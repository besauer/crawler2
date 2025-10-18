from __future__ import annotations

import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, wait
from dataclasses import dataclass
from datetime import date
from typing import Callable, Dict, Iterable, List, Optional, Pattern, Set, Tuple
from urllib.parse import urljoin, urldefrag, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser

USER_AGENT = "LocalSchoolCrawler/1.0 (+https://example.com/contact)"
REQUEST_TIMEOUT = 10  # seconds
MAX_PAGES_DEFAULT = 100
SLEEP_BETWEEN_REQUESTS = 0.5  # seconds
REPEATED_SNIPPET_MIN_CHARS = 8
REPEATED_SNIPPET_THRESHOLD = 1
_WHITESPACE_RE = re.compile(r"\s+")


@dataclass
class CrawlResult:
    source_url: str
    target_url: str
    matched_keywords: Tuple[str, ...]


@dataclass
class CrawlProgress:
    """Represents a progress update during a crawl."""

    event: str
    current_url: Optional[str] = None
    visited: Optional[int] = None
    queue_length: Optional[int] = None
    result: Optional[CrawlResult] = None


def normalize_url(url: str) -> str:
    """Normalize a URL by removing fragments and ensuring scheme."""
    cleaned, _fragment = urldefrag(url)
    parsed = urlparse(cleaned)
    if not parsed.scheme:
        # Assume HTTP if scheme missing, though urljoin should handle this earlier.
        cleaned = "http://" + cleaned
    return cleaned


def extract_links(base_url: str, html: str, *, soup: Optional[BeautifulSoup] = None) -> Iterable[str]:
    if soup is None:
        soup = BeautifulSoup(html, "html.parser")
    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href")
        if not href:
            continue
        absolute = urljoin(base_url, href)
        yield normalize_url(absolute)


def parse_date_string(value: str) -> Optional[date]:
    if not value:
        return None
    try:
        parsed = date_parser.parse(value, dayfirst=True, fuzzy=True)
    except (ValueError, OverflowError, TypeError):
        return None
    if not parsed:
        return None
    return parsed.date()


def extract_publication_date(soup: BeautifulSoup) -> Optional[date]:
    meta_fields = (
        ("property", "article:published_time"),
        ("property", "article:modified_time"),
        ("property", "og:published_time"),
        ("property", "og:updated_time"),
        ("name", "pubdate"),
        ("name", "publishdate"),
        ("name", "publish_date"),
        ("name", "date"),
        ("name", "dc.date"),
        ("name", "dc.date.issued"),
        ("itemprop", "datePublished"),
        ("itemprop", "dateCreated"),
        ("itemprop", "dateModified"),
    )

    for attr, value in meta_fields:
        tag = soup.find("meta", attrs={attr: value})
        if tag and tag.get("content"):
            parsed = parse_date_string(tag.get("content"))
            if parsed:
                return parsed

    time_tag = soup.find("time", attrs={"datetime": True})
    if time_tag:
        parsed = parse_date_string(time_tag.get("datetime"))
        if parsed:
            return parsed

    date_like = soup.find(
        lambda tag: tag.name in {"time", "span", "div", "p"}
        and any(
            cls in (tag.get("class") or []) for cls in ["date", "datum", "time", "published"]
        )
    )
    if date_like and date_like.get_text(strip=True):
        parsed = parse_date_string(date_like.get_text(strip=True))
        if parsed:
            return parsed

    text = soup.get_text(" ", strip=True)
    match = re.search(r"(20\d{2}|19\d{2})[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])", text)
    if match:
        parsed = parse_date_string(match.group(0))
        if parsed:
            return parsed

    return None


KeywordPattern = Tuple[str, Pattern[str]]


def build_keyword_patterns(keywords: Iterable[str]) -> List[KeywordPattern]:
    patterns: List[KeywordPattern] = []
    seen: Set[str] = set()
    for keyword in keywords:
        text = str(keyword).strip()
        if not text:
            continue
        lowered = text.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        escaped = re.escape(text)
        if any(char.isalnum() or char == "_" for char in text):
            pattern = re.compile(rf"(?<!\\w){escaped}(?!\\w)", re.IGNORECASE)
        else:
            pattern = re.compile(escaped, re.IGNORECASE)
        patterns.append((text, pattern))
    return patterns


def keyword_matches(text: str, patterns: Iterable[KeywordPattern]) -> Tuple[str, ...]:
    matches: List[str] = []
    for original, pattern in patterns:
        if pattern.search(text):
            matches.append(original)
    return tuple(matches)


def _split_text_segments(soup: BeautifulSoup) -> List[str]:
    raw_text = soup.get_text("\n", strip=True)
    segments: List[str] = []
    for line in raw_text.splitlines():
        cleaned = _WHITESPACE_RE.sub(" ", line).strip()
        if cleaned:
            segments.append(cleaned)
    return segments


def _normalize_segment_key(segment: str) -> str:
    return _WHITESPACE_RE.sub(" ", segment).strip().lower()


def is_html_response(response: requests.Response) -> bool:
    content_type = response.headers.get("Content-Type", "").lower()
    return "text/html" in content_type


def build_robot_parser(start_url: str) -> RobotFileParser:
    parsed = urlparse(start_url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    robot_parser = RobotFileParser()
    try:
        robot_parser.set_url(robots_url)
        robot_parser.read()
    except Exception:
        # If robots.txt cannot be fetched, fallback to allowing all.
        robot_parser = RobotFileParser()
        robot_parser.parse("User-agent: *\nAllow: /".splitlines())
    return robot_parser


def crawl_site(
    start_url: str,
    keywords: Iterable[str],
    *,
    max_pages: int = MAX_PAGES_DEFAULT,
    same_domain_only: bool = True,
    progress_callback: Optional[Callable[[CrawlProgress], None]] = None,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    cancel_event: Optional[threading.Event] = None,
    max_workers: int = 5,
    respect_robots: bool = True,
) -> List[CrawlResult]:
    """Crawl pages starting from start_url and return URLs containing keywords."""

    if max_pages <= 0:
        return []

    normalized_start = normalize_url(start_url)
    parsed_start = urlparse(normalized_start)

    keywords = tuple(filter(None, (kw.strip() for kw in keywords)))
    if not keywords:
        return []

    keyword_patterns = build_keyword_patterns(keywords)
    if not keyword_patterns:
        return []

    if respect_robots:
        robot_parser = build_robot_parser(normalized_start)

        def can_fetch(url: str) -> bool:
            return robot_parser.can_fetch(USER_AGENT, url)

    else:
        robot_parser = None

        def can_fetch(_url: str) -> bool:
            return True

    visited: Set[str] = set()
    allowed_urls: Set[str] = {normalized_start}
    results: List[CrawlResult] = []
    result_pairs: Set[Tuple[str, str, str]] = set()
    snippet_counts: Dict[str, int] = {}
    snippet_lock = threading.Lock()

    def prepare_search_text(soup_obj: BeautifulSoup) -> str:
        segments = _split_text_segments(soup_obj)
        if not segments:
            return ""
        filtered_segments: List[str] = []
        update_keys: Set[str] = set()
        with snippet_lock:
            for segment in segments:
                qualifies = len(segment) >= REPEATED_SNIPPET_MIN_CHARS
                if qualifies:
                    key = _normalize_segment_key(segment)
                    if key not in update_keys:
                        update_keys.add(key)
                    if snippet_counts.get(key, 0) >= REPEATED_SNIPPET_THRESHOLD:
                        continue
                filtered_segments.append(segment)
            for key in update_keys:
                snippet_counts[key] = snippet_counts.get(key, 0) + 1
        return " ".join(filtered_segments)

    state_lock = threading.Lock()
    state = {
        "visited_count": 0,
        "queue_remaining": 0,
        "in_progress": 0,
    }

    def emit(progress: CrawlProgress) -> None:
        if progress_callback:
            progress_callback(progress)

    if cancel_event and cancel_event.is_set():
        return []

    emit(
        CrawlProgress(
            event="start",
            current_url=normalized_start,
            visited=0,
            queue_length=1,
        )
    )

    if not can_fetch(normalized_start):
        emit(
            CrawlProgress(
                event="finish",
                current_url=None,
                visited=0,
                queue_length=0,
            )
        )
        return []

    emit(
        CrawlProgress(
            event="page",
            current_url=normalized_start,
            visited=0,
            queue_length=0,
        )
    )

    try:
        response = requests.get(
            normalized_start,
            headers={"User-Agent": USER_AGENT},
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )
    except requests.RequestException:
        with state_lock:
            visited.add(normalized_start)
            state["visited_count"] = len(visited)
            state["queue_remaining"] = 0
            visited_now = state["visited_count"]
        emit(
            CrawlProgress(
                event="visited",
                current_url=normalized_start,
                visited=visited_now,
                queue_length=0,
            )
        )
        emit(
            CrawlProgress(
                event="finish",
                current_url=None,
                visited=visited_now,
                queue_length=0,
            )
        )
        return results

    page_text = response.text if response and response.text else ""
    soup = None
    if is_html_response(response) and page_text:
        soup = BeautifulSoup(page_text, "html.parser")

    with state_lock:
        visited.add(normalized_start)
        state["visited_count"] = len(visited)
        visited_now = state["visited_count"]

    remaining_slots = max(0, max_pages - visited_now)
    allowed_links: List[str] = []

    if soup and remaining_slots:
        for link in extract_links(response.url, page_text, soup=soup):
            parsed_link = urlparse(link)
            if same_domain_only and parsed_link.netloc != parsed_start.netloc:
                continue
            if link in allowed_urls or link in visited:
                continue
            if not can_fetch(link):
                continue
            allowed_urls.add(link)
            allowed_links.append(link)
            if len(allowed_links) >= remaining_slots:
                break

    with state_lock:
        state["queue_remaining"] = len(allowed_links)
        queue_after_start = max(0, state["queue_remaining"] + state["in_progress"])
        visited_now = state["visited_count"]

    emit(
        CrawlProgress(
            event="visited",
            current_url=normalized_start,
            visited=visited_now,
            queue_length=queue_after_start,
        )
    )

    search_text = page_text
    if soup:
        search_text = prepare_search_text(soup)

    if search_text:
        matches = keyword_matches(search_text, keyword_patterns)
        for match in matches:
            result = CrawlResult(
                source_url=normalized_start,
                target_url=normalized_start,
                matched_keywords=(match,),
            )
            lowered_match = match.lower()
            with state_lock:
                key = (result.source_url, result.target_url, lowered_match)
                if key not in result_pairs:
                    result_pairs.add(key)
                    results.append(result)
            emit(
                CrawlProgress(
                    event="match",
                    current_url=normalized_start,
                    visited=visited_now,
                    queue_length=queue_after_start,
                    result=result,
                )
            )

    time.sleep(SLEEP_BETWEEN_REQUESTS)

    if cancel_event and cancel_event.is_set():
        emit(
            CrawlProgress(
                event="finish",
                current_url=None,
                visited=visited_now,
                queue_length=0,
            )
        )
        return results

    if not allowed_links:
        emit(
            CrawlProgress(
                event="finish",
                current_url=None,
                visited=visited_now,
                queue_length=0,
            )
        )
        return results

    max_workers = max(1, min(max_workers, len(allowed_links)))

    def worker(url: str) -> None:

        if cancel_event and cancel_event.is_set():
            return

        with state_lock:
            if url in visited:
                return
            if state["queue_remaining"] > 0:
                state["queue_remaining"] -= 1
            state["in_progress"] += 1
            visited_before = state["visited_count"]
            queue_length_before = max(0, state["queue_remaining"] + max(0, state["in_progress"] - 1))

        emit(
            CrawlProgress(
                event="page",
                current_url=url,
                visited=visited_before,
                queue_length=queue_length_before,
            )
        )

        if cancel_event and cancel_event.is_set():
            with state_lock:
                if state["in_progress"] > 0:
                    state["in_progress"] -= 1
                queue_length_after_cancel = max(0, state["queue_remaining"] + state["in_progress"])
                visited_count_current = state["visited_count"]
            emit(
                CrawlProgress(
                    event="visited",
                    current_url=url,
                    visited=visited_count_current,
                    queue_length=queue_length_after_cancel,
                )
            )
            return

        try:
            response = requests.get(
                url,
                headers={"User-Agent": USER_AGENT},
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )
        except requests.RequestException:
            response = None

        page_text_local = response.text if response and response.text else ""
        soup_local = None
        if response and is_html_response(response) and page_text_local:
            soup_local = BeautifulSoup(page_text_local, "html.parser")

        within_range = True
        if soup_local and (start_date is not None or end_date is not None):
            publication_date = extract_publication_date(soup_local)
            if publication_date is None:
                within_range = False
            else:
                if start_date and publication_date < start_date:
                    within_range = False
                if end_date and publication_date > end_date:
                    within_range = False

        filtered_text_local = page_text_local or ""
        if soup_local:
            filtered_text_local = prepare_search_text(soup_local)

        matches_local: Tuple[str, ...] = ()
        if within_range and filtered_text_local:
            matches_local = keyword_matches(filtered_text_local, keyword_patterns)

        with state_lock:
            visited.add(url)
            state["visited_count"] = len(visited)
            if state["in_progress"] > 0:
                state["in_progress"] -= 1
            queue_length_after = max(0, state["queue_remaining"] + state["in_progress"])
            visited_now_local = state["visited_count"]

        emit(
            CrawlProgress(
                event="visited",
                current_url=url,
                visited=visited_now_local,
                queue_length=queue_length_after,
            )
        )

        for match in matches_local:
            result_local = CrawlResult(
                source_url=normalized_start,
                target_url=url,
                matched_keywords=(match,),
            )
            lowered_match = match.lower()
            with state_lock:
                key = (result_local.source_url, result_local.target_url, lowered_match)
                if key not in result_pairs:
                    result_pairs.add(key)
                    results.append(result_local)
            emit(
                CrawlProgress(
                    event="match",
                    current_url=url,
                    visited=visited_now_local,
                    queue_length=queue_length_after,
                    result=result_local,
                )
            )

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(worker, url) for url in allowed_links]
        wait(futures)

    with state_lock:
        final_visited = state["visited_count"]

    emit(
        CrawlProgress(
            event="finish",
            current_url=None,
            visited=final_visited,
            queue_length=0,
        )
    )

    return results


__all__ = ["crawl_site", "CrawlResult", "CrawlProgress"]
