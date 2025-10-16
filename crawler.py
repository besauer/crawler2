from __future__ import annotations

import collections
import re
import threading
import time
from dataclasses import dataclass
from datetime import date
from typing import Callable, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin, urldefrag, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser

USER_AGENT = "LocalSchoolCrawler/1.0 (+https://example.com/contact)"
REQUEST_TIMEOUT = 10  # seconds
MAX_PAGES_DEFAULT = 100
SLEEP_BETWEEN_REQUESTS = 0.5  # seconds


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


def keyword_matches(text: str, keywords: Iterable[str]) -> Tuple[str, ...]:
    lower_text = text.lower()
    matches = []
    for keyword in keywords:
        keyword = keyword.strip().lower()
        if not keyword:
            continue
        # Use simple containment. Regex word boundaries optional? We'll use containment.
        if keyword in lower_text:
            matches.append(keyword)
    return tuple(matches)


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
) -> List[CrawlResult]:
    """Crawl pages starting from start_url and return URLs containing keywords."""

    if max_pages <= 0:
        return []

    normalized_start = normalize_url(start_url)
    parsed_start = urlparse(normalized_start)

    keywords = tuple(filter(None, (kw.strip() for kw in keywords)))
    if not keywords:
        return []

    robot_parser = build_robot_parser(normalized_start)

    queue: "collections.deque[str]" = collections.deque([normalized_start])
    visited: Set[str] = set()
    allowed_urls: Set[str] = {normalized_start}
    results: List[CrawlResult] = []

    if progress_callback:
        progress_callback(
            CrawlProgress(
                event="start",
                current_url=normalized_start,
                visited=0,
                queue_length=len(queue),
            )
        )

    while queue and len(visited) < max_pages:
        if cancel_event and cancel_event.is_set():
            break
        current_url = queue.popleft()
        if current_url in visited:
            continue

        parsed_current = urlparse(current_url)
        if same_domain_only and parsed_current.netloc != parsed_start.netloc:
            continue

        if not robot_parser.can_fetch(USER_AGENT, current_url):
            continue

        if progress_callback:
            progress_callback(
                CrawlProgress(
                    event="page",
                    current_url=current_url,
                    visited=len(visited),
                    queue_length=len(queue),
                )
            )

        try:
            response = requests.get(
                current_url,
                headers={"User-Agent": USER_AGENT},
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )
        except requests.RequestException:
            visited.add(current_url)
            continue

        visited.add(current_url)

        if progress_callback:
            progress_callback(
                CrawlProgress(
                    event="visited",
                    current_url=current_url,
                    visited=len(visited),
                    queue_length=len(queue),
                )
            )

        if not is_html_response(response) or not response.text:
            continue

        page_text = response.text
        soup = BeautifulSoup(page_text, "html.parser")

        should_filter = (
            current_url != normalized_start and (start_date is not None or end_date is not None)
        )
        within_range = True
        publication_date: Optional[date] = None
        if should_filter:
            publication_date = extract_publication_date(soup)
            if publication_date is None:
                within_range = False
            else:
                if start_date and publication_date < start_date:
                    within_range = False
                if end_date and publication_date > end_date:
                    within_range = False

        if not should_filter or within_range:
            matches = keyword_matches(page_text, keywords)
            if matches:
                crawl_result = CrawlResult(
                    source_url=normalized_start,
                    target_url=current_url,
                    matched_keywords=matches,
                )
                results.append(crawl_result)
                if progress_callback:
                    progress_callback(
                        CrawlProgress(
                            event="match",
                            current_url=current_url,
                            visited=len(visited),
                            queue_length=len(queue),
                            result=crawl_result,
                        )
                    )

        if current_url == normalized_start:
            for link in extract_links(response.url, page_text, soup=soup):
                parsed_link = urlparse(link)
                if same_domain_only and parsed_link.netloc != parsed_start.netloc:
                    continue
                if link in allowed_urls or link in visited:
                    continue
                allowed_urls.add(link)
                queue.append(link)

        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if progress_callback:
        progress_callback(
            CrawlProgress(
                event="finish",
                current_url=None,
                visited=len(visited),
                queue_length=len(queue),
            )
        )

    return results


__all__ = ["crawl_site", "CrawlResult", "CrawlProgress"]
