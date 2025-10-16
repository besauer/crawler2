from __future__ import annotations

import collections
import time
from dataclasses import dataclass
from typing import Callable, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin, urldefrag, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup

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


def extract_links(base_url: str, html: str) -> Iterable[str]:
    soup = BeautifulSoup(html, "html.parser")
    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href")
        if not href:
            continue
        absolute = urljoin(base_url, href)
        yield normalize_url(absolute)


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

        for link in extract_links(response.url, page_text):
            if link not in visited:
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
