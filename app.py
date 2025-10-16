from __future__ import annotations

import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from flask import Flask, jsonify, render_template, request

from crawler import CrawlProgress, CrawlResult, crawl_site, MAX_PAGES_DEFAULT

app = Flask(__name__)


@dataclass
class CrawlJob:
    """Container for the state of a running crawl job."""

    id: str
    start_urls: List[str]
    keywords: List[str]
    max_pages: int
    total_start_urls: int
    processed_start_urls: int = 0
    current_start_url: Optional[str] = None
    current_page: Optional[str] = None
    visited_in_current: int = 0
    queue_length: int = 0
    recent_pages: List[str] = field(default_factory=list)
    status: str = "pending"  # pending, running, finished, error
    error: Optional[str] = None
    results: List[CrawlResult] = field(default_factory=list)
    found_count: int = 0
    progress_percent: int = 0
    started_at: float = field(default_factory=time.time)
    completed: bool = False
    result_pairs: Set[Tuple[str, str]] = field(default_factory=set, repr=False, compare=False)
    _lock: threading.Lock = field(default_factory=threading.Lock, repr=False, compare=False)

    def as_dict(self) -> Dict[str, object]:
        with self._lock:
            return {
                "job_id": self.id,
                "status": self.status,
                "error": self.error,
                "total_start_urls": self.total_start_urls,
                "processed_start_urls": self.processed_start_urls,
                "current_start_url": self.current_start_url,
                "current_page": self.current_page,
                "visited_in_current": self.visited_in_current,
                "queue_length": self.queue_length,
                "found_count": self.found_count,
                "progress_percent": self.progress_percent,
                "completed": self.completed,
                "recent_pages": list(self.recent_pages),
                "results": [
                    {
                        "source_url": result.source_url,
                        "target_url": result.target_url,
                        "matched_keywords": list(result.matched_keywords),
                    }
                    for result in self.results
                ],
            }

    def update_progress(self, progress: CrawlProgress) -> None:
        with self._lock:
            if progress.queue_length is not None:
                self.queue_length = progress.queue_length

            if progress.current_url:
                self.current_page = progress.current_url
                if progress.event in {"page", "visited"}:
                    if not self.recent_pages or self.recent_pages[-1] != progress.current_url:
                        self.recent_pages.append(progress.current_url)
                        if len(self.recent_pages) > 10:
                            self.recent_pages = self.recent_pages[-10:]

            if progress.event == "visited" and progress.visited is not None:
                self.visited_in_current = progress.visited

            if progress.event == "match" and progress.result:
                key = (progress.result.source_url, progress.result.target_url)
                if key not in self.result_pairs:
                    self.results.append(progress.result)
                    self.result_pairs.add(key)
                self.found_count = len(self.results)

            pages_total = max(1, self.total_start_urls * self.max_pages)
            overall_pages = self.processed_start_urls * self.max_pages + self.visited_in_current
            self.progress_percent = min(100, int((overall_pages / pages_total) * 100))

    def mark_start_url_finished(self) -> None:
        with self._lock:
            self.processed_start_urls += 1
            self.visited_in_current = 0
            self.current_page = None
            self.queue_length = 0
            pages_total = max(1, self.total_start_urls * self.max_pages)
            overall_pages = self.processed_start_urls * self.max_pages
            self.progress_percent = min(100, int((overall_pages / pages_total) * 100))

    def mark_completed(self, *, error: Optional[str] = None) -> None:
        with self._lock:
            if error:
                self.status = "error"
                self.error = error
            else:
                self.status = "finished"
                self.progress_percent = 100
            self.completed = True


jobs: Dict[str, CrawlJob] = {}


def run_crawl_job(job: CrawlJob) -> None:
    try:
        for index, start_url in enumerate(job.start_urls, start=1):
            with job._lock:
                job.status = "running"
                job.current_start_url = start_url
                job.current_page = start_url
                job.visited_in_current = 0

            def progress_callback(progress: CrawlProgress) -> None:
                job.update_progress(progress)

            crawl_site(
                start_url,
                job.keywords,
                max_pages=job.max_pages,
                progress_callback=progress_callback,
            )
            job.mark_start_url_finished()

        job.mark_completed()
    except Exception as exc:  # pragma: no cover - defensive safety net
        job.mark_completed(error=str(exc))


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        raw_start_urls = request.form.get("start_urls", "")
        raw_keywords = request.form.get("keywords", "")
        max_pages = request.form.get("max_pages", type=int, default=MAX_PAGES_DEFAULT)

        start_urls = [url.strip() for url in raw_start_urls.splitlines() if url.strip()]
        keywords = [kw.strip() for kw in raw_keywords.splitlines() if kw.strip()]
        error = None
        job_id = None

        if not start_urls:
            error = "Bitte geben Sie mindestens eine Start-URL ein."
        elif not keywords:
            error = "Bitte geben Sie mindestens ein Suchstichwort ein."
        else:
            job_id = uuid.uuid4().hex
            job = CrawlJob(
                id=job_id,
                start_urls=start_urls,
                keywords=keywords,
                max_pages=max_pages,
                total_start_urls=len(start_urls),
            )
            jobs[job_id] = job
            thread = threading.Thread(target=run_crawl_job, args=(job,), daemon=True)
            thread.start()

        return render_template(
            "index.html",
            start_urls_input=raw_start_urls,
            keywords_input=raw_keywords,
            max_pages=max_pages,
            error=error,
            submitted=True,
            job_id=job_id,
        )

    return render_template(
        "index.html",
        start_urls_input="",
        keywords_input="",
        max_pages=MAX_PAGES_DEFAULT,
        error=None,
        submitted=False,
        job_id=None,
    )


@app.route("/status/<job_id>")
def job_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unbekannte Job-ID"}), 404
    return jsonify(job.as_dict())


if __name__ == "__main__":
    app.run(debug=True)
