from __future__ import annotations

import json
import threading
import time
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from flask import Flask, jsonify, render_template, request

from crawler import CrawlProgress, CrawlResult, crawl_site, MAX_PAGES_DEFAULT

app = Flask(__name__)

SAVED_SEARCHES_PATH = Path("saved_searches.json")
saved_search_lock = threading.Lock()


def _read_saved_searches_unlocked() -> List[Dict[str, object]]:
    if not SAVED_SEARCHES_PATH.exists():
        return []
    try:
        with SAVED_SEARCHES_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (json.JSONDecodeError, OSError):
        return []
    if isinstance(data, list):
        return data
    return []


def load_saved_searches() -> List[Dict[str, object]]:
    with saved_search_lock:
        return list(_read_saved_searches_unlocked())


def append_saved_search(entry: Dict[str, object]) -> None:
    with saved_search_lock:
        data = _read_saved_searches_unlocked()
        data.append(entry)
        try:
            with SAVED_SEARCHES_PATH.open("w", encoding="utf-8") as handle:
                json.dump(data, handle, ensure_ascii=False, indent=2)
        except OSError:
            # If persisting fails we silently ignore to avoid breaking the crawl UI.
            pass


@dataclass
class CrawlJob:
    """Container for the state of a running crawl job."""

    id: str
    start_urls: List[str]
    keywords: List[str]
    max_pages: int
    total_start_urls: int
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    processed_start_urls: int = 0
    current_start_url: Optional[str] = None
    current_page: Optional[str] = None
    visited_in_current: int = 0
    queue_length: int = 0
    recent_pages: List[str] = field(default_factory=list)
    status: str = "pending"  # pending, running, cancelling, cancelled, finished, error
    error: Optional[str] = None
    results: List[CrawlResult] = field(default_factory=list)
    found_count: int = 0
    progress_percent: int = 0
    started_at: float = field(default_factory=time.time)
    completed: bool = False
    result_pairs: Set[Tuple[str, str]] = field(default_factory=set, repr=False, compare=False)
    cancel_event: threading.Event = field(default_factory=threading.Event, repr=False, compare=False)
    cancelled: bool = False
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
                "start_date": self.start_date.isoformat() if self.start_date else None,
                "end_date": self.end_date.isoformat() if self.end_date else None,
                "cancelled": self.cancelled,
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

    def mark_cancelled(self) -> None:
        with self._lock:
            self.status = "cancelled"
            self.cancelled = True
            self.completed = True
            self.queue_length = 0
            self.current_page = None
            self.visited_in_current = 0

    def request_cancel(self) -> None:
        with self._lock:
            if not self.completed:
                self.status = "cancelling"
        self.cancel_event.set()


jobs: Dict[str, CrawlJob] = {}


def run_crawl_job(job: CrawlJob) -> None:
    try:
        for index, start_url in enumerate(job.start_urls, start=1):
            if job.cancel_event.is_set():
                job.mark_cancelled()
                return
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
                start_date=job.start_date,
                end_date=job.end_date,
                cancel_event=job.cancel_event,
            )
            if job.cancel_event.is_set():
                job.mark_cancelled()
                return
            job.mark_start_url_finished()

        if job.cancel_event.is_set():
            job.mark_cancelled()
        else:
            job.mark_completed()
    except Exception as exc:  # pragma: no cover - defensive safety net
        job.mark_completed(error=str(exc))


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        raw_start_urls = request.form.get("start_urls", "")
        raw_keywords = request.form.get("keywords", "")
        max_pages = request.form.get("max_pages", type=int, default=MAX_PAGES_DEFAULT)
        raw_start_date = request.form.get("start_date", "").strip()
        raw_end_date = request.form.get("end_date", "").strip()

        start_urls = [url.strip() for url in raw_start_urls.splitlines() if url.strip()]
        keywords = [kw.strip() for kw in raw_keywords.splitlines() if kw.strip()]
        error = None
        job_id = None
        start_date_value: Optional[date] = None
        end_date_value: Optional[date] = None

        if not start_urls:
            error = "Bitte geben Sie mindestens eine Start-URL ein."
        elif not keywords:
            error = "Bitte geben Sie mindestens ein Suchstichwort ein."
        else:
            if raw_start_date:
                try:
                    start_date_value = datetime.strptime(raw_start_date, "%Y-%m-%d").date()
                except ValueError:
                    error = "Das Startdatum ist ungültig."
            if not error and raw_end_date:
                try:
                    end_date_value = datetime.strptime(raw_end_date, "%Y-%m-%d").date()
                except ValueError:
                    error = "Das Enddatum ist ungültig."
            if (
                not error
                and start_date_value
                and end_date_value
                and start_date_value > end_date_value
            ):
                error = "Das Startdatum darf nicht nach dem Enddatum liegen."

        if not error:
            job_id = uuid.uuid4().hex
            job = CrawlJob(
                id=job_id,
                start_urls=start_urls,
                keywords=keywords,
                max_pages=max_pages,
                total_start_urls=len(start_urls),
                start_date=start_date_value,
                end_date=end_date_value,
            )
            jobs[job_id] = job
            thread = threading.Thread(target=run_crawl_job, args=(job,), daemon=True)
            thread.start()

        return render_template(
            "index.html",
            start_urls_input=raw_start_urls,
            keywords_input=raw_keywords,
            max_pages=max_pages,
            start_date_input=raw_start_date,
            end_date_input=raw_end_date,
            error=error,
            submitted=True,
            job_id=job_id,
        )

    return render_template(
        "index.html",
        start_urls_input="",
        keywords_input="",
        max_pages=MAX_PAGES_DEFAULT,
        start_date_input="",
        end_date_input="",
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


@app.route("/cancel/<job_id>", methods=["POST"])
def cancel_job(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unbekannte Job-ID"}), 404
    job.request_cancel()
    return jsonify({"status": job.status})


@app.route("/save-search/<job_id>", methods=["POST"])
def save_search(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unbekannte Job-ID"}), 404
    with job._lock:
        if not job.completed:
            return jsonify({"error": "Der Crawl läuft noch. Bitte warten Sie, bis er beendet ist."}), 400
        entry = {
            "saved_at": datetime.utcnow().isoformat() + "Z",
            "job_id": job.id,
            "start_urls": list(job.start_urls),
            "keywords": list(job.keywords),
            "start_date": job.start_date.isoformat() if job.start_date else None,
            "end_date": job.end_date.isoformat() if job.end_date else None,
            "status": job.status,
            "result_count": len(job.results),
            "results": [
                {
                    "source_url": result.source_url,
                    "target_url": result.target_url,
                    "matched_keywords": list(result.matched_keywords),
                }
                for result in job.results
            ],
        }
    append_saved_search(entry)
    return jsonify({"status": "saved"})


@app.route("/saved-searches")
def get_saved_searches():
    return jsonify({"searches": load_saved_searches()})


if __name__ == "__main__":
    app.run(debug=True)
