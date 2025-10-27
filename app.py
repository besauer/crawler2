from __future__ import annotations

import json
import math
import threading
import time
import unicodedata
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook

from crawler import (
    CrawlProgress,
    CrawlResult,
    crawl_site,
    MAX_PAGES_DEFAULT,
)

app = Flask(__name__)

SAVED_SEARCHES_PATH = Path("saved_searches.json")
saved_search_lock = threading.Lock()
DEFAULT_CONCURRENCY = 5
MAX_CONCURRENCY = 150
DEFAULT_RESPECT_ROBOTS = True
STAMMDATEN_PATH = Path("stammdaten.json")
stammdaten_lock = threading.Lock()

STAMMDATEN_FIELDS = [
    {"key": "schul_id", "label": "Schul ID"},
    {"key": "traeger", "label": "Träger"},
    {"key": "name", "label": "Name"},
    {"key": "ort", "label": "Ort"},
    {"key": "anzahl_schueler", "label": "Anzahl Schüler"},
    {"key": "anzahl_klassen", "label": "Anzahl Klassen"},
    {"key": "homepage", "label": "Homepage"},
]


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


def _normalize_header(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.lower().replace("-", " ").replace("_", " ")
    return " ".join(normalized.split())


def _coerce_str(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text


def _coerce_int(value: object) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int,)):
        return int(value)
    if isinstance(value, float):
        if math.isnan(value):
            return None
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    cleaned = text.replace(".", "").replace(" ", "")
    try:
        return int(cleaned)
    except ValueError:
        try:
            return int(float(cleaned.replace(",", ".")))
        except ValueError:
            return None


def load_stammdaten() -> List[Dict[str, object]]:
    with stammdaten_lock:
        if not STAMMDATEN_PATH.exists():
            return []
        try:
            with STAMMDATEN_PATH.open("r", encoding="utf-8") as handle:
                raw_data = json.load(handle)
        except (json.JSONDecodeError, OSError):
            return []

        records: List[Dict[str, object]] = []
        if isinstance(raw_data, list):
            for entry in raw_data:
                if not isinstance(entry, dict):
                    continue
                record = {
                    "schul_id": _coerce_str(entry.get("schul_id")),
                    "traeger": _coerce_str(entry.get("traeger")),
                    "name": _coerce_str(entry.get("name")),
                    "ort": _coerce_str(entry.get("ort")),
                    "anzahl_schueler": _coerce_int(entry.get("anzahl_schueler")),
                    "anzahl_klassen": _coerce_int(entry.get("anzahl_klassen")),
                    "homepage": _coerce_str(entry.get("homepage")),
                }
                if record["homepage"]:
                    records.append(record)
        return records


def save_stammdaten(records: List[Dict[str, object]]) -> None:
    with stammdaten_lock:
        try:
            with STAMMDATEN_PATH.open("w", encoding="utf-8") as handle:
                json.dump(records, handle, ensure_ascii=False, indent=2)
        except OSError:
            pass


def parse_stammdaten_excel(file_storage) -> List[Dict[str, object]]:
    try:
        file_storage.stream.seek(0)
    except (AttributeError, OSError):  # pragma: no cover - defensive
        pass

    try:
        workbook = load_workbook(file_storage, data_only=True)
    except Exception as exc:  # pragma: no cover - convert to user-facing message
        raise ValueError("Die Excel-Datei konnte nicht gelesen werden.") from exc

    worksheet = workbook.active
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Die Excel-Datei enthält keine Daten.")

    header_row = rows[0]
    header_map = {
        _normalize_header(str(cell) if cell is not None else ""): index
        for index, cell in enumerate(header_row)
    }

    expected_headers = {
        field["key"]: _normalize_header(field["label"])
        for field in STAMMDATEN_FIELDS
    }

    column_map: Dict[str, int] = {}
    for key, normalized in expected_headers.items():
        if normalized in header_map:
            column_map[key] = header_map[normalized]

    if len(column_map) < len(STAMMDATEN_FIELDS):
        for idx, field in enumerate(STAMMDATEN_FIELDS):
            column_map.setdefault(field["key"], idx)

    records: List[Dict[str, object]] = []
    for raw_row in rows[1:]:
        if raw_row is None:
            continue
        values = list(raw_row)
        if not any(cell not in (None, "") for cell in values):
            continue

        record = {}
        for field in STAMMDATEN_FIELDS:
            column_index = column_map.get(field["key"])
            cell_value = values[column_index] if column_index is not None and column_index < len(values) else None
            if field["key"] in {"anzahl_schueler", "anzahl_klassen"}:
                record[field["key"]] = _coerce_int(cell_value)
            else:
                record[field["key"]] = _coerce_str(cell_value)

        if record.get("homepage"):
            records.append(record)

    if not records:
        raise ValueError("Es konnten keine gültigen Stammdaten gefunden werden.")

    return records


@dataclass
class CrawlJob:
    """Container for the state of a running crawl job."""

    id: str
    start_urls: List[str]
    keywords: List[str]
    max_pages: int
    concurrency: int
    total_start_urls: int
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    respect_robots: bool = DEFAULT_RESPECT_ROBOTS
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
                "concurrency": self.concurrency,
                "respect_robots": self.respect_robots,
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
                max_workers=job.concurrency,
                respect_robots=job.respect_robots,
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
    stammdaten_records = load_stammdaten()
    keywords_input = ""
    selected_homepages: List[str] = []
    max_pages = MAX_PAGES_DEFAULT
    concurrency = DEFAULT_CONCURRENCY
    start_date_input = ""
    end_date_input = ""
    respect_robots = DEFAULT_RESPECT_ROBOTS
    error: Optional[str] = None
    job_id: Optional[str] = None
    submitted = False

    if request.method == "POST":
        submitted = True
        raw_keywords = request.form.get("keywords", "")
        keywords_input = raw_keywords
        max_pages = request.form.get("max_pages", type=int, default=MAX_PAGES_DEFAULT)
        concurrency = request.form.get("concurrency", type=int, default=DEFAULT_CONCURRENCY)
        start_date_input = request.form.get("start_date", "").strip()
        end_date_input = request.form.get("end_date", "").strip()
        respect_robots = bool(request.form.get("respect_robots"))

        selected_values = request.form.getlist("start_urls")
        allowed_homepages = {record["homepage"] for record in stammdaten_records if record.get("homepage")}
        selected_set = {value for value in selected_values if value in allowed_homepages}
        ordered_selection: List[str] = []
        for record in stammdaten_records:
            homepage = record.get("homepage")
            if homepage in selected_set:
                ordered_selection.append(homepage)
        selected_homepages = ordered_selection

        keywords = [kw.strip() for kw in raw_keywords.splitlines() if kw.strip()]
        start_date_value: Optional[date] = None
        end_date_value: Optional[date] = None

        if not stammdaten_records:
            error = "Es sind keine Stammdaten vorhanden. Bitte importieren Sie zuerst Schulen."
        elif not selected_homepages:
            error = "Bitte wählen Sie mindestens eine Schule aus den Stammdaten aus."
        elif not keywords:
            error = "Bitte geben Sie mindestens ein Suchstichwort ein."
        else:
            if start_date_input:
                try:
                    start_date_value = datetime.strptime(start_date_input, "%Y-%m-%d").date()
                except ValueError:
                    error = "Das Startdatum ist ungültig."
            if not error and end_date_input:
                try:
                    end_date_value = datetime.strptime(end_date_input, "%Y-%m-%d").date()
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
            if concurrency < 1:
                concurrency = 1
            elif concurrency > MAX_CONCURRENCY:
                concurrency = MAX_CONCURRENCY
            job_id = uuid.uuid4().hex
            job = CrawlJob(
                id=job_id,
                start_urls=selected_homepages,
                keywords=keywords,
                max_pages=max_pages,
                concurrency=concurrency,
                total_start_urls=len(selected_homepages),
                start_date=start_date_value,
                end_date=end_date_value,
                respect_robots=respect_robots,
            )
            jobs[job_id] = job
            thread = threading.Thread(target=run_crawl_job, args=(job,), daemon=True)
            thread.start()

    return render_template(
        "index.html",
        stammdaten_records=stammdaten_records,
        stammdaten_fields=STAMMDATEN_FIELDS,
        selected_homepages=selected_homepages,
        keywords_input=keywords_input,
        max_pages=max_pages,
        concurrency=concurrency,
        start_date_input=start_date_input,
        end_date_input=end_date_input,
        respect_robots=respect_robots,
        error=error,
        submitted=submitted,
        job_id=job_id,
        max_concurrency=MAX_CONCURRENCY,
        active_page="search",
    )


@app.route("/stammdaten", methods=["GET", "POST"])
def stammdaten():
    message: Optional[str] = None
    message_category: Optional[str] = None

    if request.method == "POST":
        uploaded = request.files.get("excel_file")
        if uploaded is None or not uploaded.filename:
            message = "Bitte wählen Sie eine Excel-Datei aus."
            message_category = "danger"
        else:
            try:
                records = parse_stammdaten_excel(uploaded)
            except ValueError as exc:
                message = str(exc)
                message_category = "danger"
            else:
                save_stammdaten(records)
                message = f"{len(records)} Stammdatensätze wurden übernommen."
                message_category = "success"

    records = load_stammdaten()
    return render_template(
        "stammdaten.html",
        records=records,
        message=message,
        message_category=message_category,
        stammdaten_fields=STAMMDATEN_FIELDS,
        storage_file=STAMMDATEN_PATH.name,
        active_page="stammdaten",
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
            "concurrency": job.concurrency,
            "respect_robots": job.respect_robots,
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
