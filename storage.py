import json
import sqlite3
import threading
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


class SQLiteStorage:
    """Simple SQLite-backed key/value storage with namespaces."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._local = threading.local()
        self._init_lock = threading.Lock()
        self._ensure_initialized()

    def ensure_schema(self) -> None:
        """Ensure the SQLite schema exists (idempotent)."""
        self._ensure_initialized()

    def _get_connection(self) -> sqlite3.Connection:
        self.ensure_schema()
        conn = getattr(self._local, "connection", None)
        if conn is None:
            conn = sqlite3.connect(
                self.path,
                detect_types=sqlite3.PARSE_DECLTYPES,
                check_same_thread=False,
                timeout=30,
            )
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA foreign_keys=ON")
            self._local.connection = conn
        return conn

    def _ensure_initialized(self) -> None:
        with self._init_lock:
            conn = sqlite3.connect(
                self.path,
                detect_types=sqlite3.PARSE_DECLTYPES,
                check_same_thread=False,
                timeout=30,
            )
            try:
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("PRAGMA foreign_keys=ON")
                conn.executescript(
                    """
                    CREATE TABLE IF NOT EXISTS kv_store (
                        namespace TEXT NOT NULL,
                        key TEXT NOT NULL,
                        value TEXT NOT NULL,
                        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        PRIMARY KEY(namespace, key)
                    );

                    CREATE TABLE IF NOT EXISTS saved_searches (
                        id TEXT PRIMARY KEY,
                        payload TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    );

                    CREATE TABLE IF NOT EXISTS audit_log (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        event TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    );

                    CREATE TABLE IF NOT EXISTS search_definitions (
                        id TEXT PRIMARY KEY,
                        payload TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                    );

                    CREATE TABLE IF NOT EXISTS search_runs (
                        id TEXT PRIMARY KEY,
                        definition_id TEXT NOT NULL,
                        payload TEXT NOT NULL,
                        status TEXT NOT NULL,
                        started_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        completed_at TEXT,
                        FOREIGN KEY(definition_id) REFERENCES search_definitions(id) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS search_results (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        run_id TEXT NOT NULL,
                        payload TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(run_id) REFERENCES search_runs(id) ON DELETE CASCADE
                    );

                    CREATE INDEX IF NOT EXISTS idx_search_runs_definition ON search_runs(definition_id);
                    CREATE INDEX IF NOT EXISTS idx_search_results_run ON search_results(run_id);
                    CREATE TABLE IF NOT EXISTS search_logs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        run_id TEXT NOT NULL,
                        payload TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(run_id) REFERENCES search_runs(id) ON DELETE CASCADE
                    );
                    CREATE INDEX IF NOT EXISTS idx_search_logs_run ON search_logs(run_id);
                    CREATE TABLE IF NOT EXISTS error_log (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        payload TEXT NOT NULL,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    );
                    CREATE INDEX IF NOT EXISTS idx_error_log_created ON error_log(created_at DESC);
                """
            )
                conn.commit()
            finally:
                conn.close()

    # Generic helpers -------------------------------------------------
    def get_json(self, namespace: str, key: str, default: Any) -> Any:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT value FROM kv_store WHERE namespace = ? AND key = ?",
            (namespace, key),
        )
        row = cursor.fetchone()
        if row is None:
            return default
        try:
            return json.loads(row["value"])
        except json.JSONDecodeError:
            return default

    def set_json(self, namespace: str, key: str, value: Any) -> None:
        conn = self._get_connection()
        payload = json.dumps(value, ensure_ascii=False)
        with conn:
            conn.execute(
                "REPLACE INTO kv_store(namespace, key, value, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP)",
                (namespace, key, payload),
            )

    def delete(self, namespace: str, keys: Iterable[str]) -> None:
        keys = list(keys)
        if not keys:
            return
        conn = self._get_connection()
        with conn:
            conn.executemany(
                "DELETE FROM kv_store WHERE namespace = ? AND key = ?",
                [(namespace, key) for key in keys],
            )

    # Saved searches --------------------------------------------------
    def list_saved_searches(self) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT id, payload FROM saved_searches ORDER BY created_at DESC, id DESC"
        )
        for row in cursor:
            try:
                data = json.loads(row["payload"])
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict):
                data.setdefault("id", row["id"])
                yield data

    def replace_saved_searches(self, searches: Iterable[Dict[str, Any]]) -> None:
        conn = self._get_connection()
        with conn:
            conn.execute("DELETE FROM saved_searches")
            conn.executemany(
                "INSERT INTO saved_searches(id, payload) VALUES(?, ?)",
                [
                    (
                        str(item.get("id")),
                        json.dumps(item, ensure_ascii=False),
                    )
                    for item in searches
                ],
            )

    def append_saved_search(self, search: Dict[str, Any]) -> None:
        conn = self._get_connection()
        payload = json.dumps(search, ensure_ascii=False)
        with conn:
            conn.execute(
                "REPLACE INTO saved_searches(id, payload, created_at) VALUES(?, ?, CURRENT_TIMESTAMP)",
                (str(search.get("id")), payload),
            )

    # Audit log -------------------------------------------------------
    def append_audit_log(self, event: Dict[str, Any]) -> None:
        conn = self._get_connection()
        payload = json.dumps(event, ensure_ascii=False)
        with conn:
            conn.execute(
                "INSERT INTO audit_log(event) VALUES(?)",
                (payload,),
            )

    def read_audit_log(self) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT event FROM audit_log ORDER BY created_at DESC, id DESC"
        )
        for row in cursor:
            try:
                data = json.loads(row["event"])
            except json.JSONDecodeError:
                continue
            yield data

    # Error log -------------------------------------------------------
    def append_error_log(self, payload: Dict[str, Any]) -> None:
        conn = self._get_connection()
        raw = json.dumps(payload, ensure_ascii=False)
        with conn:
            conn.execute(
                "INSERT INTO error_log(payload) VALUES(?)",
                (raw,),
            )

    def list_error_log(self, limit: Optional[int] = None) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        query = "SELECT id, payload, created_at FROM error_log ORDER BY created_at DESC, id DESC"
        if limit is not None:
            cursor = conn.execute(query + " LIMIT ?", (int(limit),))
        else:
            cursor = conn.execute(query)
        for row in cursor:
            try:
                data = json.loads(row["payload"])
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict):
                data.setdefault("id", row["id"])
                data.setdefault("created_at", row["created_at"])
                yield data

    def replace_error_log(self, entries: Iterable[Dict[str, Any]]) -> None:
        conn = self._get_connection()
        serialised = [json.dumps(entry, ensure_ascii=False) for entry in entries if isinstance(entry, dict)]
        with conn:
            conn.execute("DELETE FROM error_log")
            if serialised:
                conn.executemany(
                    "INSERT INTO error_log(payload) VALUES(?)",
                    [(item,) for item in serialised],
                )

    # Search definitions ---------------------------------------------
    def list_search_definitions(self) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT id, payload FROM search_definitions ORDER BY created_at DESC, id DESC"
        )
        for row in cursor:
            try:
                data = json.loads(row["payload"])
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict):
                data.setdefault("id", row["id"])
                yield data

    def get_search_definition(self, definition_id: str) -> Optional[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT payload FROM search_definitions WHERE id = ?",
            (definition_id,),
        )
        row = cursor.fetchone()
        if not row:
            return None
        try:
            data = json.loads(row["payload"])
        except json.JSONDecodeError:
            return None
        if isinstance(data, dict):
            data.setdefault("id", definition_id)
            return data
        return None

    def upsert_search_definition(self, definition: Dict[str, Any]) -> None:
        if not isinstance(definition, dict):
            return
        identifier = str(definition.get("id") or "").strip()
        if not identifier:
            raise ValueError("search definition requires an id")
        payload = json.dumps(definition, ensure_ascii=False)
        conn = self._get_connection()
        with conn:
            conn.execute(
                """
                INSERT INTO search_definitions(id, payload, created_at, updated_at)
                VALUES(?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                ON CONFLICT(id) DO UPDATE SET
                    payload = excluded.payload,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (identifier, payload),
            )

    def delete_search_definitions(self, identifiers: Iterable[str]) -> None:
        ids = [str(item) for item in identifiers if str(item)]
        if not ids:
            return
        conn = self._get_connection()
        with conn:
            conn.executemany(
                "DELETE FROM search_definitions WHERE id = ?",
                [(item,) for item in ids],
            )

    # Search runs ----------------------------------------------------
    def upsert_search_run(self, run_id: str, definition_id: str, payload: Dict[str, Any]) -> None:
        if not run_id or not definition_id:
            raise ValueError("search run requires id and definition id")
        data_payload = json.dumps(payload, ensure_ascii=False)
        conn = self._get_connection()
        with conn:
            conn.execute(
                """
                INSERT INTO search_runs(id, definition_id, payload, status, started_at)
                VALUES(?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(id) DO UPDATE SET
                    payload = excluded.payload,
                    status = excluded.status,
                    completed_at = CASE
                        WHEN json_extract(excluded.payload, '$.status') IN ('finished', 'error', 'cancelled')
                        THEN CURRENT_TIMESTAMP
                        ELSE completed_at
                    END
                """,
                (
                    run_id,
                    definition_id,
                    data_payload,
                    str(payload.get("status", "pending")),
                ),
            )

    def update_search_run_status(self, run_id: str, status: str, **fields: Any) -> None:
        if not run_id:
            return
        current = self.get_search_run(run_id) or {}
        current.update(fields)
        current["status"] = status
        self.upsert_search_run(run_id, str(current.get("definition_id", "")), current)

    def get_search_run(self, run_id: str) -> Optional[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT definition_id, payload FROM search_runs WHERE id = ?",
            (run_id,),
        )
        row = cursor.fetchone()
        if not row:
            return None
        try:
            payload = json.loads(row["payload"])
        except json.JSONDecodeError:
            payload = {}
        if isinstance(payload, dict):
            payload.setdefault("id", run_id)
            payload.setdefault("definition_id", row["definition_id"])
            return payload
        return None

    def list_search_runs(self, definition_id: Optional[str] = None) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        if definition_id:
            cursor = conn.execute(
                "SELECT id, definition_id, payload FROM search_runs WHERE definition_id = ? ORDER BY started_at DESC",
                (definition_id,),
            )
        else:
            cursor = conn.execute(
                "SELECT id, definition_id, payload FROM search_runs ORDER BY started_at DESC"
            )
        for row in cursor:
            try:
                payload = json.loads(row["payload"])
            except json.JSONDecodeError:
                payload = {}
            if isinstance(payload, dict):
                payload.setdefault("id", row["id"])
                payload.setdefault("definition_id", row["definition_id"])
                yield payload

    def append_search_results(self, run_id: str, entries: Iterable[Dict[str, Any]]) -> None:
        run_id = str(run_id or "").strip()
        if not run_id:
            raise ValueError("run_id required for search results")
        payloads: List[str] = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            cleaned = dict(entry)
            cleaned.pop("id", None)
            payloads.append(json.dumps(cleaned, ensure_ascii=False))
        if not payloads:
            return
        conn = self._get_connection()
        with conn:
            conn.executemany(
                "INSERT INTO search_results(run_id, payload) VALUES(?, ?)",
                [(run_id, payload) for payload in payloads],
            )

    def load_search_results(self, run_id: str) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT id, payload FROM search_results WHERE run_id = ? ORDER BY id",
            (run_id,),
        )
        for row in cursor:
            try:
                data = json.loads(row["payload"])
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict):
                data.setdefault("id", row["id"])
                data.setdefault("run_id", run_id)
                yield data

    def clear_search_results(self, run_id: str) -> None:
        conn = self._get_connection()
        with conn:
            conn.execute("DELETE FROM search_results WHERE run_id = ?", (run_id,))

    def append_search_logs(self, run_id: str, entries: Iterable[Dict[str, Any]]) -> None:
        run_id = str(run_id or "").strip()
        if not run_id:
            raise ValueError("run_id required for search logs")
        payloads: List[str] = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            cleaned = dict(entry)
            cleaned.pop("id", None)
            payloads.append(json.dumps(cleaned, ensure_ascii=False))
        if not payloads:
            return
        conn = self._get_connection()
        with conn:
            conn.executemany(
                "INSERT INTO search_logs(run_id, payload) VALUES(?, ?)",
                [(run_id, payload) for payload in payloads],
            )

    def load_search_logs(self, run_id: str) -> Iterable[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT id, payload FROM search_logs WHERE run_id = ? ORDER BY id",
            (run_id,),
        )
        for row in cursor:
            try:
                data = json.loads(row["payload"])
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict):
                data.setdefault("id", row["id"])
                data.setdefault("run_id", run_id)
                yield data

    def clear_search_logs(self, run_id: str) -> None:
        conn = self._get_connection()
        with conn:
            conn.execute("DELETE FROM search_logs WHERE run_id = ?", (run_id,))

    def update_search_result(self, result_id: int, payload: Dict[str, Any]) -> None:
        conn = self._get_connection()
        data = dict(payload)
        data.pop("id", None)
        serialized = json.dumps(data, ensure_ascii=False)
        with conn:
            conn.execute(
                "UPDATE search_results SET payload = ? WHERE id = ?",
                (serialized, int(result_id)),
            )

    def delete_search_results(self, result_ids: Iterable[int]) -> None:
        ids = [int(rid) for rid in result_ids if rid is not None]
        if not ids:
            return
        conn = self._get_connection()
        with conn:
            conn.executemany(
                "DELETE FROM search_results WHERE id = ?",
                [(rid,) for rid in ids],
            )

    def get_search_result(self, result_id: int) -> Optional[Dict[str, Any]]:
        conn = self._get_connection()
        cursor = conn.execute(
            "SELECT id, run_id, payload FROM search_results WHERE id = ?",
            (int(result_id),),
        )
        row = cursor.fetchone()
        if not row:
            return None
        try:
            data = json.loads(row["payload"])
        except json.JSONDecodeError:
            data = {}
        if not isinstance(data, dict):
            data = {}
        data.setdefault("id", row["id"])
        data.setdefault("run_id", row["run_id"])
        return data

    def list_search_categories(self) -> List[str]:
        data = self.get_json("searches", "categories", [])
        if isinstance(data, list):
            return [str(item).strip() for item in data if str(item).strip()]
        return []

    def save_search_categories(self, categories: Iterable[str]) -> None:
        cleaned = [str(item).strip() for item in categories if str(item).strip()]
        self.set_json("searches", "categories", cleaned)


storage = SQLiteStorage(Path("crawler_data.db"))
