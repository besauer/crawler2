import json
import sqlite3
import threading
from pathlib import Path
from typing import Any, Dict, Iterable


class SQLiteStorage:
    """Simple SQLite-backed key/value storage with namespaces."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._local = threading.local()
        self._init_lock = threading.Lock()
        self._ensure_initialized()

    def _get_connection(self) -> sqlite3.Connection:
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


storage = SQLiteStorage(Path("crawler_data.db"))
