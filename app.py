from __future__ import annotations

import json
import math
import os
import threading
import time
import unicodedata
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import requests
from flask import Flask, jsonify, render_template, request
from flask.typing import ResponseReturnValue
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
MAX_MAX_PAGES = 1000
DEFAULT_RESPECT_ROBOTS = True
STAMMDATEN_PATH = Path("stammdaten.json")
stammdaten_lock = threading.Lock()
SETTINGS_PATH = Path("settings.json")
settings_lock = threading.Lock()
SYNONYM_CACHE_PATH = Path("synonym_cache.json")
synonym_cache_lock = threading.Lock()
KEYWORD_FINDER_CACHE_PATH = Path("keyword_finder_cache.json")
keyword_finder_cache_lock = threading.Lock()
OPENAI_CHAT_COMPLETIONS_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODELS_URL = "https://api.openai.com/v1/models?limit=1"
OPENAI_MODEL_NAME = "gpt-4o-mini"
OPENAI_TIMEOUT = 20
DEFAULT_SYNONYM_PROMPT = (
    "Du bist ein Synonym-Experte für deutschsprachige Websuche und spezialisiert auf Synonyme, die sich auf Webseiten "
    "von Schulen in Baden Württemberg finden. Denke dabei vor allem an Keywords, die so banal sind, dass man sie schnell "
    "vergisst. Die aber häufig vorkommen. Gib ausschließlich eine JSON-Liste mit gebräuchlichen Synonymen, nahen Begriffen "
    "und typischen Kurzformen des folgenden Begriffs zurück. Keine Erklärungen, keine weiteren Wörter. Vermeide Homonyme "
    "und irrelevante Bedeutungen."
)
DEFAULT_KEYWORD_FINDER_PROMPT = (
    "Du kombinierst Vorschläge des Google Keyword Planner mit zusätzlicher Analyse. Du erhältst ein JSON-Objekt mit "
    "Keyword-Vorschlägen (einschließlich Suchvolumen, Wettbewerb und Relevanz). Gruppiere die Begriffe semantisch in "
    "3–8 Themencluster, entferne irrelevante oder doppelte Begriffe und gib pro Cluster die drei bis fünf relevantesten "
    "Begriffe zurück. Nutze kurze, aussagekräftige Cluster-Namen. Antworte ausschließlich im folgenden JSON-Format: "
    "{\"clusters\": [{\"label\": \"<Cluster>\", \"keywords\": [{\"term\": \"<Begriff>\", \"search_volume\": <Zahl|null>, "
    "\"competition\": \"<niedrig/mittel/hoch/null>\", \"relevance\": <Zahl|null>}]}]}}. Keine weiteren Texte."
)
DEFAULT_KEYWORD_FINDER_MAX_RESULTS = 150
DEFAULT_KEYWORD_FINDER_TEMPERATURE = 0.3
GOOGLE_KEYWORD_TIMEOUT = 20
GOOGLE_KEYWORD_API_URL = "https://keywordsearch.googleapis.com/v1beta1/suggestKeywords"
DEFAULT_KEYWORD_EVALUATION_PROMPT = (
    "Du bewertest deutsche Suchbegriffe für eine Websuche. "
    "Antworte ausschließlich mit einem JSON-Objekt mit den Schlüsseln "
    '"quality"' " (Werte: \"hoch\", \"mittel\" oder \"niedrig\") und "
    '"better_keywords"' " (Liste von bis zu fünf besseren oder spezifischeren deutschen Suchbegriffen). "
    "Falls es keine besseren Begriffe gibt, gib eine leere Liste zurück. Keine Erklärungen."
)
DEFAULT_SYNONYM_MAX_RESULTS = 50
DEFAULT_SYNONYM_TEMPERATURE = 0.2
AVAILABLE_OPENAI_MODELS = [
    "gpt-4o-mini",
    "gpt-4o",
    "gpt-4.1-mini",
    "gpt-4.1",
    "o4-mini",
    "gpt-3.5-turbo",
    "llama-3.1-70b",
]

api_key_lock = threading.Lock()
_api_key_value: Optional[str] = None
google_api_key_lock = threading.Lock()
_google_api_key_value: Optional[str] = None

BACKUP_DIR = Path("backups")
BACKUP_FILE = BACKUP_DIR / "auto_backup.json"
BACKUP_INTERVAL_SECONDS = 300
_backup_thread_started = False

STAMMDATEN_PRIMARY_FIELDS = [
    {"key": "jahr", "label": "Jahr", "type": "numeric"},
    {"key": "schul_id", "label": "Schul ID"},
    {"key": "status", "label": "Status"},
    {"key": "anzahl_aussenstellen", "label": "Anzahl Außenstellen", "type": "numeric"},
    {"key": "rb", "label": "RB"},
    {"key": "kkz", "label": "KKZ"},
    {"key": "rkz", "label": "RKZ"},
    {"key": "schulname", "label": "Schulname"},
    {"key": "strasse", "label": "Straße"},
    {"key": "plz", "label": "PLZ"},
    {"key": "ort", "label": "Ort"},
    {"key": "telefon", "label": "Telefon"},
    {"key": "fax", "label": "Fax"},
    {"key": "homepage", "label": "URL"},
    {"key": "schueler_gesamt", "label": "Schüler/-innen insgesamt", "type": "numeric"},
    {"key": "klassen_gesamt", "label": "Klassen insgesamt", "type": "numeric"},
]

STAMMDATEN_BOOLEAN_FIELDS = [
    {"key": "angebot_grundschulen", "label": "12 – Grundschulen", "type": "bool"},
    {"key": "angebot_werkreal_hauptschulen", "label": "14 – Werkreal-/Hauptschulen", "type": "bool"},
    {"key": "angebot_realschulen", "label": "15 – Realschulen", "type": "bool"},
    {"key": "angebot_gymnasien", "label": "16 – Gymnasien", "type": "bool"},
    {"key": "angebot_schulen_besonderer_art", "label": "17 – Schulen besonderer Art", "type": "bool"},
    {"key": "angebot_freie_waldorfschulen", "label": "19 – Freie Waldorfschulen", "type": "bool"},
    {"key": "angebot_gemeinschaftsschule_sek1", "label": "21 – Gemeinschaftsschulen – Sekundarstufe I", "type": "bool"},
    {"key": "angebot_gemeinschaftsschule_sek2", "label": "21 – Gemeinschaftsschulen – Sekundarstufe II", "type": "bool"},
    {"key": "angebot_berufsschulen", "label": "31 – Berufsschulen", "type": "bool"},
    {"key": "angebot_berufsfachschulen", "label": "32 – Berufsfachschulen", "type": "bool"},
    {"key": "angebot_berufskollegs", "label": "33 – Berufskollegs", "type": "bool"},
    {"key": "angebot_berufsoberschulen", "label": "34 – Berufsoberschulen", "type": "bool"},
    {"key": "angebot_fachschulen", "label": "35 – Fachschulen", "type": "bool"},
    {"key": "angebot_gesundheitswesen", "label": "35 – Schulen für Berufe des Gesundheitswesens", "type": "bool"},
    {"key": "angebot_berufliche_gymnasien", "label": "36 – Berufliche Gymnasien", "type": "bool"},
    {"key": "angebot_sbbz", "label": "51 – Sonderpädagogische Bildungs- und Beratungszentren", "type": "bool"},
    {"key": "angebot_sonderberufsschulen", "label": "52 – Sonderberufsschulen", "type": "bool"},
    {"key": "angebot_zweiter_bildungsweg", "label": "Zweiter Bildungsweg", "type": "bool"},
]

STAMMDATEN_FIELDS = STAMMDATEN_PRIMARY_FIELDS + STAMMDATEN_BOOLEAN_FIELDS
STAMMDATEN_NUMERIC_KEYS = {field["key"] for field in STAMMDATEN_FIELDS if field.get("type") == "numeric"}
STAMMDATEN_BOOLEAN_KEYS = {field["key"] for field in STAMMDATEN_FIELDS if field.get("type") == "bool"}
LEGACY_STAMMDATEN_ALIASES = {
    "schulname": ["name"],
    "schueler_gesamt": ["anzahl_schueler"],
    "klassen_gesamt": ["anzahl_klassen"],
    "homepage": ["homepage", "url"],
    "status": ["traeger"],
    "ort": ["ort"],
}


class OpenAIIntegrationError(Exception):
    """Raised when the OpenAI integration cannot provide synonyms."""


class KeywordPlannerError(Exception):
    """Raised when keyword planner suggestions cannot be fetched."""


def _default_synonym_settings() -> Dict[str, object]:
    return {
        "prompt": DEFAULT_SYNONYM_PROMPT,
        "max_results": DEFAULT_SYNONYM_MAX_RESULTS,
        "temperature": DEFAULT_SYNONYM_TEMPERATURE,
        "model": OPENAI_MODEL_NAME,
    }


def _read_settings_unlocked() -> Dict[str, object]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        with SETTINGS_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (json.JSONDecodeError, OSError):
        return {}
    if isinstance(data, dict):
        return data
    return {}


def load_settings_data() -> Dict[str, object]:
    with settings_lock:
        return dict(_read_settings_unlocked())


def save_settings_data(data: Dict[str, object]) -> None:
    with settings_lock:
        try:
            with SETTINGS_PATH.open("w", encoding="utf-8") as handle:
                json.dump(data, handle, ensure_ascii=False, indent=2)
        except OSError:
            pass


def get_synonym_defaults() -> Dict[str, object]:
    data = load_settings_data()
    defaults = _default_synonym_settings()
    result = {}
    result.update(defaults)
    settings_synonyms = data.get("synonym_defaults") if isinstance(data, dict) else {}
    if isinstance(settings_synonyms, dict):
        prompt = str(settings_synonyms.get("prompt", "")).strip()
        if prompt:
            result["prompt"] = prompt
        max_results = settings_synonyms.get("max_results")
        if isinstance(max_results, int) and 1 <= max_results <= 150:
            result["max_results"] = max_results
        temperature = settings_synonyms.get("temperature")
        try:
            temp_value = float(temperature)
        except (TypeError, ValueError):
            temp_value = None
        if temp_value is not None and 0 <= temp_value <= 2:
            result["temperature"] = temp_value
        model = str(settings_synonyms.get("model", "")).strip()
        if model:
            result["model"] = model
    return result


def update_synonym_defaults(values: Dict[str, object]) -> Dict[str, object]:
    current = load_settings_data()
    defaults = _default_synonym_settings()
    prompt = str(values.get("prompt", "")).strip() or defaults["prompt"]
    max_results_raw = values.get("max_results")
    try:
        max_results = int(max_results_raw)
    except (TypeError, ValueError):
        max_results = defaults["max_results"]
    if max_results < 1:
        max_results = 1
    elif max_results > 150:
        max_results = 150
    temperature_raw = values.get("temperature")
    try:
        temperature = float(temperature_raw)
    except (TypeError, ValueError):
        temperature = defaults["temperature"]
    if temperature < 0:
        temperature = 0.0
    elif temperature > 2:
        temperature = 2.0
    model = str(values.get("model", "")).strip() or defaults["model"]

    current.setdefault("synonym_defaults", {})
    current["synonym_defaults"] = {
        "prompt": prompt,
        "max_results": max_results,
        "temperature": temperature,
        "model": model,
    }
    save_settings_data(current)
    return current["synonym_defaults"]


def _default_keyword_finder_settings() -> Dict[str, object]:
    return {
        "prompt": DEFAULT_KEYWORD_FINDER_PROMPT,
        "max_results": DEFAULT_KEYWORD_FINDER_MAX_RESULTS,
        "temperature": DEFAULT_KEYWORD_FINDER_TEMPERATURE,
        "model": OPENAI_MODEL_NAME,
    }


def get_keyword_finder_defaults() -> Dict[str, object]:
    data = load_settings_data()
    defaults = _default_keyword_finder_settings()
    result = dict(defaults)
    if isinstance(data, dict):
        stored = data.get("keyword_finder_defaults")
        if isinstance(stored, dict):
            prompt = str(stored.get("prompt", "")).strip()
            if prompt:
                result["prompt"] = prompt
            max_results = stored.get("max_results")
            if isinstance(max_results, int):
                result["max_results"] = max(1, min(max_results, DEFAULT_KEYWORD_FINDER_MAX_RESULTS))
            temperature = stored.get("temperature")
            try:
                temp_value = float(temperature)
            except (TypeError, ValueError):
                temp_value = result["temperature"]
            else:
                temp_value = max(0.0, min(temp_value, 2.0))
            result["temperature"] = temp_value
            model_value = str(stored.get("model", "")).strip()
            if model_value:
                result["model"] = model_value
    return result


def update_keyword_finder_defaults(values: Dict[str, object]) -> Dict[str, object]:
    current = load_settings_data()
    defaults = _default_keyword_finder_settings()

    prompt = str(values.get("prompt", "")).strip() or defaults["prompt"]
    try:
        max_results = int(values.get("max_results", defaults["max_results"]))
    except (TypeError, ValueError):
        max_results = defaults["max_results"]
    max_results = max(1, min(max_results, DEFAULT_KEYWORD_FINDER_MAX_RESULTS))

    try:
        temperature = float(values.get("temperature", defaults["temperature"]))
    except (TypeError, ValueError):
        temperature = defaults["temperature"]
    temperature = max(0.0, min(temperature, 2.0))

    model = str(values.get("model", "")).strip() or defaults["model"]

    current.setdefault("keyword_finder_defaults", {})
    current["keyword_finder_defaults"] = {
        "prompt": prompt,
        "max_results": max_results,
        "temperature": temperature,
        "model": model,
    }
    save_settings_data(current)
    return current["keyword_finder_defaults"]


def _default_crawl_settings() -> Dict[str, int]:
    return {
        "max_pages": MAX_PAGES_DEFAULT,
        "concurrency": DEFAULT_CONCURRENCY,
    }


def get_crawl_defaults() -> Dict[str, int]:
    data = load_settings_data()
    defaults = _default_crawl_settings()
    result = dict(defaults)
    if isinstance(data, dict):
        stored = data.get("crawl_defaults")
        if isinstance(stored, dict):
            max_pages = stored.get("max_pages")
            if isinstance(max_pages, int):
                result["max_pages"] = max(1, min(max_pages, MAX_MAX_PAGES))
            concurrency = stored.get("concurrency")
            if isinstance(concurrency, int):
                result["concurrency"] = max(1, min(concurrency, MAX_CONCURRENCY))
    return result


def update_crawl_defaults(values: Dict[str, object]) -> Dict[str, int]:
    current = load_settings_data()
    defaults = _default_crawl_settings()

    try:
        max_pages = int(values.get("max_pages", defaults["max_pages"]))
    except (TypeError, ValueError):
        max_pages = defaults["max_pages"]
    max_pages = max(1, min(max_pages, MAX_MAX_PAGES))

    try:
        concurrency = int(values.get("concurrency", defaults["concurrency"]))
    except (TypeError, ValueError):
        concurrency = defaults["concurrency"]
    concurrency = max(1, min(concurrency, MAX_CONCURRENCY))

    current.setdefault("crawl_defaults", {})
    current["crawl_defaults"] = {
        "max_pages": max_pages,
        "concurrency": concurrency,
    }
    save_settings_data(current)
    return current["crawl_defaults"]


def _read_synonym_cache_unlocked() -> Dict[str, object]:
    if not SYNONYM_CACHE_PATH.exists():
        return {}
    try:
        with SYNONYM_CACHE_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (json.JSONDecodeError, OSError):
        return {}
    if isinstance(data, dict):
        return data
    return {}


def _write_synonym_cache_unlocked(data: Dict[str, object]) -> None:
    try:
        with SYNONYM_CACHE_PATH.open("w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_synonym_cache() -> Dict[str, object]:
    with synonym_cache_lock:
        return dict(_read_synonym_cache_unlocked())


def update_synonym_cache_entry(key: str, synonyms: List[str]) -> None:
    with synonym_cache_lock:
        cache = _read_synonym_cache_unlocked()
        cache[key] = {
            "synonyms": list(synonyms),
            "updated_at": datetime.utcnow().isoformat() + "Z",
        }
        _write_synonym_cache_unlocked(cache)


def replace_synonym_cache(data: Dict[str, object]) -> None:
    with synonym_cache_lock:
        payload = data if isinstance(data, dict) else {}
        _write_synonym_cache_unlocked(payload)


def get_synonyms_from_cache(key: str) -> Optional[List[str]]:
    cache = load_synonym_cache()
    entry = cache.get(key) if isinstance(cache, dict) else None
    if isinstance(entry, dict):
        values = entry.get("synonyms")
        if isinstance(values, list):
            result: List[str] = []
            for item in values:
                text = str(item).strip()
                if text:
                    result.append(text)
            return result
    return None


def synonym_cache_key(keyword: str, prompt: str, model: str, temperature: float) -> str:
    payload = {
        "keyword": keyword.strip().lower(),
        "prompt": prompt.strip(),
        "model": model.strip(),
        "temperature": round(float(temperature), 3),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def _read_keyword_finder_cache_unlocked() -> Dict[str, object]:
    if not KEYWORD_FINDER_CACHE_PATH.exists():
        return {}
    try:
        with KEYWORD_FINDER_CACHE_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (json.JSONDecodeError, OSError):
        return {}
    if isinstance(data, dict):
        return data
    return {}


def _write_keyword_finder_cache_unlocked(data: Dict[str, object]) -> None:
    try:
        with KEYWORD_FINDER_CACHE_PATH.open("w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
    except OSError:
        pass


def load_keyword_finder_cache() -> Dict[str, object]:
    with keyword_finder_cache_lock:
        return dict(_read_keyword_finder_cache_unlocked())


def update_keyword_finder_cache_entry(key: str, payload: Dict[str, object]) -> None:
    with keyword_finder_cache_lock:
        cache = _read_keyword_finder_cache_unlocked()
        cache[key] = {
            "data": payload,
            "updated_at": datetime.utcnow().isoformat() + "Z",
        }
        _write_keyword_finder_cache_unlocked(cache)


def replace_keyword_finder_cache(data: Dict[str, object]) -> None:
    with keyword_finder_cache_lock:
        payload = data if isinstance(data, dict) else {}
        _write_keyword_finder_cache_unlocked(payload)


def get_keyword_finder_cache_entry(key: str) -> Optional[Dict[str, object]]:
    cache = load_keyword_finder_cache()
    entry = cache.get(key) if isinstance(cache, dict) else None
    if isinstance(entry, dict):
        data = entry.get("data")
        if isinstance(data, dict):
            return data
    return None


def keyword_finder_cache_key(
    keyword: str,
    prompt: str,
    model: str,
    temperature: float,
    max_results: int,
) -> str:
    payload = {
        "keyword": keyword.strip().lower(),
        "prompt": prompt.strip(),
        "model": model.strip(),
        "temperature": round(float(temperature), 3),
        "max_results": int(max_results),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def set_api_key(value: str) -> None:
    sanitized = value.strip()
    with api_key_lock:
        global _api_key_value
        _api_key_value = sanitized or None
        if sanitized:
            os.environ["OPENAI_API_KEY"] = sanitized
        else:
            os.environ.pop("OPENAI_API_KEY", None)


def clear_api_key() -> None:
    with api_key_lock:
        global _api_key_value
        _api_key_value = None
        os.environ.pop("OPENAI_API_KEY", None)


def get_api_key() -> Optional[str]:
    with api_key_lock:
        env_value = os.environ.get("OPENAI_API_KEY", "").strip()
        if env_value:
            global _api_key_value
            _api_key_value = env_value
            return env_value
        return _api_key_value


def has_api_key() -> bool:
    return get_api_key() is not None


def set_keyword_planner_key(value: str) -> None:
    sanitized = value.strip()
    with google_api_key_lock:
        global _google_api_key_value
        _google_api_key_value = sanitized or None
        if sanitized:
            os.environ["GOOGLE_KEYWORD_PLANNER_KEY"] = sanitized
        else:
            os.environ.pop("GOOGLE_KEYWORD_PLANNER_KEY", None)


def clear_keyword_planner_key() -> None:
    with google_api_key_lock:
        global _google_api_key_value
        _google_api_key_value = None
        os.environ.pop("GOOGLE_KEYWORD_PLANNER_KEY", None)


def get_keyword_planner_key() -> Optional[str]:
    with google_api_key_lock:
        env_value = os.environ.get("GOOGLE_KEYWORD_PLANNER_KEY", "").strip()
        if env_value:
            global _google_api_key_value
            _google_api_key_value = env_value
            return env_value
        return _google_api_key_value


def has_keyword_planner_key() -> bool:
    return get_keyword_planner_key() is not None


def fetch_synonyms_from_openai(
    keyword: str,
    api_key: str,
    *,
    prompt: Optional[str] = None,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_results: Optional[int] = None,
) -> List[str]:
    """Request synonyms for a single keyword from the OpenAI API."""

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    prompt_text = (prompt or "").strip() or DEFAULT_SYNONYM_PROMPT
    model_name = (model or "").strip() or OPENAI_MODEL_NAME
    try:
        temperature_value = float(temperature) if temperature is not None else DEFAULT_SYNONYM_TEMPERATURE
    except (TypeError, ValueError):
        temperature_value = DEFAULT_SYNONYM_TEMPERATURE
    max_results_value: Optional[int]
    if max_results is None:
        max_results_value = None
    else:
        try:
            max_results_value = int(max_results)
        except (TypeError, ValueError):
            max_results_value = None
    if max_results_value is not None:
        if max_results_value < 1:
            max_results_value = 1
        elif max_results_value > 150:
            max_results_value = 150
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": prompt_text},
            {"role": "user", "content": f'Begriff: "{keyword}"'},
        ],
        "temperature": temperature_value,
        "max_tokens": 200,
    }

    try:
        response = requests.post(
            OPENAI_CHAT_COMPLETIONS_URL,
            headers=headers,
            json=payload,
            timeout=OPENAI_TIMEOUT,
        )
    except requests.RequestException as exc:  # pragma: no cover - network failure
        raise OpenAIIntegrationError(
            f'Synonymerweiterung für "{keyword}" nicht möglich: Keine Verbindung zur OpenAI-API.'
        ) from exc

    if response.status_code >= 400:
        raise OpenAIIntegrationError(
            f'Synonymerweiterung für "{keyword}" nicht möglich: OpenAI-API meldet Status {response.status_code}.'
        )

    try:
        data = response.json()
    except ValueError as exc:  # pragma: no cover - invalid JSON
        raise OpenAIIntegrationError(
            f'Synonymerweiterung für "{keyword}" nicht möglich: Antwort konnte nicht gelesen werden.'
        ) from exc

    content = ""
    choices = data.get("choices") if isinstance(data, dict) else None
    if choices and isinstance(choices, list) and choices:
        message = choices[0].get("message") or {}
        if isinstance(message, dict):
            content = message.get("content", "")

    if not content:
        raise OpenAIIntegrationError(
            f'Synonymerweiterung für "{keyword}" nicht möglich: Antwort enthielt keine Vorschläge.'
        )

    content = content.strip()
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError:
        start_idx = content.find("[")
        end_idx = content.rfind("]")
        if start_idx == -1 or end_idx == -1 or end_idx <= start_idx:
            raise OpenAIIntegrationError(
                f'Synonymerweiterung für "{keyword}" nicht möglich: Antwort war nicht im JSON-Format.'
            )
        try:
            parsed = json.loads(content[start_idx : end_idx + 1])
        except json.JSONDecodeError as exc:
            raise OpenAIIntegrationError(
                f'Synonymerweiterung für "{keyword}" nicht möglich: Antwort war nicht im JSON-Format.'
            ) from exc

    if not isinstance(parsed, list):
        raise OpenAIIntegrationError(
            f'Synonymerweiterung für "{keyword}" nicht möglich: Antwort enthielt keine Begriffsliste.'
        )

    synonyms: List[str] = []
    seen: Set[str] = set()
    for item in parsed:
        text = str(item).strip()
        if not text:
            continue
        lowered = text.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        synonyms.append(text)
        if max_results_value is not None and len(synonyms) >= max_results_value:
            break

    return synonyms


def fetch_keyword_planner_suggestions(
    keyword: str,
    api_key: str,
    *,
    max_results: int = DEFAULT_KEYWORD_FINDER_MAX_RESULTS,
) -> List[Dict[str, object]]:
    """Fetch keyword suggestions from the Google Keyword Planner."""

    if not api_key:
        raise KeywordPlannerError("Kein Keyword-Planner-Schlüssel hinterlegt. Bitte in den Einstellungen speichern.")

    try:
        max_value = int(max_results)
    except (TypeError, ValueError):
        max_value = DEFAULT_KEYWORD_FINDER_MAX_RESULTS
    max_value = max(1, min(max_value, DEFAULT_KEYWORD_FINDER_MAX_RESULTS))

    params = {"key": api_key}
    payload = {
        "query": keyword,
        "languageCode": "de-DE",
        "maxSuggestions": max_value,
        "regionCode": "DE",
    }

    try:
        response = requests.post(
            GOOGLE_KEYWORD_API_URL,
            params=params,
            json=payload,
            timeout=GOOGLE_KEYWORD_TIMEOUT,
        )
    except requests.RequestException as exc:
        raise KeywordPlannerError("Keyword-Planner-Anfrage fehlgeschlagen (Netzwerkfehler).") from exc

    if response.status_code >= 400:
        raise KeywordPlannerError(
            f"Keyword-Planner-Anfrage fehlgeschlagen (Status {response.status_code}). Bitte Schlüssel und Berechtigungen prüfen."
        )

    try:
        data = response.json()
    except ValueError as exc:
        raise KeywordPlannerError("Keyword-Planner-Antwort konnte nicht gelesen werden.") from exc

    suggestions_raw: List[Dict[str, object]] = []
    if isinstance(data, dict):
        for key in ("suggestions", "results", "keywordIdeas", "ideas", "items"):
            values = data.get(key)
            if isinstance(values, list):
                for item in values:
                    if isinstance(item, dict):
                        suggestions_raw.append(item)
        # Some APIs nest results deeper
        if not suggestions_raw:
            keyword_ideas = data.get("keyword_ideas")
            if isinstance(keyword_ideas, list):
                for item in keyword_ideas:
                    if isinstance(item, dict):
                        suggestions_raw.append(item)

    if not suggestions_raw:
        raise KeywordPlannerError("Keyword-Planner lieferte keine Vorschläge.")

    suggestions: List[Dict[str, object]] = []
    seen: Set[str] = set()
    for item in suggestions_raw:
        term = str(
            item.get("keyword")
            or item.get("text")
            or item.get("term")
            or item.get("phrase")
            or ""
        ).strip()
        if not term:
            continue
        lowered = term.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        search_volume = item.get("searchVolume") or item.get("search_volume") or item.get("avgMonthlySearches")
        competition = item.get("competition") or item.get("competitionIndex") or item.get("competition_index")
        relevance = item.get("relevance") or item.get("relevanceScore") or item.get("relevance_score")
        suggestions.append(
            {
                "term": term,
                "search_volume": search_volume,
                "competition": competition,
                "relevance": relevance,
            }
        )
        if len(suggestions) >= max_value:
            break

    if not suggestions:
        raise KeywordPlannerError("Keyword-Planner lieferte keine verwertbaren Begriffe.")

    return suggestions


def _normalise_competition(value: object) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if value <= 0:
            return "niedrig"
        if value < 0.5:
            return "niedrig"
        if value < 0.8:
            return "mittel"
        return "hoch"
    text = str(value).strip().lower()
    if not text:
        return None
    if text in {"low", "niedrig"}:
        return "niedrig"
    if text in {"medium", "mittel"}:
        return "mittel"
    if text in {"high", "hoch"}:
        return "hoch"
    return None


def _build_fallback_clusters(suggestions: List[Dict[str, object]]) -> Dict[str, object]:
    cleaned = []
    for item in suggestions:
        term = str(item.get("term", "")).strip()
        if not term:
            continue
        cleaned.append(
            {
                "term": term,
                "search_volume": item.get("search_volume"),
                "competition": _normalise_competition(item.get("competition")),
                "relevance": item.get("relevance"),
            }
        )
    return {
        "clusters": [
            {
                "label": "Weitere Vorschläge",
                "keywords": cleaned,
            }
        ]
    }


def analyse_keyword_clusters_with_openai(
    keyword: str,
    suggestions: List[Dict[str, object]],
    api_key: str,
    *,
    prompt: Optional[str] = None,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
) -> Dict[str, object]:
    if not suggestions:
        return {"clusters": []}

    prompt_text = (prompt or "").strip() or DEFAULT_KEYWORD_FINDER_PROMPT
    model_name = (model or "").strip() or OPENAI_MODEL_NAME
    try:
        temperature_value = float(temperature) if temperature is not None else DEFAULT_KEYWORD_FINDER_TEMPERATURE
    except (TypeError, ValueError):
        temperature_value = DEFAULT_KEYWORD_FINDER_TEMPERATURE

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    user_payload = {
        "keyword": keyword,
        "suggestions": suggestions,
    }
    request_payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": prompt_text},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ],
        "temperature": temperature_value,
        "max_tokens": 600,
    }

    try:
        response = requests.post(
            OPENAI_CHAT_COMPLETIONS_URL,
            headers=headers,
            json=request_payload,
            timeout=OPENAI_TIMEOUT,
        )
    except requests.RequestException as exc:
        raise OpenAIIntegrationError("LLM-Auswertung der Keyword-Vorschläge fehlgeschlagen (Netzwerkfehler).") from exc

    if response.status_code >= 400:
        raise OpenAIIntegrationError(
            f"LLM-Auswertung der Keyword-Vorschläge fehlgeschlagen (Status {response.status_code})."
        )

    try:
        data = response.json()
    except ValueError as exc:
        raise OpenAIIntegrationError("LLM-Antwort konnte nicht gelesen werden.") from exc

    content = ""
    choices = data.get("choices") if isinstance(data, dict) else None
    if isinstance(choices, list) and choices:
        message = choices[0].get("message") if isinstance(choices[0], dict) else None
        if isinstance(message, dict):
            content = message.get("content", "")

    if not content:
        raise OpenAIIntegrationError("LLM lieferte keine verwertbaren Daten.")

    content = content.strip()
    parsed: Optional[Dict[str, object]] = None
    try:
        parsed_obj = json.loads(content)
    except json.JSONDecodeError:
        start_idx = content.find("{")
        end_idx = content.rfind("}")
        if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
            try:
                parsed_obj = json.loads(content[start_idx : end_idx + 1])
            except json.JSONDecodeError as exc:
                raise OpenAIIntegrationError("LLM-Antwort war nicht im JSON-Format.") from exc
        else:
            raise OpenAIIntegrationError("LLM-Antwort war nicht im JSON-Format.")
    else:
        parsed = parsed_obj if isinstance(parsed_obj, dict) else None

    if parsed is None:
        raise OpenAIIntegrationError("LLM-Antwort enthielt keine strukturierten Daten.")

    clusters_raw = parsed.get("clusters") if isinstance(parsed, dict) else None
    clusters: List[Dict[str, object]] = []
    if isinstance(clusters_raw, list):
        for cluster in clusters_raw:
            if not isinstance(cluster, dict):
                continue
            label = str(cluster.get("label", "")).strip() or "Weitere Vorschläge"
            keywords_raw = cluster.get("keywords")
            keywords: List[Dict[str, object]] = []
            if isinstance(keywords_raw, list):
                for keyword_entry in keywords_raw:
                    if not isinstance(keyword_entry, dict):
                        continue
                    term = str(keyword_entry.get("term", "")).strip()
                    if not term:
                        continue
                    keywords.append(
                        {
                            "term": term,
                            "search_volume": keyword_entry.get("search_volume"),
                            "competition": _normalise_competition(keyword_entry.get("competition")),
                            "relevance": keyword_entry.get("relevance"),
                        }
                    )
            if keywords:
                clusters.append({"label": label, "keywords": keywords})

    if not clusters:
        return _build_fallback_clusters(suggestions)

    return {"clusters": clusters}


def expand_keywords_with_ai(
    keywords: List[str],
    *,
    enabled: bool,
    api_key: Optional[str],
) -> Tuple[List[str], List[Dict[str, List[str]]], List[Dict[str, str]], bool]:
    """Return expanded keywords, grouping information and user messages."""

    cleaned_keywords = [kw.strip() for kw in keywords if kw and kw.strip()]
    expanded: List[str] = []
    seen_lower: Set[str] = set()
    for kw in cleaned_keywords:
        lowered = kw.lower()
        if lowered not in seen_lower:
            expanded.append(kw)
            seen_lower.add(lowered)

    keyword_groups: List[Dict[str, List[str]]] = []
    messages: List[Dict[str, str]] = []
    defaults = get_synonym_defaults()

    if not cleaned_keywords:
        return expanded, keyword_groups, messages, False

    if not enabled:
        for kw in cleaned_keywords:
            keyword_groups.append({"original": kw, "additional": []})
        return expanded, keyword_groups, messages, False

    if not api_key:
        messages.append(
            {
                "category": "warning",
                "text": "OpenAI-Schlüssel erforderlich. Bitte unter Einstellungen → API-Schlüssel hinterlegen.",
            }
        )
        for kw in cleaned_keywords:
            keyword_groups.append({"original": kw, "additional": []})
        return expanded, keyword_groups, messages, False

    message_texts: Set[str] = set()
    synonyms_expanded = False

    for kw in cleaned_keywords:
        additional: List[str] = []
        try:
            synonyms = fetch_synonyms_from_openai(
                kw,
                api_key,
                prompt=defaults.get("prompt"),
                model=str(defaults.get("model") or OPENAI_MODEL_NAME),
                temperature=defaults.get("temperature"),
                max_results=defaults.get("max_results"),
            )
        except OpenAIIntegrationError as exc:
            text = str(exc)
            if text and text not in message_texts:
                messages.append({"category": "warning", "text": text})
                message_texts.add(text)
            keyword_groups.append({"original": kw, "additional": []})
            continue

        for synonym in synonyms:
            lowered = synonym.lower()
            if lowered in seen_lower:
                continue
            seen_lower.add(lowered)
            expanded.append(synonym)
            additional.append(synonym)

        if additional:
            synonyms_expanded = True

        keyword_groups.append({"original": kw, "additional": additional})

    if synonyms_expanded:
        total_added = sum(len(group["additional"]) for group in keyword_groups)
        info_text = (
            "Synonymerweiterung aktiv: "
            f"{total_added} zusätzliche Begriffe werden durchsucht."
        )
        messages.append({"category": "info", "text": info_text})

    return expanded, keyword_groups, messages, synonyms_expanded


def evaluate_keyword_with_openai(
    keyword: str,
    api_key: str,
    *,
    prompt: Optional[str] = None,
) -> Dict[str, object]:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    prompt_text = (prompt or "").strip() or DEFAULT_KEYWORD_EVALUATION_PROMPT
    payload = {
        "model": OPENAI_MODEL_NAME,
        "messages": [
            {"role": "system", "content": prompt_text},
            {"role": "user", "content": f'Begriff: "{keyword}"'},
        ],
        "temperature": 0.2,
        "max_tokens": 200,
    }

    try:
        response = requests.post(
            OPENAI_CHAT_COMPLETIONS_URL,
            headers=headers,
            json=payload,
            timeout=OPENAI_TIMEOUT,
        )
    except requests.RequestException as exc:  # pragma: no cover - network failure
        raise OpenAIIntegrationError(
            f'Bewertung für "{keyword}" nicht möglich: Keine Verbindung zur OpenAI-API.'
        ) from exc

    if response.status_code >= 400:
        raise OpenAIIntegrationError(
            f'Bewertung für "{keyword}" nicht möglich: OpenAI-API meldet Status {response.status_code}.'
        )

    try:
        response_payload = response.json()
    except ValueError as exc:  # pragma: no cover - invalid JSON
        raise OpenAIIntegrationError(
            f'Bewertung für "{keyword}" nicht möglich: Antwort konnte nicht gelesen werden.'
        ) from exc

    choices = response_payload.get("choices") if isinstance(response_payload, dict) else None
    content = ""
    if isinstance(choices, list) and choices:
        message = choices[0].get("message") if isinstance(choices[0], dict) else None
        if isinstance(message, dict):
            content = str(message.get("content", ""))

    if not content:
        raise OpenAIIntegrationError(
            f'Bewertung für "{keyword}" nicht möglich: Antwort enthielt keine Daten.'
        )

    content = content.strip()
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError:
        start_idx = content.find("{")
        end_idx = content.rfind("}")
        if start_idx == -1 or end_idx == -1 or end_idx <= start_idx:
            raise OpenAIIntegrationError(
                f'Bewertung für "{keyword}" nicht möglich: Antwort war nicht im JSON-Format.'
            )
        try:
            parsed = json.loads(content[start_idx : end_idx + 1])
        except json.JSONDecodeError as exc:
            raise OpenAIIntegrationError(
                f'Bewertung für "{keyword}" nicht möglich: Antwort war nicht im JSON-Format.'
            ) from exc

    if not isinstance(parsed, dict):
        raise OpenAIIntegrationError(
            f'Bewertung für "{keyword}" nicht möglich: Antwort enthielt kein Objekt.'
        )

    quality_raw = str(parsed.get("quality", "")).strip()
    quality = quality_raw or "unbewertet"

    better_raw = (
        parsed.get("better_keywords")
        or parsed.get("betterTerms")
        or parsed.get("better_terms")
        or parsed.get("bessere_suchbegriffe")
        or parsed.get("alternativen")
    )
    better_keywords: List[str] = []
    if isinstance(better_raw, list):
        seen: Set[str] = set()
        for item in better_raw:
            text = str(item).strip()
            if not text:
                continue
            lowered = text.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            better_keywords.append(text)
            if len(better_keywords) >= 8:
                break

    return {
        "keyword": keyword,
        "quality": quality,
        "better_keywords": better_keywords,
    }


def test_openai_api_key(api_key: str) -> None:
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        response = requests.get(OPENAI_MODELS_URL, headers=headers, timeout=OPENAI_TIMEOUT)
    except requests.RequestException as exc:  # pragma: no cover - network failure
        raise OpenAIIntegrationError("Die Verbindung zur OpenAI-API konnte nicht hergestellt werden.") from exc

    if response.status_code >= 400:
        raise OpenAIIntegrationError(
            f"OpenAI-API meldet einen Fehler (Status {response.status_code})."
        )


@app.post("/synonyms")
def generate_synonyms() -> ResponseReturnValue:
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "OpenAI-Schlüssel erforderlich."}), 400

    payload = request.get_json(silent=True) or {}
    if not isinstance(payload, dict):
        return jsonify({"error": "Ungültige Anfrage."}), 400

    keyword = str(payload.get("keyword", "")).strip()
    prompt = payload.get("prompt")
    model = payload.get("model")
    max_results = payload.get("max_results")
    temperature = payload.get("temperature")
    defaults = get_synonym_defaults()

    if not keyword:
        return jsonify({"error": "Bitte geben Sie einen Suchbegriff an."}), 400

    prompt_text = (prompt or "").strip() or str(defaults.get("prompt", DEFAULT_SYNONYM_PROMPT))
    model_name = str(model or defaults.get("model") or OPENAI_MODEL_NAME).strip() or OPENAI_MODEL_NAME
    try:
        temperature_value = float(temperature) if temperature is not None else float(defaults.get("temperature", DEFAULT_SYNONYM_TEMPERATURE))
    except (TypeError, ValueError):
        temperature_value = float(defaults.get("temperature", DEFAULT_SYNONYM_TEMPERATURE))
    if temperature_value < 0:
        temperature_value = 0.0
    elif temperature_value > 2:
        temperature_value = 2.0
    try:
        max_results_value = int(max_results) if max_results is not None else int(defaults.get("max_results", DEFAULT_SYNONYM_MAX_RESULTS))
    except (TypeError, ValueError):
        max_results_value = int(defaults.get("max_results", DEFAULT_SYNONYM_MAX_RESULTS))
    if max_results_value < 1:
        max_results_value = 1
    elif max_results_value > 150:
        max_results_value = 150

    cache_key = synonym_cache_key(keyword, prompt_text, model_name, temperature_value)
    cached = get_synonyms_from_cache(cache_key)
    if cached is not None:
        return jsonify(
            {
                "keyword": keyword,
                "synonyms": cached[:max_results_value],
                "cached": True,
            }
        )

    try:
        synonyms = fetch_synonyms_from_openai(
            keyword,
            api_key,
            prompt=prompt_text,
            model=model_name,
            temperature=temperature_value,
            max_results=max_results_value,
        )
    except OpenAIIntegrationError as exc:
        return jsonify({"error": str(exc)}), 502

    update_synonym_cache_entry(cache_key, synonyms)

    return jsonify(
        {
            "keyword": keyword,
            "synonyms": synonyms,
            "cached": False,
        }
    )


@app.post("/related-keywords")
def generate_related_keywords() -> ResponseReturnValue:
    payload = request.get_json(silent=True) or {}
    keyword = str(payload.get("keyword", "")).strip()
    if not keyword:
        return jsonify({"error": "Bitte geben Sie ein Suchwort an."}), 400

    defaults = get_keyword_finder_defaults()
    prompt = str(payload.get("prompt", "")).strip() or defaults.get("prompt", DEFAULT_KEYWORD_FINDER_PROMPT)
    model = str(payload.get("model", "")).strip() or defaults.get("model", OPENAI_MODEL_NAME)

    try:
        temperature = float(payload.get("temperature", defaults.get("temperature", DEFAULT_KEYWORD_FINDER_TEMPERATURE)))
    except (TypeError, ValueError):
        temperature = defaults.get("temperature", DEFAULT_KEYWORD_FINDER_TEMPERATURE)
    temperature = max(0.0, min(temperature, 2.0))

    try:
        max_results = int(payload.get("max_results", defaults.get("max_results", DEFAULT_KEYWORD_FINDER_MAX_RESULTS)))
    except (TypeError, ValueError):
        max_results = defaults.get("max_results", DEFAULT_KEYWORD_FINDER_MAX_RESULTS)
    max_results = max(1, min(max_results, DEFAULT_KEYWORD_FINDER_MAX_RESULTS))

    google_key = get_keyword_planner_key()
    if not google_key:
        return jsonify({"error": "Kein Keyword-Planner-Schlüssel hinterlegt. Bitte in den Einstellungen speichern."}), 400

    try:
        suggestions = fetch_keyword_planner_suggestions(keyword, google_key, max_results=max_results)
    except KeywordPlannerError as exc:
        return jsonify({"error": str(exc)}), 502

    cache_key = keyword_finder_cache_key(keyword, prompt, model, temperature, max_results)
    cached_entry = get_keyword_finder_cache_entry(cache_key)
    current_terms = [item.get("term", "").strip().lower() for item in suggestions]
    messages: List[str] = []

    if cached_entry and isinstance(cached_entry, dict):
        cached_suggestions = cached_entry.get("suggestions")
        if isinstance(cached_suggestions, list):
            cached_terms = [
                str(item.get("term", "")).strip().lower()
                for item in cached_suggestions
                if isinstance(item, dict)
            ]
            if cached_terms != current_terms:
                cached_entry = None
        else:
            cached_entry = None

    if cached_entry:
        clusters = cached_entry.get("clusters") if isinstance(cached_entry.get("clusters"), list) else []
        analysis = cached_entry.get("analysis") if isinstance(cached_entry.get("analysis"), str) else "cache"
        return jsonify(
            {
                "keyword": keyword,
                "clusters": clusters,
                "suggestions": suggestions,
                "analysis": analysis,
                "cached": True,
                "messages": messages,
                "defaults": {
                    "prompt": prompt,
                    "model": model,
                    "temperature": temperature,
                    "max_results": max_results,
                },
            }
        )

    openai_key = get_api_key()
    analysis_mode = "fallback"
    if not openai_key:
        messages.append("OpenAI-Schlüssel fehlt – Vorschläge werden ohne KI-Gruppierung angezeigt.")
        clusters_payload = _build_fallback_clusters(suggestions)
    else:
        try:
            clusters_payload = analyse_keyword_clusters_with_openai(
                keyword,
                suggestions,
                openai_key,
                prompt=prompt,
                model=model,
                temperature=temperature,
            )
        except OpenAIIntegrationError as exc:
            messages.append(str(exc))
            clusters_payload = _build_fallback_clusters(suggestions)
        else:
            analysis_mode = "llm"

    clusters = clusters_payload.get("clusters", []) if isinstance(clusters_payload, dict) else []

    cache_payload = {
        "clusters": clusters,
        "suggestions": suggestions,
        "analysis": analysis_mode,
    }
    update_keyword_finder_cache_entry(cache_key, cache_payload)

    return jsonify(
        {
            "keyword": keyword,
            "clusters": clusters,
            "suggestions": suggestions,
            "analysis": analysis_mode,
            "cached": False,
            "messages": messages,
            "defaults": {
                "prompt": prompt,
                "model": model,
                "temperature": temperature,
                "max_results": max_results,
            },
        }
    )


@app.post("/evaluate-keywords")
def evaluate_keywords() -> ResponseReturnValue:
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "OpenAI-Schlüssel erforderlich."}), 400

    payload = request.get_json(silent=True) or {}
    if not isinstance(payload, dict):
        return jsonify({"error": "Ungültige Anfrage."}), 400

    keywords_raw = payload.get("keywords")
    prompt = payload.get("prompt")

    if not isinstance(keywords_raw, list):
        return jsonify({"error": "Bitte geben Sie eine Liste von Suchbegriffen an."}), 400

    cleaned: List[str] = []
    seen_lower: Set[str] = set()
    for item in keywords_raw:
        text = str(item).strip()
        if not text:
            continue
        lowered = text.lower()
        if lowered in seen_lower:
            continue
        seen_lower.add(lowered)
        cleaned.append(text)

    if not cleaned:
        return jsonify({"evaluations": [], "messages": []})

    evaluations: List[Dict[str, object]] = []
    messages: List[str] = []

    for keyword in cleaned:
        try:
            result = evaluate_keyword_with_openai(keyword, api_key, prompt=prompt)
        except OpenAIIntegrationError as exc:
            text = str(exc)
            if text:
                messages.append(text)
        else:
            evaluations.append(result)

    status = 200 if evaluations or not messages else 502
    return jsonify({"evaluations": evaluations, "messages": messages}), status





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


def replace_saved_searches(entries: List[Dict[str, object]]) -> None:
    with saved_search_lock:
        try:
            with SAVED_SEARCHES_PATH.open("w", encoding="utf-8") as handle:
                json.dump(entries, handle, ensure_ascii=False, indent=2)
        except OSError:
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


def _coerce_bool(value: object, *, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        return bool(int(value))
    text = str(value).strip().lower()
    if not text:
        return default
    if text in {"1", "true", "wahr", "ja", "j", "y", "aktiv", "active", "x"}:
        return True
    if text in {"0", "false", "falsch", "nein", "n", "inactive", "inaktiv"}:
        return False
    return default


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

                record: Dict[str, object] = {}
                for field in STAMMDATEN_FIELDS:
                    key = field["key"]
                    field_type = field.get("type", "text")
                    value = entry.get(key)
                    if value is None:
                        for alias in LEGACY_STAMMDATEN_ALIASES.get(key, []):
                            if alias in entry:
                                value = entry.get(alias)
                                if value is not None:
                                    break
                    if field_type == "numeric":
                        record[key] = _coerce_int(value)
                    elif field_type == "bool":
                        record[key] = _coerce_bool(value, default=False)
                    else:
                        record[key] = _coerce_str(value)

                homepage = _coerce_str(
                    record.get("homepage")
                    or entry.get("homepage")
                    or entry.get("url")
                )
                if not homepage:
                    continue
                record["homepage"] = homepage

                if not record.get("schulname"):
                    record["schulname"] = _coerce_str(entry.get("name"))

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

    expected_headers: Dict[str, Set[str]] = {}
    for field in STAMMDATEN_FIELDS:
        key = field["key"]
        variants = expected_headers.setdefault(key, set())
        variants.add(_normalize_header(field["label"]))
        for alias in LEGACY_STAMMDATEN_ALIASES.get(key, []):
            variants.add(_normalize_header(str(alias)))

    column_map: Dict[str, int] = {}
    for key, options in expected_headers.items():
        for normalized in options:
            if normalized in header_map:
                column_map[key] = header_map[normalized]
                break

    if len(column_map) < len(STAMMDATEN_FIELDS):
        for field in STAMMDATEN_FIELDS:
            column_map.setdefault(field["key"], None)

    records: List[Dict[str, object]] = []
    for raw_row in rows[1:]:
        if raw_row is None:
            continue
        values = list(raw_row)
        if not any(cell not in (None, "") for cell in values):
            continue

        record: Dict[str, object] = {}
        for field in STAMMDATEN_FIELDS:
            column_index = column_map.get(field["key"])
            cell_value = (
                values[column_index]
                if column_index is not None and column_index < len(values)
                else None
            )
            field_type = field.get("type", "text")
            if field_type == "numeric":
                record[field["key"]] = _coerce_int(cell_value)
            elif field_type == "bool":
                record[field["key"]] = _coerce_bool(cell_value, default=False)
            else:
                record[field["key"]] = _coerce_str(cell_value)

        if record.get("homepage"):
            records.append(record)

    if not records:
        raise ValueError("Es konnten keine gültigen Stammdaten gefunden werden.")

    return records


def collect_backup_snapshot() -> Dict[str, object]:
    return {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "settings": load_settings_data(),
        "stammdaten": load_stammdaten(),
        "saved_searches": load_saved_searches(),
        "synonym_cache": load_synonym_cache(),
        "keyword_finder_cache": load_keyword_finder_cache(),
    }


def perform_backup() -> None:
    snapshot = collect_backup_snapshot()
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with BACKUP_FILE.open("w", encoding="utf-8") as handle:
            json.dump(snapshot, handle, ensure_ascii=False, indent=2)
    except OSError:
        pass


def restore_from_backup() -> None:
    if not BACKUP_FILE.exists():
        return
    try:
        with BACKUP_FILE.open("r", encoding="utf-8") as handle:
            snapshot = json.load(handle)
    except (json.JSONDecodeError, OSError):
        return
    if not isinstance(snapshot, dict):
        return

    settings_data = snapshot.get("settings")
    if settings_data and (not SETTINGS_PATH.exists() or SETTINGS_PATH.stat().st_size == 0):
        if isinstance(settings_data, dict):
            save_settings_data(settings_data)

    stammdaten_data = snapshot.get("stammdaten")
    if stammdaten_data and (not STAMMDATEN_PATH.exists() or STAMMDATEN_PATH.stat().st_size == 0):
        if isinstance(stammdaten_data, list):
            save_stammdaten(stammdaten_data)

    saved_searches_data = snapshot.get("saved_searches")
    if saved_searches_data and (not SAVED_SEARCHES_PATH.exists() or SAVED_SEARCHES_PATH.stat().st_size == 0):
        if isinstance(saved_searches_data, list):
            replace_saved_searches(saved_searches_data)

    synonym_cache_data = snapshot.get("synonym_cache")
    if synonym_cache_data and (not SYNONYM_CACHE_PATH.exists() or SYNONYM_CACHE_PATH.stat().st_size == 0):
        if isinstance(synonym_cache_data, dict):
            replace_synonym_cache(synonym_cache_data)

    keyword_finder_cache_data = snapshot.get("keyword_finder_cache")
    if keyword_finder_cache_data and (
        not KEYWORD_FINDER_CACHE_PATH.exists() or KEYWORD_FINDER_CACHE_PATH.stat().st_size == 0
    ):
        if isinstance(keyword_finder_cache_data, dict):
            replace_keyword_finder_cache(keyword_finder_cache_data)


def _backup_worker() -> None:
    while True:
        try:
            perform_backup()
        except Exception:
            pass
        time.sleep(BACKUP_INTERVAL_SECONDS)


def ensure_backup_thread() -> None:
    global _backup_thread_started
    if _backup_thread_started:
        return
    _backup_thread_started = True
    thread = threading.Thread(target=_backup_worker, daemon=True)
    thread.start()


restore_from_backup()
ensure_backup_thread()


@dataclass
class CrawlJob:
    """Container for the state of a running crawl job."""

    id: str
    start_urls: List[str]
    keywords: List[str]
    max_pages: int
    concurrency: int
    total_start_urls: int
    original_keywords: List[str] = field(default_factory=list)
    expanded_keywords: List[str] = field(default_factory=list)
    keyword_groups: List[Dict[str, object]] = field(default_factory=list)
    keyword_evaluations: List[Dict[str, object]] = field(default_factory=list)
    synonym_messages: List[Dict[str, str]] = field(default_factory=list)
    synonyms_enabled: bool = True
    synonyms_expanded: bool = False
    synonyms_manual: bool = False
    synonym_status: str = "noch nicht gestartet"
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
    result_pairs: Set[Tuple[str, str, str]] = field(default_factory=set, repr=False, compare=False)
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
                "original_keywords": list(self.original_keywords),
                "expanded_keywords": list(self.expanded_keywords),
                "keyword_groups": [
                    {
                        "original": group.get("original"),
                        "additional": list(group.get("additional", [])),
                        "synonyms": list(group.get("synonyms", [])),
                        "related": list(group.get("related", [])),
                        "related_details": [
                            {
                                "term": detail.get("term"),
                                "cluster": detail.get("cluster"),
                                "search_volume": detail.get("search_volume"),
                                "competition": detail.get("competition"),
                                "relevance": detail.get("relevance"),
                            }
                            for detail in group.get("related_details", [])
                            if isinstance(detail, dict)
                        ],
                    }
                    for group in self.keyword_groups
                ],
                "keyword_evaluations": [
                    {
                        "keyword": entry.get("keyword"),
                        "quality": entry.get("quality"),
                        "better_keywords": list(entry.get("better_keywords", [])),
                    }
                    for entry in self.keyword_evaluations
                ],
                "synonym_messages": list(self.synonym_messages),
                "synonyms_enabled": self.synonyms_enabled,
                "synonyms_expanded": self.synonyms_expanded,
                "synonyms_manual": self.synonyms_manual,
                "synonym_status": self.synonym_status,
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
                keywords_tuple = tuple(progress.result.matched_keywords)
                lowered = keywords_tuple[0].lower() if keywords_tuple else ""
                key = (progress.result.source_url, progress.result.target_url, lowered)
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
    all_stammdaten_records = load_stammdaten()
    stammdaten_records = [record for record in all_stammdaten_records if record.get("aktiv", True)]
    keywords_input = ""
    selected_homepages: List[str] = []
    crawl_defaults = get_crawl_defaults()
    max_pages = crawl_defaults["max_pages"]
    concurrency = crawl_defaults["concurrency"]
    start_date_input = ""
    end_date_input = ""
    respect_robots = DEFAULT_RESPECT_ROBOTS
    synonyms_enabled = True
    keyword_groups: List[Dict[str, List[str]]] = []
    expanded_keywords: List[str] = []
    keyword_evaluations: List[Dict[str, object]] = []
    synonym_messages: List[Dict[str, str]] = []
    synonyms_expanded = False
    synonym_groups_manual = False
    synonym_status = "noch nicht gestartet"
    error: Optional[str] = None
    job_id: Optional[str] = None
    submitted = False
    has_key = has_api_key()
    has_planner_key = has_keyword_planner_key()
    synonym_defaults = get_synonym_defaults()
    keyword_finder_defaults = get_keyword_finder_defaults()

    if request.method == "POST":
        submitted = True
        raw_keywords = request.form.get("keywords", "")
        keywords_input = raw_keywords
        max_pages = request.form.get("max_pages", type=int, default=crawl_defaults["max_pages"])
        concurrency = request.form.get("concurrency", type=int, default=crawl_defaults["concurrency"])
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
        expanded_keywords = list(keywords)
        keyword_groups = [{"original": kw, "additional": []} for kw in keywords]
        manual_payload_mode = request.form.get("keyword_groups_mode", "").strip().lower()
        manual_groups_raw = request.form.get("keyword_groups_payload", "").strip()
        manual_groups_data: Optional[List[Dict[str, object]]] = None
        if manual_payload_mode == "manual":
            if not manual_groups_raw:
                manual_groups_data = []
            else:
                try:
                    loaded_groups = json.loads(manual_groups_raw)
                except json.JSONDecodeError:
                    synonym_messages.append(
                        {
                            "category": "warning",
                            "text": "Synonymerweiterung: Die übermittelten Anpassungen konnten nicht gelesen werden.",
                        }
                    )
                else:
                    if isinstance(loaded_groups, list):
                        manual_groups_data = loaded_groups
                    else:
                        synonym_messages.append(
                            {
                                "category": "warning",
                                "text": "Synonymerweiterung: Die übermittelten Anpassungen waren ungültig.",
                            }
                        )
        start_date_value: Optional[date] = None
        end_date_value: Optional[date] = None

        if not stammdaten_records:
            error = "Es sind keine aktiven Stammdaten vorhanden. Bitte importieren oder aktivieren Sie Schulen."
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

        if not has_key and error:
            synonym_messages.append(
                {
                    "category": "warning",
                    "text": "OpenAI-Schlüssel erforderlich. Bitte unter Einstellungen → API-Schlüssel hinterlegen.",
                }
            )

        if not error:
            if concurrency < 1:
                concurrency = 1
            elif concurrency > MAX_CONCURRENCY:
                concurrency = MAX_CONCURRENCY

            if manual_groups_data is not None:
                synonym_groups_manual = True
                manual_map: Dict[str, Dict[str, object]] = {}
                for entry in manual_groups_data:
                    if not isinstance(entry, dict):
                        continue
                    original_value = str(entry.get("original", "")).strip()
                    if not original_value or original_value not in keywords:
                        continue
                    store = manual_map.setdefault(
                        original_value,
                        {
                            "synonyms": [],
                            "related": [],
                            "related_details": [],
                            "_synonym_set": set(),
                            "_related_set": set(),
                            "_related_detail_map": {},
                        },
                    )

                    synonyms_raw = entry.get("synonyms")
                    if not isinstance(synonyms_raw, list):
                        synonyms_raw = entry.get("additional", [])
                    if isinstance(synonyms_raw, list):
                        for candidate in synonyms_raw:
                            text = str(candidate).strip()
                            if not text:
                                continue
                            lowered_candidate = text.lower()
                            synonym_set: Set[str] = store["_synonym_set"]  # type: ignore[assignment]
                            related_set: Set[str] = store["_related_set"]  # type: ignore[assignment]
                            if lowered_candidate in synonym_set or lowered_candidate in related_set:
                                continue
                            synonym_set.add(lowered_candidate)
                            store["synonyms"].append(text)

                    related_raw = entry.get("related")
                    if isinstance(related_raw, list):
                        for candidate in related_raw:
                            text = str(candidate).strip()
                            if not text:
                                continue
                            lowered_candidate = text.lower()
                            synonym_set = store["_synonym_set"]  # type: ignore[assignment]
                            related_set = store["_related_set"]  # type: ignore[assignment]
                            if lowered_candidate in related_set or lowered_candidate in synonym_set:
                                continue
                            related_set.add(lowered_candidate)
                            store["related"].append(text)

                    details_raw = entry.get("related_details")
                    if isinstance(details_raw, list):
                        detail_map: Dict[str, Dict[str, object]] = store["_related_detail_map"]  # type: ignore[assignment]
                        for item in details_raw:
                            if not isinstance(item, dict):
                                continue
                            term_text = str(item.get("term", "")).strip()
                            if not term_text:
                                continue
                            lowered_term = term_text.lower()
                            detail_map[lowered_term] = {
                                "term": term_text,
                                "cluster": str(item.get("cluster", "")).strip() or None,
                                "search_volume": item.get("search_volume"),
                                "competition": (str(item.get("competition", "")).strip() or None),
                                "relevance": item.get("relevance"),
                            }

                seen_lower: Set[str] = set()
                expanded_keywords = []
                keyword_groups = []
                synonyms_expanded = False

                for kw in keywords:
                    expanded_keywords.append(kw)
                    seen_lower.add(kw.lower())
                    store = manual_map.get(kw, {})
                    synonyms_list = [str(value).strip() for value in store.get("synonyms", []) if str(value).strip()]
                    related_list = [str(value).strip() for value in store.get("related", []) if str(value).strip()]
                    detail_map = store.get("_related_detail_map", {})
                    if not isinstance(detail_map, dict):
                        detail_map = {}

                    additional_clean: List[str] = []
                    synonyms_clean: List[str] = []
                    related_clean: List[str] = []

                    for candidate in synonyms_list:
                        lowered = candidate.lower()
                        if lowered in seen_lower:
                            continue
                        seen_lower.add(lowered)
                        expanded_keywords.append(candidate)
                        additional_clean.append(candidate)
                        synonyms_clean.append(candidate)

                    for candidate in related_list:
                        lowered = candidate.lower()
                        if lowered in seen_lower:
                            continue
                        seen_lower.add(lowered)
                        expanded_keywords.append(candidate)
                        additional_clean.append(candidate)
                        related_clean.append(candidate)

                    related_details: List[Dict[str, object]] = []
                    for candidate in related_clean:
                        lowered = candidate.lower()
                        detail_entry = detail_map.get(lowered)
                        if isinstance(detail_entry, dict):
                            related_details.append(
                                {
                                    "term": candidate,
                                    "cluster": detail_entry.get("cluster"),
                                    "search_volume": detail_entry.get("search_volume"),
                                    "competition": detail_entry.get("competition"),
                                    "relevance": detail_entry.get("relevance"),
                                }
                            )
                        else:
                            related_details.append(
                                {
                                    "term": candidate,
                                    "cluster": None,
                                    "search_volume": None,
                                    "competition": None,
                                    "relevance": None,
                                }
                            )

                    if additional_clean:
                        synonyms_expanded = True

                    keyword_groups.append(
                        {
                            "original": kw,
                            "additional": additional_clean,
                            "synonyms": synonyms_clean,
                            "related": related_clean,
                            "related_details": related_details,
                        }
                    )

                if synonyms_expanded:
                    synonym_status = "Zusätzliche Begriffe aktiv (manuelle Auswahl)"
                    synonym_messages.append(
                        {
                            "category": "info",
                            "text": "Synonyme und verwandte Keywords wurden manuell ausgewählt.",
                        }
                    )
                else:
                    synonym_status = "Synonymerweiterung aktiv (nur Originalbegriffe)"
                    synonym_messages.append(
                        {
                            "category": "info",
                            "text": "Synonymerweiterung aktiv, es wurden keine zusätzlichen Begriffe ausgewählt.",
                        }
                    )
            else:
                expanded_keywords = list(keywords)
                keyword_groups = [
                    {
                        "original": kw,
                        "additional": [],
                        "synonyms": [],
                        "related": [],
                        "related_details": [],
                    }
                    for kw in keywords
                ]
                synonyms_expanded = False
                synonym_status = "Synonymerweiterung bereit (keine Auswahl)"
                synonym_messages.append(
                    {
                        "category": "info",
                        "text": "Keine Synonyme ausgewählt. Verwenden Sie „Synonyme finden“, um Vorschläge hinzuzufügen.",
                    }
                )

            evaluation_payload_raw = request.form.get("keyword_evaluations_payload", "").strip()
            if evaluation_payload_raw:
                try:
                    evaluation_raw = json.loads(evaluation_payload_raw)
                except json.JSONDecodeError:
                    synonym_messages.append(
                        {
                            "category": "warning",
                            "text": "Bewertung der Suchbegriffe: Die übermittelten Daten konnten nicht gelesen werden.",
                        }
                    )
                else:
                    if isinstance(evaluation_raw, list):
                        allowed_map = {kw.lower(): kw for kw in keywords}
                        evaluation_map: Dict[str, Dict[str, object]] = {}
                        for entry in evaluation_raw:
                            if not isinstance(entry, dict):
                                continue
                            keyword_value = str(entry.get("keyword", "")).strip()
                            if not keyword_value:
                                continue
                            lowered_key = keyword_value.lower()
                            if lowered_key not in allowed_map:
                                continue
                            quality_value = str(entry.get("quality", "")).strip()
                            suggestions_raw = entry.get("better_keywords", [])
                            suggestions: List[str] = []
                            if isinstance(suggestions_raw, list):
                                seen_local: Set[str] = set()
                                for candidate in suggestions_raw:
                                    text = str(candidate).strip()
                                    if not text:
                                        continue
                                    lowered_candidate = text.lower()
                                    if lowered_candidate in seen_local:
                                        continue
                                    seen_local.add(lowered_candidate)
                                    suggestions.append(text)
                            evaluation_map[lowered_key] = {
                                "keyword": allowed_map[lowered_key],
                                "quality": quality_value,
                                "better_keywords": suggestions,
                            }
                        keyword_evaluations = [
                            evaluation_map[key.lower()]
                            for key in keywords
                            if key.lower() in evaluation_map
                        ]
                    elif evaluation_raw not in (None, ""):
                        synonym_messages.append(
                            {
                                "category": "warning",
                                "text": "Bewertung der Suchbegriffe: Die übermittelten Daten waren ungültig.",
                            }
                        )

            job_id = uuid.uuid4().hex
            job = CrawlJob(
                id=job_id,
                start_urls=selected_homepages,
                keywords=list(expanded_keywords),
                original_keywords=list(keywords),
                expanded_keywords=list(expanded_keywords),
                keyword_groups=[
                    {
                        "original": group.get("original"),
                        "additional": list(group.get("additional", [])),
                        "synonyms": list(group.get("synonyms", [])),
                        "related": list(group.get("related", [])),
                        "related_details": [
                            {
                                "term": detail.get("term"),
                                "cluster": detail.get("cluster"),
                                "search_volume": detail.get("search_volume"),
                                "competition": detail.get("competition"),
                                "relevance": detail.get("relevance"),
                            }
                            for detail in group.get("related_details", [])
                            if isinstance(detail, dict)
                        ],
                    }
                    for group in keyword_groups
                ],
                keyword_evaluations=[
                    {
                        "keyword": entry.get("keyword"),
                        "quality": entry.get("quality"),
                        "better_keywords": list(entry.get("better_keywords", [])),
                    }
                    for entry in keyword_evaluations
                ],
                synonym_messages=list(synonym_messages),
                synonyms_enabled=synonyms_enabled,
                synonyms_expanded=synonyms_expanded,
                synonym_status=synonym_status,
                synonyms_manual=synonym_groups_manual,
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
        stammdaten_primary_fields=STAMMDATEN_PRIMARY_FIELDS,
        stammdaten_boolean_fields=STAMMDATEN_BOOLEAN_FIELDS,
        selected_homepages=selected_homepages,
        keywords_input=keywords_input,
        max_pages=max_pages,
        concurrency=concurrency,
        crawl_defaults=crawl_defaults,
        start_date_input=start_date_input,
        end_date_input=end_date_input,
        respect_robots=respect_robots,
        synonyms_enabled=synonyms_enabled,
        keyword_groups=keyword_groups,
        expanded_keywords=expanded_keywords,
        synonym_messages=synonym_messages,
        synonyms_expanded=synonyms_expanded,
        synonym_status=synonym_status,
        synonym_groups_manual=synonym_groups_manual,
        keyword_evaluations=keyword_evaluations,
        default_synonym_prompt=synonym_defaults.get("prompt", DEFAULT_SYNONYM_PROMPT),
        synonym_defaults=synonym_defaults,
        keyword_finder_defaults=keyword_finder_defaults,
        available_models=AVAILABLE_OPENAI_MODELS,
        has_api_key=has_key,
        has_keyword_planner_key=has_planner_key,
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
        stammdaten_primary_fields=STAMMDATEN_PRIMARY_FIELDS,
        stammdaten_boolean_fields=STAMMDATEN_BOOLEAN_FIELDS,
        storage_file=STAMMDATEN_PATH.name,
        active_page="stammdaten",
    )


@app.route("/settings", methods=["GET", "POST"])
def settings():
    message: Optional[str] = None
    message_category: Optional[str] = None
    key_present = has_api_key()
    planner_key_present = has_keyword_planner_key()
    synonym_defaults = get_synonym_defaults()
    keyword_finder_defaults = get_keyword_finder_defaults()
    crawl_defaults = get_crawl_defaults()

    if request.method == "POST":
        form_id = request.form.get("form_id", "api")
        if form_id == "synonym-defaults":
            values = {
                "prompt": request.form.get("synonym_prompt", ""),
                "max_results": request.form.get("synonym_max_results"),
                "temperature": request.form.get("synonym_temperature"),
                "model": request.form.get("synonym_model", ""),
            }
            update_synonym_defaults(values)
            synonym_defaults = get_synonym_defaults()
            message = "Die Standardwerte für die Synonymsuche wurden gespeichert."
            message_category = "success"
        elif form_id == "crawl-defaults":
            values = {
                "max_pages": request.form.get("crawl_max_pages"),
                "concurrency": request.form.get("crawl_concurrency"),
            }
            crawl_defaults = update_crawl_defaults(values)
            message = "Die Standardwerte für den Crawl wurden gespeichert."
            message_category = "success"
        elif form_id == "keyword-finder-defaults":
            values = {
                "prompt": request.form.get("finder_prompt", ""),
                "max_results": request.form.get("finder_max_results"),
                "temperature": request.form.get("finder_temperature"),
                "model": request.form.get("finder_model", ""),
            }
            keyword_finder_defaults = update_keyword_finder_defaults(values)
            message = "Die Standardwerte für den Keyword-Finder wurden gespeichert."
            message_category = "success"
        elif form_id == "google-key":
            action = request.form.get("action", "save")
            if action == "remove":
                if planner_key_present:
                    clear_keyword_planner_key()
                    planner_key_present = False
                    message = "Der Keyword-Planner-Schlüssel wurde entfernt."
                    message_category = "success"
                else:
                    message = "Es ist kein Keyword-Planner-Schlüssel hinterlegt."
                    message_category = "info"
            else:
                key_value = request.form.get("google_api_key", "").strip()
                if not key_value:
                    message = "Bitte geben Sie einen gültigen Keyword-Planner-Schlüssel ein."
                    message_category = "danger"
                else:
                    set_keyword_planner_key(key_value)
                    planner_key_present = True
                    message = "Der Keyword-Planner-Schlüssel wurde gespeichert."
                    message_category = "success"
        else:
            action = request.form.get("action", "save")
            if action == "remove":
                if key_present:
                    clear_api_key()
                    key_present = False
                    message = "Der API-Schlüssel wurde entfernt."
                    message_category = "success"
                else:
                    message = "Es ist kein API-Schlüssel hinterlegt."
                    message_category = "info"
            elif action == "test":
                candidate = request.form.get("api_key", "").strip()
                key_to_test = candidate or get_api_key()
                if not key_to_test:
                    message = "Bitte geben Sie einen API-Schlüssel ein oder speichern Sie ihn zuerst."
                    message_category = "danger"
                else:
                    try:
                        test_openai_api_key(key_to_test)
                    except OpenAIIntegrationError as exc:
                        message = str(exc)
                        message_category = "danger"
                    else:
                        if candidate:
                            set_api_key(candidate)
                            message = "Der API-Schlüssel wurde gespeichert und erfolgreich geprüft."
                        else:
                            message = "Der hinterlegte API-Schlüssel wurde erfolgreich geprüft."
                        message_category = "success"
                    key_present = has_api_key()
            else:
                key_value = request.form.get("api_key", "").strip()
                if not key_value:
                    message = "Bitte geben Sie einen gültigen API-Schlüssel ein."
                    message_category = "danger"
                else:
                    set_api_key(key_value)
                    key_present = True
                    message = "Der API-Schlüssel wurde gespeichert."
                    message_category = "success"

    key_present = has_api_key()
    planner_key_present = has_keyword_planner_key()
    keyword_finder_defaults = get_keyword_finder_defaults()
    crawl_defaults = get_crawl_defaults()

    return render_template(
        "settings.html",
        message=message,
        message_category=message_category,
        key_present=key_present,
        keyword_planner_key_present=planner_key_present,
        synonym_defaults=synonym_defaults,
        available_models=AVAILABLE_OPENAI_MODELS,
        crawl_defaults=crawl_defaults,
        max_concurrency=MAX_CONCURRENCY,
        max_max_pages=MAX_MAX_PAGES,
        keyword_finder_defaults=keyword_finder_defaults,
        keyword_finder_max_results=DEFAULT_KEYWORD_FINDER_MAX_RESULTS,
        active_page="settings",
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
            "original_keywords": list(job.original_keywords),
            "expanded_keywords": list(job.expanded_keywords),
            "keyword_groups": [
                {
                    "original": group.get("original"),
                    "additional": list(group.get("additional", [])),
                    "synonyms": list(group.get("synonyms", [])),
                    "related": list(group.get("related", [])),
                    "related_details": [
                        {
                            "term": detail.get("term"),
                            "cluster": detail.get("cluster"),
                            "search_volume": detail.get("search_volume"),
                            "competition": detail.get("competition"),
                            "relevance": detail.get("relevance"),
                        }
                        for detail in group.get("related_details", [])
                        if isinstance(detail, dict)
                    ],
                }
                for group in job.keyword_groups
            ],
            "keyword_evaluations": [
                {
                    "keyword": entry.get("keyword"),
                    "quality": entry.get("quality"),
                    "better_keywords": list(entry.get("better_keywords", [])),
                }
                for entry in job.keyword_evaluations
            ],
            "synonym_messages": list(job.synonym_messages),
            "synonyms_enabled": job.synonyms_enabled,
            "synonyms_expanded": job.synonyms_expanded,
            "synonyms_manual": job.synonyms_manual,
            "synonym_status": job.synonym_status,
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
