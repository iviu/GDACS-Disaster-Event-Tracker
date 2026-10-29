from __future__ import annotations
import contextlib, threading, traceback, webbrowser, logging
import csv, io, json, os, queue, re, sys, time, requests
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
import pandas as pd
import tkinter as tk
from requests.adapters import HTTPAdapter
from tkinter import filedialog, messagebox, simpledialog, ttk
from urllib3.util.retry import Retry
try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    HAS_MPL = True
except Exception:  # pragma: no cover
    HAS_MPL = False
try:
    import openpyxl  # noqa: F401
    HAS_XLSX = True
except Exception:
    HAS_XLSX = False

APP_NAME = "GDACS Disaster Event Tracker"
APP_VERSION = "0.1"
API_URL = "https://www.gdacs.org/gdacsapi/api/events/geteventlist/SEARCH"
LOG_PATH = Path("log.json")
ANALYSIS_PATH = Path("log_analysis.txt")
SETTINGS_PATH = Path("settings.json")
APP_LOG_PATH = Path("weather_app.log")
CACHE_PATH = Path("cache_last_df.pkl")
HISTORY_PATH = Path("run_history.json")
GEOMETRY_PATH = Path("window_geometry.json")
PRESETS_PATH = Path("filter_presets.json")
HTTP_CACHE_PATH = Path("http_cache.json")
EVENT_TYPES = ["EQ", "TC", "FL", "VO", "WF", "DR"]
ALERT_LEVELS = ["red", "orange"]
PAGE_SIZE = 100
REQUEST_TIMEOUT = 20
MAX_PAGES = 200
MAX_CONSOLE_LINES = 5000
COLUMNS: List[str] = [
    "eventid", "eventtype", "alertlevel", "country",
    "headline", "fromdate", "todate", "duration",
    "longitude", "latitude",
]
COLUMN_LABELS: Dict[str, str] = {
    "eventid": "ID", "eventtype": "Type", "alertlevel": "Alert",
    "country": "Country", "headline": "Headline",
    "fromdate": "From", "todate": "To", "duration": "Days",
    "longitude": "Lon", "latitude": "Lat",
}
COLUMN_WIDTHS: Dict[str, int] = {
    "eventid": 90, "eventtype": 60, "alertlevel": 70, "country": 130,
    "headline": 380, "fromdate": 150, "todate": 150, "duration": 70,
    "longitude": 80, "latitude": 80,
}
THEMES: Dict[str, Dict[str, str]] = {
    "dark": {
        "BG": "#1e1f26", "PANEL": "#262832", "FG": "#e6e6e6",
        "ACCENT": "#4f9dff", "SUBTLE": "#9aa0ac", "HEAD": "#33364a",
        "STATUS_BG": "#15161b", "STATUS_FG": "#b9bec7",
        "OK": "#7ee787", "WARN": "#ffa94d", "ERR": "#ff6b6b",
        "NEW": "#2ea043",
    },
    "light": {
        "BG": "#f4f5f7", "PANEL": "#ffffff", "FG": "#1e1f26",
        "ACCENT": "#2563eb", "SUBTLE": "#5b6470", "HEAD": "#e2e8f0",
        "STATUS_BG": "#e5e7eb", "STATUS_FG": "#374151",
        "OK": "#1a7f37", "WARN": "#b45309", "ERR": "#b91c1c",
        "NEW": "#15803d",
    },
}
COUNTRY_FLAGS = {
    "United States": "🇺🇸", "Japan": "🇯🇵", "China": "🇨🇳",
    "India": "🇮🇳", "Indonesia": "🇮🇩", "Philippines": "🇵🇭",
    "Mexico": "🇲🇽", "Brazil": "🇧🇷", "Australia": "🇦🇺",
    "Turkey": "🇹🇷", "Greece": "🇬🇷", "Italy": "🇮🇹",
    "France": "🇫🇷", "Germany": "🇩🇪", "Spain": "🇪🇸",
    "Chile": "🇨🇱", "Peru": "🇵🇪", "Colombia": "🇨🇴",
    "Canada": "🇨🇦", "Russia": "🇷🇺", "Bangladesh": "🇧🇩",
    "Pakistan": "🇵🇰", "Nepal": "🇳🇵", "Iran": "🇮🇷",
    "New Zealand": "🇳🇿", "Papua New Guinea": "🇵🇬",
}

def setup_logging() -> logging.Logger:
    logger = logging.getLogger("weather")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    try:
        handler: logging.Handler = RotatingFileHandler(
            APP_LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
    except OSError:
        handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger

LOG = setup_logging()

def build_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=4,
        backoff_factor=0.6,
        backoff_jitter=0.3,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=8, pool_maxsize=8)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update({
        "User-Agent": f"weather.py/{APP_VERSION} (GDACS tracker)",
        "Accept": "application/json",
    })
    return s

def _atomic_write_text(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)

def _atomic_write_json(path: Path, obj: Any) -> None:
    _atomic_write_text(path, json.dumps(obj, indent=2, default=str))

def _parse_iso(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        parsed = datetime.fromisoformat(str(s).replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)

def safe_isna(v: Any) -> bool:
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False

def country_flag(name: str) -> str:
    return COUNTRY_FLAGS.get(name, "")

@dataclass
class Settings:
    lookback_days: int = 7
    auto_refresh_minutes: int = 0
    alert_levels: List[str] = field(default_factory=lambda: ["red", "orange"])
    event_types: List[str] = field(default_factory=lambda: list(EVENT_TYPES))
    dedupe_window_days: int = 28
    page_size: int = PAGE_SIZE
    theme: str = "dark"
    filters: Dict[str, Any] = field(default_factory=dict)
    column_order: List[str] = field(default_factory=list)
    hidden_columns: List[str] = field(default_factory=list)
    http_cache_minutes: int = 0
    auto_export_csv: bool = False
    auto_export_dir: str = ""
    show_flags: bool = True

    @classmethod
    def load(cls, path: Path) -> "Settings":
        if not path.exists():
            return cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return cls()
            valid: Dict[str, Any] = {}
            fields = cls.__dataclass_fields__
            for k, v in data.items():
                if k not in fields:
                    continue
                default = getattr(cls(), k)
                if isinstance(default, list) and not isinstance(v, list):
                    continue
                if isinstance(default, int) and not isinstance(v, int):
                    try:
                        v = int(v)
                    except (TypeError, ValueError):
                        continue
                if isinstance(default, bool) and not isinstance(v, bool):
                    v = bool(v)
                valid[k] = v
            if valid.get("theme") not in THEMES:
                valid["theme"] = "dark"
            return cls(**valid)
        except (json.JSONDecodeError, TypeError, OSError):
            return cls()

    def save(self, path: Path) -> None:
        try:
            _atomic_write_json(path, asdict(self))
        except OSError as exc:  # pragma: no cover
            LOG.warning("Could not save settings: %s", exc)

def load_log(path: Path) -> List[List[Dict[str, Any]]]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            raise ValueError("log.json root must be a list")
        cleaned: List[List[Dict[str, Any]]] = []
        for week in data:
            if not isinstance(week, list):
                continue
            cleaned_week = [ev for ev in week if isinstance(ev, dict)]
            if cleaned_week:
                cleaned.append(cleaned_week)
        return cleaned
    except (json.JSONDecodeError, ValueError, OSError):
        backup = path.with_suffix(f".corrupt.{int(time.time())}.json")
        with contextlib.suppress(OSError):
            os.replace(path, backup)
            LOG.warning("Corrupt log backed up to %s", backup)
        return []

def save_log(path: Path, weekly_logs: List[List[Dict[str, Any]]]) -> None:
    _atomic_write_json(path, weekly_logs)

def append_run_history(entry: Dict[str, Any], max_entries: int = 200) -> None:
    try:
        data: List[Dict[str, Any]] = []
        if HISTORY_PATH.exists():
            try:
                data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
                if not isinstance(data, list):
                    data = []
            except (json.JSONDecodeError, OSError):
                data = []
        data.append(entry)
        data = data[-max_entries:]
        _atomic_write_json(HISTORY_PATH, data)
    except OSError as exc:  # pragma: no cover
        LOG.warning("Could not append run history: %s", exc)

def load_run_history() -> List[Dict[str, Any]]:
    if not HISTORY_PATH.exists():
        return []
    try:
        data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []

def load_presets() -> Dict[str, Dict[str, Any]]:
    if not PRESETS_PATH.exists():
        return {}
    try:
        data = json.loads(PRESETS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}

def save_presets(presets: Dict[str, Dict[str, Any]]) -> None:
    try:
        _atomic_write_json(PRESETS_PATH, presets)
    except OSError:
        pass

def _http_cache_key(params: Dict[str, str]) -> str:
    return json.dumps(params, sort_keys=True)

def load_http_cache() -> Dict[str, Any]:
    if not HTTP_CACHE_PATH.exists():
        return {}
    try:
        data = json.loads(HTTP_CACHE_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}

def save_http_cache(cache: Dict[str, Any]) -> None:
    try:
        _atomic_write_json(HTTP_CACHE_PATH, cache)
    except OSError:
        pass

def collect_past_eventids(
    weekly_logs: Sequence[Sequence[Dict[str, Any]]],
    window_days: int,
) -> set[str]:
    cutoff = datetime.now(timezone.utc) - timedelta(days=window_days)
    ids: set[str] = set()
    for week in weekly_logs:
        for ev in week:
            eid = ev.get("eventid")
            if not eid:
                continue
            fd = _parse_iso(ev.get("fromdate"))
            if fd is None or fd >= cutoff:
                ids.add(str(eid))
    return ids

def simplify_event(event: Dict[str, Any]) -> Dict[str, Any]:
    props = event.get("properties") or {}
    geom = event.get("geometry") or {}
    coords = geom.get("coordinates")
    if (
        isinstance(coords, list) and coords
        and isinstance(coords[0], list) and coords[0]
        and isinstance(coords[0][0], list)
    ):
        try:
            ring = coords[0]
            lon = sum(p[0] for p in ring) / len(ring)
            lat = sum(p[1] for p in ring) / len(ring)
            coords = [lon, lat]
        except (TypeError, IndexError, ZeroDivisionError, KeyError):
            coords = None
    if isinstance(coords, list) and len(coords) >= 2:
        try:
            coords = [float(coords[0]), float(coords[1])]
        except (TypeError, ValueError):
            coords = None
    else:
        coords = None
    return {
        "eventid": props.get("eventid"),
        "eventtype": props.get("eventtype"),
        "alertlevel": props.get("alertlevel"),
        "country": props.get("country"),
        "headline": props.get("headline") or "No headline",
        "fromdate": props.get("fromdate"),
        "todate": props.get("todate"),
        "coordinates": coords,
    }

def fetch_events(
    session: requests.Session,
    settings: Settings,
    log: Callable[[str], None] = print,
    cancelled: Optional[threading.Event] = None,
) -> List[Dict[str, Any]]:
    today = datetime.now(timezone.utc).date()
    week_ago = today - timedelta(days=settings.lookback_days)
    params_base = {
        "eventlist": ";".join(settings.event_types) or ";".join(EVENT_TYPES),
        "alertlevel": ";".join(settings.alert_levels) or ";".join(ALERT_LEVELS),
        "fromdate": week_ago.strftime("%Y-%m-%d"),
        "todate": today.strftime("%Y-%m-%d"),
        "pagesize": str(settings.page_size),
    }
    all_events: List[Dict[str, Any]] = []
    seen_ids: set[str] = set()
    page = 1
    http_cache = load_http_cache() if settings.http_cache_minutes > 0 else {}
    now_ts = time.time()
    cache_hits = 0
    while page <= MAX_PAGES:
        if cancelled and cancelled.is_set():
            log("Fetch cancelled by user.")
            break
        params = dict(params_base, pagenumber=page)
        cache_key = _http_cache_key(params)
        data: Optional[Dict[str, Any]] = None
        if settings.http_cache_minutes > 0:
            cached = http_cache.get(cache_key)
            if cached and (now_ts - cached.get("ts", 0)) < settings.http_cache_minutes * 60:
                data = cached.get("data")
                cache_hits += 1
        if data is None:
            try:
                r = session.get(API_URL, params=params, timeout=REQUEST_TIMEOUT)
                r.raise_for_status()
            except requests.RequestException as exc:
                log(f"Page {page} request failed: {exc}")
                raise
            body = r.text.strip()
            if not body:
                log(f"Pagination finished (empty body) at page {page}.")
                break
            try:
                data = r.json()
            except json.JSONDecodeError:
                log(f"Page {page}: non-JSON response, stopping.")
                break
            if settings.http_cache_minutes > 0:
                http_cache[cache_key] = {"ts": now_ts, "data": data}
        feats = data.get("features") or []
        if not feats:
            log(f"Pagination finished (no features) at page {page}.")
            break
        new_in_page = 0
        for f in feats:
            eid = (f.get("properties") or {}).get("eventid")
            if eid and eid in seen_ids:
                continue
            if eid:
                seen_ids.add(str(eid))
            all_events.append(f)
            new_in_page += 1
        if new_in_page == 0:
            log(f"Page {page}: no new events — pagination finished.")
            break
        log(f"Page {page}: {new_in_page} events (total {len(all_events)})")
        if len(feats) < settings.page_size:
            break
        page += 1
    if settings.http_cache_minutes > 0:
        try:
            save_http_cache(http_cache)
        except Exception:
            pass
    if cache_hits:
        log(f"HTTP cache hits: {cache_hits} page(s).")
    return all_events

def build_dataframe(weekly_logs: Sequence[Sequence[Dict[str, Any]]]) -> pd.DataFrame:
    flat = [ev for week in weekly_logs for ev in week if isinstance(ev, dict)]
    if not flat:
        return pd.DataFrame(columns=COLUMNS)
    df = pd.json_normalize(flat)
    if "eventid" in df.columns:
        df["eventid"] = df["eventid"].astype(str)
        df = df.drop_duplicates(subset=["eventid"], keep="last")
    for col in ("fromdate", "todate"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)
    if "country" in df.columns:
        df["country"] = df["country"].fillna("Unknown").astype(str)
        df["country"] = df["country"].str.split(r",\s*")
        df = df.explode("country").reset_index(drop=True)
        df["country"] = df["country"].fillna("Unknown").str.strip().replace("", "Unknown")
    if "coordinates" in df.columns:
        def split_coord(c: Any) -> Tuple[Optional[float], Optional[float]]:
            if isinstance(c, (list, tuple)) and len(c) >= 2:
                try:
                    return float(c[0]), float(c[1])
                except (TypeError, ValueError):
                    pass
            return None, None
        coords = df["coordinates"].apply(split_coord)
        df["longitude"] = [c[0] for c in coords]
        df["latitude"] = [c[1] for c in coords]
    if {"todate", "fromdate"}.issubset(df.columns):
        df["duration"] = (
            (df["todate"] - df["fromdate"]).dt.total_seconds() / 86400.0
        ).round(2)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = None
    return df[COLUMNS]

def build_analysis_text(df: pd.DataFrame) -> str:
    if df.empty:
        return "No events in log yet."
    parts: List[str] = [
        f"Generated: {datetime.now(timezone.utc).isoformat()}",
        f"Total rows (after country explode): {len(df)}",
        "Sample event rows:\n" + df.head(10).to_string(index=False),
    ]
    for col, title in [
        ("country", "Country counts"),
        ("eventtype", "Event type counts"),
        ("alertlevel", "Alert level counts"),
    ]:
        if col in df:
            parts.append(
                title + ":\n"
                + df[col].value_counts().to_frame("count").to_string()
            )
    if "duration" in df and df["duration"].notna().any():
        parts.append(
            "Average event duration by country (days):\n"
            + df.groupby("country")["duration"].mean()
              .sort_values(ascending=False).round(2)
              .to_frame("avg_duration").to_string()
        )
        parts.append(
            "Average event duration by event type (days):\n"
            + df.groupby("eventtype")["duration"].mean()
              .sort_values(ascending=False).round(2)
              .to_frame("avg_duration").to_string()
        )
    return "\n\n".join(parts)

@dataclass
class RunResult:
    new_events_count: int
    skipped_count: int
    total_logged: int
    unique_countries: int
    dataframe: pd.DataFrame
    analysis_text: str
    duration_seconds: float
    cancelled: bool = False
    new_event_ids: List[str] = field(default_factory=list)

def run_pipeline(
    settings: Settings,
    log_sink: Callable[[str], None],
    cancelled: Optional[threading.Event] = None,
) -> RunResult:
    started = time.monotonic()
    weekly_logs = load_log(LOG_PATH)
    log_sink(f"Loaded {len(weekly_logs)} week(s) from {LOG_PATH}.")
    past_ids = collect_past_eventids(weekly_logs, settings.dedupe_window_days)
    log_sink(f"Tracking {len(past_ids)} recent event IDs for dedupe.")
    session = build_session()
    log_sink("Fetching GDACS events...")
    raw = fetch_events(session, settings, log=log_sink, cancelled=cancelled)
    cancelled_flag = bool(cancelled and cancelled.is_set())
    log_sink(f"Fetched {len(raw)} raw events.")
    new_events: List[Dict[str, Any]] = []
    skipped = 0
    new_ids: List[str] = []
    for ev in raw:
        eid = (ev.get("properties") or {}).get("eventid")
        if eid and str(eid) in past_ids:
            skipped += 1
            continue
        new_events.append(simplify_event(ev))
        if eid:
            new_ids.append(str(eid))
    log_sink(f"New: {len(new_events)} | skipped duplicates: {skipped}")
    if new_events and not cancelled_flag:
        weekly_logs.append(new_events)
        try:
            save_log(LOG_PATH, weekly_logs)
            log_sink(f"Saved week → {LOG_PATH}")
        except OSError as exc:
            log_sink(f"Could not save log: {exc}")
    df = build_dataframe(weekly_logs)
    analysis = build_analysis_text(df)
    with contextlib.suppress(OSError):
        _atomic_write_text(ANALYSIS_PATH, analysis)
        log_sink(f"Analysis written to {ANALYSIS_PATH}")
    with contextlib.suppress(Exception):
        df.to_pickle(CACHE_PATH)
    unique_countries = int(df["country"].nunique()) if "country" in df else 0
    run_result = RunResult(
        new_events_count=len(new_events),
        skipped_count=skipped,
        total_logged=len(df),
        unique_countries=unique_countries,
        dataframe=df,
        analysis_text=analysis,
        duration_seconds=time.monotonic() - started,
        cancelled=cancelled_flag,
        new_event_ids=new_ids,
    )
    append_run_history({
        "ts": datetime.now(timezone.utc).isoformat(),
        "new": run_result.new_events_count,
        "skipped": run_result.skipped_count,
        "rows": run_result.total_logged,
        "duration": round(run_result.duration_seconds, 2),
        "cancelled": run_result.cancelled,
    })
    if settings.auto_export_csv and new_events and settings.auto_export_dir:
        try:
            out_dir = Path(settings.auto_export_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = out_dir / f"gdacs_{stamp}.csv"
            df.to_csv(path, index=False)
            log_sink(f"Auto-exported → {path}")
        except Exception as exc:
            log_sink(f"Auto-export failed: {exc}")
    return run_result

def format_event_detail(row: Dict[str, Any]) -> str:
    fields = [
        ("eventid", "Event ID"),
        ("eventtype", "Event type"),
        ("alertlevel", "Alert level"),
        ("country", "Country"),
        ("fromdate", "From (UTC)"),
        ("todate", "To (UTC)"),
        ("duration", "Duration (days)"),
        ("longitude", "Longitude"),
        ("latitude", "Latitude"),
    ]
    lines: List[str] = []
    for k, label in fields:
        v = row.get(k)
        if v is None or safe_isna(v):
            v = "—"
        if k == "country" and isinstance(v, str) and v != "—":
            flag = country_flag(v)
            if flag:
                v = f"{flag} {v}"
        lines.append(f"{label:>16}:  {v}")
    lines += ["", "Headline:", f"  {row.get('headline', '')}"]
    return "\n".join(lines)

def osm_url(lat: float, lon: float) -> str:
    return (
        f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}"
        f"#map=8/{lat}/{lon}"
    )

class WeatherApp(tk.Tk):
    FILTER_DEBOUNCE_MS = 300
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.minsize(1040, 660)
        self.settings = Settings.load(SETTINGS_PATH)
        self._queue: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()
        self._df: pd.DataFrame = pd.DataFrame(columns=COLUMNS)
        self._df_view: pd.DataFrame = self._df
        self._sort_col: Optional[str] = None
        self._sort_asc = True
        self._auto_refresh_job: Optional[str] = None
        self._last_run_summary = "Never run"
        self._debounce_job: Optional[str] = None
        self._clock_job: Optional[str] = None
        self._fetch_started_at: Optional[float] = None
        self._last_new_ids: set[str] = set()
        self._trend_canvas: Optional[FigureCanvasTkAgg] = None
        self._timeline_canvas: Optional[FigureCanvasTkAgg] = None
        self._map_canvas: Optional[FigureCanvasTkAgg] = None
        self._selection_ids: List[str] = []
        self._visible_columns: List[str] = (
            [c for c in COLUMNS if c not in self.settings.hidden_columns]
            if self.settings.hidden_columns else list(COLUMNS)
        )
        if self.settings.column_order:
            ordered = [c for c in self.settings.column_order if c in COLUMNS]
            rest = [c for c in COLUMNS if c not in ordered]
            self._visible_columns = [c for c in ordered + rest
                                     if c in self._visible_columns]
        self._setup_style()
        self._build_ui()
        self._bind_shortcuts()
        self._restore_geometry()
        self._schedule_auto_refresh()
        self._start_clock()
        self.after(80, self._drain_queue)
        self.after(200, self.start_fetch)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _theme(self) -> Dict[str, str]:
        return THEMES.get(self.settings.theme, THEMES["dark"])

    def _setup_style(self) -> None:
        t = self._theme()
        style = ttk.Style(self)
        with contextlib.suppress(tk.TclError):
            style.theme_use("clam")
        self.configure(bg=t["BG"])
        style.configure(".", background=t["BG"], foreground=t["FG"],
                        fieldbackground=t["PANEL"])
        style.configure("TFrame", background=t["BG"])
        style.configure("Card.TFrame", background=t["PANEL"])
        style.configure("TLabel", background=t["BG"], foreground=t["FG"])
        style.configure("Card.TLabel", background=t["PANEL"], foreground=t["FG"])
        style.configure("Stat.TLabel", background=t["PANEL"],
                        foreground=t["ACCENT"], font=("Segoe UI", 16, "bold"))
        style.configure("StatLabel.TLabel", background=t["PANEL"],
                        foreground=t["SUBTLE"], font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=t["BG"],
                        foreground=t["FG"], font=("Segoe UI", 15, "bold"))
        style.configure("Sub.TLabel", background=t["BG"],
                        foreground=t["SUBTLE"], font=("Segoe UI", 9))
        style.configure("TButton", background=t["PANEL"],
                        foreground=t["FG"], padding=6)
        style.map("TButton", background=[("active", t["ACCENT"])],
                  foreground=[("active", "#ffffff")])
        style.configure("Accent.TButton", background=t["ACCENT"],
                        foreground="#ffffff")
        style.map("Accent.TButton", background=[("active", "#3a7fd1")])
        style.configure("TNotebook", background=t["BG"], borderwidth=0)
        style.configure("TNotebook.Tab", background=t["PANEL"],
                        foreground=t["FG"], padding=(16, 8))
        style.map("TNotebook.Tab",
                  background=[("selected", t["ACCENT"])],
                  foreground=[("selected", "#ffffff")])
        style.configure("Treeview", background=t["PANEL"],
                        fieldbackground=t["PANEL"], foreground=t["FG"],
                        rowheight=24, borderwidth=0)
        style.configure("Treeview.Heading", background=t["HEAD"],
                        foreground=t["FG"], font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", t["ACCENT"])])
        style.map("Treeview", foreground=[("selected", "#ffffff")])
        style.configure("TProgressbar", background=t["ACCENT"],
                        troughcolor=t["PANEL"])
        style.configure("TEntry", fieldbackground=t["PANEL"],
                        foreground=t["FG"], insertcolor=t["FG"])
        style.configure("TCombobox", fieldbackground=t["PANEL"],
                        background=t["PANEL"], foreground=t["FG"],
                        arrowcolor=t["FG"])
        style.configure("Status.TLabel", background=t["STATUS_BG"],
                        foreground=t["STATUS_FG"], padding=6)
        style.configure("Red.TLabel", foreground=t["ERR"])
        style.configure("Orange.TLabel", foreground=t["WARN"])
        style.configure("Ok.TLabel", foreground=t["OK"])

    def _toggle_theme(self) -> None:
        self.settings.theme = "light" if self.settings.theme == "dark" else "dark"
        self.settings.save(SETTINGS_PATH)
        self._setup_style()
        t = self._theme()
        for widget in (self.detail_text, self.duration_text, self.console):
            with contextlib.suppress(tk.TclError):
                widget.configure(bg=t["PANEL"], fg=t["FG"],
                                 insertbackground=t["FG"])
        self.console.configure(bg=t["STATUS_BG"], fg=t["STATUS_FG"])
        self.console.tag_configure("error", foreground=t["ERR"])
        self.console.tag_configure("ok", foreground=t["OK"])
        self.console.tag_configure("warn", foreground=t["WARN"])
        self.status_var.set(f"Theme: {self.settings.theme}")
        self._log(f"Theme switched to {self.settings.theme}.", tag="ok")

    def _build_ui(self) -> None:
        self._build_header()
        self._build_stats_bar()
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=16, pady=(0, 6))
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        self._build_events_tab()
        self._build_summary_tab("country", "Countries")
        self._build_summary_tab("eventtype", "Event Types")
        self._build_summary_tab("alertlevel", "Alert Levels")
        self._build_duration_tab()
        self._build_timeline_tab()
        self._build_map_tab()
        self._build_trends_tab()
        self._build_history_tab()
        self._build_settings_tab()
        self._build_console_tab()
        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Frame(self, style="Card.TFrame")
        status_bar.pack(fill="x", side="bottom")
        ttk.Label(status_bar, textvariable=self.status_var,
                  style="Status.TLabel", anchor="w").pack(
            side="left", fill="x", expand=True)
        self.clock_var = tk.StringVar(value="")
        ttk.Label(status_bar, textvariable=self.clock_var,
                  style="Status.TLabel", anchor="e").pack(side="right", padx=10)

    def _build_header(self) -> None:
        header = ttk.Frame(self, padding=(16, 12, 16, 6))
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Label(header,
                  text=f"v{APP_VERSION} • red/orange alerts • auto-dedupe",
                  style="Sub.TLabel").pack(side="left", padx=14)
        btn_bar = ttk.Frame(header)
        btn_bar.pack(side="right")
        self.fetch_btn = ttk.Button(btn_bar, text="Refresh (F5)",
                                    style="Accent.TButton",
                                    command=self.start_fetch)
        self.fetch_btn.pack(side="left", padx=4)
        self.cancel_btn = ttk.Button(btn_bar, text="Cancel",
                                     command=self._cancel_fetch,
                                     state="disabled")
        self.cancel_btn.pack(side="left", padx=4)
        export_menu = ttk.Menubutton(btn_bar, text="Export ▾")
        menu = tk.Menu(export_menu, tearoff=0)
        menu.add_command(label="CSV (Ctrl+E)", command=self.export_csv)
        menu.add_command(label="Excel (.xlsx)", command=self.export_xlsx)
        menu.add_command(label="JSON", command=self.export_json)
        menu.add_command(label="Markdown report", command=self.export_markdown)
        menu.add_command(label="Analysis text", command=self.export_analysis)
        menu.add_separator()
        menu.add_command(label="Copy table to clipboard",
                         command=self.copy_to_clipboard)
        menu.add_command(label="Copy selected rows",
                         command=self.copy_selected)
        menu.add_separator()
        menu.add_command(label="Import & merge JSON…",
                         command=self.import_merge_json)
        export_menu["menu"] = menu
        export_menu.pack(side="left", padx=4)
        self.undo_btn = ttk.Button(btn_bar, text="Undo last fetch",
                                   command=self.undo_last_fetch)
        self.undo_btn.pack(side="left", padx=4)
        ttk.Button(btn_bar, text="Theme",
                   command=self._toggle_theme).pack(side="left", padx=4)
        ttk.Button(btn_bar, text="About",
                   command=self._show_about).pack(side="left", padx=4)

    def _build_stats_bar(self) -> None:
        bar = ttk.Frame(self, style="Card.TFrame", padding=(16, 10))
        bar.pack(fill="x", padx=16, pady=(0, 8))
        self.stat_vars: Dict[str, tk.StringVar] = {
            "new": tk.StringVar(value="—"),
            "dupes": tk.StringVar(value="—"),
            "rows": tk.StringVar(value="—"),
            "countries": tk.StringVar(value="—"),
            "last": tk.StringVar(value="—"),
        }
        labels = [
            ("new", "New this run"),
            ("dupes", "Duplicates skipped"),
            ("rows", "Rows in log"),
            ("countries", "Countries"),
            ("last", "Last run"),
        ]
        for i, (key, cap) in enumerate(labels):
            cell = ttk.Frame(bar, style="Card.TFrame")
            cell.grid(row=0, column=i, padx=(0 if i == 0 else 24, 0), sticky="w")
            ttk.Label(cell, text=cap, style="StatLabel.TLabel").pack(anchor="w")
            ttk.Label(cell, textvariable=self.stat_vars[key],
                      style="Stat.TLabel").pack(anchor="w")

    def _build_events_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Events")
        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Search:").pack(side="left")
        self.filter_var = tk.StringVar(value=self.settings.filters.get("text", ""))
        self.filter_entry = ttk.Entry(top, textvariable=self.filter_var, width=36)
        self.filter_entry.pack(side="left", padx=6)
        self.filter_entry.bind("<KeyRelease>", lambda e: self._schedule_filter())
        ttk.Label(top, text="Type:").pack(side="left")
        self.type_filter_var = tk.StringVar(
            value=self.settings.filters.get("type", "All")
        )
        self.type_combo = ttk.Combobox(
            top, textvariable=self.type_filter_var,
            values=["All"] + sorted(EVENT_TYPES), width=8, state="readonly",
        )
        self.type_combo.pack(side="left", padx=4)
        self.type_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        ttk.Label(top, text="Alert:").pack(side="left")
        self.alert_filter_var = tk.StringVar(
            value=self.settings.filters.get("alert", "All")
        )
        self.alert_combo = ttk.Combobox(
            top, textvariable=self.alert_filter_var,
            values=["All"] + sorted(ALERT_LEVELS), width=8, state="readonly",
        )
        self.alert_combo.pack(side="left", padx=4)
        self.alert_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        ttk.Label(top, text="Country:").pack(side="left", padx=(8, 0))
        self.country_filter_var = tk.StringVar(
            value=self.settings.filters.get("country", "All")
        )
        self.country_combo = ttk.Combobox(
            top, textvariable=self.country_filter_var,
            values=["All"], width=14, state="readonly",
        )
        self.country_combo.pack(side="left", padx=4)
        self.country_combo.bind("<<ComboboxSelected>>",
                                lambda e: self._apply_filter())
        ttk.Label(top, text="Min days:").pack(side="left", padx=(8, 0))
        self.min_days_var = tk.StringVar(
            value=str(self.settings.filters.get("min_days", ""))
        )
        self.min_days_entry = ttk.Entry(top, textvariable=self.min_days_var, width=6)
        self.min_days_entry.pack(side="left", padx=4)
        self.min_days_entry.bind("<KeyRelease>", lambda e: self._schedule_filter())
        ttk.Button(top, text="Clear (Ctrl+L)",
                   command=self._clear_filter).pack(side="left", padx=6)
        ttk.Label(top, text="Columns:").pack(side="left", padx=(12, 0))
        self.cols_var = tk.StringVar(
            value=f"{len(self._visible_columns)}/{len(COLUMNS)}"
        )
        ttk.Button(top, textvariable=self.cols_var,
                   command=self._open_column_chooser).pack(side="left")
        ttk.Button(top, text="Presets",
                   command=self._open_presets).pack(side="left", padx=(6, 0))
        paned = ttk.Panedwindow(frame, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left, weight=4)
        paned.add(right, weight=1)
        self._paned = paned
        self.tree = ttk.Treeview(
            left, columns=self._visible_columns, show="headings",
            selectmode="extended",
        )
        for c in self._visible_columns:
            self.tree.heading(
                c, text=COLUMN_LABELS[c],
                command=lambda col=c: self._sort_by(col),
            )
            self.tree.column(
                c, width=COLUMN_WIDTHS[c], anchor="w",
                stretch=(c == "headline"),
            )
        self.tree.tag_configure("new", background=self._theme()["HEAD"],
                                foreground=self._theme()["NEW"])
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)
        self.tree.bind("<Double-1>", self._on_row_double_click)
        self.tree.bind("<Return>", lambda e: self._open_selected_map())
        self.tree.bind("<Escape>", lambda e: self._clear_filter())
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        detail_wrap = ttk.Frame(right, padding=(8, 0, 0, 0))
        detail_wrap.pack(fill="both", expand=True)
        ttk.Label(detail_wrap, text="Event detail",
                  style="Sub.TLabel").pack(anchor="w")
        t = self._theme()
        self.detail_text = tk.Text(
            detail_wrap, wrap="word", bg=t["PANEL"], fg=t["FG"],
            insertbackground=t["FG"], relief="flat", padx=10, pady=10,
            font=("Consolas", 9), width=36, height=14,
        )
        self.detail_text.pack(fill="both", expand=True, pady=(4, 6))
        self.map_btn = ttk.Button(
            detail_wrap, text="Open in OSM",
            command=self._open_selected_map, state="disabled",
        )
        self.map_btn.pack(anchor="e")

    def _build_summary_tab(self, key: str, title: str) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text=title)
        tree = ttk.Treeview(frame, columns=(key, "count"), show="headings")
        tree.heading(key, text=COLUMN_LABELS.get(key, key.title()))
        tree.heading("count", text="Count")
        tree.column(key, width=280, anchor="w")
        tree.column("count", width=120, anchor="e")
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", lambda e, k=key: self._summary_drill(k))
        setattr(self, f"_tree_{key}", tree)

    def _summary_drill(self, key: str) -> None:
        tree: ttk.Treeview = getattr(self, f"_tree_{key}")
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        if not vals:
            return
        value = vals[0]
        if key == "country":
            self.country_filter_var.set(value)
        elif key == "eventtype":
            self.type_filter_var.set(value)
        elif key == "alertlevel":
            self.alert_filter_var.set(value)
        self._apply_filter()
        self.nb.select(0)
        self.status_var.set(f"Filtered by {key} = {value}")

    def _build_duration_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Duration")
        t = self._theme()
        self.duration_text = tk.Text(
            frame, wrap="none", bg=t["PANEL"], fg=t["FG"],
            insertbackground=t["FG"], relief="flat", padx=10, pady=10,
            font=("Consolas", 10),
        )
        self.duration_text.pack(fill="both", expand=True)

    def _build_timeline_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Timeline")
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Save PNG…",
                   command=lambda: self._save_chart(self._timeline_canvas, "timeline.png")
                   ).pack(side="left")
        self.timeline_frame = ttk.Frame(frame)
        self.timeline_frame.pack(fill="both", expand=True)
        self.timeline_placeholder = ttk.Label(
            self.timeline_frame,
            text="Timeline appears after first successful fetch.",
            style="Sub.TLabel",
        )
        self.timeline_placeholder.pack(anchor="w")

    def _build_map_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Map")
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Save PNG…",
                   command=lambda: self._save_chart(self._map_canvas, "map.png")
                   ).pack(side="left")
        self.map_frame = ttk.Frame(frame)
        self.map_frame.pack(fill="both", expand=True)
        self.map_placeholder = ttk.Label(
            self.map_frame,
            text="Event map appears after first successful fetch.",
            style="Sub.TLabel",
        )
        self.map_placeholder.pack(anchor="w")

    def _build_trends_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Trends")
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Save PNG…",
                   command=lambda: self._save_chart(self._trend_canvas, "trends.png")
                   ).pack(side="left")
        self.trends_placeholder = ttk.Label(
            frame, text="Charts appear after first successful fetch.",
            style="Sub.TLabel",
        )
        self.trends_placeholder.pack(anchor="w")
        self.trends_frame = frame

    def _build_history_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="History")
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Refresh", command=self._refresh_history_view
                   ).pack(side="left")
        ttk.Button(bar, text="Clear history",
                   command=self._clear_history).pack(side="left", padx=6)
        cols = ("ts", "new", "skipped", "rows", "duration", "cancelled")
        self.history_tree = ttk.Treeview(frame, columns=cols, show="headings")
        heads = {
            "ts": "When", "new": "New", "skipped": "Skipped",
            "rows": "Rows", "duration": "Duration (s)", "cancelled": "Cancelled",
        }
        for c in cols:
            self.history_tree.heading(c, text=heads[c])
            self.history_tree.column(c, width=140, anchor="w")
        self.history_tree.pack(fill="both", expand=True)

    def _build_settings_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=16)
        self.nb.add(frame, text="Settings")
        ttk.Label(frame, text="Fetch settings",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                             sticky="w", pady=(0, 12))
        self.lookback_var = tk.IntVar(value=self.settings.lookback_days)
        self.dedupe_var = tk.IntVar(value=self.settings.dedupe_window_days)
        self.page_var = tk.IntVar(value=self.settings.page_size)
        self.auto_var = tk.IntVar(value=self.settings.auto_refresh_minutes)
        self.cache_var = tk.IntVar(value=self.settings.http_cache_minutes)
        rows = [
            ("Lookback days", self.lookback_var, 1, 365),
            ("Dedupe window days", self.dedupe_var, 1, 365),
            ("Page size", self.page_var, 10, 500),
            ("Auto-refresh interval (minutes, 0=off)", self.auto_var, 0, 1440),
            ("HTTP cache TTL (minutes, 0=off)", self.cache_var, 0, 1440),
        ]
        for i, (label, var, lo, hi) in enumerate(rows, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Spinbox(frame, from_=lo, to=hi, textvariable=var,
                        width=8).grid(row=i, column=1, sticky="w", pady=4)
        self.alert_vars = {
            a: tk.BooleanVar(value=a in self.settings.alert_levels)
            for a in ALERT_LEVELS
        }
        self.type_vars = {
            t: tk.BooleanVar(value=t in self.settings.event_types)
            for t in EVENT_TYPES
        }
        self.auto_export_var = tk.BooleanVar(value=self.settings.auto_export_csv)
        self.auto_export_dir_var = tk.StringVar(value=self.settings.auto_export_dir)
        self.show_flags_var = tk.BooleanVar(value=self.settings.show_flags)
        base = len(rows) + 1
        ttk.Label(frame, text="Alert levels").grid(row=base, column=0,
                                                   sticky="w", pady=(12, 4))
        alert_box = ttk.Frame(frame)
        alert_box.grid(row=base, column=1, sticky="w", pady=(12, 4))
        for a in ALERT_LEVELS:
            ttk.Checkbutton(alert_box, text=a,
                            variable=self.alert_vars[a]).pack(side="left", padx=4)
        ttk.Label(frame, text="Event types").grid(row=base + 1, column=0,
                                                   sticky="w", pady=(4, 4))
        type_box = ttk.Frame(frame)
        type_box.grid(row=base + 1, column=1, sticky="w", pady=(4, 4))
        for t in EVENT_TYPES:
            ttk.Checkbutton(type_box, text=t,
                            variable=self.type_vars[t]).pack(side="left", padx=4)
        ttk.Checkbutton(frame, text="Show country flags in detail panel",
                        variable=self.show_flags_var).grid(
            row=base + 2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(frame, text="Auto-export CSV after each fetch",
                        variable=self.auto_export_var).grid(
            row=base + 3, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Label(frame, text="Auto-export directory").grid(
            row=base + 4, column=0, sticky="w", pady=(4, 0))
        export_row = ttk.Frame(frame)
        export_row.grid(row=base + 4, column=1, sticky="w", pady=(4, 0))
        ttk.Entry(export_row, textvariable=self.auto_export_dir_var,
                  width=40).pack(side="left")
        ttk.Button(export_row, text="Browse…",
                   command=self._choose_export_dir).pack(side="left", padx=4)
        self.settings_hint = ttk.Label(
            frame, text="", style="Sub.TLabel", foreground="#ff6b6b"
        )
        self.settings_hint.grid(row=base + 5, column=0, columnspan=2,
                                sticky="w", pady=(8, 0))
        btn = ttk.Frame(frame)
        btn.grid(row=base + 6, column=0, columnspan=2, pady=(16, 0), sticky="w")
        ttk.Button(btn, text="Save settings", style="Accent.TButton",
                   command=self._save_settings).pack(side="left")
        ttk.Button(btn, text="Reset to defaults",
                   command=self._reset_settings).pack(side="left", padx=8)
        ttk.Button(btn, text="Open app log file",
                   command=self._open_app_log).pack(side="left", padx=8)
        frame.columnconfigure(1, weight=1)

    def _build_console_tab(self) -> None:
        frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(frame, text="Log")
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Save log to file…",
                   command=self._save_log).pack(side="left")
        ttk.Button(bar, text="Clear",
                   command=lambda: self.console.delete("1.0", "end")
                   ).pack(side="left", padx=6)
        t = self._theme()
        self.console = tk.Text(
            frame, wrap="word", bg=t["STATUS_BG"], fg=t["STATUS_FG"],
            insertbackground=t["FG"], relief="flat", padx=10, pady=10,
            font=("Consolas", 9),
        )
        self.console.pack(fill="both", expand=True)
        self.console.tag_configure("error", foreground=t["ERR"])
        self.console.tag_configure("ok", foreground=t["OK"])
        self.console.tag_configure("warn", foreground=t["WARN"])

    def _bind_shortcuts(self) -> None:
        self.bind("<F5>", lambda e: self.start_fetch())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<Control-f>", lambda e: self.filter_entry.focus_set())
        self.bind("<Control-l>", lambda e: self._clear_filter())
        self.bind("<Control-t>", lambda e: self._toggle_theme())
        self.bind("<Control-z>", lambda e: self.undo_last_fetch())
        self.bind("<Control-a>", lambda e: self._select_all_rows())
        self.bind("<Control-i>", lambda e: self.import_merge_json())
        self.bind("<slash>", lambda e: self.filter_entry.focus_set())

    def _select_all_rows(self) -> None:
        self.tree.selection_set(self.tree.get_children())
        self._on_row_selected()

    def _restore_geometry(self) -> None:
        try:
            if GEOMETRY_PATH.exists():
                geo = json.loads(GEOMETRY_PATH.read_text(encoding="utf-8"))
                if isinstance(geo.get("geometry"), str):
                    self.geometry(geo["geometry"])
                if isinstance(geo.get("sash"), int):
                    with contextlib.suppress(tk.TclError):
                        self._paned.sashpos(0, geo["sash"])
                if isinstance(geo.get("tab"), int):
                    with contextlib.suppress(tk.TclError):
                        self.nb.select(geo["tab"])
            else:
                self.geometry("1360x860")
        except (json.JSONDecodeError, OSError, tk.TclError):
            self.geometry("1360x860")

    def _save_geometry(self) -> None:
        try:
            payload: Dict[str, Any] = {"geometry": self.geometry()}
            with contextlib.suppress(tk.TclError):
                payload["sash"] = self._paned.sashpos(0)
            with contextlib.suppress(tk.TclError):
                payload["tab"] = self.nb.index(self.nb.select())
            _atomic_write_text(GEOMETRY_PATH, json.dumps(payload))
        except OSError:
            pass

    def _start_clock(self) -> None:
        self.clock_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._clock_job = self.after(1000, self._start_clock)

    def _schedule_filter(self) -> None:
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(
            self.FILTER_DEBOUNCE_MS, self._apply_filter
        )

    def _schedule_auto_refresh(self) -> None:
        if self._auto_refresh_job:
            self.after_cancel(self._auto_refresh_job)
            self._auto_refresh_job = None
        minutes = self.settings.auto_refresh_minutes
        if minutes and minutes > 0:
            self._auto_refresh_job = self.after(
                minutes * 60_000, self._auto_refresh_tick
            )

    def _auto_refresh_tick(self) -> None:
        self._auto_refresh_job = None
        if not (self._worker and self._worker.is_alive()):
            self.start_fetch()
        self._schedule_auto_refresh()

    def start_fetch(self) -> None:
        if self._worker and self._worker.is_alive():
            self.status_var.set("Fetch already in progress.")
            return
        self._cancel_event.clear()
        self._fetch_started_at = time.monotonic()
        self.fetch_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress.start(12)
        self.status_var.set("Fetching events…")
        snapshot = Settings(**asdict(self.settings))
        self._worker = threading.Thread(
            target=self._fetch_job, args=(snapshot,), daemon=True,
        )
        self._worker.start()

    def _cancel_fetch(self) -> None:
        self._cancel_event.set()
        self.status_var.set("Cancelling…")

    def _fetch_job(self, settings: Settings) -> None:
        try:
            def sink(m: str) -> None:
                self._queue.put(("log", m))
            res = run_pipeline(settings, sink, self._cancel_event)
            self._queue.put(("done", res))
        except Exception as exc:
            tb = traceback.format_exc()
            LOG.exception("Pipeline failed")
            self._queue.put(("error", f"{exc}\n{tb}"))

    def _drain_queue(self) -> None:
        try:
            while True:
                try:
                    kind, payload = self._queue.get_nowait()
                except queue.Empty:
                    break
                try:
                    if kind == "log":
                        self._log(payload)
                        self.status_var.set(payload)
                    elif kind == "done":
                        self._on_done(payload)
                    elif kind == "error":
                        self._on_error(payload)
                except Exception:
                    LOG.exception("Error handling queue item: %s", kind)
        finally:
            with contextlib.suppress(tk.TclError):
                self.after(80, self._drain_queue)

    def _on_done(self, res: RunResult) -> None:
        self.progress.stop()
        self.fetch_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self._df = res.dataframe
        self._df_view = res.dataframe
        self._last_new_ids = set(res.new_event_ids)
        self._refresh_all_views()
        self.stat_vars["new"].set(str(res.new_events_count))
        self.stat_vars["dupes"].set(str(res.skipped_count))
        self.stat_vars["rows"].set(str(res.total_logged))
        self.stat_vars["countries"].set(str(res.unique_countries))
        ts = datetime.now().strftime("%H:%M:%S")
        self._last_run_summary = f"{ts} in {res.duration_seconds:.1f}s"
        self.stat_vars["last"].set(self._last_run_summary)
        state = "cancelled" if res.cancelled else "done"
        self.status_var.set(
            f"{state} — new: {res.new_events_count}, "
            f"duplicates: {res.skipped_count}, "
            f"rows in log: {res.total_logged} "
            f"({res.duration_seconds:.1f}s)"
        )
        self._log(
            f"Pipeline finished ({state}): {res.new_events_count} new events.",
            tag="warn" if res.cancelled else "ok",
        )
        self._refresh_history_view()

    def _on_error(self, msg: str) -> None:
        self.progress.stop()
        self.fetch_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.status_var.set("Error during fetch — see Log tab.")
        self._log("ERROR:\n" + msg, tag="error")
        with contextlib.suppress(Exception):
            messagebox.showerror("Fetch failed", msg.splitlines()[0])

    def _refresh_all_views(self) -> None:
        self._refresh_country_options()
        self._apply_filter()
        self._refresh_summary("country")
        self._refresh_summary("eventtype")
        self._refresh_summary("alertlevel")
        self._refresh_duration_view()
        self._refresh_timeline()
        self._refresh_map()
        self._refresh_trends()

    def _refresh_country_options(self) -> None:
        if self._df.empty or "country" not in self._df:
            self.country_combo.configure(values=["All"])
            self.country_filter_var.set("All")
            return
        countries = sorted(
            self._df["country"].dropna().unique().tolist()
        )
        values = ["All"] + countries
        self.country_combo.configure(values=values)
        current = self.country_filter_var.get()
        if current not in values:
            self.country_filter_var.set("All")

    def _apply_filter(self) -> None:
        self._debounce_job = None
        df = self._df
        if df.empty:
            self._df_view = df
            self._render_events()
            return
        query = self.filter_var.get().strip().lower()
        if query:
            text_cols = [c for c in ("eventid", "eventtype", "alertlevel",
                                     "country", "headline") if c in df.columns]
            mask = pd.Series(False, index=df.index)
            for c in text_cols:
                mask |= df[c].astype(str).str.lower().str.contains(
                    re.escape(query), na=False
                )
            df = df[mask]
        t = self.type_filter_var.get()
        if t and t != "All" and "eventtype" in df:
            df = df[df["eventtype"] == t]
        a = self.alert_filter_var.get()
        if a and a != "All" and "alertlevel" in df:
            df = df[df["alertlevel"] == a]
        cty = self.country_filter_var.get()
        if cty and cty != "All" and "country" in df:
            df = df[df["country"] == cty]
        min_days = self.min_days_var.get().strip()
        if min_days and "duration" in df:
            try:
                threshold = float(min_days)
                df = df[df["duration"].fillna(0) >= threshold]
            except ValueError:
                pass
        self._df_view = df.reset_index(drop=True)
        self.settings.filters = {
            "text": self.filter_var.get(),
            "type": self.type_filter_var.get(),
            "alert": self.alert_filter_var.get(),
            "country": self.country_filter_var.get(),
            "min_days": self.min_days_var.get(),
        }
        self._render_events()

    def _render_events(self) -> None:
        prev_sel = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        if self._df_view.empty:
            return
        for i, row in self._df_view.iterrows():
            vals: List[Any] = []
            for c in self._visible_columns:
                v = row.get(c)
                if safe_isna(v):
                    v = ""
                if isinstance(v, pd.Timestamp):
                    v = v.strftime("%Y-%m-%d %H:%M")
                vals.append(v)
            eid = str(row.get("eventid", ""))
            tags = ("new",) if eid in self._last_new_ids else ()
            self.tree.insert("", "end", iid=str(i), values=vals, tags=tags)
        still = [s for s in prev_sel if self.tree.exists(s)]
        if still:
            self.tree.selection_set(still)
            self._on_row_selected()

    def _refresh_summary(self, key: str) -> None:
        tree: ttk.Treeview = getattr(self, f"_tree_{key}")
        tree.delete(*tree.get_children())
        if self._df.empty or key not in self._df:
            return
        counts = self._df[key].fillna("Unknown").value_counts()
        for name, n in counts.items():
            label = str(name)
            if key == "country" and self.settings.show_flags:
                flag = country_flag(label)
                if flag:
                    label = f"{flag} {label}"
            tree.insert("", "end", values=(label, int(n)))

    def _refresh_duration_view(self) -> None:
        self.duration_text.delete("1.0", "end")
        df = self._df
        if df.empty or "duration" not in df:
            self.duration_text.insert("end", "No duration data yet.")
            return
        by_country = (
            df.groupby("country")["duration"].mean()
              .sort_values(ascending=False).round(2)
              .to_frame("avg_days")
        )
        by_type = (
            df.groupby("eventtype")["duration"].mean()
              .sort_values(ascending=False).round(2)
              .to_frame("avg_days")
        )
        self.duration_text.insert("end", "Average duration by country (days)\n")
        self.duration_text.insert("end", "=" * 44 + "\n")
        self.duration_text.insert("end", by_country.to_string() + "\n\n")
        self.duration_text.insert("end", "Average duration by event type (days)\n")
        self.duration_text.insert("end", "=" * 44 + "\n")
        self.duration_text.insert("end", by_type.to_string() + "\n")

    def _refresh_timeline(self) -> None:
        if not HAS_MPL:
            return
        df = self._df_view
        if df.empty or "fromdate" not in df or "todate" not in df:
            return
        valid = df.dropna(subset=["fromdate", "todate"]).copy()
        if valid.empty:
            return
        valid = valid.sort_values("fromdate").tail(80)
        if self._timeline_canvas is not None:
            self._timeline_canvas.get_tk_widget().destroy()
        self.timeline_placeholder.pack_forget()
        fig = Figure(figsize=(10, 4.5), dpi=100)
        ax = fig.add_subplot(111)
        color_map = {"red": "#d62728", "orange": "#ff7f0e"}
        y = range(len(valid))
        for i, (_, row) in enumerate(valid.iterrows()):
            start = row["fromdate"]
            end = row["todate"]
            if pd.isna(start) or pd.isna(end):
                continue
            c = color_map.get(str(row.get("alertlevel", "")), "#4f9dff")
            ax.barh(i, (end - start).total_seconds() / 86400.0,
                    left=start.to_pydatetime(), color=c, height=0.7)
        ax.set_yticks(list(y))
        labels = [f"{str(r.get('eventtype',''))[:3]} {str(r.get('country',''))[:14]}"
                  for _, r in valid.iterrows()]
        ax.set_yticklabels(labels, fontsize=7)
        ax.set_xlabel("Date")
        ax.set_title("Event durations (last 80)")
        fig.autofmt_xdate()
        fig.tight_layout()
        self._timeline_canvas = FigureCanvasTkAgg(fig, master=self.timeline_frame)
        self._timeline_canvas.draw()
        self._timeline_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _refresh_map(self) -> None:
        if not HAS_MPL:
            return
        df = self._df_view
        if df.empty or "latitude" not in df or "longitude" not in df:
            return
        valid = df.dropna(subset=["latitude", "longitude"])
        valid = valid[(valid["latitude"] != 0) | (valid["longitude"] != 0)]
        if valid.empty:
            return
        if self._map_canvas is not None:
            self._map_canvas.get_tk_widget().destroy()
        self.map_placeholder.pack_forget()
        fig = Figure(figsize=(10, 5), dpi=100)
        ax = fig.add_subplot(111)
        color_map = {"red": "#d62728", "orange": "#ff7f0e"}
        colors = [color_map.get(str(a), "#4f9dff")
                  for a in valid["alertlevel"].fillna("")]
        ax.scatter(valid["longitude"], valid["latitude"],
                   c=colors, alpha=0.75, edgecolors="white", linewidths=0.5,
                   s=60)
        ax.set_xlabel("Longitude")
        ax.set_ylabel("Latitude")
        ax.set_title(f"Event locations ({len(valid)} events)")
        ax.set_xlim(-180, 180)
        ax.set_ylim(-90, 90)
        ax.grid(True, linestyle=":", alpha=0.4)
        fig.tight_layout()
        self._map_canvas = FigureCanvasTkAgg(fig, master=self.map_frame)
        self._map_canvas.draw()
        self._map_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _refresh_trends(self) -> None:
        if not HAS_MPL:
            self.trends_placeholder.configure(
                text="Install matplotlib to see trend charts."
            )
            return
        history = load_run_history()
        if not history:
            return
        if self._trend_canvas is not None:
            self._trend_canvas.get_tk_widget().destroy()
        self.trends_placeholder.pack_forget()
        labels = [(h.get("ts", "") or "")[-8:] for h in history]
        new_vals = [h.get("new", 0) for h in history]
        skip_vals = [h.get("skipped", 0) for h in history]
        fig = Figure(figsize=(9, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        ax.plot(labels, new_vals, marker="o", label="new")
        ax.plot(labels, skip_vals, marker="s", label="skipped")
        ax.set_title("Events per run (last %d)" % len(history))
        ax.set_ylabel("count")
        ax.legend()
        ax.tick_params(axis="x", rotation=45, labelsize=7)
        fig.tight_layout()
        self._trend_canvas = FigureCanvasTkAgg(fig, master=self.trends_frame)
        self._trend_canvas.draw()
        self._trend_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _refresh_history_view(self) -> None:
        self.history_tree.delete(*self.history_tree.get_children())
        for h in reversed(load_run_history()):
            self.history_tree.insert("", "end", values=(
                h.get("ts", ""), h.get("new", ""), h.get("skipped", ""),
                h.get("rows", ""), h.get("duration", ""),
                "yes" if h.get("cancelled") else "",
            ))

    def _clear_history(self) -> None:
        if not messagebox.askyesno("Clear history",
                                   "Delete all run history?"):
            return
        with contextlib.suppress(OSError):
            HISTORY_PATH.unlink()
        self._refresh_history_view()

    def _save_chart(self, canvas: Optional[FigureCanvasTkAgg], name: str) -> None:
        if canvas is None:
            messagebox.showinfo("Save chart", "No chart to save yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            initialfile=name,
        )
        if not path:
            return
        try:
            canvas.figure.savefig(path, dpi=150)
            self.status_var.set(f"Chart saved → {path}")
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc))

    def _selected_row(self) -> Optional[Dict[str, Any]]:
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            idx = int(sel[0])
            return self._df_view.iloc[idx].to_dict()
        except (ValueError, IndexError):
            return None

    def _on_row_selected(self, _event: Any = None) -> None:
        row = self._selected_row()
        self.detail_text.delete("1.0", "end")
        sel_count = len(self.tree.selection())
        if row is None:
            self.map_btn.configure(state="disabled")
            return
        self.detail_text.insert("end", format_event_detail(row))
        lon, lat = row.get("longitude"), row.get("latitude")
        has_coords = (
            lon is not None and lat is not None
            and not safe_isna(lon) and not safe_isna(lat)
        )
        self.map_btn.configure(state="normal" if has_coords else "disabled")
        if sel_count > 1:
            self.status_var.set(f"{sel_count} rows selected")
        else:
            self.status_var.set(
                f"Selected: {row.get('eventid', '')} "
                f"({row.get('eventtype', '')}, {row.get('country', '')})"
            )

    def _on_row_double_click(self, event: Any = None) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        self._open_selected_map()

    def _on_tree_right_click(self, event: Any) -> None:
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        if row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open in OSM",
                         command=self._open_selected_map)
        menu.add_command(label="Copy Event ID",
                         command=self._copy_selected_id)
        menu.add_command(label="Copy selected row(s)",
                         command=self.copy_selected)
        menu.add_separator()
        row = self._selected_row()
        if row:
            cty = row.get("country")
            ety = row.get("eventtype")
            al = row.get("alertlevel")
            if cty:
                menu.add_command(
                    label=f"Filter by country: {cty}",
                    command=lambda c=cty: self._filter_by("country", c),
                )
            if ety:
                menu.add_command(
                    label=f"Filter by type: {ety}",
                    command=lambda c=ety: self._filter_by("type", c),
                )
            if al:
                menu.add_command(
                    label=f"Filter by alert: {al}",
                    command=lambda c=al: self._filter_by("alert", c),
                )
        menu.add_separator()
        menu.add_command(label="Clear filters", command=self._clear_filter)
        with contextlib.suppress(tk.TclError):
            menu.tk_popup(event.x_root, event.y_root)

    def _filter_by(self, field: str, value: str) -> None:
        if field == "country":
            self.country_filter_var.set(value)
        elif field == "type":
            self.type_filter_var.set(value)
        elif field == "alert":
            self.alert_filter_var.set(value)
        self._apply_filter()
        self.nb.select(0)

    def _copy_selected_id(self) -> None:
        row = self._selected_row()
        if not row:
            return
        self.clipboard_clear()
        self.clipboard_append(str(row.get("eventid", "")))
        self.status_var.set("Event ID copied.")

    def copy_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            self.copy_to_clipboard()
            return
        idxs = sorted(int(s) for s in sel)
        sub = self._df_view.iloc[idxs]
        buf = io.StringIO()
        sub.to_csv(buf, index=False)
        self.clipboard_clear()
        self.clipboard_append(buf.getvalue())
        self.status_var.set(f"Copied {len(sub)} row(s) to clipboard.")
        self._log(f"Copied {len(sub)} rows.", tag="ok")

    def _open_selected_map(self) -> None:
        row = self._selected_row()
        if not row:
            return
        lon, lat = row.get("longitude"), row.get("latitude")
        if lon is None or lat is None or safe_isna(lon) or safe_isna(lat):
            return
        with contextlib.suppress(Exception):
            webbrowser.open(osm_url(float(lat), float(lon)))

    def _sort_by(self, col: str) -> None:
        if self._df_view.empty or col not in self._df_view:
            return
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._df_view = self._df_view.sort_values(
            by=col, ascending=self._sort_asc, na_position="last"
        ).reset_index(drop=True)
        self._render_events()
        for c in self._visible_columns:
            arrow = ""
            if c == col:
                arrow = " ▲" if self._sort_asc else " ▼"
            self.tree.heading(c, text=COLUMN_LABELS[c] + arrow)

    def _clear_filter(self) -> None:
        self.filter_var.set("")
        self.type_filter_var.set("All")
        self.alert_filter_var.set("All")
        self.country_filter_var.set("All")
        self.min_days_var.set("")
        self._apply_filter()

    def _open_presets(self) -> None:
        presets = load_presets()
        win = tk.Toplevel(self)
        win.title("Filter presets")
        win.transient(self)
        win.geometry("380x360")
        listbox = tk.Listbox(win)
        listbox.pack(fill="both", expand=True, padx=8, pady=8)
        for name in presets:
            listbox.insert("end", name)

        def apply_selected() -> None:
            sel = listbox.curselection()
            if not sel:
                return
            name = listbox.get(sel[0])
            p = presets.get(name, {})
            self.filter_var.set(p.get("text", ""))
            self.type_filter_var.set(p.get("type", "All"))
            self.alert_filter_var.set(p.get("alert", "All"))
            if p.get("country", "All") in self.country_combo["values"]:
                self.country_filter_var.set(p.get("country", "All"))
            self.min_days_var.set(str(p.get("min_days", "")))
            self._apply_filter()
            win.destroy()

        def save_current() -> None:
            name = simpledialog.askstring("Save preset", "Preset name:",
                                          parent=win)
            if not name:
                return
            presets[name] = {
                "text": self.filter_var.get(),
                "type": self.type_filter_var.get(),
                "alert": self.alert_filter_var.get(),
                "country": self.country_filter_var.get(),
                "min_days": self.min_days_var.get(),
            }
            save_presets(presets)
            listbox.insert("end", name)
            self.status_var.set(f"Preset '{name}' saved.")

        def delete_selected() -> None:
            sel = listbox.curselection()
            if not sel:
                return
            name = listbox.get(sel[0])
            if name in presets:
                presets.pop(name, None)
                save_presets(presets)
                listbox.delete(sel[0])
                self.status_var.set(f"Preset '{name}' deleted.")
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bar, text="Apply", style="Accent.TButton",
                   command=apply_selected).pack(side="left")
        ttk.Button(bar, text="Save current…",
                   command=save_current).pack(side="left", padx=6)
        ttk.Button(bar, text="Delete",
                   command=delete_selected).pack(side="left")

    def _choose_export_dir(self) -> None:
        d = filedialog.askdirectory()
        if d:
            self.auto_export_dir_var.set(d)

    def _open_column_chooser(self) -> None:
        win = tk.Toplevel(self)
        win.title("Choose columns")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Toggle and reorder visible columns:",
                  style="Sub.TLabel").pack(anchor="w", padx=12, pady=(10, 4))
        vars_: Dict[str, tk.BooleanVar] = {}
        listbox = tk.Listbox(win, height=12, selectmode="browse")
        listbox.pack(fill="both", expand=True, padx=12, pady=4)
        for c in self._visible_columns:
            listbox.insert("end", f"{COLUMN_LABELS[c]} ({c})")
        remaining = [c for c in COLUMNS if c not in self._visible_columns]
        check_frame = ttk.Frame(win)
        check_frame.pack(fill="x", padx=12)
        for c in COLUMNS:
            v = tk.BooleanVar(value=c in self._visible_columns)
            vars_[c] = v
            ttk.Checkbutton(
                check_frame, text=COLUMN_LABELS[c], variable=v,
            ).pack(side="left", padx=3)

        def apply_and_close() -> None:
            chosen = [c for c in COLUMNS if vars_[c].get()]
            if not chosen:
                chosen = ["eventid"]
            self._visible_columns = chosen
            self.settings.hidden_columns = [
                c for c in COLUMNS if c not in chosen
            ]
            self.settings.column_order = chosen
            self.settings.save(SETTINGS_PATH)
            self.tree.configure(columns=self._visible_columns)
            for c in self._visible_columns:
                self.tree.heading(
                    c, text=COLUMN_LABELS[c],
                    command=lambda col=c: self._sort_by(col),
                )
                self.tree.column(
                    c, width=COLUMN_WIDTHS[c], anchor="w",
                    stretch=(c == "headline"),
                )
            self.cols_var.set(f"{len(self._visible_columns)}/{len(COLUMNS)}")
            self._render_events()
            win.destroy()

        def move_up() -> None:
            sel = listbox.curselection()
            if not sel or sel[0] == 0:
                return
            i = sel[0]
            key = self._visible_columns[i]
            self._visible_columns[i], self._visible_columns[i - 1] = (
                self._visible_columns[i - 1], self._visible_columns[i]
            )
            listbox.delete(0, "end")
            for c in self._visible_columns:
                listbox.insert("end", f"{COLUMN_LABELS[c]} ({c})")
            listbox.selection_set(i - 1)

        def move_down() -> None:
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self._visible_columns) - 1:
                return
            i = sel[0]
            self._visible_columns[i], self._visible_columns[i + 1] = (
                self._visible_columns[i + 1], self._visible_columns[i]
            )
            listbox.delete(0, "end")
            for c in self._visible_columns:
                listbox.insert("end", f"{COLUMN_LABELS[c]} ({c})")
            listbox.selection_set(i + 1)
        move_bar = ttk.Frame(win)
        move_bar.pack(fill="x", padx=12)
        ttk.Button(move_bar, text="▲ Up", command=move_up).pack(side="left")
        ttk.Button(move_bar, text="▼ Down", command=move_down).pack(
            side="left", padx=6)
        ttk.Button(win, text="Apply", style="Accent.TButton",
                   command=apply_and_close).pack(pady=12)

    def _log(self, msg: str, tag: Optional[str] = None) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.console.insert("end", f"[{ts}] {msg}\n", tag or ())
        total_lines = int(self.console.index("end-1c").split(".")[0])
        if total_lines > MAX_CONSOLE_LINES:
            self.console.delete("1.0", f"{total_lines - MAX_CONSOLE_LINES}.0")
        self.console.see("end")
        with contextlib.suppress(Exception):
            LOG.info(msg)

    def _save_log(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text", "*.txt")],
            initialfile="weather_log.txt",
        )
        if not path:
            return
        try:
            Path(path).write_text(self.console.get("1.0", "end"),
                                  encoding="utf-8")
            self.status_var.set(f"Log saved → {path}")
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc))

    def _open_app_log(self) -> None:
        if APP_LOG_PATH.exists():
            with contextlib.suppress(Exception):
                webbrowser.open(APP_LOG_PATH.resolve().as_uri())
        else:
            messagebox.showinfo("Log", f"No log at {APP_LOG_PATH.resolve()}")

    def _show_about(self) -> None:
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} v{APP_VERSION}\n\n"
            f"Tracks red/orange GDACS alerts with dedupe.\n\n"
            f"Data: {API_URL}\n"
            f"Settings: {SETTINGS_PATH.resolve()}\n"
            f"Log: {LOG_PATH.resolve()}\n"
            f"App log: {APP_LOG_PATH.resolve()}",
        )

    def _validate_settings(self) -> Optional[str]:
        try:
            lookback = int(self.lookback_var.get())
            dedupe = int(self.dedupe_var.get())
            page = int(self.page_var.get())
            auto = int(self.auto_var.get())
            cache = int(self.cache_var.get())
        except (tk.TclError, ValueError):
            return "All numeric settings must be integers."
        if not (1 <= lookback <= 365):
            return "Lookback days must be between 1 and 365."
        if not (1 <= dedupe <= 365):
            return "Dedupe window must be between 1 and 365."
        if not (10 <= page <= 500):
            return "Page size must be between 10 and 500."
        if not (0 <= auto <= 1440):
            return "Auto-refresh must be between 0 and 1440."
        if not (0 <= cache <= 1440):
            return "Cache TTL must be between 0 and 1440."
        if not any(v.get() for v in self.alert_vars.values()):
            return "Select at least one alert level."
        if not any(v.get() for v in self.type_vars.values()):
            return "Select at least one event type."
        return None

    def _save_settings(self) -> None:
        err = self._validate_settings()
        if err:
            self.settings_hint.configure(text=err)
            with contextlib.suppress(Exception):
                messagebox.showerror("Invalid settings", err)
            return
        self.settings_hint.configure(text="")
        self.settings.lookback_days = int(self.lookback_var.get())
        self.settings.dedupe_window_days = int(self.dedupe_var.get())
        self.settings.page_size = int(self.page_var.get())
        self.settings.auto_refresh_minutes = int(self.auto_var.get())
        self.settings.http_cache_minutes = int(self.cache_var.get())
        self.settings.alert_levels = [
            a for a, v in self.alert_vars.items() if v.get()
        ]
        self.settings.event_types = [
            t for t, v in self.type_vars.items() if v.get()
        ]
        self.settings.auto_export_csv = self.auto_export_var.get()
        self.settings.auto_export_dir = self.auto_export_dir_var.get()
        self.settings.show_flags = self.show_flags_var.get()
        self.settings.save(SETTINGS_PATH)
        self._schedule_auto_refresh()
        self.status_var.set("Settings saved.")
        self._log("Settings saved.", tag="ok")

    def _reset_settings(self) -> None:
        defaults = Settings()
        self.lookback_var.set(defaults.lookback_days)
        self.dedupe_var.set(defaults.dedupe_window_days)
        self.page_var.set(defaults.page_size)
        self.auto_var.set(defaults.auto_refresh_minutes)
        self.cache_var.set(defaults.http_cache_minutes)
        for a, v in self.alert_vars.items():
            v.set(a in defaults.alert_levels)
        for t, v in self.type_vars.items():
            v.set(t in defaults.event_types)
        self.auto_export_var.set(defaults.auto_export_csv)
        self.auto_export_dir_var.set(defaults.auto_export_dir)
        self.show_flags_var.set(defaults.show_flags)
        self.status_var.set("Defaults loaded — press Save to apply.")

    def undo_last_fetch(self) -> None:
        try:
            weekly_logs = load_log(LOG_PATH)
        except Exception as exc:
            messagebox.showerror("Undo", f"Could not read log: {exc}")
            return
        if not weekly_logs:
            messagebox.showinfo("Undo", "Nothing to undo.")
            return
        if not messagebox.askyesno(
            "Undo last fetch",
            f"Remove the most recent week ({len(weekly_logs[-1])} events)?",
        ):
            return
        weekly_logs.pop()
        try:
            save_log(LOG_PATH, weekly_logs)
        except OSError as exc:
            messagebox.showerror("Undo", str(exc))
            return
        self._df = build_dataframe(weekly_logs)
        self._df_view = self._df
        self._refresh_all_views()
        self._log("Undid last fetch.", tag="warn")
        self.status_var.set("Undid last fetch.")

    def import_merge_json(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Import", f"Could not read file: {exc}")
            return
        events: List[Dict[str, Any]] = []
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    events.append(item)
                elif isinstance(item, list):
                    events.extend(x for x in item if isinstance(x, dict))
        else:
            messagebox.showerror("Import", "Unrecognized JSON structure.")
            return
        if not events:
            messagebox.showinfo("Import", "No events found in file.")
            return
        weekly_logs = load_log(LOG_PATH)
        weekly_logs.append(events)
        try:
            save_log(LOG_PATH, weekly_logs)
        except OSError as exc:
            messagebox.showerror("Import", str(exc))
            return
        self._df = build_dataframe(weekly_logs)
        self._df_view = self._df
        self._refresh_all_views()
        self._log(f"Imported {len(events)} events from {path}", tag="ok")
        self.status_var.set(f"Imported {len(events)} events.")

    def copy_to_clipboard(self) -> None:
        if self._df_view.empty:
            messagebox.showinfo("Copy", "Nothing to copy.")
            return
        buf = io.StringIO()
        self._df_view.to_csv(buf, index=False)
        self.clipboard_clear()
        self.clipboard_append(buf.getvalue())
        self.status_var.set("Copied table to clipboard.")
        self._log("Copied table to clipboard.", tag="ok")

    def export_csv(self) -> None:
        if self._df_view.empty:
            messagebox.showinfo("Export", "Nothing to export yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="gdacs_events.csv",
        )
        if not path:
            return
        try:
            self._df_view.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL)
            self.status_var.set(f"Exported → {path}")
            self._log(f"CSV exported → {path}", tag="ok")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_xlsx(self) -> None:
        if self._df_view.empty:
            messagebox.showinfo("Export", "Nothing to export yet.")
            return
        if not HAS_XLSX:
            messagebox.showerror(
                "Export failed",
                "Install 'openpyxl' for XLSX support (pip install openpyxl).",
            )
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="gdacs_events.xlsx",
        )
        if not path:
            return
        try:
            self._df_view.to_excel(path, index=False)
            self.status_var.set(f"Exported → {path}")
            self._log(f"XLSX exported → {path}", tag="ok")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_json(self) -> None:
        if self._df_view.empty:
            messagebox.showinfo("Export", "Nothing to export yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="gdacs_events.json",
        )
        if not path:
            return
        try:
            df = self._df_view.copy()
            for col in ("fromdate", "todate"):
                if col in df:
                    df[col] = df[col].astype(str)
            payload = json.loads(df.to_json(orient="records",
                                            force_ascii=False,
                                            date_format="iso"))
            _atomic_write_json(Path(path), payload)
            self.status_var.set(f"Exported → {path}")
            self._log(f"JSON exported → {path}", tag="ok")
        except (OSError, TypeError, ValueError) as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_markdown(self) -> None:
        if self._df_view.empty:
            messagebox.showinfo("Export", "Nothing to export yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown", "*.md")],
            initialfile="gdacs_report.md",
        )
        if not path:
            return
        try:
            md = self._build_markdown_report()
            _atomic_write_text(Path(path), md)
            self.status_var.set(f"Exported → {path}")
            self._log(f"Markdown report → {path}", tag="ok")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc))

    def _build_markdown_report(self) -> str:
        df = self._df_view
        lines: List[str] = [
            f"# {APP_NAME} — Report",
            "",
            f"- Generated: `{datetime.now(timezone.utc).isoformat()}`",
            f"- Rows: **{len(df)}**",
            "",
            "## Summary",
        ]
        if "alertlevel" in df:
            lines += ["### Alert levels", "", "| Alert | Count |", "|---|---|"]
            for k, v in df["alertlevel"].value_counts().items():
                lines.append(f"| {k} | {int(v)} |")
            lines.append("")
        if "eventtype" in df:
            lines += ["### Event types", "", "| Type | Count |", "|---|---|"]
            for k, v in df["eventtype"].value_counts().items():
                lines.append(f"| {k} | {int(v)} |")
            lines.append("")
        if "country" in df:
            lines += ["### Top countries", "", "| Country | Count |", "|---|---|"]
            for k, v in df["country"].value_counts().head(15).items():
                lines.append(f"| {k} | {int(v)} |")
            lines.append("")
        lines += ["## Events table", ""]
        head = df.head(200).copy()
        for col in ("fromdate", "todate"):
            if col in head:
                head[col] = head[col].astype(str)
        try:
            lines.append(head.to_markdown(index=False))
        except Exception:
            lines.append("```\n" + head.to_string(index=False) + "\n```")
        return "\n".join(lines) + "\n"

    def export_analysis(self) -> None:
        if not ANALYSIS_PATH.exists():
            messagebox.showinfo("Export", "Analysis file not found yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text", "*.txt")],
            initialfile="log_analysis.txt",
        )
        if not path:
            return
        try:
            _atomic_write_text(Path(path),
                               ANALYSIS_PATH.read_text(encoding="utf-8"))
            self.status_var.set(f"Analysis exported → {path}")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc))

    def _on_close(self) -> None:
        try:
            self._cancel_event.set()
            if self._auto_refresh_job:
                self.after_cancel(self._auto_refresh_job)
            if self._debounce_job:
                self.after_cancel(self._debounce_job)
            if self._clock_job:
                self.after_cancel(self._clock_job)
            self.settings.save(SETTINGS_PATH)
            self._save_geometry()
        finally:
            self.destroy()

def _install_excepthook(app: tk.Tk) -> None:
    def hook(exc_type, exc, tb) -> None:
        msg = "".join(traceback.format_exception(exc_type, exc, tb))
        LOG.error("Unhandled exception:\n%s", msg)
        try:
            messagebox.showerror(
                "Unexpected error",
                f"{exc}\n\nSee weather_app.log for details.",
            )
        except Exception:
            print(msg, file=sys.stderr)
    sys.excepthook = hook

def main() -> None:
    app = WeatherApp()
    _install_excepthook(app)
    app.mainloop()

if __name__ == "__main__":
    main()
