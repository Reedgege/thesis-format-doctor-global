"""Local persistence: recently used "manual questionnaire templates" and AI-imported config files.

Stored under ``~/.paperformatpro/recents.json`` (overridable via ``TFD_RECENTS_FILE``,
the same isolation strategy as ``TFD_SETTINGS_FILE`` for test separation).

- ``manual``: recently saved questionnaires (each = one clickable "template"). Deduplicated by
  content, sorted newest-first, capped at ``MAX`` entries.
- ``ai``: recently imported AI config file paths. Deduplicated by path, sorted newest-first,
  capped at ``MAX`` entries.
- ``last_source``: the source mode selected last time (src_manual / src_ai / ...), used to
  auto-restore the last session on startup.

This module does **not import tkinter** — all functions are pure data, unit-testable offline
(same discipline as i18n.py).
"""

from __future__ import annotations

import json
import os
import shutil
import time
from pathlib import Path

MAX = 5

RECENTS_FILE = Path.home() / ".paperformatpro" / "recents.json"
RECENTS_FILE_LEGACY = Path.home() / ".thesis-format-doctor-global" / "recents.json"


def recents_file() -> Path:
    """Actual recents file path (read env each call, for runtime switch / test isolation)."""
    env = os.environ.get("TFD_RECENTS_FILE")
    if env:
        return Path(env)
    if not RECENTS_FILE.exists() and RECENTS_FILE_LEGACY.exists():
        try:
            RECENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(RECENTS_FILE_LEGACY, RECENTS_FILE)
        except Exception:
            return RECENTS_FILE_LEGACY
    return RECENTS_FILE


def _empty() -> dict:
    return {"manual": [], "ai": [], "last_source": ""}


def load_recents() -> dict:
    """Read recents; any error falls back to an empty structure (recents never blocks startup)."""
    try:
        f = recents_file()
        if f.is_file():
            raw = json.loads(f.read_text(encoding="utf-8"))
            if not isinstance(raw, dict):
                return _empty()
            out = _empty()
            out["manual"] = [x for x in (raw.get("manual") or [])
                             if isinstance(x, dict) and "data" in x]
            out["ai"] = [x for x in (raw.get("ai") or [])
                         if isinstance(x, dict) and x.get("path")]
            out["last_source"] = raw.get("last_source") or ""
            return out
    except Exception:
        pass
    return _empty()


def save_recents(recents: dict) -> bool:
    """Write recents (atomic); returns False on failure — not remembering is non-fatal."""
    try:
        f = recents_file()
        f.parent.mkdir(parents=True, exist_ok=True)
        tmp = f.with_name(f.name + ".tmp")
        tmp.write_text(json.dumps(recents, ensure_ascii=False, sort_keys=True),
                       encoding="utf-8")
        os.replace(str(tmp), str(f))
        return True
    except Exception:
        return False


def _canon(data) -> str:
    """Canonicalize a questionnaire dict to a comparable string (for content dedup)."""
    try:
        return json.dumps(data, ensure_ascii=False, sort_keys=True)
    except Exception:
        return repr(data)


def add_manual(recents: dict, data: dict, *, max_n: int = MAX) -> dict:
    """Add a saved questionnaire to "recent templates".

    Deduplicate by content, sort newest-first, keep at most ``max_n``. An empty dict is also
    recorded (represents a "blank template") but identical content dedups, so it won't pile up.
    Returns the updated recents (does not mutate the argument).
    """
    recents = dict(recents)
    items = list(recents.get("manual") or [])
    key = _canon(data)
    now = time.time()
    kept = [x for x in items if _canon(x.get("data")) != key]
    kept.insert(0, {"ts": now, "data": data})
    kept.sort(key=lambda x: x.get("ts", 0), reverse=True)
    recents["manual"] = kept[:max_n]
    return recents


def add_ai(recents: dict, path: str, *, max_n: int = MAX) -> dict:
    """Add an imported AI config file path to "recent imports".

    Deduplicate by normalized path, sort newest-first, keep at most ``max_n``.
    """
    recents = dict(recents)
    items = list(recents.get("ai") or [])
    norm = os.path.normpath(path)
    now = time.time()
    kept = [x for x in items if os.path.normpath(x.get("path", "")) != norm]
    kept.insert(0, {"ts": now, "path": path})
    kept.sort(key=lambda x: x.get("ts", 0), reverse=True)
    recents["ai"] = kept[:max_n]
    return recents


def set_last_source(recents: dict, mode: str) -> dict:
    """Record the source mode selected last (for auto-restore on startup)."""
    recents = dict(recents)
    recents["last_source"] = mode
    return recents


def remove_ai(recents: dict, path: str) -> dict:
    """Drop a (missing) AI path from history."""
    recents = dict(recents)
    norm = os.path.normpath(path)
    recents["ai"] = [x for x in (recents.get("ai") or [])
                     if os.path.normpath(x.get("path", "")) != norm]
    return recents


def first_manual(recents: dict):
    """Most recent manual template's data, or None if none."""
    items = recents.get("manual") or []
    if not items:
        return None
    return items[0].get("data")


def first_existing_ai(recents: dict):
    """Most recent AI path that still exists on disk, or None."""
    for x in (recents.get("ai") or []):
        p = x.get("path")
        if p and os.path.isfile(p):
            return p
    return None


def _fmt(v) -> str:
    try:
        f = float(v)
        return "%g" % f
    except Exception:
        return str(v)


def manual_summary(data: dict, lang: str = "en") -> str:
    """Render a questionnaire dict into a one-line clickable label (bilingual).

    Takes only a few representative fields to avoid being too long. No tkinter dependency,
    pure string join. English mode never mixes in Chinese (overseas build hard rule).
    """
    zh = (lang == "zh")
    if not data:
        return "空白模板" if zh else "Empty template"
    parts = []
    fam = data.get("font_family")
    sz = data.get("font_size_pt")
    if fam and sz:
        parts.append("%s %spt" % (fam, _fmt(sz)))
    elif fam:
        parts.append(str(fam))
    elif sz:
        parts.append("%spt" % _fmt(sz))
    mt = data.get("margin_top_in")
    if mt:
        parts.append(("%sin margins" % _fmt(mt)) if not zh else ("%s英寸页边距" % _fmt(mt)))
    cols = data.get("columns")
    if cols == "2":
        parts.append("双栏" if zh else "2 columns")
    elif cols == "1":
        parts.append("单栏" if zh else "1 column")
    hf = data.get("heading_formats") or {}
    if hf:
        parts.append(("%d levels" % len(hf)) if not zh else ("%d级标题" % len(hf)))
    if not parts:
        parts.append(("%d fields" % len(data)) if not zh else ("%d项" % len(data)))
    return " · ".join(parts)
