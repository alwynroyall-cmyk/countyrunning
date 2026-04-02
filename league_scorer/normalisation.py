"""
All normalisation helpers:
  - Gender    : M / F
    - Category  : Sen / V35 / V40 / V45 / V50 / V55 / V60 / V65 / V70+ / Jun
  - Time      : parse to seconds; display string
  - Columns   : Chip > Gun > first 'time' column
"""

import datetime
import math
import re
import logging
from typing import Optional, Tuple

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────── gender ──────

_MALE = {"m", "male", "open", "o"}
_FEMALE = {"f", "female"}


def normalise_gender(raw) -> Optional[str]:
    """Return 'M', 'F', or None if unrecognisable."""
    if raw is None or _is_nan(raw):
        return None
    val = str(raw).strip().lower()
    if val in _MALE:
        return "M"
    if val in _FEMALE:
        return "F"
    log.warning("Unrecognised gender value: '%s'", raw)
    return None


# ────────────────────────────────────────────────────────────── category ─────

def normalise_category(raw) -> Tuple[str, str]:
    """
    Normalise to: Sen | V35 | V40 | V45 | V50 | V55 | V60 | V65 | V70+ | Jun
    Returns (normalised, notes).
    """
    if raw is None or _is_nan(raw) or str(raw).strip() == "":
        log.warning("Blank/missing category — defaulting to Sen")
        return "Sen", "Missing — defaulted to Sen"

    original = str(raw).strip()
    work = original.lower()

    # Handle age ranges first (before compacting digits), e.g. "Ages 35 - 44".
    age_range = re.search(r"(?:ages?\s*)?(\d{2})\s*(?:-|to)\s*(\d{2})", work)
    if age_range:
        low = int(age_range.group(1))
        if low <= 20:
            return "Jun", ""
        if low >= 35:
            return _age_to_v_category(low), ""

    # Remove gender prefix letter when followed by 'v' or a digit
    # e.g. MV40 -> v40, FV50 -> v50, M40 -> 40
    work = re.sub(r"^[mfwlu](?=[v\d])", "", work)
    # Remove whitespace, hyphens, underscores, dots
    work = re.sub(r"[\s\-_\.]", "", work)

    # ── Juniors (must test before veteran digit catch-all) ──
    if re.match(r"(jun|junior|youth|cadet)", work):
        return "Jun", ""

    # U<number> where number ≤ 20
    m = re.match(r"u(\d+)$", work)
    if m and int(m.group(1)) <= 20:
        return "Jun", ""

    # ── Extract numeric age band ──
    # Handles V35..V70+, Vet50, V 60, FV45, F35, 70+ etc.
    m = re.search(r"v(?:et)?(\d{2})(\+)?", work)
    if not m:
        m = re.search(r"(?<!\d)(\d{2})(\+)?(?!\d)", work)

    if m:
        age = int(m.group(1))
        has_plus = bool(m.group(2))
        if age >= 35:
            if age >= 70 and has_plus:
                return "V70+", ""
            return _age_to_v_category(age), ""
        if age <= 20:
            return "Jun", ""

    # ── Senior synonyms ──
    if re.search(r"(sen|senior|open|elite|adult)", work):
        return "Sen", ""

    # ── Unrecognised ──
    # Leave unresolved categories unchanged so we do not misclassify as Senior.
    log.warning("Unrecognised category '%s' — preserving original", original)
    return original, f"Unrecognised '{original}' — preserved"


def _age_to_v_category(age: int) -> str:
    """Map veteran age to nearest EA-style 5-year master category."""
    if age >= 70:
        return "V70+"
    if age >= 65:
        return "V65"
    if age >= 60:
        return "V60"
    if age >= 55:
        return "V55"
    if age >= 50:
        return "V50"
    if age >= 45:
        return "V45"
    if age >= 40:
        return "V40"
    return "V35"


# ───────────────────────────────────────────────────────────────── time ──────

def parse_time_to_seconds(value) -> Optional[float]:
    """
    Parse a cell value to float seconds.
    Accepts: datetime.time, datetime.timedelta, Excel fraction-of-day float,
             hh:mm:ss[.xxx], mm:ss[.xxx].
    Returns None for invalid / missing values.
    """
    if value is None or _is_nan(value):
        return None

    # Native Python time object (from openpyxl/pandas)
    if isinstance(value, datetime.time):
        return (value.hour * 3600
                + value.minute * 60
                + value.second
                + value.microsecond / 1_000_000)

    # timedelta (some pandas versions return this)
    if isinstance(value, datetime.timedelta):
        return value.total_seconds()

    # datetime (should be rare for pure time cells)
    if isinstance(value, datetime.datetime):
        return (value.hour * 3600
                + value.minute * 60
                + value.second
                + value.microsecond / 1_000_000)

    # Excel fraction-of-day float (0 < v < 1)
    if isinstance(value, (int, float)) and 0 < float(value) < 1:
        return float(value) * 86_400.0

    # String formats
    s = str(value).strip()
    if not s or s.lower() in ("nan", "n/a", "none", "-", "dns", "dnf", "dq", "dsq"):
        return None

    # hh:mm:ss[.micro]
    m = re.fullmatch(r"(\d{1,3}):(\d{2}):(\d{2})(?:[.,](\d+))?", s)
    if m:
        h, mi, sc = int(m.group(1)), int(m.group(2)), int(m.group(3))
        frac = float("0." + m.group(4)) if m.group(4) else 0.0
        return h * 3600 + mi * 60 + sc + frac

    # mm:ss[.micro]
    m = re.fullmatch(r"(\d{1,3}):(\d{2})(?:[.,](\d+))?", s)
    if m:
        mi, sc = int(m.group(1)), int(m.group(2))
        frac = float("0." + m.group(3)) if m.group(3) else 0.0
        return mi * 60 + sc + frac

    return None


def time_display(value) -> str:
    """
    Return a human-readable display string for a time cell value.
    String values are returned as-is (spec: no reformatting).
    Native time objects and Excel floats are converted to hh:mm:ss[.mmm].
    """
    if isinstance(value, datetime.time):
        base = f"{value.hour:02d}:{value.minute:02d}:{value.second:02d}"
        if value.microsecond:
            return f"{base}.{value.microsecond // 1000:03d}"
        return base

    if isinstance(value, datetime.timedelta):
        ts = int(value.total_seconds())
        h, m, s = ts // 3600, (ts % 3600) // 60, ts % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    if isinstance(value, datetime.datetime):
        v = value
        base = f"{v.hour:02d}:{v.minute:02d}:{v.second:02d}"
        return f"{base}.{v.microsecond // 1000:03d}" if v.microsecond else base

    if isinstance(value, (int, float)):
        try:
            f = float(value)
            if 0 < f < 1:
                ts = int(f * 86_400)
                h, m, s = ts // 3600, (ts % 3600) // 60, ts % 60
                return f"{h:02d}:{m:02d}:{s:02d}"
        except (ValueError, TypeError):
            pass

    return str(value).strip()


# ─────────────────────────────────────────────────────── column selection ────

def find_time_column(columns) -> Optional[str]:
    """
    Return the best time column name using priority:
      1. Chip Time  2. Gun Time  3. First column whose name contains 'time'
    """
    lower_map = {c.lower().strip(): c for c in columns}

    for key in ("chip time", "chiptime", "chip"):
        if key in lower_map:
            return lower_map[key]

    for key in ("gun time", "guntime", "gun", "net time", "nettime", "nett time"):
        if key in lower_map:
            return lower_map[key]

    for col in columns:
        if "time" in col.lower():
            return col

    return None


# ─────────────────────────────────────────────────────────────── helpers ─────

def _is_nan(v) -> bool:
    try:
        return math.isnan(float(v))
    except (TypeError, ValueError):
        return False
