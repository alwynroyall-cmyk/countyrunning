"""
timeline_generator.py - Generate a WRRL Season Timeline PNG using Pillow.

Usage (from events_viewer button):
    from .timeline_generator import generate_timeline
    img = generate_timeline(schedule, year, output_path, images_dir)
"""

from __future__ import annotations

import calendar
import datetime
import re
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

from ..events_loader import EventsSchedule, EventEntry

# ---------------------------------------------------------------------------
# Colours  (RGB tuples)
# ---------------------------------------------------------------------------
NAVY       = (58,  70,  88)
GREEN      = (45, 122,  74)
AMBER      = (220, 120,  30)
WHITE      = (255, 255, 255)
NEAR_BLACK = ( 30,  30,  40)
MID_GREY   = (160, 170, 180)
DARK_GREY  = (100, 110, 120)
LIGHT_BG   = (248, 249, 251)

STATUS_COLOUR = {
    "confirmed":   GREEN,
    "provisional": AMBER,
    "tbc":        (127, 140, 141),
}

# Series band colours (up to 6 series)
SERIES_PALETTE = [
    ( 41, 128, 185),  # blue
    (142,  68, 173),  # purple
    (211,  84,   0),  # burnt orange
    ( 26, 188, 156),  # teal
    (192,  57,  43),  # red
    ( 39, 174,  96),  # emerald
]

# ---------------------------------------------------------------------------
# Canvas layout  (pixels)
# ---------------------------------------------------------------------------
IMG_W        = 1500
IMG_H        = 700
HEADER_H     = 130
TL_Y         = 330          # y of main timeline
TL_X1        = 130          # left end
TL_X2        = 1370         # right end
ABOVE_HIGH_Y = 140          # top of tall (even-index) single-event labels
ABOVE_LOW_Y  = 225          # top of short (odd-index) single-event labels
BAND_Y0      = 352          # y of first series band
BAND_STEP    = 27           # y gap between series bands
LEGEND_Y     = 570          # top of legend row


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fonts() -> dict:
    """Return a dict of ImageFont instances; fall back to default if unavailable."""
    _bold_paths    = ["C:/Windows/Fonts/segoeuib.ttf",  "C:/Windows/Fonts/arialbd.ttf"]
    _regular_paths = ["C:/Windows/Fonts/segoeui.ttf",   "C:/Windows/Fonts/arial.ttf"]

    def _f(paths, size):
        for p in paths:
            try:
                return ImageFont.truetype(p, size)
            except (OSError, IOError):
                pass
        return ImageFont.load_default()

    return {
        "title":      _f(_bold_paths,    30),
        "subtitle":   _f(_regular_paths, 14),
        "label_bold": _f(_bold_paths,    13),
        "label":      _f(_regular_paths, 12),
        "small":      _f(_regular_paths, 11),
        "badge":      _f(_bold_paths,    11),
        "month":      _f(_regular_paths, 11),
    }


def _parse_dates(text: str) -> list[datetime.date]:
    """Parse one or more 'Nth Month YYYY' dates from a comma-separated string."""
    result = []
    for part in re.split(r",\s*", str(text).strip()):
        cleaned = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", part.strip())
        try:
            result.append(datetime.datetime.strptime(cleaned, "%d %B %Y").date())
        except ValueError:
            pass
    return result


def _date_x(d: datetime.date, start: datetime.date, end: datetime.date) -> int:
    """Convert a date to an x pixel coordinate on the timeline."""
    total = max((end - start).days, 1)
    offset = (d - start).days
    return TL_X1 + int(offset / total * (TL_X2 - TL_X1))


def _short(name: str, max_len: int = 17) -> str:
    return name if len(name) <= max_len else name[: max_len - 1] + "\u2026"


def _status_col(status: str) -> tuple:
    return STATUS_COLOUR.get(status.lower(), STATUS_COLOUR["tbc"])


def _centred_text(draw: ImageDraw.Draw, cx: int, y: int, text: str,
                  font: ImageFont.FreeTypeFont, fill: tuple) -> None:
    bb = draw.textbbox((0, 0), text, font=font)
    draw.text((cx - (bb[2] - bb[0]) // 2, y), text, font=font, fill=fill)


def _pill(draw: ImageDraw.Draw, cx: int, cy: int, text: str,
          font: ImageFont.FreeTypeFont, bg: tuple, fg: tuple = WHITE,
          pad_x: int = 8, pad_y: int = 4) -> None:
    """Draw a rounded-rectangle badge centred at (cx, cy)."""
    bb = draw.textbbox((0, 0), text, font=font)
    w, h = bb[2] - bb[0] + pad_x * 2, bb[3] - bb[1] + pad_y * 2
    x0, y0 = cx - w // 2, cy - h // 2
    draw.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=4, fill=bg)
    draw.text((x0 + pad_x, y0 + pad_y), text, font=font, fill=fg)


# ---------------------------------------------------------------------------
# Main generator
# ---------------------------------------------------------------------------

def generate_timeline(
    schedule: EventsSchedule,
    year: int,
    output_path: Optional[Path] = None,
    images_dir: Optional[Path] = None,
) -> Image.Image:
    """
    Render the season timeline for *schedule* and return a PIL Image.

    Parameters
    ----------
    schedule :
        Loaded EventsSchedule.
    year :
        Season year (used in heading).
    output_path :
        If given, the PNG is also saved here.
    images_dir :
        Directory containing 'WRRL shield concept.png' for the header logo.

    Returns
    -------
    PIL.Image.Image
        The rendered timeline.
    """
    F = _fonts()

    # ------------------------------------------------------------------ #
    # Classify events and parse their dates
    # ------------------------------------------------------------------ #
    single_events: list[tuple[EventEntry, list[datetime.date]]] = []
    series_events: list[tuple[EventEntry, list[datetime.date]]] = []

    for ev in schedule.events:
        dates = _parse_dates(ev.scheduled_dates)
        if not dates:
            continue
        if ev.date_type == "Series":
            series_events.append((ev, dates))
        else:
            single_events.append((ev, dates))

    all_dates: list[datetime.date] = [
        d for (_, ds) in single_events + series_events for d in ds
    ]

    # ------------------------------------------------------------------ #
    # Season date range
    # ------------------------------------------------------------------ #
    if all_dates:
        earliest = min(all_dates)
        latest   = max(all_dates)
        start_m  = max(1, earliest.month - 1)
        start    = datetime.date(earliest.year, start_m, 1)
        end_m    = latest.month + 2
        end_y    = latest.year
        if end_m > 12:
            end_m -= 12
            end_y += 1
        end = datetime.date(end_y, end_m, calendar.monthrange(end_y, end_m)[1])
    else:
        start = datetime.date(year, 3, 1)
        end   = datetime.date(year, 11, 30)

    n_series = len(series_events)
    month_label_y = BAND_Y0 + n_series * BAND_STEP + 12

    # ------------------------------------------------------------------ #
    # Create canvas
    # ------------------------------------------------------------------ #
    img  = Image.new("RGB", (IMG_W, IMG_H), LIGHT_BG)
    draw = ImageDraw.Draw(img)

    # ── Header ────────────────────────────────────────────────────────────
    draw.rectangle([0, 0, IMG_W, HEADER_H], fill=NAVY)
    draw.rectangle([0, HEADER_H - 4, IMG_W, HEADER_H], fill=GREEN)

    # Shield logo
    text_x = 30
    if images_dir:
        shield_path = images_dir / "WRRL shield concept.png"
        if shield_path.exists():
            try:
                logo = Image.open(shield_path).convert("RGBA")
                logo.thumbnail((90, 90))
                logo_y = (HEADER_H - 4 - logo.height) // 2
                img.paste(logo, (20, logo_y), logo.split()[3])
                text_x = 125
            except Exception:
                pass

    draw.text((text_x, 22), f"{year} WRRL Season Timeline",
              font=F["title"], fill=WHITE)
    draw.text((text_x, 62), f"Championship Events  \u00b7  {len(schedule.events)} races",
              font=F["subtitle"], fill=(155, 175, 205))

    # ── Background of content area ────────────────────────────────────────
    draw.rectangle([0, HEADER_H, IMG_W, IMG_H], fill=WHITE)

    # ── Main timeline line ────────────────────────────────────────────────
    draw.line([(TL_X1, TL_Y), (TL_X2 - 10, TL_Y)], fill=NAVY, width=3)
    # Arrow head
    draw.polygon([
        (TL_X2, TL_Y),
        (TL_X2 - 12, TL_Y - 5),
        (TL_X2 - 12, TL_Y + 5),
    ], fill=NAVY)

    # ── Month markers ─────────────────────────────────────────────────────
    cur = datetime.date(start.year, start.month, 1)
    while cur <= end:
        mx = _date_x(cur, start, end)
        draw.line([(mx, TL_Y - 6), (mx, TL_Y + 6)], fill=MID_GREY, width=1)
        bb = draw.textbbox((0, 0), cur.strftime("%b"), font=F["month"])
        tw = bb[2] - bb[0]
        draw.text((mx - tw // 2, month_label_y),
                  cur.strftime("%b"), font=F["month"], fill=DARK_GREY)
        # advance month
        m, y2 = cur.month + 1, cur.year
        if m > 12:
            m, y2 = 1, y2 + 1
        cur = datetime.date(y2, m, 1)

    # ── Today marker ──────────────────────────────────────────────────────
    TODAY       = datetime.date.today()
    TODAY_RED   = (210, 30, 30)
    TODAY_RED_L = (240, 100, 100)   # lighter for outer ring

    if start <= TODAY <= end:
        tx = _date_x(TODAY, start, end)

        # Outer ring (unfilled circle)
        draw.ellipse([(tx - 14, TL_Y - 14), (tx + 14, TL_Y + 14)],
                     outline=TODAY_RED, width=2)
        # Middle ring
        draw.ellipse([(tx - 9, TL_Y - 9), (tx + 9, TL_Y + 9)],
                     outline=TODAY_RED_L, width=1)
        # Filled centre dot
        draw.ellipse([(tx - 5, TL_Y - 5), (tx + 5, TL_Y + 5)],
                     fill=TODAY_RED, outline=WHITE, width=1)

        # "Today" label just below the marker
        today_label = f"Today  {TODAY.day} {TODAY.strftime('%b')}"
        bb = draw.textbbox((0, 0), today_label, font=F["small"])
        lw = bb[2] - bb[0]
        draw.text((tx - lw // 2, TL_Y + 18), today_label, font=F["small"], fill=TODAY_RED)

    # ── Series events (coloured bands below timeline) ─────────────────────
    for i, (ev, dates) in enumerate(series_events):
        col   = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        band_y = BAND_Y0 + i * BAND_STEP
        xs    = sorted(_date_x(d, start, end) for d in dates)

        # Band line between first and last date
        draw.line([(xs[0], band_y), (xs[-1], band_y)], fill=col, width=4)

        # Individual date nodes
        for x in xs:
            draw.ellipse([(x - 5, band_y - 5), (x + 5, band_y + 5)],
                         fill=col, outline=WHITE, width=1)

        # Series name label — left of first dot if space, else right of last
        label = _short(ev.event_name, 22)
        bb    = draw.textbbox((0, 0), label, font=F["small"])
        lw    = bb[2] - bb[0]
        lx    = xs[0] - lw - 10
        if lx < TL_X1:
            lx = xs[-1] + 38   # push right of last dot (leave room for badge)
        draw.text((lx, band_y - 8), label, font=F["small"], fill=col)

        # Distance badge — always after the last dot
        bx       = xs[-1] + 10
        dist_bb  = draw.textbbox((0, 0), ev.distance, font=F["badge"])
        bw       = dist_bb[2] - dist_bb[0] + 10
        bh       = dist_bb[3] - dist_bb[1] + 6
        draw.rounded_rectangle(
            [bx, band_y - bh // 2, bx + bw, band_y + bh // 2],
            radius=3, fill=col,
        )
        draw.text((bx + 5, band_y - bh // 2 + 3), ev.distance, font=F["badge"], fill=WHITE)

        # Thin vertical connector from timeline to band start
        lighter = tuple(min(255, c + 80) for c in col)
        draw.line([(xs[0], TL_Y + 8), (xs[0], band_y - 6)],
                  fill=lighter, width=1)

    # ── Single-date events (labels above timeline) ────────────────────────
    for i, (ev, dates) in enumerate(single_events):
        d     = dates[0]
        x     = _date_x(d, start, end)
        nc    = _status_col(ev.status)
        high  = (i % 2 == 0)
        top_y = ABOVE_HIGH_Y if high else ABOVE_LOW_Y

        # Dashed stem from label bottom to node
        stem_top  = top_y + 58    # below the badge
        stem_col  = tuple(min(255, c + 100) for c in nc)  # lighter shade
        for seg in range(stem_top, TL_Y - 10, 7):
            draw.line([(x, seg), (x, min(seg + 4, TL_Y - 10))],
                      fill=stem_col, width=1)

        # Timeline node
        draw.ellipse([(x - 9, TL_Y - 9), (x + 9, TL_Y + 9)],
                     fill=nc, outline=WHITE, width=2)

        # Event name
        name = _short(ev.event_name, 16)
        _centred_text(draw, x, top_y, name, F["label_bold"], NEAR_BLACK)

        # Date
        date_str = f"{d.day} {d.strftime('%b')}"
        _centred_text(draw, x, top_y + 17, date_str, F["label"], DARK_GREY)

        # Distance pill badge
        _pill(draw, x, top_y + 43, ev.distance, F["badge"], nc)

    # ── Legend ────────────────────────────────────────────────────────────
    draw.line([(TL_X1, LEGEND_Y - 8), (IMG_W - TL_X1, LEGEND_Y - 8)],
              fill=(220, 225, 230), width=1)

    lx = TL_X1

    # Status dots
    for key, label in [("confirmed", "Confirmed"), ("provisional", "Provisional"), ("tbc", "TBC")]:
        col = STATUS_COLOUR[key]
        draw.ellipse([(lx, LEGEND_Y + 2), (lx + 12, LEGEND_Y + 14)],
                     fill=col, outline=WHITE, width=1)
        draw.text((lx + 16, LEGEND_Y), label, font=F["small"], fill=NEAR_BLACK)
        bb = draw.textbbox((0, 0), label, font=F["small"])
        lx += (bb[2] - bb[0]) + 36

    # Series swatches
    for i, (ev, _) in enumerate(series_events):
        col = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        draw.line([(lx, LEGEND_Y + 8), (lx + 22, LEGEND_Y + 8)], fill=col, width=4)
        label = _short(ev.event_name, 22)
        draw.text((lx + 28, LEGEND_Y), label, font=F["small"], fill=NEAR_BLACK)
        bb = draw.textbbox((0, 0), label, font=F["small"])
        lx += (bb[2] - bb[0]) + 50

    # Footer right-align
    footer = f"{len(schedule.events)} Championship Events  |  {year} WRRL Season"
    bb = draw.textbbox((0, 0), footer, font=F["small"])
    draw.text((IMG_W - TL_X1 - (bb[2] - bb[0]), LEGEND_Y),
              footer, font=F["small"], fill=DARK_GREY)

    # ── Save ──────────────────────────────────────────────────────────────
    if output_path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(output_path), "PNG", dpi=(150, 150))

    return img
