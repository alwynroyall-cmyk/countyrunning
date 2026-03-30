"""
report_writer.py — Branded DOCX + PDF season report.

Sections
--------
  1. Cover header  (WRRL shield + season title)
  2. Season overview stats
  3. Division 1 Club Table
  4. Division 2 Club Table
  5. Top 20 Male Individual
  6. Top 20 Female Individual
  7. Category Leaders – top 5 per category (M & F side-by-side)
"""

import logging
import re
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .models import (
    RunnerRaceEntry,
    RunnerSeasonRecord,
    TeamRaceResult,
    TeamSeasonRecord,
    UnrecognisedClub,
)

log = logging.getLogger(__name__)

# ── Brand colours ──────────────────────────────────────────────────────────────
_NAVY_HEX    = "3a4658"
_GREEN_HEX   = "2d7a4a"
_WHITE_HEX   = "ffffff"
_ALT_HEX     = "eef2f7"   # alternating data row tint

_NAVY_RGB   = RGBColor(0x3a, 0x46, 0x58)
_GREEN_RGB  = RGBColor(0x2d, 0x7a, 0x4a)
_WHITE_RGB  = RGBColor(0xff, 0xff, 0xff)
_SUBHDR_RGB = RGBColor(0xa0, 0xb0, 0xc0)

# Award / status colours
_GOLD_HEX      = "c9a84c"   # 1st place — gold
_SILVER_HEX    = "a8b4c4"   # 2nd place — silver
_BRONZE_HEX    = "b87440"   # 3rd place — bronze
_PROMOTED_HEX  = "d0ead8"   # Div 2 promotion zone (subtle green tint)
_RELEGATED_HEX = "f5d8d8"   # Div 1 relegation zone (subtle red tint)
_GOLD_RGB      = RGBColor(0xc9, 0xa8, 0x4c)
_SILVER_RGB    = RGBColor(0xa8, 0xb4, 0xc4)
_BRONZE_RGB    = RGBColor(0xb8, 0x74, 0x40)

_CATEGORIES = ["Sen", "V40", "V50", "V60", "V70+"]

# Per-category badge background colours (hex without #)
_CAT_BADGE_HEX = {
    "Sen":  "2d7a4a",   # brand green
    "V40":  "1a5fa8",   # blue
    "V50":  "c05800",   # burnt orange
    "V60":  "5e3590",   # purple
    "V70+": "a06820",   # gold-brown
}

# row_style → (bg_hex, fg_RGBColor_or_None, bold, italic)
_ROW_STYLES = {
    "leader":   (_GREEN_HEX,    _WHITE_RGB, True,  False),
    "gold":     (_GOLD_HEX,     _NAVY_RGB,  True,  False),
    "silver":   (_SILVER_HEX,   _NAVY_RGB,  True,  False),
    "bronze":   (_BRONZE_HEX,   _WHITE_RGB, True,  False),
    "promoted": (_PROMOTED_HEX, _NAVY_RGB,  False, False),
    "relegated":(_RELEGATED_HEX,_NAVY_RGB,  False, True),
}


# ── Low-level XML helpers ──────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str) -> None:
    """Set (or replace) the background fill of a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_para_bg(para, hex_color: str) -> None:
    """Shade an entire paragraph (fills to page margins — used for section bands)."""
    pPr = para._p.get_or_add_pPr()
    for existing in pPr.findall(qn("w:shd")):
        pPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def _cell_text(cell, text, bold=False, italic=False,
               color: RGBColor = None, size_pt=9, align="left") -> None:
    """Write text into a cell's first paragraph, clearing any prior content."""
    para = cell.paragraphs[0]
    para.clear()
    para.alignment = {
        "left":   WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right":  WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    run = para.add_run("" if text is None or text == "" else str(text))
    run.bold   = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color


# ── Reusable table building blocks ─────────────────────────────────────────────

def _make_header_row(table, headers: list, widths_cm: list) -> None:
    """Style the pre-existing first row as a navy header band."""
    row = table.rows[0]
    for i, hdr in enumerate(headers):
        cell = row.cells[i]
        _set_cell_bg(cell, _NAVY_HEX)
        if i < len(widths_cm) and widths_cm[i]:
            cell.width = Cm(widths_cm[i])
        _cell_text(cell, hdr, bold=True, color=_WHITE_RGB, size_pt=8, align="center")


def _add_data_row(table, values: list, widths_cm: list,
                  alt=False, row_style: str = None, name_col=1) -> None:
    """Append a data row.  *row_style* keys into _ROW_STYLES."""
    row = table.add_row()
    if row_style and row_style in _ROW_STYLES:
        bg, fg, is_bold, is_italic = _ROW_STYLES[row_style]
    else:
        bg, fg, is_bold, is_italic = (_ALT_HEX if alt else _WHITE_HEX), None, False, False
    for i, val in enumerate(values):
        cell = row.cells[i]
        _set_cell_bg(cell, bg)
        if i < len(widths_cm) and widths_cm[i]:
            cell.width = Cm(widths_cm[i])
        align = "left" if i == name_col else "center"
        _cell_text(cell, val, bold=(is_bold or i == 0),
                   italic=is_italic, color=fg, size_pt=9, align=align)


# ── Section heading ────────────────────────────────────────────────────────────

def _page_break(doc: Document) -> None:
    """Insert a hard page break."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)


def _section_heading(doc: Document, text: str, space_before_pt=14) -> None:
    """Full-width green band with white bold caps text — matches the source image style."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Pt(6)
    _set_para_bg(p, _GREEN_HEX)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = _WHITE_RGB


def _sub_heading(doc: Document, text: str) -> None:
    """Smaller green-accented sub-section label."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = _NAVY_RGB


def _spacer(doc: Document, height_pt: float = 10) -> None:
    """Insert an empty paragraph for visual breathing room."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(height_pt)


def _set_table_cell_margins(tbl, top_pt=4, bottom_pt=4, left_pt=6, right_pt=6) -> None:
    """Set uniform top/bottom/left/right cell margins on a table."""
    tbl_elem = tbl._tbl
    tbl_pr = tbl_elem.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tbl_pr)
    for existing in tbl_pr.findall(qn("w:tblCellMar")):
        tbl_pr.remove(existing)
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top_pt), ("left", left_pt),
                      ("bottom", bottom_pt), ("right", right_pt)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(val * 20)))   # twentieths of a point
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)


def _remove_vertical_borders(tbl) -> None:
    """Remove only the vertical (insideV + left + right) borders; keep horizontal lines."""
    tbl_elem = tbl._tbl
    tbl_pr = tbl_elem.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tbl_pr)
    tbl_borders = tbl_pr.find(qn("w:tblBorders"))
    if tbl_borders is None:
        tbl_borders = OxmlElement("w:tblBorders")
        tbl_pr.append(tbl_borders)
    for name in ("left", "right", "insideV"):
        existing = tbl_borders.find(qn(f"w:{name}"))
        if existing is not None:
            tbl_borders.remove(existing)
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tbl_borders.append(el)


def _no_table_borders(tbl) -> None:
    """Remove all visible borders from a table."""
    tbl_elem = tbl._tbl
    tbl_pr = tbl_elem.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tbl_pr)
    for existing in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(existing)
    tbl_borders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tbl_borders.append(el)
    tbl_pr.append(tbl_borders)


# ── Report sections ────────────────────────────────────────────────────────────

def _write_overview(doc: Document, highest: int, year: int,
                    male: List[RunnerSeasonRecord],
                    female: List[RunnerSeasonRecord],
                    unrec: List[UnrecognisedClub],
                    nonleague_filename: str = "") -> None:
    _section_heading(doc, "Season Overview", space_before_pt=6)
    headers    = ["", "Male", "Female", "Total"]
    widths_cm  = [6.0, 3.0, 3.0, 3.0]
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    _make_header_row(t, headers, widths_cm)

    clubs_m = {r.preferred_club for r in male}
    clubs_f = {r.preferred_club for r in female}
    unrec_count = len({u.raw_club_name for u in unrec})
    nonleague_note = (
        f"{unrec_count} club(s) — see {nonleague_filename}" if unrec_count
        else "None"
    )

    rows = [
        ("Season",             str(year),          "",               ""),
        ("Races processed",    str(highest),        "",               ""),
        ("Runners scored",     str(len(male)),       str(len(female)), str(len(male) + len(female))),
        ("Clubs scored",       str(len(clubs_m)),    str(len(clubs_f)),str(len(clubs_m | clubs_f))),
        ("Non-league clubs",   nonleague_note,       "",               ""),
    ]
    for i, row_data in enumerate(rows):
        _add_data_row(t, list(row_data), widths_cm, alt=i % 2 == 1, name_col=0)


def _write_club_table(doc: Document, teams: List[TeamSeasonRecord],
                      title: str, race_count: int,
                      is_div1: bool = False) -> None:
    """Render one club division table.

    Div 1: bottom 2 teams are highlighted as relegated (subtle red, italic).
    Div 2: position 1 = winner (green); position 2 = promoted (subtle teal).
    In both divisions position 1 always gets the green winner highlight.
    """
    _section_heading(doc, title)
    races      = list(range(1, race_count + 1))
    headers    = ["Pos", "Club"] + [f"R{n}" for n in races] + ["Total"]
    name_col   = 1
    widths_cm  = [1.0, 5.5] + [1.2] * len(races) + [1.5]

    sorted_teams = sorted(teams, key=lambda x: x.position)
    total        = len(sorted_teams)
    RELEGATE_N   = 2   # bottom N in Div 1
    PROMOTE_N    = 2   # top N in Div 2 (pos 1 stays as leader)

    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    _make_header_row(t, headers, widths_cm)

    for i, team in enumerate(sorted_teams):
        if team.position == 1:
            style = "leader"
        elif is_div1 and i >= total - RELEGATE_N:
            style = "relegated"
        elif not is_div1 and i < PROMOTE_N:   # pos 2 in Div 2
            style = "promoted"
        else:
            style = None

        vals = [team.position, team.display_name]
        for n in races:
            rr = team.race_results.get(n)
            vals.append(rr.team_points if rr else "")
        vals.append(team.total_points)
        _add_data_row(t, vals, widths_cm, alt=i % 2 == 1,
                      row_style=style, name_col=name_col)

    # Legend note below the table
    legend_p = doc.add_paragraph()
    legend_p.paragraph_format.space_before = Pt(3)
    legend_p.paragraph_format.space_after  = Pt(4)
    if is_div1:
        _legend_chip(legend_p, _RELEGATED_HEX, "Relegated")
    else:
        _legend_chip(legend_p, _PROMOTED_HEX, "Promoted")
        r = legend_p.add_run("  (top 2 clubs promoted to Division 1)")
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x60, 0x70, 0x80)


def _legend_chip(para, hex_color: str, label: str) -> None:
    """Inline coloured square + text label for table legends."""
    # Coloured square via background-shaded run
    run = para.add_run("  ")
    run.font.size = Pt(10)
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    rPr.append(shd)
    run2 = para.add_run(f" {label}")
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x60, 0x70, 0x80)


_MEDAL_STYLES = {1: "gold", 2: "silver", 3: "bronze"}


def _write_individual_table(doc: Document, records: List[RunnerSeasonRecord],
                             title: str, race_count: int, top_n: int = 20) -> None:
    """Top-N individual table.  Positions 1/2/3 get gold/silver/bronze highlights."""
    _section_heading(doc, title)
    races     = list(range(1, race_count + 1))
    headers   = ["Pos", "Name", "Club", "Cat"] + [f"R{n}" for n in races] + ["Total"]
    name_col  = 1
    widths_cm = [0.8, 4.5, 3.5, 1.0] + [1.1] * len(races) + [1.3]

    top = sorted(records, key=lambda r: r.position)[:top_n]

    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    _make_header_row(t, headers, widths_cm)

    for i, rec in enumerate(top):
        vals = [rec.position, rec.name, rec.preferred_club, rec.category]
        for n in races:
            vals.append(rec.race_points.get(n, ""))
        vals.append(rec.total_points)
        style = _MEDAL_STYLES.get(rec.position)
        _add_data_row(t, vals, widths_cm, alt=i % 2 == 1,
                      row_style=style, name_col=name_col)


def _write_category_leaders(doc: Document,
                              male: List[RunnerSeasonRecord],
                              female: List[RunnerSeasonRecord],
                              top_n: int = 3) -> None:
    """Category top-N tables. Position 1 is highlighted as category champion (green).
    A narrow spacer column separates the male and female halves.
    """
    _section_heading(doc, "Category Leaders")
    # Layout: 4 male cols | 1 spacer | 4 female cols  = 9 columns
    headers   = ["#", "Male Name", "Club", "Pts", "", "#", "Female Name", "Club", "Pts"]
    widths_cm = [0.6, 4.0, 3.5, 1.0, 0.4, 0.6, 4.0, 3.5, 1.0]
    name_cols = (1, 6)   # columns that left-align
    spacer_col = 4

    for cat in _CATEGORIES:
        m_top = sorted([r for r in male   if r.category == cat], key=lambda r: r.position)[:top_n]
        f_top = sorted([r for r in female if r.category == cat], key=lambda r: r.position)[:top_n]
        if not m_top and not f_top:
            continue

        _sub_heading(doc, cat)
        t = doc.add_table(rows=1, cols=9)
        t.style = "Table Grid"
        _make_header_row(t, headers, widths_cm)

        for i in range(max(len(m_top), len(f_top))):
            m = m_top[i] if i < len(m_top) else None
            f = f_top[i] if i < len(f_top) else None
            vals = [
                m.position if m else "",
                m.name     if m else "",
                m.preferred_club if m else "",
                m.total_points   if m else "",
                "",                               # spacer
                f.position if f else "",
                f.name     if f else "",
                f.preferred_club if f else "",
                f.total_points   if f else "",
            ]
            # Position 1 = category champion — green leader row
            style = "leader" if i == 0 else None
            row = t.add_row()
            if style:
                bg, fg, is_bold, _ = _ROW_STYLES[style]
            else:
                bg = _ALT_HEX if i % 2 == 1 else _WHITE_HEX
                fg = None
                is_bold = False
            for j, val in enumerate(vals):
                cell = row.cells[j]
                # Spacer column: always white, no text
                if j == spacer_col:
                    _set_cell_bg(cell, _WHITE_HEX)
                    cell.width = Cm(widths_cm[j])
                    continue
                _set_cell_bg(cell, bg)
                if j < len(widths_cm) and widths_cm[j]:
                    cell.width = Cm(widths_cm[j])
                align = "left" if j in name_cols else "center"
                _cell_text(cell, val, bold=(is_bold or j in (0, 5)),
                           color=fg, size_pt=9, align=align)


# ── Per-race report ───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

def _extract_race_name(filepath: Path) -> str:
    """Strip 'Race N -- ' prefix from the filename stem."""
    m = re.match(r'^Race\s+\d+\s*[-\u2013\u2014]+\s*(.+)$', Path(filepath).stem, re.IGNORECASE)
    return m.group(1).strip() if m else Path(filepath).stem


def _build_race_header(doc: Document, race_name: str, race_num: int,
                       total_races: int, images_dir: Optional[Path]) -> None:
    """Big race-name header matching the WRRL scoring table card style."""
    tbl = doc.add_table(rows=1, cols=2)
    logo_cell  = tbl.rows[0].cells[0]
    title_cell = tbl.rows[0].cells[1]
    logo_cell.width  = Cm(2.5)
    title_cell.width = Cm(15.5)
    _set_cell_bg(logo_cell, _NAVY_HEX)
    _set_cell_bg(title_cell, _NAVY_HEX)

    shield_path = (images_dir / "WRRL shield concept.png") if images_dir else None
    if shield_path and shield_path.exists():
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(shield_path), width=Cm(2.2))

    title_cell.paragraphs[0].clear()
    p1 = title_cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.paragraph_format.left_indent  = Cm(0.4)
    p1.paragraph_format.space_before = Pt(6)
    r1 = p1.add_run(race_name.upper())
    r1.bold = True
    r1.font.size = Pt(22)
    r1.font.color.rgb = _WHITE_RGB

    p2 = title_cell.add_paragraph()
    p2.paragraph_format.left_indent = Cm(0.4)
    r2 = p2.add_run(f"RACE {race_num} OF {total_races}")
    r2.bold = True
    r2.font.size = Pt(11)
    r2.font.color.rgb = _SUBHDR_RGB

    accent = doc.add_paragraph()
    accent.paragraph_format.space_before = Pt(0)
    accent.paragraph_format.space_after  = Pt(6)
    _set_para_bg(accent, _GREEN_HEX)


def _write_race_individual(doc: Document,
                           runners: List[RunnerRaceEntry],
                           top_n: int = 10) -> None:
    """Top-N eligible finishers by overall finish time (both genders together)."""
    _section_heading(doc, "Individual Results", space_before_pt=14)

    eligible = sorted(
        [r for r in runners if r.eligible and r.points > 0],
        key=lambda r: r.time_seconds,
    )[:top_n]

    if not eligible:
        doc.add_paragraph("  No eligible results.")
        return

    headers   = ["Pos", "Name", "Club", "Gender", "Cat", "Time", "Points"]
    widths_cm = [1.0, 5.0, 4.5, 1.5, 1.5, 2.5, 2.0]
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    _remove_vertical_borders(t)
    _set_table_cell_margins(t, top_pt=5, bottom_pt=5)
    _make_header_row(t, headers, widths_cm)

    for i, r in enumerate(eligible):
        vals  = [i + 1, r.name, r.preferred_club or "",
                 r.gender, r.normalised_category, r.time_str, r.points]
        style = _MEDAL_STYLES.get(i + 1)
        _add_data_row(t, vals, widths_cm, alt=i % 2 == 1,
                      row_style=style, name_col=1)

    _spacer(doc, height_pt=6)


def _write_race_teams(doc: Document, team_results: List[TeamRaceResult]) -> None:
    """Team results grouped by division — Division 1 left, Division 2 right."""
    _section_heading(doc, "Club Results", space_before_pt=14)

    def _sorted_div(div: int) -> List[TeamRaceResult]:
        return sorted(
            [t for t in team_results if t.division == div],
            key=lambda x: (-x.team_points, -x.team_score),
        )

    div1_teams = _sorted_div(1)
    div2_teams = [t for t in _sorted_div(2) if t.team_score > 0]

    # 13 columns: 6 (Div 1) + 1 (spacer) + 6 (Div 2)
    # Fits A4 portrait 18cm usable: (18.0 - 0.4) / 2 = 8.8cm per side
    col_widths = [0.7, 3.3, 1.2, 1.2, 1.3, 1.1]   # Pos | Club | Men | Women | Score | Pts
    div_w      = 0.4
    DIV_COL    = 6
    COLS       = 13

    t = doc.add_table(rows=2, cols=COLS)
    t.style = "Table Grid"
    _remove_vertical_borders(t)
    _set_table_cell_margins(t, top_pt=5, bottom_pt=5)

    def _pos_list(teams: List[TeamRaceResult]) -> List[int]:
        """Competition-ranked positions (tied teams share position)."""
        out: List[int] = []
        rank = 1
        i = 0
        n = len(teams)
        while i < n:
            j = i + 1
            while j < n and teams[j].team_points == teams[i].team_points:
                j += 1
            for _ in range(i, j):
                out.append(rank)
            rank = j + 1
            i = j
        return out

    pos1 = _pos_list(div1_teams)
    pos2 = _pos_list(div2_teams)

    # Row 0: merged section title headers
    r0 = t.rows[0]
    merged_l = r0.cells[0].merge(r0.cells[5])
    _set_cell_bg(merged_l, _NAVY_HEX)
    _cell_text(merged_l, "Division 1", bold=True, color=_WHITE_RGB, size_pt=10, align="center")
    _set_cell_bg(r0.cells[DIV_COL], _WHITE_HEX)
    r0.cells[DIV_COL].width = Cm(div_w)
    merged_r = r0.cells[7].merge(r0.cells[12])
    _set_cell_bg(merged_r, _NAVY_HEX)
    _cell_text(merged_r, "Division 2", bold=True, color=_WHITE_RGB, size_pt=10, align="center")

    # Row 1: column sub-headers
    r1 = t.rows[1]
    col_hdrs = ["Pos", "Club", "Men", "Women", "Score", "Pts"]
    for j, (hdr, w) in enumerate(zip(col_hdrs, col_widths)):
        c = r1.cells[j]
        _set_cell_bg(c, _NAVY_HEX)
        c.width = Cm(w)
        _cell_text(c, hdr, bold=True, color=_WHITE_RGB, size_pt=8, align="center")
    _set_cell_bg(r1.cells[DIV_COL], _WHITE_HEX)
    r1.cells[DIV_COL].width = Cm(div_w)
    for j, (hdr, w) in enumerate(zip(col_hdrs, col_widths)):
        c = r1.cells[7 + j]
        _set_cell_bg(c, _NAVY_HEX)
        c.width = Cm(w)
        _cell_text(c, hdr, bold=True, color=_WHITE_RGB, size_pt=8, align="center")

    # Data rows
    for i in range(max(len(div1_teams), len(div2_teams), 1)):
        row = t.add_row()
        for teams, positions, offset in (
            (div1_teams, pos1, 0),
            (div2_teams, pos2, 7),
        ):
            team = teams[i] if i < len(teams) else None
            pos  = positions[i] if i < len(positions) else None
            style = "leader" if (team and pos == 1) else None
            bg, fg, is_bold, _ = (
                _ROW_STYLES[style] if style
                else ((_ALT_HEX if i % 2 == 1 else _WHITE_HEX), None, False, False)
            )
            club_label = (
                f"{team.preferred_club} ({team.team_id})" if team else ""
            )
            vals = [
                pos,
                club_label,
                team.men_score   if team else "",
                team.women_score if team else "",
                team.team_score  if team else "",
                team.team_points if team else "",
            ]
            for j, (val, w) in enumerate(zip(vals, col_widths)):
                c = row.cells[offset + j]
                _set_cell_bg(c, bg)
                c.width = Cm(w)
                align = "left" if j == 1 else "center"
                _cell_text(c, val, bold=(is_bold or j == 0), color=fg,
                           size_pt=9, align=align)

        _set_cell_bg(row.cells[DIV_COL], _WHITE_HEX)
        row.cells[DIV_COL].width = Cm(div_w)

    _spacer(doc, height_pt=6)


def _write_race_category_strip(doc: Document,
                               runners: List[RunnerRaceEntry]) -> None:
    """Category leaders — coloured badge strip, male and female champion per category."""
    _section_heading(doc, "Category Leaders", space_before_pt=14)

    # Collect best M + F per category
    m_champs: dict = {}
    f_champs: dict = {}
    for cat in _CATEGORIES:
        m_champs[cat] = next(iter(sorted(
            [r for r in runners if r.eligible and r.normalised_category == cat
             and r.gender == "M" and r.points > 0],
            key=lambda r: (-r.points, r.time_seconds),
        )), None)
        f_champs[cat] = next(iter(sorted(
            [r for r in runners if r.eligible and r.normalised_category == cat
             and r.gender == "F" and r.points > 0],
            key=lambda r: (-r.points, r.time_seconds),
        )), None)

    n_cats    = len(_CATEGORIES)
    label_w   = Cm(1.5)
    cat_w     = Cm(round((18.0 - 1.5) / n_cats, 2))  # fits A4 portrait 18cm usable width

    # 3-row table: badge header | male row | female row
    tbl = doc.add_table(rows=3, cols=n_cats + 1)
    tbl.style = "Table Grid"
    _no_table_borders(tbl)

    # ── Row 0: category badge headers ────────────────────────────────────
    r0 = tbl.rows[0]
    _set_cell_bg(r0.cells[0], _NAVY_HEX)
    r0.cells[0].width = label_w
    for j, cat in enumerate(_CATEGORIES):
        c = r0.cells[j + 1]
        c.width = cat_w
        _set_cell_bg(c, _CAT_BADGE_HEX.get(cat, _NAVY_HEX))
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after  = Pt(5)
        run = p.add_run(cat.upper())
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = _WHITE_RGB

    # ── Rows 1 & 2: male / female champion cells ─────────────────────────
    for row_idx, (label, champs, label_bg) in enumerate(
        [("MALE",   m_champs, _NAVY_HEX),
         ("FEMALE", f_champs, "4a5e6e")],
        start=1,
    ):
        r = tbl.rows[row_idx]

        # Label column
        lc = r.cells[0]
        lc.width = label_w
        _set_cell_bg(lc, label_bg)
        lp = lc.paragraphs[0]
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.paragraph_format.space_before = Pt(8)
        lp.paragraph_format.space_after  = Pt(4)
        lr = lp.add_run(label)
        lr.bold = True
        lr.font.size = Pt(8)
        lr.font.color.rgb = _WHITE_RGB

        # Champion cells
        cell_bg = _WHITE_HEX if row_idx == 1 else _ALT_HEX
        for j, cat in enumerate(_CATEGORIES):
            c = r.cells[j + 1]
            c.width = cat_w
            champ = champs[cat]
            _set_cell_bg(c, cell_bg)

            # Name (bold, navy)
            p1 = c.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.paragraph_format.space_before = Pt(6)
            p1.paragraph_format.space_after  = Pt(1)
            if champ:
                rn = p1.add_run(champ.name)
                rn.bold = True
                rn.font.size = Pt(10)
                rn.font.color.rgb = _NAVY_RGB

            # Club (small italic, muted)
            p2 = c.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after  = Pt(6)
            if champ and champ.preferred_club:
                rc = p2.add_run(champ.preferred_club)
                rc.italic = True
                rc.font.size = Pt(8)
                rc.font.color.rgb = RGBColor(0x70, 0x80, 0x90)


def write_race_report(
    race_num: int,
    total_races: int,
    runners: List[RunnerRaceEntry],
    team_results: List[TeamRaceResult],
    images_dir: Optional[Path],
    year: int,
    filepath: Path,
    source_file: Optional[Path] = None,
) -> None:
    """Write the per-race scoring card DOCX (+ PDF) styled after the WRRL card."""
    race_name = _extract_race_name(source_file) if source_file else f"Race {race_num}"
    filepath  = Path(filepath).with_suffix(".docx")

    doc = Document()
    section = doc.sections[0]
    section.page_width    = Cm(21.0)
    section.page_height   = Cm(29.7)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    # Ensure no landscape attribute is present (default template may carry one)
    _pgSz = section._sectPr.find(qn("w:pgSz"))
    if _pgSz is not None:
        _pgSz.attrib.pop(qn("w:orient"), None)

    _build_race_header(doc, race_name, race_num, total_races, images_dir)
    _spacer(doc, height_pt=8)
    _write_race_individual(doc, runners, top_n=10)
    _write_race_category_strip(doc, runners)
    _page_break(doc)
    _write_race_teams(doc, team_results)
    _build_footer(doc, year)

    filepath.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(filepath))
    log.info("Race %d report written: %s", race_num, filepath.name)

    try:
        from docx2pdf import convert  # type: ignore
        pdf_path = filepath.with_suffix(".pdf")
        convert(str(filepath), str(pdf_path))
        log.info("Race %d PDF written: %s", race_num, pdf_path.name)
    except Exception as exc:
        log.warning("Race %d PDF conversion skipped — %s", race_num, exc)


# ── Document header (branding) ─────────────────────────────────────────────────

def _build_cover_header(doc: Document, year: int, images_dir: Optional[Path]) -> None:
    """Two-cell table: shield logo left | title text right, both on navy."""
    tbl = doc.add_table(rows=1, cols=2)
    logo_cell  = tbl.rows[0].cells[0]
    title_cell = tbl.rows[0].cells[1]

    logo_cell.width  = Cm(3.2)
    title_cell.width = Cm(22.8)

    _set_cell_bg(logo_cell, _NAVY_HEX)
    _set_cell_bg(title_cell, _NAVY_HEX)

    # Logo
    shield_path = (images_dir / "WRRL shield concept.png") if images_dir else None
    if shield_path and shield_path.exists():
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(str(shield_path), width=Cm(2.8))

    # Title
    title_cell.paragraphs[0].clear()
    p1 = title_cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.paragraph_format.left_indent  = Cm(0.4)
    p1.paragraph_format.space_before = Pt(8)
    r = p1.add_run(f"WRRL League Report  |  Season {year}")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = _WHITE_RGB

    p2 = title_cell.add_paragraph()
    p2.paragraph_format.left_indent = Cm(0.4)
    r2 = p2.add_run("Wiltshire Road and Running League  —  Season Summary")
    r2.font.size = Pt(10)
    r2.font.color.rgb = _SUBHDR_RGB

    # Green accent rule below header
    accent = doc.add_paragraph()
    accent.paragraph_format.space_before = Pt(0)
    accent.paragraph_format.space_after  = Pt(6)
    _set_para_bg(accent, _GREEN_HEX)


# ── Footer ─────────────────────────────────────────────────────────────────────

def _build_footer(doc: Document, year: int) -> None:
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        f"© {year} Wiltshire Road and Running League  |  Generated automatically"
    )
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x90)


# ── Public entry point ─────────────────────────────────────────────────────────

def write_combined_report(
    highest_race: int,
    year: int,
    male_records: List[RunnerSeasonRecord],
    female_records: List[RunnerSeasonRecord],
    div1_teams: List[TeamSeasonRecord],
    div2_teams: List[TeamSeasonRecord],
    unrec_all: List[UnrecognisedClub],
    images_dir: Optional[Path],
    filepath: Path,
) -> None:
    """Write the branded combined report to *filepath* (.docx) and attempt PDF.

    Parameters
    ----------
    highest_race:  highest race number included in this run.
    year:          season year (for title and footer).
    male_records / female_records: season leaderboard records.
    div1_teams / div2_teams: club table records per division.
    unrec_all:     aggregate non-league club list for overview.
    images_dir:    folder containing ``WRRL shield concept.png``.
    filepath:      output path; .docx extension used automatically.
    """
    filepath = Path(filepath).with_suffix(".docx")

    doc = Document()

    # ── A4 Landscape ──────────────────────────────────────────────────────────
    section = doc.sections[0]
    section.orientation   = WD_ORIENT.LANDSCAPE
    section.page_width    = Cm(29.7)
    section.page_height   = Cm(21.0)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    nonleague_fname = ""  # non-league DOCX no longer generated

    # ── Content ───────────────────────────────────────────────────────────────
    _build_cover_header(doc, year, images_dir)
    _write_overview(doc, highest_race, year, male_records, female_records,
                    unrec_all)

    # Club tables
    _spacer(doc, height_pt=10)  # breathing room before first club section bar
    _write_club_table(doc, div1_teams, "Division 1 — Club Table",
                      highest_race, is_div1=True)
    _page_break(doc)
    _write_club_table(doc, div2_teams, "Division 2 — Club Table",
                      highest_race, is_div1=False)

    # Individual scorer section — new page
    _page_break(doc)
    _section_heading(doc, "Individual Scorers", space_before_pt=4)
    _write_individual_table(doc, male_records,   "Top 20 — Male",   highest_race, top_n=20)
    doc.add_page_break()
    _write_individual_table(doc, female_records, "Top 20 — Female", highest_race, top_n=20)

    # Category leaders — new page
    doc.add_page_break()
    _write_category_leaders(doc, male_records, female_records, top_n=3)

    _build_footer(doc, year)

    # ── Save DOCX ─────────────────────────────────────────────────────────────
    filepath.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(filepath))
    log.info("Report written: %s", filepath.name)

    # ── PDF conversion (requires Microsoft Word on Windows) ───────────────────
    try:
        from docx2pdf import convert  # type: ignore
        pdf_path = filepath.with_suffix(".pdf")
        convert(str(filepath), str(pdf_path))
        log.info("PDF written: %s", pdf_path.name)
    except Exception as exc:
        log.warning("PDF conversion skipped — %s", exc)
