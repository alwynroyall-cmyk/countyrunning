from __future__ import annotations

from datetime import datetime, timezone
import json
import re
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import urlopen

import pandas as pd


class SporthiveRaceNotDirectlyImportableError(ValueError):
    """Raised when Sporthive race id cannot be resolved through the public API."""


def _fetch_json(url: str) -> dict:
    with urlopen(url, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def _extract_codes_from_url(url: str) -> tuple[str, int]:
    match = re.search(r"/events/([^/]+)/race/(\d+)", url)
    if not match:
        raise ValueError(
            "Could not parse event code and race id from URL. "
            "Expected format .../events/<eventCode>/race/<raceId>."
        )
    return match.group(1), int(match.group(2))


def _extract_sporthive_ids(url: str) -> tuple[int | None, int | None]:
    race_match = re.search(r"sporthive\.com/events/s/(\d+)/race/(\d+)", url)
    if race_match:
        return int(race_match.group(1)), int(race_match.group(2))

    event_match = re.search(r"sporthive\.com/events/s/(\d+)", url)
    if event_match:
        return int(event_match.group(1)), None

    return None, None


def _resolve_event_and_race(event_code: str, race_id: int) -> tuple[int, int, str]:
    event_payload = _fetch_json(f"https://results.raceroster.com/v2/api/events/{event_code}")
    event_data = event_payload.get("data", {}).get("event", {})
    result_event_id = int(event_data.get("resultEventId"))

    race_payload = _fetch_json(f"https://results.raceroster.com/v2/api/events/{event_code}/sub-events/{race_id}")
    race_data = race_payload.get("data", {})
    result_sub_event_id = int(race_data.get("resultSubEventId") or race_data.get("id"))
    race_name = str(race_data.get("name") or f"Race {race_id}")

    return result_event_id, result_sub_event_id, race_name


def _fetch_all_rows(result_event_id: int, sub_event_id: int) -> list[dict]:
    start = 0
    limit = 50
    all_rows: list[dict] = []

    while True:
        query = urlencode({"start": start, "limit": limit})
        url = (
            f"https://results.raceroster.com/v2/api/result-events/"
            f"{result_event_id}/sub-events/{sub_event_id}/results?{query}"
        )
        payload = _fetch_json(url)
        rows = payload.get("data", [])
        if not rows:
            break

        all_rows.extend(rows)
        if len(rows) < limit:
            break

        start += len(rows)

    return all_rows


def _fetch_sporthive_session(session_id: int) -> dict:
    url = f"https://eventresults-api.speedhive.com/api/v0.2.3/eventresults/sessions/{session_id}"
    return _fetch_json(url)


def _fetch_sporthive_classification_rows(session_id: int) -> list[dict]:
    all_rows: list[dict] = []
    offset = 0
    count = 200

    while True:
        query = urlencode({"count": count, "offset": offset})
        url = (
            f"https://eventresults-api.speedhive.com/api/v0.2.3/eventresults/"
            f"sessions/{session_id}/classification?{query}"
        )
        payload = _fetch_json(url)
        rows = payload.get("rows", [])
        if not rows:
            break
        all_rows.extend(rows)
        if len(rows) < count:
            break
        offset += len(rows)

    return all_rows


def _to_wrrl_rows(rows: list[dict]) -> list[dict]:
    out: list[dict] = []

    for row in rows:
        chip = str(row.get("chipTime") or "").strip()
        gun = str(row.get("gunTime") or "").strip()
        gender = str(row.get("genderSexId") or "").strip().lower()

        out.append(
            {
                "Pos": row.get("overallPlace", ""),
                "Bib": row.get("bib", ""),
                "Name": row.get("name", ""),
                "Club": row.get("custom23101578", ""),
                "Gender": "F" if gender.startswith("w") else "M",
                "Category": row.get("custom23096209", ""),
                "Time": chip or gun,
                "Chip Time": chip,
                "Gun Time": gun,
                "Delta": row.get("custom23096212", ""),
                "Category Pos": row.get("custom23096210", ""),
                "Source ID": row.get("id", ""),
            }
        )

    return out


def _guess_gender_from_category(category: str) -> str:
    cat = str(category or "").strip().upper()
    if cat.startswith(("F", "W", "L")):
        return "F"
    if cat:
        return "M"
    return ""


def _extract_time_text(row: dict) -> str:
    direct_candidates = [
        row.get("time"),
        row.get("totalTime"),
        row.get("netTime"),
        row.get("gunTime"),
    ]
    for value in direct_candidates:
        text = str(value or "").strip()
        if text:
            return text

    for value in row.get("additionalFields", []) or []:
        text = str(value or "").strip()
        if re.fullmatch(r"\d{1,2}:\d{2}:\d{2}", text):
            return text
    return ""


def _to_wrrl_rows_sporthive(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        additional = row.get("additionalFields", []) or []
        club_guess = str(additional[0]).strip() if additional else ""
        category = str(row.get("resultClass") or "").strip()
        out.append(
            {
                "Pos": row.get("position", ""),
                "Bib": row.get("startNumber", ""),
                "Name": row.get("name", ""),
                "Club": club_guess,
                "Gender": _guess_gender_from_category(category),
                "Category": category,
                "Time": _extract_time_text(row),
                "Chip Time": "",
                "Gun Time": "",
                "Delta": "",
                "Category Pos": row.get("positionInClass", ""),
                "Source ID": row.get("id", ""),
            }
        )
    return out


def _build_sporthive_header_index(headers: list[str]) -> dict[str, int]:
    def _key(value: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())

    mapping: dict[str, int] = {}
    for idx, header in enumerate(headers):
        key = _key(header)
        if key:
            mapping[key] = idx
    return mapping


def _row_value(cells: list[str], header_index: dict[str, int], *aliases: str) -> str:
    for alias in aliases:
        idx = header_index.get(alias)
        if idx is None:
            continue
        if 0 <= idx < len(cells):
            return str(cells[idx] or "").strip()
    return ""


def _to_wrrl_rows_sporthive_rendered(headers: list[str], raw_rows: list[list[str]]) -> list[dict]:
    header_index = _build_sporthive_header_index(headers)
    out: list[dict] = []

    for cells in raw_rows:
        stripped = [str(cell or "").strip() for cell in cells]
        if not any(stripped):
            continue

        pos = _row_value(stripped, header_index, "POS", "POSITION")
        name = _row_value(stripped, header_index, "NAME", "PARTICIPANT")
        bib = _row_value(stripped, header_index, "BIB", "BIBNO", "BIBNUMBER", "STARTNUMBER")
        club = _row_value(stripped, header_index, "TEAMCLUB", "TEAM", "CLUB")
        category = _row_value(stripped, header_index, "CATEGORY", "CLASS", "RESULTCLASS")
        gun_time = _row_value(stripped, header_index, "GUNTIME")
        chip_time = _row_value(stripped, header_index, "CHIPTIME", "NETTIME")
        cat_pos = _row_value(stripped, header_index, "CATEGORYPOS", "CLASSPOS", "POSITIONINCLASS")

        if not name:
            continue

        out.append(
            {
                "Pos": pos,
                "Bib": bib,
                "Name": name,
                "Club": club,
                "Gender": _guess_gender_from_category(category),
                "Category": category,
                "Time": chip_time or gun_time,
                "Chip Time": chip_time,
                "Gun Time": gun_time,
                "Delta": "",
                "Category Pos": cat_pos,
                "Source ID": "",
            }
        )

    return out


def _fetch_sporthive_rendered_rows(race_url: str) -> tuple[str, list[dict]]:
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        raise RuntimeError(
            "Automatic Sporthive scraping requires Playwright (Python package + browser runtime)."
        ) from exc

    all_rows: list[dict] = []
    race_name = "Sporthive Race"

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto(race_url, wait_until="networkidle", timeout=120000)
            page.wait_for_selector(".sporthive-data-table__tbody .sporthive-data-table__tr", timeout=120000)

            title_text = page.locator(".sporthive-race-page__race-name")
            if title_text.count() > 0:
                race_name = title_text.first.inner_text().strip() or race_name

            toggle = page.locator(".sporthive-race-page__radio-button-container", has_text="Individual")
            if toggle.count() > 0:
                toggle.first.click(timeout=3000)
                page.wait_for_timeout(700)

            # Track already-collected pages by the first row position, not by
            # the paginator active-button number.  Sporthive sometimes omits the
            # active-button highlight (e.g. page 28 of 29), so relying on the
            # active indicator causes spurious early loop breaks.
            visited_first_positions: set[str] = set()
            max_pages = 500

            for _ in range(max_pages):
                snapshot = page.evaluate(
                    r"""
                    () => {
                        const headers = Array.from(
                          document.querySelectorAll('.sporthive-data-table__thead .sporthive-data-table__th')
                        ).map(el => (el.innerText || '').trim());
                        const rows = Array.from(
                          document.querySelectorAll('.sporthive-data-table__tbody .sporthive-data-table__tr')
                        ).map(tr => Array.from(tr.querySelectorAll('.sporthive-data-table__td')).map(td => (td.innerText || '').trim()));
                        const firstRowPos = rows.length > 0 && rows[0].length > 0 ? rows[0][0] : '';
                        const active = document.querySelector('.sporthive-paginator__button--active')?.textContent?.trim() || '';
                        const numberedBtns = Array.from(document.querySelectorAll('button.sporthive-paginator__button'))
                          .map(b => (b.textContent || '').trim())
                          .filter(t => /^\d+$/.test(t));
                        const nextBtn = Array.from(document.querySelectorAll('button.sporthive-paginator__button'))
                          .find(btn => (btn.textContent || '').trim() === '>');
                        return {
                          headers,
                          rows,
                          firstRowPos,
                          active,
                          numberedBtns,
                          hasNextBtn: !!nextBtn && !nextBtn.disabled
                        };
                    }
                    """
                )

                first_row_pos = str(snapshot.get("firstRowPos") or "")
                # Table still loading — wait and retry
                if not first_row_pos:
                    page.wait_for_timeout(400)
                    continue

                if first_row_pos in visited_first_positions:
                    break
                visited_first_positions.add(first_row_pos)

                headers = [str(value or "").strip() for value in snapshot.get("headers", [])]
                raw_rows = snapshot.get("rows", [])
                all_rows.extend(_to_wrrl_rows_sporthive_rendered(headers, raw_rows))

                # --- Determine the next page to navigate to ---
                active_page = str(snapshot.get("active") or "")
                numbered_btns: list[str] = [str(b) for b in snapshot.get("numberedBtns", [])]
                has_next_btn: bool = bool(snapshot.get("hasNextBtn"))

                # Try to click the specific next numbered button visible in the window
                next_page_number: str | None = None
                if active_page.isdigit():
                    candidate = str(int(active_page) + 1)
                    if candidate in numbered_btns:
                        next_page_number = candidate

                if next_page_number is not None:
                    page.evaluate(
                        r"""
                        nextPage => {
                            const btn = Array.from(document.querySelectorAll('button.sporthive-paginator__button'))
                              .find(b => (b.textContent || '').trim() === nextPage);
                            if (btn && !btn.disabled) btn.click();
                        }
                        """,
                        arg=next_page_number,
                    )
                elif has_next_btn:
                    moved = page.evaluate(
                        r"""
                        () => {
                            const btn = Array.from(document.querySelectorAll('button.sporthive-paginator__button'))
                              .find(b => (b.textContent || '').trim() === '>');
                            if (!btn || btn.disabled) return false;
                            btn.click();
                            return true;
                        }
                        """
                    )
                    if not moved:
                        break
                else:
                    break  # No navigation possible — last page reached

                # Wait for first visible row position to change.  On timeout, verify
                # whether navigation actually occurred; if not, we are on the last page.
                try:
                    page.wait_for_function(
                        r"""
                        prev => {
                            const td = document.querySelector(
                              '.sporthive-data-table__tbody .sporthive-data-table__tr .sporthive-data-table__td'
                            );
                            if (!td) return false;
                            const pos = (td.innerText || '').trim();
                            return pos !== '' && pos !== prev;
                        }
                        """,
                        arg=first_row_pos,
                        timeout=15000,
                    )
                except PlaywrightTimeoutError:
                    # Check whether the page was actually navigated
                    current_pos = page.evaluate(
                        r"() => (document.querySelector('.sporthive-data-table__tbody .sporthive-data-table__tr .sporthive-data-table__td')?.innerText || '').trim()"
                    )
                    if not current_pos or current_pos == first_row_pos:
                        break  # Navigation had no effect — last page
                    # Slow response but content changed — continue normally
        finally:
            browser.close()

    if not all_rows:
        raise RuntimeError("No result rows were found on the rendered Sporthive race page.")

    deduped: dict[tuple[str, str, str], dict] = {}
    for row in all_rows:
        key = (
            str(row.get("Pos", "")).strip(),
            str(row.get("Bib", "")).strip(),
            str(row.get("Name", "")).strip().lower(),
        )
        deduped[key] = row

    def _sort_key(item: dict) -> tuple[int, str, str]:
        pos_text = str(item.get("Pos", "")).strip()
        pos_num = int(pos_text) if pos_text.isdigit() else 999999
        return (pos_num, str(item.get("Bib", "")), str(item.get("Name", "")))

    rows = sorted(deduped.values(), key=_sort_key)
    return race_name, rows


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 _-]", "", str(value)).strip()
    return cleaned or "Race"


def _history_path(input_dir: Path) -> Path:
    return input_dir / "Race Roster Import History.csv"


def _append_history_row(
    input_dir: Path,
    race_url: str,
    event_code: str,
    race_id: int,
    result_event_id: int,
    result_sub_event_id: int,
    league_race_number: int,
    race_name: str,
    row_count: int,
    output_path: Path,
) -> Path:
    history_file = _history_path(input_dir)
    row = {
        "Imported UTC": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "Race URL": race_url,
        "Event Code": event_code,
        "Source Race ID": race_id,
        "Result Event ID": result_event_id,
        "Result Sub-Event ID": result_sub_event_id,
        "League Race Number": league_race_number,
        "Race Name": race_name,
        "Rows Imported": row_count,
        "Output File": str(output_path),
    }

    frame = pd.DataFrame([row])
    if history_file.exists():
        frame.to_csv(history_file, mode="a", header=False, index=False)
    else:
        frame.to_csv(history_file, index=False)
    return history_file


def _parse_sporthive_table_page(page_text: str) -> list[dict]:
    rows: list[dict] = []
    for line in str(page_text).splitlines():
        if "|" not in line:
            continue
        cells = [cell.strip() for cell in line.split("|")]
        cells = [cell for cell in cells if cell != ""]
        if len(cells) < 7:
            continue
        if not re.fullmatch(r"\d+", cells[0]):
            continue

        pos = cells[0]
        name = cells[1]
        bib = cells[2]
        club = cells[3]
        category = cells[4]
        gun_time = cells[5]
        chip_time = cells[6]

        if not name:
            continue

        rows.append(
            {
                "Pos": pos,
                "Bib": bib,
                "Name": name,
                "Club": club,
                "Gender": _guess_gender_from_category(category) or "M",
                "Category": category,
                "Time": chip_time or gun_time,
                "Chip Time": chip_time,
                "Gun Time": gun_time,
                "Delta": "",
                "Category Pos": "",
                "Source ID": "",
            }
        )
    return rows


def import_sporthive_manual_pages(
    race_url: str,
    pages_text: list[str],
    input_dir: Path,
    league_race_number: int,
    race_name_override: str | None = None,
) -> tuple[Path, int, Path]:
    parsed_rows: list[dict] = []
    for page in pages_text:
        parsed_rows.extend(_parse_sporthive_table_page(page))

    if not parsed_rows:
        raise ValueError("No runner rows were detected in pasted Sporthive pages.")

    deduped: dict[tuple[str, str, str], dict] = {}
    for row in parsed_rows:
        key = (str(row.get("Pos", "")), str(row.get("Bib", "")), str(row.get("Name", "")).lower())
        deduped[key] = row

    rows = [deduped[key] for key in sorted(deduped, key=lambda item: int(item[0]) if item[0].isdigit() else 999999)]

    event_id, race_id = _extract_sporthive_ids(race_url)
    race_name = race_name_override.strip() if race_name_override else (f"Sporthive Race {race_id or ''}".strip())

    file_stem = f"Race #{league_race_number} - {_safe_name(race_name)}"
    output_path = input_dir / f"{file_stem}.xlsx"
    pd.DataFrame(rows).to_excel(output_path, index=False)

    history_file = _append_history_row(
        input_dir=input_dir,
        race_url=race_url,
        event_code=str(event_id or ""),
        race_id=int(race_id or 0),
        result_event_id=0,
        result_sub_event_id=0,
        league_race_number=league_race_number,
        race_name=race_name,
        row_count=len(rows),
        output_path=output_path,
    )

    return output_path, len(rows), history_file


def import_raceroster_results(
    race_url: str,
    input_dir: Path,
    league_race_number: int,
    race_name_override: str | None = None,
    sporthive_race_id_hint: int | None = None,
) -> tuple[Path, int, Path]:
    lower_url = race_url.lower()

    provider_label = "Race Roster"
    event_code = ""
    result_event_id = 0
    result_sub_event_id = 0

    if "raceroster.com" in lower_url:
        event_code, race_id = _extract_codes_from_url(race_url)
        result_event_id, result_sub_event_id, race_name = _resolve_event_and_race(event_code, race_id)
        rows = _fetch_all_rows(result_event_id, result_sub_event_id)
        if not rows:
            raise ValueError("No results were returned by Race Roster.")
        wrll_rows = _to_wrrl_rows(rows)
    elif "sporthive.com" in lower_url:
        provider_label = "Sporthive"
        event_id, race_id = _extract_sporthive_ids(race_url)
        if race_id is None:
            race_id = sporthive_race_id_hint
        if race_id is None:
            raise ValueError(
                "Sporthive event-summary links do not include the race/session id. "
                "Open 'View results' on Sporthive and use that URL (contains '/race/<id>'), "
                "or enter the race id when prompted."
            )

        try:
            session = _fetch_sporthive_session(race_id)
            race_name = str(session.get("name") or f"Sporthive Session {race_id}")
            rows = _fetch_sporthive_classification_rows(race_id)
            if not rows:
                raise ValueError("No classification rows were returned by Sporthive.")

            result_event_id = int(session.get("eventId") or 0)
            result_sub_event_id = int(session.get("id") or race_id)
            event_code = str(event_id or "")
            wrll_rows = _to_wrrl_rows_sporthive(rows)
        except Exception as exc:
            if "HTTP Error 404" not in str(exc):
                raise

            try:
                rendered_race_name, rendered_rows = _fetch_sporthive_rendered_rows(race_url)
            except Exception as rendered_exc:
                raise SporthiveRaceNotDirectlyImportableError(
                    "Sporthive returned 404 for direct API import, and automatic page scraping failed. "
                    "Use manual page-paste mode for this race.\n\n"
                    f"Auto-scrape error: {rendered_exc}"
                ) from rendered_exc

            race_name = rendered_race_name or f"Sporthive Session {race_id}"
            rows = rendered_rows
            result_event_id = int(event_id or 0)
            result_sub_event_id = int(race_id)
            event_code = str(event_id or "")
            wrll_rows = rendered_rows
    else:
        raise ValueError(
            "Unsupported URL. Use a Race Roster link or a Sporthive race link containing '/race/<id>'."
        )

    name_for_file = race_name_override.strip() if race_name_override else race_name
    file_stem = f"Race #{league_race_number} - {_safe_name(name_for_file)}"
    output_path = input_dir / f"{file_stem}.xlsx"

    dataframe = pd.DataFrame(wrll_rows)
    dataframe.to_excel(output_path, index=False)

    history_file = _append_history_row(
        input_dir=input_dir,
        race_url=race_url,
        event_code=event_code,
        race_id=race_id,
        result_event_id=result_event_id,
        result_sub_event_id=result_sub_event_id,
        league_race_number=league_race_number,
        race_name=name_for_file,
        row_count=len(rows),
        output_path=output_path,
    )

    return output_path, len(rows), history_file
