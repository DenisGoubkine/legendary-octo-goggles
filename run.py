#!/usr/bin/env python3
"""Market Data Automation — single-script framework."""

from __future__ import annotations

import argparse
import base64
import json
import logging
import sys
import tempfile
import time
from datetime import date, datetime
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path
from typing import Any

import win32com.client as win32
from PIL import Image

SCRIPT_DIR = Path(__file__).resolve().parent
LOG = logging.getLogger("carlopdf")


def _outlook_inbox(namespace: Any) -> Any:
    return namespace.GetDefaultFolder(6)  # olFolderInbox


def image_data_uri(path: Path) -> str:
    """Embed a local image in HTML (no Outlook MAPI properties needed)."""
    mime = {".png": "png", ".jpg": "jpeg", ".jpeg": "jpeg", ".gif": "gif"}.get(
        path.suffix.lower(), "png"
    )
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/{mime};base64,{encoded}"


def decode_data_uri(src: str) -> bytes | None:
    if not src.lower().startswith("data:image"):
        return None
    try:
        _header, encoded = src.split(",", 1)
        return base64.b64decode(encoded)
    except (ValueError, TypeError):
        return None


# =============================================================================
# CONFIG — edit settings here (the only place you need to change)
# Optional: pass --config path.json to override with an external JSON file.
# =============================================================================

CONFIG: dict[str, Any] = {

    # --- EMAIL ---
    "email_source": {
        "folder": "Inbox/QES",                    # Outlook folder, e.g. "Inbox/QES"
        "subject": "Daily Market Summary",        # exact daily email subject line
        "optional": False,
        "received_after": "16:20",                # HH:MM — start of today's window
        "received_before": "16:45",               # HH:MM — end of today's window
        "poll_interval_seconds": 60,
        "poll_duration_minutes": 15,
        "chart": {
            "title_patterns": ["S&P 500", "S&P500"],
            "exclude_patterns": ["logo", "header", "banner", "unsubscribe"],
            "min_width": 200,
            "min_height": 150,
            "max_banner_height": 200,
            "require_reddish": False,
        },
    },

    # --- EXCEL ---
    "excel": {
        "workbook": r"C:\path\to\data.xlsm",
        "bloomberg_addin": r"C:\blp\API\Office Tools\BloombergUI.xla",
        "refresh_mode": "workbook",               # "workbook", "worksheets", or "none"
        "timeout_seconds": 300,
        "poll_interval_seconds": 2,
        "read_only": True,
        "close_workbook": True,
        "quit_excel": False,
        "optional": False,
        "tabs": [
            {
                "name": "Carlo Sectors",
                "wait_cells": ["F4"],
                "sections": [
                    {
                        "label": "Carlo Sectors",
                        "range": "F4:T77",
                        "skip_empty_rows": True,
                    },
                ],
            },
            {
                "name": "Carlo Desk Flows",
                "wait_cells": ["A2", "G2"],
                "sections": [
                    {
                        "label": "Desk Flows 5 Day",
                        "header_range": "A1:D1",
                        "range": "A2:D15",
                    },
                    {
                        "label": "Desk Flows 1 Day",
                        "header_range": "G1:J1",
                        "range": "G2:J15",
                    },
                ],
            },
        ],
    },

    # --- BLOOMBERG / SPTSX ---
    "bloomberg_direct": {
        "optional": False,
        "index_ticker": "SPTSX Index",
        "index_fields": [
            "PX_LAST",
            "CHG_NET_1D",
            "CHG_PCT_1D",
            "VOLUME",
            "VOLUME_AVG_20D",
        ],
        "sptsx": {
            "enabled": True,
            "index_summary": True,
            "mov_by_subindustry": True,
            "top_bottom_chart": True,
            "top_n": 10,
            "movers_per_side": 3,
            "member_fields": {
                "name": "SHORT_NAME",
                "price": "PX_LAST",
                "change_pct": "CHG_PCT_1D",
                "points": "INDX_POINTS",
                "pct_ind_mkt_val": "PCT_IND_MKT_VAL",
                "subindustry_code": "GICS_SUB_INDUSTRY",
                "subindustry_name": "GICS_SUB_INDUSTRY_NAME",
            },
        },
        "tables": [],                             # optional extra xbbg bdp/bdh/bds tables
    },

    # --- OUTPUT ---
    "output": {
        "draft_to": "",                           # recipient email address(es)
        "draft_cc": "",
        "draft_subject": "Daily Market Update — {date}",
        "intro_text": "Auto-generated draft. Review before sending.",
    },

    "state_file": ".state.json",                  # processed email IDs (do not edit)
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_config(path: Path) -> dict:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def load_state(path: Path) -> dict:
    if path.exists():
        with path.open(encoding="utf-8") as f:
            return json.load(f)
    return {"processed_entry_ids": []}


def save_state(path: Path, state: dict) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def parse_hhmm(value: str) -> tuple[int, int]:
    hour, minute = value.split(":")
    return int(hour), int(minute)


def today_window(after: str, before: str) -> tuple[datetime, datetime]:
    today = date.today()
    ah, am = parse_hhmm(after)
    bh, bm = parse_hhmm(before)
    start = datetime(today.year, today.month, today.day, ah, am)
    end = datetime(today.year, today.month, today.day, bh, bm)
    return start, end


def py_datetime(value: Any) -> datetime:
    return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)


def fmt_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        return f"{value:.4g}"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def fmt_pct(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        return f"{float(value):.2f}%"
    except (TypeError, ValueError):
        return str(value)


def normalize_range_value(value: Any) -> list[list[Any]]:
    if value is None:
        return [[None]]
    if isinstance(value, tuple):
        rows: list[list[Any]] = []
        for row in value:
            if isinstance(row, tuple):
                rows.append(list(row))
            else:
                rows.append([row])
        return rows
    return [[value]]


def row_is_empty(row: list[Any]) -> bool:
    return all(v is None or str(v).strip() == "" for v in row)


def filter_empty_rows(grid: list[list[Any]]) -> list[list[Any]]:
    return [row for row in grid if not row_is_empty(row)]


def is_bloomberg_pending(value: Any) -> bool:
    if isinstance(value, str):
        return (
            "#N/A Requesting Data" in value
            or "#N/A Requesting Data..." in value
            or "#N/A Real Time Data" in value
        )
    if isinstance(value, tuple):
        for row in value:
            if isinstance(row, tuple):
                for cell in row:
                    if is_bloomberg_pending(cell):
                        return True
            elif is_bloomberg_pending(row):
                return True
    return False


def df_to_table(df: Any) -> dict:
    if df is None:
        return {"columns": [], "rows": []}
    try:
        empty = df.empty
    except AttributeError:
        empty = len(df) == 0
    if empty:
        return {"columns": [], "rows": []}
    df = df.reset_index()
    columns = [str(c) for c in df.columns]
    rows = [[fmt_cell(v) for v in row] for row in df.values.tolist()]
    return {"columns": columns, "rows": rows}


def parse_folder_path(folder_spec: str | list[str]) -> list[str]:
    if isinstance(folder_spec, list):
        return [str(p).strip() for p in folder_spec if str(p).strip()]
    parts = [p.strip() for p in str(folder_spec).replace("\\", "/").split("/") if p.strip()]
    return parts or ["Inbox"]


# ---------------------------------------------------------------------------
# Component 1: Email chart extraction
# ---------------------------------------------------------------------------


class ImgTagParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tags: list[dict[str, str]] = []
        self._text_chunks: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "img":
            return
        attr_map = {k.lower(): (v or "") for k, v in attrs}
        self.tags.append(
            {
                "src": attr_map.get("src", ""),
                "alt": attr_map.get("alt", ""),
                "width": attr_map.get("width", ""),
                "height": attr_map.get("height", ""),
                "context": " ".join(self._text_chunks[-3:]),
            }
        )

    def handle_data(self, data: str) -> None:
        text = data.strip()
        if text:
            self._text_chunks.append(text)


def get_outlook_folder(namespace: Any, folder_spec: str | list[str]) -> Any:
    parts = parse_folder_path(folder_spec)
    if not parts:
        return _outlook_inbox(namespace)

    if parts[0].lower() in ("inbox", "6"):
        folder = _outlook_inbox(namespace)
        remaining = parts[1:]
    else:
        folder = None
        remaining = parts
        for store in namespace.Folders:
            try:
                if str(store.Name).lower() == parts[0].lower():
                    folder = store
                    remaining = parts[1:]
                    break
            except Exception:
                continue
        if folder is None:
            folder = namespace.Folders.Item(1)
            remaining = parts

    for part in remaining:
        matched = None
        for sub in folder.Folders:
            if str(sub.Name).lower() == part.lower():
                matched = sub
                break
        if matched is None:
            raise RuntimeError(f"Outlook folder not found: {part} (path: {' / '.join(parts)})")
        folder = matched

    return folder


def attachment_cid_map(mail_item: Any) -> dict[str, bytes]:
    """Map attachment lookup keys (filename, index) to image bytes."""
    mapping: dict[str, bytes] = {}
    for i in range(1, mail_item.Attachments.Count + 1):
        attachment = mail_item.Attachments.Item(i)
        fname = str(attachment.FileName or f"attachment{i}")
        tmp = Path(tempfile.gettempdir()) / f"carlopdf_att_{i}_{fname}"
        attachment.SaveAsFile(str(tmp))
        data = tmp.read_bytes()
        tmp.unlink(missing_ok=True)
        stem = fname.rsplit(".", 1)[0].lower()
        for key in (fname.lower(), stem, f"att{i}", str(i)):
            mapping[key] = data
    return mapping


def image_dimensions(image_bytes: bytes) -> tuple[int, int]:
    with Image.open(BytesIO(image_bytes)) as img:
        return img.size


def reddish_ratio(image_bytes: bytes) -> float:
    with Image.open(BytesIO(image_bytes)) as img:
        rgb = img.convert("RGB")
        pixels = list(rgb.getdata())
    if not pixels:
        return 0.0
    redish = sum(1 for r, g, b in pixels if r > g + 30 and r > b + 30)
    return redish / len(pixels)


def score_chart_candidate(tag: dict[str, str], image_bytes: bytes, chart_cfg: dict) -> float:
    score = 0.0
    haystack = " ".join([tag.get("alt", ""), tag.get("context", ""), tag.get("src", "")]).lower()

    for pattern in chart_cfg.get("title_patterns", []):
        if pattern.lower() in haystack:
            score += 3.0

    for pattern in chart_cfg.get("exclude_patterns", []):
        if pattern.lower() in haystack:
            score -= 5.0

    width = height = 0
    if tag.get("width", "").isdigit():
        width = int(tag["width"])
    if tag.get("height", "").isdigit():
        height = int(tag["height"])
    if width == 0 or height == 0:
        try:
            width, height = image_dimensions(image_bytes)
        except Exception:
            width = height = 0

    min_w = chart_cfg.get("min_width", 200)
    min_h = chart_cfg.get("min_height", 150)
    max_banner_h = chart_cfg.get("max_banner_height", 200)

    if width >= min_w and height >= min_h:
        score += 1.0
    if width > 900 and height < max_banner_h:
        score -= 4.0
    if chart_cfg.get("require_reddish") and reddish_ratio(image_bytes) < 0.03:
        score -= 2.0
    elif reddish_ratio(image_bytes) >= 0.03:
        score += 0.5

    return score


def extract_chart_from_mail(mail_item: Any, chart_cfg: dict) -> bytes | None:
    html = mail_item.HTMLBody or ""
    parser = ImgTagParser()
    parser.feed(html)
    cid_map = attachment_cid_map(mail_item)

    best_score = float("-inf")
    best_bytes: bytes | None = None

    for tag in parser.tags:
        src = tag.get("src", "")
        embedded = decode_data_uri(src)
        if embedded:
            score = score_chart_candidate(tag, embedded, chart_cfg)
            if score > best_score:
                best_score = score
                best_bytes = embedded
            continue
        if not src.lower().startswith("cid:"):
            continue
        cid = src[4:].strip("<>").lower()
        cid_keys = [cid, cid.split("@")[0]]
        image_bytes = None
        for key in cid_keys:
            image_bytes = cid_map.get(key)
            if image_bytes:
                break
        if not image_bytes:
            continue
        score = score_chart_candidate(tag, image_bytes, chart_cfg)
        if score > best_score:
            best_score = score
            best_bytes = image_bytes

    if best_bytes is None and cid_map:
        for cid, image_bytes in cid_map.items():
            tag = {"src": cid, "alt": cid, "context": html[:500], "width": "", "height": ""}
            score = score_chart_candidate(tag, image_bytes, chart_cfg)
            if score > best_score:
                best_score = score
                best_bytes = image_bytes

    return best_bytes


def find_matching_email(cfg: dict, state: dict) -> Any | None:
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = get_outlook_folder(namespace, cfg.get("folder", "Inbox"))
    start, end = today_window(cfg["received_after"], cfg["received_before"])
    processed = set(state.get("processed_entry_ids", []))

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    for item in items:
        try:
            subject = str(item.Subject)
            received = py_datetime(item.ReceivedTime)
            entry_id = str(item.EntryID)
        except Exception:
            continue

        if subject != cfg["subject"]:
            continue
        if received < start or received > end:
            continue
        if entry_id in processed:
            continue
        return item

    return None


def fetch_chart(cfg: dict, state_path: Path, poll: bool = True) -> Path:
    chart_cfg = cfg.get("chart", {})
    state = load_state(state_path)
    interval = cfg.get("poll_interval_seconds", 60)
    duration_min = cfg.get("poll_duration_minutes", 15)
    deadline = time.time() + duration_min * 60

    while True:
        mail_item = find_matching_email(cfg, state)
        if mail_item:
            image_bytes = extract_chart_from_mail(mail_item, chart_cfg)
            if not image_bytes:
                raise RuntimeError("Email found but no matching chart image was extracted")

            out_dir = Path(tempfile.gettempdir()) / "carlopdf"
            out_dir.mkdir(exist_ok=True)
            out_path = out_dir / f"chart_{date.today().isoformat()}.png"
            out_path.write_bytes(image_bytes)

            state.setdefault("processed_entry_ids", []).append(str(mail_item.EntryID))
            save_state(state_path, state)
            LOG.info("Chart saved to %s", out_path)
            return out_path

        if not poll or time.time() >= deadline:
            raise RuntimeError(
                f"No email found with subject '{cfg['subject']}' in today's time window"
            )

        LOG.info("Email not found yet; retrying in %ss", interval)
        time.sleep(interval)


# ---------------------------------------------------------------------------
# Component 2: Excel + Bloomberg refresh
# ---------------------------------------------------------------------------


def get_excel_app(addin_path: Path | None) -> Any:
    try:
        return win32.GetActiveObject("Excel.Application")
    except Exception:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = True
        xl.DisplayAlerts = False
        if addin_path and addin_path.exists():
            xl.Workbooks.Open(str(addin_path), ReadOnly=True)
        return xl


def bloomberg_refresh(xl: Any, wb: Any, refresh_mode: str, tabs: list[dict]) -> None:
    if refresh_mode == "none":
        return
    if refresh_mode == "worksheets":
        for tab in tabs:
            wb.Worksheets(tab["name"]).Activate()
            xl.Application.Run("BloombergUI.xla!RefreshEntireWorksheet")
        return
    xl.Application.Run("BloombergUI.xla!RefreshEntireWorkbook")


def collect_wait_cells(tabs: list[dict]) -> list[tuple[str, str]]:
    cells: list[tuple[str, str]] = []
    for tab in tabs:
        sheet = tab["name"]
        for addr in tab.get("wait_cells", []):
            cells.append((sheet, addr))
        for section in tab.get("sections", []):
            if section.get("range"):
                cells.append((sheet, section["range"].split(":")[0]))
        for addr in tab.get("cells", []):
            cells.append((sheet, addr.split(":")[0]))
    return cells


def wait_for_bloomberg(
    xl: Any,
    wb: Any,
    tabs: list[dict],
    timeout_seconds: int,
    poll_interval: float,
) -> None:
    deadline = time.time() + timeout_seconds
    wait_list = collect_wait_cells(tabs)
    while time.time() < deadline:
        pending = False
        for sheet_name, addr in wait_list:
            ws = wb.Worksheets(sheet_name)
            if is_bloomberg_pending(ws.Range(addr).Value):
                pending = True
                break
        if not pending and xl.CalculationState == 0:  # xlDone
            return
        time.sleep(poll_interval)
    raise TimeoutError("Bloomberg refresh did not complete within timeout")


def read_range_grid(ws: Any, range_addr: str, skip_empty_rows: bool = False) -> list[list[Any]]:
    grid = normalize_range_value(ws.Range(range_addr).Value)
    if skip_empty_rows:
        grid = filter_empty_rows(grid)
    return grid


def read_excel_tabs(wb: Any, tabs: list[dict]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for tab in tabs:
        ws = wb.Worksheets(tab["name"])
        tab_label = tab.get("label", tab["name"])

        if tab.get("sections"):
            sections: dict[str, dict] = {}
            for section in tab["sections"]:
                label = section.get("label", section.get("range", "Section"))
                header = (
                    read_range_grid(ws, section["header_range"])
                    if section.get("header_range")
                    else []
                )
                data = read_range_grid(
                    ws,
                    section["range"],
                    skip_empty_rows=section.get("skip_empty_rows", False),
                )
                sections[label] = {"header": header, "rows": data}
            result[tab_label] = {"sections": sections}
            continue

        cells: dict[str, list[list[Any]]] = {}
        for addr in tab.get("cells", []):
            cells[addr] = normalize_range_value(ws.Range(addr).Value)
        result[tab_label] = {"cells": cells}
    return result


def refresh_and_read(cfg: dict) -> dict[str, Any]:
    workbook_path = Path(cfg["workbook"])
    if not workbook_path.exists():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    addin_path = Path(cfg.get("bloomberg_addin", ""))
    tabs = cfg.get("tabs", [])
    xl = get_excel_app(addin_path if addin_path else None)

    wb = xl.Workbooks.Open(
        str(workbook_path),
        UpdateLinks=0,
        ReadOnly=cfg.get("read_only", True),
    )
    try:
        bloomberg_refresh(xl, wb, cfg.get("refresh_mode", "workbook"), tabs)
        wait_for_bloomberg(
            xl,
            wb,
            tabs,
            cfg.get("timeout_seconds", 300),
            cfg.get("poll_interval_seconds", 2),
        )
        return read_excel_tabs(wb, tabs)
    finally:
        if cfg.get("close_workbook", True):
            wb.Close(SaveChanges=False)
        if cfg.get("quit_excel", False):
            xl.Quit()


# ---------------------------------------------------------------------------
# Component 3: SPTSX Bloomberg (xbbg)
# ---------------------------------------------------------------------------


def _flatten_bdp(df: Any) -> dict[str, dict[str, Any]]:
    if df is None or df.empty:
        return {}
    out: dict[str, dict[str, Any]] = {}
    for idx, row in df.iterrows():
        ticker = str(idx[0] if isinstance(idx, tuple) else idx)
        out[ticker] = {str(col): row[col] for col in df.columns}
    return out


def _chunked(items: list[str], size: int) -> list[list[str]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def fetch_sptsx_members(index_ticker: str, fields: list[str], batch_size: int = 80) -> Any:
    import xbbg

    members_df = None
    try:
        members_df = xbbg.bds(index_ticker, "INDX_MEMBERS", backend="pandas")
    except Exception as exc:
        LOG.warning("bds INDX_MEMBERS failed: %s", exc)

    if members_df is None or members_df.empty:
        try:
            members_df = xbbg.index_members(index_ticker, backend="pandas")
        except Exception as exc:
            raise RuntimeError(f"Could not fetch members for {index_ticker}: {exc}") from exc

    member_col = None
    for col in members_df.columns:
        if "member" in str(col).lower() or "ticker" in str(col).lower() or "id" in str(col).lower():
            member_col = col
            break
    if member_col is None:
        member_col = members_df.columns[0]

    tickers = [str(t).strip() for t in members_df[member_col].dropna().tolist() if str(t).strip()]
    if not tickers:
        raise RuntimeError(f"No index members returned for {index_ticker}")

    frames = []
    for batch in _chunked(tickers, batch_size):
        part = xbbg.bdp(batch, fields, backend="pandas")
        if part is not None and not part.empty:
            frames.append(part)

    if not frames:
        raise RuntimeError(f"No member data returned for {index_ticker}")

    import pandas as pd

    return pd.concat(frames)


def fetch_sptsx_index_summary(cfg: dict) -> dict:
    import xbbg

    ticker = cfg.get("index_ticker", "SPTSX Index")
    fields = cfg.get(
        "index_fields",
        ["PX_LAST", "CHG_NET_1D", "CHG_PCT_1D", "VOLUME", "VOLUME_AVG_20D"],
    )
    df = xbbg.bdp(ticker, fields, backend="pandas")
    row = df.iloc[0] if not df.empty else {}

    last = row.get("PX_LAST", row.get("px_last"))
    chg = row.get("CHG_NET_1D", row.get("chg_net_1d"))
    chg_pct = row.get("CHG_PCT_1D", row.get("chg_pct_1d"))
    volume = row.get("VOLUME", row.get("volume"))
    avg_vol = row.get("VOLUME_AVG_20D", row.get("volume_avg_20d"))

    vol_ratio = ""
    if volume not in (None, "") and avg_vol not in (None, "", 0):
        try:
            vol_ratio = f"{float(volume) / float(avg_vol):.2f}x"
        except (TypeError, ValueError, ZeroDivisionError):
            vol_ratio = ""

    return {
        "label": f"{ticker} Summary",
        "type": "table",
        "table": {
            "columns": ["Last", "Change", "Change %", "Volume", "20D Avg Volume", "Vol vs 20D Avg"],
            "rows": [[
                fmt_cell(last),
                fmt_cell(chg),
                fmt_pct(chg_pct),
                fmt_cell(volume),
                fmt_cell(avg_vol),
                vol_ratio,
            ]],
        },
        "index_return_pct": float(chg_pct) if chg_pct not in (None, "") else 0.0,
    }


def _member_row_label(ticker: str, name: Any) -> str:
    if name not in (None, ""):
        return f"{ticker} — {name}"
    return ticker


def _subindustry_key(row: dict, code_field: str, name_field: str) -> str:
    code = row.get(code_field, row.get(code_field.lower(), ""))
    name = row.get(name_field, row.get(name_field.lower(), ""))
    code_s = "" if code in (None, "") else str(code).strip()
    name_s = "" if name in (None, "") else str(name).strip()
    if code_s and name_s:
        return f"{code_s} ({name_s})"
    return code_s or name_s or "Unclassified"


def build_mov_by_subindustry(members_df: Any, cfg: dict) -> dict:
    fields = cfg.get(
        "member_fields",
        {
            "name": "SHORT_NAME",
            "price": "PX_LAST",
            "change_pct": "CHG_PCT_1D",
            "points": "INDX_POINTS",
            "pct_ind_mkt_val": "PCT_IND_MKT_VAL",
            "subindustry_code": "GICS_SUB_INDUSTRY",
            "subindustry_name": "GICS_SUB_INDUSTRY_NAME",
        },
    )
    movers_per_side = cfg.get("movers_per_side", 3)
    columns = ["Ticker", "Price", "% Change", "Points", "%Ind Mv"]

    df = members_df.copy()
    df.index = [str(i[0] if isinstance(i, tuple) else i) for i in df.index]

    records = []
    for ticker, row in df.iterrows():
        rec = {"ticker": ticker}
        for key, bbg_field in fields.items():
            val = row.get(bbg_field, row.get(bbg_field.lower()))
            rec[key] = val
        try:
            rec["change_pct_num"] = float(rec.get("change_pct", 0) or 0)
        except (TypeError, ValueError):
            rec["change_pct_num"] = 0.0
        records.append(rec)

    groups: dict[str, list[dict]] = {}
    for rec in records:
        key = _subindustry_key(rec, fields["subindustry_code"], fields["subindustry_name"])
        groups.setdefault(key, []).append(rec)

    mov_groups = []
    for group_name in sorted(groups.keys()):
        items = sorted(groups[group_name], key=lambda r: r["change_pct_num"], reverse=True)
        leaders = items[:movers_per_side]
        laggards = list(reversed(items[-movers_per_side:])) if len(items) >= movers_per_side else list(reversed(items))
        mov_groups.append({
            "name": group_name,
            "leaders": [
                [
                    _member_row_label(r["ticker"], r.get("name")),
                    fmt_cell(r.get("price")),
                    fmt_pct(r.get("change_pct")),
                    fmt_cell(r.get("points")),
                    fmt_pct(r.get("pct_ind_mkt_val")),
                ]
                for r in leaders
            ],
            "laggards": [
                [
                    _member_row_label(r["ticker"], r.get("name")),
                    fmt_cell(r.get("price")),
                    fmt_pct(r.get("change_pct")),
                    fmt_cell(r.get("points")),
                    fmt_pct(r.get("pct_ind_mkt_val")),
                ]
                for r in laggards
            ],
        })

    return {
        "label": "SPTSX MOV by Sub-Industry (GICS Level 4)",
        "type": "mov_dual",
        "columns": columns,
        "groups": mov_groups,
    }


def build_top_bottom_chart(members_df: Any, index_return_pct: float, cfg: dict) -> Path:
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches

    top_n = cfg.get("top_n", 10)
    change_field = cfg.get("member_fields", {}).get("change_pct", "CHG_PCT_1D")
    name_field = cfg.get("member_fields", {}).get("name", "SHORT_NAME")

    rows = []
    for ticker, row in members_df.iterrows():
        ticker_s = str(ticker[0] if isinstance(ticker, tuple) else ticker)
        chg = row.get(change_field, row.get(change_field.lower()))
        name = row.get(name_field, row.get(name_field.lower(), ticker_s))
        try:
            chg_f = float(chg)
        except (TypeError, ValueError):
            continue
        label = str(name) if name not in (None, "") else ticker_s
        rows.append({"label": label, "change": chg_f})

    if not rows:
        raise RuntimeError("No member change data available for performers chart")

    rows.sort(key=lambda r: r["change"], reverse=True)
    gainers = rows[:top_n]
    losers = list(reversed(rows[-top_n:]))

    n_rows = max(len(gainers), len(losers))
    labels = [""] * n_rows
    left_vals = [0.0] * n_rows
    right_vals = [0.0] * n_rows

    for i, item in enumerate(losers):
        labels[i] = item["label"]
        left_vals[i] = item["change"]

    for i, item in enumerate(gainers):
        labels[i] = item["label"] or labels[i]
        right_vals[i] = item["change"]

    fig_h = max(7, n_rows * 0.55)
    fig, ax = plt.subplots(figsize=(13, fig_h))
    y_pos = list(range(n_rows))

    ax.barh(y_pos, left_vals, color="#c0392b", height=0.65, align="center")
    ax.barh(y_pos, right_vals, color="#27ae60", height=0.65, align="center")
    ax.axvline(0, color="#333333", linewidth=1.2)

    for i, label in enumerate(labels):
        ax.text(0, i, label, ha="center", va="center", fontsize=9, fontweight="bold", clip_on=False)

    ax.set_yticks([])
    ax.set_xlabel("Price Return %")
    ax.set_title(
        f"SPTSX Composite — Total Return {index_return_pct:.2f}%",
        fontsize=13,
        fontweight="bold",
        pad=12,
    )
    ax.legend(
        handles=[
            mpatches.Patch(color="#c0392b", label="Worst performers (left)"),
            mpatches.Patch(color="#27ae60", label="Best performers (right)"),
        ],
        loc="lower center",
        ncol=2,
        bbox_to_anchor=(0.5, -0.08),
    )
    ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()

    out_dir = Path(tempfile.gettempdir()) / "carlopdf"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"sptsx_performers_{date.today().isoformat()}.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    LOG.info("Performers chart saved to %s", out_path)
    return out_path


def fetch_sptsx_data(cfg: dict) -> dict[str, dict]:
    sptsx_cfg = cfg.get("sptsx", {})
    if not sptsx_cfg.get("enabled", True):
        return {}

    index_ticker = cfg.get("index_ticker", "SPTSX Index")
    member_fields_dict = sptsx_cfg.get(
        "member_fields",
        {
            "name": "SHORT_NAME",
            "price": "PX_LAST",
            "change_pct": "CHG_PCT_1D",
            "points": "INDX_POINTS",
            "pct_ind_mkt_val": "PCT_IND_MKT_VAL",
            "subindustry_code": "GICS_SUB_INDUSTRY",
            "subindustry_name": "GICS_SUB_INDUSTRY_NAME",
        },
    )
    member_field_names = list(dict.fromkeys(member_fields_dict.values()))

    result: dict[str, dict] = {}
    index_return = 0.0

    if sptsx_cfg.get("index_summary", True):
        summary = fetch_sptsx_index_summary(cfg)
        result[summary["label"]] = summary
        index_return = summary.get("index_return_pct", 0.0)

    members_df = None
    need_members = sptsx_cfg.get("mov_by_subindustry", True) or sptsx_cfg.get("top_bottom_chart", True)
    if need_members:
        members_df = fetch_sptsx_members(index_ticker, member_field_names)

    if sptsx_cfg.get("mov_by_subindustry", True) and members_df is not None:
        mov = build_mov_by_subindustry(members_df, {**sptsx_cfg, "member_fields": member_fields_dict})
        result[mov["label"]] = mov

    if sptsx_cfg.get("top_bottom_chart", True) and members_df is not None:
        chart_path = build_top_bottom_chart(
            members_df,
            index_return,
            {**sptsx_cfg, "member_fields": member_fields_dict},
        )
        label = "SPTSX Top/Bottom Performers"
        result[label] = {
            "label": label,
            "type": "image",
            "path": str(chart_path),
        }

    return result


def fetch_bbg_tables(cfg: dict) -> dict[str, dict]:
    result: dict[str, dict] = {}

    if cfg.get("sptsx"):
        result.update(fetch_sptsx_data(cfg))

    for table in cfg.get("tables", []):
        label = table.get("label", table.get("type", "table"))
        try:
            import xbbg

            table_type = table["type"].lower()
            if table_type == "bdp":
                df = xbbg.bdp(table["tickers"], table["fields"], backend="pandas")
            elif table_type == "bdh":
                df = xbbg.bdh(
                    table.get("ticker") or table.get("tickers"),
                    table["fields"],
                    table.get("start_date"),
                    table.get("end_date"),
                    backend="pandas",
                )
            elif table_type == "bds":
                df = xbbg.bds(
                    table.get("ticker") or table.get("tickers"),
                    table["field"],
                    backend="pandas",
                )
            else:
                raise ValueError(f"Unsupported table type: {table_type}")
            result[label] = {"label": label, "type": "table", "table": df_to_table(df)}
        except Exception as exc:
            if table.get("optional") or cfg.get("optional"):
                LOG.warning("Skipping optional Bloomberg table '%s': %s", label, exc)
                result[label] = {
                    "label": label,
                    "type": "table",
                    "table": {"columns": [], "rows": []},
                    "error": str(exc),
                }
            else:
                raise

    return result


# ---------------------------------------------------------------------------
# Output: Outlook draft popup
# ---------------------------------------------------------------------------


def excel_data_to_html(excel_data: dict[str, Any]) -> str:
    parts: list[str] = []
    for label, payload in excel_data.items():
        parts.append(f"<h3>{label}</h3>")

        if payload.get("sections"):
            for section_label, section in payload["sections"].items():
                parts.append(f"<h4>{section_label}</h4>")
                parts.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>")
                for row in section.get("header", []):
                    parts.append("<tr>" + "".join(f"<th>{fmt_cell(v)}</th>" for v in row) + "</tr>")
                for row in section.get("rows", []):
                    parts.append("<tr>" + "".join(f"<td>{fmt_cell(v)}</td>" for v in row) + "</tr>")
                parts.append("</table>")
            continue

        for addr, grid in payload.get("cells", {}).items():
            parts.append(f"<p><strong>{addr}</strong></p>")
            parts.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>")
            for row in grid:
                parts.append("<tr>" + "".join(f"<td>{fmt_cell(v)}</td>" for v in row) + "</tr>")
            parts.append("</table>")
    return "\n".join(parts)


def mov_dual_table_html(payload: dict) -> str:
    columns = payload.get("columns", [])
    parts = [
        "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;width:100%;'>",
        "<tr><th colspan='5' style='text-align:center;background:#eee;'>Leaders</th>"
        "<th colspan='5' style='text-align:center;background:#eee;'>Laggards</th></tr>",
        "<tr>"
        + "".join(f"<th>{c}</th>" for c in columns)
        + "".join(f"<th>{c}</th>" for c in columns)
        + "</tr>",
    ]
    for group in payload.get("groups", []):
        parts.append(
            f"<tr><td colspan='10' style='background:#f5f5f5;font-weight:bold;'>{group['name']}</td></tr>"
        )
        leaders = group.get("leaders", [])
        laggards = group.get("laggards", [])
        max_rows = max(len(leaders), len(laggards), 1)
        for i in range(max_rows):
            left = leaders[i] if i < len(leaders) else [""] * len(columns)
            right = laggards[i] if i < len(laggards) else [""] * len(columns)
            parts.append(
                "<tr>"
                + "".join(f"<td>{v}</td>" for v in left)
                + "".join(f"<td>{v}</td>" for v in right)
                + "</tr>"
            )
    parts.append("</table>")
    return "\n".join(parts)


def bbg_data_to_html(bbg_data: dict[str, dict]) -> str:
    parts: list[str] = []
    for payload in bbg_data.values():
        label = payload.get("label", "Bloomberg")
        parts.append(f"<h3>{label}</h3>")
        if payload.get("error"):
            parts.append(f"<p><em>{payload['error']}</em></p>")
            continue

        payload_type = payload.get("type", "table")
        if payload_type == "image":
            img_path = payload.get("path")
            if img_path and Path(img_path).exists():
                uri = image_data_uri(Path(img_path))
                parts.append(f'<p><img src="{uri}" alt="{label}"></p>')
            continue
        if payload_type == "mov_dual":
            parts.append(mov_dual_table_html(payload))
            continue

        table = payload.get("table", {"columns": [], "rows": []})
        cols = table.get("columns", [])
        rows = table.get("rows", [])
        if not cols:
            parts.append("<p><em>No data</em></p>")
            continue
        parts.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>")
        parts.append("<tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")
        for row in rows:
            parts.append("<tr>" + "".join(f"<td>{v}</td>" for v in row) + "</tr>")
        parts.append("</table>")
    return "\n".join(parts)


def build_draft_html(
    output_cfg: dict,
    excel_data: dict,
    bbg_data: dict,
    chart_path: Path | None,
) -> str:
    intro = output_cfg.get("intro_text", "")
    sections = [f"<p>{intro}</p>"]
    if chart_path and chart_path.exists():
        uri = image_data_uri(chart_path)
        sections.append(f'<p><img src="{uri}" alt="QES Chart"></p>')
    sections.append("<h2>Excel Data</h2>")
    sections.append(excel_data_to_html(excel_data))
    sections.append("<h2>Bloomberg Data</h2>")
    sections.append(bbg_data_to_html(bbg_data))
    body = "\n".join(sections)
    return f"<html><body style='font-family:Calibri,Arial,sans-serif;'>{body}</body></html>"


def open_outlook_draft(
    output_cfg: dict,
    chart_path: Path | None,
    excel_data: dict,
    bbg_data: dict,
) -> None:
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem

    subject = output_cfg.get("draft_subject", "Daily Market Update — {date}")
    subject = subject.replace("{date}", date.today().strftime("%Y-%m-%d"))
    mail.Subject = subject

    if output_cfg.get("draft_to"):
        mail.To = output_cfg["draft_to"]
    if output_cfg.get("draft_cc"):
        mail.CC = output_cfg["draft_cc"]

    mail.HTMLBody = build_draft_html(output_cfg, excel_data, bbg_data, chart_path)
    mail.Display()
    LOG.info("Outlook draft opened")


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------


def run_source(name: str, fn: Any, cfg: dict, optional: bool) -> Any:
    try:
        return fn(cfg)
    except Exception as exc:
        if optional:
            LOG.warning("%s failed (optional): %s", name, exc)
            return None
        LOG.error("%s failed: %s", name, exc)
        raise


def check_deployment(cfg: dict) -> int:
    """Preflight checks for the Bloomberg deployment machine."""
    ok = True

    def report(name: str, passed: bool, detail: str = "") -> None:
        nonlocal ok
        if passed:
            LOG.info("[OK] %s%s", name, f" — {detail}" if detail else "")
        else:
            ok = False
            LOG.error("[FAIL] %s%s", name, f" — {detail}" if detail else "")

    email_cfg = cfg.get("email_source", {})
    excel_cfg = cfg.get("excel", {})
    bbg_cfg = cfg.get("bloomberg_direct", {})

    subject = email_cfg.get("subject", "").strip()
    if not subject:
        report("email_source.subject", False, "subject is empty — set your QES email subject")
    else:
        report("email_source.subject", True, subject)

    folder_spec = email_cfg.get("folder", "Inbox")
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, folder_spec)
        report("email_source.folder", True, str(folder.Name))
    except Exception as exc:
        report("email_source.folder", False, str(exc))

    workbook = excel_cfg.get("workbook", "")
    if not workbook or "path/to" in workbook.replace("\\", "/").lower():
        report("excel.workbook", False, "set the real .xlsm path in CONFIG at top of run.py")
    elif not Path(workbook).exists():
        report("excel.workbook", False, f"file not found: {workbook}")
    else:
        report("excel.workbook", True, workbook)

    addin = excel_cfg.get("bloomberg_addin", "")
    addin_path = Path(addin) if addin else None
    if not addin_path or not addin_path.exists():
        alt = Path(r"C:\Program Files (x86)\blp\API\Office Tools\BloombergUI.xla")
        if alt.exists():
            report("excel.bloomberg_addin", True, f"configured path missing; found {alt}")
        else:
            report("excel.bloomberg_addin", False, f"not found: {addin}")
    else:
        report("excel.bloomberg_addin", True, str(addin_path))

    for tab in excel_cfg.get("tabs", []):
        report(f"excel tab '{tab.get('name')}'", True, "configured")

    ticker = bbg_cfg.get("index_ticker", "SPTSX Index")
    try:
        import xbbg

        df = xbbg.bdp(ticker, "PX_LAST", backend="pandas")
        if df is None or df.empty:
            report("xbbg / Bloomberg Terminal", False, f"no data for {ticker}")
        else:
            val = df.iloc[0, 0]
            report("xbbg / Bloomberg Terminal", True, f"{ticker} PX_LAST = {val}")
    except Exception as exc:
        report("xbbg / Bloomberg Terminal", False, str(exc))

    draft_to = cfg.get("output", {}).get("draft_to", "")
    if not draft_to:
        report("output.draft_to", False, "set recipient email address(es)")
    else:
        report("output.draft_to", True, draft_to)

    if ok:
        LOG.info("All preflight checks passed — ready to run.")
    else:
        LOG.error("Preflight checks failed — fix CONFIG in run.py before scheduling.")
    return 0 if ok else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Market Data Automation")
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="Optional JSON config override (default: CONFIG dict at top of run.py)",
    )
    parser.add_argument("--no-poll", action="store_true", help="Do not poll for late email")
    parser.add_argument(
        "--check",
        action="store_true",
        help="Run preflight checks only (Outlook, Excel paths, Bloomberg/xbbg)",
    )
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    if args.config:
        if not args.config.exists():
            LOG.error("Config file not found: %s", args.config)
            return 1
        cfg = load_config(args.config)
        LOG.info("Loaded config from %s", args.config)
    else:
        cfg = CONFIG

    if args.check:
        return check_deployment(cfg)

    state_path = SCRIPT_DIR / cfg.get("state_file", ".state.json")

    chart_path: Path | None = None
    excel_data: dict = {}
    bbg_data: dict = {}

    email_cfg = cfg.get("email_source", {})
    if email_cfg and not email_cfg.get("skip"):
        try:
            chart_path = fetch_chart(email_cfg, state_path, poll=not args.no_poll)
        except Exception as exc:
            if email_cfg.get("optional"):
                LOG.warning("Email chart skipped (optional): %s", exc)
            else:
                raise

    excel_cfg = cfg.get("excel", {})
    if excel_cfg and excel_cfg.get("workbook"):
        excel_data = run_source(
            "Excel",
            refresh_and_read,
            excel_cfg,
            excel_cfg.get("optional", False),
        ) or {}

    bbg_cfg = cfg.get("bloomberg_direct", {})
    if bbg_cfg and (bbg_cfg.get("sptsx") or bbg_cfg.get("tables")):
        try:
            bbg_data = fetch_bbg_tables(bbg_cfg)
        except Exception as exc:
            if bbg_cfg.get("optional"):
                LOG.warning("Bloomberg failed (optional): %s", exc)
            else:
                raise

    open_outlook_draft(cfg.get("output", {}), chart_path, excel_data, bbg_data)
    return 0


if __name__ == "__main__":
    sys.exit(main())
