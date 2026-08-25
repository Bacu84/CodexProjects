"""Build the static Pricing Engine HTML application from the pricing workbook.

IT review notes:
- This script reads a local Excel macro-enabled workbook as an OOXML zip file.
- It does not execute Excel, VBA, macros, external links, or Power Query.
- It writes one local HTML file next to this script and embeds the required data.
- Plotly is loaded from the local bundled `plotly-2.27.0.min.js` file, not from a CDN.
- The generated browser app persists manual price inputs only in browser storage
  and optional user-downloaded JSON/XLSX files; it does not upload data.
"""

from __future__ import annotations

import datetime as dt
import json
import math
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


# Local source workbook supplied by the business user.
WORKBOOK_PATH = Path(r"C:\Users\s.porciello\Downloads\PricingToolSchema26_FInal_IMPORT (1).xlsm")
ROOT = Path(__file__).resolve().parent
# Local Plotly bundle used by the generated static HTML application.
PLOTLY_PATH = ROOT / "plotly-2.27.0.min.js"
OUTPUT_PATH = ROOT / "pricing_engine_app.html"

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PACKAGE_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    """Convert an Excel column label such as 'A' or 'AK' to a 1-based index."""

def col_to_idx(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char.upper()) - 64
    return value
 """Convert a 1-based column index back to an Excel column label."""

def idx_to_col(idx: int) -> str:
    label = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        label = chr(65 + rem) + label
    return label

"""Split a cell reference such as 'AK10' into column and row numbers."""
def split_cell_ref(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"Unsupported cell reference: {ref}")
    return col_to_idx(match.group(1)), int(match.group(2))

 """Convert an Excel range such as 'A10:BK2904' into numeric boundaries."""
def range_bounds(ref: str) -> tuple[int, int, int, int]:
    start, end = ref.split(":")
    start_col, start_row = split_cell_ref(start)
    end_col, end_row = split_cell_ref(end)
    return start_col, start_row, end_col, end_row

"""Normalize workbook relationship targets to zip-internal OOXML paths."""
def package_path(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return f"xl/{target}"

 """Normalize material numbers and lookup keys so '50754' and 50754 match."""
def normalize_key(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

"""Return an Excel XML value as a number when possible, otherwise text."""
def numeric_or_text(raw: str | None):
    if raw is None:
        return None
    try:
        value = float(raw)
    except ValueError:
        return raw
    if math.isfinite(value) and value.is_integer():
        return int(value)
    return value


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Read Excel's shared string table without opening or executing Excel."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    strings: list[str] = []
    for _, elem in ET.iterparse(zf.open("xl/sharedStrings.xml"), events=("end",)):
        if elem.tag == MAIN_NS + "si":
            strings.append("".join((node.text or "") for node in elem.iter(MAIN_NS + "t")))
            elem.clear()
    return strings


def decode_cell(elem: ET.Element, shared_strings: list[str]):
    """Decode a raw OOXML cell value into a Python scalar."""
    cell_type = elem.attrib.get("t")
    value_elem = elem.find(MAIN_NS + "v")
    raw = value_elem.text if value_elem is not None else None

    if cell_type == "s" and raw is not None:
        try:
            return shared_strings[int(raw)]
        except (IndexError, ValueError):
            return raw
    if cell_type == "inlineStr":
        inline = elem.find(MAIN_NS + "is")
        if inline is None:
            return None
        return "".join((node.text or "") for node in inline.iter(MAIN_NS + "t"))
    if cell_type == "str":
        return raw
    if cell_type == "b" and raw is not None:
        return bool(int(raw))
    return numeric_or_text(raw)


def workbook_sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map visible worksheet names to their internal OOXML file paths."""
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    paths: dict[str, str] = {}
    sheets = workbook.find(MAIN_NS + "sheets")
    if sheets is None:
        return paths
    for sheet in sheets:
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[REL_NS + "id"]
        paths[name] = package_path(rel_map[rel_id])
    return paths


def table_definitions(zf: zipfile.ZipFile) -> dict[str, dict]:
    """Extract Excel table metadata so the PricingEngine table range is stable."""
    tables: dict[str, dict] = {}
    for name in zf.namelist():
        if not (name.startswith("xl/tables/") and name.endswith(".xml")):
            continue
        root = ET.fromstring(zf.read(name))
        table_name = root.attrib.get("name") or root.attrib.get("displayName")
        table_columns = root.find(MAIN_NS + "tableColumns")
        columns = [col.attrib.get("name", "") for col in table_columns] if table_columns is not None else []
        tables[table_name] = {
            "path": name,
            "ref": root.attrib["ref"],
            "columns": columns,
        }
    return tables


def table_owner_sheets(zf: zipfile.ZipFile) -> dict[str, str]:
    """Resolve each Excel table to the worksheet that owns it."""
    sheet_paths = workbook_sheet_paths(zf)
    path_to_name = {path: name for name, path in sheet_paths.items()}
    owners: dict[str, str] = {}
    for rel_path in zf.namelist():
        if not (rel_path.startswith("xl/worksheets/_rels/") and rel_path.endswith(".rels")):
            continue
        sheet_file = rel_path.rsplit("/", 1)[-1].replace(".rels", "")
        sheet_path = f"xl/worksheets/{sheet_file}"
        sheet_name = path_to_name.get(sheet_path, sheet_path)
        root = ET.fromstring(zf.read(rel_path))
        for rel in root:
            target = rel.attrib.get("Target", "")
            if "/tables/" in target or target.startswith("../tables/"):
                owners[target.rsplit("/", 1)[-1]] = sheet_name
    return owners


def parse_cells(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
    wanted_rows: set[int] | None = None,
    wanted_cols: set[int] | None = None,
) -> dict[tuple[int, int], object]:
    """Stream only the requested worksheet cells to keep memory usage controlled."""
    cells: dict[tuple[int, int], object] = {}
    for _, elem in ET.iterparse(zf.open(sheet_path), events=("end",)):
        if elem.tag == MAIN_NS + "c":
            ref = elem.attrib.get("r")
            if ref:
                col, row = split_cell_ref(ref)
                if (wanted_rows is None or row in wanted_rows) and (
                    wanted_cols is None or col in wanted_cols
                ):
                    cells[(row, col)] = decode_cell(elem, shared_strings)
            elem.clear()
    return cells


def parse_column_widths(zf: zipfile.ZipFile, sheet_path: str, col_count: int) -> list[int]:
    """Convert Excel column widths into practical pixel widths for the HTML table."""
    widths = [120 for _ in range(col_count)]
    text_like = {1, 2, 3, 4, 6, 7, 8, 51}
    for col in text_like:
        if col <= col_count:
            widths[col - 1] = 210 if col in {2, 6, 7, 8, 51} else 140

    for _, elem in ET.iterparse(zf.open(sheet_path), events=("end",)):
        if elem.tag == MAIN_NS + "col":
            min_col = int(elem.attrib.get("min", "1"))
            max_col = int(elem.attrib.get("max", str(min_col)))
            if elem.attrib.get("hidden") == "1":
                px = 0
            else:
                width = float(elem.attrib.get("width", "13"))
                px = max(76, min(260, int(width * 7.3 + 18)))
            for idx in range(min_col, min(max_col, col_count) + 1):
                widths[idx - 1] = px
            elem.clear()
        elif elem.tag == MAIN_NS + "sheetData":
            break
    return widths


def parse_pricing_engine(zf: zipfile.ZipFile, shared_strings: list[str]):
    """Read the front-end PricingEngine table and its summary row from the workbook."""
    sheet_paths = workbook_sheet_paths(zf)
    tables = table_definitions(zf)
    sheet_path = sheet_paths.get("PricingEngine") or sheet_paths.get("Pricing Engine")
    if sheet_path is None:
        raise RuntimeError("Could not find the PricingEngine sheet.")
    table = tables["PricingEngineMain"]
    start_col, header_row, end_col, end_row = range_bounds(table["ref"])
    col_count = end_col - start_col + 1
    wanted_rows = set(range(header_row, end_row + 1))
    wanted_rows.add(header_row - 1)
    wanted_cols = set(range(start_col, end_col + 1))
    cells = parse_cells(zf, sheet_path, shared_strings, wanted_rows, wanted_cols)
    headers = [cells.get((header_row, col), "") for col in range(start_col, end_col + 1)]
    summary = [cells.get((header_row - 1, col)) for col in range(start_col, end_col + 1)]
    rows = [
        [cells.get((row, col)) for col in range(start_col, end_col + 1)]
        for row in range(header_row + 1, end_row + 1)
    ]
    col_widths = parse_column_widths(zf, sheet_path, col_count)
    return {
        "headers": headers,
        "summary": summary,
        "rows": rows,
        "col_widths": col_widths,
        "header_row": header_row,
        "data_start_row": header_row + 1,
        "data_end_row": end_row,
    }


def extract_projected_nsv(zf: zipfile.ZipFile, shared_strings: list[str]) -> dict[str, float]:
    """Read helper values used to reproduce the workbook's Projected NSV logic."""
    sheet_paths = workbook_sheet_paths(zf)
    tables = table_definitions(zf)
    table_owners = table_owner_sheets(zf)
    trend_table = tables["LinearForecastTable"]
    mix_table = tables["FY27Ch.MixTable"]
    trend_sheet = table_owners.get(trend_table["path"].rsplit("/", 1)[-1])
    mix_sheet = table_owners.get(mix_table["path"].rsplit("/", 1)[-1])
    if not trend_sheet or trend_sheet != mix_sheet:
        raise RuntimeError("Could not resolve LinearForecastTable and FY27Ch.MixTable owner sheet.")
    helper_path = sheet_paths[trend_sheet]

    trend_start_col, trend_header_row, _, trend_end_row = range_bounds(trend_table["ref"])
    mix_start_col, mix_header_row, _, mix_end_row = range_bounds(mix_table["ref"])
    trend_col = trend_start_col + trend_table["columns"].index("TREND")
    projected_col = mix_start_col + mix_table["columns"].index("Projected NSV")
    first_row = max(trend_header_row, mix_header_row) + 1
    last_row = min(trend_end_row, mix_end_row)
    wanted_rows = set(range(first_row, last_row + 1))
    wanted_cols = {trend_col, projected_col}
    cells = parse_cells(zf, helper_path, shared_strings, wanted_rows, wanted_cols)

    projected: dict[str, float] = {}
    for row in range(first_row, last_row + 1):
        key = normalize_key(cells.get((row, trend_col)))
        value = cells.get((row, projected_col))
        if key and isinstance(value, (int, float)):
            projected[key] = float(value)
    return projected


def extract_cogs_rate_sum(zf: zipfile.ZipFile, shared_strings: list[str]) -> float:
    """Read the COGs rate constants used in GP FY and cash margin calculations."""
    sheet_paths = workbook_sheet_paths(zf)
    cogs_path = sheet_paths["COGs"]
    rows = {2, 3, 4}
    cols = {col_to_idx("AK")}
    cells = parse_cells(zf, cogs_path, shared_strings, rows, cols)
    return sum(float(value or 0) for value in cells.values() if isinstance(value, (int, float)))


def excel_div_minus_one(num: float, den: float) -> float:
    return num / den - 1 if den else 0


def as_num(value) -> float:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)) and math.isfinite(value):
        return float(value)
    try:
        cleaned = str(value).replace("%", "").replace(",", "").strip()
        return float(cleaned)
    except (TypeError, ValueError):
        return 0.0


def strict_num(value) -> float:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)) and math.isfinite(value):
        return float(value)
    try:
        cleaned = str(value).replace("%", "").replace(",", "").strip()
        if not cleaned:
            return math.nan
        return float(cleaned)
    except (TypeError, ValueError):
        return math.nan


def is_discontinued_status(value) -> bool:
    """Return True for both 'discontinued' and 'to be discontinued' product states."""
    status = str(value or "").strip().lower()
    return "discontinued" in status


def excel_launch_fy(value):
    """Convert Excel launch-year values to an integer fiscal year when possible."""
    if isinstance(value, (int, float)) and math.isfinite(value):
        if value > 30000:
            date_value = dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
            return date_value.year + (1 if date_value.month >= 10 else 0)
        return value
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return excel_launch_fy(float(text))
    except ValueError:
        return None


def extract_linear_forecast_model(
    zf: zipfile.ZipFile,
    shared_strings: list[str],
    pricing_rows: list[list],
    pricing_headers: list[str],
) -> dict:
    """Extract data needed to recalculate Linear Forecast FY27 in the browser.

    The returned model stores historical price and channel quantity data by
    material number. The generated JavaScript uses the same inputs when a user
    manually changes ZEVP FY.
    """
    idx = {header: i for i, header in enumerate(pricing_headers)}
    pricing_keys = {
        normalize_key(row[idx["Material No."]])
        for row in pricing_rows
        if normalize_key(row[idx["Material No."]])
    }

    sheet_paths = workbook_sheet_paths(zf)
    tables = table_definitions(zf)
    linear_table = tables["LinearForecastTable"]
    mix_table = tables["FY27Ch.MixTable"]
    channel_names = linear_table["columns"][1:]
    channel_count = len(channel_names)

    price_cols = {col_to_idx(col) for col in ["B", "N", "O", "P"]}
    price_cells = parse_cells(
        zf,
        sheet_paths["Prices"],
        shared_strings,
        wanted_rows=set(range(2, 2258)),
        wanted_cols=price_cols,
    )
    price_map: dict[str, list[float]] = {}
    for row in range(2, 2258):
        key = normalize_key(price_cells.get((row, col_to_idx("B"))))
        if key:
            price_map[key] = [
                as_num(price_cells.get((row, col_to_idx("P")))),
                as_num(price_cells.get((row, col_to_idx("O")))),
                as_num(price_cells.get((row, col_to_idx("N")))),
            ]

    start_col, header_row, _, end_row = range_bounds(linear_table["ref"])
    mix_start_col, _, _, _ = range_bounds(mix_table["ref"])
    total_col = mix_start_col + mix_table["columns"].index("TOT")
    trend_cols = [start_col + 1 + offset for offset in range(channel_count)]
    qty_cols = [
        [col_to_idx("D") + offset, col_to_idx("AI") + offset, col_to_idx("BN") + offset]
        for offset in range(channel_count)
    ]

    wanted_cols = {col_to_idx("A"), col_to_idx("B"), total_col, *trend_cols}
    for cols in qty_cols:
        wanted_cols.update(cols)
    forecast_cells = parse_cells(
        zf,
        sheet_paths["Channel Trend"],
        shared_strings,
        wanted_rows=set(range(header_row + 1, end_row + 1)),
        wanted_cols=wanted_cols,
    )

    models = {}
    for row_num in range(header_row + 1, end_row + 1):
        key = normalize_key(forecast_cells.get((row_num, col_to_idx("A"))))
        if not key or key not in pricing_keys or key not in price_map:
            continue
        trend_sum = sum(
            as_num(forecast_cells.get((row_num, col_index)))
            for col_index in trend_cols
            if isinstance(forecast_cells.get((row_num, col_index)), (int, float))
        )
        total = as_num(forecast_cells.get((row_num, total_col)))
        qty = [
            [as_num(forecast_cells.get((row_num, col_index))) for col_index in cols]
            for cols in qty_cols
        ]
        models[key] = {
            "launchFY": excel_launch_fy(forecast_cells.get((row_num, col_to_idx("B")))),
            "baseTotal": total - trend_sum,
            "priceHistory": price_map[key],
            "qty": qty,
        }

    return {
        "histYears": [2023, 2024, 2025],
        "channels": channel_names,
        "models": models,
        "modelRows": len(models),
    }


def recalc_for_qa(row: list, projected_base: float, idx: dict[str, int], u10: float, cogs_rate: float) -> dict[str, float]:
    """Recalculate key row formulas for QA against the workbook values.

    This is intentionally a small mirror of the browser-side formulas. It helps
    detect extraction or formula drift before writing the final HTML file.
    """
    out = row[:]
    zevp_fy = as_num(out[idx["ZEVP FY"]])
    zgru_2027 = as_num(out[idx["ZGRU 2027"]])
    out[idx["ZBRU FY"]] = zevp_fy / u10 if u10 else 0
    out[idx["ZGRU FY"]] = zgru_2027
    out[idx["ZNEK FY"]] = as_num(out[idx["ZBRU FY"]]) * (1 + as_num(out[idx["ZGRU FY"]]))
    out[idx["RET Diff"]] = excel_div_minus_one(zevp_fy, as_num(out[idx["ZEVP 2026"]]))
    out[idx["Trade Diff ZNEK"]] = excel_div_minus_one(as_num(out[idx["ZNEK FY"]]), as_num(out[idx["ZNEK 2026"]]))
    oase_avg = (as_num(out[idx["Oase Cheapest"]]) + as_num(out[idx["Oase Highest"]])) / 2
    out[idx["Street Price FY27"]] = zevp_fy * (1 - 0.18) if oase_avg == 0 else oase_avg
    out[idx["Street Price Diff SY"]] = excel_div_minus_one(as_num(out[idx["Street Price FY27"]]), zevp_fy)
    nsv26_strict = strict_num(out[idx["NSV 26 YTD"]])
    theoretical = nsv26_strict + excel_div_minus_one(zevp_fy, as_num(out[idx["ZEVP 2025"]]))
    out[idx["Theoretical NSV FY"]] = theoretical if math.isfinite(theoretical) else as_num(out[idx["NSV 25 Full Year"]])
    condition_mix = as_num(out[idx["Condition Mix"]])
    base = strict_num(out[idx["NSV 26 YTD"]]) if projected_base == 0 else projected_base
    raw_projected = base * (1 + as_num(out[idx["RET Diff"]]) + condition_mix)
    theoretical_nsv = as_num(out[idx["Theoretical NSV FY"]])
    if not math.isfinite(raw_projected):
        out[idx["Projected NSV 27"]] = 0
    elif raw_projected >= theoretical_nsv * 1.25 or raw_projected <= theoretical_nsv * 0.75:
        out[idx["Projected NSV 27"]] = theoretical_nsv
    else:
        out[idx["Projected NSV 27"]] = raw_projected
    out[idx["Forecasted NSV Change FY27"]] = excel_div_minus_one(as_num(out[idx["Projected NSV 27"]]), theoretical_nsv)
    projected_nsv = as_num(out[idx["Projected NSV 27"]])
    if projected_nsv:
        out[idx["GP FY"]] = (
            projected_nsv
            - (as_num(out[idx["FIFO Stack Cost"]]) + projected_nsv * cogs_rate)
        ) / projected_nsv
    else:
        out[idx["GP FY"]] = 0
    out[idx["Margin Erosion"]] = as_num(out[idx["GP FY"]]) - as_num(out[idx["GP Margin SY"]])
    pacemaker_qty = as_num(out[idx[" PaceMaker Forecast FY27"]])
    fallback_to_linear_qty = pacemaker_qty == 0 and not is_discontinued_status(out[idx["Product Status"]])
    effective_pacemaker_qty = as_num(out[idx["Linear Forecast FY27"]]) if fallback_to_linear_qty else pacemaker_qty
    out[idx["Revenue @Forecast Old Price"]] = as_num(out[idx["Theoretical NSV FY"]]) * effective_pacemaker_qty
    out[idx["Revenue @Forecast New Price Pacemaker"]] = effective_pacemaker_qty * as_num(out[idx["Projected NSV 27"]])
    out[idx["Revenue @Forecast New Price Linear"]] = as_num(out[idx["Projected NSV 27"]]) * as_num(out[idx["Linear Forecast FY27"]])
    out[idx["Cash Margin FY26 FY"]] = (
        as_num(out[idx["I.O. QTY 26 Full Year"]])
        * (as_num(out[idx["NSV 26 YTD"]]) * as_num(out[idx["GP Margin SY"]]))
    )
    out[idx["Cash Margin FY27FY Pacemaker"]] = (
        as_num(out[idx["Projected NSV 27"]])
        * effective_pacemaker_qty
        * as_num(out[idx["GP FY"]])
    )
    return {header: as_num(out[col]) for header, col in idx.items()}


def qa_recalc(rows: list[list], projected_base: list[float], headers: list[str], u10: float, cogs_rate: float) -> str:
    """Return maximum formula differences between extracted and recalculated values."""
    idx = {header: i for i, header in enumerate(headers)}
    fields = [
        "ZBRU FY",
        "ZGRU FY",
        "ZNEK FY",
        "RET Diff",
        "Trade Diff ZNEK",
        "Theoretical NSV FY",
        "Projected NSV 27",
        "GP FY",
        "Margin Erosion",
        "Revenue @Forecast New Price Pacemaker",
        "Cash Margin FY26 FY",
        "Cash Margin FY27FY Pacemaker",
    ]
    max_diffs = {field: 0.0 for field in fields}
    for row, base in zip(rows, projected_base):
        calc = recalc_for_qa(row, base, idx, u10, cogs_rate)
        for field in fields:
            diff = abs(calc[field] - as_num(row[idx[field]]))
            if math.isfinite(diff):
                max_diffs[field] = max(max_diffs[field], diff)
    return ", ".join(f"{field}: {diff:.6g}" for field, diff in max_diffs.items())


def build_payload() -> dict:
    """Build the JSON payload embedded in the static HTML app."""
    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(WORKBOOK_PATH)
    if not PLOTLY_PATH.exists():
        raise FileNotFoundError(PLOTLY_PATH)

    # The .xlsm file is treated as a zip archive of XML files. Excel itself is
    # not launched and workbook macros are not executed.
    with zipfile.ZipFile(WORKBOOK_PATH) as zf:
        shared_strings = load_shared_strings(zf)
        pricing = parse_pricing_engine(zf, shared_strings)
        linear_forecast = extract_linear_forecast_model(
            zf,
            shared_strings,
            pricing["rows"],
            pricing["headers"],
        )
        projected_map = extract_projected_nsv(zf, shared_strings)
        cogs_rate = extract_cogs_rate_sum(zf, shared_strings)

    headers = pricing["headers"]
    idx = {header: i for i, header in enumerate(headers)}
    projected_base = [
        projected_map.get(normalize_key(row[idx["Material No."]]), 0.0)
        for row in pricing["rows"]
    ]
    u10 = as_num(pricing["summary"][idx["ZBRU 2026"]]) or 1.19
    # QA output is printed by main() and is useful evidence for IT/business
    # validation that the exported app still mirrors the workbook formulas.
    qa = qa_recalc(pricing["rows"], projected_base, headers, u10, cogs_rate)
    hit_count = sum(1 for value in projected_base if value)

    return {
        "sourceFile": WORKBOOK_PATH.name,
        "generatedAt": dt.datetime.now().isoformat(timespec="seconds"),
        "sheetName": "PricingEngine",
        "headerRow": pricing["header_row"],
        "dataStartRow": pricing["data_start_row"],
        "dataEndRow": pricing["data_end_row"],
        "headers": headers,
        "summaryRow": pricing["summary"],
        "rows": pricing["rows"],
        "projectedBase": projected_base,
        "projectedBaseHits": hit_count,
        "linearForecast": linear_forecast,
        "colWidths": pricing["col_widths"],
        "constants": {
            "u10": u10,
            "cogsRateSum": cogs_rate,
        },
        "qa": qa,
    }


# CSS template embedded into the generated single-file HTML app.
STYLE = r"""
:root {
  color-scheme: light;
  --text: #202124;
  --muted: #667085;
  --line: #d9dee7;
  --panel: #ffffff;
  --soft: #f6f8fb;
  --teal: #0f766e;
  --blue: #2563eb;
  --orange: #c2410c;
  --beige: #f5ead6;
  --beige-border: #d9bd8c;
  --positive: #047857;
  --negative: #b42318;
}
* { box-sizing: border-box; }
html, body { height: 100%; }
body {
  margin: 0;
  font-family: "Segoe UI", Arial, sans-serif;
  color: var(--text);
  background: #eef2f6;
}
button, input, select {
  font: inherit;
}
.app {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}
.topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
  padding: 16px 22px;
  background: #ffffff;
  border-bottom: 1px solid var(--line);
}
.title h1 {
  margin: 0;
  font-size: 22px;
  font-weight: 700;
  letter-spacing: 0;
}
.title span {
  display: block;
  margin-top: 3px;
  color: var(--muted);
  font-size: 12px;
}
.actions {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  justify-content: flex-end;
}
.save-status {
  min-width: 132px;
  color: var(--muted);
  font-size: 11px;
  text-align: right;
}
.save-status.saved {
  color: var(--positive);
}
.save-status.error {
  color: var(--negative);
}
.file-input-hidden {
  display: none;
}
.btn {
  min-height: 36px;
  border: 1px solid var(--line);
  background: #fff;
  color: var(--text);
  border-radius: 6px;
  padding: 0 12px;
  cursor: pointer;
}
.btn.primary {
  background: var(--teal);
  border-color: var(--teal);
  color: #fff;
  font-weight: 650;
}
.btn:active {
  transform: translateY(1px);
}
.workspace {
  display: grid;
  grid-template-columns: 270px minmax(0, 1fr);
  gap: 12px;
  padding: 12px;
  flex: 1;
  align-items: start;
}
.sidebar {
  grid-column: 1;
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 10px;
}
.panel h2 {
  margin: 0 0 9px;
  font-size: 14px;
  font-weight: 700;
}
.field {
  display: grid;
  gap: 5px;
  margin-bottom: 8px;
}
.field label,
.slider-row label {
  color: #344054;
  font-size: 12px;
  font-weight: 650;
}
.field input:not([type="checkbox"]),
.field select {
  min-width: 0;
  height: 34px;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 0 9px;
  background: #fff;
}
.multi-filter {
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  overflow: hidden;
}
.multi-filter summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  min-height: 34px;
  padding: 0 9px;
  cursor: pointer;
  list-style: none;
  color: var(--text);
  font-size: 12px;
}
.multi-filter summary::-webkit-details-marker {
  display: none;
}
.multi-filter summary::after {
  content: "";
  width: 0;
  height: 0;
  border-left: 4px solid transparent;
  border-right: 4px solid transparent;
  border-top: 5px solid var(--muted);
  flex: 0 0 auto;
}
.multi-filter[open] summary {
  border-bottom: 1px solid var(--line);
}
.multi-filter[open] summary::after {
  transform: rotate(180deg);
}
.multi-filter .summary-text {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.multi-filter-actions {
  display: flex;
  gap: 6px;
  padding: 7px 6px 5px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfe;
}
.multi-filter-actions button {
  flex: 1;
  min-height: 26px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #ffffff;
  color: #344054;
  cursor: pointer;
  font-size: 11px;
}
.multi-filter-actions button:hover {
  background: var(--soft);
}
.multi-filter-search-wrap {
  padding: 7px 6px 5px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfe;
}
.multi-filter-search {
  width: 100%;
  height: 30px;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 0 8px;
  background: #ffffff;
  color: var(--text);
  font-size: 12px;
}
.multi-options {
  max-height: 190px;
  overflow: auto;
  padding: 6px;
  display: grid;
  gap: 2px;
}
.multi-option {
  display: grid;
  grid-template-columns: 16px minmax(0, 1fr);
  align-items: center;
  gap: 7px;
  padding: 5px 4px;
  border-radius: 5px;
  color: #344054;
  font-size: 12px;
}
.multi-option:hover {
  background: var(--soft);
}
.multi-option input {
  width: 14px;
  height: 14px;
  margin: 0;
  accent-color: var(--teal);
}
.multi-option span {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.slider-row {
  display: grid;
  grid-template-columns: 1fr auto;
  align-items: center;
  gap: 6px;
  margin-bottom: 8px;
}
.slider-row input[type="range"] {
  grid-column: 1 / -1;
  width: 100%;
  accent-color: var(--teal);
}
.slider-value {
  min-width: 56px;
  text-align: right;
  color: var(--muted);
  font-size: 12px;
}
.metrics {
  display: grid;
  grid-template-columns: repeat(9, minmax(0, 1fr));
  justify-content: stretch;
  align-items: stretch;
  gap: 3px;
  margin-bottom: 9px;
}
.metric {
  position: relative;
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 5px 4px;
  min-height: 48px;
  min-width: 0;
}
.metric .label {
  color: var(--muted);
  font-size: 8.8px;
  line-height: 1.12;
  min-height: 20px;
  margin-bottom: 2px;
  overflow: hidden;
}
.metric .value {
  font-size: 11.8px;
  line-height: 1.12;
  font-weight: 750;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.metric .sub {
  margin-top: 1px;
  font-size: 7.8px;
  color: var(--muted);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.metric.has-badge .label {
  padding-right: 48px;
}
.metric-badge {
  position: absolute;
  top: 4px;
  right: 4px;
  max-width: 50px;
  padding: 2px 4px;
  border-radius: 999px;
  background: #edfdf8;
  color: var(--positive);
  border: 1px solid #b9eadb;
  font-size: 7.5px;
  font-weight: 750;
  line-height: 1;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.metric-badge.negative {
  background: #fff1f0;
  color: var(--negative);
  border-color: #ffd3cf;
}
.main {
  grid-column: 2;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 14px;
}
.charts {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
}
.insight-charts {
  margin-top: -2px;
}
.chart {
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 8px;
  min-height: 330px;
  padding: 8px;
}
.insight-charts .chart {
  min-height: 300px;
}
.chart-scroll {
  max-height: 360px;
  overflow-y: auto;
  overflow-x: hidden;
}
.chart-scroll-inner {
  min-height: 300px;
}
.table-panel {
  grid-column: 1 / -1;
  min-width: 0;
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 8px;
  overflow: hidden;
}
.table-toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfe;
}
.table-toolbar .left,
.table-toolbar .right {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}
.table-toolbar select {
  height: 32px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
}
.row-count {
  color: var(--muted);
  font-size: 12px;
}
.grid-wrap {
  overflow: auto;
  max-height: 650px;
}
table {
  border-collapse: separate;
  border-spacing: 0;
  width: max-content;
  min-width: 100%;
  font-size: 12px;
}
thead th {
  position: sticky;
  top: 0;
  z-index: 4;
  background: #2f3a4a;
  color: #fff;
  text-align: left;
  vertical-align: top;
  font-weight: 650;
  border-right: 1px solid #465366;
  border-bottom: 1px solid #1f2937;
  padding: 7px 6px;
  white-space: normal;
  box-sizing: border-box;
}
.th-content {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 22px 8px;
  align-items: stretch;
  gap: 5px;
  min-height: 30px;
}
.th-title {
  min-width: 0;
  display: block;
  overflow: visible;
  overflow-wrap: anywhere;
  padding-top: 3px;
  line-height: 1.15;
  white-space: normal;
}
.th-filter {
  width: 22px;
  height: 22px;
  border: 1px solid rgba(255, 255, 255, 0.28);
  border-radius: 5px;
  background: rgba(255, 255, 255, 0.08);
  cursor: pointer;
  position: relative;
}
.th-filter::before {
  content: "";
  position: absolute;
  left: 6px;
  top: 5px;
  width: 0;
  height: 0;
  border-left: 4px solid transparent;
  border-right: 4px solid transparent;
  border-top: 7px solid #ffffff;
}
.th-filter::after {
  content: "";
  position: absolute;
  left: 9px;
  top: 11px;
  width: 2px;
  height: 5px;
  background: #ffffff;
}
.th-filter:hover,
.th-filter.active {
  background: #0f766e;
  border-color: #8ee5d8;
}
.th-resize {
  width: 8px;
  align-self: stretch;
  border: 0;
  border-radius: 4px;
  background: transparent;
  cursor: col-resize;
  padding: 0;
  position: relative;
  touch-action: none;
}
.th-resize::after {
  content: "";
  position: absolute;
  top: 2px;
  bottom: 2px;
  left: 3px;
  width: 2px;
  border-radius: 999px;
  background: rgba(255, 255, 255, 0.24);
}
.th-resize:hover::after,
.th-resize:focus-visible::after {
  background: #8ee5d8;
}
body.resizing-column {
  cursor: col-resize;
  user-select: none;
}
thead th.input-col {
  background: #8a6f3a;
  color: #fff;
}
tbody td {
  border-right: 1px solid #e6e9ef;
  border-bottom: 1px solid #edf0f5;
  padding: 6px 7px;
  white-space: nowrap;
  background: #fff;
  max-width: 280px;
  overflow: hidden;
  text-overflow: ellipsis;
}
tbody tr:nth-child(even) td {
  background: #fbfcfe;
}
tbody td.num {
  text-align: right;
  font-variant-numeric: tabular-nums;
}
tbody td.input-col {
  background: var(--beige);
  border-right-color: var(--beige-border);
  border-left: 1px solid var(--beige-border);
}
tbody tr:nth-child(even) td.input-col {
  background: #f7eddd;
}
.cell-input {
  width: 100%;
  min-width: 76px;
  height: 24px;
  border: 1px solid var(--beige-border);
  border-radius: 5px;
  background: #fffaf1;
  padding: 0 5px;
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.cell-input.invalid {
  border-color: #f04438;
  background: #fff1f0;
}
.sticky-a {
  position: sticky;
  left: 0;
  z-index: 3;
}
.sticky-b {
  position: sticky;
  left: 140px;
  z-index: 3;
}
thead .sticky-a,
thead .sticky-b {
  z-index: 5;
}
tbody .sticky-a,
tbody .sticky-b {
  background: #fff;
}
tbody tr:nth-child(even) .sticky-a,
tbody tr:nth-child(even) .sticky-b {
  background: #fbfcfe;
}
.good { color: var(--positive); }
.bad { color: var(--negative); }
.column-filter-popover {
  position: fixed;
  z-index: 50;
  width: min(300px, calc(100vw - 24px));
  max-height: min(430px, calc(100vh - 24px));
  background: #ffffff;
  border: 1px solid var(--line);
  border-radius: 8px;
  box-shadow: 0 18px 44px rgba(15, 23, 42, 0.18);
  overflow: hidden;
}
.column-filter-head {
  display: grid;
  gap: 8px;
  padding: 10px;
  border-bottom: 1px solid var(--line);
  background: #fbfcfe;
}
.column-filter-title {
  font-size: 12px;
  font-weight: 700;
  color: #344054;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.column-filter-search {
  height: 32px;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 0 8px;
}
.column-filter-actions {
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
}
.column-filter-actions button {
  height: 28px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #ffffff;
  cursor: pointer;
  padding: 0 8px;
  font-size: 12px;
}
.column-filter-actions button.active {
  background: #0f766e;
  border-color: #0f766e;
  color: #ffffff;
}
.column-filter-list {
  max-height: 300px;
  overflow: auto;
  padding: 6px;
}
.column-filter-option {
  display: grid;
  grid-template-columns: 16px minmax(0, 1fr);
  gap: 7px;
  align-items: center;
  padding: 5px 4px;
  border-radius: 5px;
  font-size: 12px;
}
.column-filter-option:hover {
  background: var(--soft);
}
.column-filter-option input {
  width: 14px;
  height: 14px;
  margin: 0;
  accent-color: var(--teal);
}
.column-filter-option span {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
@media (max-width: 1100px) {
  .workspace {
    grid-template-columns: 1fr;
  }
  .sidebar,
  .main,
  .table-panel {
    grid-column: 1;
  }
  .metrics {
    grid-template-columns: repeat(9, minmax(0, 1fr));
  }
  .charts,
  .insight-charts {
    grid-template-columns: 1fr;
  }
}
@media (max-width: 640px) {
  .topbar {
    align-items: flex-start;
    flex-direction: column;
  }
  .metrics {
    grid-template-columns: repeat(9, minmax(72px, 1fr));
    overflow-x: auto;
  }
}
"""


APP_JS = r"""
(() => {
  // Runtime data is embedded in the HTML by the Python builder. The app does
  // not call a server or external API after the file is opened in the browser.
  const payload = JSON.parse(document.getElementById("pricing-data").textContent);
  const headers = payload.headers;
  // rows is the mutable working copy used by filters, charts, and manual edits.
  // baseRows keeps the starting values so scenario sliders can reset cleanly.
  const rows = payload.rows.map((row) => row.slice());
  const baseRows = payload.rows.map((row) => row.slice());
  const projectedBase = payload.projectedBase.map((value) => Number(value) || 0);
  const linearForecast = payload.linearForecast || { histYears: [2023, 2024, 2025], channels: [], models: {} };
  const constants = payload.constants;
  const col = Object.fromEntries(headers.map((header, index) => [header, index]));

  const required = [
    "Material No.", "Material Descr.", "Product Status", "Business Segment", "Market",
    "Product Group", "Product Family", "Market Segment", "Segment", "ZEVP 2026",
    "ZGRU 2027", "ZEVP 2025", "ZNEK 2026", "I.O. QTY 26 Full Year",
    "Sales Qty YTD", "% Share Qty YTD", " PaceMaker Forecast FY27", "Linear Forecast FY27", "NSV 26 YTD",
    "NSV 25 Full Year", "GP Margin SY", "ZEVP FY", "ZBRU FY", "ZGRU FY",
    "ZNEK FY", "RET Diff", "Trade Diff ZNEK", "Street Price Diff SY",
    "Theoretical NSV FY", "Projected NSV 27", "Forecasted NSV Change FY27",
    "GP FY", "Street Price FY27", "Margin Erosion", "Condition Mix",
    "Oase Cheapest", "Oase Highest", "Revenue FY26 Full Year",
    "Revenue @Forecast Old Price", "Revenue @Forecast New Price Pacemaker",
    "Revenue @Forecast New Price Linear", "Cash Margin FY26 FY",
    "Cash Margin FY27FY Pacemaker", "FIFO Stack Cost", "FIFO Stack Cost Change",
    "Vol.Change FY27 pace.ai"
  ];
  for (const name of required) {
    if (col[name] === undefined) throw new Error(`Missing column: ${name}`);
  }

  const percentHeaders = new Set([
    "Controlling Cost Change", "FIFO Stack Cost Change", "% Share Qty YTD",
    "ZGRU 2026", "ZGRU 2027", "Vol.Change FY26vs25", "Vol.Change FY27 pace.ai",
    "Revenue YOY % 26vs25",
    "GP Margin SY", "ZGRU FY", "RET Diff", "Trade Diff ZNEK",
    "Street Price Diff SY", "Forecasted NSV Change FY27", "GP FY",
    "Margin Erosion", "Condition Mix"
  ]);
  const currencyHeaders = new Set([
    "CoGs SY", "COGs FY", "FIFO Stack Cost", "Rolling Sum",
    "ZEVP 2026", "ZBRU 2026", "ZEVP 2025", "ZNEK 2026", "ZNEK 2025",
    "Revenue FY26 Full Year", "Revenue FY25 Full Year",
    "NSV 26 YTD", "NSV 25 Full Year", "NSV 24 Full Year",
    "ZEVP FY", "ZBRU FY", "ZNEK FY", "Theoretical NSV FY", "Projected NSV 27",
    "Street Price FY27", "Cheapest COMP.Price", "Highest COMP.Price",
    "Oase Cheapest", "Oase Highest", "Revenue @Forecast Old Price",
    "Revenue @Forecast New Price Linear", "Cash Margin FY26 FY",
    "Cash Margin FY27FY Pacemaker"
  ]);
  const editableHeaders = new Set(["ZGRU 2027", "ZEVP FY"]);
  // Browser-side persistence for manual pricing inputs. Data remains local to
  // the user's browser unless the user explicitly downloads/exports it.
  const savedInputKey = `pricing-engine-inputs:v1:${payload.sourceFile || "pricing_engine_app"}`;
  const savedInputBackupKey = `${savedInputKey}:backup`;
  const savedInputDbName = "pricing-engine-saved-prices";
  const savedInputDbStore = "snapshots";
  const savedInputFields = ["ZEVP FY", "ZGRU 2027"];
  const dirtyInputs = new Map();
  let saveDebounceTimer = null;
  let saveStatusTimer = null;
  let pricingDbPromise = null;
  // Security limits: prevent unbounded saved-price imports from freezing the browser tab.
  // Maximum file size for imported JSON files (10 MB).
  const MAX_IMPORT_FILE_SIZE = 10 * 1024 * 1024;
  // Maximum number of rows in a saved payload (10x typical dataset size).
  const MAX_SAVED_ROWS = 100000;
  const state = {
    filters: {},
    search: "",
    minRevenue: 0,
    page: 0,
    pageSize: 100,
    filtered: rows.map((_, index) => index),
    columnFilters: {},
    tableSort: { column: null, direction: null },
    chartReady: false,
    columnWidths: payload.colWidths.map((width, index) => {
      const parsed = Number(width) || 120;
      const min = headers[index] === "Product Status" ? 140 : 90;
      return Math.max(parsed, min);
    })
  };
  const filterFields = ["Product Status", "Business Segment", "Market", "Market Segment", "Segment", "Product Group", "Product Family"];

  function minColumnWidth(index, header = headers[index]) {
    if (header === "Product Status") return 140;
    if (index === col["Material Descr."]) return 150;
    if (index === col["Material No."]) return 96;
    return 88;
  }

  function columnWidth(index, header = headers[index]) {
    const width = Number(state.columnWidths[index] ?? payload.colWidths[index] ?? 120);
    return Math.max(Math.round(width || 120), minColumnWidth(index, header));
  }

  function stickyMaterialWidth() {
    return columnWidth(col["Material No."], "Material No.");
  }

  function tableCellStyle(index, header = headers[index]) {
    const width = columnWidth(index, header);
    const stickyLeft = index === col["Material Descr."] ? `left:${stickyMaterialWidth()}px;` : "";
    return `${stickyLeft}width:${width}px;min-width:${width}px;max-width:${width}px;`;
  }

  function applyTableColumnWidths() {
    document.querySelectorAll("[data-col-index]").forEach((element) => {
      const index = Number(element.dataset.colIndex);
      const width = columnWidth(index);
      element.style.width = `${width}px`;
      element.style.minWidth = `${width}px`;
      element.style.maxWidth = `${width}px`;
      if (index === col["Material Descr."]) {
        element.style.left = `${stickyMaterialWidth()}px`;
      } else if (index === col["Material No."]) {
        element.style.left = "0";
      }
    });
  }

  const els = {
    sourceMeta: document.getElementById("sourceMeta"),
    saveStatus: document.getElementById("saveStatus"),
    rowCount: document.getElementById("rowCount"),
    filterArea: document.getElementById("filterArea"),
    search: document.getElementById("search"),
    minRevenue: document.getElementById("minRevenue"),
    minRevenueValue: document.getElementById("minRevenueValue"),
    zevpSlider: document.getElementById("zevpSlider"),
    zevpSliderValue: document.getElementById("zevpSliderValue"),
    zgruSlider: document.getElementById("zgruSlider"),
    zgruSliderValue: document.getElementById("zgruSliderValue"),
    importPrices: document.getElementById("importPrices"),
    importPricesFile: document.getElementById("importPricesFile"),
    pageSize: document.getElementById("pageSize"),
    prevPage: document.getElementById("prevPage"),
    nextPage: document.getElementById("nextPage"),
    pageLabel: document.getElementById("pageLabel"),
    tableHead: document.getElementById("tableHead"),
    tableBody: document.getElementById("tableBody"),
    metrics: document.getElementById("metrics"),
    columnFilterPopup: document.getElementById("columnFilterPopup")
  };

  function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value === "boolean") return value ? 1 : 0;
    let text = String(value ?? "")
      .replace(/[\u20ac%]/g, "")
      .replace(/\s/g, "")
      .trim();
    if (!text) return NaN;
    const commaCount = (text.match(/,/g) || []).length;
    const dotCount = (text.match(/\./g) || []).length;
    if (commaCount && dotCount) {
      text = text.lastIndexOf(",") > text.lastIndexOf(".")
        ? text.replace(/\./g, "").replace(",", ".")
        : text.replace(/,/g, "");
    } else if (commaCount) {
      const parts = text.split(",");
      text = parts.length === 2 && parts[1].length > 0 && parts[1].length <= 2
        ? `${parts[0]}.${parts[1]}`
        : text.replace(/,/g, "");
    }
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  function toNum(value) {
    const parsed = parseNumber(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function strictNum(value) {
    return parseNumber(value);
  }

  function divMinusOne(num, den) {
    return den ? num / den - 1 : 0;
  }

  function finite(value) {
    return Number.isFinite(value) ? value : 0;
  }

  function materialKey(value) {
    if (typeof value === "number" && Number.isFinite(value) && Number.isInteger(value)) return String(value);
    return String(value ?? "").trim();
  }

  function savedFieldName(field) {
    if (field === "ZEVP FY") return "zevpFY";
    if (field === "ZGRU 2027") return "zgru2027";
    return field;
  }

  function formatSaveTime(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleString(undefined, {
      day: "2-digit",
      month: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    });
  }

  function updateSaveStatus(message, kind = "") {
    if (!els.saveStatus) return;
    window.clearTimeout(saveStatusTimer);
    els.saveStatus.textContent = message;
    els.saveStatus.className = `save-status${kind ? ` ${kind}` : ""}`;
    if (kind === "saved") {
      saveStatusTimer = window.setTimeout(() => {
        els.saveStatus.className = "save-status";
      }, 5000);
    }
  }

  function normalizeSavedPayload(raw) {
    if (!raw) return null;
    const source = Array.isArray(raw) ? { rows: raw } : raw;
    if (!Array.isArray(source.rows)) return null;
    // Enforce row count limit to prevent browser tab freeze/crash.
    if (source.rows.length > MAX_SAVED_ROWS) {
      console.warn(`Saved payload exceeds maximum row count (${source.rows.length} > ${MAX_SAVED_ROWS}). Truncating.`);
      source.rows = source.rows.slice(0, MAX_SAVED_ROWS);
    }
    return {
      version: source.version || 2,
      sourceFile: source.sourceFile || payload.sourceFile || "",
      generatedAt: source.generatedAt || "",
      savedAt: source.savedAt || new Date(0).toISOString(),
      reason: source.reason || "",
      rows: source.rows
    };
  }

  function readSavedPayloadFromStorage(storage, key) {
    try {
      return normalizeSavedPayload(JSON.parse(storage.getItem(key) || "null"));
    } catch (error) {
      console.warn("Unable to read saved pricing inputs.", error);
      return null;
    }
  }

  function browserSavedPayloads() {
    const saved = [];
    try {
      saved.push(readSavedPayloadFromStorage(localStorage, savedInputKey));
      saved.push(readSavedPayloadFromStorage(localStorage, savedInputBackupKey));
    } catch (error) {
      console.warn("Local pricing save is not available.", error);
    }
    try {
      saved.push(readSavedPayloadFromStorage(sessionStorage, savedInputKey));
    } catch (error) {
      // Session storage is optional.
    }
    return saved.filter(Boolean);
  }

  function openPricingDb() {
    // IndexedDB is used as a second local browser storage layer. It is not a
    // remote database and it does not require network access.
    if (pricingDbPromise) return pricingDbPromise;
    pricingDbPromise = new Promise((resolve) => {
      if (!window.indexedDB) {
        resolve(null);
        return;
      }
      const request = indexedDB.open(savedInputDbName, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(savedInputDbStore)) {
          db.createObjectStore(savedInputDbStore, { keyPath: "key" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => {
        console.warn("Unable to open IndexedDB pricing save.", request.error);
        resolve(null);
      };
      request.onblocked = () => resolve(null);
    });
    return pricingDbPromise;
  }

  async function savePayloadToIndexedDb(payloadToSave) {
    const db = await openPricingDb();
    if (!db) return false;
    return new Promise((resolve) => {
      const tx = db.transaction(savedInputDbStore, "readwrite");
      tx.objectStore(savedInputDbStore).put({
        key: savedInputKey,
        payload: payloadToSave,
        savedAt: payloadToSave.savedAt
      });
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => {
        console.warn("Unable to write IndexedDB pricing save.", tx.error);
        resolve(false);
      };
    });
  }

  async function readPayloadFromIndexedDb() {
    const db = await openPricingDb();
    if (!db) return null;
    return new Promise((resolve) => {
      const tx = db.transaction(savedInputDbStore, "readonly");
      const request = tx.objectStore(savedInputDbStore).get(savedInputKey);
      request.onsuccess = () => resolve(normalizeSavedPayload(request.result?.payload));
      request.onerror = () => {
        console.warn("Unable to read IndexedDB pricing save.", request.error);
        resolve(null);
      };
    });
  }

  async function latestSavedPayload() {
    const saved = browserSavedPayloads();
    const indexedDbPayload = await readPayloadFromIndexedDb();
    if (indexedDbPayload) saved.push(indexedDbPayload);
    if (!saved.length) return null;
    saved.sort((a, b) => Date.parse(b.savedAt || 0) - Date.parse(a.savedAt || 0));
    return saved[0];
  }

  function savedEntryFromRow(rowIndex) {
    const material = materialKey(rows[rowIndex][col["Material No."]]);
    const entry = { material, updatedAt: new Date().toISOString() };
    for (const field of savedInputFields) {
      entry[savedFieldName(field)] = finite(toNum(rows[rowIndex][col[field]]));
    }
    return entry;
  }

  function markDirtyInput(rowIndex, field, value) {
    if (!savedInputFields.includes(field)) return;
    const material = materialKey(rows[rowIndex][col["Material No."]]);
    if (!material) return;
    const entry = { ...(dirtyInputs.get(material) || savedEntryFromRow(rowIndex)) };
    for (const editableField of savedInputFields) {
      entry[savedFieldName(editableField)] = finite(toNum(rows[rowIndex][col[editableField]]));
    }
    entry[savedFieldName(field)] = finite(toNum(value));
    entry.updatedAt = new Date().toISOString();
    dirtyInputs.set(material, entry);
  }

  function buildSavedPayload(reason = "auto") {
    return {
      version: 3,
      sourceFile: payload.sourceFile || "",
      generatedAt: payload.generatedAt || "",
      savedAt: new Date().toISOString(),
      reason,
      rows: [...dirtyInputs.values()].map((entry) => ({ ...entry }))
    };
  }

  function savePayloadToBrowserStorage(payloadToSave) {
    const json = JSON.stringify(payloadToSave);
    let ok = false;
    try {
      localStorage.setItem(savedInputKey, json);
      localStorage.setItem(savedInputBackupKey, json);
      ok = true;
    } catch (error) {
      console.warn("Unable to write local pricing save.", error);
    }
    try {
      sessionStorage.setItem(savedInputKey, json);
      ok = true;
    } catch (error) {
      // Session storage is optional.
    }
    return ok;
  }

  function downloadSavedPayload(payloadToSave) {
    try {
      const stamp = new Date(payloadToSave.savedAt).toISOString().replace(/[:.]/g, "-");
      const blob = new Blob([JSON.stringify(payloadToSave, null, 2)], { type: "application/json" });
      const href = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.download = `pricing_saved_prices_${stamp}.json`;
      link.href = href;
      document.body.appendChild(link);
      link.click();
      setTimeout(() => {
        URL.revokeObjectURL(href);
        link.remove();
      }, 0);
      return true;
    } catch (error) {
      console.warn("Unable to download pricing backup.", error);
      return false;
    }
  }

  function saveEditableValues(reason = "auto", options = {}) {
    // Manual ZEVP FY / ZGRU 2027 inputs are saved to multiple local browser
    // stores, with an optional JSON download for user-controlled backup.
    window.clearTimeout(saveDebounceTimer);
    if (!dirtyInputs.size) {
      updateSaveStatus("No saved prices", "");
      return false;
    }
    const payloadToSave = buildSavedPayload(reason);
    const rowCount = payloadToSave.rows.length;
    const localOk = savePayloadToBrowserStorage(payloadToSave);
    const downloaded = options.downloadBackup ? downloadSavedPayload(payloadToSave) : false;
    savePayloadToIndexedDb(payloadToSave).then((indexedDbOk) => {
      if (localOk || indexedDbOk || downloaded) {
        updateSaveStatus(
          `Saved ${rowCount} rows ${formatSaveTime(payloadToSave.savedAt)}${downloaded ? " + backup file" : ""}`,
          "saved"
        );
      } else {
        updateSaveStatus("Save failed", "error");
      }
    });
    if (localOk || downloaded) {
      updateSaveStatus(
        `Saved ${rowCount} rows ${formatSaveTime(payloadToSave.savedAt)}${downloaded ? " + backup file" : ""}`,
        "saved"
      );
    } else {
      updateSaveStatus("Saving...", "");
    }
    return localOk || downloaded;
  }

  function scheduleEditableSave(delay = 350) {
    window.clearTimeout(saveDebounceTimer);
    updateSaveStatus("Saving...", "");
    saveDebounceTimer = window.setTimeout(() => saveEditableValues("debounced-edit"), delay);
  }

  function applySavedPayload(saved) {
    if (!saved) return { cells: 0, rows: 0, savedAt: "" };
    const rowByMaterial = new Map(rows.map((row, index) => [materialKey(row[col["Material No."]]), index]));
    let appliedCells = 0;
    let appliedRows = 0;
    dirtyInputs.clear();
    for (const entry of saved.rows) {
      const material = materialKey(entry.material);
      const rowIndex = rowByMaterial.get(material);
      if (rowIndex === undefined) continue;
      const normalized = { material, updatedAt: entry.updatedAt || saved.savedAt || new Date().toISOString() };
      let rowTouched = false;
      for (const field of savedInputFields) {
        const key = savedFieldName(field);
        if (entry[key] === undefined) continue;
        const value = parseNumber(entry[key]);
        if (!Number.isFinite(value)) continue;
        rows[rowIndex][col[field]] = value;
        baseRows[rowIndex][col[field]] = value;
        normalized[key] = value;
        appliedCells += 1;
        rowTouched = true;
      }
      if (rowTouched) {
        dirtyInputs.set(material, normalized);
        appliedRows += 1;
      }
    }
    return { cells: appliedCells, rows: appliedRows, savedAt: saved.savedAt || "" };
  }

  async function applySavedEditableValues() {
    const saved = await latestSavedPayload();
    return applySavedPayload(saved);
  }

  async function importSavedPricesFile(file) {
    if (!file) return;
    try {
      // Enforce file size limit to prevent browser tab freeze/crash.
      if (file.size > MAX_IMPORT_FILE_SIZE) {
        throw new Error(`File size (${(file.size / 1024 / 1024).toFixed(1)} MB) exceeds maximum allowed size (${MAX_IMPORT_FILE_SIZE / 1024 / 1024} MB).`);
      }
      const imported = normalizeSavedPayload(JSON.parse(await file.text()));
      if (!imported) throw new Error("Invalid saved price file.");
      imported.savedAt = new Date().toISOString();
      imported.reason = "import-file";
      imported.sourceFile = payload.sourceFile || imported.sourceFile || "";
      const applied = applySavedPayload(imported);
      recalcAll();
      applyFilters();
      saveEditableValues("import-file");
      updateSaveStatus(`Imported ${applied.rows} rows ${formatSaveTime(imported.savedAt)}`, "saved");
    } catch (error) {
      console.warn("Unable to import saved pricing file.", error);
      updateSaveStatus("Import failed", "error");
    } finally {
      if (els.importPricesFile) els.importPricesFile.value = "";
    }
  }

  function median(values) {
    const sorted = values.slice().sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function growthEstimate(yValues, xValues, newX) {
    const count = yValues.length;
    const meanX = xValues.reduce((total, value) => total + value, 0) / count;
    const logs = yValues.map((value) => Math.log(value));
    const meanLog = logs.reduce((total, value) => total + value, 0) / count;
    const denominator = xValues.reduce((total, value) => total + (value - meanX) ** 2, 0);
    if (!denominator) return median(yValues);
    const slope = xValues.reduce((total, value, index) => total + (value - meanX) * (logs[index] - meanLog), 0) / denominator;
    const intercept = meanLog - slope * meanX;
    return Math.exp(intercept + slope * newX);
  }

  function calcChannelTrend(qtyRaw, priceHistory, newPrice, launchFY, channelName) {
    if (!Number.isFinite(newPrice)) return null;
    const histYears = linearForecast.histYears || [2023, 2024, 2025];
    const valid = [];
    for (let index = 0; index < histYears.length; index += 1) {
      const qty = Number(qtyRaw?.[index]);
      const price = Number(priceHistory?.[index]);
      if (
        Number.isFinite(qty) &&
        Number.isFinite(price) &&
        qty !== 0 &&
        qty > -1 &&
        histYears[index] !== launchFY
      ) {
        valid.push({ qty, price });
      }
    }
    if (!valid.length) return null;
    const multipliers = valid.map(({ qty }) => 1 + Math.min(3.5, Math.max(-0.9, qty)));
    const validPrices = valid.map(({ price }) => price);
    const medianMultiplier = median(multipliers);
    let growthMultiplier = medianMultiplier;
    if (valid.length >= 2 && Math.min(...validPrices) !== Math.max(...validPrices)) {
      growthMultiplier = growthEstimate(multipliers, validPrices, newPrice);
    }
    const shrunkMultiplier = Math.exp(0.65 * Math.log(growthMultiplier) + 0.35 * Math.log(medianMultiplier));
    const rawForecast = shrunkMultiplier - 1;
    let scenarioForecast = rawForecast;
    if (String(channelName || "").includes("Amazon e-commerce (0025)")) {
      scenarioForecast = (1 + rawForecast) * 1.2 - 1;
    } else if (String(channelName || "").includes("DIY Store (0040)")) {
      scenarioForecast = (1 + rawForecast) * 0.8 - 1;
    }
    const adjustedForecast = (1 + scenarioForecast) * 1.02 * 1.03 - 1;
    return Math.min(2.5, Math.max(-0.8, adjustedForecast));
  }

  function recalcLinearForecast(rowIndex) {
    // Recalculate the linear forecast for the edited material using extracted
    // historical channel and price data from the workbook.
    const row = rows[rowIndex];
    const model = linearForecast.models?.[materialKey(row[col["Material No."]])];
    if (!model) return toNum(row[col["Linear Forecast FY27"]]);
    const channelNames = linearForecast.channels || [];
    let trendTotal = 0;
    for (let index = 0; index < channelNames.length; index += 1) {
      const trend = calcChannelTrend(
        model.qty?.[index],
        model.priceHistory,
        toNum(row[col["RET Diff"]]),
        model.launchFY,
        channelNames[index]
      );
      if (Number.isFinite(trend)) trendTotal += trend;
    }
    return finite(toNum(model.baseTotal) + trendTotal);
  }

  function effectivePacemakerForecast(row) {
    // Business rule from the workbook: if Pacemaker volume is zero and the item
    // is not discontinued, use Linear Forecast FY27 as the effective quantity.
    const pacemakerForecast = toNum(row[col[" PaceMaker Forecast FY27"]]);
    const productStatus = String(row[col["Product Status"]] || "").trim().toLowerCase();
    const canUseLinearFallback = pacemakerForecast === 0 && !productStatus.includes("discontinued");
    return canUseLinearFallback ? toNum(row[col["Linear Forecast FY27"]]) : pacemakerForecast;
  }

  function recalcRow(rowIndex) {
    // Recalculate all dependent row-level pricing outputs after manual price
    // edits or scenario-slider changes.
    const row = rows[rowIndex];
    const zevpFY = toNum(row[col["ZEVP FY"]]);
    const zgru2027 = toNum(row[col["ZGRU 2027"]]);
    const zbruFY = constants.u10 ? zevpFY / constants.u10 : 0;
    row[col["ZBRU FY"]] = finite(zbruFY);
    row[col["ZGRU FY"]] = finite(zgru2027);
    row[col["ZNEK FY"]] = finite(toNum(row[col["ZBRU FY"]]) * (1 + toNum(row[col["ZGRU FY"]])));
    row[col["RET Diff"]] = finite(divMinusOne(zevpFY, toNum(row[col["ZEVP 2026"]])));
    row[col["Trade Diff ZNEK"]] = finite(divMinusOne(toNum(row[col["ZNEK FY"]]), toNum(row[col["ZNEK 2026"]])));
    row[col["Linear Forecast FY27"]] = recalcLinearForecast(rowIndex);

    const streetAverage = (toNum(row[col["Oase Cheapest"]]) + toNum(row[col["Oase Highest"]])) / 2;
    row[col["Street Price FY27"]] = streetAverage === 0 ? finite(zevpFY * 0.82) : streetAverage;
    row[col["Street Price Diff SY"]] = finite(divMinusOne(toNum(row[col["Street Price FY27"]]), zevpFY));

    const nsv26Strict = strictNum(row[col["NSV 26 YTD"]]);
    const theoretical = nsv26Strict + divMinusOne(zevpFY, toNum(row[col["ZEVP 2025"]]));
    row[col["Theoretical NSV FY"]] = Number.isFinite(theoretical) ? theoretical : toNum(row[col["NSV 25 Full Year"]]);

    const conditionMix = toNum(row[col["Condition Mix"]]);
    const baseNSV = projectedBase[rowIndex] === 0 ? strictNum(row[col["NSV 26 YTD"]]) : projectedBase[rowIndex];
    let projected = baseNSV * (1 + toNum(row[col["RET Diff"]]) + conditionMix);
    const theoreticalNSV = toNum(row[col["Theoretical NSV FY"]]);
    if (!Number.isFinite(projected)) {
      row[col["Projected NSV 27"]] = 0;
    } else if (projected >= theoreticalNSV * 1.25 || projected <= theoreticalNSV * 0.75) {
      row[col["Projected NSV 27"]] = theoreticalNSV;
    } else {
      row[col["Projected NSV 27"]] = projected;
    }
    row[col["Forecasted NSV Change FY27"]] = finite(divMinusOne(toNum(row[col["Projected NSV 27"]]), theoreticalNSV));

    const projectedNSV = toNum(row[col["Projected NSV 27"]]);
    row[col["GP FY"]] = projectedNSV
      ? finite((projectedNSV - (toNum(row[col["FIFO Stack Cost"]]) + projectedNSV * constants.cogsRateSum)) / projectedNSV)
      : 0;
    row[col["Margin Erosion"]] = finite(toNum(row[col["GP FY"]]) - toNum(row[col["GP Margin SY"]]));
    const pacemakerForecast = effectivePacemakerForecast(row);
    row[col["Revenue @Forecast Old Price"]] = finite(toNum(row[col["Theoretical NSV FY"]]) * pacemakerForecast);
    row[col["Revenue @Forecast New Price Pacemaker"]] = finite(pacemakerForecast * toNum(row[col["Projected NSV 27"]]));
    row[col["Revenue @Forecast New Price Linear"]] = finite(toNum(row[col["Projected NSV 27"]]) * toNum(row[col["Linear Forecast FY27"]]));
    row[col["Cash Margin FY26 FY"]] = finite(
      toNum(row[col["I.O. QTY 26 Full Year"]]) * (toNum(row[col["NSV 26 YTD"]]) * toNum(row[col["GP Margin SY"]]))
    );
    row[col["Cash Margin FY27FY Pacemaker"]] = finite(
      toNum(row[col["Projected NSV 27"]]) * pacemakerForecast * toNum(row[col["GP FY"]])
    );
  }

  function recalcAll() {
    for (let i = 0; i < rows.length; i += 1) recalcRow(i);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function formatNumber(value, header) {
    if (header === "Material No.") return String(value ?? "");
    const n = toNum(value);
    if (currencyHeaders.has(header)) {
      return n.toLocaleString(undefined, {
        style: "currency",
        currency: "EUR",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }
    if (percentHeaders.has(header)) return `${(n * 100).toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 })}%`;
    if (Math.abs(n) >= 1000000) return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    if (Math.abs(n) >= 1000) return n.toLocaleString(undefined, { maximumFractionDigits: 1 });
    if (Math.abs(n) >= 10) return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
    return n.toLocaleString(undefined, { maximumFractionDigits: 4 });
  }

  function formatCompact(value) {
    const n = toNum(value);
    return n.toLocaleString(undefined, { notation: "compact", maximumFractionDigits: 2 });
  }

  function formatEuroCompact(value) {
    return `€${formatCompact(value)}`;
  }

  function formatPct(value) {
    return `${(toNum(value) * 100).toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 })}%`;
  }

  function formatPctSigned(value) {
    const n = toNum(value) * 100;
    const sign = n > 0 ? "+" : "";
    return `${sign}${n.toLocaleString(undefined, { maximumFractionDigits: 1, minimumFractionDigits: 1 })}%`;
  }

  function isNumericColumn(index) {
    if (headers[index] === "Material No.") return false;
    for (let i = 0; i < Math.min(rows.length, 25); i += 1) {
      const value = rows[i][index];
      if (typeof value === "number") return true;
      if (typeof value === "string" && value.trim() && Number.isFinite(Number(value))) return true;
    }
    return false;
  }

  function filterId(field) {
    return field.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase();
  }

  function selectedValues(field) {
    return state.filters[field] || [];
  }

  function updateFilterSummary(field, labelEl, totalCount = 0) {
    if (selectedValues(field).includes("__NO_FILTER_MATCH__")) {
      labelEl.textContent = "0 selected";
      return;
    }
    const count = selectedValues(field).length;
    labelEl.textContent = count && count !== totalCount ? `${count} selected` : "All";
  }

  function columnFilterKey(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value);
  }

  function columnFilterLabel(value, header) {
    if (value === null || value === undefined || value === "") return "(blank)";
    return isNumericColumn(col[header]) ? formatNumber(value, header) : String(value);
  }

  function selectedColumnValues(colIndex) {
    return state.columnFilters[colIndex] || [];
  }

  function compareColumnValues(aRowIndex, bRowIndex, colIndex, direction) {
    const aRaw = rows[aRowIndex][colIndex];
    const bRaw = rows[bRowIndex][colIndex];
    const multiplier = direction === "desc" ? -1 : 1;
    const aBlank = aRaw === null || aRaw === undefined || aRaw === "";
    const bBlank = bRaw === null || bRaw === undefined || bRaw === "";
    if (aBlank && bBlank) return aRowIndex - bRowIndex;
    if (aBlank) return 1;
    if (bBlank) return -1;
    let result;
    if (isNumericColumn(colIndex)) {
      result = toNum(aRaw) - toNum(bRaw);
    } else {
      result = String(aRaw).localeCompare(String(bRaw), undefined, { numeric: true, sensitivity: "base" });
    }
    if (result === 0) return aRowIndex - bRowIndex;
    return result * multiplier;
  }

  function sortFilteredRows() {
    if (state.tableSort.column === null || !state.tableSort.direction) return;
    const colIndex = Number(state.tableSort.column);
    state.filtered.sort((a, b) => compareColumnValues(a, b, colIndex, state.tableSort.direction));
  }

  function rowMatchesFilters(rowIndex, options = {}) {
    const row = rows[rowIndex];
    for (const field of filterFields) {
      const selected = selectedValues(field);
      if (selected.length && !selected.includes(String(row[col[field]]))) return false;
    }
    if (state.minRevenue && toNum(row[col["Revenue @Forecast New Price Pacemaker"]]) < state.minRevenue) return false;
    if (state.search.trim()) {
      const haystack = [
        row[col["Material No."]],
        row[col["Material Descr."]],
        row[col["Market Segment"]],
        row[col["Product Family"]],
        row[col["Product Group"]]
      ].join(" ").toLowerCase();
      if (!haystack.includes(state.search.trim().toLowerCase())) return false;
    }
    for (const [columnIndex, selected] of Object.entries(state.columnFilters)) {
      const index = Number(columnIndex);
      if (options.skipColumnIndex === index) continue;
      if (selected.length && !selected.includes(columnFilterKey(row[index]))) return false;
    }
    return true;
  }

  function getColumnFilterOptions(colIndex) {
    const header = headers[colIndex];
    const map = new Map();
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      if (!rowMatchesFilters(rowIndex, { skipColumnIndex: colIndex })) continue;
      const value = rows[rowIndex][colIndex];
      const key = columnFilterKey(value);
      if (!map.has(key)) map.set(key, columnFilterLabel(value, header));
    }
    return [...map.entries()]
      .map(([key, label]) => ({ key, label }))
      .sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true }));
  }

  function updateColumnFilterButtons() {
    document.querySelectorAll(".th-filter").forEach((button) => {
      const colIndex = Number(button.dataset.col);
      const isSorted = state.tableSort.column === colIndex && Boolean(state.tableSort.direction);
      button.classList.toggle("active", selectedColumnValues(colIndex).length > 0 || isSorted);
    });
  }

  function closeColumnFilter() {
    els.columnFilterPopup.hidden = true;
    els.columnFilterPopup.innerHTML = "";
  }

  function positionColumnFilter(anchor) {
    const rect = anchor.getBoundingClientRect();
    const popup = els.columnFilterPopup;
    const width = Math.min(300, window.innerWidth - 24);
    let left = rect.right - width;
    left = Math.max(12, Math.min(left, window.innerWidth - width - 12));
    let top = rect.bottom + 6;
    const maxHeight = Math.min(430, window.innerHeight - 24);
    if (top + maxHeight > window.innerHeight) top = Math.max(12, rect.top - maxHeight - 6);
    popup.style.left = `${left}px`;
    popup.style.top = `${top}px`;
  }

  function renderColumnFilterOptions(colIndex, query = "") {
    const popup = els.columnFilterPopup;
    const list = popup.querySelector(".column-filter-list");
    const allOptions = getColumnFilterOptions(colIndex);
    const currentFilter = selectedColumnValues(colIndex);
    const selected = new Set(currentFilter.length ? currentFilter : allOptions.map((item) => item.key));
    const filteredOptions = allOptions.filter((item) => item.label.toLowerCase().includes(query.trim().toLowerCase()));
    list.innerHTML = filteredOptions.map((item, index) => {
      const optionId = `column-filter-${colIndex}-${index}`;
      const checked = selected.has(item.key) ? " checked" : "";
      return `
        <label class="column-filter-option" for="${optionId}" title="${escapeHtml(item.label)}">
          <input id="${optionId}" type="checkbox" value="${escapeHtml(item.key)}"${checked}>
          <span>${escapeHtml(item.label)}</span>
        </label>
      `;
    }).join("");
    list.querySelectorAll("input[type='checkbox']").forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        selected.delete("__NO_COLUMN_MATCH__");
        if (checkbox.checked) selected.add(checkbox.value);
        else selected.delete(checkbox.value);
        if (selected.size === 0) state.columnFilters[colIndex] = ["__NO_COLUMN_MATCH__"];
        else if (selected.size === allOptions.length) delete state.columnFilters[colIndex];
        else state.columnFilters[colIndex] = [...selected];
        state.page = 0;
        applyFilters();
      });
    });
  }

  function applyColumnSearchFilter(colIndex, query) {
    const search = query.trim().toLowerCase();
    if (!search) return false;
    const matches = getColumnFilterOptions(colIndex).filter((item) => {
      const label = item.label.toLowerCase();
      const key = item.key.toLowerCase();
      return label.includes(search) || key.includes(search);
    });
    state.columnFilters[colIndex] = matches.length ? matches.map((item) => item.key) : ["__NO_COLUMN_MATCH__"];
    state.page = 0;
    applyFilters();
    return true;
  }

  function openColumnFilter(colIndex, anchor) {
    const header = headers[colIndex];
    const options = getColumnFilterOptions(colIndex);
    const selected = selectedColumnValues(colIndex);
    const selectedCount = selected.includes("__NO_COLUMN_MATCH__") ? 0 : (selected.length || options.length);
    const sortAscActive = state.tableSort.column === colIndex && state.tableSort.direction === "asc" ? " active" : "";
    const sortDescActive = state.tableSort.column === colIndex && state.tableSort.direction === "desc" ? " active" : "";
    const ascLabel = isNumericColumn(colIndex) ? "Smallest to largest" : "Sort A to Z";
    const descLabel = isNumericColumn(colIndex) ? "Largest to smallest" : "Sort Z to A";
    els.columnFilterPopup.innerHTML = `
      <div class="column-filter-head">
        <div class="column-filter-title">${escapeHtml(header)} (${selectedCount}/${options.length})</div>
        <input class="column-filter-search" type="search" placeholder="Search values">
        <div class="column-filter-actions">
          <button type="button" data-action="sort-asc" class="${sortAscActive.trim()}">${escapeHtml(ascLabel)}</button>
          <button type="button" data-action="sort-desc" class="${sortDescActive.trim()}">${escapeHtml(descLabel)}</button>
          <button type="button" data-action="clear-sort">Clear sort</button>
          <button type="button" data-action="apply-search">Apply search</button>
          <button type="button" data-action="all">Select all</button>
          <button type="button" data-action="none">Select none</button>
        </div>
      </div>
      <div class="column-filter-list"></div>
    `;
    els.columnFilterPopup.hidden = false;
    positionColumnFilter(anchor);
    renderColumnFilterOptions(colIndex);

    const searchInput = els.columnFilterPopup.querySelector(".column-filter-search");
    searchInput.addEventListener("input", () => renderColumnFilterOptions(colIndex, searchInput.value));
    searchInput.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      event.stopPropagation();
      if (applyColumnSearchFilter(colIndex, searchInput.value)) closeColumnFilter();
    });
    const reopenAfterSort = () => {
      const nextAnchor = els.tableHead.querySelector(`.th-filter[data-col="${colIndex}"]`) || anchor;
      openColumnFilter(colIndex, nextAnchor);
    };
    els.columnFilterPopup.querySelector("[data-action='sort-asc']").addEventListener("click", () => {
      state.tableSort = { column: colIndex, direction: "asc" };
      state.page = 0;
      applyFilters();
      reopenAfterSort();
    });
    els.columnFilterPopup.querySelector("[data-action='sort-desc']").addEventListener("click", () => {
      state.tableSort = { column: colIndex, direction: "desc" };
      state.page = 0;
      applyFilters();
      reopenAfterSort();
    });
    els.columnFilterPopup.querySelector("[data-action='clear-sort']").addEventListener("click", () => {
      if (state.tableSort.column === colIndex) {
        state.tableSort = { column: null, direction: null };
        state.page = 0;
        applyFilters();
      }
      reopenAfterSort();
    });
    els.columnFilterPopup.querySelector("[data-action='apply-search']").addEventListener("click", () => {
      if (applyColumnSearchFilter(colIndex, searchInput.value)) closeColumnFilter();
    });
    els.columnFilterPopup.querySelector("[data-action='all']").addEventListener("click", () => {
      delete state.columnFilters[colIndex];
      state.page = 0;
      applyFilters();
      openColumnFilter(colIndex, anchor);
    });
    els.columnFilterPopup.querySelector("[data-action='none']").addEventListener("click", () => {
      state.columnFilters[colIndex] = ["__NO_COLUMN_MATCH__"];
      state.page = 0;
      applyFilters();
      openColumnFilter(colIndex, anchor);
    });
    searchInput.focus();
  }

  function populateFilters() {
    els.filterArea.innerHTML = "";
    for (const field of filterFields) {
      const values = [...new Set(rows.map((row) => row[col[field]]).filter((value) => value !== null && value !== "" && value !== undefined))]
        .map(String)
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      const wrapper = document.createElement("div");
      wrapper.className = "field";
      const safeId = filterId(field);
      wrapper.innerHTML = `<label for="filter-${safeId}">${escapeHtml(field)}</label>`;
      const details = document.createElement("details");
      details.className = "multi-filter";
      details.id = `filter-${safeId}`;
      const searchableFilter = field === "Product Family" || field === "Product Group";
      const searchMarkup = searchableFilter ? `
        <div class="multi-filter-search-wrap">
          <input class="multi-filter-search" type="search" placeholder="Search ${escapeHtml(field.toLowerCase())}" aria-label="Search ${escapeHtml(field)}">
        </div>
      ` : "";
      const options = values.map((value, index) => {
        const optionId = `filter-${safeId}-${index}`;
        return `
          <label class="multi-option" for="${optionId}" title="${escapeHtml(value)}">
            <input id="${optionId}" type="checkbox" value="${escapeHtml(value)}">
            <span>${escapeHtml(value)}</span>
          </label>
        `;
      }).join("");
      details.innerHTML = `
        <summary><span class="summary-text">All</span></summary>
        <div class="multi-filter-actions">
          <button type="button" data-action="all">Select all</button>
          <button type="button" data-action="none">Deselect all</button>
        </div>
        ${searchMarkup}
        <div class="multi-options">${options}</div>
      `;
      const summaryLabel = details.querySelector(".summary-text");
      const checkboxInputs = [...details.querySelectorAll("input[type='checkbox']")];
      const searchInput = details.querySelector(".multi-filter-search");
      let optionLabels = [];
      const resetOptionSearch = () => {
        if (!searchInput) return;
        searchInput.value = "";
        optionLabels.forEach((label) => {
          label.hidden = false;
        });
      };
      if (searchInput) {
        optionLabels = [...details.querySelectorAll(".multi-option")];
        const applyOptionSearch = () => {
          const query = searchInput.value.trim().toLowerCase();
          const matchedValues = [];
          optionLabels.forEach((label) => {
            const matches = !query || label.textContent.toLowerCase().includes(query);
            label.hidden = !matches;
            if (matches && query) matchedValues.push(label.querySelector("input").value);
          });
          if (searchableFilter) {
            checkboxInputs.forEach((checkbox) => {
              checkbox.checked = !query || matchedValues.includes(checkbox.value);
            });
            if (!query) delete state.filters[field];
            else if (matchedValues.length === 0) state.filters[field] = ["__NO_FILTER_MATCH__"];
            else state.filters[field] = matchedValues;
            updateFilterSummary(field, summaryLabel, checkboxInputs.length);
            state.page = 0;
            applyFilters();
          }
        };
        searchInput.addEventListener("input", applyOptionSearch);
        searchInput.addEventListener("click", (event) => event.stopPropagation());
        searchInput.addEventListener("keydown", (event) => event.stopPropagation());
      }
      const syncFilterFromChecks = () => {
        const checkedValues = checkboxInputs.filter((item) => item.checked).map((item) => item.value);
        if (checkedValues.length === 0) state.filters[field] = ["__NO_FILTER_MATCH__"];
        else if (checkedValues.length === checkboxInputs.length) delete state.filters[field];
        else state.filters[field] = checkedValues;
        updateFilterSummary(field, summaryLabel, checkboxInputs.length);
        state.page = 0;
        applyFilters();
      };
      checkboxInputs.forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          syncFilterFromChecks();
        });
      });
      details.querySelector("[data-action='all']").addEventListener("click", () => {
        resetOptionSearch();
        checkboxInputs.forEach((item) => {
          item.checked = true;
        });
        delete state.filters[field];
        updateFilterSummary(field, summaryLabel, checkboxInputs.length);
        state.page = 0;
        applyFilters();
      });
      details.querySelector("[data-action='none']").addEventListener("click", () => {
        resetOptionSearch();
        checkboxInputs.forEach((item) => {
          item.checked = false;
        });
        state.filters[field] = ["__NO_FILTER_MATCH__"];
        updateFilterSummary(field, summaryLabel, checkboxInputs.length);
        state.page = 0;
        applyFilters();
      });
      details.addEventListener("toggle", () => {
        if (!details.open) return;
        els.filterArea.querySelectorAll("details.multi-filter[open]").forEach((item) => {
          if (item !== details) item.open = false;
        });
        if (searchInput) window.setTimeout(() => searchInput.focus(), 0);
      });
      updateFilterSummary(field, summaryLabel, checkboxInputs.length);
      wrapper.appendChild(details);
      els.filterArea.appendChild(wrapper);
    }
  }

  function setupRevenueSlider() {
    const values = rows.map((row) => toNum(row[col["Revenue @Forecast New Price Pacemaker"]])).filter((value) => value > 0).sort((a, b) => a - b);
    const p95 = values.length ? values[Math.floor(values.length * 0.95)] : 100000;
    els.minRevenue.max = Math.max(1, Math.ceil(p95));
    els.minRevenue.step = Math.max(1, Math.round(Number(els.minRevenue.max) / 200));
    els.minRevenue.value = 0;
    els.minRevenueValue.textContent = "0";
  }

  function applyFilters() {
    state.filtered = [];
    for (let i = 0; i < rows.length; i += 1) {
      if (rowMatchesFilters(i)) state.filtered.push(i);
    }
    sortFilteredRows();
    updateColumnFilterButtons();
    renderDashboard();
  }

  function summary(indices) {
    const sum = (field) => indices.reduce((total, rowIndex) => total + toNum(rows[rowIndex][col[field]]), 0);
    const avg = (field) => indices.length ? sum(field) / indices.length : 0;
    const visibleQtyTotal = sum("Sales Qty YTD");
    const visibleShare = (rowIndex) => {
      if (visibleQtyTotal) return toNum(rows[rowIndex][col["Sales Qty YTD"]]) / visibleQtyTotal;
      return 0;
    };
    const weighted = (field, positiveOnly = false) => indices.reduce((total, rowIndex) => {
      const row = rows[rowIndex];
      const value = toNum(row[col[field]]);
      if (positiveOnly && value <= 0) return total;
      return total + value * visibleShare(rowIndex);
    }, 0);
    const revenueFY26 = sum("Revenue FY26 Full Year");
    const revenueOld = sum("Revenue @Forecast Old Price");
    const revenueNew = sum("Revenue @Forecast New Price Pacemaker");
    const revenueLinear = sum("Revenue @Forecast New Price Linear");
    const cash26 = sum("Cash Margin FY26 FY");
    const cash27 = sum("Cash Margin FY27FY Pacemaker");
    return {
      rows: indices.length,
      revenueFY26,
      revenueOld,
      revenueNew,
      revenueLinear,
      cash26,
      cash27,
      revenueChange: revenueOld ? revenueNew / revenueOld - 1 : 0,
      revenuePacemakerVsFY26: revenueFY26 ? revenueNew / revenueFY26 - 1 : 0,
      cashChange: cash26 ? cash27 / cash26 - 1 : 0,
      avgZevpFY: avg("ZEVP FY"),
      avgZgruFY: avg("ZGRU FY"),
      weightedNSV: weighted("Forecasted NSV Change FY27", true),
      weightedGP: weighted("GP FY", true),
      marginErosion: weighted("Margin Erosion", false),
      fifoStackCostChange: weighted("FIFO Stack Cost Change", true),
      volumeChangeFY27: weighted("Vol.Change FY27 pace.ai", false)
    };
  }

  function renderMetrics(data) {
    const metricData = [
      { label: "Filtered rows", value: data.rows.toLocaleString(), sub: "PricingEngine records" },
      {
        label: "Revenue FY27 pacemaker",
        value: formatCompact(data.revenueNew),
        sub: "Current card",
        badge: formatPctSigned(data.revenuePacemakerVsFY26),
        badgeClass: data.revenuePacemakerVsFY26 < 0 ? " negative" : "",
        badgeTitle: `YoY vs Revenue FY26 Full Year: ${formatPctSigned(data.revenuePacemakerVsFY26)}`
      },
      { label: "Revenue FY26 Full Year", value: formatCompact(data.revenueFY26), sub: "AE10 subtotal" },
      { label: "Cash Margin FY26 FY", value: formatCompact(data.cash26), sub: "BK10 subtotal" },
      { label: "Cash Margin FY27 Pacemaker", value: formatCompact(data.cash27), sub: "BL10 subtotal" },
      { label: "GP FY", value: formatPct(data.weightedGP), sub: "AY10 weighted" },
      { label: "Margin Erosion", value: formatPct(data.marginErosion), sub: "BA10 weighted" },
      { label: "FIFO Stack Cost Change", value: formatPct(data.fifoStackCostChange), sub: "N10 weighted" },
      { label: "Volume Change", value: formatPct(data.volumeChangeFY27), sub: "AM10 weighted" }
    ];
    els.metrics.innerHTML = metricData.map((item) => `
      <div class="metric${item.badge ? " has-badge" : ""}" title="${escapeHtml(`${item.label}: ${item.value} | ${item.sub}${item.badgeTitle ? ` | ${item.badgeTitle}` : ""}`)}">
        ${item.badge ? `<div class="metric-badge${item.badgeClass || ""}" title="${escapeHtml(item.badgeTitle || item.badge)}">${escapeHtml(item.badge)}</div>` : ""}
        <div class="label">${escapeHtml(item.label)}</div>
        <div class="value">${escapeHtml(item.value)}</div>
        <div class="sub">${escapeHtml(item.sub)}</div>
      </div>
    `).join("");
  }

  function aggregateTop(field, valueField, indices, limit = 10) {
    const map = new Map();
    for (const rowIndex of indices) {
      const row = rows[rowIndex];
      const key = String(row[col[field]] || "Unassigned");
      map.set(key, (map.get(key) || 0) + toNum(row[col[valueField]]));
    }
    return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit);
  }

  function forecastVolumeByFamily(indices) {
    const map = new Map();
    for (const rowIndex of indices) {
      const row = rows[rowIndex];
      const family = String(row[col["Product Family"]] || "Unassigned");
      const entry = map.get(family) || { family, pacemaker: 0, linear: 0 };
      entry.pacemaker += toNum(row[col[" PaceMaker Forecast FY27"]]);
      entry.linear += toNum(row[col["Linear Forecast FY27"]]);
      map.set(family, entry);
    }
    return [...map.values()]
      .map((item) => ({ ...item, delta: item.linear - item.pacemaker }))
      .filter((item) => item.pacemaker || item.linear)
      .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta))
      .reverse();
  }

  function segmentRevenueMaterialStats(indices) {
    const map = new Map();
    for (const rowIndex of indices) {
      const row = rows[rowIndex];
      const segment = String(row[col["Segment"]] || "Unassigned").trim() || "Unassigned";
      const entry = map.get(segment) || { segment, revenue: 0, materials: new Set() };
      entry.revenue += toNum(row[col["Revenue @Forecast New Price Pacemaker"]]);
      entry.materials.add(materialKey(row[col["Material No."]]));
      map.set(segment, entry);
    }
    const preferred = ["A", "B", "C"];
    return [...map.values()]
      .map((item) => ({ segment: item.segment, revenue: item.revenue, count: item.materials.size }))
      .sort((a, b) => {
        const ai = preferred.indexOf(a.segment);
        const bi = preferred.indexOf(b.segment);
        if (ai !== -1 || bi !== -1) return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
        return a.segment.localeCompare(b.segment, undefined, { numeric: true });
      });
  }

  function renderCharts(data) {
    const layoutBase = {
      margin: { l: 48, r: 22, t: 44, b: 46 },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff",
      font: { family: "Segoe UI, Arial, sans-serif", size: 12, color: "#202124" },
      hoverlabel: { bgcolor: "#ffffff", bordercolor: "#d9dee7", font: { color: "#202124" } }
    };
    const config = { responsive: true, displayModeBar: false };
    const paddedBarRange = (values) => {
      const nums = values.map(toNum);
      const minValue = Math.min(...nums, 0);
      const maxValue = Math.max(...nums, 0);
      if (minValue < 0 && maxValue > 0) return [minValue * 1.16, maxValue * 1.16];
      if (minValue < 0) return [minValue * 1.16, 0];
      return maxValue ? [0, maxValue * 1.16] : [0, 1];
    };
    const revenueValues = [data.revenueFY26, data.revenueOld, data.revenueNew, data.revenueLinear];
    const cashValues = [data.cash26, data.cash27];

    Plotly.react("revenueChart", [{
      type: "bar",
      x: ["FY26", "Old FY27", "Pacemaker FY27", "Linear FY27"],
      y: revenueValues,
      width: 0.42,
      text: revenueValues.map(formatEuroCompact),
      textposition: "outside",
      textfont: { size: 12, color: "#202124", weight: 700 },
      cliponaxis: false,
      marker: {
        color: ["#94a3b8", "#60a5fa", "#2dd4bf", "#fb923c"],
        line: { color: "#ffffff", width: 1.5 }
      },
      hovertemplate: "%{x}<br>€%{y:,.0f}<extra></extra>"
    }], {
      ...layoutBase,
      title: { text: "Revenue Scenario", font: { size: 15 } },
      bargap: 0.58,
      yaxis: { tickformat: ",.2s", gridcolor: "#edf0f5", zeroline: false, range: paddedBarRange(revenueValues) },
      xaxis: { tickangle: 0, showline: false }
    }, config);

    Plotly.react("cashChart", [{
      type: "bar",
      x: ["FY26", "FY27 pacemaker"],
      y: cashValues,
      width: 0.34,
      text: cashValues.map(formatEuroCompact),
      textposition: "outside",
      textfont: { size: 12, color: "#202124", weight: 700 },
      cliponaxis: false,
      marker: {
        color: ["#94a3b8", "#2dd4bf"],
        line: { color: "#ffffff", width: 1.5 }
      },
      hovertemplate: "%{x}<br>€%{y:,.0f}<extra></extra>"
    }], {
      ...layoutBase,
      title: { text: "Cash Margin", font: { size: 15 } },
      bargap: 0.68,
      yaxis: { tickformat: ",.2s", gridcolor: "#edf0f5", zeroline: false, range: paddedBarRange(cashValues) },
      xaxis: { showline: false }
    }, config);

    const topGroups = aggregateTop("Product Group", "Revenue @Forecast New Price Pacemaker", state.filtered, 10).reverse();
    Plotly.react("groupChart", [{
      type: "bar",
      orientation: "h",
      x: topGroups.map(([, value]) => value),
      y: topGroups.map(([name]) => name),
      marker: { color: "#2563eb" },
      hovertemplate: "%{y}<br>%{x:,.0f}<extra></extra>"
    }], {
      ...layoutBase,
      title: { text: "Top Product Groups", font: { size: 15 } },
      xaxis: { tickformat: ",.2s", gridcolor: "#edf0f5" },
      yaxis: { automargin: true }
    }, config);

    const rrpFields = [
      { field: "ZEVP 2025", name: "RRP 2025", color: "#64748b" },
      { field: "ZEVP 2026", name: "RRP 2026", color: "#2563eb" },
      { field: "ZEVP FY", name: "RRP FY27", color: "#0f766e" }
    ];
    const familyPalette = ["#2563eb", "#0f766e", "#f97316", "#7c3aed", "#db2777", "#0891b2", "#65a30d", "#b45309", "#475569", "#dc2626"];
    const colorForFamily = (family) => {
      let hash = 0;
      for (const char of String(family || "Unassigned")) hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
      return familyPalette[Math.abs(hash) % familyPalette.length];
    };
    const rrpSortValue = (rowIndex) => {
      const fy27 = toNum(rows[rowIndex][col["ZEVP FY"]]);
      if (fy27 > 0) return fy27;
      const fy26 = toNum(rows[rowIndex][col["ZEVP 2026"]]);
      if (fy26 > 0) return fy26;
      return toNum(rows[rowIndex][col["ZEVP 2025"]]);
    };
    const familyOrGroupFocused =
      selectedValues("Product Family").length > 0 ||
      selectedValues("Product Group").length > 0 ||
      Boolean(state.columnFilters[col["Product Family"]]?.length) ||
      Boolean(state.columnFilters[col["Product Group"]]?.length);
    const productLimit = familyOrGroupFocused ? 120 : 50;
    const rrpCandidates = state.filtered
      .filter((rowIndex) => rrpFields.some((item) => toNum(rows[rowIndex][col[item.field]]) > 0))
      .sort((a, b) => {
        const priceCompare = rrpSortValue(a) - rrpSortValue(b);
        if (priceCompare !== 0) return priceCompare;
        const familyCompare = String(rows[a][col["Product Family"]] || "").localeCompare(String(rows[b][col["Product Family"]] || ""), undefined, { numeric: true, sensitivity: "base" });
        if (familyCompare !== 0) return familyCompare;
        return String(rows[a][col["Material Descr."]] || "").localeCompare(String(rows[b][col["Material Descr."]] || ""), undefined, { numeric: true, sensitivity: "base" });
      });
    const rrpProducts = rrpCandidates.slice(0, productLimit);
    const productLabel = (rowIndex) => {
      const material = rows[rowIndex][col["Material No."]];
      const description = String(rows[rowIndex][col["Material Descr."]] || "Product");
      return `${material} | ${description}`;
    };
    const productLabels = rrpProducts.map(productLabel);
    const tickText = productLabels.map((name) => name.length > 28 ? `${name.slice(0, 25)}...` : name);
    const productHoverRow = (rowIndex) => {
      const row = rows[rowIndex];
      return [
        row[col["Product Family"]] || "Unassigned",
        row[col["Material No."]] || "",
        row[col["Material Descr."]] || ""
      ];
    };
    const productHover = rrpProducts.map(productHoverRow);
    const visibleFamilies = [...new Set(rrpCandidates.map((rowIndex) => rows[rowIndex][col["Product Family"]] || "Unassigned"))];
    let rrpTraces;
    if (familyOrGroupFocused && visibleFamilies.length > 1) {
      const byFamily = new Map();
      for (const rowIndex of rrpProducts) {
        const family = rows[rowIndex][col["Product Family"]] || "Unassigned";
        if (!byFamily.has(family)) byFamily.set(family, []);
        byFamily.get(family).push(rowIndex);
      }
      rrpTraces = [...byFamily.entries()].map(([family, familyRows]) => {
        const color = colorForFamily(family);
        const displayName = family.length > 38 ? `${family.slice(0, 35)}...` : family;
        return {
          type: "scatter",
          mode: "lines+markers",
          name: displayName,
          x: familyRows.map(productLabel),
          y: familyRows.map((rowIndex) => rrpSortValue(rowIndex) || null),
          line: { color, width: 2.6 },
          marker: { color, size: 6 },
          customdata: familyRows.map(productHoverRow),
          hovertemplate: "Family: %{customdata[0]}<br>Material: %{customdata[1]}<br>%{customdata[2]}<br>RRP FY27 %{y:,.2f}<extra></extra>"
        };
      });
    } else {
      rrpTraces = rrpFields.map((item) => ({
        type: "scatter",
        mode: "lines+markers",
        name: item.name,
        x: productLabels,
        y: rrpProducts.map((rowIndex) => toNum(rows[rowIndex][col[item.field]]) || null),
        line: { color: item.color, width: 2.5 },
        marker: { color: item.color, size: 6 },
        connectgaps: true,
        customdata: productHover,
        hovertemplate: "Family: %{customdata[0]}<br>Material: %{customdata[1]}<br>%{customdata[2]}<br>" + item.name + " %{y:,.2f}<extra></extra>"
      }));
    }
    Plotly.react("scatterChart", rrpTraces, {
      ...layoutBase,
      margin: { ...layoutBase.margin, b: 138 },
      title: { text: "RRPs by Product Family", font: { size: 15 } },
      xaxis: { title: { text: "Product inside Product Family", standoff: 34 }, tickangle: -35, automargin: true, tickfont: { size: 10 }, tickvals: productLabels, ticktext: tickText },
      yaxis: { title: "RRP", tickformat: ",.0f", gridcolor: "#edf0f5" },
      legend: {
        orientation: "h",
        x: 0,
        xanchor: "left",
        y: -0.44,
        yanchor: "top",
        bgcolor: "rgba(255, 255, 255, 0.92)",
        bordercolor: "#d9dee7",
        borderwidth: 1
      }
    }, config);

    const volumeRows = forecastVolumeByFamily(state.filtered);
    const volumeChartHeight = Math.max(300, 120 + volumeRows.length * 26);
    Plotly.react("volumeDiffChart", [
      {
        type: "bar",
        orientation: "h",
        name: "FY27 Pacemaker",
        x: volumeRows.map((item) => item.pacemaker),
        y: volumeRows.map((item) => item.family),
        customdata: volumeRows.map((item) => item.delta),
        marker: { color: "#2563eb", line: { color: "#ffffff", width: 1 } },
        hovertemplate: "%{y}<br>Pacemaker %{x:,.0f}<br>Linear minus Pacemaker %{customdata:,.0f}<extra></extra>"
      },
      {
        type: "bar",
        orientation: "h",
        name: "Linear FY27",
        x: volumeRows.map((item) => item.linear),
        y: volumeRows.map((item) => item.family),
        customdata: volumeRows.map((item) => item.delta),
        marker: { color: "#0f766e", line: { color: "#ffffff", width: 1 } },
        hovertemplate: "%{y}<br>Linear %{x:,.0f}<br>Linear minus Pacemaker %{customdata:,.0f}<extra></extra>"
      }
    ], {
      ...layoutBase,
      height: volumeChartHeight,
      margin: { l: 170, r: 24, t: 46, b: 48 },
      title: { text: "Forecast Volume by Product Family", font: { size: 15 } },
      barmode: "group",
      bargap: 0.24,
      xaxis: { title: "Forecast quantity", tickformat: ",.2s", gridcolor: "#edf0f5", zeroline: false },
      yaxis: { automargin: true },
      legend: { orientation: "h", x: 0, y: -0.18, xanchor: "left", yanchor: "top" }
    }, config);

    const segmentRows = segmentRevenueMaterialStats(state.filtered);
    Plotly.react("segmentRevenueChart", [
      {
        type: "bar",
        name: "Revenue",
        x: segmentRows.map((item) => item.segment),
        y: segmentRows.map((item) => item.revenue),
        width: 0.34,
        marker: { color: "#2563eb", line: { color: "#ffffff", width: 1.2 } },
        hovertemplate: "Segment %{x}<br>Revenue %{y:,.0f}<extra></extra>"
      },
      {
        type: "bar",
        name: "Materials",
        x: segmentRows.map((item) => item.segment),
        y: segmentRows.map((item) => item.count),
        yaxis: "y2",
        width: 0.34,
        marker: { color: "#f97316", line: { color: "#ffffff", width: 1.2 } },
        hovertemplate: "Segment %{x}<br>Materials %{y:,.0f}<extra></extra>"
      }
    ], {
      ...layoutBase,
      margin: { l: 58, r: 58, t: 46, b: 58 },
      title: { text: "ABC Segment Revenue and Material Count", font: { size: 15 } },
      barmode: "group",
      bargap: 0.68,
      bargroupgap: 0.14,
      xaxis: { title: "Segment", categoryorder: "array", categoryarray: ["A", "B", "C"], showline: false },
      yaxis: { title: "Revenue", tickformat: ",.2s", gridcolor: "#edf0f5", zeroline: false },
      yaxis2: {
        title: "Materials",
        overlaying: "y",
        side: "right",
        rangemode: "tozero",
        showgrid: false,
        zeroline: false
      },
      legend: { orientation: "h", x: 0, y: -0.2, xanchor: "left", yanchor: "top" }
    }, config);
  }

  function renderTableHead() {
    const cells = headers.map((header, index) => {
      const inputClass = editableHeaders.has(header) ? " input-col" : "";
      const sticky = index === col["Material No."] ? " sticky-a" : index === col["Material Descr."] ? " sticky-b" : "";
      const cellStyle = tableCellStyle(index, header);
      const active = selectedColumnValues(index).length || (state.tableSort.column === index && state.tableSort.direction) ? " active" : "";
      return `
        <th class="${inputClass}${sticky}" data-col-index="${index}" style="${cellStyle}">
          <div class="th-content">
            <span class="th-title" title="${escapeHtml(header)}">${escapeHtml(header)}</span>
            <button class="th-filter${active}" data-col="${index}" type="button" aria-label="Filter ${escapeHtml(header)}"></button>
            <button class="th-resize" data-col="${index}" type="button" aria-label="Resize ${escapeHtml(header)} column"></button>
          </div>
        </th>
      `;
    }).join("");
    els.tableHead.innerHTML = `<tr>${cells}</tr>`;
    els.tableHead.querySelectorAll(".th-filter").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        const colIndex = Number(button.dataset.col);
        if (!els.columnFilterPopup.hidden && els.columnFilterPopup.dataset.col === String(colIndex)) {
          closeColumnFilter();
          return;
        }
        els.columnFilterPopup.dataset.col = String(colIndex);
        openColumnFilter(colIndex, button);
      });
    });
    els.tableHead.querySelectorAll(".th-resize").forEach((handle) => {
      handle.addEventListener("mousedown", startColumnResize);
      handle.addEventListener("click", (event) => event.stopPropagation());
    });
  }

  function startColumnResize(event) {
    event.preventDefault();
    event.stopPropagation();
    closeColumnFilter();
    const colIndex = Number(event.currentTarget.dataset.col);
    const startX = event.clientX;
    const startWidth = columnWidth(colIndex);
    const minWidth = minColumnWidth(colIndex);
    document.body.classList.add("resizing-column");

    const onMove = (moveEvent) => {
      const nextWidth = Math.max(minWidth, Math.min(640, Math.round(startWidth + moveEvent.clientX - startX)));
      if (state.columnWidths[colIndex] === nextWidth) return;
      state.columnWidths[colIndex] = nextWidth;
      applyTableColumnWidths();
    };
    const onUp = () => {
      document.body.classList.remove("resizing-column");
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }

  function renderTableBody() {
    const start = state.page * state.pageSize;
    const pageRows = state.filtered.slice(start, start + state.pageSize);
    const html = pageRows.map((rowIndex) => {
      const row = rows[rowIndex];
      const cells = headers.map((header, index) => {
        const inputClass = editableHeaders.has(header) ? " input-col" : "";
        const sticky = index === col["Material No."] ? " sticky-a" : index === col["Material Descr."] ? " sticky-b" : "";
        const cellStyle = tableCellStyle(index, header);
        const numeric = isNumericColumn(index);
        const className = `${numeric ? "num" : ""}${inputClass}${sticky}`;
        if (header === "ZGRU 2027") {
          return `<td class="${className}" data-col-index="${index}" style="${cellStyle}"><input class="cell-input" data-row="${rowIndex}" data-field="${escapeHtml(header)}" value="${(toNum(row[index]) * 100).toFixed(2)}"></td>`;
        }
        if (header === "ZEVP FY") {
          return `<td class="${className}" data-col-index="${index}" style="${cellStyle}"><input class="cell-input" data-row="${rowIndex}" data-field="${escapeHtml(header)}" value="${toNum(row[index]).toFixed(2)}"></td>`;
        }
        const value = numeric ? formatNumber(row[index], header) : escapeHtml(row[index]);
        return `<td class="${className}" data-col-index="${index}" style="${cellStyle}" title="${escapeHtml(row[index])}">${value}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    }).join("");
    els.tableBody.innerHTML = html;
    els.tableBody.querySelectorAll(".cell-input").forEach((input) => {
      input.addEventListener("input", () => {
        if (commitCellInput(input, { render: false, save: false, status: false })) {
          scheduleEditableSave();
        }
      });
      input.addEventListener("change", handleCellEdit);
      input.addEventListener("blur", handleCellEdit);
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          handleCellEdit({ currentTarget: input });
          input.blur();
        }
      });
    });
  }

  function commitCellInput(input, options = {}) {
    const rowIndex = Number(input.dataset.row);
    const field = input.dataset.field;
    if (!Number.isInteger(rowIndex) || !field || col[field] === undefined) return false;
    let value = parseNumber(input.value);
    if (!Number.isFinite(value)) {
      input.classList.add("invalid");
      if (options.status !== false) updateSaveStatus("Invalid value not saved", "error");
      return false;
    }
    input.classList.remove("invalid");
    if (field === "ZGRU 2027") value /= 100;
    rows[rowIndex][col[field]] = value;
    if (savedInputFields.includes(field)) {
      baseRows[rowIndex][col[field]] = value;
      markDirtyInput(rowIndex, field, value);
    }
    recalcRow(rowIndex);
    if (options.save !== false) saveEditableValues();
    if (options.render !== false) renderDashboard();
    return true;
  }

  function commitActiveCellInput(options = {}) {
    const active = document.activeElement;
    if (active && active.classList && active.classList.contains("cell-input")) {
      commitCellInput(active, options);
      return true;
    }
    return false;
  }

  function handleCellEdit(event) {
    commitCellInput(event.currentTarget);
  }

  function renderPagination() {
    const pages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
    if (state.page >= pages) state.page = pages - 1;
    els.pageLabel.textContent = `${state.page + 1} / ${pages}`;
    els.prevPage.disabled = state.page === 0;
    els.nextPage.disabled = state.page >= pages - 1;
    els.rowCount.textContent = `${state.filtered.length.toLocaleString()} rows`;
  }

  function renderDashboard() {
    const data = summary(state.filtered);
    renderMetrics(data);
    renderCharts(data);
    renderPagination();
    renderTableBody();
  }

  function applyScenario() {
    const zevpDelta = toNum(els.zevpSlider.value) / 100;
    const zgruDelta = toNum(els.zgruSlider.value) / 100;
    els.zevpSliderValue.textContent = `${Number(els.zevpSlider.value).toFixed(1)}%`;
    els.zgruSliderValue.textContent = `${Number(els.zgruSlider.value).toFixed(1)} pp`;
    for (let i = 0; i < rows.length; i += 1) {
      rows[i][col["ZEVP FY"]] = toNum(baseRows[i][col["ZEVP FY"]]) * (1 + zevpDelta);
      rows[i][col["ZGRU 2027"]] = toNum(baseRows[i][col["ZGRU 2027"]]) + zgruDelta;
      recalcRow(i);
    }
    applyFilters();
  }

  function resetScenario() {
    for (let i = 0; i < rows.length; i += 1) rows[i] = baseRows[i].slice();
    els.zevpSlider.value = 0;
    els.zgruSlider.value = 0;
    els.zevpSliderValue.textContent = "0.0%";
    els.zgruSliderValue.textContent = "0.0 pp";
    recalcAll();
    applyFilters();
  }

  function xmlEscape(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&apos;");
  }

  function crc32(bytes) {
    let table = crc32.table;
    if (!table) {
      table = crc32.table = new Uint32Array(256);
      for (let i = 0; i < 256; i += 1) {
        let c = i;
        for (let j = 0; j < 8; j += 1) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
        table[i] = c >>> 0;
      }
    }
    let crc = 0xffffffff;
    for (const byte of bytes) crc = table[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function u16(value) {
    return [value & 255, (value >>> 8) & 255];
  }

  function u32(value) {
    return [value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255];
  }

  function concatBytes(parts) {
    const size = parts.reduce((total, part) => total + part.length, 0);
    const out = new Uint8Array(size);
    let offset = 0;
    for (const part of parts) {
      out.set(part, offset);
      offset += part.length;
    }
    return out;
  }

  function createZip(files) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    let offset = 0;
    for (const file of files) {
      const nameBytes = encoder.encode(file.name);
      const dataBytes = encoder.encode(file.data);
      const crc = crc32(dataBytes);
      const localHeader = new Uint8Array([
        ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
        ...u32(crc), ...u32(dataBytes.length), ...u32(dataBytes.length),
        ...u16(nameBytes.length), ...u16(0)
      ]);
      localParts.push(localHeader, nameBytes, dataBytes);
      const centralHeader = new Uint8Array([
        ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
        ...u32(crc), ...u32(dataBytes.length), ...u32(dataBytes.length),
        ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
        ...u32(0), ...u32(offset)
      ]);
      centralParts.push(centralHeader, nameBytes);
      offset += localHeader.length + nameBytes.length + dataBytes.length;
    }
    const centralOffset = offset;
    const central = concatBytes(centralParts);
    const end = new Uint8Array([
      ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length),
      ...u32(central.length), ...u32(centralOffset), ...u16(0)
    ]);
    return concatBytes([...localParts, central, end]);
  }

  function columnName(index) {
    let name = "";
    let value = index;
    while (value > 0) {
      const rem = (value - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name;
  }

  function sheetXml(table) {
    const rowsXml = table.map((row, rowIdx) => {
      const cells = row.map((value, colIdx) => {
        const ref = `${columnName(colIdx + 1)}${rowIdx + 1}`;
        if (typeof value === "number" && Number.isFinite(value)) {
          return `<c r="${ref}"><v>${value}</v></c>`;
        }
        return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
      }).join("");
      return `<row r="${rowIdx + 1}">${cells}</row>`;
    }).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`;
  }

  function bytesToBase64(bytes) {
    let binary = "";
    const chunkSize = 32768;
    for (let offset = 0; offset < bytes.length; offset += chunkSize) {
      const chunk = bytes.subarray(offset, offset + chunkSize);
      binary += String.fromCharCode(...chunk);
    }
    return btoa(binary);
  }

  function exportXlsx() {
    const outHeaders = ["Material No.", "Material Descr.", "ZEVP FY", "ZBRU FY", "ZGRU FY", "ZNEK FY"];
    const table = [outHeaders];
    for (const rowIndex of state.filtered) {
      const row = rows[rowIndex];
      table.push(outHeaders.map((header) => row[col[header]]));
    }
    const created = new Date().toISOString();
    const files = [
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>` },
      { name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Pricing Output</dc:title><dc:creator>Pricing Engine App</dc:creator><cp:lastModifiedBy>Pricing Engine App</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified></cp:coreProperties>` },
      { name: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Pricing Engine App</Application></Properties>` },
      { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Pricing Output" sheetId="1" r:id="rId1"/></sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>` },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml(table) }
    ];
    const zip = createZip(files);
    const href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${bytesToBase64(zip)}`;
    const downloadLink = document.getElementById("downloadXlsxLink");
    downloadLink.href = href;
    downloadLink.style.display = "inline-flex";
    const link = document.createElement("a");
    link.download = "pricing_output.xlsx";
    link.href = href;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      link.remove();
    }, 0);
  }

  async function init() {
    const restoredInputs = await applySavedEditableValues();
    els.sourceMeta.textContent = `${payload.sourceFile} | ${payload.rows.length.toLocaleString()} rows | generated ${payload.generatedAt}${restoredInputs.cells ? ` | restored ${restoredInputs.cells} saved inputs` : ""}`;
    if (restoredInputs.rows) {
      updateSaveStatus(`Restored ${restoredInputs.rows} rows ${formatSaveTime(restoredInputs.savedAt)}`, "saved");
    }
    recalcAll();
    populateFilters();
    setupRevenueSlider();
    renderTableHead();
    applyFilters();

    els.search.addEventListener("input", () => {
      state.search = els.search.value;
      state.page = 0;
      applyFilters();
    });
    els.minRevenue.addEventListener("input", () => {
      state.minRevenue = Number(els.minRevenue.value) || 0;
      els.minRevenueValue.textContent = formatCompact(state.minRevenue);
      state.page = 0;
      applyFilters();
    });
    els.zevpSlider.addEventListener("input", applyScenario);
    els.zgruSlider.addEventListener("input", applyScenario);
    document.getElementById("savePrices").addEventListener("click", () => {
      commitActiveCellInput({ save: false, render: false });
      saveEditableValues("manual-button", { downloadBackup: true });
    });
    els.importPrices.addEventListener("click", () => els.importPricesFile.click());
    els.importPricesFile.addEventListener("change", () => {
      importSavedPricesFile(els.importPricesFile.files?.[0]);
    });
    document.getElementById("resetScenario").addEventListener("click", resetScenario);
    document.getElementById("exportXlsx").addEventListener("click", exportXlsx);
    els.columnFilterPopup.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    els.pageSize.addEventListener("change", () => {
      state.pageSize = Number(els.pageSize.value);
      state.page = 0;
      renderDashboard();
    });
    els.prevPage.addEventListener("click", () => {
      state.page = Math.max(0, state.page - 1);
      renderDashboard();
    });
    els.nextPage.addEventListener("click", () => {
      const pages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
      state.page = Math.min(pages - 1, state.page + 1);
      renderDashboard();
    });
    document.addEventListener("click", (event) => {
      if (els.columnFilterPopup.hidden) return;
      if (els.columnFilterPopup.contains(event.target)) return;
      if (event.target.closest(".th-filter")) return;
      closeColumnFilter();
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeColumnFilter();
    });
    window.addEventListener("resize", closeColumnFilter);
    window.addEventListener("beforeunload", () => {
      commitActiveCellInput({ save: false, render: false });
      saveEditableValues();
    });
    window.addEventListener("pagehide", () => {
      commitActiveCellInput({ save: false, render: false });
      saveEditableValues();
    });
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") {
        commitActiveCellInput({ save: false, render: false });
        saveEditableValues("tab-hidden");
      }
    });
    setInterval(() => {
      commitActiveCellInput({ save: false, render: false });
      saveEditableValues();
    }, 60 * 60 * 1000);
  }

  init().catch((error) => {
    console.error("Unable to initialize Pricing Engine.", error);
    updateSaveStatus("Restore failed", "error");
    recalcAll();
    populateFilters();
    setupRevenueSlider();
    renderTableHead();
    applyFilters();
  });
})();
"""


HTML_SHELL_TOP = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Pricing Engine App</title>
  <style>
"""


HTML_BODY_START = r"""  </style>
</head>
<body>
  <div class="app">
    <header class="topbar">
      <div class="title">
        <h1>Pricing Engine</h1>
        <span id="sourceMeta"></span>
      </div>
      <div class="actions">
        <span class="save-status" id="saveStatus">Not saved yet</span>
        <button class="btn" id="savePrices" type="button">Save Prices</button>
        <button class="btn" id="importPrices" type="button">Import Prices</button>
        <input class="file-input-hidden" id="importPricesFile" type="file" accept="application/json,.json">
        <button class="btn" id="resetScenario" type="button">Reset Scenario</button>
        <button class="btn primary" id="exportXlsx" type="button">Output XLSX</button>
        <a class="btn" id="downloadXlsxLink" download="pricing_output.xlsx" style="display:none">Download XLSX</a>
      </div>
    </header>
    <div class="workspace">
      <aside class="sidebar">
        <section class="panel">
          <h2>Filters</h2>
          <div class="field">
            <label for="search">Search</label>
            <input id="search" type="search" placeholder="Material or product">
          </div>
          <div id="filterArea"></div>
          <div class="slider-row">
            <label for="minRevenue">Min FY27 Revenue</label>
            <span class="slider-value" id="minRevenueValue">0</span>
            <input id="minRevenue" type="range" min="0" max="100000" value="0">
          </div>
        </section>
        <section class="panel">
          <h2>Scenario</h2>
          <div class="slider-row">
            <label for="zevpSlider">ZEVP FY Delta</label>
            <span class="slider-value" id="zevpSliderValue">0.0%</span>
            <input id="zevpSlider" type="range" min="-20" max="20" value="0" step="0.5">
          </div>
          <div class="slider-row">
            <label for="zgruSlider">ZGRU 2027 Delta</label>
            <span class="slider-value" id="zgruSliderValue">0.0 pp</span>
            <input id="zgruSlider" type="range" min="-20" max="20" value="0" step="0.5">
          </div>
        </section>
      </aside>
      <main class="main">
        <section class="metrics" id="metrics"></section>
        <section class="charts">
          <div class="chart" id="revenueChart"></div>
          <div class="chart" id="cashChart"></div>
          <div class="chart" id="groupChart"></div>
          <div class="chart" id="scatterChart"></div>
        </section>
        <section class="charts insight-charts">
          <div class="chart chart-scroll"><div class="chart-scroll-inner" id="volumeDiffChart"></div></div>
          <div class="chart" id="segmentRevenueChart"></div>
        </section>
      </main>
      <section class="table-panel">
        <div class="table-toolbar">
          <div class="left">
            <span class="row-count" id="rowCount"></span>
          </div>
          <div class="right">
            <select id="pageSize">
              <option value="50">50 rows</option>
              <option value="100" selected>100 rows</option>
              <option value="250">250 rows</option>
            </select>
            <button class="btn" id="prevPage" type="button">Prev</button>
            <span class="row-count" id="pageLabel">1 / 1</span>
            <button class="btn" id="nextPage" type="button">Next</button>
          </div>
        </div>
        <div class="grid-wrap">
          <table>
            <thead id="tableHead"></thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
      </section>
    </div>
  </div>
  <div id="columnFilterPopup" class="column-filter-popover" hidden></div>
"""


def write_html(payload: dict) -> None:
    """Write the self-contained static HTML application to disk.

    The output includes local CSS, local Plotly JavaScript, the workbook-derived
    JSON payload, and the app JavaScript. There are no external script links.
    """
    data_json = json.dumps(payload, ensure_ascii=True, separators=(",", ":")).replace("</", "<\\/")
    plotly_js = PLOTLY_PATH.read_text(encoding="utf-8").replace("</", "<\\/")
    html = (
        HTML_SHELL_TOP
        + STYLE
        + HTML_BODY_START
        + "\n  <script>\n"
        + plotly_js
        + "\n  </script>\n"
        + '  <script id="pricing-data" type="application/json">'
        + data_json
        + "</script>\n"
        + "  <script>\n"
        + APP_JS
        + "\n  </script>\n"
        + "</body>\n</html>\n"
    )
    OUTPUT_PATH.write_text(html, encoding="utf-8")


def main() -> None:
    """Build the app and print a small validation summary for reviewers."""
    payload = build_payload()
    write_html(payload)
    print(f"Wrote: {OUTPUT_PATH}")
    print(f"Rows: {len(payload['rows'])}")
    print(f"Projected NSV helper hits: {payload['projectedBaseHits']}")
    print(f"Linear forecast model rows: {payload['linearForecast']['modelRows']}")
    print(f"Constants: {payload['constants']}")
    print(f"QA max diffs: {payload['qa']}")


if __name__ == "__main__":
    main()
