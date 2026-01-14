from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

from openpyxl.worksheet.worksheet import Worksheet

MONTHS = {
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december",
}


@dataclass(frozen=True)
class WeekBlock:
    label: str
    start_col: int
    end_col: int


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = " ".join(text.split())
    return text.casefold()


def normalize_header(value: object) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def is_qty_header(value: object) -> bool:
    text = normalize_header(value)
    return text in {"q ty", "qty", "q'ty", "q ty"} or text.replace(" ", "") == "qty"


def is_manhour_header(value: object) -> bool:
    text = normalize_header(value)
    return text in {"m hour", "man hour", "man hours", "manhour"}


def is_timesheet_header(value: object) -> bool:
    text = normalize_header(value)
    return text in {"timesheet", "time sheet", "timesheet hours", "timesheet h"}


def is_month_label(label: str) -> bool:
    return label in MONTHS


def parse_number(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def header_pattern_matches(values: Iterable[object]) -> bool:
    items = list(values)
    if len(items) != 5:
        return False
    return (
        is_qty_header(items[0])
        and is_manhour_header(items[1])
        and is_qty_header(items[2])
        and is_manhour_header(items[3])
        and is_timesheet_header(items[4])
    )


def detect_week_blocks(ws: Worksheet) -> List[WeekBlock]:
    blocks: List[WeekBlock] = []
    max_col = ws.max_column

    for merged in ws.merged_cells.ranges:
        if merged.min_row == 10 and merged.max_row == 10 and merged.max_col - merged.min_col == 4:
            label_value = ws.cell(row=10, column=merged.min_col).value
            label = normalize_text(label_value)
            if not label or is_month_label(label):
                continue
            header_values = [ws.cell(row=13, column=col).value for col in range(merged.min_col, merged.max_col + 1)]
            if header_pattern_matches(header_values):
                blocks.append(
                    WeekBlock(
                        label=str(label_value).strip(),
                        start_col=merged.min_col,
                        end_col=merged.max_col,
                    )
                )

    if blocks:
        return blocks

    for start_col in range(1, max_col - 3):
        end_col = start_col + 4
        if end_col > max_col:
            continue
        label_value = ws.cell(row=10, column=start_col).value
        label = normalize_text(label_value)
        if not label or is_month_label(label):
            continue
        header_values = [ws.cell(row=13, column=col).value for col in range(start_col, end_col + 1)]
        if header_pattern_matches(header_values):
            blocks.append(
                WeekBlock(
                    label=str(label_value).strip(),
                    start_col=start_col,
                    end_col=end_col,
                )
            )

    return blocks


def build_week_map(blocks: Iterable[WeekBlock]) -> Dict[str, WeekBlock]:
    return {normalize_text(block.label): block for block in blocks}


def extract_week_number(label: str) -> Optional[str]:
    match = re.search(r"(\d+)", label)
    return match.group(1) if match else None


def match_week_label(selected_label: str, block_map: Dict[str, WeekBlock]) -> Optional[WeekBlock]:
    normalized = normalize_text(selected_label)
    if normalized in block_map:
        return block_map[normalized]
    selected_number = extract_week_number(normalized)
    if not selected_number:
        return None
    for key, block in block_map.items():
        if extract_week_number(key) == selected_number:
            return block
    return None


def find_data_start_row(ws: Worksheet, min_row: int = 14) -> int:
    for row in range(min_row, ws.max_row + 1):
        value = ws.cell(row=row, column=3).value
        if value is None:
            continue
        if normalize_text(value) and normalize_text(value) != normalize_text("Description of Work"):
            return row
    return min_row
