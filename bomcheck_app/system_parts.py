from __future__ import annotations

from dataclasses import dataclass
from math import isclose
import csv
import re
from pathlib import Path
from typing import Dict

from openpyxl import Workbook, load_workbook

from .excel_processor import format_quantity_text, normalize_part_no
from .text_utils import normalized_variants


@dataclass
class SystemPartRecord:
    part_no: str
    description: str
    unit: str
    applicant: str
    inventory: float | None

    @property
    def categories(self) -> tuple[str, ...]:
        parts = [segment.strip() for segment in self.description.split(";") if segment.strip()]
        return tuple(parts) if parts else ("未分类",)

    @property
    def inventory_display(self) -> str:
        if self.inventory is None:
            return ""
        if isclose(self.inventory, round(self.inventory), abs_tol=1e-6):
            return str(int(round(self.inventory)))
        return f"{round(self.inventory, 4):g}"


class SystemPartRepository:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.records: list[SystemPartRecord] = []
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(f"系统料号文件不存在：{self.path}")

        workbook = load_workbook(self.path, data_only=True, read_only=True)
        sheet = workbook.active
        records: list[SystemPartRecord] = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            part_no = _safe_str(row[0])
            description = _safe_str(row[1])
            unit = _safe_str(row[2])
            applicant = _clean_applicant_text(_safe_str(row[3]))
            inventory_value = row[4] if len(row) > 4 else None
            if not part_no:
                continue
            inventory = _convert_inventory(inventory_value)
            records.append(SystemPartRecord(part_no, description, unit, applicant, inventory))

        workbook.close()
        self.records = records

    def build_hierarchy(self, query: str | None = None) -> Dict[str, Dict]:
        keywords = _prepare_keywords(query)
        root: Dict[str, Dict] = {"children": {}, "parts": []}

        for record in self.records:
            if keywords and not _matches_query(record, keywords):
                continue
            node = root
            for category in record.categories:
                node = node["children"].setdefault(category, {"children": {}, "parts": []})
            node.setdefault("parts", []).append(record)

        return root

    def search(self, query: str) -> list[SystemPartRecord]:
        keywords = _prepare_keywords(query)
        if not keywords:
            return list(self.records)
        return [record for record in self.records if _matches_query(record, keywords)]


def generate_system_part_excel(
    source_path: Path,
    invalid_part_path: Path,
    blocked_applicant_path: Path,
) -> Path:
    if not source_path.exists():
        raise FileNotFoundError(f"找不到系统料号原始文件：{source_path}")

    invalid_part_numbers = _load_invalid_part_numbers(invalid_part_path)
    blocked_applicants = _load_blocked_applicants(blocked_applicant_path)

    records = _parse_system_parts(source_path)

    filtered_records: list[SystemPartRecord] = []
    seen_parts: set[str] = set()
    for record in records:
        normalized_part = normalize_part_no(record.part_no)
        if not normalized_part.startswith("UC3"):
            continue
        if normalized_part in invalid_part_numbers:
            continue
        if _should_block(record.applicant, blocked_applicants):
            continue
        if normalized_part in seen_parts:
            continue
        seen_parts.add(normalized_part)
        filtered_records.append(record)

    destination = source_path.with_suffix(".xlsx")
    if destination == source_path:
        destination = source_path.with_name(source_path.name + ".xlsx")

    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "系统料号"
    sheet.append(["料号", "描述", "单位", "申请人", "库存"])

    for record in filtered_records:
        sheet.append(
            [
                record.part_no,
                record.description,
                record.unit,
                record.applicant,
                _format_inventory_cell(record.inventory),
            ]
        )

    workbook.save(destination)
    workbook.close()
    return destination


def _parse_system_parts(path: Path) -> list[SystemPartRecord]:
    suffix = path.suffix.lower()
    if suffix in {".tsv", ".txt"}:
        return _parse_tsv(path)
    if suffix in {".xlsx", ".xlsm"}:
        return _parse_excel(path)
    raise ValueError(f"不支持的系统料号文件格式：{path.suffix}")


def _parse_tsv(path: Path) -> list[SystemPartRecord]:
    records: list[SystemPartRecord] = []
    with path.open("r", encoding="utf-8") as handle:
        reader = csv.reader(handle, delimiter="\t")
        for row in reader:
            if len(row) < 11:
                continue
            part_no = _safe_str(row[1])
            description = _safe_str(row[3])
            if not part_no:
                continue
            unit = _safe_str(row[6])
            applicant = _clean_applicant_text(_safe_str(row[9]))
            inventory = _convert_inventory(row[10])
            records.append(SystemPartRecord(part_no, description, unit, applicant, inventory))
    return records


def _parse_excel(path: Path) -> list[SystemPartRecord]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    records: list[SystemPartRecord] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        part_no = _safe_str(row[0])
        description = _safe_str(row[1])
        if not part_no:
            continue
        unit = _safe_str(row[2])
        applicant = _clean_applicant_text(_safe_str(row[3]))
        inventory = _convert_inventory(row[4] if len(row) > 4 else None)
        records.append(SystemPartRecord(part_no, description, unit, applicant, inventory))
    workbook.close()
    return records


def _convert_inventory(value) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        cleaned = str(value).strip()
        try:
            return float(cleaned)
        except ValueError:
            return None


def _format_inventory_cell(value: float | None):
    if value is None:
        return ""
    formatted = format_quantity_text(value)
    try:
        return float(formatted)
    except (TypeError, ValueError):
        return formatted


def _load_invalid_part_numbers(path: Path) -> set[str]:
    if not path.exists():
        return set()
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    invalid_numbers: set[str] = set()
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        invalid_no = _safe_str(row[0])
        if invalid_no:
            invalid_numbers.add(normalize_part_no(invalid_no))
    workbook.close()
    return invalid_numbers


def _load_blocked_applicants(path: Path) -> "BlockedApplicantMatcher":
    matcher = BlockedApplicantMatcher()
    if not path.exists():
        return matcher
    content = path.read_text(encoding="utf-8")
    tokens = re.split(r"[\s,;，；]+", content)
    for token in tokens:
        matcher.add(token)
    return matcher


def _should_block(applicant: str, blocked: "BlockedApplicantMatcher") -> bool:
    if not applicant:
        return False
    return blocked.matches(applicant)


def _matches_query(record: SystemPartRecord, keywords: list[str]) -> bool:
    if not keywords:
        return True
    text = " ".join(
        [record.part_no.lower(), record.description.lower(), record.applicant.lower()]
    )
    return all(keyword in text for keyword in keywords)


def _safe_str(value) -> str:
    return str(value).strip() if value not in (None, "") else ""


def _clean_applicant_text(value: str) -> str:
    return value.strip().strip(",，;；")


class BlockedApplicantMatcher:
    def __init__(self) -> None:
        self._variant_lengths: dict[str, set[int]] = {}

    def add(self, value: str) -> None:
        if value is None:
            return
        cleaned = value.strip()
        if not cleaned:
            return
        length = len(cleaned)
        for variant in normalized_variants(cleaned):
            if not variant:
                continue
            self._variant_lengths.setdefault(variant, set()).add(length)

    def matches(self, applicant: str) -> bool:
        if not self._variant_lengths or not applicant:
            return False
        for token in _split_applicant_tokens(applicant):
            token_length = len(token)
            for variant in normalized_variants(token):
                lengths = self._variant_lengths.get(variant)
                if not lengths:
                    continue
                if token_length == 2:
                    if 2 in lengths:
                        return True
                else:
                    return True
        return False


def _split_applicant_tokens(value: str) -> list[str]:
    raw_tokens = re.split(r"[\s,;，；/、\|]+", value)
    tokens: list[str] = []
    for token in raw_tokens:
        cleaned = token.strip().strip("()（）[]【】'\"")
        if cleaned:
            tokens.append(cleaned)
    return tokens


def _prepare_keywords(query: str | None) -> list[str]:
    if not query:
        return []
    tokens = [segment.strip().lower() for segment in re.split(r"[\s,;，；]+", query)]
    return [token for token in tokens if token]

