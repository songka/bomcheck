from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .binding_library import BindingLibrary, BindingGroup
from .models import (
    BindingProjectResult,
    ExecutionResult,
    ImportantMaterialHit,
    MissingItem,
    ReplacementRecord,
    ReplacementSummary,
    RequirementGroupResult,
)
from .text_utils import normalize_text

BLACK_FILL = PatternFill(start_color="000000", end_color="000000", fill_type="solid")


class ExcelProcessor:
    def __init__(self, config):
        self.config = config

    def execute(self, excel_path: Path, binding_library: BindingLibrary) -> ExecutionResult:
        wb = load_workbook(excel_path)
        replacement_summary, _ = self._apply_replacements(wb)
        part_quantities, part_desc = self._extract_part_quantities(wb)
        for record in replacement_summary.records:
            if not record.replacement_part_no:
                continue
            qty = part_quantities.pop(record.invalid_part_no, 0.0)
            if qty:
                part_quantities[record.replacement_part_no] += qty
                if record.replacement_desc:
                    part_desc.setdefault(record.replacement_part_no, record.replacement_desc)
        binding_results, missing_items = self._evaluate_binding_requirements(part_quantities, part_desc, binding_library)
        important_hits = self._scan_important_materials(part_quantities)
        self._write_result_sheets(
            wb,
            replacement_summary,
            binding_results,
            important_hits,
            missing_items,
            part_quantities,
        )
        wb.save(excel_path)
        return ExecutionResult(
            replacement_summary=replacement_summary,
            binding_results=binding_results,
            important_hits=important_hits,
            missing_items=missing_items,
        )

    def _apply_replacements(self, wb) -> Tuple[ReplacementSummary, Dict[str, str]]:
        summary = ReplacementSummary()
        replaced_parts: Dict[str, str] = {}
        invalid_wb = load_workbook(self.config.invalid_part_db)
        invalid_ws = invalid_wb.active
        invalid_entries: List[Tuple[str, str, Optional[str], Optional[str]]] = []
        for row in invalid_ws.iter_rows(min_row=2, values_only=True):
            invalid_no = str(row[0]).strip() if row[0] else ""
            invalid_desc = str(row[1]).strip() if row[1] else ""
            replacement_no = str(row[2]).strip() if row[2] else None
            replacement_desc = str(row[3]).strip() if row[3] else None
            if invalid_no:
                invalid_entries.append((invalid_no, invalid_desc, replacement_no, replacement_desc))
        for ws in wb.worksheets:
            for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                cell_value = row[0].value
                if not cell_value:
                    continue
                part_no = str(cell_value).strip()
                match = next((entry for entry in invalid_entries if entry[0] == part_no), None)
                if not match:
                    continue
                summary.total_invalid_found += 1
                invalid_no, invalid_desc, replacement_no, replacement_desc = match
                for cell in row:
                    cell.fill = BLACK_FILL
                if replacement_no:
                    replacement_col = ws.max_column + 1
                    ws.cell(row=row_idx, column=replacement_col).value = replacement_no
                    ws.cell(row=row_idx, column=replacement_col + 1).value = replacement_desc
                    replaced_parts[part_no] = replacement_no
                    summary.total_replaced += 1
                summary.records.append(
                    ReplacementRecord(
                        invalid_part_no=invalid_no,
                        invalid_desc=invalid_desc,
                        replacement_part_no=replacement_no,
                        replacement_desc=replacement_desc,
                        sheet_name=ws.title,
                        row_index=row_idx,
                    )
                )
        return summary, replaced_parts

    def _extract_part_quantities(self, wb) -> Tuple[Dict[str, float], Dict[str, str]]:
        part_quantities: Dict[str, float] = defaultdict(float)
        part_descriptions: Dict[str, str] = {}
        for ws in wb.worksheets:
            qty_col_idx = self._identify_quantity_column(ws)
            for row in ws.iter_rows(min_row=2):
                part_cell = row[0]
                if not part_cell.value:
                    continue
                part_no = str(part_cell.value).strip()
                desc_cell = row[1] if len(row) > 1 else None
                if desc_cell and desc_cell.value:
                    part_descriptions.setdefault(part_no, str(desc_cell.value).strip())
                quantity = 1.0
                if qty_col_idx is not None and qty_col_idx < len(row):
                    quantity_cell = row[qty_col_idx]
                    try:
                        quantity = float(quantity_cell.value)
                    except (TypeError, ValueError):
                        quantity = 1.0
                part_quantities[part_no] += quantity
        return part_quantities, part_descriptions

    def _identify_quantity_column(self, ws) -> Optional[int]:
        numeric_scores = []  # (col_idx, numeric_count, decimal_count)
        for col_idx in range(3, ws.max_column):
            numeric_count = 0
            decimal_count = 0
            valid_column = True
            for cell in ws.iter_rows(min_row=2, min_col=col_idx + 1, max_col=col_idx + 1, values_only=True):
                value = cell[0]
                if value is None or value == "":
                    continue
                try:
                    number = float(value)
                except (TypeError, ValueError):
                    valid_column = False
                    break
                numeric_count += 1
                if abs(number - round(number)) > 1e-6:
                    decimal_count += 1
            if valid_column and numeric_count:
                numeric_scores.append((col_idx, numeric_count, decimal_count))
        if not numeric_scores:
            return None
        numeric_scores.sort(key=lambda item: (-item[1], item[2]))
        best_col = numeric_scores[0][0]
        return best_col

    def _evaluate_binding_requirements(
        self,
        part_quantities: Dict[str, float],
        part_desc: Dict[str, str],
        binding_library: BindingLibrary,
    ) -> Tuple[List[BindingProjectResult], List[MissingItem]]:
        results: List[BindingProjectResult] = []
        missing_items: Dict[str, MissingItem] = {}
        for project in binding_library.iter_projects():
            qty = part_quantities.get(project.index_part_no)
            if not qty:
                continue
            group_results: List[RequirementGroupResult] = []
            for group in project.required_groups:
                result = self._evaluate_group(group, part_quantities)
                group_results.append(result)
                if result.missing_qty > 0:
                    for part_no in result.missing_choices:
                        description = part_desc.get(part_no, "") or self._lookup_choice_desc(group, part_no)
                        item = missing_items.setdefault(
                            part_no,
                            MissingItem(
                                part_no=part_no,
                                desc=description,
                                missing_qty=0.0,
                            ),
                        )
                        item.desc = item.desc or description
                        item.missing_qty += result.missing_qty
            results.append(
                BindingProjectResult(
                    project_desc=project.project_desc,
                    index_part_no=project.index_part_no,
                    matched_quantity=qty,
                    requirement_results=group_results,
                )
            )
        return results, list(missing_items.values())

    def _evaluate_group(self, group: BindingGroup, part_quantities: Dict[str, float]) -> RequirementGroupResult:
        required_qty = group.number or 1.0
        available_qty = 0.0
        missing_choices: List[str] = []
        for choice in group.choices:
            if not choice.part_no:
                continue
            if not self._choice_applicable(choice, part_quantities):
                continue
            quantity = part_quantities.get(choice.part_no, 0.0)
            if quantity:
                contributes = quantity if choice.number is None else min(quantity, choice.number)
                available_qty += contributes
            else:
                missing_choices.append(choice.part_no)
        missing_qty = max(required_qty - available_qty, 0.0)
        if missing_qty > 0 and not missing_choices:
            missing_choices = [choice.part_no for choice in group.choices if choice.part_no]
        return RequirementGroupResult(
            group_name=group.group_name,
            required_qty=required_qty,
            available_qty=available_qty,
            missing_qty=missing_qty,
            missing_choices=missing_choices,
        )

    def _lookup_choice_desc(self, group: BindingGroup, part_no: str) -> str:
        for choice in group.choices:
            if choice.part_no == part_no and choice.desc:
                return choice.desc
        return ""

    def _choice_applicable(self, choice, part_quantities: Dict[str, float]) -> bool:
        mode = (choice.condition_mode or "").upper()
        if mode == "ALL":
            return all(part_quantities.get(part_no) for part_no in choice.condition_part_nos)
        if mode == "ANY":
            return any(part_quantities.get(part_no) for part_no in choice.condition_part_nos)
        if mode == "NOTANY":
            return not any(part_quantities.get(part_no) for part_no in choice.condition_part_nos)
        return True

    def _scan_important_materials(self, part_quantities: Dict[str, float]) -> List[ImportantMaterialHit]:
        important_path = self.config.important_materials
        hits: List[ImportantMaterialHit] = []
        if not important_path.exists():
            return hits
        keywords = [line.strip() for line in important_path.read_text(encoding="utf-8").splitlines() if line.strip()]
        for keyword in keywords:
            normalized_keyword = normalize_text(keyword)
            total_qty = 0.0
            for part_no, qty in part_quantities.items():
                if normalized_keyword in normalize_text(part_no):
                    total_qty += qty
            if total_qty:
                hits.append(
                    ImportantMaterialHit(keyword=keyword, converted_keyword=normalized_keyword, total_quantity=total_qty)
                )
        return hits

    def _write_result_sheets(
        self,
        wb,
        replacement_summary: ReplacementSummary,
        binding_results: List[BindingProjectResult],
        important_hits: List[ImportantMaterialHit],
        missing_items: List[MissingItem],
        part_quantities: Dict[str, float],
    ) -> None:
        if "执行统计" in wb.sheetnames:
            del wb["执行统计"]
        if "剩余物料" in wb.sheetnames:
            del wb["剩余物料"]
        summary_ws = wb.create_sheet("执行统计")
        summary_ws.append(["失效料号数量", replacement_summary.total_invalid_found])
        summary_ws.append(["已替换数量", replacement_summary.total_replaced])
        summary_ws.append([])
        summary_ws.append(["绑定料号统计"])
        summary_ws.append(["项目描述", "索引料号", "需求分组", "需求数量", "可用数量", "缺少数量"])
        for result in binding_results:
            for group_result in result.requirement_results:
                summary_ws.append([
                    result.project_desc,
                    result.index_part_no,
                    group_result.group_name,
                    group_result.required_qty,
                    group_result.available_qty,
                    group_result.missing_qty,
                ])
        summary_ws.append([])
        summary_ws.append(["缺失物料"])
        summary_ws.append(["料号", "描述", "缺少数量"])
        for item in missing_items:
            summary_ws.append([item.part_no, item.desc, item.missing_qty])
        summary_ws.append([])
        summary_ws.append(["重要物料统计"])
        summary_ws.append(["关键字", "标准关键字", "数量"])
        for hit in important_hits:
            summary_ws.append([hit.keyword, hit.converted_keyword, hit.total_quantity])

        remainder_ws = wb.create_sheet("剩余物料")
        remainder_ws.append(["料号", "数量"])
        for part_no, qty in part_quantities.items():
            remainder_ws.append([part_no, qty])
