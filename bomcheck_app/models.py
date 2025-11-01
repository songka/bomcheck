from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class ReplacementRecord:
    invalid_part_no: str
    invalid_desc: str
    replacement_part_no: Optional[str]
    replacement_desc: Optional[str]
    sheet_name: str
    row_index: int


@dataclass
class ReplacementSummary:
    total_invalid_found: int = 0
    total_replaced: int = 0
    records: List[ReplacementRecord] = field(default_factory=list)


@dataclass
class RequirementGroupResult:
    group_name: str
    required_qty: float
    available_qty: float
    missing_qty: float
    missing_choices: List[str] = field(default_factory=list)
    matched_details: Dict[str, float] = field(default_factory=dict)


@dataclass
class BindingProjectResult:
    project_desc: str
    index_part_no: str
    matched_quantity: float
    requirement_results: List[RequirementGroupResult] = field(default_factory=list)

    @property
    def has_missing(self) -> bool:
        return any(group.missing_qty > 0 for group in self.requirement_results)


@dataclass
class ImportantMaterialHit:
    keyword: str
    converted_keyword: str
    total_quantity: float
    matched_parts: Dict[str, float] = field(default_factory=dict)


@dataclass
class MissingItem:
    part_no: str
    desc: str
    missing_qty: float


@dataclass
class ExecutionResult:
    replacement_summary: ReplacementSummary
    binding_results: List[BindingProjectResult]
    important_hits: List[ImportantMaterialHit]
    missing_items: List[MissingItem]
    debug_logs: List[str] = field(default_factory=list)

    @property
    def has_missing(self) -> bool:
        return bool(self.missing_items) or any(result.has_missing for result in self.binding_results)
