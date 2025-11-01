"""Utilities for loading and editing the binding part number library."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Sequence

import pandas as pd


@dataclass
class BindingChoice:
    part_no: str
    desc: str
    condition_mode: str | None = None
    condition_part_nos: List[str] = field(default_factory=list)
    number: float | None = None

    @classmethod
    def from_mapping(cls, mapping: Dict) -> "BindingChoice":
        return cls(
            part_no=str(mapping.get("partNo", "")).strip(),
            desc=str(mapping.get("desc", "")).strip(),
            condition_mode=(mapping.get("conditionMode") or None),
            condition_part_nos=[str(v).strip() for v in mapping.get("conditionPartNos", []) if str(v).strip()],
            number=_to_float(mapping.get("number")),
        )

    def to_mapping(self) -> Dict:
        mapping: Dict[str, object] = {
            "partNo": self.part_no,
            "desc": self.desc,
        }
        if self.condition_mode:
            mapping["conditionMode"] = self.condition_mode
        if self.condition_part_nos:
            mapping["conditionPartNos"] = self.condition_part_nos
        if self.number is not None:
            mapping["number"] = self.number
        return mapping


@dataclass
class BindingGroup:
    group_name: str
    number: float | None
    choices: List[BindingChoice] = field(default_factory=list)

    @classmethod
    def from_mapping(cls, mapping: Dict) -> "BindingGroup":
        return cls(
            group_name=str(mapping.get("groupName", "")).strip(),
            number=_to_float(mapping.get("number")),
            choices=[BindingChoice.from_mapping(item) for item in mapping.get("choices", [])],
        )

    def to_mapping(self) -> Dict:
        mapping: Dict[str, object] = {
            "groupName": self.group_name,
            "choices": [choice.to_mapping() for choice in self.choices],
        }
        if self.number is not None:
            mapping["number"] = self.number
        return mapping


@dataclass
class BindingProject:
    project_desc: str
    index_part_no: str
    index_part_desc: str
    required_groups: List[BindingGroup] = field(default_factory=list)

    @classmethod
    def from_mapping(cls, mapping: Dict) -> "BindingProject":
        return cls(
            project_desc=str(mapping.get("projectDesc", "")).strip(),
            index_part_no=str(mapping.get("indexPartNo", "")).strip(),
            index_part_desc=str(mapping.get("indexPartDesc", "")).strip(),
            required_groups=[BindingGroup.from_mapping(item) for item in mapping.get("requiredGroups", [])],
        )

    def to_mapping(self) -> Dict:
        return {
            "projectDesc": self.project_desc,
            "indexPartNo": self.index_part_no,
            "indexPartDesc": self.index_part_desc,
            "requiredGroups": [group.to_mapping() for group in self.required_groups],
        }


class BindingLibrary:
    """Read and write binding library data from disk."""

    def __init__(self, path: Path):
        self.path = path

    def load(self) -> List[BindingProject]:
        if not self.path.exists():
            return []
        raw_text = self.path.read_text(encoding="utf-8")
        raw_text = raw_text.strip()
        if not raw_text:
            return []
        if not raw_text.startswith("["):
            raw_text = f"[{raw_text}]"
        data = json.loads(raw_text)
        return [BindingProject.from_mapping(item) for item in data]

    def save(self, projects: Sequence[BindingProject]) -> None:
        data = [project.to_mapping() for project in projects]
        serialized = json.dumps(data, ensure_ascii=False, indent=2)
        self.path.write_text(serialized + "\n", encoding="utf-8")

    def export_to_excel(self, target: Path, projects: Sequence[BindingProject]) -> None:
        rows = []
        for project in projects:
            for group in project.required_groups:
                for choice in group.choices:
                    rows.append(
                        {
                            "projectDesc": project.project_desc,
                            "indexPartNo": project.index_part_no,
                            "indexPartDesc": project.index_part_desc,
                            "groupName": group.group_name,
                            "groupNumber": group.number,
                            "choicePartNo": choice.part_no,
                            "choiceDesc": choice.desc,
                            "choiceConditionMode": choice.condition_mode,
                            "choiceConditionPartNos": ",".join(choice.condition_part_nos),
                            "choiceNumber": choice.number,
                        }
                    )
        df = pd.DataFrame(rows)
        if df.empty:
            df = pd.DataFrame(
                columns=[
                    "projectDesc",
                    "indexPartNo",
                    "indexPartDesc",
                    "groupName",
                    "groupNumber",
                    "choicePartNo",
                    "choiceDesc",
                    "choiceConditionMode",
                    "choiceConditionPartNos",
                    "choiceNumber",
                ]
            )
        df.to_excel(target, index=False)

    def import_from_excel(self, source: Path) -> List[BindingProject]:
        df = pd.read_excel(source)
        required_columns = {"projectDesc", "indexPartNo", "indexPartDesc", "groupName", "choicePartNo", "choiceDesc"}
        if not required_columns.issubset(df.columns):
            missing = ", ".join(sorted(required_columns - set(df.columns)))
            raise ValueError(f"Excel文件缺少必要欄位: {missing}")
        projects: Dict[tuple[str, str], BindingProject] = {}
        for _, row in df.iterrows():
            project_key = (str(row["projectDesc"]).strip(), str(row["indexPartNo"]).strip())
            project = projects.get(project_key)
            if not project:
                project = BindingProject(
                    project_desc=project_key[0],
                    index_part_no=project_key[1],
                    index_part_desc=str(row.get("indexPartDesc", "")).strip(),
                    required_groups=[],
                )
                projects[project_key] = project
            group_name = str(row.get("groupName", "")).strip()
            if not group_name:
                continue
            group = next((g for g in project.required_groups if g.group_name == group_name), None)
            if not group:
                group_number = _to_float(row.get("groupNumber"))
                group = BindingGroup(group_name=group_name, number=group_number, choices=[])
                project.required_groups.append(group)
            choice_part_no = str(row.get("choicePartNo", "")).strip()
            choice_desc = str(row.get("choiceDesc", "")).strip()
            if not choice_part_no:
                continue
            condition_mode = str(row.get("choiceConditionMode", "")).strip() or None
            condition_part_nos_raw = str(row.get("choiceConditionPartNos", "")).strip()
            condition_part_nos = [item.strip() for item in condition_part_nos_raw.split(",") if item.strip()]
            choice_number = _to_float(row.get("choiceNumber"))
            group.choices.append(
                BindingChoice(
                    part_no=choice_part_no,
                    desc=choice_desc,
                    condition_mode=condition_mode,
                    condition_part_nos=condition_part_nos,
                    number=choice_number,
                )
            )
        return list(projects.values())


def _to_float(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        stripped = str(value).strip()
        if not stripped:
            return None
        return float(stripped)
    except ValueError:
        return None
