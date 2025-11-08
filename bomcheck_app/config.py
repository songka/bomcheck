from __future__ import annotations

import json
from json import JSONDecodeError
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict


DEFAULT_CONFIG = {
    "invalid_part_db": "失效料号.xlsx",
    "binding_library": "绑定料号.js",
    "important_materials": "重要物料.txt",
}


@dataclass
class AppConfig:
    invalid_part_db: Path
    binding_library: Path
    important_materials: Path

    @classmethod
    def from_dict(cls, data: Dict[str, Any], base_dir: Path) -> "AppConfig":
        return cls(
            invalid_part_db=_resolve_path(data.get("invalid_part_db"), base_dir),
            binding_library=_resolve_path(data.get("binding_library"), base_dir),
            important_materials=_resolve_path(data.get("important_materials"), base_dir),
        )

    def to_dict(self, base_dir: Path) -> Dict[str, str]:
        return {
            "invalid_part_db": _to_relative(self.invalid_part_db, base_dir),
            "binding_library": _to_relative(self.binding_library, base_dir),
            "important_materials": _to_relative(self.important_materials, base_dir),
        }


def load_config(path: Path) -> AppConfig:
    base_dir = path.parent
    if not path.exists():
        save_config(path, AppConfig.from_dict(DEFAULT_CONFIG, base_dir))

    raw_text = path.read_text(encoding="utf-8")
    corrected = False
    try:
        data = json.loads(raw_text)
    except JSONDecodeError as error:
        sanitized_text = raw_text.replace("\\", "\\\\")
        try:
            data = json.loads(sanitized_text)
        except JSONDecodeError as secondary_error:
            raise error from secondary_error
        corrected = True

    config = AppConfig.from_dict(data, base_dir)
    if corrected:
        save_config(path, config)
    return config


def save_config(path: Path, config: AppConfig) -> None:
    base_dir = path.parent
    path.write_text(json.dumps(config.to_dict(base_dir), ensure_ascii=False, indent=2), encoding="utf-8")


def _escape_invalid_backslashes(raw_text: str) -> str:
    pattern = r"(?<!\\)\\(?![\\/\"bfnrtu])"
    return re.sub(pattern, r"\\\\", raw_text)


def _resolve_path(value: str | None, base_dir: Path) -> Path:
    if not value:
        return base_dir
    p = Path(value)
    if not p.is_absolute():
        p = base_dir / p
    return p


def _to_relative(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir))
    except ValueError:
        return str(path)
