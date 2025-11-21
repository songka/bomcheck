from __future__ import annotations

import json
import re
from json import JSONDecodeError
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict


DEFAULT_CONFIG = {
    "invalid_part_db": "失效料号.xlsx",
    "binding_library": "绑定料号.js",
    "important_materials": "重要物料.txt",
    "system_part_db": "系统料号.xlsx",
    "blocked_applicants": "屏蔽申请人.txt",
    "part_asset_dir": "料号资源",
    "account_store": "accounts.json",
    "ua_lookup_dir": "",
}


@dataclass
class AppConfig:
    invalid_part_db: Path
    binding_library: Path
    important_materials: Path
    system_part_db: Path
    blocked_applicants: Path
    part_asset_dir: Path
    account_store: Path
    ua_lookup_dir: Path | None

    @classmethod
    def from_dict(cls, data: Dict[str, Any], base_dir: Path) -> "AppConfig":
        return cls(
            invalid_part_db=_resolve_path(
                data.get("invalid_part_db") or DEFAULT_CONFIG["invalid_part_db"],
                base_dir,
            ),
            binding_library=_resolve_path(
                data.get("binding_library") or DEFAULT_CONFIG["binding_library"],
                base_dir,
            ),
            important_materials=_resolve_path(
                data.get("important_materials") or DEFAULT_CONFIG["important_materials"],
                base_dir,
            ),
            system_part_db=_resolve_path(
                data.get("system_part_db") or DEFAULT_CONFIG["system_part_db"],
                base_dir,
            ),
            blocked_applicants=_resolve_path(
                data.get("blocked_applicants") or DEFAULT_CONFIG["blocked_applicants"],
                base_dir,
            ),
            part_asset_dir=_resolve_path(
                data.get("part_asset_dir") or DEFAULT_CONFIG["part_asset_dir"],
                base_dir,
            ),
            account_store=_resolve_path(
                data.get("account_store") or DEFAULT_CONFIG["account_store"],
                base_dir,
            ),
            ua_lookup_dir=_resolve_optional_path(
                data.get("ua_lookup_dir"), base_dir
            ),
        )

    def to_dict(self, base_dir: Path) -> Dict[str, str]:
        return {
            "invalid_part_db": _to_relative(self.invalid_part_db, base_dir),
            "binding_library": _to_relative(self.binding_library, base_dir),
            "important_materials": _to_relative(self.important_materials, base_dir),
            "system_part_db": _to_relative(self.system_part_db, base_dir),
            "blocked_applicants": _to_relative(self.blocked_applicants, base_dir),
            "part_asset_dir": _to_relative(self.part_asset_dir, base_dir),
            "account_store": _to_relative(self.account_store, base_dir),
            "ua_lookup_dir": _to_relative(self.ua_lookup_dir, base_dir)
            if self.ua_lookup_dir
            else "",
        }


def load_config(path: Path) -> AppConfig:
    base_dir = path.parent
    if not path.exists():
        save_config(path, AppConfig.from_dict(DEFAULT_CONFIG, base_dir))

    raw_text = path.read_text(encoding="utf-8")
    sanitized_text = _sanitize_json_text(raw_text)
    corrected = sanitized_text != raw_text

    try:
        data = json.loads(sanitized_text)
    except JSONDecodeError:
        # If we still cannot load the configuration, fall back to defaults and
        # preserve the original text for manual inspection.
        backup_path = path.with_suffix(path.suffix + ".bak")
        backup_path.write_text(raw_text, encoding="utf-8")
        data = DEFAULT_CONFIG
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


def _sanitize_json_text(raw_text: str) -> str:
    # Strip BOM, normalize newlines, remove comments, repair stray backslashes and
    # trailing commas so config files copied between machines remain loadable.
    text = raw_text.lstrip("\ufeff").replace("\r\n", "\n")
    text = _strip_json_comments(text)
    text = _escape_invalid_backslashes(text)
    return _remove_trailing_commas(text)


def _resolve_path(value: str | None, base_dir: Path) -> Path:
    if not value:
        return base_dir
    p = Path(value)
    if not p.is_absolute():
        p = base_dir / p
    return p


def _resolve_optional_path(value: str | None, base_dir: Path) -> Path | None:
    if not value:
        return None
    return _resolve_path(value, base_dir)


def _to_relative(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir))
    except ValueError:
        return str(path)


def _strip_json_comments(text: str) -> str:
    # Remove // line comments that start a line and /* block comments */ while
    # leaving inline URLs untouched.
    text = re.sub(r"(?m)^\s*//.*$", "", text)
    return re.sub(r"/\*.*?\*/", "", text, flags=re.DOTALL)


def _remove_trailing_commas(text: str) -> str:
    return re.sub(r",(\s*[}\]])", r"\1", text)
