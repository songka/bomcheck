"""Configuration management utilities for the BOM check application."""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

DEFAULT_CONFIG = {
    "invalid_database": "失效料号.xlsx",
    "binding_database": "绑定料号.js",
    "important_materials": "重要物料.txt",
}


@dataclass
class AppConfig:
    """Data container for application configuration values."""

    invalid_database: Path
    binding_database: Path
    important_materials: Path

    @classmethod
    def from_mapping(cls, mapping: Dict[str, str], base_dir: Path | None = None) -> "AppConfig":
        base_dir = base_dir or Path.cwd()
        return cls(
            invalid_database=_resolve_path(mapping.get("invalid_database"), base_dir),
            binding_database=_resolve_path(mapping.get("binding_database"), base_dir),
            important_materials=_resolve_path(mapping.get("important_materials"), base_dir),
        )

    def to_mapping(self, base_dir: Path | None = None) -> Dict[str, str]:
        base_dir = base_dir or Path.cwd()
        return {
            "invalid_database": _relativize(self.invalid_database, base_dir),
            "binding_database": _relativize(self.binding_database, base_dir),
            "important_materials": _relativize(self.important_materials, base_dir),
        }


class ConfigManager:
    """Helper to load and store :class:`AppConfig` instances."""

    def __init__(self, config_path: Path):
        self._config_path = config_path
        self._config_path.parent.mkdir(parents=True, exist_ok=True)

    @property
    def config_path(self) -> Path:
        return self._config_path

    def load(self) -> AppConfig:
        if not self._config_path.exists():
            self.save(AppConfig.from_mapping(DEFAULT_CONFIG, self._config_path.parent))
        with self._config_path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return AppConfig.from_mapping(data, self._config_path.parent)

    def save(self, config: AppConfig) -> None:
        data = config.to_mapping(self._config_path.parent)
        with self._config_path.open("w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)


def _resolve_path(value: str | None, base_dir: Path) -> Path:
    if not value:
        return base_dir
    path = Path(value)
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def _relativize(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir))
    except ValueError:
        return str(path)
