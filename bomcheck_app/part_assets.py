from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import webbrowser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable
from urllib.parse import urlparse

import requests
from PIL import Image

from .excel_processor import normalize_part_no


@dataclass
class PartAsset:
    part_no: str
    images: list[str] = field(default_factory=list)
    model_file: str | None = None
    local_paths: list[str] = field(default_factory=list)
    remote_links: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, part_no: str, data: Dict) -> "PartAsset":
        return cls(
            part_no=data.get("part_no", part_no),
            images=list(data.get("images", []) or []),
            model_file=data.get("model_file") or None,
            local_paths=list(data.get("local_paths", []) or []),
            remote_links=list(data.get("remote_links", []) or []),
        )

    def to_dict(self) -> Dict:
        return {
            "part_no": self.part_no,
            "images": list(self.images),
            "model_file": self.model_file,
            "local_paths": list(self.local_paths),
            "remote_links": list(self.remote_links),
        }


class PartAssetStore:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)
        self.index_path = self.root / "assets.json"
        self.assets: Dict[str, PartAsset] = {}
        self._load()

    def _load(self) -> None:
        if not self.index_path.exists():
            self.assets = {}
            return
        try:
            data = json.loads(self.index_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.assets = {}
            return
        self.assets = {
            key: PartAsset.from_dict(value.get("part_no", key), value)
            for key, value in data.items()
        }

    def save(self) -> None:
        payload = {key: asset.to_dict() for key, asset in self.assets.items()}
        self.index_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def get(self, part_no: str) -> PartAsset | None:
        normalized = normalize_part_no(part_no)
        if not normalized:
            return None
        asset = self.assets.get(normalized)
        if asset:
            return asset
        return None

    def list_assets(self) -> list[PartAsset]:
        return sorted(self.assets.values(), key=lambda item: item.part_no)

    def upsert(self, asset: PartAsset) -> None:
        normalized = normalize_part_no(asset.part_no)
        if not normalized:
            raise ValueError("无效的料号")
        asset.part_no = normalized
        self.assets[normalized] = asset
        self.save()

    def remove(self, part_no: str) -> None:
        normalized = normalize_part_no(part_no)
        if not normalized:
            return
        if normalized in self.assets:
            del self.assets[normalized]
            self.save()

    def add_images(self, part_no: str, image_paths: Iterable[Path]) -> list[str]:
        asset = self._ensure(part_no)
        saved: list[str] = []
        for path in image_paths:
            destination = self._copy_to_part_folder(asset.part_no, path)
            asset.images.append(destination)
            saved.append(destination)
        self.upsert(asset)
        return saved

    def set_model_file(self, part_no: str, source: Path) -> str:
        asset = self._ensure(part_no)
        destination = self._copy_to_part_folder(asset.part_no, source)
        asset.model_file = destination
        self.upsert(asset)
        return destination

    def set_local_paths(self, part_no: str, paths: list[str]) -> None:
        asset = self._ensure(part_no)
        asset.local_paths = paths
        self.upsert(asset)

    def set_remote_links(self, part_no: str, links: list[str]) -> None:
        asset = self._ensure(part_no)
        asset.remote_links = links
        self.upsert(asset)

    def download_image(self, part_no: str, url: str) -> str:
        asset = self._ensure(part_no)
        file_name = _safe_filename(urlparse(url).path.rsplit("/", 1)[-1]) or "image"
        extension = _guess_extension(file_name)
        target = self._generate_unique_path(asset.part_no, f"{file_name}{extension}")
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        target.write_bytes(response.content)
        asset.images.append(str(target.relative_to(self.root)))
        self.upsert(asset)
        return str(target.relative_to(self.root))

    def download_first_image_from_search(self, part_no: str, keyword: str) -> str | None:
        url = "https://www.bing.com/images/search"
        response = requests.get(
            url,
            params={"q": keyword},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=15,
        )
        response.raise_for_status()
        match = re.search(r"murl\":\"(.*?)\"", response.text)
        if not match:
            return None
        image_url = match.group(1).replace("\\/", "/")
        try:
            return self.download_image(part_no, image_url)
        except Exception:
            return None

    def _ensure(self, part_no: str) -> PartAsset:
        normalized = normalize_part_no(part_no)
        if not normalized:
            raise ValueError("无效的料号")
        existing = self.assets.get(normalized)
        if existing:
            return existing
        asset = PartAsset(part_no=normalized)
        self.assets[normalized] = asset
        return asset

    def _copy_to_part_folder(self, part_no: str, source: Path) -> str:
        part_folder = self.root / part_no
        part_folder.mkdir(parents=True, exist_ok=True)
        destination = self._generate_unique_path(part_no, source.name)
        shutil.copy2(source, destination)
        return str(destination.relative_to(self.root))

    def _generate_unique_path(self, part_no: str, file_name: str) -> Path:
        part_folder = self.root / part_no
        part_folder.mkdir(parents=True, exist_ok=True)
        candidate = part_folder / file_name
        stem = Path(file_name).stem
        suffix = Path(file_name).suffix
        counter = 1
        while candidate.exists():
            candidate = part_folder / f"{stem}_{counter}{suffix}"
            counter += 1
        return candidate

    def load_image_preview(self, relative_path: str, max_size: tuple[int, int] = (420, 420)):
        image_path = self.root / relative_path
        with Image.open(image_path) as img:
            img.thumbnail(max_size)
            return img.copy()

    def resolve_path(self, relative_path: str) -> Path:
        return self.root / relative_path


def open_file(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        webbrowser.open(path.as_uri())


def _safe_filename(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]", "_", name)


def _guess_extension(name: str) -> str:
    if Path(name).suffix:
        return ""
    return ".jpg"
