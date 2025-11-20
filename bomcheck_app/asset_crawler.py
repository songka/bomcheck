from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

from .excel_processor import normalize_part_no
from .part_assets import PartAsset, PartAssetStore


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36"
)


@dataclass
class CrawlStatus:
    part_no: str
    status: str = "pending"  # pending | done | failed
    message: str = ""

    def to_dict(self) -> Dict:
        return {
            "part_no": self.part_no,
            "status": self.status,
            "message": self.message,
        }

    @classmethod
    def from_dict(cls, data: Dict) -> "CrawlStatus":
        return cls(
            part_no=data.get("part_no", ""),
            status=data.get("status", "pending"),
            message=data.get("message", ""),
        )


class AssetCrawler:
    def __init__(
        self,
        asset_root: Path,
        progress_path: Optional[Path] = None,
        delay_seconds: float = 1.0,
    ) -> None:
        self.store = PartAssetStore(asset_root)
        self.progress_path = progress_path or (asset_root / "crawl_progress.json")
        self.delay_seconds = delay_seconds
        self._tasks: Dict[str, CrawlStatus] = {}
        self._load_progress()

    def _load_progress(self) -> None:
        if not self.progress_path.exists():
            return
        try:
            raw = json.loads(self.progress_path.read_text(encoding="utf-8"))
            for item in raw:
                status = CrawlStatus.from_dict(item)
                if status.part_no:
                    self._tasks[status.part_no] = status
        except json.JSONDecodeError:
            # 如果进度文件损坏，则忽略并重新开始
            self._tasks = {}

    def _save_progress(self) -> None:
        payload = [task.to_dict() for task in self._tasks.values()]
        self.progress_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def add_tasks(self, part_numbers: Iterable[str]) -> None:
        for part in part_numbers:
            normalized = normalize_part_no(part)
            if not normalized:
                continue
            if normalized in self._tasks and self._tasks[normalized].status == "done":
                continue
            self._tasks.setdefault(normalized, CrawlStatus(part_no=normalized))
        self._save_progress()

    def remove_tasks(self, part_numbers: Iterable[str]) -> None:
        removed = False
        for part in part_numbers:
            normalized = normalize_part_no(part)
            if normalized and normalized in self._tasks:
                del self._tasks[normalized]
                removed = True
        if removed:
            self._save_progress()

    def clear(self) -> None:
        if not self._tasks:
            return
        self._tasks = {}
        self._save_progress()

    def pending(self) -> List[str]:
        return [p for p, task in self._tasks.items() if task.status != "done"]

    def run(self, limit: Optional[int] = None) -> None:
        processed = 0
        for part_no in list(self.pending()):
            if limit is not None and processed >= limit:
                break
            status = self._tasks[part_no]
            try:
                message = self._process_part(part_no)
                status.status = "done"
                status.message = message
            except Exception as exc:  # noqa: BLE001
                status.status = "failed"
                status.message = str(exc)
            self._tasks[part_no] = status
            self._save_progress()
            processed += 1
            if self.delay_seconds:
                time.sleep(self.delay_seconds)

    def statuses(self) -> List[CrawlStatus]:
        return sorted(self._tasks.values(), key=lambda item: item.part_no)

    def summary(self) -> tuple[int, int]:
        total = len(self._tasks)
        done = len([t for t in self._tasks.values() if t.status == "done"])
        return done, total

    def _process_part(self, part_no: str) -> str:
        asset = self.store.get(part_no) or PartAsset(part_no=part_no)
        self.store.upsert(asset)

        updates: list[str] = []
        official = self._search_official_site(part_no)
        if official:
            updated_links = list(asset.remote_links)
            if official not in updated_links:
                updated_links.append(official)
                self.store.set_remote_links(part_no, updated_links)
                updates.append("官网链接")

        if not asset.images:
            image_path = self.store.download_first_image_from_search(
                part_no, f"{part_no} 产品 图片"
            )
            if image_path:
                updates.append("图片")

        if not updates:
            return "已存在资源，未更新"
        return f"已更新 {', '.join(updates)}"

    def _search_official_site(self, keyword: str) -> Optional[str]:
        response = requests.get(
            "https://www.bing.com/search",
            params={"q": f"{keyword} 官网", "setlang": "zh-cn"},
            headers={"User-Agent": USER_AGENT},
            timeout=15,
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        normalized = (normalize_part_no(keyword) or keyword).lower()
        fallback: Optional[str] = None
        for link in soup.select("li.b_algo h2 a, ol#b_results h2 a"):
            href = link.get("href")
            if not href or not self._is_http_url(href):
                continue
            if normalized in href.lower():
                return href
            if fallback is None:
                fallback = href
        return fallback

    def _is_http_url(self, url: str) -> bool:
        parsed = urlparse(url)
        return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


__all__ = ["AssetCrawler", "CrawlStatus"]
