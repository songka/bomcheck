from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from bomcheck_app.asset_crawler import AssetCrawler


def read_parts(source: Path) -> list[str]:
    if not source.exists():
        return []
    return [
        line.strip()
        for line in source.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="自动爬取料号的图片与官网链接，更新到料号资源库。",
    )
    parser.add_argument(
        "parts",
        nargs="*",
        help="需要处理的料号；如不提供则从 --parts-file 中读取",
    )
    parser.add_argument(
        "--parts-file",
        type=Path,
        default=Path("parts.txt"),
        help="包含料号列表的文本文件（每行一个）",
    )
    parser.add_argument(
        "--asset-root",
        type=Path,
        default=Path("料号资源"),
        help="料号资源库根目录，默认使用项目下的 料号资源/",
    )
    parser.add_argument(
        "--progress",
        type=Path,
        default=None,
        help="自定义进度文件路径，默认为资产目录下 crawl_progress.json",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="每个任务之间的等待秒数，避免频繁请求触发风控",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="本次最多处理多少个任务，便于分批执行",
    )
    parser.add_argument(
        "--ua-dir",
        type=Path,
        default=None,
        help="UA 成品资料目录，自动生成成品资源时会在该目录下查找",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    parts: Iterable[str] = args.parts or read_parts(args.parts_file)
    crawler = AssetCrawler(
        args.asset_root,
        args.progress,
        delay_seconds=args.delay,
        ua_lookup_dir=args.ua_dir,
    )
    crawler.add_tasks(parts)
    pending = crawler.pending()
    if not pending:
        print("没有待处理的料号。")
        return
    print(f"开始处理 {len(pending)} 个料号……")
    crawler.run(limit=args.limit)
    remaining = crawler.pending()
    if remaining:
        print(f"本次已处理部分任务，剩余 {len(remaining)} 个待处理，可再次运行继续。")
    else:
        print("所有任务已完成。")


if __name__ == "__main__":
    main()
