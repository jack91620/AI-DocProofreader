#!/usr/bin/env python
"""离线导入外部术语词典到 cs_terms.yaml。

为什么离线
==========

全国科学技术名词审定委员会的官方平台 https://www.cnterm.cn/ 没有公开 API；
直接抓站违反对方使用条款。务实做法是：用户从该平台手工检索、导出 CSV，
然后用本脚本合并到本地 ``cs_terms.yaml``。

支持的输入格式
==============

CSV / TSV，UTF-8。最少 2 列，最多 4 列：

    english,chinese,alternatives,section
    queue,队列,伫列;队伍,competing_translations
    robustness,鲁棒性,健壮性;稳健性,competing_translations
    operating_system,操作系统,操做系统,regional_variants

列说明
- english          外文（仅 competing_translations 段使用，其它段可留空）
- chinese          推荐中文译名（必填）
- alternatives     备选译名 / 港台变体，多个用 ; 分隔（可空）
- section          ``regional_variants`` 或 ``competing_translations``
                   ；缺省时按 alternatives 是否含 CJK 推断（含 → regional）

用法
====

    # 干跑（只看会增/改什么，不写文件）
    uv run python scripts/import_terms.py path/to/external.csv --dry-run

    # 实际合并到默认 cs_terms.yaml
    uv run python scripts/import_terms.py path/to/external.csv

    # 输出到自定义文件
    uv run python scripts/import_terms.py external.csv -o data/terminology/extra.yaml
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DICT = REPO_ROOT / "data" / "terminology" / "cs_terms.yaml"


def detect_dialect(path: Path) -> csv.Dialect | type[csv.Dialect]:
    """根据扩展名/内容简单判定分隔符。"""
    if path.suffix.lower() == ".tsv":
        return csv.excel_tab
    return csv.excel


def has_cjk(s: str) -> bool:
    return any("一" <= c <= "鿿" for c in s)


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"regional_variants": {}, "competing_translations": {}, "english_should_be_kept": []}
    with path.open(encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    data.setdefault("regional_variants", {})
    data.setdefault("competing_translations", {})
    data.setdefault("english_should_be_kept", [])
    return data


def merge_row(
    data: dict[str, Any],
    *,
    english: str,
    chinese: str,
    alternatives: list[str],
    section: str,
    overwrite: bool,
) -> str:
    """把一行合并入 data。返回操作描述（用于 dry-run 输出）。"""
    if section == "competing_translations":
        if not english:
            return f"  [skip] no english key for '{chinese}' in competing"
        bucket = data["competing_translations"]
        existing = bucket.get(english) or []
        # 合并：推荐译名永远在第一位；备选合并去重保序
        merged = [chinese]
        for alt in (*existing, *alternatives):
            if alt and alt not in merged:
                merged.append(alt)
        if existing == merged:
            return f"  [same] competing/{english}: {merged}"
        bucket[english] = merged
        verb = "update" if english in bucket and not overwrite else "add"
        return f"  [{verb}] competing/{english}: {existing} → {merged}"

    # regional_variants 段
    bucket = data["regional_variants"]
    existing_alts = list(bucket.get(chinese) or [])
    new_alts = [a for a in alternatives if a and a not in existing_alts]
    if not new_alts:
        return f"  [same] regional/{chinese}: {existing_alts}"
    merged = existing_alts + new_alts
    bucket[chinese] = merged
    verb = "update" if existing_alts else "add"
    return f"  [{verb}] regional/{chinese}: {existing_alts} → {merged}"


def import_csv(
    csv_path: Path,
    out_path: Path,
    *,
    dry_run: bool = False,
    overwrite: bool = False,
) -> int:
    data = load_yaml(out_path)

    with csv_path.open(encoding="utf-8") as f:
        reader = csv.reader(f, dialect=detect_dialect(csv_path))
        try:
            header = next(reader)
        except StopIteration:
            print(f"❌ 空 CSV: {csv_path}", file=sys.stderr)
            return 1

        # 容错：如果第一行看起来像数据而不是 header（不含 'english' / 'chinese'）
        # 则把它当数据处理
        looks_like_header = any(
            h.strip().lower() in {"english", "en", "chinese", "zh", "section"} for h in header
        )
        rows = list(reader)
        if not looks_like_header:
            rows = [header, *rows]

    if not rows:
        print("❌ 没有可导入的数据行", file=sys.stderr)
        return 1

    actions: list[str] = []
    for row in rows:
        # 补齐到 4 列，尾部空字符串
        cells = (row + [""] * 4)[:4]
        english, chinese, alts_field, section = (c.strip() for c in cells)

        if not chinese:
            actions.append(f"  [skip] empty chinese in row: {row!r}")
            continue

        alternatives = [a.strip() for a in alts_field.split(";") if a.strip()]

        if not section:
            # 推断：alternatives 含 CJK ⇒ regional；否则 competing
            section = (
                "regional_variants"
                if alternatives and any(has_cjk(a) for a in alternatives)
                else "competing_translations"
            )

        actions.append(
            merge_row(
                data,
                english=english,
                chinese=chinese,
                alternatives=alternatives,
                section=section,
                overwrite=overwrite,
            )
        )

    print(f"📋 导入 {len(rows)} 行：")
    for a in actions:
        print(a)

    if dry_run:
        print("🛑 dry-run 模式，未写入文件")
        return 0

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        yaml.safe_dump(
            data, f, allow_unicode=True, sort_keys=False, default_flow_style=False, width=120
        )
    print(f"✅ 写入 {out_path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("csv", type=Path, help="外部 CSV / TSV 文件路径")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_DICT,
        help=f"输出 yaml（默认 {DEFAULT_DICT}）",
    )
    parser.add_argument("--dry-run", action="store_true", help="只看会变什么，不写文件")
    parser.add_argument("--overwrite", action="store_true", help="存在同 key 时覆盖而非合并")

    args = parser.parse_args(argv)
    if not args.csv.exists():
        print(f"❌ 找不到 CSV: {args.csv}", file=sys.stderr)
        return 1

    return import_csv(args.csv, args.output, dry_run=args.dry_run, overwrite=args.overwrite)


if __name__ == "__main__":
    sys.exit(main())
