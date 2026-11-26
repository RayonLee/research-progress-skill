#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用本地存档脚本（不依赖任何云平台）

把通过复核的 DOCX 复制到 config.json 的 archive_dir（本地归档目录）。
这是「云端存档」的最小通用实现；有具体云平台时（WPS / 灵犀 / Google Drive /
OneDrive 等）应改用对应 CLI，但同名冲突语义保持一致：

    **同名文件已存在 → 停止并报告，不覆盖、不加序号**，
    由人决定是覆盖原稿还是顺延日期。

用法:
    python3 archive_local.py "<YYYY年M月D日 李某研究进展.docx>"

退出码:
    0 = 成功
    1 = 配置错误（archive_dir 未配置）
    4 = 源文件缺失/命名不合规/目标已有同名文件（拒绝存档）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys

REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_config() -> dict:
    path = os.path.join(REPO_DIR, "config.json")
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


CFG = _load_config()
NAME_CN = CFG.get("report_name_cn") or "李某研究进展"

DOCX_NAME_RE = re.compile(
    r'^\d{4}年(?:[1-9]|1[0-2])月(?:[1-9]|[12]\d|3[01])日 '
    + re.escape(NAME_CN) + r'\.docx$')


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", help="通过复核的 DOCX 绝对路径")
    args = ap.parse_args()

    archive_dir = CFG.get("archive_dir")
    if not archive_dir:
        print("[FATAL] config.json 未配置 archive_dir（本地归档目录）。",
              file=sys.stderr)
        sys.exit(1)
    archive_dir = os.path.expanduser(archive_dir)

    if not os.path.isfile(args.docx):
        print(f"[FATAL] 源文件不存在：{args.docx}", file=sys.stderr)
        sys.exit(4)
    name = os.path.basename(args.docx)
    if not DOCX_NAME_RE.match(name):
        print(f"[FATAL] 命名不合规，拒绝存档：{name}", file=sys.stderr)
        print(f"[FATAL] 须为 YYYY年M月D日 {NAME_CN}.docx"
              "（月日不补零、一个半角空格、无后缀/序号）", file=sys.stderr)
        sys.exit(4)

    os.makedirs(archive_dir, exist_ok=True)
    dst = os.path.join(archive_dir, name)
    if os.path.exists(dst):
        print(f"[FATAL] 归档目录已存在同名文件，停止，不覆盖、不加序号：{dst}",
              file=sys.stderr)
        print("[FATAL] 由人决定：覆盖原稿，或改期生成。", file=sys.stderr)
        sys.exit(4)

    shutil.copy2(args.docx, dst)
    print(f"[OK] 已存档：{dst}")


if __name__ == "__main__":
    main()
