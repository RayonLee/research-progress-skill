#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究进展定时任务 · 数据抓取阶段（通用本地文件系统版）

不依赖任何云文档平台：递归扫描 config.json 的 source_dirs（本地目录），
按文件 mtime 取增量窗口内的记录，输出与灵犀版 collect_recent.py **完全相同**
的 progress_auto/records.json 结构，供同一份归纳/渲染流程使用。

与灵犀版的分工：
    - 数据源在云文档（WPS / 灵犀）     → scripts/collect_recent.py
    - 数据源在本地目录（任意平台可用） → scripts/collect_recent_fs.py

水位线（state.json 的 last_success_at）与审批状态与灵犀版**共用同一份**：
--commit-state / --resolve-approval 的参数与语义完全一致，两个抓取脚本可以
混用（比如平时用云文档版，断网/换机时用本地版）。

事实边界与灵犀版相同：mtime 只能说明「文件被修改过」，无法区分本人或协作者，
也无法断定「当天新写」；归纳时禁止把推断写成既成事实。

用法:
    python3 collect_recent_fs.py                # 用水位线增量扫描
    python3 collect_recent_fs.py --since "2026-11-09T00:00:00+08:00"
    python3 collect_recent_fs.py --commit-state "<生成的docx绝对路径>"
    python3 collect_recent_fs.py --resolve-approval approved

退出码:
    1 = 参数/配置错误（source_dirs 未配置等）
    4 = 文档不存在或命名不合规（拒绝推进水位线）
    5 = 窗口内无任何记录（不生成空文档、不推进水位线）
    （本脚本不接触网络，无 2/3 鉴权与接口错误码）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, timedelta

# 仓库根：scripts/ 的上一级
REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_config() -> dict:
    path = os.path.join(REPO_DIR, "config.json")
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"[WARN] config.json 解析失败，使用内置默认值：{e}", file=sys.stderr)
        return {}


CFG = _load_config()

WORKSPACE = CFG.get("workspace") or os.path.dirname(REPO_DIR)
LOCAL_TZ = datetime.now().astimezone().tzinfo

NAME_CN = CFG.get("report_name_cn") or "李某研究进展"
NAME_EN = CFG.get("report_name_en") or "LiMou_Research_Progress"

# 扫描的本地目录（在 config.json 的 source_dirs 里配置，可多个）
SOURCE_DIRS: list[str] = [os.path.expanduser(d) for d in
                          (CFG.get("source_dirs") or [])]

# 噪声与自产排除规则与灵犀版保持一致
NOISE_KEYS = ("验证码", "签到表", "备份码", "backup codes", "公务卡",
              "报酬", "杂志领取", "学籍", "征文大赛", "tmp", ".s2t")
SELF_PRODUCED_KEYS = (NAME_CN, NAME_EN)


def _looks_mojibake(name: str) -> bool:
    """文件名是否为 UTF-8 被按 latin-1 误解码的乱码（密度启发式）。"""
    if name.isascii():
        return False
    return sum(1 for ch in name if 0x80 <= ord(ch) <= 0xFF) >= 3


def is_self_produced(name: str) -> bool:
    return any(k in name for k in SELF_PRODUCED_KEYS) or _looks_mojibake(name)


def week_monday(dt: datetime) -> datetime:
    d = dt.astimezone(LOCAL_TZ)
    return (d - timedelta(days=d.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0)


STATE_NAME = "state.json"

# 命名硬约束与灵犀版一致（月/日不补零，日期与姓名间一个半角空格）
DOCX_NAME_RE = re.compile(
    r'^\d{4}年(?:[1-9]|1[0-2])月(?:[1-9]|[12]\d|3[01])日 '
    + re.escape(NAME_CN) + r'\.docx$')


def validate_docx(docx_path: str) -> None:
    name = os.path.basename(docx_path)
    if not os.path.isfile(docx_path):
        print(f"[FATAL] 文档不存在，拒绝推进水位线：{docx_path}", file=sys.stderr)
        sys.exit(4)
    if not DOCX_NAME_RE.match(name):
        print("[FATAL] 文档命名不符合规范，拒绝推进水位线：" + name,
              file=sys.stderr)
        print(f"[FATAL] 须为 YYYY年M月D日 {NAME_CN}.docx"
              "（月日不补零、一个半角空格、无后缀/序号）", file=sys.stderr)
        sys.exit(4)


def read_state(out_dir: str) -> dict:
    p = os.path.join(out_dir, STATE_NAME)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def write_state(out_dir: str, state: dict) -> None:
    os.makedirs(out_dir, exist_ok=True)
    p = os.path.join(out_dir, STATE_NAME)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)  # 原子写，避免半截状态


def commit_state(out_dir: str, docx_path: str) -> None:
    """文档成功生成后才调用：推进水位线并标记待审批（与灵犀版同语义）。"""
    validate_docx(docx_path)
    rec_path = os.path.join(out_dir, "records.json")
    with open(rec_path, encoding="utf-8") as f:
        rec = json.load(f)
    if not rec.get("total_in_window"):
        print("[FATAL] 本窗口无任何记录，不生成文档、不推进水位线。",
              file=sys.stderr)
        sys.exit(5)
    state = read_state(out_dir)
    state["last_success_at"] = rec["window_end"]
    state["last_output"] = docx_path
    state["pending_approval"] = {
        "docx_path": docx_path,
        "window_start": rec["window_start"],
        "window_end": rec["window_end"],
        "file_count": rec["total_in_window"],
        "status": "pending",
    }
    write_state(out_dir, state)
    print(f"[OK] 水位线已推进至 {rec['window_end']}，文档标记待审批")


def clear_approval(out_dir: str, action: str) -> None:
    """审批阶段上传/发送完成后调用，清除待审批标记。"""
    state = read_state(out_dir)
    pa = state.get("pending_approval")
    if not pa or pa.get("status") != "pending":
        print("[WARN] 当前无待审批文档，未做改动")
        return
    pa["status"] = "approved" if action == "approved" else "rejected"
    pa["resolved_at"] = datetime.now(LOCAL_TZ).isoformat()
    if action == "rejected":
        ws = pa.get("window_start")
        if ws:
            state["last_success_at"] = ws
            pa["watermark_rolled_back_to"] = ws
    state["pending_approval"] = pa
    write_state(out_dir, state)
    extra = ("；水位线已回退至 " + pa["watermark_rolled_back_to"][:19]
             if action == "rejected" and pa.get("watermark_rolled_back_to") else "")
    print(f"[OK] 审批状态已更新为 {pa['status']}{extra}")


def scan_dirs(start: datetime) -> list[dict]:
    """递归扫描 source_dirs，返回 mtime 在窗口内的文件记录。"""
    out = []
    for root in SOURCE_DIRS:
        if not os.path.isdir(root):
            print(f"[WARN] source_dir 不存在，跳过：{root}", file=sys.stderr)
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            # 跳过隐藏目录与常见依赖/输出目录
            dirnames[:] = [d for d in dirnames
                           if not d.startswith(".")
                           and d not in ("node_modules", "__pycache__",
                                         "venv", ".venv", "dist", "build")]
            for fn in filenames:
                p = os.path.join(dirpath, fn)
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                mtime = datetime.fromtimestamp(st.st_mtime, LOCAL_TZ)
                if mtime < start:
                    continue
                ctime = datetime.fromtimestamp(st.st_ctime, LOCAL_TZ)
                out.append({
                    "file_id": os.path.abspath(p),  # 本地路径即唯一 id
                    "name": fn,
                    "drive": os.path.relpath(dirpath, root),
                    "ctime_bj": ctime.isoformat(),
                    "mtime_bj": mtime.isoformat(),
                    "link": "file://" + os.path.abspath(p),
                    "size": st.st_size,
                })
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--weeks", type=int, default=3,
                    help="首次运行（无水位线）的兜底窗口，按自然周计，默认 3")
    ap.add_argument("--since", default=None,
                    help="显式指定窗口起点（ISO 8601，带时区），优先于水位线")
    ap.add_argument("--out-dir", default=os.path.join(WORKSPACE, "progress_auto"))
    ap.add_argument("--commit-state", default=None,
                    help="文档成功生成后传入其绝对路径，推进水位线并标记待审批")
    ap.add_argument("--resolve-approval", choices=["approved", "rejected"],
                    default=None, help="审批阶段结束时更新审批状态")
    args = ap.parse_args()

    if args.commit_state:
        commit_state(args.out_dir, args.commit_state)
        return
    if args.resolve_approval:
        clear_approval(args.out_dir, args.resolve_approval)
        return

    if not SOURCE_DIRS:
        print("[FATAL] config.json 未配置 source_dirs，无法扫描本地目录。",
              file=sys.stderr)
        sys.exit(1)

    now = datetime.now(LOCAL_TZ)
    state = read_state(args.out_dir)
    window_mode = "incremental"
    if args.since:
        start = datetime.fromisoformat(args.since).astimezone(LOCAL_TZ)
        window_mode = "explicit"
    elif state.get("last_success_at"):
        start = datetime.fromisoformat(state["last_success_at"]).astimezone(LOCAL_TZ)
    else:
        start = week_monday(now) - timedelta(weeks=args.weeks - 1)
        window_mode = "first_run_fallback"

    items = scan_dirs(start)

    # 同名副本合并看待（同目录同名才算副本）
    name_counts = Counter(i.get("name", "") for i in items)
    dup_names = {n for n, c in name_counts.items() if c > 1}

    records = []
    self_produced: list[str] = []
    for it in items:
        name = it["name"]
        if is_self_produced(name):
            self_produced.append(name)
            continue
        t = datetime.fromisoformat(it["mtime_bj"])
        rec = dict(it)
        rec["week_monday"] = week_monday(t).strftime("%Y-%m-%d")
        rec["duplicate_name"] = name in dup_names
        rec["noise"] = any(k in name for k in NOISE_KEYS)
        rec["created_by"] = ""
        rec["modified_by"] = ""
        records.append(rec)
    records.sort(key=lambda r: r["mtime_bj"], reverse=True)

    weeks = {}
    for r in records:
        weeks.setdefault(r["week_monday"], []).append(r)

    os.makedirs(args.out_dir, exist_ok=True)
    out_path = os.path.join(args.out_dir, "records.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({
            "generated_at": now.isoformat(),
            "window_start": start.isoformat(),
            "window_end": now.isoformat(),
            "window_mode": window_mode,
            "window_start_source": "last_success_at" if window_mode == "incremental" else window_mode,
            "source": "local_fs",
            "total_raw": len(items),
            "total_unique": len(items),
            "total_in_window": len(records),
            "excluded_self_produced": len(self_produced),
            "self_produced_names": sorted(set(self_produced)),
            "week_counts": {k: len(v) for k, v in sorted(weeks.items())},
            "records": records,
        }, f, ensure_ascii=False, indent=2)

    print(f"[OK] 窗口({window_mode}) {start:%Y-%m-%d %H:%M} ~ {now:%Y-%m-%d %H:%M}，"
          f"窗口内 {len(records)} 个文件")
    if self_produced:
        print(f"[SKIP] 已排除 {len(self_produced)} 个自产文件，不计入研究进展：")
        for n in sorted(set(self_produced)):
            print(f"         · {n}")
    print("[NOTE] 水位线未推进；文档成功生成后再用 --commit-state 推进")
    for k, v in sorted(weeks.items()):
        print(f"  {k} 起：{len(v)} 个文件")
    print(f"[OK] 已写入 {out_path}")


if __name__ == "__main__":
    main()
