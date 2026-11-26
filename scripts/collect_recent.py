#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究进展定时任务 · 数据抓取阶段（灵犀 / WPS 云文档版）

从 WPS 云文档拉取最近访问/编辑记录，按 file_id 去重、UTC 转北京时间、
按自然周分组，输出 progress_auto/records.json 供后续归纳使用。

无灵犀环境请改用通用版 scripts/collect_recent_fs.py（扫描本地目录）。

用法:
    python3 collect_recent.py --weeks 3
    python3 collect_recent.py --weeks 3 --out-dir /abs/path
    python3 collect_recent.py --commit-state "<生成的docx绝对路径>"
    python3 collect_recent.py --resolve-approval approved

约束:
    - 凭据由运行时会话注入的环境变量提供，本脚本不读取、不保存、不打印。
    - 统计窗口为增量：取上次成功运行时刻至本次的全部记录，失败不推进水位线。
    - 提交水位线前强校验 DOCX 存在性与命名规范（YYYY年M月D日 <中文名>研究进展.docx，
      中文名取 config.json 的 report_name_cn，示例默认「李某研究进展」）。
    - 所有路径使用绝对路径，不依赖当前工作目录。

退出码:
    2 = 凭据缺失/鉴权失效（不重试）
    3 = 接口连续失败或非瞬时错误（停止，不产出半成品）
    4 = 文档不存在或命名不合规（拒绝推进水位线）
    5 = 窗口内无任何记录（不生成空文档、不推进水位线）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime, timedelta, timezone

# 仓库根：scripts/ 的上一级
REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_config() -> dict:
    """读取仓库根的 config.json；缺失或解析失败时返回空字典（走下方的兼容回退）。"""
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

# 工作区：state.json / papers.json / 输出 DOCX 所在目录
WORKSPACE = CFG.get("workspace") or os.path.dirname(REPO_DIR)
# 灵犀（WPS）云文档 CLI；config.json 的 wps_cli 优先，
# 缺省按 macOS 标准安装位置推导（expanduser，不写死任何用户名）
CLI = CFG.get("wps_cli") or os.path.expanduser(
    "~/Library/Application Support/WPS 灵犀/"
    "serverdir/target_skills/wps-docs/scripts/wps_docs_cli.py")
BJ = timezone(timedelta(hours=8))

# 报告人姓名（用于命名硬约束与自产文件排除）
NAME_CN = CFG.get("report_name_cn") or "李某研究进展"
NAME_EN = CFG.get("report_name_en") or "LiMou_Research_Progress"

# 归入"其他"且不单独展开的噪声文件特征
NOISE_KEYS = ("验证码", "签到表", "备份码", "backup codes", "公务卡",
              "报酬", "杂志领取", "学籍", "征文大赛", "tmp", ".s2t")

# 自动化自身产出的文件：报告正文、邮件 ASCII 副本、样式模板、乱码名变体。
# 它们不是研究进展。若混入统计窗口，会被当成"本周工作"写进报告，形成自指噪声。
SELF_PRODUCED_KEYS = (NAME_CN, NAME_EN)


def _looks_mojibake(name: str) -> bool:
    """判断文件名是否为 UTF-8 被按 latin-1 误解码产生的乱码。

    例：`2026å¹´11æ26æ¥ æçæ´ç ç©¶è¿å±.docx`

    为什么不用 round-trip 校验：这类名字在传递过程中被**有损截断**（丢了部分字节），
    实测 `name.encode("latin-1").decode("utf-8")` 会抛 UnicodeDecodeError，
    所以只能改用密度启发式。

    原理：中文 UTF-8 字节集中在 0xE4–0xE9，误解码后会产生大量 U+0080–U+00FF
    范围内的字符（å ¹ ´ æ ¥ ç © ¶ è ¿ ± 等）。正常英文名和正常中文名都不会
    落入这个区间。

    取舍：带重音符号的欧洲语言名（如 `Müller.docx`，3 个）可能误判，
    但本场景是中文文档盘，概率极低。
    """
    if name.isascii():
        return False
    latin1_chars = sum(1 for ch in name if 0x80 <= ord(ch) <= 0xFF)
    return latin1_chars >= 3


def is_self_produced(name: str) -> bool:
    """是否为自动化自身产出的文件（不进入统计窗口）。"""
    return any(k in name for k in SELF_PRODUCED_KEYS) or _looks_mojibake(name)


def bj_time(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(BJ)


def week_monday(dt: datetime) -> datetime:
    d = dt.astimezone(BJ)
    return (d - timedelta(days=d.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0)


# 网络/token 换取的瞬时错误：允许有限重试；鉴权类错误：立即停止不重试
_TRANSIENT = ("eof", "request failed", "timeout", "timed out",
              "connection reset", "persist failed", "获取 token 失败",
              "token 已换取")
# 仅匹配明确的鉴权错误短语；避免裸数字误命中正常数据中的时间戳/体积
_AUTH_FATAL = ("缺少认证凭证", "unauthorized", "invalid token",
               "token 无效", "authentication failed", "http 401", "http 403")


def _call_cli(cmd: list[str], attempts: int = 3) -> dict:
    """调用 CLI 并解析 JSON；仅对瞬时错误做有限重试。"""
    env = os.environ.copy()  # 显式继承，确保会话注入的凭据完整传递
    last = ""
    for n in range(1, attempts + 1):
        proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
        out = (proc.stdout or "").strip()
        err = (proc.stderr or "").strip()
        payload = None
        try:
            payload = json.loads(out)
        except json.JSONDecodeError:
            pass
        # 先判定成功返回，直接返回，不对数据正文做错误关键字匹配
        if proc.returncode == 0 and isinstance(payload, dict) and \
           payload.get("success", True) is not False:
            return payload
        # 仅失败响应才提取错误消息做判断
        err_msg = ""
        if isinstance(payload, dict):
            err_msg = str(payload.get("error") or "")
        low = (err_msg + " " + err).lower()
        last = out or err
        # 瞬时特征优先：token 换取失败（EOF）会伪装成 HTTP 401，
        # 这类可恢复错误必须重试，不能误判为凭据失效
        transient = any(k in low for k in _TRANSIENT)
        if transient:
            if n < attempts:
                print(f"[WARN] 第 {n} 次调用遇瞬时错误，3s 后重试：{err_msg[:160]}",
                      file=sys.stderr)
                time.sleep(3)
                continue
            print(f"[FATAL] 云文档接口连续 {attempts} 次瞬时失败，停止执行：{err_msg[:300]}",
                  file=sys.stderr)
            sys.exit(3)
        if any(k in low for k in _AUTH_FATAL):
            print("[FATAL] 云文档鉴权失败：会话凭据无效或未注入，停止执行（不重试）。",
                  file=sys.stderr)
            sys.exit(2)
        print(f"[FATAL] 云文档接口调用失败（非瞬时错误，不重试）：{last[:400]}",
              file=sys.stderr)
        sys.exit(3)


def fetch_all(max_pages: int = 20) -> list[dict]:
    items: list[dict] = []
    token = None
    for _ in range(max_pages):
        cmd = [sys.executable, CLI, "drive", "list-latest-items",
               "--page-size", "200"]
        if token:
            cmd += ["--page-token", token]
        payload = _call_cli(cmd)
        data = payload.get("data", payload)
        items.extend(data.get("items") or [])
        token = data.get("next_page_token")
        if not token:
            break
        time.sleep(1)  # 翻页间隔，降低 token 换取接口瞬时失败概率
    return items


STATE_NAME = "state.json"

# 命名硬约束：存档目录历史文档全部遵循此格式，检索以文档名为准
# YYYY年M月D日 <中文名>研究进展.docx（月/日不补零，日期与姓名间一个半角空格，无后缀/序号）
DOCX_NAME_RE = re.compile(
    r'^\d{4}年(?:[1-9]|1[0-2])月(?:[1-9]|[12]\d|3[01])日 '
    + re.escape(NAME_CN) + r'\.docx$')


def validate_docx(docx_path: str) -> None:
    """提交水位线前强校验：文档必须存在且命名符合规范。"""
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
    """文档成功生成后才调用：推进水位线并标记待审批。"""
    validate_docx(docx_path)  # 先校验，不合规绝不推进水位线
    rec_path = os.path.join(out_dir, "records.json")
    with open(rec_path, encoding="utf-8") as f:
        rec = json.load(f)
    # 空窗口不生成文档、不推进水位线，避免产出空进展并跳过后续可补数据
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
    pa["resolved_at"] = datetime.now(BJ).isoformat()
    # 拒绝时必须回退水位线到该期窗口起点，否则这批记录会被永久跳过
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

    # 仅状态维护模式，不访问云文档
    if args.commit_state:
        commit_state(args.out_dir, args.commit_state)
        return
    if args.resolve_approval:
        clear_approval(args.out_dir, args.resolve_approval)
        return

    if not os.environ.get("TMP_LX_UUID") and not os.environ.get("WPS_SID"):
        print("[FATAL] 未检测到会话凭据环境变量，停止执行（不读取、不保存凭据）。",
              file=sys.stderr)
        sys.exit(2)

    now = datetime.now(BJ)
    state = read_state(args.out_dir)
    window_mode = "incremental"
    if args.since:
        start = datetime.fromisoformat(args.since).astimezone(BJ)
        window_mode = "explicit"
    elif state.get("last_success_at"):
        start = datetime.fromisoformat(state["last_success_at"]).astimezone(BJ)
    else:
        start = week_monday(now) - timedelta(weeks=args.weeks - 1)
        window_mode = "first_run_fallback"
    items = fetch_all()

    # 按 file_id 去重，保留最新 mtime
    by_id: dict[str, dict] = {}
    for it in items:
        fid = it.get("file_id")
        if not fid or not it.get("mtime"):
            continue
        prev = by_id.get(fid)
        if prev is None or bj_time(it["mtime"]) > bj_time(prev["mtime"]):
            by_id[fid] = it

    # 同名副本（我的云文档 / 我的漫游箱）合并看待
    name_counts = Counter(i.get("name", "") for i in by_id.values())
    dup_names = {n for n, c in name_counts.items() if c > 1}

    records = []
    self_produced: list[str] = []
    for fid, it in by_id.items():
        t = bj_time(it["mtime"])
        if t < start:
            continue
        name = it.get("name", "")
        if is_self_produced(name):
            self_produced.append(name)
            continue
        records.append({
            "file_id": fid,
            "name": name,
            # 接口实际字段为 file_src.name（我的云文档 / 我的漫游箱）
            "drive": (it.get("file_src") or {}).get("name", ""),
            "ctime_bj": (bj_time(it["ctime"]).isoformat()
                          if it.get("ctime") else ""),
            "mtime_bj": t.isoformat(),
            "link": it.get("link_url", ""),
            "week_monday": week_monday(t).strftime("%Y-%m-%d"),
            "duplicate_name": name in dup_names,
            "noise": any(k in name for k in NOISE_KEYS),
            # list-latest-items 不返回作者字段，保留占位供后续扩展
            "created_by": "",
            "modified_by": "",
        })
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
            "total_raw": len(items),
            "total_unique": len(by_id),
            "total_in_window": len(records),
            "excluded_self_produced": len(self_produced),
            "self_produced_names": sorted(set(self_produced)),
            "week_counts": {k: len(v) for k, v in sorted(weeks.items())},
            "records": records,
        }, f, ensure_ascii=False, indent=2)

    print(f"[OK] 窗口({window_mode}) {start:%Y-%m-%d %H:%M} ~ {now:%Y-%m-%d %H:%M}，"
          f"原始 {len(items)} 条，去重后 {len(by_id)} 个文件，"
          f"窗口内 {len(records)} 个文件")
    if self_produced:
        print(f"[SKIP] 已排除 {len(self_produced)} 个自产文件（报告/邮件副本/模板/乱码名），"
              f"不计入研究进展：")
        for n in sorted(set(self_produced)):
            print(f"         · {n}")
    print("[NOTE] 水位线未推进；文档成功生成后再用 --commit-state 推进")
    for k, v in sorted(weeks.items()):
        print(f"  {k} 起：{len(v)} 个文件")
    print(f"[OK] 已写入 {out_path}")


if __name__ == "__main__":
    main()
