#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究进展周报 · 通用 SMTP 发信脚本（不依赖任何特定 agent 平台）

任何能跑 Python 的环境都能用：标准库 smtplib / imaplib，无第三方依赖。
凭据只从 config.json 或环境变量读取，**绝不写入代码、日志与产物**。

铁律在脚本层强制执行（不合规直接拒绝发送，退出码 4）：
    1. 邮件附件名必须纯 ASCII，格式 YYYY-MM-DD_<英文名>_Research_Progress.docx
    2. 邮件主题 == 附件名去掉 .docx，逐字符相等
    3. 首次使用先 --self-test：只发自己、不带抄送，验证主题/正文/附件名/可打开
    4. 正式发送后 --verify-sent 回读已发文件夹核对主题与附件名

用法:
    # 预览自检（只发自己，不带 cc）
    python3 send_mail_smtp.py --docx "<YYYY-MM-DD_..._Research_Progress.docx>" \
        --self-test

    # 正式发送
    python3 send_mail_smtp.py --docx "<...>.docx" --to advisor@example.com \
        --cc partner@example.com --body body.txt

    # 发送并回读校验（需在 config.json 配置 imap_*）
    python3 send_mail_smtp.py --docx "<...>.docx" --to ... --verify-sent

凭据优先级：命令行 > 环境变量 > config.json
    环境变量：SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASSWORD
              MAIL_FROM / MAIL_TO / MAIL_CC / IMAP_HOST

退出码:
    0 = 成功
    1 = 参数/配置错误（缺 host/user/password 等）
    4 = 附件命名不合规或主题不一致（拒绝发送）
    5 = 回读校验未找到已发邮件（可能是 IMAP 未配置已发箱映射，需人工查收）
"""
from __future__ import annotations

import argparse
import imaplib
import json
import os
import re
import smtplib
import sys
from email.header import Header
from email.message import EmailMessage
from email.utils import formataddr

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
NAME_EN = CFG.get("report_name_en") or "LiMou_Research_Progress"

# 附件名硬约束：纯 ASCII，YYYY-MM-DD_<英文名>.docx
ATTACH_RE = re.compile(
    r'^\d{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01])_'
    + re.escape(NAME_EN) + r'\.docx$')


def fail(msg: str, code: int) -> None:
    print(f"[FATAL] {msg}", file=sys.stderr)
    sys.exit(code)


def validate(docx_path: str, subject: str) -> None:
    name = os.path.basename(docx_path)
    if not os.path.isfile(docx_path):
        fail(f"附件不存在：{docx_path}", 4)
    if not ATTACH_RE.match(name):
        fail(f"附件命名不合规：{name}\n"
             f"       须为 YYYY-MM-DD_{NAME_EN}.docx（纯 ASCII，月日补零）\n"
             f"       从中文存档名复制改名：cp \"<中文名>.docx\" \"{name}\"", 4)
    if subject != name[:-len(".docx")]:
        fail(f"邮件主题与附件名不一致：\n"
             f"       主题={subject!r}\n"
             f"       附件名去后缀={name[:-len('.docx')]!r}\n"
             f"       两者必须逐字符相等", 4)
    if not name.isascii():
        fail(f"附件名含非 ASCII 字符：{name}", 4)


def build_message(*, mail_from: str, to: list[str], cc: list[str],
                  subject: str, body: str, docx_path: str) -> EmailMessage:
    msg = EmailMessage()
    # 主题仍按 MIME 规范编码（非 ASCII 主题也安全），但铁律要求主题为 ASCII
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((str(Header(os.path.basename(mail_from), "utf-8")),
                              mail_from))
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.set_content(body)
    with open(docx_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument."
                                   "wordprocessingml.document",
                           filename=os.path.basename(docx_path))
    return msg


def send(*, host: str, port: int, user: str, password: str,
         msg: EmailMessage, rcpts: list[str]) -> str:
    with smtplib.SMTP_SSL(host, port, timeout=30) as s:
        s.login(user, password)
        s.send_message(msg, to_addrs=rcpts)
    return msg["Message-ID"]


def verify_sent(imap_host: str, user: str, password: str, subject: str,
                timeout: int = 30) -> bool:
    """在已发文件夹里按主题找到刚发的邮件。找不到返回 False（退出码 5）。"""
    try:
        with imaplib.IMAP4_SSL(imap_host, timeout=timeout) as m:
            m.login(user, password)
            for folder in ('"Sent Messages"', '"已发送"', "Sent", "Sent Items"):
                try:
                    typ, _ = m.select(folder, readonly=True)
                except imaplib.IMAP4.error:
                    continue
                if typ != "OK":
                    continue
                typ, data = m.search(None, "SUBJECT", f'"{subject}"')
                if typ == "OK" and data and data[0]:
                    print(f"[OK] 已发文件夹 {folder} 中回读到主题匹配的邮件")
                    return True
        return False
    except (imaplib.IMAP4.error, OSError) as e:
        print(f"[WARN] 回读校验未完成：{e}", file=sys.stderr)
        return False


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True, help="ASCII 命名的附件绝对路径")
    ap.add_argument("--to", default=os.environ.get("MAIL_TO")
                    or CFG.get("mail_to"),
                    help="收件人，多个用逗号分隔")
    ap.add_argument("--cc", default=os.environ.get("MAIL_CC")
                    or CFG.get("mail_cc") or "",
                    help="抄送，多个用逗号分隔")
    ap.add_argument("--subject", default=None,
                    help="主题；缺省取附件名去 .docx（推荐，保证逐字符相等）")
    ap.add_argument("--body", default=None, help="正文纯文本文件路径")
    ap.add_argument("--greeting", default=None,
                    help="正文开头的称呼行，如「X老师，您好：」；与 --body 二选一")
    ap.add_argument("--self-test", action="store_true",
                    help="预览自检：只发自己（发件人），不带抄送")
    ap.add_argument("--verify-sent", action="store_true",
                    help="发送后在已发文件夹回读校验主题")
    args = ap.parse_args()

    host = os.environ.get("SMTP_HOST") or CFG.get("smtp_host")
    port = int(os.environ.get("SMTP_PORT") or CFG.get("smtp_port") or 465)
    user = os.environ.get("SMTP_USER") or CFG.get("smtp_user")
    password = os.environ.get("SMTP_PASSWORD") or CFG.get("smtp_password")
    mail_from = os.environ.get("MAIL_FROM") or CFG.get("mail_from") or user
    imap_host = os.environ.get("IMAP_HOST") or CFG.get("imap_host")

    if not host or not user or not password or not mail_from:
        fail("SMTP 配置不完整。请配置 config.json 的 smtp_host/smtp_port/"
             "smtp_user/smtp_password/mail_from，或设置对应环境变量。", 1)

    name = os.path.basename(args.docx)
    subject = args.subject if args.subject is not None else name[:-len(".docx")]
    validate(args.docx, subject)

    if args.body:
        with open(args.body, encoding="utf-8") as f:
            body = f.read()
    elif args.greeting is not None:
        body = (args.greeting.rstrip() + "\n\n"
                + "（正文请在此粘贴本周进展概述与下周进展计划，\n"
                + "  发送前人工复核。结构示例见 references/delivery.md）")
    else:
        body = "（正文为空，发送前请人工复核并填写。）"

    if args.self_test:
        to = [mail_from]
        cc: list[str] = []
        print("[NOTE] 自检模式：只发自己，不带抄送")
    else:
        to = [t.strip() for t in (args.to or "").split(",") if t.strip()]
        cc = [c.strip() for c in (args.cc or "").split(",") if c.strip()]
        if not to:
            fail("未指定收件人（--to / MAIL_TO / config.json mail_to）", 1)

    msg = build_message(mail_from=mail_from, to=to, cc=cc, subject=subject,
                        body=body, docx_path=args.docx)
    rcpts = to + cc
    mid = send(host=host, port=port, user=user, password=password,
               msg=msg, rcpts=rcpts)
    print(f"[OK] 已发送 → to={to} cc={cc or '（无）'}")
    print(f"[OK] 主题={subject}")
    print(f"[OK] 附件={name}（{os.path.getsize(args.docx)} 字节）")
    print(f"[OK] Message-ID: {mid}")

    if args.verify_sent:
        if not imap_host:
            print("[WARN] 未配置 imap_host，跳过回读校验，请人工查收已发文件夹。",
                  file=sys.stderr)
            return
        if verify_sent(imap_host, user, password, subject):
            print("[OK] 回读校验通过")
        else:
            fail("已发文件夹中未找到主题匹配的邮件，请人工查收确认。", 5)


if __name__ == "__main__":
    main()
