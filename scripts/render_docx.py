#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究进展 DOCX 渲染器（模板驱动）

职责：把 content.json 渲染为 DOCX。**一切格式来自 template_spec.json**，
本脚本不含任何硬编码排版参数。会话只负责产出内容，不得手写排版，
从而保证每次生成的页面、字体、段落、缩进、表格样式完全一致。

模板 = template_spec.json（排版）+ refs 下的样式底座 DOCX（styles.xml /
numbering.xml / docDefaults）。改模板只需改 JSON，无需改本脚本。

用法:
    python3 render_docx.py --content content.json --out "2026年12月2日 李某研究进展.docx"
    python3 render_docx.py --content content.json --out ... --spec template_spec.json
    python3 render_docx.py --content content.json --out ... --papers-from "上期文档.docx"

--papers-from：论文计划表不从 content.json 取，而是从上期 DOCX 的第一个表格
    **原样照抄**（列名、行数、每个单元格文字全部继承）。论文是长期工作、
    动态很少，记录来源多是文档格式修改，不能反映论文实质进展，因此禁止
    从记录推断改写；确需变更时由人在审批时口头提出，改后的版本成为下一期基准。

content.json 结构:
    {
      "title": "2026年12月2日 李某研究进展",
      "summary":  [{"no": "1.", "label": "内参：", "text": "..."}, ...],
      "plan":     [{"no": "1.", "label": "内参：", "text": "..."}, ...],
      "papers":   {"columns": [...7列...], "rows": [[...], ...]}
    }

命名说明：
    报告人姓名/英文名从 config.json 的 report_name_cn / report_name_en 读取
    （示例默认「李某 / LiMou」），命名硬约束为
    YYYY年M月D日 <中文名>研究进展.docx（月/日不补零、一个半角空格）。

退出码:
    0 = 成功
    1 = 参数/内容/模板配置错误
    2 = 样式底座缺失且自动下载失败
    3 = 输出文件名或标题不合规（拒绝落盘）
    4 = 上期文档缺失或无可用表格（无法继承论文计划表）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

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

WORKSPACE = CFG.get("workspace") or os.path.dirname(REPO_DIR)
# 灵犀（WPS）云文档 CLI；config.json 的 wps_cli 优先，
# 缺省按 macOS 标准安装位置推导（expanduser，不写死任何用户名）
CLI = CFG.get("wps_cli") or os.path.expanduser(
    "~/Library/Application Support/WPS 灵犀/"
    "serverdir/target_skills/wps-docs/scripts/wps_docs_cli.py")
REFS_DIR = os.path.join(WORKSPACE, "refs")
TEMPLATE_FILE_ID = CFG.get("template_file_id") or ""
TEMPLATE_NAME = CFG.get("template_name") or "template.docx"
DEFAULT_TEMPLATE = os.path.join(REFS_DIR, TEMPLATE_NAME)
DEFAULT_SPEC = CFG.get("template_spec") or os.path.join(REPO_DIR, "template_spec.json")

# 报告人姓名（中文名用于本地/云端存档命名，英文名用于邮件附件命名）
NAME_CN = CFG.get("report_name_cn") or "李某研究进展"
NAME_EN = CFG.get("report_name_en") or "LiMou_Research_Progress"

# 输出文件名硬约束（与 collect_recent.py 保持一致）
DOCX_NAME_RE = re.compile(
    r'^\d{4}年(?:[1-9]|1[0-2])月(?:[1-9]|[12]\d|3[01])日 '
    + re.escape(NAME_CN) + r'\.docx$')


def ensure_template(template: str) -> None:
    """样式底座缺失时从云文档下载，避免静默失败。"""
    if os.path.isfile(template):
        return
    if not TEMPLATE_FILE_ID:
        print(f"[FATAL] 样式底座缺失：{template}", file=sys.stderr)
        print("[FATAL] 请在 config.json 配置 template_file_id（云文档 ID），"
              "或手动把样式底座 DOCX 放到上述路径。", file=sys.stderr)
        sys.exit(2)
    print(f"[WARN] 样式底座缺失，尝试从云文档下载：{TEMPLATE_FILE_ID}",
          file=sys.stderr)
    os.makedirs(os.path.dirname(template), exist_ok=True)
    p = subprocess.run(
        [sys.executable, CLI, "drive", "download-file",
         "--file-id", TEMPLATE_FILE_ID, "--save-dir", os.path.dirname(template)],
        capture_output=True, text=True)
    if p.returncode != 0 or not os.path.isfile(template):
        msg = (p.stderr or p.stdout or "").strip()
        print(f"[FATAL] 样式底座下载失败，请手动放置到 {template}：{msg[:300]}",
              file=sys.stderr)
        sys.exit(2)
    print(f"[OK] 样式底座已下载：{template}", file=sys.stderr)


def validate_output(out_path: str, title: str) -> None:
    """输出文件名必须合规，且文档标题与文件名严格一致（检索以名为准）。"""
    name = os.path.basename(out_path)
    if not DOCX_NAME_RE.match(name):
        print("[FATAL] 输出文件名不符合规范：" + name, file=sys.stderr)
        print(f"[FATAL] 须为 YYYY年M月D日 {NAME_CN}.docx"
              "（月日不补零、一个半角空格、无后缀/序号）", file=sys.stderr)
        sys.exit(3)
    if title != os.path.splitext(name)[0]:
        print(f"[FATAL] 文档标题与文件名不一致：标题={title!r} 文件名={name!r}",
              file=sys.stderr)
        sys.exit(3)


def el(tag: str, **attrs) -> OxmlElement:
    """创建带 w: 命名空间属性的元素（tag 自动补 w: 前缀）。"""
    e = OxmlElement(tag if ":" in tag else f"w:{tag}")
    for k, v in attrs.items():
        if v is None:
            continue
        e.set(qn(f"w:{k}"), str(v))
    return e


def make_rpr(*, ascii: str, eastAsia: str, sz: int, bold: bool,
             cs: str | None = None, color: str | None = None,
             hint: str = "eastAsia", kern: str | None = None) -> OxmlElement:
    rpr = OxmlElement("w:rPr")
    rpr.append(el("rFonts", hint=hint, ascii=ascii, hAnsi=ascii,
                  eastAsia=eastAsia, cs=cs or ascii))
    if bold:
        rpr.append(el("b"))
        rpr.append(el("bCs"))
    else:
        rpr.append(el("b", val="0"))
        rpr.append(el("bCs", val="0"))
    if color:
        rpr.append(el("color", val=color))
    if kern is not None:
        rpr.append(el("kern", val=kern))
    rpr.append(el("sz", val=sz))
    rpr.append(el("szCs", val=sz))
    rpr.append(el("lang", val="en-US", eastAsia="zh-CN"))
    return rpr


def make_run(text: str, rpr: OxmlElement) -> OxmlElement:
    r = OxmlElement("w:r")
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def add_para(doc, runs: list[tuple[str, OxmlElement]],
             ppr: OxmlElement | None = None) -> None:
    """添加段落。ppr 已是完整 pPr，直接作为段落属性，不得再嵌套。"""
    p = OxmlElement("w:p")
    p.append(ppr if ppr is not None else OxmlElement("w:pPr"))
    for text, rpr in runs:
        p.append(make_run(text, rpr))
    doc.element.body.append(p)


def make_ppr(*, style: str | None = None, ind: dict | None = None,
             align: str | None = None, num_id: int | None = None,
             keep_next: bool = False, page_break_before: bool = False,
             mark_rpr: OxmlElement | None = None) -> OxmlElement:
    """按 OOXML 规定顺序构造 pPr。

    pStyle → keepNext → pageBreakBefore → numPr → ind → jc → rPr
    """
    ppr = OxmlElement("w:pPr")
    if style:
        ppr.append(el("pStyle", val=style))
    if keep_next:
        ppr.append(el("keepNext"))
    if page_break_before:
        ppr.append(el("pageBreakBefore"))
    if num_id is not None:
        numpr = OxmlElement("w:numPr")
        numpr.append(el("ilvl", val=0))
        numpr.append(el("numId", val=num_id))
        ppr.append(numpr)
    if ind:
        ppr.append(el("ind", **ind))
    if align:
        ppr.append(el("jc", val=align))
    if mark_rpr is not None:
        ppr.append(mark_rpr)
    return ppr


def extract_papers(path: str) -> dict:
    """读取论文计划表。

    支持两种来源：
      *.json  -> 直接读 {"columns": [...], "rows": [[...], ...]}（人可直接编辑）
      其他     -> 从上期 DOCX 的第一个表格原样提取

    DOCX 单元格内多个段落用 \\n 连接，与 build_table 的拆分规则互逆，
    保证“照抄”后渲染出来与上期完全一致。
    """
    if not os.path.isfile(path):
        print(f"[FATAL] 上期文档/基线缺失，无法继承论文计划表：{path}",
              file=sys.stderr)
        sys.exit(4)

    if path.lower().endswith(".json"):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            print(f"[FATAL] 论文表基线无法解析：{path}：{e}", file=sys.stderr)
            sys.exit(4)
        columns = [str(c).strip() for c in data.get("columns") or []]
        rows = [[str(c) for c in r] for r in data.get("rows") or []]
        rows = [r for r in rows if any(x.strip() for x in r)]
        if not columns or not rows:
            print(f"[FATAL] 论文表基线缺少 columns/rows：{path}", file=sys.stderr)
            sys.exit(4)
        bad = [i for i, r in enumerate(rows) if len(r) != len(columns)]
        if bad:
            print(f"[FATAL] 论文表基线第 {bad} 行列数与表头不一致"
                  f"（应 {len(columns)} 列）", file=sys.stderr)
            sys.exit(4)
        return {"columns": columns, "rows": rows}

    d = Document(path)
    if not d.tables:
        print(f"[FATAL] 上期文档没有表格：{path}", file=sys.stderr)
        sys.exit(4)
    rows = d.tables[0].rows
    if len(rows) < 2:
        print(f"[FATAL] 上期文档表格只有 {len(rows)} 行，无数据行：{path}",
              file=sys.stderr)
        sys.exit(4)
    columns = [c.text.strip() for c in rows[0].cells]
    data = []
    for r in rows[1:]:
        cells = [c.text.strip() for c in r.cells]
        if not any(cells):
            continue
        data.append(cells)
    if not data:
        print(f"[FATAL] 上期文档表格没有有效数据行：{path}", file=sys.stderr)
        sys.exit(4)
    return {"columns": columns, "rows": data}


def spacer_ppr(mark: dict, **ppr_kwargs) -> OxmlElement:
    """空段落的 pPr：段落标记 rPr 的字号决定行高，必须显式给出。"""
    mr = make_rpr(ascii=mark["ascii"], eastAsia=mark["eastAsia"],
                  sz=mark["sz"], bold=mark["bold"],
                  cs=mark.get("cs"), color=mark.get("color"),
                  hint=mark.get("hint", "eastAsia"), kern=mark.get("kern"))
    return make_ppr(mark_rpr=mr, **ppr_kwargs)


def build_table(doc, spec: dict, columns: list[str], rows: list[list[str]]) -> None:
    """按模板规格重建表格：固定列宽、全边框、表头加粗居中。"""
    t = spec["table"]
    col_w = t["col_widths"]
    tbl = OxmlElement("w:tbl")
    tblpr = OxmlElement("w:tblPr")
    tblpr.append(el("tblStyle", val=t["style_id"]))
    tblpr.append(el("tblpPr", **t["pos"]))
    tblpr.append(el("tblOverlap", val="never"))
    tblpr.append(el("tblW", w=t["width_dxa"], type="dxa"))
    tblpr.append(el("tblInd", w=0, type="dxa"))
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(el(side, val="single", sz=t["border_sz"], space="0",
                          color="auto"))
    tblpr.append(borders)
    tblpr.append(el("tblLayout", type="fixed"))
    cellmar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        cellmar.append(el(side, w=t["cell_margin"][side], type="dxa"))
    tblpr.append(cellmar)
    tbl.append(tblpr)

    grid = OxmlElement("w:tblGrid")
    for w in col_w[:len(columns)]:
        grid.append(el("gridCol", w=w))
    tbl.append(grid)

    hf = t["header_font"]
    for ri, values in enumerate([columns] + rows):
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        # trPr 子元素必须按 OOXML 顺序：cantSplit → trHeight → tblHeader
        if t.get("cant_split", True):
            trpr.append(el("cantSplit"))
        trpr.append(el("trHeight",
                       val=t["row_height_header"] if ri == 0 else t["row_height_body"],
                       hRule="atLeast"))
        # 表格跨页时表头行在续页重复，避免续页看不到列名
        if ri == 0 and t.get("repeat_header", True):
            trpr.append(el("tblHeader"))
        tr.append(trpr)
        for ci, value in enumerate(values[:len(columns)]):
            tc = OxmlElement("w:tc")
            tcpr = OxmlElement("w:tcPr")
            tcpr.append(el("tcW", w=col_w[ci], type="dxa"))
            tcpr.append(el("vAlign", val="center"))
            tc.append(tcpr)
            # 参考文档中单元格内多行是多个独立段落（非软换行），逐行重建
            # keep_together：除最后一行外全部 keepNext，把整张表锁在同一页，
            # 避免表头留在上页、数据行被推到下页
            last_row = ri == len([columns] + rows) - 1
            keep = t.get("keep_together", True) and not last_row
            for line in str(value or "").split("\n"):
                p = OxmlElement("w:p")
                p.append(make_ppr(
                    align="center", keep_next=keep,
                    mark_rpr=make_rpr(ascii=hf["ascii"], eastAsia=hf["eastAsia"],
                                      sz=hf["sz"], bold=ri == 0,
                                      cs=hf["cs"],
                                      hint="eastAsia" if ri == 0 else "default")))
                p.append(make_run(line, make_rpr(
                    ascii=hf["ascii"], eastAsia=hf["eastAsia"], sz=hf["sz"],
                    bold=ri == 0, cs=hf["cs"],
                    hint="eastAsia" if ri == 0 else "default")))
                tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    doc.element.body.append(tbl)


def render(content: dict, spec: dict, template: str, out_path: str,
           skip_validate: bool = False) -> None:
    ensure_template(template)
    if not skip_validate:
        validate_output(out_path, content["title"])

    doc = Document(template)
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    # 页面
    pg = spec["page"]
    for s in doc.sections:
        s.page_width, s.page_height = Twips(pg["width_twips"]), Twips(pg["height_twips"])
        s.top_margin = s.bottom_margin = Twips(pg["margin_twips"])
        s.left_margin = s.right_margin = Twips(pg["margin_twips"])

    t, sec, lst, pap = spec["title"], spec["section"], spec["list"], spec["paper"]
    sp = spec["spacers"]
    mk = sp["mark"]
    heads = spec["sections"]

    # 标题
    add_para(doc, [(content["title"], make_rpr(
        ascii=t["ascii"], eastAsia=t["eastAsia"], sz=t["sz"], bold=t["bold"],
        cs=t["cs"], kern=t["kern"]))],
        make_ppr(align=t["align"]))
    for _ in range(sp["after_title"]):
        add_para(doc, [], spacer_ppr(mk["after_title"],
                                     align=mk["after_title"]["align"]))

    # 本周进展概述
    add_para(doc, [(heads["summary_heading"], make_rpr(
        ascii=sec["ascii"], eastAsia=sec["eastAsia"], sz=sec["sz"],
        bold=sec["bold"], cs=sec["cs"], hint="default"))], make_ppr())
    for item in content.get("summary", []):
        add_para(doc, [
            (item.get("no", ""), make_rpr(ascii=lst["ascii"], eastAsia=lst["eastAsia"],
                                          sz=lst["sz"], bold=True)),
            (item.get("label", ""), make_rpr(ascii=lst["ascii"], eastAsia=lst["eastAsia"],
                                             sz=lst["sz"], bold=True)),
            (item.get("text", ""), make_rpr(ascii=lst["ascii"], eastAsia=lst["eastAsia"],
                                            sz=lst["sz"], bold=False)),
        ], make_ppr(style=lst["style_id"], ind=lst["ind"],
                    num_id=lst["num_id_summary"]))
    for _ in range(sp["after_summary"]):
        add_para(doc, [], spacer_ppr(mk["after_summary"], style=lst["style_id"],
                                     ind=lst["ind"],
                                     num_id=lst["num_id_summary"]))

    # 下周进展计划（标题红色仅作用于段落标记，文字为黑色）
    add_para(doc, [(heads["plan_heading"], make_rpr(
        ascii=sec["ascii"], eastAsia=sec["eastAsia"], sz=sec["sz"],
        bold=sec["bold"], cs=sec["cs"], hint="default"))],
        make_ppr(mark_rpr=make_rpr(
            ascii=sec["ascii"], eastAsia=sec["eastAsia"], sz=sec["sz"],
            bold=True, cs=sec["cs"], color=sec["plan_title_color"],
            hint="default")))
    for item in content.get("plan", []):
        # 编号由 Word 自动生成（num_id_plan），文本不再拼接 "1."
        add_para(doc, [
            (item.get("label", ""), make_rpr(ascii=lst["ascii"], eastAsia=lst["eastAsia"],
                                             sz=lst["sz"], bold=True)),
            (item.get("text", ""), make_rpr(ascii=lst["ascii"], eastAsia=lst["eastAsia"],
                                            sz=lst["sz"], bold=False)),
        ], make_ppr(style=lst["style_id"], ind=lst["ind"],
                    num_id=lst["num_id_plan"]))
    for _ in range(sp["after_plan"]):
        add_para(doc, [], spacer_ppr(mk["after_plan"], style=lst["style_id"],
                                     ind=lst["ind"],
                                     num_id=lst["num_id_summary"]))

    # 论文计划
    add_para(doc, [(heads["paper_heading"], make_rpr(
        ascii=pap["ascii"], eastAsia=pap["eastAsia"], sz=pap["sz"],
        bold=pap["bold"], cs=pap["ascii"]))],
        make_ppr(style=pap["style_id"], ind=pap["ind"], num_id=pap["num_id"],
                 keep_next=pap.get("keep_next", True),
                 page_break_before=pap.get("page_break_before", False)))

    papers = content.get("papers") or {}
    columns = papers.get("columns") or []
    if columns:
        build_table(doc, spec, columns, papers.get("rows") or [])

    for _ in range(sp["tail"]):
        add_para(doc, [], spacer_ppr(mk["tail"]))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"[OK] 已生成 {out_path}")
    print(f"[OK] 段落数={len(doc.paragraphs)} 表格数={len(doc.tables)}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", required=True, help="content.json 路径")
    ap.add_argument("--out", required=True, help="输出 docx 绝对路径")
    ap.add_argument("--spec", default=DEFAULT_SPEC, help="template_spec.json 路径")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="样式底座 docx")
    ap.add_argument("--papers-from",
                    help="上期 DOCX 路径；给出时论文计划表从中原样继承")
    ap.add_argument("--skip-validate", action="store_true",
                    help="仅用于生成空白模板示例，跳过命名/标题校验")
    args = ap.parse_args()

    if not os.path.isfile(args.content):
        print(f"[FATAL] 内容文件缺失：{args.content}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.spec):
        print(f"[FATAL] 模板配置缺失：{args.spec}", file=sys.stderr)
        sys.exit(1)
    try:
        with open(args.content, encoding="utf-8") as f:
            content = json.load(f)
        with open(args.spec, encoding="utf-8") as f:
            spec = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"[FATAL] JSON 无法解析：{e}", file=sys.stderr)
        sys.exit(1)
    if not content.get("title"):
        print("[FATAL] content.json 缺少 title", file=sys.stderr)
        sys.exit(1)

    if args.papers_from:
        content["papers"] = extract_papers(args.papers_from)
        print(f"[OK] 已从《{os.path.basename(args.papers_from)}》"
              f"原样继承论文计划表 {len(content['papers']['rows'])} 行"
              f"（{len(content['papers']['columns'])} 列）", file=sys.stderr)

    render(content, spec, args.template, args.out,
           skip_validate=args.skip_validate)


if __name__ == "__main__":
    main()
