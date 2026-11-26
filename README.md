# research-progress-skill

把每周的工作记录（云文档 / 本地目录），整理成一份格式固定的 Word 研究进展周报，**经人工复核后**才上传存档、发送邮件。

> **核心设计：文档能重生成，邮件发出去收不回来。** 所以自动阶段只生成、不发送；发送必须人工点头。

流程与具体平台解耦：数据源、存档、发信各有「灵犀实现」和「通用实现」两份，核心不变量是**生成 → 人工复核 → 发送**的两阶段硬闸门。

---

## 它解决什么问题

每周要交一份固定六类（内参/项目/论文/课题/基础学习/其他）的进展报告。手工做有四件事很烦：

1. 翻记录回忆这周干了什么
2. 按六类归类、按尺度精炼（六类合计 300–400 字）
3. 调 Word 格式（标题、缩进、编号、七列表格）
4. 上传、写邮件、发出去

前三件能自动化，第四件必须人看。所以拆成两阶段。

---

## 两阶段

```
┌─ 阶段一：生成（自动，如每周三 09:00）───────────────┐
│                                                      │
│  抓取增量记录 → 六类归纳 → 论文表继承上期             │
│    → 渲染 DOCX → 导出 PNG 目检 → 推进水位线            │
│    → 输出链接 + 摘要 + 「待复核，未上传、未发送」       │
│    → ⏸  停住等人                                     │
└──────────────────────────────────────────────────────┘

┌─ 阶段二：交付（人工明确说「可以发」之后）─────────────┐
│                                                      │
│  复制 ASCII 副本 → 上传云端存档 → 发邮件              │
│    → 回读校验 → 结清审批状态                         │
└──────────────────────────────────────────────────────┘
```

五条铁律（任何情况下不得例外）：

| # | 规则 |
|---|---|
| 1 | DOCX 生成后必须交用户复核，**用户明确同意才能上传/发信**；上次批准过不算本次批准 |
| 2 | 邮件附件名必须纯 ASCII：`YYYY-MM-DD_<英文名>_Research_Progress.docx` |
| 3 | 邮件主题 = 附件名去掉 `.docx`，逐字符相等 |
| 4 | 云端存档仍用中文名 `YYYY年M月D日 <中文名>研究进展.docx`；邮件附件是独立 ASCII 副本 |
| 5 | 失败即停：任一环节失败立即停止并报告，不推进水位线、不产出半成品 |

---

## 目录结构

```
research-progress-skill/
├── SKILL.md                    # 技能入口（支持 SKILL.md 的 agent 自动加载）
├── README.md                   # 本文件
├── LICENSE                     # MIT
├── .gitignore
├── config.example.json         # 配置模板（config.json 为本地私有，已忽略）
├── template_spec.json          # 排版模板（唯一格式来源）
├── scripts/
│   ├── collect_recent.py       # 抓取【灵犀/WPS 云文档】记录 + 水位线/审批状态
│   ├── collect_recent_fs.py    # 抓取【通用本地目录】记录（零依赖，与上者同构）
│   ├── render_docx.py          # 模板驱动渲染器（含命名/标题强校验）
│   ├── send_mail_smtp.py       # 通用 SMTP 发信（ASCII 附件/主题强校验、自检、回读）
│   └── archive_local.py        # 通用本地存档（同名冲突即停）
├── references/
│   ├── pipeline.md             # 数据源、统计窗口、去重、噪声、维护操作污染
│   ├── writing-style.md        # 六类写法规范 + 尺度示例
│   └── delivery.md             # 命名、排版定稿值、上传、邮件规范、踩坑
└── docs/
    ├── 研究进展_整理配置.md      # 完整执行规范（最详细）
    └── PORTING.md              # 移植适配指南：三个适配点 + 不变量 + 检查单
```

### 改规则时改哪份

| 想改什么 | 改哪里 |
|---|---|
| 六类写法、篇幅要求 | `references/writing-style.md` |
| 页面/字体/缩进/表格列宽 | `template_spec.json` |
| 邮件渠道、收件人、命名 | `references/delivery.md` + `config.json` |
| 数据源、窗口、去重 | `references/pipeline.md` |
| 移植到新平台 | `docs/PORTING.md` |
| 完整背景与踩坑记录 | `docs/研究进展_整理配置.md` |

`SKILL.md` 是入口导航，保持精简；细节都在 `references/` 与 `docs/`。

---

## 安装

### 1. 依赖

- Python 3.10+，`pip install python-docx`（仅渲染器需要）
- 灵犀环境：WPS 灵犀桌面端（提供 `wpscli`、云文档 CLI、邮件连接器）
- 通用环境：仅标准库（本地抓取 / SMTP 发信 / 本地存档三个脚本零第三方依赖）

### 2. 配置

```bash
cp config.example.json config.json
```

按本机情况填写（全部字段说明见 `config.example.json` 内注释）：

| 字段 | 说明 |
|---|---|
| `report_name_cn` / `report_name_en` | 报告人中文名/英文名，**决定所有命名规则**，先填这个 |
| `workspace` | 工作区绝对路径（`state.json`、`records.json`、输出 DOCX 所在目录） |
| 灵犀版 | `wps_cli` / `template_file_id` / `template_name` / `archive_folder_id` / `mail_connector_id` |
| 通用版 | `source_dirs`（本地记录目录）/ `archive_dir`（本地归档）/ `smtp_*` / `imap_host` |
| 邮件 | `mail_account` / `mail_to` / `mail_cc` |

**`config.json` 含绝对路径、邮箱与凭据，永远不要提交**（已在 `.gitignore`）。仓库内不出现任何真实个人信息：姓名示例为「李某 / LiMou」，个人化的邮件正文套路只在你本地维护。

### 3. 样式底座

渲染需要一份含目标样式的 DOCX 作为 `styles.xml` / `numbering.xml` 底座：

- 灵犀版：首次运行自动从云文档下载 `template_file_id` 到 `<workspace>/refs/`
- 通用版：手动放任意一份排版满意的旧报告 DOCX 到 `refs/`，文件名与 `template_name` 一致

### 4. 安装为技能（可选）

支持 SKILL.md 的 agent（如 Kimi Work）：

```
skill(action="install", path="<本仓库绝对路径>")
```

装完**当前会话看不到变化**（技能清单是启动时快照），需开新会话。其他环境直接按 `SKILL.md` 的流程手动执行即可。

---

## 手动运行（一次完整周期）

```bash
S=scripts

# ── 阶段一 ──────────────────────────────────────────
# 1. 抓取增量（灵犀版 / 通用版二选一）
python3 $S/collect_recent.py                 # 或 collect_recent_fs.py
#    可加 --since "2026-11-09T00:00:00+08:00" 显式窗口；0 条记录退出码 5 直接收工

# 2. 按 references/writing-style.md 归纳，写 progress_auto/content.json
#    （title / summary / plan 三个字段；papers 不写）

# 3. 渲染（论文表从上期原样继承）
python3 $S/render_docx.py \
  --content progress_auto/content.json \
  --out "<workspace>/2026年12月2日 李某研究进展.docx" \
  --papers-from progress_auto/papers.json    # 或上期 DOCX

# 4. 目检（灵犀）
wpscli doc export "<docx>" --format png -o <输出目录>
#    通用环境用 Word/LibreOffice 导出 PDF 逐页看：边距、缩进、表格裁切、空白页

# 5. 推进水位线（渲染成功后才执行）
python3 $S/collect_recent.py --commit-state "<生成的docx>"

# ── 停住等复核：给人 file:// 链接 + 全文摘要 + 「待复核，未上传、未发送」──

# ── 阶段二（用户明确说「可以发」之后）────────────────
# 6. 复制 ASCII 副本
cp "<中文名>.docx" "2026-12-02_LiMou_Research_Progress.docx"

# 7. 存档（灵犀 / 通用）
wps_docs_cli.py drive upload-file --file-path "<中文名docx>" --folder-id <id>
# 或
python3 $S/archive_local.py "<中文名docx>"

# 8. 发信：首次先自检（只发自己、不带 cc）
python3 $S/send_mail_smtp.py --docx "2026-12-02_LiMou_Research_Progress.docx" --self-test
#    正式发送 + 回读校验
python3 $S/send_mail_smtp.py --docx "<ASCII副本>" --to <收件人> --cc <抄送> \
  --body body.txt --verify-sent
#    灵犀环境则走连接器 connector_exec（见 references/delivery.md）

# 9. 结清审批
python3 $S/collect_recent.py --resolve-approval approved   # 拒绝则 rejected
```

---

## 退出码

| 码 | 含义 | 处理 |
|---|---|---|
| 0 | 成功 | — |
| 1 | 参数/内容/模板配置错误 | 停止报告 |
| 2 | 凭据缺失或鉴权失效（仅灵犀版） | 停止，不重试不换来源 |
| 3 | 接口连续失败或非瞬时错误（仅灵犀版） | 停止，不产出残缺文档 |
| 4 | 文档不存在或命名不合规（发信脚本含主题不一致） | 停止，不推进水位线/拒绝发送 |
| 5 | 窗口内 0 条记录（发信脚本：回读未命中） | 说明情况并结束 |

---

## 命名规则

| 场景 | 格式 | 示例 |
|---|---|---|
| 本地文件 | `YYYY年M月D日 <中文名>研究进展.docx` | `2026年11月26日 李某研究进展.docx` |
| 云端存档 | 同上（中文） | 同上 |
| 邮件附件 | `YYYY-MM-DD_<英文名>_Research_Progress.docx` | `2026-11-26_LiMou_Research_Progress.docx` |
| 邮件主题 | 附件名去掉 `.docx` | `2026-11-26_LiMou_Research_Progress` |

存档名月/日**不补零**，`日` 与姓名之间**一个半角空格**，禁止任何后缀、括号、版本号。附件名为纯 ASCII（起源是某云端连接器中文附件名 MIME 编码缺陷导致收件端乱码；通用 SMTP 脚本仍强制此规则，见 `references/delivery.md`）。

上传接口普遍以 `basename` 作为云端文件名，不自动改名——**本地文件名即云端检索名**。

---

## 定时任务

| 项 | 值 |
|---|---|
| 调度 | 每周三 09:00（cron `0 0 9 * * 3`）或平台定时任务 |
| 范围 | **只做阶段一**，产出待审批稿后停止 |

任务指令只引用技能名和铁律摘要，不内联路径参数与凭据。改流程时改技能，不用动任务。

---

## 踩过的坑

这些都是实际调试中撞出来的，已全部固化进配置或脚本强校验。

| # | 现象 | 根因 | 修法 |
|---|---|---|---|
| 1 | 凭空多出空白第 2 页 | `after_plan` 空段 + 段前分页叠加，空段溢到下一页再被分页顶开 | `spacers.after_plan = 0` |
| 2 | 「拟完成时间」表头换行 | `tblLayout=autofit` 按内容压缩列宽 | 改 `fixed` |
| 3 | 全文整体上移 16pt | 空段行高只由段落标记 rPr 字号决定 | `spacers.mark` 逐档配置字号 |
| 4 | 表格跨页断行 | 浮动表格（`tblpPr`）不响应 `keepNext` | `cant_split` + `page_break_before` |
| 5 | 中文附件名乱码 | 连接器 MIME 编码缺陷 | 邮件附件改用 ASCII 名（脚本强校验） |
| 6 | 重命名污染统计窗口 | 批量操作刷新 mtime | 水位线推进到操作完成时刻 |
| 7 | 论文状态列写进周报叙述 | 从记录推断改写 | 原样继承上期表格 |
| 8 | token 接口 EOF 伪装成 401 | 鉴权误判 | 瞬时特征优先 + 3 次重试 |
| 9 | 标签冒号消失 | 拆分 label 时把全角冒号一起吃掉 | label 保留冒号，text 不重复 |
| 10 | 把自己的输出当周报 | 自动化产出的报告/副本/模板落入统计窗口 | `SELF_PRODUCED_KEYS` + 乱码名密度启发式排除 |

---

## 设计取舍

**为什么用模板 JSON 而不是 Word 占位符？** 占位符替换在动态行数（六类列表、论文表多行）时极易破坏格式。JSON 配置 + 渲染器既让模板可维护，又保证 100% 确定性——同一输入连续两次渲染，16 个 XML 部件逐字节一致。

**为什么论文表原样继承？** 论文是长期工作，极少周级动态；记录多为格式修改和版本迭代，不能反映实质进展。据此改写会把噪声写进状态列（实测后果：单元格变长撑高行高，整张表超过一页被拆开）。

**为什么拆两阶段？** 定时任务跑在无人值守会话，无法中途等待人工审批。上传和发信必须由人在审批后单独触发。

**为什么邮件附件强制 ASCII，连通用 SMTP 也是？** 一是与存档命名形成清晰分工（中文=存档，ASCII=邮件），二是规避个别企业网关/客户端的 MIME 解码差异。规则在脚本层强校验，不合规直接拒绝发送。

---

## 隐私与安全设计

- 仓库内**没有任何真实个人信息**：姓名/邮箱/连接器 id/云文档 id/个人邮件正文一律不进仓库，全部通过 `config.json`（已 gitignore）注入
- 灵犀版凭据由会话注入环境变量（`TMP_LX_UUID` / `WPS_SID`），脚本不读取、不保存、不打印
- SMTP 凭据只接受环境变量或本地 config.json，严禁写入代码、日志与产物

## License

MIT
