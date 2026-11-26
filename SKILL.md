---
name: research-progress
description: 生成并交付每周研究进展 Word 文档。适用于：整理近期云文档/本地目录工作记录、生成「YYYY年M月D日 <姓名>研究进展.docx」、按六类（内参/项目/论文/课题/基础学习/其他）撰写进展与下周计划、从上一期继承论文计划表、导出 PNG 目检、上传云端存档、发送研究进展邮件。触发词：研究进展、每周汇报、进展整理、周报、research progress。
---

# 研究进展周报

把一周的工作记录整理成一份格式固定的 Word 进展报告，经人复核后上传存档并发邮件。

## 铁律（优先级最高，任何情况下不得例外）

| # | 规则 |
|---|---|
| **1** | **DOCX 生成后必须交用户复核，用户明确同意才能上传/发信**。生成完只给「文档链接 + 正文摘要 + 待复核标注」，然后停住等人。**严禁**自动上传或发信；上次批准过不算本次批准。 |
| **2** | **邮件附件名必须是纯 ASCII**：`YYYY-MM-DD_<英文名>_Research_Progress.docx`（英文名取 config.json 的 `report_name_en`） |
| **3** | **邮件主题 = 附件名去掉 `.docx`**，逐字符相等 |
| **4** | **云端存档仍用中文名** `YYYY年M月D日 <中文名>研究进展.docx`；邮件附件是独立 ASCII 副本，不改存档名 |
| **5** | **失败即停**：任一环节失败立即停止并报告，不推进水位线、不产出半成品 |

> 第 1 条的理由：定时任务跑在无人值守会话，用户不在现场。写错的周报发给导师不可撤回——文档能重生成，邮件发出去收不回来。

## 两阶段

| 阶段 | 触发 | 动作 |
|---|---|---|
| 一：生成 | 定时（周三 09:00）或用户要求 | 抓增量 → 归纳 → 渲染 DOCX → PNG 目检 → 推进水位线 → 标记待审批 → **停住** |
| 二：交付 | 用户明确说「可以发」后 | 复制 ASCII 副本 → 上传云端存档 → 发邮件 → 回读校验 → 结清审批 |

## 仓库与关键文件

脚本已配置驱动：读仓库根 `config.json`（缺失时回退内置默认值，姓名默认「李某 / LiMou」）。

| 角色 | 路径（相对仓库根） |
|---|---|
| 仓库根 | 本仓库 |
| 本地配置 | `config.json`（含本机路径/邮箱，已 gitignore） |
| 配置示例 | `config.example.json` |
| 抓取脚本（灵犀/WPS 云文档） | `scripts/collect_recent.py` |
| 抓取脚本（通用本地目录） | `scripts/collect_recent_fs.py` |
| 渲染器 | `scripts/render_docx.py` |
| 通用发信（SMTP） | `scripts/send_mail_smtp.py` |
| 通用本地存档 | `scripts/archive_local.py` |
| 排版模板 | `template_spec.json` |
| 完整规范 | `docs/研究进展_整理配置.md` |
| 适配指南（移植到其他平台） | `docs/PORTING.md` |
| 水位线 | `<workspace>/progress_auto/state.json` |
| 论文表基线 | `<workspace>/progress_auto/papers.json` |
| 样式底座 | `<workspace>/refs/<template_name>` |

工作区（workspace）在 `config.json` 配置；缺省为仓库的上一级目录。

## 阶段一：生成

1. **抓取增量**（按数据源二选一，输出结构相同）
   ```bash
   # 灵犀 / WPS 云文档
   python3 scripts/collect_recent.py
   # 通用：本地目录（config.json 的 source_dirs）
   python3 scripts/collect_recent_fs.py
   ```
   窗口自动取自 `state.json` 的 `last_success_at` 至当前时刻。窗口内 0 条 → 退出码 5，直接回「本期无新增记录」并结束，不生成、不推进水位线。

2. **六类归纳**
   固定顺序：内参 / 项目 / 论文 / 课题 / 基础学习 / 其他。
   **写法必须精炼精要**，详见 [references/writing-style.md](references/writing-style.md)：每类 1–2 句、12–110 字，六类合计 300–400 字（超 500 字视为未收敛，重写）；动词开头；不列文件名、版本号、字数、意见来源；同类事项合并计数；下周无推进写「暂无。」

3. **确定论文表继承源**
   优先 `state.json` 的 `last_output`；缺失则从云端存档文件夹（config.json 的 `archive_folder_id`）取最新一份；都没有 → 停止报告，不生成缺表文档。

4. **写 content.json**
   只写 `title` / `summary` / `plan`，**不要写 `papers`**（由 `--papers-from` 注入）。

5. **渲染**
   ```bash
   python3 scripts/render_docx.py \
     --content <content.json 绝对路径> \
     --out "<输出目录>/YYYY年M月D日 <中文名>研究进展.docx" \
     --papers-from "<上期 DOCX 绝对路径>"
   ```
   `--spec` 有默认值（仓库根的 `template_spec.json`），无需显式传。

   **禁止手写 python-docx / JS 排版代码，禁止改 `template_spec.json` 或渲染器。**
   渲染器强校验：文件名合规 + 文档标题 == 文件名去 `.docx`，不合规退出码 3 且不落盘。

6. **PNG 目检**（灵犀环境；其他环境用 Word/LibreOffice 导出 PDF 或逐页查看）
   ```bash
   wpscli grant-dir-access <workspace>
   wpscli doc export "<docx>" --format png -o <输出目录>
   ```
   检查：页面边距、标题、列表缩进、表格有无裁切、**有无空白页**。异常只改 `content.json`。

7. **推进水位线**
   ```bash
   python3 scripts/collect_recent.py --commit-state "<生成的docx绝对路径>"
   ```
   （通用数据源则把脚本名换成 `collect_recent_fs.py`）
   仅 DOCX 确实生成成功后才执行；退出码非 0 须停止报告。

8. **停住等复核**
   输出：`file://` 链接 + 六类正文与下周计划全文 + 论文表状态列 + 「待复核，未上传、未发送」。

## 阶段二：交付

**前置条件：用户已明确同意发送。** 未获同意一律停在阶段一。

1. **复制 ASCII 副本**
   ```bash
   cp "<中文名>.docx" "<YYYY-MM-DD_<英文名>_Research_Progress>.docx"
   ```

2. **上传云端存档**（中文名）
   - 灵犀环境：
     ```bash
     wps_docs_cli.py drive upload-file --file-path "<中文名docx>" --folder-id <archive_folder_id>
     ```
   - 通用环境（本地归档目录）：
     ```bash
     python3 scripts/archive_local.py "<中文名docx>"
     ```
   上传前检查目标文件夹是否已有同名文件，有则停止请人确认，不覆盖、不加序号。

3. **发邮件**（按渠道二选一）
   - 灵犀 / Kimi Work 连接器：id 取 config.json 的 `mail_connector_id`，正文模板、附件规则、回读校验见 [references/delivery.md](references/delivery.md)
   - 通用 SMTP：
     ```bash
     # 首次先自检（只发自己、不带 cc）
     python3 scripts/send_mail_smtp.py --docx "<ASCII副本>" --self-test
     # 正式发送 + 回读校验
     python3 scripts/send_mail_smtp.py --docx "<ASCII副本>" --to <收件人> --cc <抄送> --body <正文.txt> --verify-sent
     ```
   - **首次发送前先给自己发预览**（只发自己、不带 cc），确认主题/正文/附件名/可打开四项

4. **结清审批**
   ```bash
   python3 scripts/collect_recent.py --resolve-approval approved
   ```
   用户拒绝则 `--resolve-approval rejected`，不上传不发信，水位线回退到该期窗口起点。

## 退出码

| 码 | 含义 | 处理 |
|---|---|---|
| 0 | 成功 | — |
| 1 | 参数/内容/模板配置错误 | 停止报告 |
| 2 | 凭据缺失或鉴权失效（仅灵犀版） | 停止，不重试不换来源 |
| 3 | 接口连续失败或非瞬时错误（仅灵犀版） | 停止，不产出残缺文档 |
| 4 | 文档不存在或命名不合规 | 停止，不推进水位线 |
| 5 | 窗口内 0 条记录（发信脚本：回读未命中） | 说明情况并结束 |

## 参考文档

- [references/pipeline.md](references/pipeline.md) — 数据源、统计窗口、去重、字段限制、噪声、维护操作污染
- [references/writing-style.md](references/writing-style.md) — 六类写法规范、量化标准、尺度示例
- [references/delivery.md](references/delivery.md) — 命名硬约束、排版模板定稿值、云端上传、邮件规范、连接器坑位
- [docs/PORTING.md](docs/PORTING.md) — 移植到其他 agent 平台 / 邮件渠道 / 数据源的适配指南

## 其他约束

- 只写记录可证实的内容；记录是文件级修改时间，无法区分本人或协作者，不得把推断写成既成事实
- 严禁把凭据值写入任何文件、脚本或产物
- 灵犀版云文档 CLI 依赖会话注入的 `TMP_LX_UUID` / `WPS_SID`；翻页时 token 接口偶发 EOF 并伪装成 `HTTP 401`，脚本已做 3 次重试，连续失败即退出码 3
