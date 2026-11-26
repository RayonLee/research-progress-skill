# 移植适配指南（PORTING）

本仓库最初在「灵犀 / Kimi Work + WPS 云文档 + QQ 邮箱连接器」环境跑通，但**流程本身与平台无关**。核心循环永远是：

```
生成（抓增量 → 归纳 → 渲染 DOCX → 目检 → 推水位线）
  → 人工复核（硬闸门，不可省略）
  → 交付（存档 → 发信 → 回读校验）
```

移植到任何新平台，只需要替换三个「适配点」，其余一律不动。

## 三个适配点

| 适配点 | 灵犀实现 | 通用实现 | 你要做的事 |
|---|---|---|---|
| **数据源**（工作记录从哪来） | `scripts/collect_recent.py`：WPS 云文档 `list-latest-items`，凭据为会话注入的 `TMP_LX_UUID` / `WPS_SID` | `scripts/collect_recent_fs.py`：扫描本地目录 mtime，零依赖 | 两者输出**完全相同的** `records.json`；若都不合适，写一个输出同构 records.json 的脚本即可（字段见下） |
| **存档**（DOCX 存到哪） | 云文档 CLI：`drive upload-file --folder-id <archive_folder_id>` | `scripts/archive_local.py`：复制到本地 `archive_dir` | 换平台时实现同一语义：上传前查同名冲突，**有同名即停止，不覆盖、不加序号** |
| **发信**（邮件怎么发） | Kimi Work 连接器：`connector_exec` 调 `mail_connector_id` | `scripts/send_mail_smtp.py`：标准库 SMTP + IMAP 回读 | 换渠道时保留脚本层硬校验：ASCII 附件名、主题==附件名去后缀、先发预览给本人 |

## 平台无关的不变量（移植时禁止丢弃）

1. **两阶段 + 人工硬闸门**：自动生成只到「待复核」为止；发送必须由人明确点头。这条不是功能，是事故防线——写错的周报发给导师不可撤回。
2. **水位线语义**：`last_success_at` 只在 DOCX 成功生成后推进（`--commit-state`）；审批拒绝回退到窗口起点；任何失败不推进。退出码 2/3/4/5 的语义保持不变。
3. **命名硬约束**：存档中文名 `YYYY年M月D日 <中文名>研究进展.docx`（月日不补零、一个半角空格），邮件 ASCII 名 `YYYY-MM-DD_<英文名>_Research_Progress.docx`，主题==附件名去后缀。渲染器与状态脚本强校验，不合规不落盘/不推进/不发送。
4. **论文表原样继承**：从 `--papers-from` 注入，不从记录推断改写。
5. **自产文件排除**：报告正文、ASCII 副本、样式模板、乱码名变体不进入统计窗口。
6. **事实边界**：mtime 只是「文件被修改过」，不能区分谁改的、是不是当天新写；不得把推断写成既成事实。

## records.json 输出契约

自定义数据源脚本必须输出到 `<workspace>/progress_auto/records.json`，顶层字段：

| 字段 | 类型 | 说明 |
|---|---|---|
| `window_start` / `window_end` | ISO8601 字符串 | 统计窗口（含时区） |
| `window_mode` | string | `incremental` / `explicit` / `first_run_fallback` |
| `total_in_window` | int | 窗口内有效记录数；为 0 时 `--commit-state` 会拒绝推进（退出码 5） |
| `excluded_self_produced` / `self_produced_names` | int / list | 自产文件排除计数与名单 |
| `week_counts` | dict | 按自然周（周一为界）分组计数 |
| `records` | list | 每条至少含：`file_id`（唯一）、`name`、`drive`、`ctime_bj`、`mtime_bj`、`link`、`week_monday`、`duplicate_name`、`noise`、`created_by`、`modified_by`（留空） |

## 移植检查单

- [ ] `config.json` 从 `config.example.json` 复制并填写：`report_name_cn` / `report_name_en`（决定所有命名规则）
- [ ] 数据源脚本跑出 `records.json`，`total_in_window > 0`
- [ ] `render_docx.py --papers-from <上期>` 生成成功，目检无空白页/裁切
- [ ] `--commit-state` 推进水位线 → `--resolve-approval approved/rejected` 语义验证（拒绝后水位线回退）
- [ ] 发信先 `--self-test` 只发本人，核对四项；正式发送后回读校验
- [ ] 定时任务**只做阶段一**；任务指令只引用技能/规范文件，不内联凭据

## 环境特有能力对照

| 能力 | 灵犀 / Kimi Work | 通用环境 |
|---|---|---|
| 云文档记录抓取 | `TMP_LX_UUID` / `WPS_SID` 会话注入 | 本地目录扫描 / 平台 API + 自行管理凭据 |
| PNG 目检 | `wpscli doc export --format png` | Word / LibreOffice 导出 PDF；或 docx2pdf 等 |
| 发信 | 连接器 `connector_exec` | SMTP（`send_mail_smtp.py`） |
| 定时 | 平台定时任务 | cron / Task Scheduler / 任何调度器 |
| 技能加载 | SKILL.md 自动加载 | 手动按 SKILL.md 流程执行，或接入支持 SKILL.md 的 agent |
