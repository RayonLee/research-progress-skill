# 交付规范

## 命名硬约束（检索以文档名为准）

**本地与云端存档唯一合法格式**：`YYYY年M月D日 <中文名>研究进展.docx`
（中文名 = config.json 的 `report_name_cn`，下文示例用「李某」）

- 年份四位；月/日**不补前导零**（如 `2026年9月9日`）
- `日` 与姓名之间**一个半角空格**（U+0020）
- 固定结尾 `<中文名>研究进展.docx`
- **禁止**任何后缀、括号、起止日期、版本号、`(2)`、`-v2`、状态标记
- 统计窗口起止写在文档正文（标题下方说明行），不写进文件名

**依据**：历史存档目录 52+ 份文档 100% 遵循此格式，检索以文档名为准。

**上传接口普遍以 `os.path.basename(file_path)` 作为云端文件名，不自动改名**——本地文件名即云端检索名，命名错误会直接污染存档。

同日已有同名文件：**停止并报告，不覆盖、不加序号**；由人决定是覆盖原稿还是顺延日期。

`collect_recent.py --commit-state`（及通用版 `collect_recent_fs.py`）会强校验文件名与文件存在性，命名不合规退出码 4 且拒绝推进水位线。

## 邮件附件命名（与存档不同）

**唯一合法格式**：`YYYY-MM-DD_<英文名>_Research_Progress.docx`
（英文名 = config.json 的 `report_name_en`，示例 `LiMou_Research_Progress`）

**邮件主题 = 附件名去掉 `.docx`**，逐字符相等。

### 为什么必须 ASCII

实测某云端邮件连接器写入 MIME 头时对**中文附件名编码有缺陷**，收件端显示成乱码。同封三附件对照实验（内容完全相同，只有文件名不同）结果：

| 文件名 | 收件端显示 |
|---|---|
| `2026年11月26日_李某研究进展.docx` | 乱码 |
| `2026-11-26_LiMou_Research_Progress.docx` | **正常** |
| `测试A中文名.docx` | 乱码 |

连接器是云端官方实现，本地无源码，无法修编码，只能让附件名避开非 ASCII。**不要试图靠改主题、改编码、改扩展名绕过。**

通用 SMTP 通道（`scripts/send_mail_smtp.py`）在协议层能正确处理中文名，但仍**强制 ASCII**：
一来保持与存档命名的清晰分工（中文名=存档，ASCII 名=邮件），二来避免个别企业网关/客户端的老旧解码问题。该规则已写进脚本的强校验，不合规拒绝发送。

发附件前先 `cp` 一份改名为 ASCII，**不要直接改云端存档名**。

## 排版模板定稿值

`template_spec.json` 是**唯一格式来源**，渲染器只执行模板。

| 键 | 规格 |
|---|---|
| `page` | A4 11906×16838 twips，四边边距 1134 twips |
| `title` | 居中，Times New Roman Bold / 方正小标宋简体，sz 32（16pt），加粗，kern=0 |
| `section` | 左对齐，Times New Roman / 宋体，sz 24（12pt），加粗；`plan_title_color=FF0000` 仅作用于「下周进展计划」段落标记 |
| `list` | List Paragraph（style_id 13），ind left=218 leftChars=104 firstLine=221 firstLineChars=100，sz 22（11pt）；编号与标签加粗、正文不加粗 |
| `list.num_id_summary` / `num_id_plan` | 本周手写「1.」～「6.」（numId=0）；下周用 Word 自动编号（numId=1），文本不再拼编号 |
| `paper` | List Paragraph，ind left=360 hanging=360，sz 24（12pt），加粗，numId=0 |
| `table` | style_id 9；`tblLayout=fixed`；7 列宽 2549/1370/1704/1310/1394/1505/1225 twips；`tblW=11057`；全边框 sz=4；单元格左右边距 108、上下 0；浮动定位 `tblpX=751 tblpY=299 tblOverlap=never`；表头行高 742、正文行高 90；表头字体 sz 21 |
| `table.cant_split` / `keep_together` / `repeat_header` | 行不断开 / 整表尽量不拆页 / 跨页时表头行重复。三项均为 true |
| `paper.page_break_before` | `true`，论文计划另起一页 |
| `spacers` | 标题后 1、本周六类后 1、**下周六类后 0**、末尾 1 |
| `spacers.mark` | 空段**段落标记字号**：标题后 16pt（微软雅黑）、本周后 12pt、下周后 11pt、末尾 11pt |
| `sections` | `本周进展概述` / `下周进展计划` / `论文计划：` |

### 三个已踩过的坑

1. **`spacers.after_plan` 必须为 0**：`page_break_before=true` 已把论文表顶到页首，若再留 4 个空段，空段会溢到下一页，分页再把它顶开，**整张空白页凭空多出来**。这两个值是一对：开了分页就归零空段。

2. **`tblLayout` 必须为 `fixed`**：`autofit` 会按内容压缩列宽，导致「拟完成时间」表头换行、表格宽度与参考不一致。

3. **空段行高只由段落标记 rPr 字号决定**：漏配 `spacers.mark` 会使全文上移 16pt（实测 32px）。

### 验证记录

- 确定性：同一 `content.json` + 同一 `template_spec.json` 连续两次渲染，16 个 XML 部件逐字节一致
- 版面：以参考文档内容渲染后与其 PNG 逐文本块比对，页面 1 的 17 个文本块、页面 2 的表格水平线（y=114/302/490/647）**偏移均为 0**

## 云端上传

- 目标文件夹：config.json 的 `archive_folder_id`（灵犀版）或 `archive_dir`（通用本地版）
- 灵犀版命令：
  ```bash
  wps_docs_cli.py drive upload-file --file-path <docx> --folder-id <archive_folder_id>
  ```
- 通用版命令：
  ```bash
  python3 scripts/archive_local.py "<docx>"
  ```
- 上传前复核：① 文件名符合命名硬约束；② 目标文件夹是否已有同名文件，有则停止请人确认

## 邮件

### 渠道（两种，按环境选一）

**A. 灵犀 / Kimi Work 连接器**（凭据存在连接器内，脚本只引用 id）

- 连接器：`connector_manage` id = config.json 的 `mail_connector_id`（QQ 邮箱类，SMTP/IMAP）
- 能力：`search` / `read` / `send`；**没有删除能力**，测试邮件需人工清理
- 调用：
  ```
  connector_exec(action="call", id=<mail_connector_id>, tool="send",
    arguments={account, to, cc, subject, body, content_type, attachments})
  ```
- `attachments` 传绝对路径列表；**文件名取本地 basename，无任何改写接口** → ASCII 规则只能靠本地副本改名实现

**B. 通用 SMTP**（`scripts/send_mail_smtp.py`，任何 Python 环境可用）

- 配置 config.json 的 `smtp_host` / `smtp_port` / `smtp_user` / `smtp_password` / `mail_from` / `imap_host`，或对应环境变量（`SMTP_*` / `MAIL_*` / `IMAP_HOST`）
- 脚本层强校验：附件 ASCII 命名、主题==附件名去后缀，不合规退出码 4 拒绝发送
- 回读校验：`--verify-sent` 用 IMAP 在已发文件夹按主题检索

### 收件人

config.json 的 `mail_to` / `mail_cc`。**测试发送必须只发自己、不带 cc**，否则正式版发出前就已抄送给合作者。

### 正文结构（`text/plain`）

正文是**高度个人化的固定套路**（称呼、客套话、落款），随仓库只给结构不给定稿文字——请在自己的 `config.json` 或本地正文模板里维护真实版本，**不要提交到仓库**。结构约束：

```
<称呼>，您好：
      <一句客套话，行首 6 个半角空格缩进>
      <一句祝福语，行首 6 个半角空格缩进>

本周进展概述
1.内参：……
……
6.其他：……

下周进展计划
1. 内参：……
……
6. 其他：……

此致
敬礼
```

- 行首 6 个半角空格缩进是定稿格式的一部分
- 概述部分编号后不加空格（`1.内参`），计划部分编号后加一个空格（`1. 内参`）——两处风格不同是刻意的，与 DOCX 版式一致

### 发送纪律

1. **首次发送前先给自己发预览**（只发自己、**不带 cc**），确认主题、正文、附件名、附件可打开四项均正常
2. 回读验证：连接器渠道 `search(folder="Sent Messages", since=...)` 拿 uid → `read` 核对 `to` / `cc` / `subject` / `attachments[].filename` / `size`；SMTP 渠道用 `--verify-sent`
3. 附件 SHA256 必须与本地文件一致

### 历史踩坑

- 连接器不在 `connector_manage list` 里时不要断言「未配置」，以实际调用结果为准
- 测试发送必须只发自己、**不得带 cc**
- 预览邮件的主题与正式版相同即可，便于确认最终效果

## 审批状态

- 通过：`python3 scripts/collect_recent.py --resolve-approval approved`（通用版把脚本名换成 `collect_recent_fs.py`）
- 拒绝：`--resolve-approval rejected`，不上传不发信；**脚本会把水位线回退到该期窗口起点**，使这批记录在下期重新纳入，避免被永久跳过
