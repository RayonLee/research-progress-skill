# 数据源与统计窗口

## 两种数据源

| 数据源 | 脚本 | 凭据 | 适用 |
|---|---|---|---|
| WPS / 灵犀云文档「最近访问/编辑」 | `scripts/collect_recent.py` | 会话注入的 `TMP_LX_UUID` / `WPS_SID` | 灵犀 / WPS 环境 |
| 本地目录 mtime 扫描 | `scripts/collect_recent_fs.py` | 无（纯本地） | 任何环境，最通用 |

两个脚本输出**完全相同的** `progress_auto/records.json` 结构，水位线与审批状态共用同一份 `state.json`，归纳、渲染、交付流程不做任何区分。

### 灵犀版来源细节

- 接口：WPS 云文档 `drive list-latest-items`（最近访问/编辑记录）
- 分页：`--page-size 200`，沿 `next_page_token` 翻页直到覆盖窗口起点，最多 20 页
- 同名副本（「我的云文档」/「我的漫游箱」等）按 `file_id` 去重、保留最新 mtime，归纳阶段再按文件名二次去重

### 通用版来源细节

- config.json 的 `source_dirs`（可多个本地目录），递归扫描 mtime 落在窗口内的文件
- 自动跳过隐藏目录与 `node_modules` / `__pycache__` / `venv` / `dist` / `build`
- 文件绝对路径即唯一 id，不存在云端副本问题

## 统计窗口

**自上次成功运行时刻至本次运行时刻的全部记录**（增量，不重叠、不遗漏）。

- 水位线存于 `progress_auto/state.json` 的 `last_success_at`
- 首次运行（无水位线）兜底为滚动近三个自然周，`window_mode=first_run_fallback`；此时会与已交付文档内容部分重叠，须在交付说明中注明
- **只有阶段一 DOCX 成功生成后**，才调用 `--commit-state <docx绝对路径>` 推进水位线并标记待审批
- 任何环节失败都不得推进水位线；下次运行会自动补齐，不会漏数据
- 可用 `--since <ISO8601>` 显式覆盖窗口起点

## 维护操作污染（重要）

批量重命名、移动、权限变更等操作会刷新文档 mtime，使其被误判为「研究进展」。

此类操作完成后**必须**把 `last_success_at` 推进到操作完成时刻（用 `--since` 或直接改 `state.json`），否则下期会把这些文档当成新增修改写入。

## 时区与分组

云端接口返回 UTC 时一律换算本地时区后按自然周（周一为界）分组。通用版直接使用本地时区。

## 字段限制

- 灵犀版接口返回 `file_src.name`（盘名），**不返回 `created_by` / `modified_by`**，因此不得声称某文件由谁创建或修改

## 噪声

备份码、验证码、签到表、报销单、临时 txt 等默认归入「其他」且不单独展开。

两个抓取脚本内置相同的 `NOISE_KEYS`：验证码、签到表、备份码、backup codes、公务卡、报酬、杂志领取、学籍、征文大赛、tmp、.s2t。

## 自产文件排除（不进入统计窗口）

自动化**自己产出**的文件不是研究进展。若混入窗口，会被当成「本周工作」写进报告，形成自指噪声。这类文件在 `records` 组装阶段就被剔除，并在 `records.json` 的 `excluded_self_produced` 里计数。

### 排除规则

| 特征 | 命中的文件 |
|---|---|
| `<report_name_cn>`（如 `李某研究进展`） | 报告正文、云端存档、样式模板 |
| `<report_name_en>`（如 `LiMou_Research_Progress`） | 邮件 ASCII 副本 |
| 乱码名（密度启发式） | UTF-8 被按 latin-1 误解码的变体 |

### 乱码名怎么认

例：`2026å¹´11æ26æ¥ æçæ´ç ç©¶è¿å±.docx`

**不用 round-trip 校验**：这类名字在传递中被**有损截断**（丢了部分字节），`encode("latin-1").decode("utf-8")` 会抛 `UnicodeDecodeError`。

**改用密度启发式**：中文 UTF-8 字节集中在 `0xE4–0xE9`，误解码后会产生大量 `U+0080–U+00FF` 范围内的字符（`å ¹ ´ æ ¥ ç © ¶ è ¿ ±`）。正常英文名和正常中文名都不会落入这个区间。

```python
def _looks_mojibake(name: str) -> bool:
    if name.isascii():
        return False
    return sum(1 for ch in name if 0x80 <= ord(ch) <= 0xFF) >= 3
```

取舍：带重音符号的欧洲语言名（如 `Müller.docx`）可能误判，但中文文档盘场景概率极低。可在 config.json 扩展排除词表。

## 事实边界

记录是**文件级修改时间**，无法区分本人操作或协作者更新，也无法据此断定「当天新写」。文档末尾必须附这一口径说明。

禁止把推断写成既成事实，禁止记录以外的杜撰。确需推断的语句标注「待确认」。

## 鉴权与重试（仅灵犀版）

- 云文档 CLI 依赖会话注入的 `TMP_LX_UUID`（优先）或 `WPS_SID`
- 翻页时 token 换取接口会偶发 EOF，并**伪装成 `HTTP 401 Unauthorized`**；这不是凭据失效
- 已按「瞬时特征优先于鉴权判断」做最多 3 次重试（间隔 3s），翻页间再隔 1s；连续失败则退出码 3 停止
- 成功判定必须先于错误关键字匹配，且只对失败响应的 `error` 字段做鉴权判断，避免正常数据里的数字（时间戳、体积）误命中 `401` / `403`
