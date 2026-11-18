# 需求：deskhand 增加 `ppt` 服务（PowerPoint COM）

状态：已实现，2026-11-17。提出者：需求方（经 Codex，材料来自调用方项目的演示文稿工作）。实现者：deskhand 仓库的 agent。

## 1. 目标

新增第二个 app：`deskhand ppt`。它通过 COM 驱动 Windows 上的 PowerPoint 桌面客户端，从 WSL 提供 MCP 工具：读一份 deck 的结构与文字、按 1280x720 画布增删改元素、把幻灯片渲染成 PNG/PDF 给人看、另存为 `.pptx`。

一句话：把"做 PPT"里**必须由 PowerPoint 本体完成**的那部分收进 deskhand；内容、论点、数字口径不进来。

## 2. 为什么要收进来

2026-11-04/05 的一次 session 在调用方项目里做了 20 页和 5 页两份 deck，链路是：

| 步骤 | 当时怎么做的 | 现在应该在哪 |
| --- | --- | --- |
| 读用户手做的 pptx | 解 zip、正则读 slide XML，取文字、`a:off/a:ext` 坐标、`a:srcRect` 裁切、notes | deskhand `read_deck` |
| 定论点、页面计划、数字白名单 | `prompt.md` + `facts.md` + `check_numbers.py` | 留在调用方项目 |
| 从数据产物生成数字块 | `gen_blocks.py`，断言每个数字仍出现在源产物里 | 留在调用方项目 |
| 写页面 | 子 agent 写 `pages/NN.json`（演示文稿工具的 JSON） | deskhand 的元素级写工具 |
| 校验 | `build_deck.py --check`（id、越界、重叠、文字溢出） | 调用方算，deskhand 只给真实几何 |
| 看效果 | 浏览器自动化工具 1280x720 逐页截图，再 Read PNG | deskhand `render` |
| 出 pptx | `build_pptx.py`（python-pptx）+ `check_pptx.py` + `render.ps1`（PowerPoint COM 导 PNG） | deskhand 直接写 + 导 |

结论：**渲染和"读回真实 deck"只有 PowerPoint 能做**，而当时是三条互不相干的脚本（OOXML 解析、python-pptx、COM），每条都要自己处理路径、进程、超时和失败。deskhand 已经有这套东西（`core/powershell.py`、`apply_token.py`、`server.py`），所以这部分应该收敛进来，其余留在项目里。

## 3. 非目标（明确不要做）

- 不做 `facts.md`、数字白名单、禁用词表：这是调用方的内容纪律，deskhand 不认识"数字必须来自哪里"。
- 不做叙事、页面计划、主题审美、模板库。
- 不做 HTML deck / 浏览器截图。
- 不引入 python-pptx 作为写路径：一个文件只能有一个写者，否则"读回"和"写入"会不一致。python-pptx 只允许出现在测试里做只读断言。
- 不做 LLM 调用、不做联网。

## 4. 落点

```text
src/deskhand/ppt/
  server.py     create_ppt_server()，照 onenote/server.py
  tools.py      ToolSpec 列表与处理函数
  schemas.py    入参/出参 JSON Schema
  format.py     给人看的文本（列表、预览、错误）
  bridge.ps1    PowerPoint COM worker
  AGENTS.md     COM 的坑与不变量
```

- `cli.py` 的 `APPS` 加 `"ppt": _ppt`。
- `core/` 只在第二个 app 真的需要时改：`core/powershell.py` 的错误文案现在写死 OneNote，改成接受 app 名；`CALL_TIMEOUT_S = 40` 对 PowerPoint 冷启动加多页导出偏短，加一个 per-call timeout 参数（默认不变）。
- `tests/parity/fixtures/` 按 app 分目录，现有 `tools.json` 移到 `onenote/` 下。

## 5. 单位与坐标（最容易错的地方）

- 工具 API 一律用 **px @ 96 dpi**，画布默认 1280x720（= 13.333x7.5 in = 960x540 pt）。
- bridge 内部转 pt：`pt = px * 0.75`。PowerPoint 的 `Left/Top/Width/Height`、`SlideWidth/SlideHeight` 都是 pt。
- 读回时取整到 0.1 px，避免浮点噪声进 fingerprint。
- 颜色 `#RRGGBB`；COM 的 `RGB` 是 `0xBBGGRR`，转换只在一个函数里做。

## 6. 工具面

### 6.1 只读（不加 token）

| 工具 | 入参 | 出参 |
| --- | --- | --- |
| `presentations` | - | 打开的 deck：`name, path, slides, widthPx, heightPx, saved`；没开 PowerPoint 时 `running: false` |
| `read_deck` | `presentation?`, `slides?`（如 `"1-5"`）, `includeNotes?` | 每页 `index, layout, title, shapes[]`；shape 含 `id, name, type, x, y, w, h, text, paragraphs`；表格给 `rows/cols/cells`，图表给 `chartType/seriesNames/categories/series`，图片给 `path/widthPx/heightPx` |
| `read_slide` | `presentation?`, `slide` | 同上，单页；带 `fingerprint` |
| `render` | `presentation?`, `slides?`, `outDir`, `width?`, `height?`, `format?`(png/pdf), `overwrite?` | 写出的文件路径与尺寸；默认 1600x900 |

`render` 会写文件但不改 deck，所以不走 token；`outDir` 必须显式给，同名文件存在时除非 `overwrite: true` 否则拒绝。`overwrite: true` 会替换已有文件，因此 `render` 被标注为 `destructive_hint=true`。PDF 固定用 deck 的页面尺寸，`width`/`height`/`slides` 只对 PNG 有效，传给 PDF 会报错而不是静默忽略。

`read_deck` 默认每页只打一行概览（页号、版式、标题、shape 数），传 `slides` 才展开 shape；展开时每页最多列 5 个 shape。完整模型始终在 `structuredContent` 里。

### 6.2 写入（dry run -> applyToken -> apply）

| 工具 | 关键入参 |
| --- | --- |
| `create_deck` | `path?`, `width?=1280`, `height?=720`, `theme?{background,text,accent,font}` |
| `add_slide` | `presentation?`, `layout?`(blank/title/title-content/two-column/section), `position?`, `title?`, `subtitle?`, `background?` |
| `add_text` | `slide`, `x,y,w,h`, `paragraphs[]`（每段 `runs[]`，run 有 `text,bold,italic,color,size,font`；段有 `align,spaceBefore,spaceAfter,lineSpacing`）, `padding?`, `anchor?`, `fill?`, `line?` |
| `add_table` | `slide`, `x,y,w,h`, `header?[]`, `rows[][]`, `columnWidths?`, `headerFill?`, `fontSize?`, `headerFontSize?` |
| `add_chart` | `slide`, `x,y,w,h`, `type`(bar/column/line/pie), `categories[]`, `series[{name,values,color?}]`, `title?`, `legend?`, `valueAxisTitle?`, `dataLabels?`, `numberFormat?` |
| `add_picture` | `slide`, `x,y,w,h`, `path` 或 `dataBase64`, `fit?`(stretch/contain/cover) |
| `set_notes` | `slide`, `text` |
| `update_shape` | `slide`, `shapeId`, `x? y? w? h?`, `text?`, `paragraphs?`, `fill?`, `line?` |
| `delete_shape` | `slide`, `shapeId` |
| `move_slide` | `from`, `to` |
| `delete_slide` | `slide` |
| `save_deck` | `presentation?`, `path`, `overwrite?=false` |
| `close_deck` | `presentation?`, `discard?=true` |

规则：

- 所有写入默认 `dryRun: true`，返回 `applyToken` 和预览（"会在第 3 页 (320,96,640,120) 加一个 4 段文本框"）。预览第一行是 `applyToken:`，第二行是 `fingerprint:`；预览只显示摘要与有界的 shape 列表，完整数据在 `structuredContent`。
- 服务端门控不依赖客户端：没有 token 或 fingerprint 的 apply 一律拒绝。仓库里的 pi 扩展目前只识别 OneNote 工具名，PPT 写入还没有客户端确认对话框。
- `apply` 必须带同一个 `applyToken` 和 `expectFingerprint`；token 单次、10 分钟、按请求体绑定（照 `core/apply_token.py`）。
- fingerprint：目标页 = `sha256(layout | 每个 shape 的 id/name/type/几何/文字/填充/线条/表格/图表/图片链接 | notes)`；`create_deck` = `sha256("new:"+size+theme)`；`save_deck` = 目标文件的 `size+mtime` 或 `absent`，再与内存 deck 的 fingerprint 合并。
- 不一致返回 `changed_since_preview`，要求重新读。
- COM 在写入开始后失败，返回 `applied_unverified`，不要说"没写"。
- `dataLabels` 按 series 的每个 point 设置（`Series.HasDataLabels` + `Point.DataLabel.ShowValue`）；`numberFormat` 同时用于 value axis 和已启用的 data labels。
- 字号以 px 表示，但 PowerPoint 最小 1 pt（1.34 px）；小于这个值的 run/table 字号在 preview 阶段返回 `invalid_request`，不会发 token。
- 填色文本框不能当细分割线用：PowerPoint 会把文本框的最小高度撑到约 40 px，即使请求的 `h` 更小。

## 7. bridge 的不变量

1. `bridge.ps1` 必须能被 PS 5.1 解析、纯 ASCII、永远不 `exit`；`tests/unit/test_bridge_script.py` 的静态检查要扩到新文件。
2. 一个 resident `powershell.exe`，一次一个 op；PowerPoint COM 是 STA，不能并发。
3. **不要杀掉用户的 PowerPoint**：启动时记 `$was = (Get-Process POWERPNT).Count -gt 0`；只有"不是我们起的"且"没有打开的 presentation"才 `Quit`。
4. 不抢焦点：`Presentations.Open(path, ReadOnly, Untitled:=0, WithWindow:=0)`；需要窗口的操作把 `WindowState` 设最小化。图表必须有一个 presentation 窗口，`Ensure-ChartWindow` 会建窗口并最小化；图表数据直接写 `Series.Values`/`Series.XValues`，不走 `ChartData.Workbook`（那会起 Excel 且无窗口时失败）。最小化窗口仍会留在任务栏，这是 quiet-desktop 的未决冲突。
5. 关掉所有对话框：`Application.DisplayAlerts = ppAlertsNone`；`AutomationSecurity = msoAutomationSecurityForceDisable`。
6. 保存只写调用方给的新路径；`save_deck` 的 `path` 若等于已打开的源文件，必须 `overwrite: true` 才允许。PDF 用 `SaveCopyAs(path, 32)`，不改变已打开 deck 的路径。
7. 冷启动 PowerPoint 可能 5-10 s，`render` 20 页可能超过 40 s：读 10 s、写 30 s、`render` 55 s，并在返回里写实际耗时。
8. 路径：WSL 路径在 Python 侧转成 `C:\...`（`core/powershell.py` 已有 Windows temp 的处理，复用），bridge 只收 Windows 路径。

## 8. 错误码

`no_powerpoint`（COM 起不来）、`no_presentation`、`slide_not_found`、`shape_not_found`、`file_exists`、`file_locked`、`changed_since_preview`、`apply_requires_fingerprint`、`com_error`（附 HRESULT 与原文）、`applied_unverified`、`timeout`。

## 9. 验收

必须能用一个脚本完成：

1. `create_deck` 1280x720 -> `add_slide` x5 -> 每页 `add_text`（标题加正文，含一个加粗 run）-> 一页 `add_table` -> 一页 `add_chart`（bar，3 类 2 系列）-> 一页 `add_picture`（PNG）-> `set_notes`。
2. `render` 出 5 张 PNG，尺寸 1600x900；人看一眼没有重叠或溢出。
3. `read_deck` 读回：5 页、每页 shape 数量与几何和写入一致（±0.5 px）、文字逐字相同、notes 相同。
4. `save_deck` 到 `%TEMP%`，`close_deck`，再打开该文件 `read_deck`，内容一致。
5. 第二次运行能对已保存的 deck 做 `update_shape`（改位置和文字）并 `save_deck` 覆盖。
6. 全流程结束后：PowerPoint 进程数回到开始前的值，没有窗口留下；`uv run pytest` 与 `uvx ruff check src tests` 通过。

## 10. 测试

- 单测：假 bridge（照 `tests/unit/test_onenote.py`），覆盖每个工具的 dry run/apply、fingerprint 漂移、错误码映射、px<->pt 转换、预览文本。
- bridge 静态：PS 5.1 解析、ASCII、无 `exit`。
- live（`tests/live/ppt.py`，手工跑）：真实 PowerPoint，建 deck -> 渲染 -> 读回 -> 保存 -> 删临时文件；断言进程数恢复。
- parity：`tests/parity/fixtures/ppt/tools.json`，从 `dump_tools.py` 生成。

## 11. 分两刀落地

2026-11-17 的实现一次落地了全部 17 个工具，没有按下面的两刀拆开：写入链路、fingerprint、错误分类和 live 测试是同一套机制，拆开会让第二刀重复验证同一批不变量；需求方的验收脚本也同时需要表格、图表、图片、notes 和已保存 deck 的再编辑。

- **第一刀（先合）**：`presentations`、`read_deck`、`read_slide`、`render`、`create_deck`、`add_slide`、`add_text`、`save_deck`、`close_deck`。已经覆盖"读用户 deck、从零写文字页、看效果"。
- **第二刀**：`add_table`、`add_chart`、`add_picture`、`set_notes`、`update_shape`、`delete_shape`、`move_slide`、`delete_slide`。

## 12. 实现前要先回答的问题

1. 图表走原生 `Shapes.AddChart2`（要 Excel，本机 Excel 16.0 可用）还是画成形状？结论：用原生图表，但数据直接写 `Series.Values`/`Series.XValues`，不用 `ChartData.Workbook`；图表前确保存在一个最小化的 presentation 窗口。
2. `check_layout`（越界、重叠、文字溢出）要不要进 deskhand？建议**不要**：它是内容纪律，留在调用方；deskhand 只保证 `read_deck` 给出真实几何，调用方自己算。
3. 一个 server 管所有 deck，还是每 deck 一个？建议一个 server，`presentation` 用 name/path 寻址。

## 13. 评审

这是第二个 app，会碰到 deskhand 现有的 stance：写入门控、热延迟、安静桌面、契约连续性。已按 `AGENTS.md` 走一次 hydra-crux review，六位 reviewer 的结论与未决冲突记录在 `.agents/skills/deskhand-stance-registry/SKILL.md`。
