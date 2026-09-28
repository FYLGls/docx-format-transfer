# 检测规则手册（逐维度实现手段）

提取器 `extract_format.py` 的实现依据。所有规则 = **定位（XML/正则）→ 归纳（统计）→ 置信度**。
命名空间前缀：`w`=wordprocessingml 主命名空间，`wp`=drawing，`a`=drawingml，`m`=math，`v`=vml。

## 通用机制
- **解析**：`zipfile` 读 docx，`lxml.etree` 解析 `document.xml`、`styles.xml`、`numbering.xml`、`settings.xml`、`header*.xml`、`footer*.xml`、`footnotes.xml`。
- **样式继承解析**：段落直接 pPr/rPr → 段落样式 → basedOn 链 → docDefaults，四层合并取"最深定义值"。
- **域解析**：遍历段落子节点，`w:fldSimple/@w:instr` 与 `w:fldChar(begin)…w:instrText…w:fldChar(end)` 两种形态都收集为域列表。
- **众数归纳** `mode_with_conf(values, tol)`：组内容差聚类（cm±0.05 / pt±0.5 / 倍数±0.05）→ 最大簇为规则值；n≥3 且占比≥0.9 → high；n=2 或占比 0.5~0.9 → medium；n=1 或并列冲突 → low；离群样本写入 `flags`。
- **文本工具**：`para_text(p)` 拼接 `w:t`（跳过域指令文本 `w:instrText`）。

## A 文档级
- **页面**：每个 `w:sectPr` 读 `w:pgSz`（w/h twips ÷567 → cm；`orient="landscape"` 判横页；节间不一致 → `mixed_orientation=true`）。
- **页边距**：`w:pgMar` 的 top/bottom/left/right/gutter/header/footer，多节取众数（正文节优先：取最后一个 next-page 节）。
- **页码格式**：`w:pgNumType @w:fmt/@w:start`；同时读页脚 XML 中 PAGE 域 instrText（`PAGE \* ROMAN`）——WPS 渲染以 instrText 为准，二者冲突时以域为准并记 flag。
- **页码位置**：含 PAGE 域的页脚段落的 `w:jc`（center→footer-center；right + 奇偶设置 → footer-outside）。
- **页眉**：`settings.xml` 的 `evenAndOddHeaders`；`sectPr` 的 `titlePg`；页眉文字与 rPr 抽样；页眉段落若有 `w:pBdr/bottom` → has_rule_line。
- **分节结构**：按节序列输出 [封面(无页码), 前置(罗马), 正文(阿拉伯)] 的节角色推断：无 footer 域的首页节=封面；fmt=roman=front_matter；否则 body。
- **分栏**：`w:cols @w:num @w:space`。

## B 样式
- **标题识别**：段落 `w:pPr/w:outlineLvl`（0-based）或样式名匹配 `Heading\s*(\d)` / `标题\s*(\d)`；两者冲突时以 outlineLvl 为准。
- **自动编号标题**（Word 多级编号在 XML、文本看不到号）：无样式/大纲级别但有 `numPr`、ilvl<3、文本<30 字、无句内标点且加粗 → 按-ilvl 定级（如"实验目的"实际显示"一、实验目的"）。
- **列表项排除**（防误判为标题）：
  - 连续 ≥3 段"1./2./3."式无样式编号段 → 列表，降级正文；
  - 编号序列延续规则（**倒序遍历**实现连锁降级）：无样式"N. ××"的下一段是"N+1."开头的正文 → 列表项。
- **标题编号**：
  - 自动：`numPr` 的 numId → numbering.xml lvlText（`%1`→`{n}`；numFmt=chinese* 时→`{c}` 中文数字占位符）；
  - 手打正则族：`第X章`（数字→`第{n}章`，汉字→`第{c}章`）；点分数字按深度 `{n}`/`{n}.{n2}`/…（**编号后分隔符形态进 pattern**：`1、`→`{n}、`，`1.`→`{n}.`，`1 ××`→`{n}`）；`一、`→`{c}、`；`（一）`→`（{c2}）`（次级中文计数）。
- **标题格式**：按级别聚合继承解析后的 rPr/pPr（字体三槽、sz 半磅÷2、b、jc、spacing、pageBreakBefore、keepNext/keepLines）。
- **特殊标题**（摘要/Abstract/关键词/目录/参考文献/致谢/附录/结论/结束语，英文大小写不敏感）不计入标题层级统计。
- **正文**：样式为 Normal（或正文最多样式）且长度>30 字、非标题非题注非列表的段落为样本集，聚合众数。
- **特殊段落**：文本匹配 `^(摘\s*要|ABSTRACT|关键词|Key\s*words|目\s*录|参考文献|致\s*谢|附录)` 的段落及其样式记录。

## C/D 图与题注
- **图片**：`w:drawing` 下 `wp:inline`（嵌入）与 `wp:anchor`（浮动，记 flag）；宽度取所有 `wp:extent @cx` ÷360000 → cm。
- **题注识别**（两种来源）：
  - 域：段落域 instr 含 `SEQ\s*(图|表|式|Figure|Table)`。
  - 手打正则：`^(图|表|式|Figure\.?|Table\.?|公式)\s*([0-9]+(?:[-－\.．][0-9]+)?)\s*([:：\-—．]?)\s*(.+)$`。
  - **句子排除守卫**：标题部分含 `。；`、或含 `，` 且全文>30 字 → 是"图1结论：…"式分析句，不是题注。
- **题注字体取值**：段落无直接子 run（文字全在域内）时回退到后代 run（`fldSimple` 内的 `w:r`）。
- **位置**：题注段落的前/后 2 块内的图表相对位置，统计众数。
- **编号体系**：域含 `STYLEREF` → chapter_linked=True（**域字面结果可能是未更新的旧值，不作数**）；或编号文本含 `-`；纯 SEQ/单数字 → 连续。
- **分隔符**：捕获组归类：半角空格/全角空格/冒号/横线/点/无。
- **交叉引用**：域 instr 含 `REF `_`_Ref` 计数；正文文本正则计数（题注段自身不计）。

## E 表格
- **边框形态规则**（不止三线，模板不是三线时同样要能规整目标）：每表分类
  three-line（上下框线+表头下线、无竖线）/ grid（含内框线）/ outline（仅上下框线）/ none / other，取众数为 `border_rule`；`three_line.enabled` = 规则为 three-line。
- **三线表判定**：表级 `tblBorders` top/bottom sz≥6 且 insideV/left/right=nil，且 insideH=nil 或仅表头行有单元格级下边框；线宽实测（非硬编码）。
- **跨页属性**：首行 `tblHeader`（repeat_header 众数）；各行 `cantSplit`。
- **续表**：全文段落文本正则 `续(上)?表`。
- **表内字体**：单元格段落 rPr 众数（继承穿透）。
- **表宽**：`tblW @type(pct/dxa/auto)` + `@w`。
- **嵌套/合并**：`gridSpan`、`vMerge`、嵌套 tbl 计数记 flags（应用时结构不动）。

## F 公式
- **载体统计**：`m:oMath`（含 `m:oMathPara` 判独立行）、`w:object`（OLE，progId 不可读时按存在计）、`w:pict`（图片公式）。
- **对齐/编号**：含 oMathPara 的段落 jc；段内文本正则 `\((\d+)([-–]\d+)?\)$|式\s*\d+`。
- **OMML 字体**：`m:r` 下 `w:rPr` 或 `m:rPr` 的字体字号众数（缺省 Cambria Math）。

## G 代码
- **识别信号**（段落级打分，≥2 分入候选）：
  - ascii 字体 ∈ 等宽集合{Consolas, Courier New, Cascadia*, JetBrains*, Source Code*, Menlo, Monaco} +2
  - 段落 `w:shd @w:fill` 非 auto/FFFFFF +1
  - 段落 `w:pBdr` 存在 +1
  - 文本代码特征密度（`;{}()=<>#` 占比>8% 或行首 2+ 空格缩进）+1
  - 连续候选 ≥2 段聚合为块。
- **装框方式**：块外层是单格 `w:tbl` → single-cell-table；段自带 pBdr/shd → paragraph-border；两者皆无 → none。
- **附录代码**：块位于"附录"标题之后的，标记 appendix=true（应用时不移动位置）。

## H 其他
- **参考文献**："参考文献"标题后至下一标题前的段落：正则 `^\[(\d+)\]` / `^\((\d+)\)` / `^(\d+)[\.、]` 定编号格式；`w:ind @w:hanging/hangingChars` 定悬挂缩进。
- **目录**：域 instr `TOC\s+\\o\s+"(\d)-(\d)"` → mode=field+depth；无域但有"目录"标题+前导符制表位 → manual。
- **脚注**：`footnotes.xml` 中非分隔符项的 rPr 众数。
- **孤行**：`w:widowControl` 众数。
- **说明文字矛盾检测**：模板正文里"标题X号/正文X号"字样与实例实测字号（号数映射表）不一致 → flag 交用户裁决（真实教师模板常见：说明写五号、实例是四号）。
- **水印**：header 内 `v:shape` 含 watermark；**行号**：`w:lnNumType`。

## 目标文档侧（元素映射器，同一套检测器复用）
提取器对目标文档额外输出 `element_map`（不进 profile）：
- `headings`: [{index, level, text, current_numbering}]
- `captions`: [{index, kind: figure|table, text, parsed{label,num,sep,title}}]
- `figures` / `tables` / `code_blocks` / `formulas`: 块索引与位置
- `ambiguities`: 无法自动定级的标题（无样式+无编号文本）、疑似题注的非题注段等 → 交模型问用户。
