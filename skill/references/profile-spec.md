# 格式档案（Format Profile）IR 规范 v1.0

格式档案是本技能的**中间表示**：提取器产出它、用户确认修改它、应用器消费它、验证器比对它。设计为格式无关（未来 LaTeX 提取器/应用器可复用同一 schema）。

## 总则
- 单位统一：长度 `cm`、字号 `pt`、缩进优先用**字符数**（中文文档标准）、EMU/twips 只在脚本内部出现。
- 每个规则叶子可带 `"confidence"`（high/medium/low）与 `"evidence"`（如 `"11/12 张图宽度一致"`）。low 必须在确认环节问用户。
- 模板中不存在的维度记入顶层 `"dimensions_absent"`（如模板无公式 → `"formulas"`），应用器对 absent 维度**跳过该类修改**（不改目标文档的公式）。
- 模板内部冲突（9:1 不一致）取众数为规则值，并在 `"flags"` 中记录异常，清单里展示给用户裁决。
- `"schema_version"` 变更需同步更新提取器/应用器/验证器三方。

## 结构树（关键字段与枚举）

```
profile
├─ meta
│  ├─ source_file / extracted_at / generator
│  └─ doc_stats{sections, paragraphs, images, tables, captions_figure, captions_table,
│               formulas_omml, formulas_ole, code_blocks}
├─ page
│  ├─ size{width_cm, height_cm, orientation: portrait|landscape, mixed_orientation: bool}
│  ├─ margins{top_cm, bottom_cm, left_cm, right_cm, gutter_cm}
│  ├─ header_footer_distance{header_cm, footer_cm}
│  └─ columns{count, space_cm}
├─ page_numbers
│  ├─ cover{has_page_number: bool}
│  ├─ front_matter{format: decimal|lowerRoman|upperRoman|chineseCounting, start, position: footer-center|footer-outside|header-right…}
│  └─ body{format, start, position, font, size_pt}
├─ headers{even_odd_different, first_page_different, content_font, content_size_pt, has_rule_line}
├─ headings
│  ├─ levels_in_use: [1,2,3]
│  └─ levelN{
│       style_name, font_eastasia, font_ascii, size_pt, bold, italic, align,
│       space_before_pt, space_after_pt,
│       line_spacing{rule: auto|exact|atLeast, value_pt_or_multiple},
│       page_break_before, keep_next, keep_lines,
│       numbering{mode: auto-numbering|manual, pattern: "第{n}章"|"{n}.{n2}"|"一、",
│                 chapter_linked: bool, confidence}}
├─ body{font_eastasia, font_ascii, size_pt, line_spacing{…},
│       first_line_indent_chars, align: both|left|center|justify,
│       space_before_pt, space_after_pt, widow_control}
├─ figures{width_rule{type: fixed|pct-page-width, value_cm|value_pct},
│           align, keep_with_caption, space_before_pt, space_after_pt}
├─ captions
│  ├─ figure{label:"图", position: below|above,
│  │         numbering{mode: field-seq-styleref|field-seq|manual, chapter_linked},
│  │         format_template:"图{chapter}-{index}{sep}{text}", separator:" "|"　"|":"|"-"|"",
│  │         font_eastasia, size_pt, bold, align}
│  ├─ table{… position: above|below …}
│  └─ equation{template:"({chapter}-{index})"|"式{chapter}-{index}", position: right}
├─ tables{three_line{enabled, top_bottom_pt, header_underline_pt},
│          header_row{bold, align, shading_fill},
│          repeat_header, cant_split_rows, keep_caption_with_table,
│          continuation_convention{detected, label:"续表"},
│          cell_font_eastasia, cell_font_ascii, cell_size_pt, cell_align,
│          width_mode{type: pct|dxa|auto, value}}
├─ formulas{omml_font, omml_size_pt, align, number_position, number_template,
│            ole_detected: bool}
├─ code_blocks{box_style: single-cell-table|paragraph-border|none,
│               font_ascii, size_pt, line_spacing_single: bool, shading_fill}
├─ bibliography{detected, numbering_format:"[{n}]"|"({n})"|"{n}.",
│                hanging_indent_chars|cm, size_pt, align}
├─ toc{mode: field|manual, depth, leader: dot|none}
├─ cross_references{mode: field|text, count}
├─ flags: [字符串列表，记录模板异常/警告]
└─ dimensions_absent: [formulas|code_blocks|toc|bibliography|…]
```

## 语义约定
- `numbering.pattern` 中 `{n}`=章序、`{n2}`=节序；`chapter_linked=true` 表示编号含章号（图3-1）。
- `captions.numbering.mode` 三态：`field-seq-styleref`（域+按章）、`field-seq`（域+连续）、`manual`（手打，应用时需维护计数器重排）。
- `line_spacing`：`rule=auto` 时 value 为倍数（1.5）；`exact/atLeast` 时 value 为磅。
- `page_numbers.body.position` 由页脚段落 jc 推断；`cover.has_page_number=false` 由 titlePg+空页脚或独立无页码节推断。

## 消费方式
- **确认环节**：`report.md` 是档案的人类可读投影；用户自然语言修改由模型解析为 JSON patch 后写回 confirmed profile。
- **应用器**：只认 confirmed profile；`--dry-run` 输出计划修改清单（按类别计数+低置信明细）不落盘。
- **验证器**：对输出文档重提取，逐叶子与 confirmed profile 比对（数值容差：cm±0.05、pt±0.5、倍数±0.05）。
