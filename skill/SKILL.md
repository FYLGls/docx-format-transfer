---
name: docx-format-transfer
description: 按 Word 模板文档检测全套格式规则并一键套用到用户的文档。Use whenever the user wants to 套模板格式/按模板改格式/格式体检/对齐论文或比赛格式要求，or mentions template docx formatting, format checking, caption/三线表/页边距/标题格式 alignment. Even if they just say "把这个文档改成模板的格式" — this skill handles detection, confirmation, application, and verification.
---

# docx-format-transfer：模板格式检测与一键套用

把「模板.docx + 用户的文档.docx」变成「格式完全对齐模板的新文档」。流程固定为：
**提取模板 → 清单确认 → 目标分析+歧义确认 → dry-run → 应用 → 验证 → 交付**。
任何一步都不要跳过确认直接改文件；**原件永不修改，输出写新文件**。

## 可视化界面（用户要"逐条列出/可排序/点击修改"时优先推荐）

```
python "$SCRIPTS/serve_ui.py"        # 打开 http://127.0.0.1:8765
```
界面右侧会在检测时**逐条流式列出**每项格式要求（自动排序：容器→样式→内容→
分页类标注"从上到下单遍、排最后"→收尾），用户可勾选/去勾选、直接改值、▲▼ 调整顺序
（分页项提前会显示风险提示），点「开始修改」后逐条执行并逐条显示验证结果。
本项目根目录的 `启动界面.bat` 可双击启动。

## 脚本

```
SCRIPTS = <本技能目录>/scripts
python "$SCRIPTS/extract_format.py" <doc> [-o profile.json] [--report report.md] [--map map.json]
python "$SCRIPTS/apply_format.py"  <target.docx> <confirmed_profile.json> [-o out.docx] [--map map.json] [--dry-run] [--report changes.md]
python "$SCRIPTS/verify_format.py" <out.docx> <confirmed_profile.json> [--original target.docx] [--report verify.md]
```

- `extract_format.py` 对任何 docx 输出三件套：profile.json（格式档案）、report.md（人类可读清单）、map.json（元素映射：标题/题注/图/表/代码块清单 + 歧义列表）；`profile_steps()` 支持流式逐维度输出（界面用）。
- `apply_format.py` 按确认后的档案**逐条规则**改写（RULES 目录 25 条：容器→样式定义→内容元素→分页布局→域收尾），支持 `--rules`（自定义顺序/禁用）与 `--overrides`（值覆盖），幂等可重跑；单条规则失败不中断整体并记录。
- `verify_format.py` 重提取比对档案 + 内容零损坏断言（正文/题注标题/标题文字逐字一致）。
- `serve_ui.py` 本地可视化界面（127.0.0.1，无需额外依赖）。
- 修改顺序与理由见 `references/apply-order.md`；逐维度检测手段见 `references/detection-rules.md`；档案 schema 见 `references/profile-spec.md`。

## 工作流（严格按序执行）

### 第 1 步：输入体检
1. 确认两个文件：模板 docx + 目标 docx（用户没给就先要）。`.doc` 老格式或加密文档 → 停止并告知（加密不猜密码；.doc 请先另存 .docx）。
2. 跑 `extract_format.py` 对模板。

### 第 2 步：清单确认（参照 references/ux-playbook.md）
1. 把 report.md 的内容以对话消息呈现（分组表格，只说人话：✅已确认 / ⚠️基本确定 / ❓需要你拍板）。
2. 只把 ❓ 项逐个问用户（选择题，一次 ≤4 问）；⚠️ 项汇总一句"有异议再提"。
3. 用户口语修改（"7 改黑体三号"）→ 你解析成 JSON patch 改 profile → **复述确认后**存为 `<目标stem>.confirmed.json`。
4. 模板 flags（如 9:1 不一致）必须原样转述，让用户裁决。

### 第 3 步：目标分析与歧义确认
跑 `extract_format.py` 对目标文档，读 map.json：
- `ambiguities` 非空（无样式标题、疑似题注等）→ 列出行号+前 20 字，让用户确认或纠正；把纠正写进一个 mapping JSON（`{"headings": [{"idx":3,"level":1}, ...]}` 合并进 map 的对应项）。
- 目标某维度完全缺失（如没有公式）→ 该类不修改，交付时明说。

### 第 4 步：dry-run
`apply_format.py <target> <confirmed.json> --map <修正后的map.json> --dry-run` → 把"将修改 N 处（分类计数）+ 低置信明细"给用户，得到认可后才执行。

### 第 5 步：应用与验证
1. 正式跑 apply（不带 --dry-run）。
2. 跑 `verify_format.py <输出> <confirmed.json> --original <目标原件>`。
3. FAIL=0 才交付；有 FAIL → 读报告定位 → 修 profile 或 map → 重跑（幂等）→ 再验证。
4. 交付：输出文件 + changes.md + verify 报告，提醒「Word 里 Ctrl+A → F9 刷新目录/域」。

## 硬性规则
- **原件只读**。输出文件名带 `.formatted.docx` 后缀，重跑只覆盖上次的输出。
- dry-run 未获用户认可不得正式应用。
- `dimensions_absent` 里的维度（模板没有的）一律跳过，不拿默认值硬套。
- OLE 公式 / 图片公式 / 图内文字 / 参考文献字段内容：只检测提示，绝不自动改。
- "续表"用 cantSplit+tblHeader 等价实现并在报告里说明。
- 按章域编号的题注：v1 重排为字面编号（显示正确且稳定），报告里说明。
- 超大文档（>150 页）告知耗时，正常执行。

## 常见追问
- "能改成域编号吗" / "想要三线表线宽不同"：改 confirmed profile 对应字段重跑 apply 即可。
- "模板和目标反了"：重跑第 1 步即可，流程对称。
