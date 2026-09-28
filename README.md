# docx-format-transfer

按 Word 模板检测全套格式规则，经用户逐项确认后一键套用到自己的文档。
独立技能项目：模板格式检测 → 清单确认 → 目标分析 → dry-run → 应用 → 验证，全程原件只读。
**v1.2.0 新增**：可上传老师发布的**格式要求文档/图片**（docx/pdf/txt/截图，或直接粘贴文字），
解析为结构化规则并以更高优先级合并——与模板不一致处按要求执行并明确标出。
分发形态对齐"码文"（code-for-word）：**可下载发布包 + 自托管网页版 + ZCode skill** 三用。

## 发布包（可下载）

`release/docx-format-transfer-v1.2.0.zip`（含 .sha256 校验和），解压即用：
- 双击 **启动界面.bat** → 网页版（自动开浏览器）
- 双击 **install-skill.bat** → 注册为 ZCode 个人技能
- `scripts/` 命令行全功能 + `check_deps.py` 依赖体检 + `README.md`（发布版说明）
重新构建：`python packaging/build_release.py`（版本号取自根目录 VERSION）。
要像码文那样挂到 GitHub/Gitee Release：把 zip 传到仓库 Releases、README 传 Pages 即可。

## 快速使用

**可视化界面（推荐）**：双击 `启动界面.bat`（或 `python skill/scripts/serve_ui.py`），
浏览器打开 http://127.0.0.1:8765 —— 右侧面板在检测时**逐条流式列出**每项格式要求，
自动按执行顺序排列（分页类标注"从上到下单遍检查，排最后"），可逐条勾选、改值、
▲▼ 调整顺序，点「开始修改」逐条执行并逐条显示验证结果。

**局域网/公网网页版**：`python skill/scripts/serve_ui.py 8765 --host 0.0.0.0`
（启动时打印局域网地址，任何设备浏览器可用；放云主机即公网版，注意防火墙放行端口）。

**命令行**：

```bash
cd skill/scripts
python extract_format.py 模板.docx          # → 模板.profile.json + 模板.report.md + 模板.map.json
python extract_format.py 我的文档.docx       # 目标元素映射 + 歧义清单
python apply_format.py 我的文档.docx 模板.profile.json --dry-run   # 预演
python apply_format.py 我的文档.docx 模板.profile.json -o 我的文档.formatted.docx
python apply_format.py 我的文档.docx 模板.profile.json --rules my_rules.json --overrides my_values.json  # 自定义顺序/值
python verify_format.py 我的文档.formatted.docx 模板.profile.json --original 我的文档.docx
```

对话用法：把两个文件交给装了本技能的 ZCode，说"按模板改格式"，它会走
SKILL.md 的确认流程（清单确认 → 歧义确认 → dry-run → 应用 → 验证）。

## 目录

```
skill/                     交付物（安装到 ~/.agents/skills/docx-format-transfer/）
├── SKILL.md               工作流编排
├── scripts/               extract / apply（25 条规则逐条执行）/ verify / serve_ui + ooxml_utils
│   └── ui/index.html      可视化界面（右侧逐条清单 + 排序 + 点击执行 + 逐条验证）
├── references/            profile-spec（档案 IR）/ detection-rules / apply-order / ux-playbook
└── assets/sample-profile.json
启动界面.bat               双击启动可视化界面（GBK+CRLF，cmd 原生兼容）
release/                   可下载发布包（zip + sha256）
packaging/                 build_release.py 构建脚本 + zip 内模板（bat/README/check_deps）
tests/                     测试资产（不随技能安装）
├── lib_fixture.py         夹具构造库（含域题注/三线表/OMML/浮动图/合并单元格）
├── gen_fixtures.py        20 个场景生成器（template+target+ground truth）
├── gen_ai_reports.py      10 种 AI 生成器风格 × 2 轮实验报告
├── run_eval.py            评测：extract / full 两阶段
├── run_ai_eval.py         真实模板 × AI 报告两轮评测
├── test_web.py            网页版 20 轮端到端测试
├── test_release.py        下载包 20 轮安装-运行测试
├── smoke_ui.py            界面服务器全链路烟雾测试
└── sanity.py              可打开性 + 幂等性
docs/                      需求-格式维度全表 / 设计-修改顺序与确认UX
```

## 质量状态（2026-09-28，v1.1.0）

- 20 个合成复杂场景：检测准确率 **114/114（100%）**，应用后 verify **全部 0 FAIL**
- **真实模板多轮实测**：课程教师模板 + 华科毕业论文模板 × 20 份 AI 生成器风格报告
  → 暴露并修复 11 类真实问题（见 `tests/ai_reports/REPORT-问题清单.md`）
- **双端发布测试**：网页版 20 轮（上传→流式检测→逐条执行 6 种条件变体→验证）
  与下载包 20 轮（全新解压→依赖体检→CLI 全流程→验证）全部通过，
  40 轮覆盖全部 40 份测试文档，外加 install-skill.bat 安装验证
- 内容零损坏断言：正文/题注标题/标题文字在应用前后逐字一致（编号与交叉引用归一化后）
- 输出可被 python-docx 正常打开；应用幂等可重跑

## 已知边界（v1）

- 按章域编号题注重排为字面编号（显示正确稳定；域自动编号需用户另要求）
- 续表 = cantSplit + tblHeader 等价实现（Word 无原生续表机制）
- OLE(MathType)/图片公式/图内文字/参考文献字段内容：只检测提示，不自动改
- 分节重组（封面/罗马前置页）不自动新建，只在现有节上改页码格式
- 渲染级分页复查需要 LibreOffice（未安装时降级为静态 keep 规则）
- 模板自身矛盾（说明文字 vs 实例、9:1 不一致）会以 flag 形式交用户裁决，不静默取值
