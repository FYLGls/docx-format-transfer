# docx-format-transfer v{VERSION}

按 Word 模板检测全套格式规则，逐条确认后一键套用到你的文档，并逐条验证。
三种用法任选，原件永不修改，结果写新文件。

## 用法一：网页版（推荐）

双击 **启动界面.bat**（或 `python scripts/serve_ui.py`），浏览器自动打开
http://127.0.0.1:8765 —— 右侧面板逐条列出检测到的每项格式要求（自动按执行顺序
排列，分页类标注"从上到下单遍检查、排最后"），可逐条勾选、改值、调整顺序，
点「开始修改」逐条执行并逐条显示验证结果。

**格式要求文件/图片（可选）**：老师发布的格式要求说明可直接上传
（docx/pdf/txt/md/图片截图，或点开"粘贴要求文字"直接粘贴）。
系统会解析出结构化规则（页边距/字体字号/行距/缩进/页码/题注/三线表/目录/代码等）
并以**更高优先级**合并进档案——与模板不一致处按要求执行，并在清单中用
"要求"徽章和琥珀色提示标明。解析不了的句子会列出原文，不静默丢弃。

**局域网/公网部署**（让其他设备浏览器访问）：

```bash
python scripts/serve_ui.py 8765 --host 0.0.0.0
```

启动时会打印局域网地址（如 http://192.168.1.5:8765），其他设备直接访问即可
（注意防火墙放行端口）。放到云主机/服务器上运行即为公网网页版。

## 用法二：命令行

```bash
cd scripts
python extract_format.py 模板.docx        # 检测 → 模板.profile.json / .report.md / .map.json
python extract_format.py 我的文档.docx     # 目标分析 + 歧义清单
python parse_requirements.py 格式要求.docx --merge 模板.profile.json   # 要求并入档案（可选）
python apply_format.py 我的文档.docx 模板.profile.json --dry-run   # 预演
python apply_format.py 我的文档.docx 模板.profile.json -o 输出.docx
python verify_format.py 输出.docx 模板.profile.json --original 我的文档.docx
```

进阶：`--rules 规则顺序.json`（自定义执行顺序/禁用项）、`--overrides 值.json`
（覆盖检测值）。首次使用先跑 `python check_deps.py` 体检依赖。

## 用法三：ZCode skill

双击 **install-skill.bat**，把本包注册为 ZCode 个人技能（安装到
`%USERPROFILE%\.agents\skills\docx-format-transfer`，重启 ZCode 生效）。
之后把模板和文档放进对话，说"按模板改格式"即可走完整确认流程。

## 能检测/套用什么

页边距、纸张、页码格式与位置、页眉页脚、各级标题样式与编号体系（自动/手打、
第X章/X.Y/一、/（一）、中文数字）、正文（中西文字体/字号/行距/缩进/对齐）、
图片尺寸与对齐、题注（位置/编号按章或连续/分隔符/字体）、三线表/网格表边框、
表头重复与行不拆分、单元格字体、公式字体（OMML）、代码块装框、参考文献版式、
目录域、交叉引用同步、分页行为（题注图表绑定、防孤行、每章另起页）。
模板中不存在的维度自动跳过；模板自相矛盾处会列出交你裁决。

## 已知边界

- 按章域编号的题注重排为字面编号（显示正确稳定；需 Word 域自动编号请在对话流中说明）
- "续表"以"行不拆分+跨页重复表头"等价实现（Word 无原生续表机制）
- MathType/图片公式、图内文字、参考文献字段内容：只检测提示，不自动改
- 分节重组（封面/罗马前置页）不自动新建，只在现有节上改页码格式

## 环境要求

Windows / macOS / Linux + Python 3.10+，`pip install python-docx lxml`。
可选增强：`pip install pypdf`（pdf 格式要求）、`pip install rapidocr-onnxruntime`
（图片格式要求 OCR）——缺失时仅该输入类型不可用，其余功能不受影响。
网页版仅本机使用时零配置；局域网/公网部署见上文。
