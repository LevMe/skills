# 标准中文文档格式

一个用于 Codex 的中文 DOCX 非内容格式控制 skill。

它面向中文毕业论文、正式报告、方案和说明书，统一控制字体、字号、页边距、行距、标题层级、目录、图表、公式、引用、参考文献、附录和页码；不擅自改写正文内容，也不臆造参考文献。

## 功能特点

- 双模式：论文模式和通用正式文档模式。
- 默认 A4 纵向，上、左、右边距 2 cm，下边距 2.7 cm。
- 正文小四号宋体、两端对齐、首行缩进两字符、1.5 倍行距；代码和接口标识使用左对齐样式。
- 标题使用真实的 Heading 1、Heading 2、Heading 3 样式和模板多级编号，均为不加粗黑体；一级居中，二级和三级左对齐。
- 一级标题使用 1.5 倍行距、段后 1 行；二级和三级使用 1.5 倍行距、段前段后为 0，所有标题缩进和制表位清零。
- 多级编号使用 Word `isLgl` 规则，一级显示“第一章”，二、三级显示“1.1”“1.1.1”，避免“二.1”。
- 摘要、英文摘要、目录、图表目录、参考文献和附录按实际内容条件生成。
- 图题和表题使用五号楷体；图像优先使用行内对象；表格使用可见网格。
- 图题、表题和图表目录条目使用 1.25 倍行距。
- 公式使用原生可编辑 OMML，不使用 MathType OLE 或原始 LaTeX。
- 自带只读结构审计脚本，检查 DOCX 结构和常见格式问题。

## 安装

将本目录复制到 Codex 用户技能目录：

Windows：

    C:\Users\你的用户名\.codex\skills\chinese-docx-format

macOS 或 Linux：

    ~/.codex/skills/chinese-docx-format

目录中应保留 SKILL.md、agents、references、assets 和 scripts。

## 使用

显式调用：

    $chinese-docx-format

    请按标准中文文档格式排版这个 DOCX，只调整非内容格式，不改写正文。

自然语言触发词包括：

- 标准中文文档格式
- 按标准中文文档格式排版
- 按标准中文文档格式生成 DOCX
- 标准中文 DOCX 格式

毕业论文或学位论文建议明确写“论文模式”；报告、方案或制度建议明确写“通用正式文档模式”。

已有 DOCX 应迁移到新输出文件，不要覆盖原文档：

    python chinese-docx-format/scripts/format_chinese_docx.py 输入.docx --output 输出.docx --profile auto

该脚本以空白基线模板承载迁移后的内容，清除源文档的标题样式、节属性和直接格式，重建标题编号，保留正文、表格、图片和 OMML。源文档已有 MathType/OLE 时不会静默删除或伪装转换，会在审计中标记。新文档使用 `assets/chinese-docx-base-template.docx`，输入和输出路径必须不同。

## 结构审计

使用 Codex 文档运行环境中的 Python 执行：

    python chinese-docx-format/scripts/audit_chinese_docx.py 输入.docx --profile auto --require-baseline --json 审计报告.json

审计脚本只读检查，不会自动修改输入文档，主要检查页面和节、样式、标题编号、条件模块、图表、OMML、引用、参考文献和元数据警告。

## 视觉验收

结构审计不能代替视觉验收。Windows 有 Microsoft Word 时，执行：

    powershell -File chinese-docx-format/scripts/render_docx_windows.ps1 -InputPath 输入.docx -OutputDirectory 渲染目录

脚本使用 Windows PowerShell 5.1 调用 Microsoft Word 原生导出 PDF，再用环境中的 Poppler 转为页面图片并逐页检查；论文模式会保留用户页眉文字并使用页眉距 1.5 cm、页脚距 1.75 cm的节属性；所有模式不启用前置罗马数字页码。LibreOffice 不是硬依赖。脚本会先检查 Word 是否能创建临时文档，失败时直接报告工作文件环境问题。没有可用 Word 渲染器时，只能交付结构检查结果，并明确标注“视觉验收未完成”。

## 文件说明

| 文件或目录 | 作用 |
| --- | --- |
| SKILL.md | 强制规则和执行流程 |
| agents/openai.yaml | 自动调用和界面元数据 |
| references/format-spec.md | 详细格式规范 |
| references/qa-checklist.md | 交付前检查清单 |
| assets/chinese-docx-base-template.docx | 去内容化的 Word 基线模板 |
| scripts/build_base_template.py | 可重复生成干净基线模板 |
| scripts/format_chinese_docx.py | 将已有 DOCX 迁移到新输出文件 |
| scripts/audit_chinese_docx.py | 只读结构审计器 |
| scripts/render_docx_windows.ps1 | 使用 Microsoft Word 导出 PDF 的视觉验收脚本 |

## 适用边界

本 skill 只处理 DOCX。用户提供的学校规范或新模板优先于本 skill；参考论文只用于补充未明确的版式细节，不会复制其中的个人信息、学校信息、正文或嵌入对象。
