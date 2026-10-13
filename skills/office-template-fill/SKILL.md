---
name: office-template-fill
description: Windows 下老式 Word .doc/.docx 与 Excel .xls 模板的保格式批量填充经验（源自生产项目踩坑），并附带公司研发项目文档模板 10 份（项目建议书/任务书/设计开发方案/计划书/输入输出清单/评审验证报告/客户试用报告/结项报告/工时分配表）。涵盖 Word COM (pywin32) doc→docx 转换与 Range 填充、自动编号与合并单元格陷阱、xlrd/xlutils/xlwt 写 .xls、PDF 预览核验、旧内容污染自检。Use when filling Office document templates (project forms, reports, checklists, 工时表) with generated content while preserving the original template formatting, especially legacy binary .doc/.xls files on Windows.
---

# 老式 Office 模板保格式填充（Windows）

适用：把生成内容填入 Word/Excel 模板且**不能破坏原格式**，尤其是老式 .doc/.xls 二进制格式。
技术路线：**Word COM (pywin32)** 负责 .doc 转换与格式敏感填充；**python-docx** 负责 .docx 结构读取与单元格重写；**xlrd + xlutils + xlwt** 负责 .xls 写入。

若是全新 .docx/.xlsx 文档的创建或普通编辑（不涉及老式格式），优先用内置 `office-docx` / `office-xlsx` skill；本 skill 专管遗留格式与 COM 保格式填充。

## 工作流

0. **取用模板**（复用公司表单体系时）：`assets/templates/` 内含研发项目文档模板 10 份（9 份 .doc + 净化版工时表 .xls）。复制到项目 `模板/` 目录后开工；**填前必读 `references/template-manifest.md`**——模板预印的具体值（编号/日期/样例数量/预写结论）全是格式样例，须按清单覆写。
1. **转换**：.doc → .docx 批量转换
   `python scripts/convert_doc_to_docx.py <目录或单文件> [<输出目录>]`（Word COM SaveAs2 FileFormat=16）
2. **探测**（填之前必做，别直接猜坐标）：
   - 结构：`python scripts/dump_docx_structure.py <文件.docx> [--full]` —— body 块顺序、表格逻辑坐标、合并单元格去重
   - 自动编号：`python scripts/probe_list_numbering.py <文件.doc|docx> [--all|--contains 文本]` —— 凡模板出现"（一）（1）"式编号必读
3. **填充**（按目标格式选路线，细节读对应 reference）：
   - Word COM Range 填充 → `references/word-com-fill.md`
   - python-docx 单元格重写（仅 .docx）→ `references/python-docx-rewrite.md`
   - .xls → `scripts/fill_cells_xls.py`（作为库导入）+ `references/xls-fill.md`
4. **核验**：`pwsh -File scripts/export_pdf_preview.ps1 -Docs <文档...> -OutDir <目录>` 导出 PDF → PyMuPDF 渲染页图 → read_image 逐页检查 → `references/preview-verify.md`
5. **污染自检（交付门禁）**：`python scripts/check_contamination.py <全部交付文档>` —— FORBIDDEN 词零命中才能交付；REVIEW 词命中须逐一说明合理。新项目开工时用 `--extra` 把本项目将废弃的旧值追加进黑名单。

## 硬性约束（动手前通读，均为生产验证过的坑）

- **模板预印具体值 = 格式样例**：空格填充模式不会触碰预印值，不覆写就会带进新项目文档。按 `references/template-manifest.md` 的必改清单逐项覆写，交付前跑污染自检。
- **`Find.Execute(Replace=2)` / `ReplaceWith` 在 pywin32 下静默失效**（返回 True 但不替换）。必须 `Find.Execute()` 命中后对命中 Range 直接改 `.Text`。
- **自动编号区域写多行文本会产生双重编号错乱**：展开的新段落继承编号。编号只保留在标签段，正文段逐段 `ListFormat.RemoveNumbers(1)`。
- **`Table.Cell(r,c)` 的 c 是行内第 N 个合并单元格**，不是物理列号。先探测坐标再填。
- **段落索引随写入漂移**：同一单元格内多处写入时从后往前写。
- **`ExportAsFixedFormat` 目标 PDF 已存在时静默失败**（不报错、旧文件保留）。先删旧文件。
- **xlwt 在 Python 3.13 保存崩溃**：模板 `_num_formats` 含 None 键。xlutils.copy 后删除该键（`scripts/fill_cells_xls.py` 已内置修复）。
- **绝不用 COM Range.Copy 搬运中文值到 .xls**：会污染同列其他单元格（整行被覆写成 46023.0 之类）。用 xlwt 直接写值。
- COM 会话必须 `finally: word.Quit()`；结束后清理残留 `~$` 锁文件。
- 路径含中文或 `\01` 类数字段：Python 里用 r-string 或正斜杠（`\01` 是八进制转义）；不要在 `python -c` 里内联中文路径；用 Python 生成 PS1 时以 `utf-8-sig` 编码写文件。

## 依赖

- COM 路线（转换/填充/导出 PDF）：宿主机需装 MS Word；`pip install pywin32 pymupdf -i https://mirrors.cloud.tencent.com/pypi/simple`
- .xls 路线：`pip install xlrd xlutils xlwt -i https://mirrors.cloud.tencent.com/pypi/simple`（DSH 内置 Python 不含这三个库，用宿主 Python）
- python-docx：DSH 内置 Python 自带，`load_workspace_dependencies` 取路径
- **宿主 Python 必须用绝对路径调用**（如 `C:\Program Files\Python310\python.exe`，前提是该解释器已装 pywin32）：裸 `python` 常解析到 Microsoft Store 别名 stub，会间歇性静默失败（exit 9009、无输出）。先 `where python` 确认实际解释器再写进脚本。

## 参考文档（按需读取）

| 文件 | 何时读 |
|---|---|
| `references/template-manifest.md` | 使用 assets/templates 模板前必读：逐份模板的格式样例值、必改清单、软样板确认项、污染黑名单 |
| `references/word-com-fill.md` | 用 COM 填 Word 模板：空格/续填/整格替换写法、Find 替换正确姿势、编号去除、合并单元格、索引漂移 |
| `references/python-docx-rewrite.md` | 用 python-docx 重写 .docx 表格单元格：合并去重、段落重建、加粗控制 |
| `references/xls-fill.md` | 写 .xls：样式规则、工时表块结构、合计对账方法 |
| `references/preview-verify.md` | PDF 导出 + 页图渲染核验清单、路径/编码坑、收尾清理 |
