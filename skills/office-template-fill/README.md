# 老式 Office 模板保格式填充 (office-template-fill)

Windows 下把生成内容填入老式 Word `.doc/.docx` 与 Excel `.xls` 模板且**不破坏原格式**的完整方法论，附带一套已脱敏的研发项目文档模板（11 份）。

## 核心能力

- **Word COM (pywin32)**：.doc → .docx 批量转换、Range 精确填充、自动编号清理、合并单元格处理
- **python-docx**：.docx 表格结构探测与单元格重写
- **xlrd + xlutils + xlwt**：老式 .xls 写入（含 xlwt 在 Python 3.13 的崩溃修复）
- **PDF 预览核验**：导出 PDF → 渲染页图 → 逐页目检
- **污染自检门禁**：交付前扫描旧项目文号/姓名/编号残留，FORBIDDEN 词零命中才放行

## 附带模板（assets/templates/，已脱敏）

| 文件 | 用途 |
|---|---|
| 01 项目建议书.doc | 项目立项建议 |
| 02 项目任务书.doc | 任务下达与阶段排期 |
| 03 设计开发方案.doc | 方案正文（含自动编号骨架） |
| 04 设计开发计划书.doc | 开发计划 |
| 05 设计开发输入清单.doc | 输入清单 |
| 07 设计开发评审报告.doc | 评审记录 |
| 08 设计开发验证报告.doc | 验证记录 |
| 09 设计开发输出清单.doc | 输出清单 |
| 10  工时分配表.xls | 工时表（数据已净化） |
| 12 客户试用报告.doc | 试用反馈 |
| 13 项目结项报告.doc | 结项验收 |

> 模板内预印的文号/序号/日期/结论均为**格式样例**，复用前必读 `references/template-manifest.md` 逐项覆写，交付前跑 `scripts/check_contamination.py`。

## 快速开始

```bash
# 1. .doc 转 .docx
python scripts/convert_doc_to_docx.py <目录或单文件> [<输出目录>]

# 2. 填前探测（结构 + 自动编号）
python scripts/dump_docx_structure.py <文件.docx> --full
python scripts/probe_list_numbering.py <文件.doc>

# 3. 填充（按格式选路线，详见 SKILL.md 工作流）

# 4. PDF 预览核验
pwsh -File scripts/export_pdf_preview.ps1 -Docs <文档...> -OutDir <目录>

# 5. 污染自检（交付门禁）
python scripts/check_contamination.py <全部交付文档>
```

## 依赖

- COM 路线：宿主机安装 MS Word/Excel，`pip install pywin32 pymupdf`
- .xls 路线：`pip install xlrd xlutils xlwt`
- 详细约束与踩坑清单见 [SKILL.md](SKILL.md)
