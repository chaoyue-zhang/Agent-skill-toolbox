# PDF 预览核验

流程：填充完成 → Word COM 导出 PDF → PyMuPDF 渲染页图 → read_image 逐页目检。
文本回读通过≠格式正确：双重编号、合并格溢出、字体丢失只在渲染图上暴露，**每份填充完的文档都要走一遍**。

## 导出 PDF

用 `scripts/export_pdf_preview.ps1`（已内置防覆盖与只读打开）：

```powershell
pwsh -NoProfile -File scripts/export_pdf_preview.ps1 -Docs "D:\产物\01 项目建议书.docx","D:\产物\02 项目任务书.docx" -OutDir "$env:TEMP\preview"
```

- **坑（生产验证）**：`ExportAsFixedFormat(pdf, 17)` 在目标 PDF **已存在同名文件时静默失败**——不报错、保留旧文件，导致"核验的是上一版"。脚本已先删旧 PDF；手写逻辑时必须同样处理（或换新名/带版本号）。
- 文档以 ReadOnly 打开：用户 Word 会话占着文件（或有 `~$` 锁文件）也不阻塞导出。
- 也可用内置 `office-docx` skill 的 LibreOffice 渲染，但 Word COM 导出的分页与 MS Word 完全一致，核验保真度更高；两者分页可能有差异，别混用结论。

## 渲染页图

```python
import fitz   # PyMuPDF: pip install pymupdf -i https://mirrors.cloud.tencent.com/pypi/simple

doc = fitz.open(r'D:/preview/01 项目建议书.pdf')
for i, page in enumerate(doc):
    page.get_pixmap(dpi=150).save(f'page-{i+1:02d}.png')
```

用 read_image 工具逐页查看，核对清单：

- [ ] 编号序列正常：无"（十）（1）…"双重编号，标签段编号连续
- [ ] 合并单元格内容未溢格、未串格
- [ ] 中文无丢字体（不出现方块/宋体替身）
- [ ] 该留空的签名栏、审批栏、日期栏、公章位仍是空白
- [ ] 表格列宽行高未被长文本撑爆；换行位置正常
- [ ] 页眉页脚、页码未被填充操作破坏

## 路径与编码坑

- 导出/渲染脚本落成**脚本文件**执行；不要在 `python -c` 内联中文盘符路径。
- Python 源码里含中文或 `\01` 类数字段路径：r-string 或正斜杠（`'\01'` = 八进制转义 chr(1)）。
- 用 Python 生成 PS1 时以 BOM 编码写文件：`codecs.open(path, 'w', 'utf-8-sig')`，否则 PowerShell 解析中文出错。
- 基础路径经环境变量传递（`$env:PROJECT_BASE`），避开转义问题。

## 收尾清理

- 核验用的临时 PDF/PNG 按项目约定删除（除非用户要求 PDF 也是交付物）。
- Word 异常退出会在产物目录残留 `~$xxx.docx` 锁文件：确认 WINWORD.EXE 已结束再删。
- 交付前最后一遍：`present` 最终文档；不要把核验产物当交付物。
