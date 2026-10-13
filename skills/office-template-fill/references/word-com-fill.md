# Word COM 保格式填充（.doc/.docx）

前提：Windows + 本机 MS Word + pywin32。
安装：`pip install pywin32 -i https://mirrors.cloud.tencent.com/pypi/simple`
适用：格式敏感的模板填充（边框、字体、编号必须原样保留），以及 python-docx 无法直接解析的老式 .doc。

## 会话骨架

```python
import win32com.client as win32

word = win32.DispatchEx('Word.Application')  # 独立实例，不抢占用户已打开的 Word
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(r'D:/产物/01 项目建议书.docx')
    # ... 填充操作 ...
    doc.Save()
    doc.Close(False)
finally:
    word.Quit()   # 必须，否则残留 WINWORD.EXE 并产生 ~$ 锁文件
```

- 用 `DispatchEx` 而非 `Dispatch`：避免复用用户已打开的 Word 实例（Visible/弹窗设置互相干扰）。
- 路径含中文、空格或 `\01` 类数字段：用 r-string 或正斜杠（`'\01'` 会被解释成八进制转义）。
- 不要在 `python -c` 内联中文路径 —— 落成脚本文件执行；基础路径可用环境变量（如 `$env:PROJECT_BASE`）传入。

## 模式 A：空值格填充 / 大格末尾续填

单元格的格式挂在其段落标记上。核心做法：定位单元格**最后一个段落**，去掉末尾单元格标记 `\x07`，collapse 到段首后写 `Range.Text` —— 以 `\r` 分段展开的新段落会继承该段的字体与段落格式（生产验证，见原 fill_01.py 路线）：

```python
cell = doc.Tables(1).Cell(2, 3)
paras = cell.Range.Paragraphs
rng = paras(paras.Count).Range.Duplicate   # 最后一个段落
rng.End = rng.End - 1                      # 去掉单元格标记 \x07
rng.Collapse(1)                            # 1 = wdCollapseStart，collapse 到段首
rng.Text = '第一行内容\r第二行内容'          # \r 分段
```

- 大格"续填"（已有内容后追加）同理：最后一个段落位于既有内容之后，在其段首写入即自然接排。
- 若格内有占位符（如全角空格），可先对占位符 Range 置空再写，或直接用模式 B。

## 模式 B：整格替换

```python
rng = cell.Range.Duplicate
rng.End = rng.End - 1        # 去掉 \x07，保留单元格自身段落标记
rng.Text = '新的整格内容\r第二行'
```

替换范围覆盖格内全部旧文本；单元格级格式（边框、对齐）不受影响。

## 模式 C：文档级文本替换 —— Replace=2 静默失效！

**坑（生产验证）**：pywin32 下 `Find.Execute(Replace=2)` 与命名 `ReplaceWith` 绑定**返回 True 但实际不生效**。
正确做法：Find 定位命中后，对命中 Range 直接改写文本：

```python
rng = doc.Content.Duplicate
find = rng.Find
find.Text = '旧文本'
find.Forward = True
find.Wrap = 0                    # 0 = wdFindStop
find.MatchCase = False
if find.Execute():               # 命中后 rng 已收缩为命中范围
    rng.Duplicate.Text = '新文本'
```

- 循环替换时每次重新从 `doc.Content.Duplicate` 开始查找。
- 替换完必须独立校验（再查一遍旧文本应为 0 命中），不要相信 Execute 返回值。

**遍历统计多处命中时（生产验证的坑）**：不要每轮用 `SetRange(pos, end)` 重建查找范围再 Execute——范围起点落在表格单元格内时会被 Word **吸附到单元格边界**，Find 会向后重复命中已统计过的目标（实测一个出现 3 次的词被计成 238 次，pos 在格内每前进 1 字符重复计一次）。
正确姿势是规范 VBA 走法：同一个 Find 对象，命中后 `rng.Collapse(0)`（wdCollapseEnd）折叠到命中末尾再 `Execute()`，直到返回 False；另加"命中位置不前进即 break"的保险。参考 `scripts/check_contamination.py` 的 `find_hits_in_story`。

另：`doc.StoryRanges` 的 `Count` 不可信、按索引访问会抛"所需要的成员不存在"（索引是 story 类型号而非序号）；**只能用 For Each 枚举**。正文=StoryType 1，正文外内容（文本框/页眉页脚）在各自的 story 里，全文检索需遍历所有 story 及其 `NextStoryRange` 链。

## 自动编号陷阱（填大格前必做）

模板大单元格里的标签段（如"产品主要功能："）常由 Word **自动编号**提供"（一）（二）（三）"前缀，带编号的标签段与无编号的待填空段交替排列。

**坑**：向空段直接写多行 `Range.Text` 时，展开出的新段落会**继承编号格式**，渲染成"（十）（1）…"双重编号错乱。

处理步骤：

1. **填前探测**：对原 .doc/.docx 跑 `scripts/probe_list_numbering.py`，打印每段 `Range.ListFormat.ListString`，记录编号归属的段落（标签段）及其索引。
2. **填后正文去编号**，仅保留标签段的编号：

```python
for i in range(first, last + 1):      # 覆盖填充区域的段落范围
    p = doc.Paragraphs(i)
    if p.Range.ListFormat.ListString and i not in label_idx:
        p.Range.ListFormat.RemoveNumbers(1)
```

3. **复测**：再跑一遍探测脚本，确认编号序列只剩标签段持有、正文段 ListString 为空，然后才进入 PDF 核验。

## 合并单元格坐标

**坑**：`Table.Cell(r, c)` 的 c 是**该行内第 N 个（合并后的）单元格**，不是物理列号。

填前探测行内实际单元格：

```python
row = doc.Tables(1).Rows(3)
for c in range(1, row.Cells.Count + 1):
    print(3, c, repr(row.Cells(c).Range.Text[:20]))
```

或用 `scripts/dump_docx_structure.py`（python-docx，按 `id(_tc)` 去重）直接得到每行的"起始单元格索引"逻辑坐标，无需开 Word。
整行合并的格只有 1 个起始单元格，内容写起始格，不要按固定列号猜。

## 段落索引漂移

向单元格写入多行文本后，文档段落总数变化，后续段落索引整体**漂移**。

规则：同一单元格内多处写入时，**从后往前写**（先写索引大的段落），先写的内容不会移动后写的坐标；或每次写入前重新定位目标 Range，不缓存旧索引。

## 收尾

- 原地保存 `doc.Save()`；另存 .docx 用 `doc.SaveAs2(new_path, FileFormat=16)`（目标已存在时先删，防静默失败）。
- Word 异常退出后：清理残留 `WINWORD.EXE` 进程与产物目录下的 `~$*.docx` 锁文件（须先确认用户的 Word 已关闭）。
- 文本回读通过≠格式正确：编号错乱、合并格溢出、字体丢失只在渲染图上暴露，务必按 `preview-verify.md` 做视觉核验。
