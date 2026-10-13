# python-docx 单元格重写（.docx）

适用：整格重写 .docx 表格单元格为生成内容（评审意见节、试用意见节等），无需 Word COM。
运行环境：DSH 内置 Python 自带 python-docx（用 `load_workspace_dependencies` 取解释器路径），不装任何包。
仅支持 .docx；老式 .doc 先经 `scripts/convert_doc_to_docx.py` 转换。

## 合并单元格去重（定位坐标系）

模板常含合并单元格；遍历 `row.cells` 时同一个底层 `_tc` 对象会重复出现多次。用 `id(_tc)` 去重得到每行"起始单元格索引 + cell 对象"：

```python
def get_first_cells(row):
    seen, out = set(), []
    for i, c in enumerate(row.cells):
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append((i, c))
    return out
```

- 该逻辑坐标与 Word COM `Table.Cell(r, c)` 的"行内第 N 个合并单元格"一致，两条路线可共用探测结果（`scripts/dump_docx_structure.py` 已内置去重输出）。
- **id() 去重不可跨行/跨文档持久化**：lxml 元素代理对象被 GC 后 id 会被复用，只存 id 整数而不持有元素引用的"全文档 seen 集合"会把无辜单元格误判为重复而**静默丢内容**（生产验证：整篇文档标题格被丢，检索结果全错）。去重范围限定在单行内，且被保留的 cell 对象要持引用（放入返回列表）。
- **整行合并的格**在该行只有 1 个起始单元格：内容写 `cells[0]`（起始格），不要写 `cells[1]`——那是合并重复位，写进去内容会消失或错位。

## 整格重写段落（可控加粗）

先删光单元格内所有 `<w:p>`，再逐段重建（生产验证，原 doc12/doc13 填充路线）：

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def rewrite_cell(cell, blocks):
    """blocks: [(text, is_bold), ...]，每项重建为一个段落。"""
    tc = cell._tc
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    for text, is_bold in blocks:
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        if is_bold:
            rpr = OxmlElement('w:rPr')
            rpr.append(OxmlElement('w:b'))
            r.append(rpr)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')   # 保留首尾空格
        t.text = text
        r.append(t)
        p.append(r)
        tc.append(p)
```

要点：

- 单元格必须至少保留一个 `<w:p>`，全删不重建会产出损坏文件。
- 需要沿用原段落格式（居中、行距）时：**先 `deepcopy` 原 `_element` 再清空其 runs** 作模板用；直接改原元素会破坏原内容。
- **变量名错位是高频坑**：`(title, content)` 与 `(title, is_bold)` 两种元组语义不要混用——混用后内容段被静默吞掉。动手前先定死 blocks 的元组结构。
- 中文字体：仅设 `run.font.name` 不保证中文字形，必要时补 `w:rPr/w:rFonts` 的 `w:eastAsia` 属性。

## 校验

1. 重写后用 python-docx 重新打开，回读单元格文本逐段比对；
2. 再按 `preview-verify.md` 转 PDF 渲染页图目检（合并格溢出、字体丢失只有渲染图能发现）。
