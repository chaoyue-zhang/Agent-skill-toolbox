# .xls（老式 Excel 二进制格式）填充

工具链：xlrd（`formatting_info=True`）+ xlutils.copy + xlwt。修复已封装在 `scripts/fill_cells_xls.py`，作为库导入使用：

```python
from fill_cells_xls import open_xls, make_style, save_xls

rb, wb = open_xls('template.xls')          # 自动应用 None 键修复
ws = wb.get_sheet(0)
ws.write(2, 1, '项目名称', make_style())    # 0 基行列；样式必须新建
save_xls(wb, 'filled.xls')                 # 先删同名旧文件再保存

# 对账回读
sh = rb.sheet_by_index(0)
print(sh.nrows, sh.cell_value(2, 1))
```

安装：`pip install xlrd xlutils xlwt -i https://mirrors.cloud.tencent.com/pypi/simple`
注意：DSH 内置 Python **不含**这三个库（内置的是 openpyxl/XlsxWriter，只支持 .xlsx），.xls 任务用宿主 Python。

## 坑 1：xlwt 在 Python 3.13 保存崩溃（_num_formats None 键）

- 症状：xlutils.copy 后 `wb.save()` 抛 `unicode(s, encoding)` 类 TypeError。
- 原因：模板样式表自带一条 `_num_formats` 条目 `k=None, v=165`，xlwt 序列化时对 None 键调用了 Python 2 式 `unicode()` 路径。
- 修复（open_xls 已内置）：

```python
wb = xlcopy(rb)
nf = wb._Workbook__styles._num_formats
if None in nf:
    del nf[None]
```

## 坑 2：绝不复用模板 XF 索引

copy 之后样式表 XF 索引已整体偏移，按原索引引用会套错样式甚至越界。写任何单元格都**新建 XFStyle**：

```python
style = xlwt.XFStyle()
b = xlwt.Borders()
b.left = b.right = b.top = b.bottom = xlwt.Borders.THIN
style.borders = b                  # 保持模板细边框观感
style.num_format_str = 'General'   # 数值格按需改，如 '0.0'
ws.write(r, c, value, style)
```

## 坑 3：不要用 COM Range.Copy 搬运中文值

生产事故记录：用 Excel COM `Range.Copy` 复制含中文的单元格时，**同列其他单元格被污染**（整行数值被覆写成 46023.0 之类的序列值）。数值一律用 xlutils/xlwt 直接 `write()`，COM 只用于只读查看或最终另存。

## 工时分配表类模板（按天小时模式）的块结构

生产模板每月一个块，共 11 行：

| 块内行 | 内容 |
|---|---|
| 1 | 月份标签行 |
| 2 | 日期行 A（1–16 号） |
| 3–6 | 数据行 A（每人一行） |
| 7 | 日期行 B（17 号–月末，含合计） |
| 8–11 | 数据行 B（每人一行） |

4 个月 = 44 行。**先用 xlrd 读回单元格值确认块边界再填，不要硬编码行号**（不同模板版本行数可能不同）。多余月份列留空即可，不要删列。

## 合计对账

"按天数值累加 ≠ 表格合计行"时的排查顺序：

1. 检查按天数值数组（dv）是否与当月实际工作日数对齐（数组长度、每月块内偏移）；
2. 正确方向是**由"每人每月总工时"反推单日值**再回填，而不是先拍单日值再凑合计；
3. 填完用 xlrd 回读合计行与数据行逐格比对，误差为 0 才算通过。

## 保存与交付

- `save_xls` 先删同名旧文件再保存，避免半成品叠加。
- 大批量填充时先写临时目录（如 `%TEMP%`），xlrd 回读验证通过后再复制到正式产物目录。
- 最终文件用 Excel COM 或 LibreOffice 打开一次做目检（列宽、边框），文本值核对以 xlrd 回读为准。
