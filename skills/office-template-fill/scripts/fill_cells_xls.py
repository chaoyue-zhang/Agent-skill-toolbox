# -*- coding: utf-8 -*-
"""Fill legacy .xls templates with xlrd + xlutils.copy + xlwt, preserving layout.

Library use (preferred):
    from fill_cells_xls import open_xls, make_style, save_xls

    rb, wb = open_xls('template.xls')            # applies the None-key fix below
    ws = wb.get_sheet(0)
    ws.write(2, 1, '项目名称', make_style())      # 0-based row/col, fresh style
    save_xls(wb, 'filled.xls')                   # removes an existing target first

    # read back for reconciliation (see references/xls-fill.md):
    sh = rb.sheet_by_index(0)
    print(sh.nrows, sh.cell_value(2, 1))

CLI (quick one-off writes):
    python fill_cells_xls.py <template.xls> <out.xls> "<sheet>,<row>,<col>,<value>" ...
    sheet = sheet name or 0-based index; row/col 0-based;
    value stays a string unless it parses as int/float.

Production pitfalls baked in here:
1. xlwt save crashes on Python 3.13 when the template carries a _num_formats
   entry with key None (k=None, v=165) -> unicode(s, encoding) TypeError.
   open_xls() deletes the None key right after xlutils.copy.
2. Template XF indexes shift during copy. NEVER write with an original XF
   index; always pass a freshly built XFStyle (thin borders + num_format_str).
3. Do not move values through Excel COM Range.Copy: copying Chinese text
   polluted other cells of the same column (whole runs overwritten with values
   like 46023.0). Plain xlwt writes are the reliable path.
4. xlrd must open with formatting_info=True or xlutils.copy loses the layout.

Install: pip install xlrd xlutils xlwt -i https://mirrors.cloud.tencent.com/pypi/simple
(The DSH bundled Python does NOT ship these; use a host Python.)
"""
import os
import sys

import xlrd
import xlwt
from xlutils.copy import copy as xlcopy


def open_xls(path):
    """Return (read_book, write_book); write_book has the None num-format key removed."""
    rb = xlrd.open_workbook(path, formatting_info=True)
    wb = xlcopy(rb)
    num_formats = wb._Workbook__styles._num_formats
    if None in num_formats:
        del num_formats[None]
    return rb, wb


def make_style(num_format_str='General', bold=False):
    """Fresh XFStyle per written cell - never reuse template XF indexes."""
    style = xlwt.XFStyle()
    borders = xlwt.Borders()
    borders.left = borders.right = borders.top = borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    style.num_format_str = num_format_str
    if bold:
        font = xlwt.Font()
        font.bold = True
        style.font = font
    return style


def save_xls(wb, path):
    if os.path.exists(path):
        os.remove(path)
    wb.save(path)


def coerce(value):
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    rb, wb = open_xls(src)
    names = [s.name for s in rb.sheets()]
    for spec in sys.argv[3:]:
        sheet, row, col, value = spec.split(',', 3)
        idx = names.index(sheet) if sheet in names else int(sheet)
        ws = wb.get_sheet(idx)
        ws.write(int(row), int(col), coerce(value), make_style())
        print('wrote', spec[:60])
    save_xls(wb, dst)
    print('saved', dst)


if __name__ == '__main__':
    main()
