# -*- coding: utf-8 -*-
"""Dump the block structure of a .docx: paragraphs and tables in document order.

Usage:
    python dump_docx_structure.py <file.docx> [--full]
        --full   print entire text instead of a 60-char preview

Uses python-docx only (no Word COM needed) - the DSH bundled Python works:
call load_workspace_dependencies to get its path.

Why the "starting cell index" matters: templates are full of merged cells.
python-docx repeats the same underlying _tc across row.cells for merges; this
dump dedupes by id(_tc) and prints, per row, the starting cell index of each
distinct cell. That logical numbering matches Word COM Table.Cell(r, c), where
c counts merged cells within the row, NOT physical grid columns. Use these
coordinates when planning fills (see references/word-com-fill.md).
"""
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc):
    """Yield Paragraph and Table objects in document body order."""
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)


def first_cells(row):
    """[(starting_index, cell)] deduped by underlying _tc identity."""
    seen, out = set(), []
    for i, c in enumerate(row.cells):
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append((i, c))
    return out


def preview(text, full):
    if full:
        return text.replace('\n', '\\n')
    text = ' '.join(text.split())
    if len(text) > 60:
        return text[:60] + '...'
    return text


def dump_table(table, full, indent='  '):
    print('{}rows={} grid_cols={}'.format(indent, len(table.rows), len(table.columns)))
    for r, row in enumerate(table.rows):
        cells = first_cells(row)
        parts = ['c{}={}'.format(ci, preview(c.text, full)) for ci, c in cells]
        print('{}row {:>2} ({} cells): {}'.format(indent, r, len(cells), ' | '.join(parts)))


def main():
    full = '--full' in sys.argv[1:]
    args = [a for a in sys.argv[1:] if a != '--full']
    if len(args) != 1:
        print(__doc__)
        sys.exit(1)
    doc = Document(args[0])
    pidx = 0
    for bi, block in enumerate(iter_blocks(doc)):
        if isinstance(block, Paragraph):
            print('[{:>3}] P{:<3} style={:<20} {}'.format(
                bi, pidx, str(block.style.name), preview(block.text, full)))
            pidx += 1
        else:
            print('[{:>3}] TABLE'.format(bi))
            dump_table(block, full)


if __name__ == '__main__':
    main()
