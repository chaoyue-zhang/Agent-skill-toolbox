# -*- coding: utf-8 -*-
"""Probe Word auto-numbering (ListFormat.ListString) per paragraph of a .doc/.docx.

Run this BEFORE filling any template whose cells show （一）（二）/ 1. /（1）-style
prefixes: those prefixes are often Word AUTO-NUMBERING owned by label paragraphs
(e.g. "产品主要功能："), alternating with empty paragraphs meant for content.
Writing multi-line text into an empty paragraph makes the expanded paragraphs
INHERIT the numbering and render as double-number garbage like （十）（1）…
The printout shows which paragraphs own numbering, so you can keep it on labels
and call ListFormat.RemoveNumbers(1) on content paragraphs after filling
(see references/word-com-fill.md).

Usage:
    python probe_list_numbering.py <file.doc|file.docx> [--all] [--contains TEXT]

    default        only paragraphs that carry a ListString (numbered ones)
    --all          every paragraph (numbering column shows '')
    --contains T   only paragraphs whose text contains T (numbered or not)

Requires: pywin32 + Microsoft Word (Windows). Works on legacy .doc directly.
Output paragraph indexes are 1-based doc.Paragraphs indexes and cover table
cells too - the same coordinate system used when removing numbers.
"""
import argparse
import os
import sys

import win32com.client as win32


def main():
    ap = argparse.ArgumentParser(description='Probe Word paragraph auto-numbering.')
    ap.add_argument('path', help='target .doc or .docx file')
    ap.add_argument('--all', action='store_true', help='print every paragraph')
    ap.add_argument('--contains', default=None, help='only paragraphs whose text contains this')
    args = ap.parse_args()

    path = os.path.abspath(args.path)
    if not os.path.isfile(path):
        print('file not found:', path)
        sys.exit(1)

    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    shown = numbered = 0
    try:
        doc = word.Documents.Open(path, ReadOnly=True)
        try:
            total = doc.Paragraphs.Count
            print('paragraphs:', total)
            for i in range(1, total + 1):
                para = doc.Paragraphs(i)
                text = para.Range.Text.replace('\r', ' ').replace('\x07', '').strip()
                ls = para.Range.ListFormat.ListString or ''
                if ls:
                    numbered += 1
                if args.contains is not None:
                    keep = args.contains in text
                elif args.all:
                    keep = True
                else:
                    keep = bool(ls)
                if not keep:
                    continue
                if len(text) > 44:
                    text = text[:44] + '...'
                print('{:>5} [{}] list={!r:<10} {}'.format(
                    i, 'NUM' if ls else '   ', ls, text))
                shown += 1
        finally:
            doc.Close(False)
    finally:
        word.Quit()
    print('shown {} paragraph(s); {} numbered in document'.format(shown, numbered))


if __name__ == '__main__':
    main()
