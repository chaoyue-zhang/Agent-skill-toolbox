# -*- coding: utf-8 -*-
"""Batch-convert legacy .doc files to .docx via Word COM (Windows + MS Word).

Usage:
    python convert_doc_to_docx.py <src> [<dst_dir>]

    <src>      one .doc file, or a directory (all .doc files inside, non-recursive)
    <dst_dir>  output directory; default = next to each source file

Dependency:
    pip install pywin32 -i https://mirrors.cloud.tencent.com/pypi/simple

Production notes:
- SaveAs2(FileFormat=16) = wdFormatXMLDocument (.docx).
- An existing target .docx is deleted first: SaveAs2 onto an existing file may
  silently keep the old content.
- Word lock files (~$*.doc) are skipped.
- The Word instance is created with DispatchEx (independent of any user session)
  and always Quit() in finally to avoid orphaned WINWORD.EXE processes.
"""
import os
import sys

import win32com.client as win32


def collect(src):
    src = os.path.abspath(src)
    if os.path.isfile(src):
        return [src]
    out = []
    for name in sorted(os.listdir(src)):
        p = os.path.join(src, name)
        if os.path.isfile(p) and name.lower().endswith('.doc') and not name.startswith('~$'):
            out.append(p)
    return out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    dst_dir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else None
    docs = collect(src)
    if not docs:
        print('no .doc files found under:', src)
        sys.exit(1)
    if dst_dir:
        os.makedirs(dst_dir, exist_ok=True)

    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    ok = fail = 0
    try:
        for d in docs:
            base = os.path.splitext(os.path.basename(d))[0]
            out = os.path.join(dst_dir or os.path.dirname(d), base + '.docx')
            if os.path.exists(out):
                os.remove(out)
            try:
                doc = word.Documents.Open(d, ReadOnly=False)
                doc.SaveAs2(out, FileFormat=16)
                doc.Close(False)
                ok += 1
                print('OK  ', d, '->', out)
            except Exception as exc:
                fail += 1
                print('FAIL', d, exc)
    finally:
        word.Quit()
    print('done: {} converted, {} failed'.format(ok, fail))
    sys.exit(0 if fail == 0 else 2)


if __name__ == '__main__':
    main()
