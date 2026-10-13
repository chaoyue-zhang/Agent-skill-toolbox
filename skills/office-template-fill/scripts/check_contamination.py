# -*- coding: utf-8 -*-
"""Contamination self-check: search filled documents for previous-project blacklist strings.

Scans .doc/.docx (Word COM; body + headers/footers stories, correct Find usage -
NOT the broken Replace=2 path) and .xls/.xlsx (Excel COM, used-range values).

Usage:
    python check_contamination.py <file>... [--extra 词1,词2] [--review 词1,词2]
                                  [--skip-defaults]

    --extra          additional FORBIDDEN terms (comma separated)
    --review         additional REVIEW terms (comma separated)
    --skip-defaults  use only the terms given above

Exit codes: 0 = clean or REVIEW-only hits; 1 = FORBIDDEN hit; 2 = input/runtime error.
Every REVIEW hit must be individually explained before delivery
(see references/template-manifest.md for the default blacklist rationale).

Requires: pywin32 + MS Word/Excel (Windows).
"""
import argparse
import os
import sys

import win32com.client as win32

# Default blacklist = the PRESET SAMPLE VALUES shipped in assets/templates/
# (see references/template-manifest.md). Real previous-project identities
# (member names, old project title, real doc numbers) are NOT hard-coded here:
# pass them per project via --extra / --review, or keep a private list file
# and expand it with: --extra "$(cat my_blacklist.txt)".
#
# Chinese/mixed phrases must cover BOTH spacing variants: delivered docs often
# use the spaced form while generated text uses the unspaced one, and an
# exact-string blacklist silently misses the other.
DEFAULT_FORBIDDEN = (
    ['XXXX—2026{:03d}'.format(n) for n in range(1, 13)]   # 模板预印文号样例
)
DEFAULT_REVIEW = [
    'NO2026.001', 'NO2026.002', '技术部', 'V001', '会议验收结项',
    '电子设计企业', '2026.01.02', '2026.04.30',
]

WORD_EXT = ('.doc', '.docx')
EXCEL_EXT = ('.xls', '.xlsx')


def split_terms(raw):
    return [t.strip() for t in raw.split(',') if t.strip()] if raw else []


def iter_stories(doc):
    """Yield every story range (body, headers, footers, ...) including linked ones."""
    for story in doc.StoryRanges:
        s = story
        while s is not None:
            yield s
            try:
                s = s.NextStoryRange
            except Exception:
                s = None


def find_hits_in_story(story, term):
    """Count literal hits of term within a story range; return list of start offsets.

    Canonical VBA walk: ONE Find object on a self-advancing range - after a hit,
    Collapse(0) to the hit end and Execute again. Do NOT rebuild the range with
    SetRange(pos, end) per iteration: Word snaps the range start to table-cell
    boundaries and Find then re-returns matches lying before pos, inflating
    counts (observed: 238 hits for a term occurring 3 times).
    """
    hits = []
    rng = story.Duplicate
    f = rng.Find
    f.Text = term
    f.Forward = True
    f.Wrap = 0                 # wdFindStop
    f.MatchCase = False
    f.MatchWildcards = False
    last_start = -1
    while f.Execute():
        if rng.Start <= last_start:
            break              # non-advancing hit: safety stop
        hits.append(rng.Start)
        last_start = rng.Start
        rng.Collapse(0)        # 0 = wdCollapseEnd: continue after this hit
    return hits


def scan_word(paths, terms, word):
    """Return {path: {term: hit_count}} for Word documents."""
    results = {}
    for p in paths:
        hits_by_term = {}
        doc = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
        try:
            for term in terms:
                total = 0
                for story in iter_stories(doc):
                    total += len(find_hits_in_story(story, term))
                if total:
                    hits_by_term[term] = total
        finally:
            doc.Close(False)
        results[p] = hits_by_term
    return results


def scan_excel(paths, terms, excel):
    """Return {path: {term: hit_count}} for Excel workbooks (cell value substring)."""
    results = {}
    for p in paths:
        hits_by_term = {}
        wb = excel.Workbooks.Open(os.path.abspath(p), ReadOnly=True)
        try:
            texts = []
            for ws in wb.Worksheets:
                vals = ws.UsedRange.Value
                if vals is None:
                    continue
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                for row in vals:
                    for v in row:
                        if v is not None:
                            texts.append(str(v))
            blob = '\n'.join(texts)
            for term in terms:
                n = blob.count(term)
                if n:
                    hits_by_term[term] = n
        finally:
            wb.Close(False)
        results[p] = hits_by_term
    return results


def main():
    ap = argparse.ArgumentParser(description='Contamination self-check for filled documents.')
    ap.add_argument('files', nargs='+')
    ap.add_argument('--extra', default=None)
    ap.add_argument('--review', default=None)
    ap.add_argument('--skip-defaults', action='store_true')
    args = ap.parse_args()

    forbidden = [] if args.skip_defaults else list(DEFAULT_FORBIDDEN)
    review = [] if args.skip_defaults else list(DEFAULT_REVIEW)
    forbidden += [t for t in split_terms(args.extra) if t not in forbidden]
    review += [t for t in split_terms(args.review) if t not in review and t not in forbidden]
    all_terms = forbidden + review

    missing = [f for f in args.files if not os.path.isfile(f)]
    if missing:
        for m in missing:
            print('MISSING:', m)
        sys.exit(2)

    word_files = [f for f in args.files if f.lower().endswith(WORD_EXT) and '~$' not in os.path.basename(f)]
    excel_files = [f for f in args.files if f.lower().endswith(EXCEL_EXT) and '~$' not in os.path.basename(f)]
    if not word_files and not excel_files:
        print('no scannable Office files in input')
        sys.exit(2)

    results = {}
    if word_files:
        word = win32.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            results.update(scan_word(word_files, all_terms, word))
        finally:
            word.Quit()
    if excel_files:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            results.update(scan_excel(excel_files, all_terms, excel))
        finally:
            excel.Quit()

    forbidden_total = 0
    for path, hits_by_term in results.items():
        print('=' * 60)
        print('FILE:', path)
        if not hits_by_term:
            print('  clean (no blacklist hits)')
        for term, n in sorted(hits_by_term.items()):
            level = 'FORBIDDEN' if term in forbidden else 'REVIEW   '
            print('  [{}] {!r}: {} hit(s)'.format(level, term, n))
            if term in forbidden:
                forbidden_total += n

    print('=' * 60)
    if forbidden_total:
        print('RESULT: FAIL - {} FORBIDDEN hit(s); fix the documents before delivery'.format(forbidden_total))
        sys.exit(1)
    print('RESULT: PASS (REVIEW hits, if any, need individual justification)')
    sys.exit(0)


if __name__ == '__main__':
    main()
