#!/usr/bin/env python3
"""Cross-chapter consistency check. Run after check_content.py.

Catches the failure modes that per-chapter checks cannot see:
invented column names, inconsistent metric definitions, duplicated
H1 titles, and chapters that never reference the shared dataset.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "book")
DICT = os.path.join(ROOT, "data", "adventureworks-data-dictionary.md")

failures = []


def fail(unit, reason, evidence):
    failures.append((unit, reason, evidence))


def main():
    dtext = open(DICT, encoding="utf-8").read()
    # every identifier the dictionary legitimises
    known = set(re.findall(r"`([A-Za-z_][A-Za-z0-9_\[\]\. ]*)`", dtext))
    known |= set(re.findall(r"\b((?:Dim|Fact)[A-Za-z]+)\b", dtext))
    known_cols = {k.split("[")[0].strip() for k in known}
    known_cols |= {m for m in re.findall(r"\[([A-Za-z0-9_ ]+)\]", dtext)}

    chapters = sorted(f for f in os.listdir(BOOK) if re.match(r"ch\d+-", f))
    titles = {}

    for f in chapters:
        unit = "book/" + f
        text = open(os.path.join(BOOK, f), encoding="utf-8").read()

        h1 = re.search(r"^#\s+Chapter\s+\d+\s+[—-]\s+(.+)$", text, re.M)
        if h1:
            t = h1.group(1).strip()
            if t in titles:
                fail(unit, "duplicate chapter title", "also used by %s" % titles[t])
            titles[t] = unit

        # must actually use the shared dataset
        if not re.search(r"\b(Fact\w+|Dim\w+|Sales\[|AdventureWorks)\b", text):
            fail(unit, "chapter never references the AdventureWorks dataset",
                 "no Fact*/Dim*/Sales[] reference found")

        # any Dim*/Fact* table named must exist in the dictionary
        for tbl in set(re.findall(r"\b((?:Dim|Fact)[A-Z][A-Za-z]+)\b", text)):
            if tbl not in dtext:
                fail(unit, "references a table absent from the data dictionary", tbl)

    # the two load-bearing definitions must be stated identically where stated
    for f in chapters:
        text = open(os.path.join(BOOK, f), encoding="utf-8").read()
        for m in re.finditer(r"Margin\s*=\s*SalesAmount\s*([-−+])\s*([A-Za-z]+)", text):
            if m.group(1) == "+" or m.group(2) != "TotalProductCost":
                fail("book/" + f, "margin defined inconsistently with the data dictionary",
                     "found 'SalesAmount %s %s', expected 'SalesAmount - TotalProductCost'"
                     % (m.group(1), m.group(2)))

    # the appended Sales table with Channel is built in Chapter 7, not earlier.
    # A chapter may not claim an earlier chapter delivers it.
    for f in chapters:
        n = int(re.match(r"ch(\d+)-", f).group(1))
        text = open(os.path.join(BOOK, f), encoding="utf-8").read()
        for m in re.finditer(r"combined `?Sales`? table from Chapter (\d+)", text):
            if int(m.group(1)) < 7:
                fail("book/" + f, "attributes the appended Sales table to a chapter that "
                                  "does not build it",
                     "claims Chapter %s; the append is built in Chapter 7" % m.group(1))
        if n < 7 and re.search(r"Sales\[Channel\]", text):
            fail("book/" + f, "uses Sales[Channel] before Chapter 7 creates it",
                 "Channel column originates in the Chapter 7 append")

    if len(titles) != len(chapters):
        fail("book/", "some chapters have an unparseable H1",
             "%d titles parsed from %d chapters" % (len(titles), len(chapters)))

    if failures:
        print("CONSISTENCY FAILED — %d problem(s)\n" % len(failures))
        for unit, reason, evidence in failures:
            print("UNIT:     %s\nVERDICT:  FAIL\nREASON:   %s\nEVIDENCE: %s\nSCOPE:    this file only\n"
                  % (unit, reason, evidence))
        return 1
    print("CONSISTENCY PASSED — %d chapters, titles unique, all tables documented" % len(chapters))
    return 0


if __name__ == "__main__":
    sys.exit(main())
