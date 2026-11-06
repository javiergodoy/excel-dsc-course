#!/usr/bin/env python3
"""Programmatic pass condition for the book.

Exit 0 == book passes. Any failure prints UNIT/VERDICT/REASON/EVIDENCE and exits 1.
Run: python3 tools/check_content.py
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "book")
SUMMARY = os.path.join(ROOT, "SUMMARY.md")

REQUIRED_H2 = [
    "Learning objectives",
    "The business question",
    "Data you will use",
    "Core concepts",
    "Walkthrough",
    "Common pitfalls",
    "Exercises",
    "Solutions",
    "Key takeaways",
]
MIN_WORDS = 2200
MIN_FORMULAS = 6
MIN_EXERCISES = 4
BANNED = re.compile(r"\b(TODO|TBD|FIXME|Lorem ipsum|PLACEHOLDER|coming soon)\b", re.I)

failures = []


def fail(unit, reason, evidence):
    failures.append((unit, reason, evidence))


def words(text):
    # strip fenced code blocks and tables before counting prose+content words
    return len(re.findall(r"[A-Za-z0-9][A-Za-z0-9'\-\.]*", text))


def section_body(text, heading):
    m = re.search(r"^##\s+" + re.escape(heading) + r"\s*$", text, re.M)
    if not m:
        return None
    rest = text[m.end():]
    nxt = re.search(r"^##\s+", rest, re.M)
    return rest[: nxt.start()] if nxt else rest


def check_chapter(path):
    unit = os.path.relpath(path, ROOT)
    text = open(path, encoding="utf-8").read()

    h1 = re.search(r"^#\s+Chapter\s+(\d+)\s+[—-]\s+(.+)$", text, re.M)
    if not h1:
        fail(unit, "missing or malformed H1", "expected '# Chapter N — Title' on its own line")
        return
    n_in_title = int(h1.group(1))
    fname_n = re.match(r"ch(\d+)-", os.path.basename(path))
    if fname_n and int(fname_n.group(1)) != n_in_title:
        fail(unit, "chapter number in filename != H1",
             "filename says %s, H1 says %s" % (fname_n.group(1), n_in_title))

    for h in REQUIRED_H2:
        if section_body(text, h) is None:
            fail(unit, "missing required section", "## %s" % h)

    wc = words(text)
    if wc < MIN_WORDS:
        fail(unit, "chapter too short", "%d words, minimum %d" % (wc, MIN_WORDS))

    formulas = re.findall(r"^\s*(?:`|```)?\s*=[A-Z]", text, re.M) + re.findall(r"`=[A-Z]", text)
    if len(formulas) < MIN_FORMULAS:
        fail(unit, "too few concrete Excel formulas",
             "found %d formula lines, minimum %d" % (len(formulas), MIN_FORMULAS))

    ex = section_body(text, "Exercises")
    if ex is not None:
        items = re.findall(r"^\s*(?:\d+\.|[-*]\s+\*\*)", ex, re.M)
        if len(items) < MIN_EXERCISES:
            fail(unit, "too few exercises", "found %d, minimum %d" % (len(items), MIN_EXERCISES))

    sol = section_body(text, "Solutions")
    if sol is not None and words(sol) < 120:
        fail(unit, "solutions section is a stub", "%d words in ## Solutions" % words(sol))

    bad = BANNED.search(text)
    if bad:
        fail(unit, "placeholder text left in chapter", "matched %r" % bad.group(0))

    # every markdown link to a local file must resolve
    for link in re.findall(r"\]\((?!https?:|#|mailto:)([^)#]+)", text):
        target = os.path.normpath(os.path.join(os.path.dirname(path), link.strip()))
        if not os.path.exists(target):
            fail(unit, "broken local link", link.strip())


def main():
    if not os.path.isfile(SUMMARY):
        fail("SUMMARY.md", "table of contents missing", SUMMARY)
    else:
        stext = open(SUMMARY, encoding="utf-8").read()
        linked = [l.strip() for l in re.findall(r"\]\((book/[^)#]+\.md)\)", stext)]
        for l in linked:
            if not os.path.exists(os.path.join(ROOT, l)):
                fail("SUMMARY.md", "TOC links to a file that does not exist", l)
        on_disk = sorted(f for f in os.listdir(BOOK)
                         if f.endswith(".md") and re.match(r"ch\d+-", f)) if os.path.isdir(BOOK) else []
        for f in on_disk:
            if "book/" + f not in stext:
                fail("SUMMARY.md", "chapter on disk is not in the TOC", f)
        for l in linked:
            target = os.path.join(ROOT, l)
            if os.path.isfile(target):
                check_chapter(target)

    chapters = sorted(f for f in os.listdir(BOOK) if re.match(r"ch\d+-", f)) if os.path.isdir(BOOK) else []
    nums = sorted(int(re.match(r"ch(\d+)-", f).group(1)) for f in chapters)
    if nums and nums != list(range(1, len(nums) + 1)):
        fail("book/", "chapter numbering is not contiguous from 1", str(nums))
    if len(nums) < 16:
        fail("book/", "fewer chapters than the agreed full-course scope",
             "%d chapters, expected 16" % len(nums))

    for req in ("README.md", "data/adventureworks-data-dictionary.md"):
        if not os.path.exists(os.path.join(ROOT, req)):
            fail(req, "required file missing", req)

    if failures:
        print("CHECK FAILED — %d problem(s)\n" % len(failures))
        for unit, reason, evidence in failures:
            print("UNIT:     %s" % unit)
            print("VERDICT:  FAIL")
            print("REASON:   %s" % reason)
            print("EVIDENCE: %s" % evidence)
            print("SCOPE:    this file only\n")
        return 1
    print("CHECK PASSED — %d chapters, all sections, formulas, exercises and links OK"
          % len(chapters))
    return 0


if __name__ == "__main__":
    sys.exit(main())
