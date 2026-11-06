# Chapter contract

Any chapter must satisfy `python3 tools/check_content.py`. That means:

- H1 exactly `# Chapter N — Title` (em dash), N matching the filename.
- These `##` sections, spelled exactly, in this order:
  `Learning objectives`, `The business question`, `Data you will use`,
  `Core concepts`, `Walkthrough`, `Common pitfalls`, `Exercises`,
  `Solutions`, `Key takeaways`.
- At least 2,200 words.
- At least 6 concrete Excel formulas written as `` `=FUNCTION(...)` `` or in fenced blocks.
- At least 4 numbered exercises; `## Solutions` at least 120 words.
- No TODO / TBD / placeholder text. All local links must resolve.

Style: direct, specific, no filler. Use real AdventureWorks column names from
`data/adventureworks-data-dictionary.md`. Prefer structured references
(`Sales[SalesAmount]`). Explain *why*, not only *how*. State assumptions.
