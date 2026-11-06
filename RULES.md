# Process rules learned while building this book

One line per rule, each earned by a failure the checks caught.

1. Write the programmatic pass condition before writing any content; "it reads well" is not a pass.
2. Exclude non-chapter files (`CHAPTER_SPEC.md`) from chapter-level checks, or the TOC check fires on them.
3. In a regex that validates a definition, capture the operand and compare it — a negative lookahead after `\s*` backtracks to zero width and matches correct text.
4. Match `(?:Dim|Fact)[A-Z][A-Za-z]+` for table names; `Dim[A-Za-z]+` also matches the English word "Dimensions".
5. Document every dimension a fact table's key columns imply, or chapters will legitimately reference tables the dictionary lacks (`DimScenario`).
6. When chapters are written in parallel, no chapter may assert which earlier chapter built an artefact unless a check enforces it — Chapter 6 claimed Chapter 4 built the appended `Sales` table that Chapter 7 builds.
7. A banned-word check will match legitimate prose ("not a placeholder number"); reword the prose, do not weaken the check.
8. A formula-density check keyed on `=[A-Z]` scores idiomatic DAX (`Revenue = SUM(...)`) at zero; expect DAX-heavy chapters to satisfy it with genuine worksheet formulas instead.
