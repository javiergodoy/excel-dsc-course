# Excel for Data Science and Consulting

### A course book built on the AdventureWorks case

This is the book for the IE course *Excel for Data Science and Consulting*. It
teaches Excel the way a consultant actually uses it: as an analytical
instrument that has to produce a defensible answer, on a deadline, for someone
who will act on it.

Every chapter works on the same dataset — **AdventureWorks**, a fictional
multinational bicycle manufacturer — so the analysis compounds. By Chapter 15
you are not practising formulas. You are presenting a recommendation to a
management team about where AdventureWorks should put its next dollar of
growth investment, and defending it with exhibits you built yourself.

## Who this is for

Students and practitioners who can already use Excel for everyday work and
want to use it for analysis. You should be comfortable typing a formula, naming
a sheet and making a chart. You do not need statistics, SQL, VBA or Python.

## How the book is organised

**Part I — Foundations (Chapters 1–4).** Workbook architecture, the shape of
the AdventureWorks data model, data hygiene, and joining tables in Excel.
This is the part most courses skip and most analysts' errors come from.

**Part II — Analysis (Chapters 5–8).** Conditional aggregation and dynamic
arrays, PivotTables used properly, Power Query for repeatable ETL, and
Power Pivot with DAX measures.

**Part III — Quantitative methods (Chapters 9–12).** Descriptive statistics,
time series and forecasting, regression and driver analysis, and customer
segmentation.

**Part IV — Consulting delivery (Chapters 13–16).** Financial and scenario
modelling, dashboard design, turning analysis into a recommendation, and
workbook quality, audit and handoff.

Read [SUMMARY.md](SUMMARY.md) for the full table of contents.

## Before you start

1. Get the data. Follow
   [data/adventureworks-data-dictionary.md](data/adventureworks-data-dictionary.md),
   which explains which AdventureWorks release to download, how to load it, and
   what every table and column means. Keep that file open while you work; the
   chapters refer to it constantly rather than repeating column lists.
2. Check your Excel. The book assumes **Microsoft 365 Excel for Windows**.
   Chapters 5 onwards use dynamic arrays (`FILTER`, `UNIQUE`, `SORT`, `LET`)
   that need a current build. Chapters 7 and 8 use Power Query and Power Pivot;
   Power Pivot is Windows-only, and Chapter 8 notes the Mac workaround.
3. Create one workbook per part, not one per exercise. Chapter 1 explains the
   layout and why it matters more than it sounds like it does.

## How to use each chapter

Chapters share a fixed shape so you can navigate them predictably:

- **Learning objectives** — what you will be able to do.
- **The business question** — the consulting question driving the chapter.
  Techniques are introduced because a question needs them, never for their own sake.
- **Data you will use** — exact tables and columns.
- **Core concepts** — the ideas, with the reasoning behind them.
- **Walkthrough** — build it, step by step, with real formulas.
- **Common pitfalls** — the specific ways this goes wrong in practice.
- **Exercises** — work to do yourself.
- **Solutions** — worked answers, including the interpretation, not just the formula.
- **Key takeaways** — what to remember when the deadline is tomorrow.

Do the exercises. The walkthroughs are readable; the exercises are where the
skill forms.

## Checking the book

The repository includes a content check that verifies structure, required
sections, formula density, exercise counts and internal links:

```bash
python3 tools/check_content.py
```

## A note on what this book argues

Excel's reputation as a serious analytical tool is mixed, and some of that
reputation is earned — by workbooks nobody can audit, models whose assumptions
are buried in cell `K47`, and charts that mislead by accident. The argument of
this book is that those are failures of discipline rather than of the tool, and
that the discipline is learnable: separate inputs from logic, define every
metric once, make every number traceable to its source, and never present a
number you could not reconstruct in front of the client.

That discipline is also what transfers. The analyst who has internalised it
moves to Python, SQL or a BI tool and carries the judgment across. The one who
has only memorised functions starts again.
