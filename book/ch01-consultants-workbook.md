# Chapter 1 — The Consultant's Workbook

A workbook is not a scratchpad. On an engagement it is a deliverable, an audit
trail and a liability all at once. Someone will open it after you have moved to
another project, change one assumption, and present the result to a board. What
you build in the next sixteen chapters has to survive that.

This chapter is about structure. It contains almost no analysis, and it is the
chapter that will save you the most time. Analysts who skip it spend the rest of
the course — and the rest of their careers — rebuilding workbooks that became
impossible to change.

## Learning objectives

By the end of this chapter you will be able to:

- Lay out an analytical workbook in distinct layers — raw, transform, analysis,
  output — and explain to a client which sheet to look at and which to ignore.
- Name and order sheets so that a stranger can navigate the file without asking.
- Apply the inputs-blue / formulas-black / cross-sheet-green convention, and say
  what it communicates.
- Choose deliberately between an Excel Table, a named range and a bare range,
  and use structured references that survive a change in row count.
- Use absolute, relative and mixed references on purpose rather than by
  trial and error.
- Maintain a Definitions sheet where each metric is defined exactly once.
- Set up the Part I workbook for the AdventureWorks case and produce a first
  sanity-checked revenue total you are willing to defend.

## The business question

AdventureWorks is a fictional multinational manufacturer and retailer of
bicycles, components, clothing and accessories. It sells through two channels:
direct to consumers on its website, and wholesale to bike shops and department
stores. Four product categories, three continental sales groups, roughly four
years of daily transactions.

The management team has asked a deceptively simple question, and it is the
question the whole book answers:

> **Where should AdventureWorks put its next dollar of growth investment?**

You cannot answer that in Chapter 1. What you can do in Chapter 1 is the thing
that makes the answer defensible later. Before anyone asks about growth, someone
will ask the entry-level question that stops engagements dead:

> **What was total revenue last year, and why does your number differ from the
> one in the sales director's deck?**

Answering that reliably requires knowing exactly which rows you summed, which
column you called revenue, and whether you included both channels. That is a
structural question, not an analytical one. Hence this chapter.

## Data you will use

For this chapter, two tables, both documented in
[../data/adventureworks-data-dictionary.md](../data/adventureworks-data-dictionary.md):

- **`FactInternetSales`** — direct-to-consumer order lines. Columns used here:
  `SalesOrderNumber`, `SalesOrderLineNumber`, `ProductKey`, `OrderDateKey`,
  `OrderQuantity`, `UnitPrice`, `SalesAmount`, `TotalProductCost`,
  `ProductStandardCost`.
- **`DimProduct`** — the product catalogue. Columns used here: `ProductKey`,
  `ProductAlternateKey`, `EnglishProductName`, `StandardCost`, `ListPrice`,
  `ProductSubcategoryKey`.

Keep the data dictionary open in a second window. The chapters refer to it
instead of repeating column lists, and that is deliberate: in real work the data
dictionary is a separate artefact that you maintain, not something you remember.

Two assumptions, stated now because every number in Part I depends on them:

1. **Revenue means `SalesAmount`.** Not `ExtendedAmount` (which ignores
   discounts), not `SalesAmount + TaxAmt + Freight` (which is what the customer
   paid, not what AdventureWorks earned).
2. **Margin means `SalesAmount − TotalProductCost`.** It is gross margin at
   standard cost. It excludes freight, selling costs and overhead. It is not
   profit, and you should never call it profit in front of a CFO.

## Core concepts

### The four layers

Every analytical workbook has four jobs, and the single most useful structural
decision is to stop one sheet from doing two of them.

**Raw.** Data exactly as it arrived. Never typed into, never sorted, never
"just fixed". If a value is wrong here, it stays wrong here and you correct it
in the transform layer where the correction is visible. A raw sheet that someone
has edited by hand is a workbook whose numbers cannot be reconciled to source,
which means they cannot be defended.

**Transform.** Cleaning, type conversion, derived columns, joins. This is where
`20130715` becomes a date and where `HouseOwnerFlag` stops being text. Every
change to the data is a formula or a Power Query step, so every change is
visible and reversible.

**Analysis.** Aggregation, pivots, statistics, models. Reads the transform
layer, never the raw layer.

**Output.** Exhibits: the charts and tables that go into the deck. Formatted for
a reader, not for an analyst. Contains as little logic as possible — ideally
none, just references to the analysis layer.

The direction of dependency is strictly one way: raw → transform → analysis →
output. The moment an output sheet feeds a transform sheet you have a circular
workbook that nobody can trace, and tracing is the whole point.

### Sheet naming and tab order

Order tabs left to right in dependency order, so reading the tab bar tells you
the data flow. Prefix with the layer so the grouping is visible even when tab
colours are lost in a PDF:

```
00_README   01_Definitions   10_RAW_Sales   11_RAW_Products
20_TRN_Sales   30_ANL_Revenue   90_OUT_Exhibit1
```

Numbered prefixes leave gaps on purpose. When you need a second transform sheet
you insert `21_TRN_Products` without renaming anything downstream. Names with
spaces force single quotes into every cross-sheet reference
(`='20 TRN Sales'!B4`), which is noise you will read a thousand times.

The `00_README` sheet holds, in plain English: what this workbook does, where
the data came from and when it was pulled, who built it, what the key
assumptions are, and what it should not be used for. Five minutes to write. It
is the difference between a file someone can pick up and a file someone
rebuilds.

### Colour as a protocol

The financial-modelling convention is simple and near-universal:

- **Blue font** — a hard-coded input. A number a human chose. Change these.
- **Black font** — a formula that depends only on cells on the same sheet.
- **Green font** — a formula that references another sheet.
- **Red font** — a link to another workbook, which is a risk and should be
  temporary.

This is not decoration. It answers, at a glance and without clicking, the
question every reviewer has: *which cells can I change, and which will break if
I touch them?* A workbook where every number is black tells the reviewer
nothing, so the reviewer clicks every cell, and reviewers who have to click
every cell stop reviewing.

The discipline has a corollary that matters more than the colours: if a cell is
blue it should be an input, and if it is an input it should be blue. A hard-coded
number sitting inside a formula cannot be coloured, which is one reason the rule
below exists.

### Tables, named ranges and bare ranges

Excel gives you three ways to refer to data, and they are not interchangeable.

**Excel Tables** (`Ctrl+T`) are the default for anything row-shaped that might
grow. A Table has a name, typed columns, headers that persist when you scroll,
and — critically — it expands automatically. A formula written against
`Sales[SalesAmount]` covers 60,000 rows today and 80,000 next quarter without
edit. `=SUM(F2:F60398)` covers exactly the rows that existed the day you typed
it, and will silently under-report forever after.

**Named ranges** are for single cells and small fixed blocks: scalar
assumptions, scenario switches, lookup constants. `=Revenue_Target` reads better
than `=Assumptions!$C$7` and does not break when someone inserts a row.

**Bare ranges** are for throwaway work only. If a bare range appears in a
formula that will ship, replace it.

Structured references read as business language rather than geometry:

```
=SUM(Sales[SalesAmount])
=SUMIFS(Sales[SalesAmount], Sales[OrderDateKey], ">=20130101", Sales[OrderDateKey], "<=20131231")
=SUM(Sales[SalesAmount]) - SUM(Sales[TotalProductCost])
```

Anyone can read the second formula and check whether it means what the exhibit
claims. Nobody can do that with `=SUMIFS(F:F,D:D,">=20130101",D:D,"<=20131231")`.

### Referencing on purpose

`$` is not a formatting flourish. It is a declaration of what should move when
the formula is copied.

- `A1` — relative. Both parts move. Use inside a block of parallel calculations.
- `$A$1` — absolute. Nothing moves. Use for a single assumption referenced
  everywhere.
- `$A1` — column locked, row free. Use when copying across a row of months that
  all read the same input column.
- `A$1` — row locked, column free. Use when copying down a column that all reads
  the same header row.

Mixed references are the ones people avoid, and they are the ones that make a
grid of calculations collapse into one formula copied over a rectangle. If you
find yourself writing a different formula in each cell of a grid, you have
probably skipped a `$`.

### One definition, one place

The most common reason two analysts produce two revenue numbers is not an error.
It is that they defined revenue differently and nobody wrote either definition
down.

The `01_Definitions` sheet carries one row per metric, with four columns: metric
name, plain-English definition, the exact Excel expression, and what it
deliberately excludes. The exclusions column is the one that prevents arguments.

| Metric | Definition | Expression | Excludes |
|---|---|---|---|
| Revenue | Net invoiced sales value of goods, after line discounts | `SUM(Sales[SalesAmount])` | Tax, freight, returns |
| COGS | Standard cost of goods sold | `SUM(Sales[TotalProductCost])` | Freight-in, overhead absorption |
| Gross margin | Revenue less COGS | Revenue − COGS | Selling cost, overhead |
| Gross margin % | Gross margin as a share of revenue | Gross margin / Revenue | — |
| Orders | Distinct customer orders | Distinct `SalesOrderNumber` | Order lines |

When the sales director's number differs from yours, this sheet turns a
credibility problem into a five-minute reconciliation.

### File naming and version discipline

`AW_Growth_Analysis_v03_2026-02-14_JG.xlsx` beats `analysis_final_FINAL(2).xlsx`.
Project, content, version, ISO date, initials. ISO dates sort chronologically in
a file listing; `14-02-2026` does not.

Never overwrite a version you have already sent to a client. The file they are
looking at must still exist on your disk, unchanged, or you cannot answer
questions about it.

### Keyboard-first

Analysis speed is mostly navigation speed. The irreducible set:

`Ctrl+Arrow` to the edge of a region · `Ctrl+Shift+Arrow` to select to it ·
`Ctrl+T` make Table · `Alt+=` AutoSum · `F4` toggle `$` while editing a
reference · `Ctrl+[` jump to a formula's precedent · `F5`+Enter jump back ·
`Alt+N+V` insert PivotTable · `Ctrl+Shift+L` toggle filters · `Ctrl+PgUp/PgDn`
move between sheets.

`F4` and `Ctrl+[` are the two that change how you work. The first makes mixed
referencing effortless; the second turns auditing from a chore into a reflex.

### Never hard-code a number inside a formula

`=Sales[@SalesAmount]*1.21` is a time bomb. In six months nobody knows whether
`1.21` is Spanish VAT, an uplift assumption, or a typo. It cannot be coloured
blue, it cannot be found by search, and it cannot be changed in one place
because there are eleven copies of it.

`=Sales[@SalesAmount]*(1+VAT_Rate)`, with `VAT_Rate` a blue cell on the
Definitions sheet, is auditable, labelled and changeable. The only numbers that
belong inside formulas are structural constants that cannot be otherwise: the
`0` in `=IFERROR(x,0)`, the `12` in an annualisation that is definitionally
twelve months.

This single rule catches more engagement-threatening errors than any other, and
it is the one most often abandoned under deadline.

## Walkthrough

Build the Part I workbook. You will extend it in Chapters 2 to 4.

**Step 1 — Create the skeleton.** New workbook, saved as
`AW_PartI_v01_2026-02-14_XX.xlsx`. Create these sheets in this order:
`00_README`, `01_Definitions`, `10_RAW_Sales`, `11_RAW_Products`,
`20_TRN_Sales`, `30_ANL_Revenue`.

Fill `00_README` with six labelled rows: Purpose, Data source, Extract date,
Author, Key assumptions, Known limitations. Under Key assumptions write the two
assumptions from the *Data you will use* section above, in words.

**Step 2 — Load the sales fact.** On `10_RAW_Sales`, use **Data → Get Data →
From Text/CSV** and import `FactInternetSales`. Load it as a Table, not a
range. Rename the Table via **Table Design → Table Name** to `Sales`.

Two things to check before going further. First, that `SalesAmount` arrived as a
number and not as text — select the column and confirm the status bar shows a
sum. Second, that `OrderDateKey` arrived as an integer like `20130715`. It will
have, and it is not a date. Chapter 2 deals with that properly; for now simply
note that you cannot do arithmetic on it.

**Step 3 — Load the product dimension.** On `11_RAW_Products`, import
`DimProduct` the same way. Name the Table `Products`.

**Step 4 — Write the Definitions sheet.** On `01_Definitions`, reproduce the
five-row metric table from *Core concepts* above. Then add a small block of
named scalar inputs, formatted in blue font:

| Cell | Name | Value |
|---|---|---|
| `C12` | `Analysis_Year` | 2013 |
| `C13` | `Margin_Target_Pct` | 0.42 |

Name them with **Formulas → Define Name**, or by typing into the Name Box. Both
are inputs; both are blue. Nothing else on this sheet is blue.

**Step 5 — First totals.** On `30_ANL_Revenue`, in green font because these
reference other sheets:

```
=SUM(Sales[SalesAmount])                                   → total internet revenue, all years
=SUM(Sales[TotalProductCost])                              → total standard COGS
=SUM(Sales[SalesAmount])-SUM(Sales[TotalProductCost])      → gross margin
=(SUM(Sales[SalesAmount])-SUM(Sales[TotalProductCost]))/SUM(Sales[SalesAmount])   → gross margin %
```

In the 2019 release, total internet `SalesAmount` comes to roughly $29 million
across the full window, at a gross margin of about 41 per cent. Your release may
differ slightly. What matters is not the figure but that you can say which rows
and which column produced it.

**Step 6 — Sanity-check before you believe it.** Three checks, each one row on
`30_ANL_Revenue`:

```
=ROWS(Sales)                                               → line count
=SUMPRODUCT(Sales[UnitPrice],Sales[OrderQuantity])         → price × quantity, independently
=SUM(Sales[ExtendedAmount])-SUM(Sales[DiscountAmount])     → a third route to the same number
```

The second and third should land close to `SUM(Sales[SalesAmount])`. Small
differences are the discount mechanics; a large difference means you have the
wrong column. Reconciling a total by two independent routes before quoting it is
the habit that separates analysts whose numbers hold up from analysts who get
corrected in meetings.

**Step 7 — Scope with the named input.** Add a year-scoped revenue figure that
reads the blue input rather than a typed year:

```
=SUMIFS(Sales[SalesAmount], Sales[OrderDateKey], ">="&Analysis_Year*10000+101, Sales[OrderDateKey], "<="&Analysis_Year*10000+1231)
```

Change `Analysis_Year` on `01_Definitions` to 2012 and the figure moves. The
`*10000+101` arithmetic is ugly — it exists only because `OrderDateKey` is an
integer. Chapter 2 removes the ugliness by making a real date column, and that
is exactly the kind of debt the transform layer is for.

**Step 8 — Colour and save.** Walk the workbook and apply the convention: blue
inputs on `01_Definitions`, green anything referencing another sheet, black for
same-sheet formulas. If you find a cell you cannot colour because it is a number
buried in a formula, that is the rule from *Core concepts* catching a real
instance. Pull it out into a named input.

## Common pitfalls

**Analysing on the raw sheet.** Sorting the raw import to "have a look" is how
the raw layer stops matching source. Look at it with a filter, which changes
nothing, or copy it to the transform layer first.

**Whole-column references.** `=SUM(F:F)` survives new rows but drags every
calculation and silently includes anything anyone pastes below the data. Tables
give you growth without the exposure.

**Tables that were never renamed.** `Table1`, `Table3`, `Table7` are worse than
range references, because they look meaningful. Rename at creation.

**Merged cells.** They break sorting, filtering, structured references and
copy-paste. Use **Center Across Selection** for the same visual effect with none
of the damage.

**Two definitions of the same metric in two places.** A margin formula on the
analysis sheet and a different one on the exhibit sheet will diverge, and the
exhibit is the one the client sees. Define once, reference everywhere.

**Formatting to hide a problem.** Rounding a figure to millions to make a
reconciliation gap disappear is not formatting, it is concealment. Fix the gap
or disclose it.

**Treating `OrderDateKey` as a date.** `=Sales[@OrderDateKey]-20130101` returns
a meaningless integer, not a day count. It will not error, which is what makes
it dangerous.

## Exercises

1. Build the skeleton workbook described in the walkthrough, including the
   `00_README` sheet with all six labelled rows filled in. Write the two Part I
   assumptions in your own words, and add one limitation that is true of the
   dataset — look at the *Known data quirks* list in the data dictionary for
   candidates.

2. Add a sixth metric to `01_Definitions`: **Average order value**. Give the
   plain-English definition, the Excel expression, and the exclusions. Note that
   it needs distinct orders rather than order lines, and that
   `=SUM(Sales[SalesAmount])/ROWS(Sales)` is therefore *not* average order
   value. Write down what that formula actually measures.

3. Replace every reference in your `30_ANL_Revenue` formulas with structured
   references, then insert three blank rows at the top of `10_RAW_Sales` above
   the Table. Record which of your formulas broke and which did not, and explain
   the pattern.

4. On `30_ANL_Revenue`, build a small grid: product `StandardCost` bands down the
   rows, calendar years across the columns, revenue in the cells. Write the top-
   left formula once with the right mix of `$` signs and copy it across the whole
   rectangle. If you had to edit any copied cell, diagnose which reference should
   have been mixed.

5. Audit this formula, which a colleague has put on an output sheet:
   `=SUM('30_ANL_Revenue'!B4:B9)*1.07-150000`. List everything wrong with it on
   the conventions in this chapter, and rewrite it so that each number is named,
   coloured and defined in one place.

6. Take the gross margin percentage you produced in Step 5 and compare it to
   `Margin_Target_Pct` on `01_Definitions`. Write one sentence, in the register
   you would use with a client, stating the gap and explicitly naming what your
   margin figure excludes.

## Solutions

**1.** A good limitation line is specific and sourced. For example: *"Reseller
sales are excluded from this workbook, so revenue here is roughly a quarter of
company turnover; the final partial month in the window is truncated and should
not be read as a decline."* Both come straight from the data dictionary's quirks
list. A limitation like *"data may contain errors"* is not a limitation, it is a
disclaimer, and it protects nobody.

**2.** `=SUM(Sales[SalesAmount])/ROWS(Sales)` is **average revenue per order
line**, which for internet sales — where `OrderQuantity` is almost always 1 — is
close to average unit price. Average order value needs distinct
`SalesOrderNumber`, and a single order of three products is one order, not
three. The honest expression at this stage is revenue divided by a distinct
count of `SalesOrderNumber`, which plain Excel does not do in one function.
Writing the definition now and the formula in Chapter 2 is the correct order of
work: define the metric, then find the mechanics. Doing it the other way round
produces whatever number the available function happens to give.

**3.** Nothing breaks. That is the point. Structured references are resolved by
Table name and column name, not by position, so rows inserted above the Table
shift it without affecting a single formula. Had you written `=SUM(F2:F60398)`,
the three inserted rows would have shifted the data down while your range stayed
put, and your total would have quietly dropped the last three rows while
including three blanks. No error, no warning, wrong number — the most dangerous
failure mode in Excel, and the reason this chapter insists on Tables before it
teaches any analysis.

**5.** Four faults. `1.07` is an unlabelled, unnamed, uncoloured assumption —
possibly tax, possibly growth, nobody can tell. `150000` is a hard-coded
adjustment with no stated basis. `B4:B9` is a positional range that breaks the
moment a row is inserted on the analysis sheet. And the formula performs
business logic on an *output* sheet, which inverts the layer dependency and
hides the calculation from anyone reviewing the analysis layer. The rewrite moves
the arithmetic to `30_ANL_Revenue`, names the two numbers as blue inputs on
`01_Definitions` with their basis documented, replaces the range with a
structured or named reference, and leaves the output sheet holding a single
green reference to the result.

**6.** Something close to: *"Gross margin on internet sales runs at roughly 41
per cent against a 42 per cent target, a gap of about one point. This is margin
at standard cost and excludes freight, selling costs and overhead, so it is not
comparable to the net margin in the management accounts."* The second sentence
is the one that matters. Quoting a margin without saying what it excludes is how
an analyst ends up defending a number they never claimed.

## Key takeaways

- Separate raw, transform, analysis and output into different sheets, and keep
  the dependency one-directional. Every number must be traceable back to source.
- Number and prefix your tabs so the tab bar documents the data flow.
- Colour is a protocol, not decoration: blue inputs, black local formulas, green
  cross-sheet. It tells a reviewer what is safe to change without clicking.
- Load row-shaped data as named Excel Tables and write structured references.
  They survive growth and row insertion; positional ranges fail silently.
- Use `$` deliberately. Mixed references are what let one formula fill a grid.
- Define each metric exactly once, in writing, with its exclusions. Most
  disagreements about numbers are undocumented disagreements about definitions.
- Never bury a number inside a formula. If a human chose it, name it, colour it
  and document its basis.
- Reconcile any total by two independent routes before you quote it.
