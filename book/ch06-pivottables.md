# Chapter 6 — PivotTables as the Analyst's Workhorse

The PivotTable is the most productive feature Excel has ever shipped and the
one most likely to put a wrong number on a slide.

Both halves of that sentence matter. A pivot answers a question in fifteen
seconds that a formula block takes twenty minutes to answer — which, during the
exploratory phase of an engagement, is the difference between understanding the
business on day two and understanding it on day five. But a pivot also hides
its own arithmetic. Drag a field into Values and Excel picks an aggregation for
you. Drop a ratio into a calculated field and Excel computes it in a way that
is almost certainly not what you meant. Leave a filter off and a junk dimension
row walks onto your exhibit.

This chapter teaches pivots as a grammar you control, and teaches the specific
places where the default behaviour is wrong.

## Learning objectives

By the end of this chapter you will be able to:

- Explain what the pivot cache is, and why a refreshed pivot and a stale pivot
  can disagree without any visible sign.
- Use rows, columns, values and filters deliberately, as four distinct roles.
- Choose value field settings correctly, including `% of row`, `% of column`,
  `% of parent total`, running totals and difference from previous period.
- Group dates and numeric fields into bins, and fix the sort orders Excel gets
  wrong — including the `CommuteDistance` trap.
- Handle the `NA` sales territory row explicitly instead of letting it appear.
- Explain why a calculated field on a ratio is usually wrong, and compute
  margin percentage correctly instead.
- Use `GETPIVOTDATA` to reference a pivot without depending on cell positions.
- Wire one slicer or timeline to several pivots, and use Show Report Filter
  Pages to generate a per-territory pack in one action.
- Recognise when a pivot is the wrong tool and a formula block is the right one.

## The business question

> AdventureWorks sells through two channels with very different economics. The
> management team believes bikes are the profitable part of the business,
> because bikes are where the revenue is. Before the growth model is built,
> establish the truth: revenue *and* gross margin percentage by channel and
> category, and the same view by sales territory group, with any unassigned
> revenue disclosed rather than buried.

This is the question the engagement turns on, and it is a trap. Revenue by
category says one thing; margin percentage by category frequently says the
opposite, because Bikes carry high unit prices and thin percentage margins
while Accessories carry low prices and fat ones. An analyst who reports only
revenue confirms the client's existing belief and adds nothing. The whole value
of the exhibit is in the second column.

Getting that second column right is where the calculated-field trap bites.

## Data you will use

The combined `Sales` table from Chapter 4, or the two fact tables separately:

- `SalesAmount` — revenue.
- `TotalProductCost` — cost of goods.
- `OrderQuantity` — units.
- `SalesOrderNumber` — for distinct order counts.
- `Channel` — `"Internet"` or `"Reseller"`.
- `Category` and `Subcategory` — from the product hierarchy.
- `CalendarYear`, `EnglishMonthName`, `MonthNumberOfYear` — from `DimDate`.
- `SalesTerritoryGroup`, `SalesTerritoryRegion` — from `DimSalesTerritory`.

And, for the sorting exercise, `DimCustomer[CommuteDistance]` joined through
`CustomerKey`. Definitions and quirks are in
[the data dictionary](../data/adventureworks-data-dictionary.md).

One critical preparation step. Add a calculated column to the `Sales` table
before you build anything:

`=[@SalesAmount]-[@TotalProductCost]`

Name it `Margin`. You need it as a column, not as a pivot calculated field, for
reasons the Core concepts section explains. If your source is a Power Query
load, add it in the query instead, so it survives a refresh.

## Core concepts

### The pivot cache

When you create a PivotTable, Excel copies the source data into an internal
store — the pivot cache — and every subsequent drag, filter and sort operates
on that copy, not on your worksheet. This is why pivots are fast on large
tables. It is also why a pivot can be wrong.

Change the source data and the pivot does not update until you refresh it
(**PivotTable Analyze → Refresh**, or `Alt+F5`). Nothing on screen indicates
staleness. A pivot built on Monday and screenshotted on Thursday shows
Monday's numbers, confidently.

Three consequences for practice:

1. **Refresh before you read, always.** Make it reflex, not a checklist item.
2. **Set refresh-on-open** for delivered workbooks: PivotTable Options → Data →
   *Refresh data when opening the file*. A client who opens your workbook
   should not see numbers that predate their own data update.
3. **Pivots built on the same source share a cache**, which is what lets one
   slicer drive several pivots — and also means a grouping you apply in one
   pivot can appear in another. If two pivots must behave independently, build
   the second from a separate cache (paste the source reference afresh, or load
   the table twice).

Pivots built on the Power Pivot data model behave differently and better; see
Chapter 8.

### Rows, columns, values, filters — four roles, not four boxes

Treat the four drop zones as a grammar:

- **Rows** are the thing you are comparing. Usually the longer list.
- **Columns** are the dimension you want read side by side. Keep it short —
  three to six values. A pivot with twelve columns is a table nobody reads.
- **Values** are the measures. Every field here gets an aggregation.
- **Filters** are the scope statement. Anything in Filters is a claim about
  what the exhibit covers, and it must appear in the exhibit's title.

That last point is the one consultants get wrong. A pivot filtered to 2013 with
no title saying so is a document that will be misread, and you will be the one
who gets blamed. Write scope into the title: *"Revenue and gross margin by
channel and category, calendar 2013, excluding unassigned territory"*.

### Value field settings: sum, count, and the percentage family

Double-click a value field header to open Value Field Settings. Two tabs
matter.

**Summarize Values By** controls the aggregation. `Sum` for revenue, cost,
margin and units. `Count` counts non-empty rows — which on a fact table counts
*order lines*, not orders. The number of distinct orders is a different
question: `SalesOrderNumber` has one row per line, so `Count` of it
overstates orders by the average lines-per-order. Distinct Count is available
only on data-model pivots (Chapter 8); on a classic pivot you must prepare the
column or use a formula.

**Show Values As** is where the real power is, and it is underused:

- **`% of Grand Total`** — share of everything.
- **`% of Row Total`** / **`% of Column Total`** — share within the row or
  column. Use `% of Row` when rows are the entity and you want its mix;
  `% of Column` when columns are the entity.
- **`% of Parent Row Total`** — share within the immediate parent in a nested
  hierarchy. This is the right setting for a Category → Subcategory pivot,
  because it tells you each subcategory's share *of its own category*, which is
  the question you actually have. `% of Row Total` in the same layout gives
  share of the whole, which is a different and usually less useful number.
- **`Running Total In`** — cumulative along a field. Set it on a date field for
  year-to-date revenue.
- **`Difference From`** / **`% Difference From`** with base field `CalendarYear`
  and base item `(previous)` — year-on-year growth, computed by the pivot
  rather than by hand. This single setting replaces a column of error-prone
  manual subtraction.

You can place the same field in Values twice: once as `Sum` and once as
`% of Column Total`. That is how you get absolute and share side by side.

### Grouping dates and numbers

Right-click a date field in Rows → **Group** → choose Years, Quarters, Months.
Excel creates the grouping in the cache and adds pseudo-fields.

Use it, but know its limits. Grouping on `OrderDate` requires a real date
column; `OrderDateKey` is an integer (`20130715`) and grouping it will produce
numeric bins, which is nonsense. Either use the real `OrderDate`, or join
`DimDate` and use its `CalendarYear` and `EnglishMonthName` fields — which is
the better habit anyway, because `DimDate` also carries fiscal periods that
Excel's grouping cannot invent.

Numeric grouping bins a continuous field. Right-click `YearlyIncome` in Rows →
Group → start at 10,000, end at 170,000, by 20,000. You now have income bands
for segmentation. Choose the bin width deliberately: too narrow and the pattern
disappears into noise, too wide and it disappears into averages. State the bin
width in the exhibit.

### Sorting, and the `CommuteDistance` trap

Excel sorts text alphabetically, which is correct for product names and wrong
for anything ordinal stored as text. `DimCustomer[CommuteDistance]` sorts as:

```
0-1 Miles
1-2 Miles
10+ Miles
2-5 Miles
5-10 Miles
```

`10+ Miles` sits third. Any chart built on this is misleading, and a reader who
spots it stops trusting the deck.

The fix is a custom list. **File → Options → Advanced → Edit Custom Lists**,
enter the five values in the correct order, then right-click the field in the
pivot → **Sort → More Sort Options → More Options → Sort automatically**, and
ensure *First key sort order* uses your list. The order now persists through
refreshes and travels with the workbook.

The same trap applies to anything ordinal: `EnglishEducation` (Partial High
School → Graduate Degree), `BusinessType`, income bands you created as text,
and month names if you ever store them without `MonthNumberOfYear`.

For non-ordinal fields, sort by value instead: right-click a cell in the values
area → Sort → Largest to Smallest. This sorts rows by that measure and holds
through refresh, which is what you want for a ranked exhibit.

### The `NA` territory row

`DimSalesTerritory` contains a row — key 11 — named `NA`, which absorbs sales
with no assigned territory. Build an unfiltered territory pivot and it appears,
typically as a small but visible slice.

Do not filter it out quietly. Two defensible treatments:

1. **Exclude and disclose.** Filter it out, and put the excluded amount in a
   footnote: *"Excludes $X of revenue with no assigned sales territory (0.Y% of
   total)."* Quantify it with
   `=SUMIFS(Sales[SalesAmount], Sales[SalesTerritoryRegion], "NA")`.
2. **Keep and label.** Rename it `Unassigned` in the source and leave it in the
   exhibit.

The one indefensible treatment is deleting it and saying nothing, because the
exhibit then does not reconcile to total revenue and someone will eventually
notice. The general principle: **every exclusion is a disclosure**.

### Calculated fields, calculated items, and why ratios break

**PivotTable Analyze → Fields, Items & Sets → Calculated Field** lets you define
a new field from existing ones. For sums, it works:
`Margin = SalesAmount - TotalProductCost` is correct, because the difference of
sums equals the sum of differences.

For ratios, it is a trap. Define a calculated field
`MarginPct = (SalesAmount - TotalProductCost) / SalesAmount` and Excel evaluates
it **on the aggregated values of each cell**, which is actually what you want —
but the *subtotals and grand total* are then computed the same way from their
own aggregates, which is also correct. The genuine failure arrives when the
underlying field is itself a ratio, or when you compute the ratio in a source
column and then let the pivot `Sum` or `Average` it. Summing percentages is
meaningless; averaging them is an unweighted average that ignores the size of
each row. A category with one $5 order at 90% margin moves an unweighted
average as much as ten thousand orders at 10%.

The safe pattern, and the one to use by default:

1. Put `Sum of SalesAmount` and `Sum of Margin` in the pivot as separate value
   fields.
2. Define margin percentage as a calculated field dividing those two aggregates,
   or compute it outside the pivot with `GETPIVOTDATA`.

Never place a pre-computed percentage column into Values. If you see
`Average of MarginPct` in a pivot, that exhibit is wrong.

**Calculated items** create a new row or column from existing items — a
"Bikes + Components" combined row, for instance. They are fragile: they
interact badly with grouping and filtering, and they disable some Show Values As
options. Prefer restructuring the source or using a grouping column.

### `GETPIVOTDATA`

Click a cell outside a pivot, type `=`, then click a pivot cell, and Excel
writes something like:

```
=GETPIVOTDATA("Margin", $A$3, "Channel", "Reseller", "Category", "Bikes")
```

This looks verbose and it is the right answer. It references the pivot by
*field and item names*, not by cell address. Add a category, refresh, and the
rows shift — a `=B7` reference now points somewhere else, silently; the
`GETPIVOTDATA` reference still returns reseller bike margin.

Any summary cell, chart title, or model input that reads from a pivot should
use `GETPIVOTDATA`. Build the arguments from cells so the formula is
driveable: `=GETPIVOTDATA("Margin", $A$3, "Channel", $B$1, "Category", $A5)`.

If an item does not exist in the current pivot layout, `GETPIVOTDATA` returns
`#REF!` — loudly, which is the correct behaviour. Wrap it only where you have
decided what missing means: `=IFERROR(GETPIVOTDATA(...), 0)`.

To turn the auto-generation off when you genuinely want a cell reference:
PivotTable Analyze → Options dropdown → *Generate GetPivotData*.

### Slicers, timelines, and report filter pages

A **slicer** is a filter you can see and click. **Insert → Slicer**, pick a
field. A **timeline** is the date equivalent, with a draggable range.

Their real power is connection. Right-click a slicer → **Report Connections**
and tick every pivot it should drive. One `CalendarYear` slicer can now control
four pivots and their charts simultaneously — which is how a dashboard behaves
(Chapter 14). The constraint: connected pivots must share the same cache, so
build them from the same source table.

Visible filters beat hidden ones. A filter in the Filters zone is a dropdown
nobody notices; a slicer sitting on the sheet with `2013` highlighted is
self-documenting. For client-facing work, prefer slicers.

**Show Report Filter Pages** is the most undersold feature in Excel. Put
`SalesTerritoryGroup` in the Filters zone, then PivotTable Analyze → Options
dropdown → *Show Report Filter Pages*. Excel generates one worksheet per
territory group, each a complete copy of the pivot filtered to that group. A
three-region pack in one click. Use it for per-region appendices and per-manager
versions of the same exhibit.

### Pivot charts and their limits

A pivot chart is bound to its pivot and reconfigures itself when the pivot
changes — convenient, and the source of some frustration. Limitations worth
knowing: you cannot make an XY scatter or a bubble chart from a pivot; some
formatting resets on refresh; and the chart inherits the pivot's layout, so you
cannot reorder series independently.

When chart layout matters more than live binding — which on a client exhibit it
usually does — read the pivot into a small staging block with `GETPIVOTDATA`
and chart that. You get full control and a stable reference. Chapter 14 builds
on this pattern.

### When a pivot is the wrong tool

Use a formula block instead when:

- The output feeds a model and needs **fixed cell addresses**. Pivot rows move.
- The calculation is **not an aggregation** — a ratio of ratios, a weighted
  average with an unusual weight, an allocation, a scenario-dependent figure.
- You need **row-level output**, not a summary. `FILTER` beats a pivot for
  "show me the orders that match this".
- The result must **recalculate automatically**. Formulas do; pivots need a
  refresh.
- You need **distinct counts** and are not on the data model.

Use a pivot when exploring, when the shape of the question is still changing,
and when you need an answer in the next two minutes. Both tools, chosen on
purpose.

## Walkthrough

**Step 1 — prepare the source.** Confirm `Sales` is a proper Excel table with
the `Margin` calculated column described above, plus `Channel`, `Category`,
`CalendarYear` and `SalesTerritoryGroup`. Pivots on a table rather than a range
pick up new rows on refresh; pivots on a range do not.

**Step 2 — create the pivot.** Select any cell in `Sales` → **Insert →
PivotTable → New Worksheet**. Name the sheet `Pvt_Channel_Category`. Naming
sheets by content is not fussiness; Chapter 16 explains why a reviewer needs it.

**Step 3 — lay out the grammar.**

- Rows: `Category`
- Columns: `Channel`
- Values: `Sum of SalesAmount`, then `Sum of Margin`

You now have a 4×2 grid of revenue and a 4×2 grid of margin dollars. Read it
before going further — this is the moment the Bikes assumption starts to wobble,
because Bikes dominate revenue by a wide margin while their margin dollars are
proportionally much closer to the other categories.

**Step 4 — margin percentage, computed correctly.** Fields, Items & Sets →
Calculated Field. Name: `MarginPct`. Formula:

```
= Margin / SalesAmount
```

Add it to Values and format as a percentage with one decimal. Because both
operands are summed before dividing, every cell — including subtotals and the
grand total — is a true weighted margin percentage. Verify it: the grand total
cell should equal

`=SUM(Sales[Margin])/SUM(Sales[SalesAmount])`

computed on a blank cell elsewhere. If those two disagree, you have an
aggregation problem, not a rounding problem.

**Step 5 — sort and format defensibly.** Right-click a revenue cell → Sort →
Largest to Smallest, so categories rank by revenue. Set the number format
through Value Field Settings → Number Format, not by formatting cells — cell
formatting is lost on refresh and layout change. Turn off **Grand Totals for
columns** if the row grand total is the only one that means anything, and
switch to **Report Layout → Show in Tabular Form** with **Repeat All Item
Labels** if anyone will ever copy this into another tool.

**Step 6 — title it with its scope.** In the cell above the pivot, write the
full claim: *"Revenue, gross margin and margin % by category and channel, all
years, all territories. Margin = SalesAmount − TotalProductCost; excludes tax
and freight."* That last clause pre-empts the first question you will be asked.

**Step 7 — the territory view, with `NA` handled.** Build a second pivot on the
same source: Rows `SalesTerritoryGroup`, Columns `Channel`, Values revenue and
`MarginPct`. The `NA` group appears. Compute its size on a blank cell:

`=SUMIFS(Sales[SalesAmount], Sales[SalesTerritoryRegion], "NA")`

and as a share:

`=SUMIFS(Sales[SalesAmount], Sales[SalesTerritoryRegion], "NA")/SUM(Sales[SalesAmount])`

Then filter it out of the pivot via the row-label dropdown, and write the
footnote underneath using those two cells, so the disclosed figure updates with
the data rather than being typed once and forgotten.

**Step 8 — wire a slicer.** Insert → Slicer → `CalendarYear`. Right-click →
Report Connections → tick both pivots. Both views now move together, and the
selected year is visible on the sheet rather than hidden in a dropdown.

**Step 9 — extract the headline numbers.** On a `Findings` sheet, pull the four
numbers your recommendation will rest on using `GETPIVOTDATA` — total revenue,
bike margin percentage, accessory margin percentage, and reseller share of
revenue. These cells are the ones Chapter 15's exhibit will cite, and because
they are `GETPIVOTDATA` they survive every layout change you make between now
and the final deck.

## Common pitfalls

- **Reading a stale pivot.** Refresh before you look. Set refresh-on-open for
  anything you deliver.
- **`Count` of `SalesOrderNumber` reported as "number of orders".** It counts
  order *lines*. The grain of the fact table is one row per line, as the data
  dictionary states.
- **An unweighted average of a percentage column.** Wrong at every level of
  aggregation. Compute ratios from summed numerator and denominator.
- **A filter in the Filters zone with no mention in the title.** The exhibit now
  makes a claim it does not state.
- **The `NA` territory row appearing silently**, or disappearing silently.
  Disclose and quantify.
- **`CommuteDistance` and `EnglishEducation` sorted alphabetically.** Set a
  custom list once.
- **Cell-formatting a pivot instead of using Number Format in Value Field
  Settings.** Lost on the next refresh.
- **Referencing a pivot cell by address.** Use `GETPIVOTDATA`, or accept that a
  future layout change will break your model without an error message.
- **Grouping on `OrderDateKey`.** It is an integer; you will get numeric bins.
  Use the real date or `DimDate`.
- **Calculated items.** Fragile, and they disable features you will want.
  Restructure the source instead.

## Exercises

1. Build a pivot of revenue and margin percentage by `Subcategory` nested under
   `Category`, using `% of Parent Row Total` for a third value field. Explain
   in writing what `% of Parent Row Total` tells you that `% of Row Total` does
   not, and which you would show a category manager.
2. Using the same source, produce year-on-year revenue growth by category
   entirely through Show Values As — no manual subtraction, no helper columns.
   State which base field and base item you used and why the first year is
   blank.
3. Deliberately build the wrong margin exhibit: add a `MarginPct` column to the
   source table, place it in Values as `Average`, and compare it cell by cell to
   the correct calculated-field version. Quantify the largest discrepancy, name
   the category it occurs in, and explain the mechanism in two sentences.
4. Fix the `CommuteDistance` sort with a custom list, then build revenue per
   customer by commute distance band. Refresh the pivot and confirm the order
   holds. State whether the resulting pattern is strong enough to act on, and
   what confound you would want ruled out first.
5. Create a pivot with `SalesTerritoryGroup` in the Filters zone and use Show
   Report Filter Pages to generate the per-region pack. Then write the single
   footnote that every one of those sheets needs in order to be defensible on
   its own, away from the rest of the workbook.
6. Replace five hard-coded cell references in a summary block with
   `GETPIVOTDATA` equivalents, then add a category to the source data and
   refresh. Document what would have broken under the original references.

## Solutions

**1.** With `Category` and `Subcategory` both in Rows, `% of Parent Row Total`
gives each subcategory's share of its own category — Mountain Bikes as a share
of Bikes — while `% of Row Total` gives its share of all company revenue. The
category manager wants the former: they control mix within their category and
cannot act on the denominator. Show the parent-total version in the body and
keep grand-total share for the executive summary, where the comparison across
categories is the point.

**2.** Add `Sum of SalesAmount` to Values twice. On the second, Show Values As →
`% Difference From`, base field `CalendarYear`, base item `(previous)`. The
first year is blank because there is no prior period to compare against — which
is correct and should be left blank rather than shown as zero or 100%. A zero
would read as flat growth, which is a false statement about the data.

**3.** The `Average of MarginPct` version will differ most in Accessories,
where a large number of small-value lines sit alongside a few larger ones: the
unweighted average treats a $3 line and a $300 line as equal observations, so it
drifts toward the margin of the typical *line* rather than the margin of the
typical *dollar*. Expect the unweighted figure to read several percentage points
higher than the weighted one in Accessories and Clothing, and to be closer in
Bikes where line values are more uniform. The mechanism is simply that
`AVERAGE` weights by row count while the business question weights by revenue.
This exercise is worth doing precisely because the wrong number looks entirely
plausible — there is no error, no `#DIV/0!`, nothing to notice. The only
defence is the rule: ratios are computed from aggregates, never aggregated
themselves.

**4.** Custom list `0-1 Miles, 1-2 Miles, 2-5 Miles, 5-10 Miles, 10+ Miles`,
then Sort automatically with that list as the first key. Revenue per customer
typically rises with commute distance, which is tempting to read as "distant
customers are more valuable". Before acting on it, rule out the obvious
confound: commute distance correlates with `YearlyIncome` and with home
ownership in this dataset, so the pattern may be an income effect wearing a
geography costume. Chapter 11's regression is the tool that separates them; a
pivot can only show you that the question exists.

**5.** Every generated sheet needs the scope and definition travelling with it,
because these sheets get emailed individually: *"Revenue and gross margin by
category, [region], calendar 2011–2014. Margin = SalesAmount − TotalProductCost,
excluding tax and freight. Excludes revenue with no assigned sales territory."*
Without it, a regional sheet is a set of unlabelled numbers, and the first
person to compare two regions on different years will produce a wrong
conclusion with your name on it.

**6.** `=GETPIVOTDATA("Margin", $A$3, "Category", $A5, "Channel", $B$1)` and
similar. Under the original `=B7`-style references, adding a category inserts a
row and every reference below the insertion point silently shifts to the wrong
category — no error, no `#REF!`, just four wrong numbers in the summary block.
`GETPIVOTDATA` either returns the right value or returns `#REF!`, and a visible
error is always cheaper than a plausible wrong number.

## Key takeaways

- The pivot cache means a pivot shows the data as of its last refresh. Refresh
  before reading; set refresh-on-open before delivering.
- Rows, columns, values and filters are four roles. Anything in Filters is a
  claim about scope and belongs in the title.
- Show Values As — `% of Parent Row Total`, `Running Total In`,
  `% Difference From (previous)` — replaces most manual arithmetic, and is
  harder to get wrong than doing it by hand.
- Ratios are computed from aggregates. Put summed numerator and summed
  denominator in the pivot and divide; never aggregate a percentage column.
- Ordinal text fields sort alphabetically and therefore wrongly. Set a custom
  list once, for `CommuteDistance`, `EnglishEducation` and anything like them.
- Every exclusion is a disclosure. Quantify the `NA` territory row with a
  formula so the footnote updates itself.
- Reference pivots with `GETPIVOTDATA`, never by cell address. Layouts move;
  field names do not.
- Pivots are for exploration and for views. Formula blocks are for anything that
  feeds a model, needs stable addresses, or has to recalculate on its own.
