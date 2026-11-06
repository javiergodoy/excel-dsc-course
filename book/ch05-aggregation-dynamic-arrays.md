# Chapter 5 — Conditional Aggregation and Dynamic Arrays

You can now read the AdventureWorks model, clean it and join it. This chapter
is where you start producing numbers that someone else will act on.

There are two ways to summarise a table in Excel. One is a PivotTable, which
Chapter 6 covers. The other is formulas — and formulas are what you reach for
when the summary has to feed something else: a model, a scenario, a chart whose
layout you control, an exhibit that must recalculate the instant an input
changes. A pivot is a snapshot you refresh. A formula block is live.

Consultants need both, and need to know which one the situation calls for. The
distinction is not about elegance. It is about whether your deliverable breaks
silently when the data updates the night before the steering committee.

## Learning objectives

By the end of this chapter you will be able to:

- Write `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` and `MAXIFS` with multiple criteria,
  including operators, date bounds, wildcards and cell references.
- Recognise the cases the `*IFS` family cannot handle, and use `SUMPRODUCT`
  for them — weighted averages, row-by-row arithmetic, and boolean logic.
- Use `FILTER`, `UNIQUE`, `SORT`, `SORTBY` and `SEQUENCE`, and reference spill
  ranges with the `#` operator.
- Make long formulas readable with `LET`, and reusable with `LAMBDA`.
- Build a summary block whose row and column headers are themselves formulas,
  so new categories and new years appear without anyone editing the workbook.
- Explain why a formula-driven summary beats a pivot when the output feeds a
  model, and why the reverse is true for exploration.
- Keep a 60,000-row workbook responsive instead of letting it degrade into a
  three-second pause on every keystroke.

## The business question

> The AdventureWorks management team wants a one-page operating summary:
> revenue and gross margin by product category and year, across both sales
> channels, that updates itself when next month's data lands. It has to tie to
> the reported revenue total exactly, and it has to be the single block that
> feeds the growth model in Chapter 13.

Read what that question is actually asking for. Not "show me revenue by
category" — you could do that with a pivot in ninety seconds. It is asking for
something *structural*: a block with a known shape, known cell addresses, and
no manual maintenance, because other parts of the workbook will point at it.

That requirement is what rules out a pivot. A pivot's cells move. Add a
category and the rows below shift down; your model's `=B14` now points at the
wrong thing and nothing warns you. A formula block built on `UNIQUE` grows in a
place you decided in advance.

## Data you will use

From `FactInternetSales` and `FactResellerSales`, loaded as tables named
`InternetSales` and `ResellerSales`, or as the combined `Sales` table you built
in Chapter 4:

- `SalesAmount` — revenue.
- `TotalProductCost` — cost of goods for the line.
- `OrderQuantity` — units.
- `ProductKey`, `OrderDateKey`, `SalesTerritoryKey`.

From the dimensions, via the lookup columns you added in Chapter 4:

- `Category` — derived from `DimProduct` → `DimProductSubcategory` →
  `DimProductCategory`. Four values: Bikes, Components, Clothing, Accessories.
- `CalendarYear` — derived from `DimDate` on `OrderDateKey`.
- `Channel` — a text column, `"Internet"` or `"Reseller"`, added when the two
  fact tables were unioned.

Full column definitions are in
[the data dictionary](../data/adventureworks-data-dictionary.md). As always,
"margin" means `SalesAmount − TotalProductCost`, and there is no margin column
in the source — you derive it, which is exactly why the definition has to live
in one place.

One assumption, stated up front: this chapter works on the combined `Sales`
table with a `Channel` column. If you have kept the two fact tables separate,
every formula here works unchanged on each one; you just add the two results,
and you lose the ability to slice by channel in a single block.

## Core concepts

### The `*IFS` family, and what the criteria actually mean

`SUMIFS` sums a range where every condition holds. Its argument order is the
opposite of `SUMIF`'s, which is a genuine design wart and a reliable source of
errors:

```
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
```

Revenue for Bikes in 2013:

`=SUMIFS(Sales[SalesAmount], Sales[Category], "Bikes", Sales[CalendarYear], 2013)`

The criteria arguments are where the real subtlety lives, because a criterion
is a *string that Excel parses*, not an expression you write.

- **Equality** is the bare value: `"Bikes"`, `2013`.
- **Operators** must be inside the string: `">1000"`, `"<=0"`, `"<>Bikes"`.
- **A cell reference with an operator** needs concatenation:
  `">" & $B$2`. Writing `">$B$2"` compares against the literal text `$B$2` and
  returns zero without complaint. This is the single most common `SUMIFS` bug.
- **Dates** work the same way, and you should always anchor them to real dates
  rather than typed text: `">=" & DATE(2013,1,1)`.
- **Wildcards** work on text criteria: `"Road*"` matches any subcategory
  starting with Road; `"?-?"` matches single characters. If you need a literal
  asterisk, escape it with a tilde: `"~*"`.
- **Blank and non-blank**: `""` matches truly empty cells, `"<>"` matches
  non-empty. These behave differently from cells containing a zero-length
  string, which is a distinction that matters on the sparse `Color` and `Size`
  columns in `DimProduct`.

`COUNTIFS`, `AVERAGEIFS`, `MINIFS` and `MAXIFS` take the same criteria grammar.
`AVERAGEIFS` returns `#DIV/0!` when nothing matches — which is useful, because
a silent zero would be worse. Wrap it only when you have decided what an empty
set should display:

`=IFERROR(AVERAGEIFS(Sales[SalesAmount], Sales[Category], $A5, Sales[CalendarYear], B$4), "n/a")`

Note `$A5` and `B$4`: mixed references, so one formula fills across and down a
whole grid. Getting the dollar signs right once is worth more than any function
you will learn this chapter.

### Where `*IFS` runs out, and `SUMPRODUCT` takes over

`SUMIFS` can only sum a range that already exists. It cannot sum an
*expression* evaluated row by row. Three situations hit this wall constantly.

**Margin, when there is no margin column.** You could add a calculated column
`=[@SalesAmount]-[@TotalProductCost]` to the table — and often you should. But
when you cannot modify the source table, `SUMPRODUCT` computes it in place:

```
=SUMPRODUCT((Sales[Category]="Bikes")*(Sales[CalendarYear]=2013)
           *(Sales[SalesAmount]-Sales[TotalProductCost]))
```

The mechanics are worth understanding rather than memorising. Each comparison
produces an array of `TRUE`/`FALSE`. Multiplying them coerces to 1/0 and gives
you AND logic. Multiply by the value expression and you have summed only the
matching rows. Addition gives OR logic:
`((Sales[Category]="Bikes")+(Sales[Category]="Components"))`, though you must
guard against double counting when conditions can both be true.

**Weighted averages.** The average of the `UnitPrice` column is not the average
price AdventureWorks achieved, because reseller lines carry many units each.
The honest number is revenue divided by units:

```
=SUMPRODUCT((Sales[Category]=$A5)*Sales[SalesAmount])
 /SUMPRODUCT((Sales[Category]=$A5)*Sales[OrderQuantity])
```

A client who is shown an unweighted average price and later discovers it will
not trust any other number in the deck. Say which you used.

**Boolean columns stored as text.** `DimCustomer[HouseOwnerFlag]` is text
`"0"`/`"1"`, as Chapter 3 warned. `SUMPRODUCT` lets you coerce it inline:
`=SUMPRODUCT(--(Customers[HouseOwnerFlag]="1"))`. The double negative turns
booleans into numbers.

`SUMPRODUCT` is slower than `SUMIFS` because it evaluates every row rather than
using Excel's indexed criteria path. On 60,000 rows you will not notice. On a
grid of 200 `SUMPRODUCT` cells over 60,000 rows you will. Reach for it when you
need it, not by default.

### Dynamic arrays: the change that mattered

Before dynamic arrays, a formula returned one value into one cell. Now a single
formula returns an array that *spills* into neighbouring cells, and that array
resizes when the data does. This is not a convenience feature. It changes what
a well-built workbook looks like.

- `=UNIQUE(Sales[Category])` returns the distinct categories, in first-appearance
  order. Wrap it in `SORT` for stability: `=SORT(UNIQUE(Sales[Category]))`.
- `=FILTER(Sales[SalesOrderNumber], (Sales[Category]="Bikes")*(Sales[SalesAmount]>3000), "none")`
  returns the matching orders, with a fallback value instead of `#CALC!` when
  nothing matches. Always supply the third argument.
- `=SORT(array, sort_index, order, by_col)` sorts; `SORTBY(array, by_array1, order1, ...)`
  sorts one range by another, which is how you rank categories by revenue
  without adding a helper column.
- `=SEQUENCE(4, 1, 2011, 1)` generates the year headers 2011–2014. Better still,
  derive them: `=TRANSPOSE(SORT(UNIQUE(Sales[CalendarYear])))`.

The `#` operator is what makes this compose. If `A5` holds a spilled `UNIQUE`,
then `A5#` means "the whole spill range, whatever size it currently is". A
chart or a downstream formula pointed at `A5#` tracks the data automatically.
`COUNTA(A5#)` tells you how many categories there are today.

### `LET` for readability, `LAMBDA` for reuse

A formula that takes thirty seconds to parse is a formula nobody will audit,
and an unaudited formula is a liability. `LET` names intermediate results:

```
=LET(
   rev,  SUMIFS(Sales[SalesAmount],      Sales[Category], $A5, Sales[CalendarYear], B$4),
   cost, SUMIFS(Sales[TotalProductCost], Sales[Category], $A5, Sales[CalendarYear], B$4),
   IF(rev=0, "n/a", (rev-cost)/rev)
 )
```

Besides reading better, this evaluates each `SUMIFS` once rather than twice —
`LET` genuinely improves performance on repeated sub-expressions.

`LAMBDA` goes further: define the calculation once, name it in **Formulas →
Name Manager**, and call it everywhere. Define `MarginPct` as:

```
=LAMBDA(cat, yr,
   LET(rev,  SUMIFS(Sales[SalesAmount],      Sales[Category], cat, Sales[CalendarYear], yr),
       cost, SUMIFS(Sales[TotalProductCost], Sales[Category], cat, Sales[CalendarYear], yr),
       IF(rev=0, "n/a", (rev-cost)/rev)))
```

Then every cell in your margin grid is `=MarginPct($A5, B$4)`. The definition of
margin now exists in exactly one place in the workbook. When the client asks
whether freight is in cost — and they will — you change one named formula and
every exhibit moves together. Chapter 16 makes this discipline the centre of
the audit argument.

### `GROUPBY` and `PIVOTBY`

Current Microsoft 365 builds include `GROUPBY` and `PIVOTBY`, which produce a
complete aggregated table from one formula:

```
=PIVOTBY(Sales[Category], Sales[CalendarYear], Sales[SalesAmount], SUM, , 0, , 0)
```

They are excellent for exploration and increasingly good for delivery. Two
cautions for consulting work. First, availability: if your client is on
Excel 2021 or a semi-annual channel build, these do not exist, and a deck built
on them cannot be handed over. Second, the output is a single spill range whose
width and height are data-dependent, which reintroduces the moving-cell problem
the `UNIQUE`-plus-`SUMIFS` pattern was designed to solve. Use them to check
your work; build the deliverable on the pattern that survives being emailed.

### Performance discipline

Three habits keep a 60,000-row workbook fast:

1. **Never reference whole columns.** `SUMIFS(A:A, B:B, "Bikes")` asks Excel
   about a million rows. `Sales[SalesAmount]` asks about the rows that exist.
   Structured references are faster *and* correct when the table grows.
2. **Avoid repeated expensive sub-expressions.** Use `LET`, or compute once in
   a helper cell and reference it. A grid of `SUMPRODUCT` formulas each
   re-scanning the fact table is the usual cause of a workbook that pauses.
3. **Avoid volatile functions in large chains.** `NOW`, `TODAY`, `RAND`,
   `OFFSET` and `INDIRECT` force recalculation of everything downstream on
   every edit. `INDIRECT` is also invisible to Excel's dependency tracer, so
   trace-precedents audits miss it — doubly disqualifying.

If a workbook becomes sluggish, switch to **Formulas → Calculation Options →
Manual** while building, and use **Formulas → Evaluate Formula** to find the
expensive step rather than guessing.

## Walkthrough

Build the operating summary: revenue and margin percentage by category and
year, self-updating, tying to a control total.

**Step 1 — reserve the space.** On a new sheet named `Summary`, put the block
title in `A1`. Leave `A2:A3` for notes. The grid starts at `A4`. Decide now
that nothing else goes on this sheet to the right of column `H` or below row
`30`; the block will grow into that space and must not collide with anything.

**Step 2 — the year headers, derived.** In `B4`:

`=TRANSPOSE(SORT(UNIQUE(Sales[CalendarYear])))`

This spills across `B4:E4` for a four-year dataset. When 2015 data arrives, it
spills to `F4` on its own.

**Step 3 — the category labels, derived.** In `A5`:

`=SORT(UNIQUE(Sales[Category]))`

Four rows, `A5:A8`. If AdventureWorks adds a category, a fifth row appears.

**Step 4 — revenue.** In `B5`, one formula that fills the whole grid:

`=SUMIFS(Sales[SalesAmount], Sales[Category], $A5, Sales[CalendarYear], B$4)`

Copy across and down to cover a grid slightly larger than the current spills —
say `B5:H12`. Cells beyond the live data return `0`; suppress them with a
custom number format or an `IF` on the header, but do not delete them, because
they are the room the block grows into.

**Step 5 — margin percentage, via `LAMBDA`.** Define `MarginPct` in Name
Manager as above. In `B15` (leaving a blank row and a sub-heading in row 14):

`=MarginPct($A15, B$14)`

with `A15` holding `=A5#` and `B14` holding `=B4#`, so the second grid inherits
its headers from the first. One edit to the category list propagates to both.

**Step 6 — validate against a control total.** This step is not optional, and
in client work it is the step that saves you. In `A20`:

`=SUM(B5:H12)`

In `B20`, the independent total:

`=SUM(Sales[SalesAmount])`

In `C20`, the reconciliation:

`=IF(ROUND(A20-B20,2)=0, "TIES", "BREAK: " & TEXT(A20-B20,"#,##0.00"))`

If this says `BREAK`, the usual cause is a category or year falling outside the
grid — a product with a null `ProductSubcategoryKey` producing a blank
`Category`, or an order date outside the years your header formula found. Both
are real findings about the data, not nuisances to paper over. A blank-category
row means products are being sold that your category hierarchy does not
classify, and the client should know.

Add a second control on units if margin matters:
`=SUM(Sales[OrderQuantity])` against the sum of a parallel units grid.

**Step 7 — make the block addressable.** Name the revenue grid `Summary_Revenue`
and the margin grid `Summary_MarginPct` in Name Manager. Chapter 13's model
references those names, never cell addresses. The block can now move within the
sheet without breaking anything downstream.

## Common pitfalls

- **`">$B$2"` instead of `">" & $B$2`.** Returns zero, looks like a real
  answer. Build the habit of concatenating every operator criterion.
- **Averaging an average.** `AVERAGE` of a column of category margin
  percentages is not the company margin percentage. Recompute from summed
  revenue and summed cost at every level of aggregation.
- **Calculated-ratio grids that do not tie.** Related to the above: a total row
  computed as the sum of percentages is meaningless. Compute the total row from
  the underlying sums.
- **Spill collisions.** A dynamic array that cannot spill returns `#SPILL!`.
  The usual cause is a stray value left in the spill path from an earlier
  draft. Select the whole intended region and clear it before pointing a
  dynamic array at it.
- **`FILTER` without a fallback.** Returns `#CALC!` when nothing matches, which
  propagates into every downstream formula. Always pass the third argument.
- **Forgetting that `UNIQUE` includes blanks.** A blank `Category` becomes a
  row labelled `0` in your block. Filter it explicitly —
  `=SORT(UNIQUE(FILTER(Sales[Category], Sales[Category]<>"")))` — and report the
  excluded rows separately rather than silently dropping them.
- **Whole-column references.** Slow, and wrong the moment the table grows past
  the range you typed.
- **`INDIRECT` to build criteria dynamically.** Tempting, volatile, and
  invisible to the dependency tracer. Use `CHOOSE`, `SWITCH`, `INDEX` or
  `FILTER` instead.

## Exercises

1. Write a single `SUMIFS` returning reseller revenue for Mountain Bikes in
   calendar 2013 where the line discount was greater than zero, using cell
   references for both the subcategory and the year so the formula can be
   driven from input cells.
2. Compute the weighted average selling price for each product category, and
   alongside it the unweighted `AVERAGE` of `UnitPrice`. Quantify the gap for
   each category and explain, in two sentences per category, which number you
   would put in front of a client and why.
3. Build a self-updating block of **units sold** by subcategory and year, with
   subcategory rows sorted by total units descending — not alphabetically.
   Use `SORTBY` rather than a helper column, and reconcile the grid total to
   `=SUM(Sales[OrderQuantity])`.
4. Define a `LAMBDA` named `RevShare` taking a category and a year and returning
   that cell's share of the year's total company revenue. Use it to build a
   percentage-of-year grid, and verify that each year's column sums to exactly
   100%.
5. Using `FILTER`, produce a list of the ten largest single order lines by
   margin across both channels, showing order number, product name, revenue,
   cost and margin. Then state what you notice about which channel and
   category they come from, and what you would check before reporting it.
6. Take any one of the grids above and deliberately break it: add a row to the
   source table with a blank `Category`. Document what the control total shows,
   where the row lands, and the smallest change that makes the failure visible
   rather than silent.

## Solutions

**1.** `=SUMIFS(ResellerSales[SalesAmount], ResellerSales[Subcategory], $B$2, ResellerSales[CalendarYear], $B$3, ResellerSales[UnitPriceDiscountPct], ">0")`
with `B2` holding `Mountain Bikes` and `B3` holding `2013`. The literal
`">0"` is fine because the threshold is constant; if it came from a cell it
would need `">" & $B$4`. The result will be a small fraction of Mountain Bike
revenue, because `DimPromotion` key 1 (`No Discount`) covers most lines — and
that sparsity is itself the finding. Any promotion conclusion drawn from this
subset needs its sample size stated.

**2.** Weighted: `=SUMPRODUCT((Sales[Category]=$A5)*Sales[SalesAmount])/SUMPRODUCT((Sales[Category]=$A5)*Sales[OrderQuantity])`.
Unweighted: `=AVERAGEIFS(Sales[UnitPrice], Sales[Category], $A5)`. The two
diverge most for Accessories and Clothing, where reseller lines carry high
`OrderQuantity` at low unit prices — the unweighted average is pulled up toward
the internet list price, which no reseller actually paid. For Bikes the gap is
narrower because internet orders are almost always a single unit. Report the
weighted figure, labelled "revenue per unit", and keep the unweighted one out
of the deck entirely; it answers a question nobody asked.

**3.** Row labels: `=SORTBY(UNIQUE(Sales[Subcategory]), MAP(UNIQUE(Sales[Subcategory]), LAMBDA(s, SUMIFS(Sales[OrderQuantity], Sales[Subcategory], s))), -1)`.
If `MAP` is unavailable, compute the totals in an adjacent helper column and
`SORTBY` on that — equally defensible, and easier for a reviewer to follow.
The reconciliation will tie only if you have handled the null-subcategory
products; expect a break equal to their units, and resolve it by adding an
explicit `Unclassified` row rather than by widening the grid until the
difference disappears.

**4.** `=LAMBDA(cat, yr, LET(r, SUMIFS(Sales[SalesAmount], Sales[Category], cat, Sales[CalendarYear], yr), t, SUMIFS(Sales[SalesAmount], Sales[CalendarYear], yr), IF(t=0, "n/a", r/t)))`.
Columns sum to 100% only when every revenue row is captured by some category,
so this grid doubles as a completeness test — a column summing to 97% is
telling you three percent of revenue has no category. That is a better
diagnostic than the control total in Step 6, because it localises the problem
to a year.

**5.** `=TAKE(SORT(FILTER(CHOOSECOLS(Sales, 1, 2, 3), Sales[SalesAmount]>0), 3, -1), 10)`
is the compact form, but build it in stages: add a `Margin` calculated column
first, then `SORT` on it, then `TAKE` the top ten. Expect the list to be
dominated by reseller Bikes lines, because `OrderQuantity` multiplies through.
Before reporting it, check whether any line is a data artefact — a quantity or
price outlier — and check whether the same order number appears repeatedly,
which would make this a story about one customer rather than about a product.

**6.** The blank-category row lands outside the `UNIQUE` row labels only if you
filtered blanks out, in which case the control total breaks by that row's
revenue. If you did not filter, a row labelled `0` appears and the total ties —
which is worse, because the workbook now silently contains a category that does
not exist in the business. The smallest fix is a dedicated reconciliation cell:
`=SUMIFS(Sales[SalesAmount], Sales[Category], "")` displayed next to the control
total with a label like "revenue with no category". Visible, quantified, and
impossible to present by accident.

## Key takeaways

- `SUMIFS` criteria are parsed strings. Concatenate every operator with `&`,
  and never compare against a reference inside quotes.
- `SUMPRODUCT` is the escape hatch for row-by-row arithmetic, weighted averages
  and text-encoded booleans. Use it deliberately, not by habit.
- Dynamic arrays let row and column headers be *derived* rather than typed. A
  block built on `UNIQUE` and `SORT` absorbs new categories and new years with
  no edits, which is what "self-updating" actually means.
- `LET` makes formulas auditable; `LAMBDA` makes a definition singular. Define
  margin once and every exhibit in the deck agrees by construction.
- Build deliverables on `UNIQUE`-plus-`SUMIFS` rather than `PIVOTBY` when the
  workbook has to survive being emailed to a client on an older Excel.
- Every summary block gets a control total and an explicit reconciliation cell.
  A grid that does not tie is not a formatting problem; it is a finding.
- Structured references, `LET` for repeated sub-expressions, and no volatile
  functions. That is the whole of performance discipline at this data size.
