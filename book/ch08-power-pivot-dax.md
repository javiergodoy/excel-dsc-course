# Chapter 8 — Power Pivot and DAX Measures

Chapter 7 ended with one wide `Sales` table: 120,000 rows carrying channel,
margin, product category and territory on every line. That table works, and for
many engagements it is enough.

It is also the wrong architecture, and this chapter explains why. The wide table
repeats `"Mountain Bikes"` sixty thousand times. It fixes the definition of
margin in a column, so changing it means re-running the pipeline. And it cannot
answer a question like "revenue year-to-date versus the same period last year, by
category, for the reseller channel only" without a sheet full of helper columns
that nobody will want to audit.

Power Pivot replaces the wide table with a *model*: narrow tables joined by
relationships, and calculations defined once as *measures* that recompute for
whatever slice of the model the reader is looking at. It is the difference
between a spreadsheet that contains answers and a model that answers questions.

## Learning objectives

By the end of this chapter you will be able to:

- Explain why a star schema with relationships beats a flattened table, in terms
  of file size, maintenance and analytical reach.
- Load tables to the data model and build relationships between them.
- Mark a date table and say why time intelligence is broken without it.
- Decide correctly between a calculated column and a measure.
- Distinguish implicit from explicit measures, and write only explicit ones.
- Write the core DAX vocabulary: `SUM`, `SUMX`, `DIVIDE`, `CALCULATE`.
- Reason about filter context and modify it with `ALL`, `ALLEXCEPT`,
  `KEEPFILTERS` and `FILTER`.
- Compute year-to-date and prior-year comparisons with `TOTALYTD`,
  `SAMEPERIODLASTYEAR` and `DATEADD`.
- Use `RELATED` and `RELATEDTABLE` to cross a relationship inside a calculation.
- Recognise a many-to-many problem and handle it with a bridge table or
  `CROSSFILTER`.
- Work productively on macOS, where Power Pivot does not exist.

## The business question

The question this chapter answers is the one that ends most arguments in client
meetings:

> **Two analysts on the same team produced two different gross margin numbers for
> the Bikes category in FY2013. Both used the same source data. How do we build
> an environment in which that cannot happen again?**

And the analytical question it unlocks, which the previous chapters could not
reach cleanly:

> **Which product categories are growing, which are growing profitably, and how
> does each channel's year-to-date performance compare with the same period last
> year?**

The first question is about governance and the second is about analysis, and the
same mechanism answers both. A measure is a single, named, centrally-stored
definition of a number. Once `[Margin %]` exists in the model, every PivotTable
in every workbook connected to that model computes margin percentage the same
way, because there is only one way. The disagreement becomes structurally
impossible rather than discouraged.

The second question needs time intelligence, which needs a real date table,
which needs relationships. That is the whole chapter in one dependency chain.

## Data you will use

From [the AdventureWorks data dictionary](../data/adventureworks-data-dictionary.md),
loaded as **separate tables** — this is the point:

| Table | Grain | Role |
|---|---|---|
| `Sales` | One row per order line, both channels, from Chapter 7's append | Fact |
| `Products` | One row per product version | Dimension |
| `Customers` | One row per internet customer | Dimension |
| `Resellers` | One row per reseller | Dimension |
| `Calendar` (`DimDate`) | One row per day | Date dimension |
| `Territory` (`DimSalesTerritory`) | One row per territory | Dimension |

Key columns for the relationships: `Sales[ProductKey]` → `Products[ProductKey]`,
`Sales[OrderDateKey]` → `Calendar[DateKey]`, `Sales[CustomerKey]` →
`Customers[CustomerKey]`, `Sales[SalesTerritoryKey]` →
`Territory[SalesTerritoryKey]`, `Sales[ResellerKey]` →
`Resellers[ResellerKey]`.

Important change from Chapter 7: **strip the lookup columns out of `Sales`.** The
expanded `EnglishProductCategoryName`, `SalesTerritoryGroup` and the rest were
necessary when `Sales` had to stand alone. In a model they are redundant — the
relationship supplies them — and they are the bulk of the file size. Edit the
`Sales` query, delete the merge-and-expand steps, and keep only keys, quantities
and amounts. Keep `Channel`, which is a genuine attribute of the fact and
belongs nowhere else. Keep `Margin` for now; you will delete it in the
walkthrough once a measure replaces it, and comparing the two is instructive.

Assume a Windows build of Microsoft 365 with the Power Pivot add-in enabled
(**File → Options → Add-ins → COM Add-ins → Microsoft Power Pivot for Excel**).
The macOS route is covered at the end of **Core concepts**.

## Core concepts

### Why a model, and not a wider table

Three reasons, in increasing order of importance.

**Size and speed.** The data model stores columns, not rows, and compresses each
column by its distinct values. A `Channel` column with two distinct values across
120,000 rows costs almost nothing compressed; the same column as worksheet text
costs real bytes per row. Replacing repeated text with an integer key that joins
to a short dimension routinely cuts a file by 80–90%. On AdventureWorks volumes
this is a convenience. On a client's 20 million transaction rows it is the
difference between a workbook that opens and one that does not.

**Maintenance.** Add `ProductLine` to the catalogue and in the flat world you edit
the pipeline, re-expand the merge, refresh 120,000 rows and re-point every
PivotTable field. In the model you add one column to `Products` and it is
immediately available to every pivot, because the relationship was already there.

**Analytical reach.** This is the real reason. Worksheet formulas and ordinary
PivotTables compute over the rows in front of them. A measure computes over
whatever the model's filters currently select — and it can *change* those filters
mid-calculation. "Revenue as a percentage of the category total" is a nuisance
with formulas; it is one line of DAX that works in any pivot, at any level, with
any slicers applied. That capability does not exist in the flat table, at any
amount of effort.

### The star schema, and why it is one-to-many

Load the tables and relate them, in the **Power Pivot → Manage → Diagram View**
window, by dragging key to key:

```
         Calendar
            |  1
            |
            *                    1
Products --*  Sales  *-------------- Territory
    1          *  *
               |  | 1
     Customers ----  Resellers
```

Every relationship runs **one** from the dimension to **many** in the fact. One
product, many sales lines. One day, many orders. Power Pivot will refuse to
create the relationship if the dimension side is not unique, which is a useful
forcing function: it makes you confront the grain of your dimension table.

Here AdventureWorks has a trap the data dictionary flags. `DimProduct` is a
slowly changing dimension, so the same business product can appear several times
with different `ProductKey` values. That is fine — `ProductKey` is still unique,
and the relationship works. But it means `=DISTINCTCOUNT(Products[ProductKey])`
is not "number of products" and never was. If someone asks how many products the
company sells, the honest measure counts the business key:

```dax
Product Count = DISTINCTCOUNT(Products[ProductAlternateKey])
```

**Filters flow down the "one" side to the "many" side.** Select "Bikes" in a
pivot and the filter lands on `Products`, propagates along the relationship to
`Sales`, and every measure sees only bike rows. It does not flow back up by
default: filtering `Sales` does not filter `Products`. That asymmetry is the
single most important mechanical fact in the chapter, and the cause of most
"why is my number wrong?" moments. When you need the filter to travel the other
way, you ask for it explicitly with `CROSSFILTER`.

### Mark as Date Table, and why it is not optional

Select `Calendar`, then **Design → Mark as Date Table**, and nominate
`FullDateAlternateKey` as the unique date column.

Time intelligence functions — `TOTALYTD`, `SAMEPERIODLASTYEAR`, `DATEADD` — do
not do arithmetic on your dates. They take the set of dates currently in filter
context and *replace* it with a different set: the same months last year, the
days from 1 January to here, the equivalent period shifted by a quarter. To do
that they need a table they can trust to be a complete, gap-free, one-row-per-day
calendar. Marking the table is how you promise that.

The failure mode when you skip it is nasty because it is silent. `TOTALYTD`
returns numbers. The numbers are wrong at year boundaries, or for any month
where the fact table happens to have no transactions on some days — because
without a marked date table the function is working from dates that exist in the
*facts*, and a day with no sales simply is not there. The business-sense check:
your YTD line goes flat across a quiet week instead of staying level, or
December's YTD does not equal the annual total.

Two further requirements. The calendar must cover **complete years** — if your
facts start in December 2010, the calendar still needs 1 January 2010 onward, or
prior-year comparisons for early 2011 have nothing to compare against. And you
must group and slice by the calendar's columns, never the fact's. Put
`Calendar[CalendarYear]` on the pivot, not a year derived from `Sales`. A field
from the fact table bypasses the relationship, so the time intelligence has
nothing to shift.

The fiscal-year point from the data dictionary applies here too. `Calendar`
carries both `CalendarYear` and `FiscalYear`, and the AdventureWorks fiscal year
starts in July. A chart labelled "2013" means different things to the controller
and the sales director. The model lets you offer both; your job is to label which
one the exhibit uses.

### Calculated columns versus measures

The decision students get wrong most often, so here is the rule:

**A calculated column** is computed once per row, at refresh, and stored. It has
*row context*: it knows which row it is on. Use it for an **attribute** — a
property of the row that you want to slice, filter or group by.

```dax
Price Band = SWITCH(TRUE(),
    Products[ListPrice] = 0,     "No price",
    Products[ListPrice] < 50,    "Under 50",
    Products[ListPrice] < 500,   "50 to 500",
    Products[ListPrice] < 2000,  "500 to 2000",
                                 "2000 plus")
```

That belongs in a column, because you want it on the rows of a pivot.

**A measure** is computed on demand, for each cell, from whatever filter context
that cell implies. It is not stored. Use it for a **number you want to see**.

```dax
Revenue = SUM(Sales[SalesAmount])
```

Two consequences that explain most of the confusion:

*A calculated column cannot respond to a slicer.* It was computed at refresh,
before any slicer existed. If you want "this product's share of its category",
that depends on what else is selected and must be a measure.

*A measure cannot go on the rows of a pivot.* It has no row identity. If you
want to group by something, it has to be a column.

And a cost: every calculated column is stored for every row, so a calculated
column on a 20-million-row fact table is expensive, while a measure over the
same table is nearly free until someone looks at it. **When in doubt, write a
measure.** Experienced modellers write almost no calculated columns on fact
tables, and the ones they write are keys.

### Implicit versus explicit measures

Drag `Sales[SalesAmount]` into a pivot's Values area and Excel invents a
calculation for you, labelled "Sum of SalesAmount". That is an *implicit*
measure. It is convenient and you should stop using it.

An *explicit* measure is one you wrote and named:

```dax
Revenue = SUM(Sales[SalesAmount])
```

Why it matters, in the terms of the business question. The implicit measure is
created inside one PivotTable. The next analyst drags the same field, gets
another implicit measure, maybe with `AVERAGE` because they clicked something,
and nothing in the model records that the two differ. The explicit measure lives
in the model, carries its name into every pivot, can be documented, and can be
changed in one place with every exhibit following.

The two-analysts problem is not solved by care. It is solved by there being only
one `[Margin %]`, defined in the model, with its definition visible. Hide the
underlying numeric columns from client tools once your measures exist
(right-click → **Hide from Client Tools**), so the only way to get revenue is to
use `[Revenue]`. That is the governance move, and it takes ten seconds.

### SUM, SUMX, and why margin needs the X

`SUM` adds a column. `SUMX` iterates a table, evaluates an expression per row,
and adds the results.

For margin you can write either, and they agree:

```dax
Margin = SUM(Sales[SalesAmount]) - SUM(Sales[TotalProductCost])
```

```dax
Margin = SUMX(Sales, Sales[SalesAmount] - Sales[TotalProductCost])
```

Now consider a case where they do *not* agree. Suppose cost were not stored on
the fact at all, only `Products[StandardCost]`, and you need cost of goods sold:

```dax
COGS = SUMX(Sales, Sales[OrderQuantity] * RELATED(Products[StandardCost]))
```

There is no column to `SUM`. The product has to be formed row by row — quantity
on this line times the cost of this line's product — and only then added. `SUMX`
supplies the *row context* that makes `Sales[OrderQuantity]` mean "this row's
quantity" and lets `RELATED` reach across the relationship to fetch this row's
product cost.

The general principle: **if the calculation involves a multiplication or a ratio
that must happen before the addition, you need an iterator.** Sum of products is
not the product of sums. This is the same trap as averaging a percentage, which
Chapter 9 treats properly.

### DIVIDE, not the slash

```dax
Margin % = DIVIDE([Margin], [Revenue])
```

Not `[Margin] / [Revenue]`. `DIVIDE` returns blank on a zero denominator instead
of an error, and takes an optional third argument for what to return instead. One
`#DIV/0!` in a pivot cell propagates visually through an entire exhibit and
looks, to a client, like the model is broken.

Note also what is being divided: two measures, each of which aggregates its own
column over the current filter context. This gives the correct weighted
percentage at every level of the pivot automatically. The naive alternative —
averaging a per-row margin percentage — gives a figure in which a $4 inner tube
counts as much as a $3,400 road bike. That is not a rounding difference; it is a
different and wrong number.

### CALCULATE: the whole language in one function

Everything above computes over the filter context handed to it. `CALCULATE`
changes that context. It is the only function that can, and essentially all of
intermediate DAX is `CALCULATE` with different arguments.

```dax
Revenue Internet = CALCULATE([Revenue], Sales[Channel] = "Internet")
```

Read it as: compute `[Revenue]`, but with the filter `Channel = "Internet"`
added to whatever filters are already in play. In a pivot sliced to Bikes and
2013, this gives internet bike revenue in 2013. The measure composes with the
reader's selections rather than overriding them — which is exactly what you want
and what a hard-coded helper column can never do.

The modifiers are where the power is:

**`ALL`** removes filters. The denominator for a share calculation:

```dax
Revenue % of Total = DIVIDE([Revenue], CALCULATE([Revenue], ALL(Products)))
```

In every cell, the numerator respects the row's product filter and the
denominator ignores it, so you get a share that totals to 100%.

**`ALLEXCEPT`** removes all filters from a table except the ones you name. Share
*within* category rather than of the whole:

```dax
Revenue % of Category = DIVIDE([Revenue],
    CALCULATE([Revenue], ALLEXCEPT(Products, Products[EnglishProductCategoryName])))
```

**`FILTER`** builds a table with a condition that can reference measures, which a
simple predicate cannot:

```dax
Revenue Large Orders = CALCULATE([Revenue],
    FILTER(Sales, Sales[SalesAmount] > 1000))
```

**`KEEPFILTERS`** changes how your filter combines with the existing one. By
default `CALCULATE([Revenue], Products[EnglishProductCategoryName] = "Bikes")`
*replaces* any category filter, so in the Clothing row of a pivot it still
reports Bikes. Wrap it in `KEEPFILTERS` and the filters intersect instead, so the
Clothing row reports blank — usually what the reader expects.

That last distinction is subtle and worth sitting with. It is the most common
source of a measure that looks right in the grand total and wrong in every row.

### Time intelligence

With a marked date table, comparisons become declarative:

```dax
Revenue YTD = TOTALYTD([Revenue], Calendar[FullDateAlternateKey])
```

```dax
Revenue LY = CALCULATE([Revenue], SAMEPERIODLASTYEAR(Calendar[FullDateAlternateKey]))
```

```dax
Revenue YoY % = DIVIDE([Revenue] - [Revenue LY], [Revenue LY])
```

`DATEADD` generalises the shift to any offset and grain — a quarter back, three
months forward:

```dax
Revenue PQ = CALCULATE([Revenue], DATEADD(Calendar[FullDateAlternateKey], -1, QUARTER))
```

For a fiscal year ending 30 June, `TOTALYTD` takes a year-end-date argument:

```dax
Revenue FYTD = TOTALYTD([Revenue], Calendar[FullDateAlternateKey], "06-30")
```

Two cautions. These functions need a *contiguous* date selection; applying
`SAMEPERIODLASTYEAR` to a non-contiguous filter such as "January and March" is
undefined behaviour, not an error. And the data dictionary's quirk 9 bites here:
the last partial month of the AdventureWorks window shows an apparent collapse
in revenue, and a YoY measure will faithfully report it as a catastrophic
decline. It is a truncated period. Exclude incomplete periods explicitly, and
say in the exhibit footnote that you did.

### RELATED and RELATEDTABLE

`RELATED` goes from the many side to the one side and returns a single value:
from a sales row to its product's cost, as in the `COGS` measure above. It needs
row context, so it only works inside a calculated column or an iterator.

`RELATEDTABLE` goes the other way, from one to many, and returns a table. A
calculated column on `Products`:

```dax
Units Sold = SUMX(RELATEDTABLE(Sales), Sales[OrderQuantity])
```

Useful, but think before you reach for it. On a product dimension it is cheap;
the same pattern on a large dimension re-scanning a huge fact table at every
refresh is slow, and a measure usually does the job better.

### Many-to-many and CROSSFILTER

AdventureWorks is a clean star, so build the problem yourself to see it.
Suppose a `ProductTags` table where one product can carry several tags and one
tag applies to several products. There is no unique key on either side, so a
direct relationship is impossible.

Two answers. The structural one, and the one I would ship: a **bridge table** at
the grain `ProductKey` + `Tag`, with one-to-many relationships from `Products`
and from `Tags` into it. Filters now travel through the bridge, and the model
documents the many-to-many honestly.

The tactical one: `CROSSFILTER`, which changes a relationship's direction for the
duration of one calculation.

```dax
Products Sold = CALCULATE(DISTINCTCOUNT(Products[ProductAlternateKey]),
    CROSSFILTER(Sales[ProductKey], Products[ProductKey], BOTH))
```

That answers "how many distinct products actually sold in this slice?" by letting
the sales filter propagate *up* to `Products`, against the default direction.

Be sparing. You can set bidirectional filtering permanently on a relationship,
and in a model with several dimensions it creates ambiguous filter paths, which
produce either an error or — worse — a number whose derivation nobody can
explain. Keep relationships single-direction and override per measure where you
need it.

### Measure organisation

Create an empty table called `_Measures` — add a one-cell table to the model and
hide its only column — and put every measure in it. In the PivotTable field list
it appears as a single folder of calculations with no columns to drag by
mistake. Name measures as a business user would say them: `Revenue`, `Margin %`,
`Revenue YoY %`. Not `SumOfSalesAmount_v2`. The names appear in client exhibits;
treat them as deliverables.

Use the **Description** field in the measure dialog to record the definition in
words. It surfaces as a tooltip, and it is the cheapest documentation in the
entire toolchain.

### If you are on a Mac

Power Pivot is Windows-only. Excel for Mac has Power Query (reduced) but no data
model, no relationships view, no DAX. There is no workaround inside Excel for
Mac, and anyone telling you otherwise is describing Windows.

Three honest routes:

1. **Power BI Desktop** — also Windows-only, but it is the same engine and the
   same DAX, and it is free. If you can reach a Windows VM, this is where to
   learn the material. Measures written here are portable to Power Pivot.
2. **Power Query plus formulas on Mac.** Do the joins in Power Query as in
   Chapter 7, producing the wide table, and compute with `SUMIFS` and
   `SUMPRODUCT`. You lose measure reuse and time intelligence; you can still do
   the analysis. The YTD equivalent is a dated range condition:
   `=SUMIFS(Sales[SalesAmount],Sales[OrderDate],">="&DATE(2013,1,1),Sales[OrderDate],"<="&D4)`.
3. **Read this chapter anyway.** Filter context, grain, and definitions living in
   one place are concepts, not features. They transfer to Power BI, to SQL window
   functions and to pandas. The Mac limitation costs you the exercises, not the
   understanding.

Say which platform your deliverable requires, early. A client on Mac who receives
a Power Pivot workbook receives a file that opens, shows stale cached pivots, and
cannot refresh. That is a worse outcome than a plainer workbook.

## Walkthrough

Build the AdventureWorks model and the measure set that every later chapter uses.

**Step 1 — Trim the fact table.** In Power Query, edit `Sales` and delete the
merge-and-expand steps that attached product and territory attributes. Keep
`SalesOrderNumber`, `SalesOrderLineNumber`, `OrderDateKey`, `OrderDate`,
`ProductKey`, `CustomerKey`, `ResellerKey`, `SalesTerritoryKey`, `Channel`,
`OrderQuantity`, `UnitPrice`, `DiscountAmount`, `SalesAmount`,
`TotalProductCost`. Note the file size before and after.

**Step 2 — Load tables to the model, not the sheet.** For `Sales`, `Products`,
`Customers`, `Resellers`, `Territory` and `Calendar`: **Close & Load To → Only
Create Connection**, with **Add this data to the Data Model** ticked. Nothing
lands on a worksheet. This is the architecture: the data lives in the model and
the sheets hold only exhibits.

**Step 3 — Build the relationships.** **Power Pivot → Manage → Diagram View**.
Drag:

- `Sales[ProductKey]` → `Products[ProductKey]`
- `Sales[OrderDateKey]` → `Calendar[DateKey]`
- `Sales[CustomerKey]` → `Customers[CustomerKey]`
- `Sales[ResellerKey]` → `Resellers[ResellerKey]`
- `Sales[SalesTerritoryKey]` → `Territory[SalesTerritoryKey]`

Each should show `1` at the dimension end and `*` at `Sales`. If one is refused
for non-unique keys, your dimension has the wrong grain — fix it in Power Query,
do not force it here.

The customer and reseller relationships are each only meaningful for one channel,
and half the rows carry nulls in each key. That is correct and expected: `Channel`
tells you which key to trust. A pivot by customer attribute will simply show
reseller revenue as blank, which is honest.

**Step 4 — Mark the date table.** Select `Calendar`, **Design → Mark as Date
Table**, choose `FullDateAlternateKey`. Confirm it spans complete years from
before your earliest sale.

**Step 5 — A measures home.** Add a trivial one-cell table to the model, rename
it `_Measures`, hide its column from client tools.

**Step 6 — Write the base measures.** In the Power Pivot window, on `_Measures`:

```dax
Revenue = SUM(Sales[SalesAmount])
```

```dax
Cost = SUM(Sales[TotalProductCost])
```

```dax
Margin = [Revenue] - [Cost]
```

```dax
Margin % = DIVIDE([Margin], [Revenue])
```

```dax
Units = SUM(Sales[OrderQuantity])
```

```dax
Orders = DISTINCTCOUNT(Sales[SalesOrderNumber])
```

```dax
Avg Order Value = DIVIDE([Revenue], [Orders])
```

Format each as you create it — currency with no decimals for money, percentage
with one decimal for `Margin %`. Formatting is a property of the measure, so it
is right in every pivot forever, which is a small thing that saves real time.

Note `Margin` is now `[Revenue] - [Cost]`, measures referencing measures. If the
definition of revenue ever changes — net of freight, say — margin follows
automatically. This is the DRY principle, and it is why you can now delete the
`Margin` column from the `Sales` query. Do that, refresh, and confirm the measure
total equals what the column used to give.

**Step 7 — Channel and share measures.**

```dax
Revenue Internet = CALCULATE([Revenue], KEEPFILTERS(Sales[Channel] = "Internet"))
```

```dax
Revenue Reseller = CALCULATE([Revenue], KEEPFILTERS(Sales[Channel] = "Reseller"))
```

```dax
Revenue % of Total = DIVIDE([Revenue], CALCULATE([Revenue], ALL(Products), ALL(Territory)))
```

```dax
Revenue % of Category = DIVIDE([Revenue],
    CALCULATE([Revenue], ALLEXCEPT(Products, Products[EnglishProductCategoryName])))
```

**Step 8 — Time intelligence.**

```dax
Revenue YTD = TOTALYTD([Revenue], Calendar[FullDateAlternateKey])
```

```dax
Revenue LY = CALCULATE([Revenue], SAMEPERIODLASTYEAR(Calendar[FullDateAlternateKey]))
```

```dax
Revenue YTD LY = CALCULATE([Revenue YTD], SAMEPERIODLASTYEAR(Calendar[FullDateAlternateKey]))
```

```dax
Revenue YoY % = DIVIDE([Revenue] - [Revenue LY], [Revenue LY])
```

```dax
Margin % LY = CALCULATE([Margin %], SAMEPERIODLASTYEAR(Calendar[FullDateAlternateKey]))
```

That last one is worth noticing: you shifted a *ratio* in time by wrapping the
existing ratio measure, with no new arithmetic. In the flat-table world this
needed a new helper column and a new chance to get it wrong.

**Step 9 — Build the pivot that formulas could not.**

**Insert → PivotTable → Use this workbook's Data Model.**

- Rows: `Products[EnglishProductCategoryName]`, then
  `Products[EnglishProductSubcategoryName]` nested.
- Columns: `Calendar[CalendarYear]`.
- Values: `[Revenue]`, `[Margin %]`, `[Revenue YoY %]`, `[Revenue % of Category]`.
- Slicers: `Sales[Channel]`, `Territory[SalesTerritoryGroup]`.

Now interrogate it. Click the Reseller slicer: every measure recomputes,
including the ratios and the prior-year comparisons, at every level of the
hierarchy, correctly weighted. Switch to Europe. Expand Bikes into subcategories
and watch `Revenue % of Category` still sum to 100% within Bikes.

Then confirm the whole point: open a second PivotTable, from the same model, and
put `[Margin %]` on it. It is the same number, because there is one definition.
There is no second margin to disagree with.

**Step 10 — Reconcile to Chapter 7.** The model's numbers must tie to the
pipeline's. Because the fact table no longer lands on a worksheet, you cannot
`SUM` a column any more — you read the model from the grid with `CUBEVALUE`,
which is how a measure is pulled into an ordinary cell:

```excel
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]")
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]","[Sales].[Channel].[Internet]")
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Revenue]","[Sales].[Channel].[Reseller]")
=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Margin]")
```

Put the Chapter 7 control totals in the cells beside them and difference the
pairs:

```excel
=ROUND(B4-C4,2)
=IF(SUMPRODUCT(--(ROUND($B$4:$B$7-$C$4:$C$7,2)<>0))=0,"RECONCILED","CHECK FAILED")
```

That second formula is the whole `Controls` sheet in one cell, and it is the
thing you look at after every refresh. Then check `[Revenue]` against the
Chapter 7 total, `[Revenue Internet]` and
`[Revenue Reseller]` against the `SUMIFS` figures, and verify the YTD measure
for the final complete year equals that year's total. Any gap is a relationship
or a filter problem, not a rounding problem. Find it now.

## Common pitfalls

**No marked date table.** Time intelligence returns plausible wrong numbers.
Mark it, cover complete years, and slice by `Calendar` columns rather than fact
columns.

**Slicing by a fact-table date.** Putting a year derived from `Sales` on the
pivot bypasses the relationship, so `SAMEPERIODLASTYEAR` has nothing to shift and
quietly returns the current period.

**Calculated columns everywhere.** Each one is stored for every row. If it is a
number you want to look at, it is a measure. Columns are for things you group by.

**Implicit measures.** "Sum of SalesAmount" is a per-pivot accident. Write
explicit measures and hide the raw numeric columns from client tools so nobody
can create implicit ones.

**`/` instead of `DIVIDE`.** One zero denominator and `#DIV/0!` spreads across
an exhibit.

**Averaging a ratio.** `AVERAGE` of a margin-percentage column weights a $4 tube
equally with a $3,400 bike. Always `DIVIDE(SUM(...), SUM(...))`.

**Forgetting `KEEPFILTERS`.** `CALCULATE([Revenue], Products[Category]="Bikes")`
replaces the row's own category filter, so every row of the pivot reports Bikes.
The grand total looks fine; every row is wrong.

**`SUM` where you needed `SUMX`.** Any calculation that multiplies before it adds
needs an iterator and row context. Sum of products is not the product of sums.

**Bidirectional relationships by default.** Ambiguous filter paths, unexplainable
numbers. Single-direction relationships, `CROSSFILTER` where needed.

**Reporting the truncated final period.** The last partial month looks like a
collapse and `Revenue YoY %` will report it as one. Filter to complete periods
and footnote it.

**Counting `ProductKey` as products.** Versioned dimension; use
`ProductAlternateKey`.

**Shipping a Power Pivot workbook to a Mac client.** It opens, shows stale
pivots, and cannot refresh. Establish the platform before you choose the
architecture.

## Exercises

1. **Build the model.** Work through steps 1 to 8. Report the file size before
   and after trimming the fact table, and verify on the `Controls` sheet that
   `[Revenue]`, `[Revenue Internet]` and `[Revenue Reseller]` tie exactly to the
   Chapter 7 `SUMIFS` figures.

2. **Prove a measure beats a column.** Note the grand total of the old `Margin`
   column. Delete that column from the `Sales` query, refresh, and confirm
   `[Margin]` gives the identical total. Then redefine `[Revenue]` as
   `=SUM(Sales[SalesAmount]) - SUM(Sales[DiscountAmount])` and state, in one
   sentence, every exhibit that changed and how many formulas you had to edit.

3. **Four share measures, one distinction.** Write `Revenue % of Total`,
   `Revenue % of Category`, and two versions of a Bikes-only measure — one with
   `KEEPFILTERS` and one without. Put all four on a pivot with category on rows
   and explain, from what you observe, exactly what `KEEPFILTERS` did.

4. **Year-over-year, honestly.** Build a pivot of `[Revenue]`, `[Revenue LY]` and
   `[Revenue YoY %]` by `Calendar[EnglishMonthName]` within
   `Calendar[CalendarYear]`. Identify every cell that is misleading because of a
   truncated or absent comparison period, then write a `Revenue YoY % Clean`
   measure that returns blank unless both periods are complete. State your rule
   for "complete".

5. **An iterator you cannot avoid.** Write `COGS at List Cost` as
   `SUMX(Sales, Sales[OrderQuantity] * RELATED(Products[StandardCost]))` and
   compare it with `[Cost]`. Explain why the two differ, which one you would put
   in a client deck, and what the difference tells you about AdventureWorks'
   product versioning.

6. **Fiscal versus calendar.** Build the same category revenue exhibit twice,
   once by `Calendar[CalendarYear]` and once by `Calendar[FiscalYear]`, and write
   `Revenue FYTD` with the `"06-30"` year end. Quantify the largest single
   category difference between the two views and write the one-line footnote you
   would put under the exhibit.

7. **Many-to-many.** Create a small `ProductTags` table (`ProductKey`, `Tag`)
   assigning two or three tags to some products, with products carrying multiple
   tags. Attempt a direct relationship to `Products`, observe what happens, then
   implement the bridge-table solution and report revenue by tag. Explain why
   revenue by tag does not sum to total revenue.

## Solutions

**Exercise 1.** Trimming the lookup columns typically cuts the file by well over
half — the repeated category and territory text was most of the weight, and the
model's compression handles the remaining integer keys almost for free. The
reconciliation must be exact, not close. A mismatch is nearly always one of
three things: a relationship not created, so a filter is not propagating; a
relationship pointing the wrong way; or the `NA` territory row (key 11) from the
data dictionary appearing in one view and not the other. Check all three before
suspecting the data.

**Exercise 2.** The totals must match to the cent, because the measure computes
exactly the subtraction the column stored. The second half is the lesson.
Redefining `[Revenue]` changes `[Margin]`, `[Margin %]`, `[Revenue YTD]`,
`[Revenue LY]`, `[Revenue YoY %]`, `[Revenue Internet]`, `[Revenue Reseller]`,
both share measures and `[Avg Order Value]` — because each of them is written in
terms of `[Revenue]` rather than in terms of a column. The number of formulas you
edited is one. In the flat-table world it would have been one edit per helper
column per sheet, and the near-certainty of missing one. That asymmetry is the
entire argument for measures, and it is also the answer to the chapter's business
question: the two analysts disagreed because each owned a copy of the definition.
Remove the copies and the disagreement has nowhere to live.

**Exercise 3.** `Revenue % of Total` sums to 100% across all categories.
`Revenue % of Category` sums to 100% within each category when subcategories are
on the rows, and shows 100% on every category row itself. The Bikes measure
*without* `KEEPFILTERS` shows the same Bikes figure on every row, including
Clothing and Accessories, because the explicit filter replaced the row's own
category filter. *With* `KEEPFILTERS`, the two filters intersect, so the Bikes
row shows bike revenue and the other rows show blank. The practical rule: if a
measure's value is suspiciously identical down a column, you have replaced a
filter you meant to intersect.

**Exercise 4.** Every month of the first year has a blank `[Revenue LY]` and
therefore a blank or infinite YoY — correct behaviour, but it needs a footnote
rather than a dash in a client exhibit. The dangerous cell is the final partial
month: it reports a collapse of 60% or more that is purely an artefact of the
data window. A defensible `Revenue YoY % Clean` returns blank unless both the
current and prior periods are complete:

```dax
Revenue YoY % Clean =
VAR LastFullDate = CALCULATE(MAX(Sales[OrderDate]), ALL(Sales))
VAR PeriodEnd = MAX(Calendar[FullDateAlternateKey])
RETURN IF(PeriodEnd <= EOMONTH(LastFullDate, -1), [Revenue YoY %])
```

"Complete" here means the period ends on or before the last day of the last
month in which data exists for the whole month. State whichever rule you use, in
the exhibit, because a reader cannot infer it.

**Exercise 5.** The two differ because `Sales[TotalProductCost]` is the cost
captured *at the time of the transaction*, while `Products[StandardCost]` is the
cost on the product version the key points to. For most rows they agree. Where
they diverge, you are seeing the slowly-changing dimension at work. Put
`[Cost]` — the transactional figure — in the client deck, because it is what
actually happened and it reconciles to the ledger. `SUMX` with `RELATED` is the
right tool when the fact genuinely lacks the figure, and the comparison is a
useful data-quality probe: a large gap means product versioning is more active
than the catalogue suggests.

**Exercise 6.** With a 30 June year end, roughly half of each calendar year lands
in a different fiscal year, so category revenue shifts substantially between the
two views — and the apparent *growth rate* can change sign for a category with
strong seasonality, which is the uncomfortable finding. Bikes, with their summer
peak, are most affected. The footnote should name the basis and the year end
explicitly: "Fiscal year ending 30 June; FY2013 covers July 2012 to June 2013."
Never ship an exhibit labelled only "2013".

**Exercise 7.** The direct relationship is refused, because neither side is
unique — `ProductTags` repeats `ProductKey` and `Products` would need to repeat
too. The bridge table at grain `ProductKey` + `Tag` resolves it: one-to-many from
`Products` into the bridge, one-to-many from a distinct `Tags` table into the
bridge, and `CROSSFILTER` or a bidirectional setting on the bridge-to-`Products`
relationship so a tag selection reaches the facts. Revenue by tag does not sum to
total revenue for two reasons, and both must be stated on any such exhibit:
untagged products contribute to the total and to no tag, while multi-tagged
products contribute to several tags and are therefore double-counted in the
column. A tag breakdown is a set of overlapping slices, not a partition. Label
it as such, or a client will add the column up and ask why it exceeds revenue.

## Key takeaways

- **A model beats a wide table.** Narrow tables joined by relationships are
  smaller, easier to extend, and analytically capable of things a flat table
  cannot do at any effort.
- **Filters flow one way, from the one side to the many.** Internalise this and
  most "wrong number" mysteries dissolve.
- **Mark the date table.** Time intelligence does not do date arithmetic; it
  substitutes date sets, and it needs a complete calendar to substitute from.
- **Columns are for grouping, measures are for looking at.** When in doubt, write
  a measure: it costs nothing until someone asks for it.
- **Write explicit measures and hide the raw columns.** This is the governance
  answer to two analysts producing two margin numbers. One definition, in one
  place, inherited by every exhibit.
- **Measures should reference measures.** `[Margin] = [Revenue] - [Cost]` means a
  change to revenue propagates everywhere, in one edit.
- **`DIVIDE`, never `/`. `DIVIDE(SUM, SUM)`, never `AVERAGE` of a ratio.**
- **`CALCULATE` is the language.** `ALL` for shares, `ALLEXCEPT` for within-group
  shares, `FILTER` for measure-dependent conditions, `KEEPFILTERS` so your filter
  intersects rather than replaces.
- **Iterators where multiplication precedes addition.** `SUMX` with `RELATED`
  when the fact table does not carry the number you need.
- **Keep relationships single-direction.** Bridge tables for genuine
  many-to-many; `CROSSFILTER` per measure when you must.
- **Say which calendar and which periods.** Fiscal versus calendar year, and
  complete versus truncated periods, change the story. Footnote both.
- **Check the platform before the architecture.** Power Pivot is Windows-only,
  and a client on a Mac needs a different deliverable or a Power BI route.
