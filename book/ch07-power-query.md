# Chapter 7 — Power Query: Repeatable ETL in Excel

The three chapters before this one taught you to analyse data that was already
sitting in your workbook, in good shape, with the columns you needed. That is
not the state in which data arrives. It arrives as six CSV files with headers on
row 4, or a database you can only reach read-only, or a monthly export whose
column order changes when somebody in finance edits the report template.

The gap between "data exists somewhere" and "data is in a table I can analyse"
is where most analyst hours go, and where most analyst errors are made. This
chapter is about closing that gap with a recorded pipeline instead of an
afternoon of manual editing.

## Learning objectives

By the end of this chapter you will be able to:

- Explain why manual data preparation is a defect, not a task, and what
  replaces it.
- Read the Applied Steps list as what it is: auditable code with a visual
  front end.
- Connect to CSV files, whole folders of files, databases and tables inside the
  current workbook.
- Apply the transformation set that covers most real work: promote headers,
  change type with locale, remove and reorder columns, split and merge, and
  unpivot.
- Use Group By to aggregate in the query rather than in the sheet.
- Choose the right Merge Queries join kind, including the anti joins, and map
  each one to the join concepts from Chapter 4.
- Append `FactInternetSales` and `FactResellerSales` into one `Sales` table
  despite their differing columns, tagged with a `Channel` column.
- Parameterise a file path so the pipeline survives being moved or handed over.
- Set staging queries to connection-only and understand the resulting
  dependency graph.
- Write small amounts of M for the steps the user interface cannot express.
- Diagnose the formula-firewall error and set privacy levels deliberately.

## The business question

The question for this chapter is not an analytical question. It is the question
a partner asks on day three of an engagement, and it is the one that decides
whether the engagement is pleasant or miserable:

> **The client will send us a refreshed sales extract every Monday for the next
> eleven weeks. How do we build the analysis once, and then spend four minutes
> rather than four hours on each refresh?**

Underneath it sits a second question, which is really about risk:

> **When the numbers in next Monday's deck differ from last Monday's, can we say
> precisely why — and prove it was the data that changed and not us?**

Those two questions have the same answer. If every transformation between the
raw extract and the analysed table is recorded as a step, then the refresh is
automatic and the diff is explainable. If any transformation was a human
deleting rows, neither is true.

Note what this chapter adds to the AdventureWorks case specifically. Everything
so far has been about internet sales, because `FactInternetSales` is a single
convenient table. But as the data dictionary warns, internet sales are the
smaller channel. Any statement about "AdventureWorks revenue" that uses only
internet sales is wrong about the business by a wide margin. Unioning the two
fact tables is the first job in this chapter because it is the job that makes
every later chapter honest.

## Data you will use

From [the AdventureWorks data dictionary](../data/adventureworks-data-dictionary.md):

| Table | Role here | Columns that matter |
|---|---|---|
| `FactInternetSales` | Fact, direct channel | `SalesOrderNumber`, `SalesOrderLineNumber`, `OrderDateKey`, `ProductKey`, `CustomerKey`, `SalesTerritoryKey`, `OrderQuantity`, `SalesAmount`, `TotalProductCost`, `DiscountAmount` |
| `FactResellerSales` | Fact, wholesale channel | the same, plus `ResellerKey` and `EmployeeKey`, and no `CustomerKey` |
| `DimProduct` | Product attributes | `ProductKey`, `EnglishProductName`, `ProductSubcategoryKey` |
| `DimProductSubcategory` | Hierarchy hop | `ProductSubcategoryKey`, `EnglishProductSubcategoryName`, `ProductCategoryKey` |
| `DimProductCategory` | Hierarchy top | `ProductCategoryKey`, `EnglishProductCategoryName` |
| `DimSalesTerritory` | Geography rollup | `SalesTerritoryKey`, `SalesTerritoryRegion`, `SalesTerritoryCountry`, `SalesTerritoryGroup` |
| `DimDate` | Calendar | `DateKey`, `FullDateAlternateKey`, `CalendarYear`, `MonthNumberOfYear`, `EnglishMonthName`, `FiscalYear` |

Assume the four dimension tables and two fact tables are available as CSV files
in a single folder. If you have them as worksheets instead, every step in the
walkthrough is identical except the first; use **Data → Get Data → From Other
Sources → From Table/Range** in place of the CSV connector.

Two derived fields will be created inside the query, not in the sheet:

- `Channel` — the text `"Internet"` or `"Reseller"`, stamped on each fact row
  before the append.
- `Margin` — `SalesAmount − TotalProductCost`, computed once, in the pipeline.

The second one deserves a word. In Chapter 5 you computed margin with a worksheet
formula such as `=SUM([@SalesAmount]-[@TotalProductCost])`. That works, but it puts
the definition of margin in a column that any analyst can overwrite in any
workbook, which is how two analysts end up with two margin numbers. Computing it
in the query moves the definition upstream, where it is defined once and
inherited by everything downstream. Chapter 8 moves it upstream again, into a
measure. This is a theme: **push every definition as far up the pipeline as it
will go.**

## Core concepts

### ETL is a program, and Applied Steps is its source code

Power Query is an ETL tool: **E**xtract from a source, **T**ransform the shape,
**L**oad to a destination. The thing that makes it valuable is not that it can
do transformations — Excel could always do transformations. It is that the
transformations are *recorded, named, ordered and replayable*.

When you click "Remove Columns", Power Query does not remove columns. It appends
a line to a program. The Applied Steps pane on the right of the editor is that
program, read top to bottom:

```
Source
Promoted Headers
Changed Type with Locale
Removed Other Columns
Inserted Margin
Added Channel
Appended Reseller Sales
Merged Product Category
Expanded Category
```

You can click any step and see the table as it was at that point. You can insert
a step in the middle, delete one, or reorder two. You can rename a step so the
next person understands your intent — and you should, because `Removed Columns1`
tells a reviewer nothing while `Keep only reporting columns` tells them
everything.

This is why I introduce Power Query as *code* rather than as a feature. The
discipline it enforces is the same discipline Chapter 1 argued for in workbook
layout: separate the raw input from the logic applied to it, and make the logic
visible. The Applied Steps list is a diff-able, reviewable, re-runnable
transformation log. A sequence of manual edits in a worksheet is none of those
things, however careful the person making them.

### The two-layer rule: never transform the source in place

Open a query. Look at the first step, `Source`. Everything after that step is
yours; the `Source` step itself should do nothing but connect. Resist the strong
temptation to "just fix" something in the source file. The moment you edit the
raw extract, the pipeline is no longer reproducible from what the client sent
you, and when the Monday refresh arrives you will have to remember what you
fixed. You will not remember.

### Connecting: four sources you will actually use

**From Text/CSV.** The default for extracts. Power Query infers the delimiter and
samples the first 200 rows to guess types. The sampling is the risk; see the
type section below.

**From Folder.** The one people discover too late. Point it at a directory,
and it returns a table of files — one row per file, with a `Content` column
holding the binary. Transform one file, then let Power Query apply that
transformation to every file and stack the results. This is how you turn
"eleven weekly extracts" into one query that grows by itself. If the client
sends `Sales_2024_W01.csv` through `Sales_2024_W11.csv`, you build the
transformation once against the first file and the pipeline absorbs the rest.

**From Database.** SQL Server, Azure SQL, PostgreSQL, Oracle and the rest. The
important behaviour here is **query folding**: where it can, Power Query
translates your steps into SQL and makes the server do the work, so a `Group By`
over fifty million rows never travels down the wire. Folding breaks the moment
you use a step the source cannot express, and everything after that point runs
locally. Right-click a step and choose **View Native Query**; if it is greyed
out, folding has stopped there. Put your filters and groupings early, while
folding still works.

**From Table/Range.** Reads a table in the current workbook. Useful for the
assumption tables of Chapter 13 — a parameter sheet the user edits, consumed by
a query. Be careful of circularity: a query that reads a table which is itself
the output of a query is a refresh-order problem waiting to happen.

### Types, locale, and the comma that cost a client a week

The second step in almost every query should be a deliberate type assignment.
Power Query will guess for you, from a sample, and the guess is usually right
and occasionally catastrophic.

Two specific failure modes:

**Sampling.** Types are inferred from the first rows. A column that holds
integers for 10,000 rows and then a text value in row 10,001 will be typed as
whole number and the text row will become an error. You will not see it unless
you look.

**Locale.** `1.234,56` is one thousand two hundred thirty-four point five six in
Germany and a malformed number in the United States. `03/04/2013` is 3 April or
4 March depending on who exported it. Power Query converts using *your* machine's
locale by default, which means the same query produces different numbers on the
Madrid analyst's laptop and the London analyst's laptop.

The fix is **Transform → Data Type → Using Locale**, which pins the
interpretation to a stated culture. Do this explicitly on every date and every
decimal that came from outside your own country. It generates M like this:

```m
= Table.TransformColumnTypes(
    #"Promoted Headers",
    {{"OrderDate", type date}},
    "en-GB"
  )
```

The `"en-GB"` is the whole point. It is now a property of the query, not of the
laptop. This is the single most under-used feature in Power Query and the cause
of a disproportionate share of international reporting errors.

### Unpivot: the highest-value transformation in the product

Most business data arrives in *cross-tab* form, because that is how humans read
it: months across the top, products down the side, numbers in the middle.

| Product | Jan | Feb | Mar |
|---|---|---|---|
| Road-150 | 42,000 | 38,500 | 51,200 |
| Mountain-200 | 31,400 | 29,900 | 36,100 |

Every analytical tool, including PivotTables, wants the opposite — *tidy* form,
one row per observation:

| Product | Month | Amount |
|---|---|---|
| Road-150 | Jan | 42,000 |
| Road-150 | Feb | 38,500 |
| Mountain-200 | Jan | 31,400 |

Select the month columns, **Transform → Unpivot Columns**, and you have it. Three
clicks replace an hour of copy-paste-transpose, and unlike the hour of
copy-paste, it survives the client adding an April column — if you use
**Unpivot Other Columns** instead, which means "keep `Product` as a key and melt
everything else", and therefore adapts to new columns automatically.

Learn to recognise the symptom: if you find yourself writing a formula that
references a column *by its month name*, you needed to unpivot and did not.

### Group By: aggregate upstream

**Transform → Group By** collapses rows. Group `Sales` by `ProductKey` and
produce `Sum of SalesAmount`, and sixty thousand rows become a few hundred.

Use it when the analysis genuinely does not need line-level detail, because a
smaller table refreshes faster, and because against a database it may fold into
a `GROUP BY` that the server executes. Do *not* use it to pre-aggregate
speculatively, because an aggregation you did upstream is detail you cannot get
back downstream. The rule I use: **keep the grain you might be asked about, and
no finer.** For AdventureWorks, that is line level, because the obvious
follow-up question — "what is the discount on that line?" — needs it.

### Merge Queries, and the join kinds

Merge is a join. Chapter 4 explained joins in worksheet terms, with `XLOOKUP`
and `INDEX`/`MATCH`; Power Query exposes the full relational set:

| Join kind | Keeps | Use it when |
|---|---|---|
| **Left outer** | All left rows, matched right rows | The default and the right answer 80% of the time. You are enriching a fact table with attributes. |
| **Inner** | Only matched rows | You deliberately want to discard unmatched facts — rarely what you want on a fact table, because it silently deletes revenue. |
| **Right outer** | All right rows | Almost never. Swap the tables and use left outer instead; it is easier to reason about. |
| **Full outer** | Everything from both | Reconciliation. Two extracts of the same population that should agree. |
| **Left anti** | Left rows with *no* match | **The diagnostic join.** Orphan keys, referential-integrity failures. |
| **Right anti** | Right rows with no match | Unused dimension members — products that never sold. |

The anti joins are the ones that separate a careful analyst from a trusting one.
Before you accept a left-outer merge, run the same merge as a **left anti** join
and look at the row count. If it is zero, every fact row found its dimension and
you may proceed. If it is 1,400, you have a data quality finding to report —
and, more urgently, 1,400 rows whose category will show as blank in every exhibit
you build. The data dictionary flags exactly this case: products with a null
`ProductSubcategoryKey`.

After merging, Power Query gives you a column of nested tables. Click the expand
icon in its header and pick the columns you want, and **untick "Use original
column name as prefix"** unless you actually want `DimProductCategory.
EnglishProductCategoryName` as a header.

### Append Queries, and the union of unlike tables

Append stacks rows. Merge adds columns; append adds rows. Here they are needed
for the central structural fact of AdventureWorks: the business has two sales
channels in two tables with overlapping but not identical columns.

Power Query's append is **by column name, not position**. Matching names line up;
names present in one table and not the other produce `null` in the rows that
lack them. This is forgiving behaviour, and it is also how a typo turns into two
half-populated columns called `SalesAmount` and `Sales Amount`. Normalise the
column names in each staging query *before* appending, which is both why staging
queries exist and why the walkthrough below removes unneeded columns early.

`CustomerKey` and `ResellerKey` are a genuine, not a cosmetic, difference. Do
not force them into one column; a customer and a reseller are different entities
with different dimension tables. Keep both columns, accept the nulls, and let
`Channel` tell you which one is meaningful on a given row.

### Parameters: the pipeline that survives being handed over

A query whose `Source` step contains a literal path works on your laptop:

```m
= Csv.Document(File.Contents("C:\Users\jgodoy\Desktop\aw\FactInternetSales.csv"))
```

Send that workbook to a colleague and every query fails. The fix is a parameter.
**Home → Manage Parameters → New Parameter**, name it `DataFolder`, type Text,
current value the folder path. Then edit the source step to use it:

```m
= Csv.Document(File.Contents(DataFolder & "\FactInternetSales.csv"))
```

Now one value, editable without opening the query editor, relocates the entire
pipeline. For a client engagement, parameterise the data folder, the reporting
period and any threshold the client might argue about. For maximum
friendliness, hold the parameter values in a small worksheet table and read it
with `Excel.CurrentWorkbook`, so the user never opens the editor at all.

### Staging queries and connection-only loading

A real pipeline is a graph, not a line. In the walkthrough you will build nine
queries, and only one of them should land in the workbook as a table. The others
are *staging* queries — intermediate results consumed by later queries.

Set them to **Close & Load To → Only Create Connection**. They then exist in the
Queries pane, participate in refresh, and occupy no sheet. Loading all nine as
tables is the most common beginner mistake in Power Query: it bloats the file,
refreshes everything twice, and presents the reader with eight tables they must
not touch.

Use **View → Query Dependencies** to see the graph. If it is a tangle rather
than a tree, the pipeline will be hard to refresh in a predictable order.

### Privacy levels and the formula firewall

Sooner or later you will see:

> *Formula.Firewall: Query 'X' (step 'Y') references other queries or steps, so
> it may not directly access a data source. Please rebuild this data
> combination.*

This is Power Query's privacy system. It classifies each source as Private,
Organizational or Public, and refuses to let data from a private source be sent
to another source — the scenario it is protecting you against is your query
quietly posting client data to a web service as part of a lookup.

Two legitimate responses. The correct one: restructure so that the step
combining data does not also connect to a source — put each source in its own
staging query, and combine the staging queries. The pragmatic one, when you are
certain all sources are equally trusted: **Options → Current Workbook → Privacy
→ Ignore Privacy Levels**. That setting is per workbook, so say so in the
documentation sheet, because it is a security decision and the next person
inherits it.

### A little M

The UI generates M. You will occasionally need to write some, and reading it is
non-negotiable if you want to review someone else's query. Three patterns worth
knowing:

A custom column with a conditional, from **Add Column → Custom Column**:

```m
= if [SalesAmount] = 0 then null else [DiscountAmount] / [SalesAmount]
```

Note `if/then/else` rather than a function call, `null` rather than blank, and
case sensitivity — `If` is an error. Guarding the divide-by-zero is not
optional; one zero-revenue line produces an error value that propagates.

A step that references a previous step by name:

```m
= Table.SelectRows(#"Added Channel", each [SalesAmount] > 0)
```

`each` introduces a function of the current row, `[Column]` reads a field, and
`#"Step Name"` is how a name with spaces is quoted.

And the single most useful line in M, for turning an integer `DateKey` into a
real date — a job the UI cannot do in one move, because `20130715` is a perfectly
valid whole number and Power Query has no reason to suspect it is a date:

```m
= Date.FromText(Text.From([OrderDateKey]), [Format = "yyyyMMdd"])
```

## Walkthrough

Build the canonical AdventureWorks sales pipeline: two fact tables in, one
analysable table out, with channel, margin, category and territory attached, and
a refresh that reconciles.

**Step 1 — Parameterise the folder.**

**Data → Get Data → Launch Power Query Editor → Home → Manage Parameters → New**.
Name `DataFolder`, Type `Text`, Current Value the folder holding your CSVs.
Every source step below uses it.

**Step 2 — Stage the internet sales fact.**

**New Source → Text/CSV**, select `FactInternetSales.csv`, then **Transform
Data** — not **Load**. In the editor:

1. Rename the query `stgInternetSales`. Naming matters; you will have nine of
   these.
2. **Use First Row as Headers** if it did not happen automatically.
3. **Remove Other Columns**, keeping only: `SalesOrderNumber`,
   `SalesOrderLineNumber`, `OrderDateKey`, `ProductKey`, `CustomerKey`,
   `SalesTerritoryKey`, `PromotionKey`, `OrderQuantity`, `UnitPrice`,
   `DiscountAmount`, `SalesAmount`, `TotalProductCost`. Rename the step
   `Keep reporting columns`.
4. Set types explicitly: keys to Whole Number, the money columns to Fixed
   Decimal Number (currency — not Decimal Number, which is floating point and
   will give you 0.30000000000000004 in a reconciliation), `SalesOrderNumber` to
   Text. If this extract came from outside your locale, use **Data Type → Using
   Locale** on the money columns.
5. **Add Column → Custom Column**, name `Channel`, formula `= "Internet"`.
6. Add a `ResellerKey` column of nulls and an `EmployeeKey` column of nulls, so
   both staging tables have the same shape before appending. Custom column,
   formula `= null`, then set the type to Whole Number.
7. **Close & Load To → Only Create Connection**.

**Step 3 — Stage the reseller sales fact.**

Same sequence against `FactResellerSales.csv`, query named `stgResellerSales`,
keeping the identical column list but with `ResellerKey` and `EmployeeKey` as
real columns and a `CustomerKey` column of nulls. `Channel` is `= "Reseller"`.

Check the two queries have identical column names, spelled identically, in the
same order. The append is name-based so order does not strictly matter, but
matching order makes the review trivial, and a review that is trivial is a
review that happens.

**Step 4 — Append into one fact table.**

**Home → Append Queries → Append Queries as New**, pick `stgInternetSales` and
`stgResellerSales`, name the result `Sales`.

Sanity check now, before going further. Add a temporary Group By on `Channel`
counting rows, confirm both channels are present with plausible counts, then
delete that step. Roughly 60,000 rows each is what you expect; one channel at
zero rows means an append picked up the wrong query.

**Step 5 — Derive the date and the margin.**

With `Sales` selected:

1. **Add Column → Custom Column**, name `OrderDate`:

   ```m
   = Date.FromText(Text.From([OrderDateKey]), [Format = "yyyyMMdd"])
   ```

   Set the type to Date. You now have a real date you can group by month
   without `DateKey` arithmetic nonsense — the trap the data dictionary lists as
   quirk 7.

2. **Add Column → Custom Column**, name `Margin`:

   ```m
   = [SalesAmount] - [TotalProductCost]
   ```

   Set the type to Fixed Decimal Number. Margin is now defined exactly once for
   the whole workbook.

3. Add `MarginPct` with the guard discussed above:

   ```m
   = if [SalesAmount] = 0 then null else ([SalesAmount] - [TotalProductCost]) / [SalesAmount]
   ```

   A `null` here is honest: a line with no revenue has no margin percentage, and
   averaging it later would be wrong anyway. Chapter 9 explains why you should
   almost never average a percentage.

**Step 6 — Stage the product hierarchy.**

Three staging queries — `stgProduct`, `stgSubcategory`, `stgCategory` — each
reduced to the columns in the table above, all connection-only. Then:

1. In `stgProduct`: **Merge Queries**, `ProductSubcategoryKey` to
   `stgSubcategory`, **Left Outer**. Expand
   `EnglishProductSubcategoryName` and `ProductCategoryKey`, prefix unticked.
2. Still in `stgProduct`: merge the result to `stgCategory` on
   `ProductCategoryKey`, Left Outer, expand `EnglishProductCategoryName`.
3. Rename the query `Products`. It now carries product name, subcategory and
   category in one flat table, which is what the fact table wants.

Before you move on, run the diagnostic. Duplicate `stgProduct`, change the first
merge to **Left Anti**, and look at the count. Those are the products with a null
`ProductSubcategoryKey` — the never-sold manufacturing components. Expect a few
hundred. Now you know, and you can state in the deck that they are excluded
because they have no sales rather than discovering a blank row in a client
meeting. Delete the duplicate.

**Step 7 — Attach product and territory to Sales.**

In `Sales`:

1. **Merge Queries**, `ProductKey` to `Products`, Left Outer. Expand
   `EnglishProductName`, `EnglishProductSubcategoryName`,
   `EnglishProductCategoryName`.
2. **Merge Queries**, `SalesTerritoryKey` to `stgTerritory`, Left Outer. Expand
   `SalesTerritoryRegion`, `SalesTerritoryCountry`, `SalesTerritoryGroup`.

Rename each merge step to say what it attached. `Merged Queries2` is noise;
`Attach territory rollup` is documentation.

**Step 8 — Load exactly one table.**

**Close & Load To → Table → New Worksheet**. One table named `Sales`, around
120,000 rows, with channel, margin, category and territory on every row. Every
other query is connection-only.

**Step 9 — Reconcile.**

This step is not optional and it is the reason the chapter exists. In a scratch
area, compute the control totals:

```excel
=SUM(Sales[SalesAmount])
=SUMIFS(Sales[SalesAmount],Sales[Channel],"Internet")
=SUMIFS(Sales[SalesAmount],Sales[Channel],"Reseller")
=SUM(Sales[Margin])
=COUNTA(Sales[SalesOrderNumber])
=SUMPRODUCT(--(Sales[EnglishProductCategoryName]=""))
```

Add two more, because a reconciliation that only checks the grand total hides
offsetting errors:

```excel
=COUNTA(UNIQUE(Sales[SalesOrderNumber]))
=MAX(Sales[OrderDate])-MIN(Sales[OrderDate])
```

The distinct order count tells you whether the append duplicated rows, and the
date span tells you whether a type conversion quietly turned some dates into
errors — the span should be about 1,100 days, not 40,000.

Then open the two raw CSVs and total `SalesAmount` in each, independently of the
pipeline. The internet total must equal the internet `SUMIFS`, the reseller
total must equal the reseller `SUMIFS`, and the two must add to the grand total
to the cent. The last formula must return zero — a non-zero count means fact
rows failed to find a category and your category exhibits are understated.

Record those five numbers on a `Controls` sheet with today's date.

**Step 10 — Refresh and re-reconcile.**

Simulate next Monday. Append a few hundred rows to one of the source CSVs — or
just re-save it unchanged — and hit **Data → Refresh All**. The pipeline re-runs
every step in order. Check the control totals again: they should have moved by
exactly the amount you added, and the blank-category count should still be zero.

If that works, you have what the business question asked for. The analysis was
built once; the refresh took four minutes, most of which was waiting; and when
the number changes you can say which source row changed it.

## Common pitfalls

**Loading every query as a table.** Nine sheets of intermediate data, a 40 MB
file, and a reader who does not know which table is the real one. Connection-only
for everything but the output.

**Trusting inferred types.** Power Query sampled 200 rows and guessed. Set types
explicitly, as a named step, and use the locale variant for anything foreign.

**Using Inner Join on a fact table.** It silently deletes revenue. Left outer,
then an anti join to count what did not match, then a decision about those rows
that you write down.

**Not running the anti join at all.** This is the big one. A left outer merge
never fails; it just produces nulls. Your category chart quietly excludes
€2 million of sales and the total at the bottom of the chart no longer matches
the total in the summary table. Anti-join first, always.

**Renaming columns in the source file.** The query references column names. Edit
the CSV header and the step breaks with "column not found". Rename inside the
query instead, and keep the raw file pristine.

**Appending before normalising names.** `Sales Amount` and `SalesAmount` append
into two columns, each half null, and the total halves. Normalise in staging.

**Filtering with the dropdown checkboxes.** Click a column dropdown, untick a
few values, and Power Query generates a step listing the *values* you kept:
`each ([Color] <> "Red" and [Color] <> "Blue")`. Next month's extract has a new
colour and it comes through unfiltered. Write the condition you mean, as a
custom filter.

**Ignoring query folding.** Thirty seconds becomes thirty minutes on a large
database because an early step broke folding. View Native Query, and move
filters and groupings before the step that breaks it.

**Turning off privacy levels reflexively.** It usually works and it is a
security decision. Prefer restructuring; if you do disable it, document it.

**Letting `Changed Type` auto-steps accumulate.** Power Query inserts a type step
after many operations, so a long query can carry six of them, each re-casting
every column. Delete the redundant ones; a query with 40 steps where 12 are
`Changed Type` is slow and unreadable.

**Hard-coded paths.** Parameterise on day one, not on the day you hand over.

## Exercises

1. **Build the pipeline.** Work through the walkthrough end to end and produce a
   single loaded `Sales` table with `Channel`, `OrderDate`, `Margin`,
   `MarginPct`, product category and territory group. Record the five control
   totals from Step 9 on a `Controls` sheet.

2. **Prove the referential integrity.** For each of `ProductKey`,
   `SalesTerritoryKey` and `CustomerKey` (internet rows only), create a left
   anti merge from `Sales` to the relevant dimension and report the orphan count.
   Write one sentence per key stating whether the join is clean and, if not, how
   much revenue sits on the orphan rows. Then delete the three diagnostic
   queries — or keep them connection-only and refer to them in the workbook
   documentation as integrity checks.

3. **Unpivot a cross-tab.** Use your `Sales` table to build a PivotTable of
   `SalesAmount` by `EnglishProductCategoryName` (rows) and `CalendarYear`
   (columns). Copy it as values into a new sheet — now you have a cross-tab, the
   format a client would email you. Load *that* into Power Query and unpivot it
   back to tidy form with columns `Category`, `Year`, `SalesAmount`. Then
   confirm the unpivoted total equals the cross-tab total. Use **Unpivot Other
   Columns** and explain in one line why that choice matters when a 2015 column
   appears.

4. **Parameterise and relocate.** Convert every hard-coded path to use
   `DataFolder`. Move the CSV folder somewhere else, change the parameter value
   only, and refresh. Then go further: put the parameter value in a one-cell
   worksheet table and have the query read it with `Excel.CurrentWorkbook`, so a
   non-technical user can repoint the pipeline without opening the editor.

5. **Group upstream, and measure the cost.** Create a new query from `Sales`
   that groups by `CalendarYear`, `EnglishProductCategoryName` and `Channel`,
   producing `Sum of SalesAmount`, `Sum of Margin` and `Count of Rows`. Load it
   as a second table. Compare its row count and the file size before and after.
   Then state one analytical question the grouped table can answer and one it
   cannot, and say which table you would ship to a client.

6. **Break it deliberately.** Rename `SalesAmount` to `Revenue` in one of the
   source CSVs and refresh. Read the error, note which step failed and why, then
   fix it in the query — not in the file — in a way that would survive the same
   change happening again.

## Solutions

**Exercise 1.** The control totals are the deliverable, not the table. Your
internet and reseller `SUMIFS` must each tie to an independent total taken from
the raw CSV, and the blank-category count must be zero. The number that surprises
most students is the channel split: reseller sales are roughly three times
internet sales by revenue. If your previous chapters' analysis used only
`FactInternetSales`, it was describing about a quarter of the business. That is
the finding to carry forward, and it is a reminder that the structural question
— *which rows am I even looking at?* — outranks every technique in this book.

**Exercise 2.** `SalesTerritoryKey` and `CustomerKey` should join cleanly, with
zero orphans; AdventureWorks is a well-behaved warehouse at the fact level.
`ProductKey` also joins cleanly from the fact side. The nulls appear in the other
direction — the right-anti join from `stgProduct` to the facts shows products
that never sold, and the left-anti merge inside `stgProduct` shows products with
no subcategory. Both are findings about the catalogue, not about the sales, and
the correct treatment is to exclude them from sales exhibits and mention the
exclusion. Note also the two-path problem the data dictionary raises: territory
reached through `DimCustomer.GeographyKey` does not always agree with the order's
`SalesTerritoryKey`. Pick one, say which, and be consistent. I use the order's
territory, because the question "where did we sell it?" is about the sale.

**Exercise 3.** The unpivoted total must match the cross-tab total exactly; if it
does not, you unpivoted a row of column subtotals along with the data, which is
the classic error and the reason you should never paste a PivotTable with totals
enabled into a pipeline. **Unpivot Other Columns** matters because it is defined
negatively — "melt everything that is not `Category`" — so a new 2015 column is
absorbed automatically. **Unpivot Columns** names the columns to melt, so the new
column would sit there untouched, contributing nothing to the total and producing
a quiet understatement. Negative definitions adapt; positive ones go stale.

**Exercise 4.** The relocation should take one edit. The `Excel.CurrentWorkbook`
version is the better answer for a client deliverable, and the step is:

```m
= Excel.CurrentWorkbook(){[Name="Params"]}[Content]{0}[Value]
```

That reads the first row's `Value` from a worksheet table named `Params`. Watch
for the refresh-order subtlety: the parameter table must not itself be query
output, or you have a dependency cycle.

**Exercise 5.** Grouping by year, category and channel collapses roughly 120,000
rows to under a hundred, and the file shrinks accordingly. It can answer "what is
margin percentage by category by year by channel?" instantly. It cannot answer
"what is the average discount on orders above $5,000?", because order-level and
line-level detail is gone, and no amount of downstream cleverness recovers it.
Ship the detail table to a client who will ask follow-up questions, and the
grouped table only as a performance optimisation behind a dashboard whose
questions are fixed.

**Exercise 6.** The `Keep reporting columns` step fails, because it names
`SalesAmount` and that column no longer exists. The naive fix is to add a rename
step before it — which works until the column name changes again. The robust fix
is to rename on the way in, as the second step of the staging query, and to make
the rename tolerant:

```m
= Table.RenameColumns(Source, {{"Revenue", "SalesAmount"}}, MissingField.Ignore)
```

`MissingField.Ignore` means the step succeeds whether the source arrives with the
old name or the new one. That is the general lesson: when you know a source is
unstable, build the instability into the pipeline rather than fixing it by hand
each Monday. A pipeline that tolerates the client's inconsistency is worth more
than a pipeline that is elegant.

## Key takeaways

- **Manual data preparation is a defect.** If a transformation is not recorded as
  a step, it is not reproducible, and an analysis you cannot reproduce is an
  analysis you cannot defend.
- **Applied Steps is source code.** Name the steps. A reviewer should be able to
  read the list and know what the query does without clicking anything.
- **Never edit the source.** Everything after `Source` is yours; the raw extract
  stays exactly as the client sent it.
- **Set types deliberately, with locale.** Inferred types come from a 200-row
  sample, and date and decimal formats are not universal. Pin them.
- **Unpivot is the highest-leverage transformation you will learn.** Cross-tabs
  are for reading; tidy data is for analysing. Prefer **Unpivot Other Columns**
  so new columns are absorbed.
- **Left outer merge, then anti merge.** A left outer join never errors — it just
  produces nulls that silently shrink your totals. Count the non-matches before
  you believe the join.
- **Append is by column name.** Normalise names in staging queries, and keep
  `CustomerKey` and `ResellerKey` separate with a `Channel` column to say which
  one applies.
- **One loaded table, many connection-only queries.** The query dependency graph
  should be a tree with a single leaf in the workbook.
- **Parameterise paths on day one.** The pipeline that only runs on your laptop
  is not a deliverable.
- **Reconcile after every refresh.** Control totals on a dated `Controls` sheet
  are the cheapest insurance in analytics, and the only way to tell "the data
  changed" from "I broke something".
- **Push definitions upstream.** Margin computed in the query beats margin
  computed in a column, because there is one of it. Chapter 8 pushes it one step
  further still, into a DAX measure, and that is where it finally stops moving.
