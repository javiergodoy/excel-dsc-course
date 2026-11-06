# Chapter 4 — Joining Tables: Lookups and Relationships

## Learning objectives

By the end of this chapter you will be able to:

- Explain why the join is the central analytical operation in Excel, and what it is doing conceptually when you write a lookup.
- Use `XLOOKUP` fully: exact and approximate match, the `if_not_found` argument, search modes, and reverse search for latest-version lookups.
- Use `INDEX`/`MATCH` and explain the specific situations where it is still the better choice.
- Name `VLOOKUP`'s failure modes precisely enough to justify retiring it from your own work.
- Perform two-dimensional lookups with `INDEX`/`MATCH`/`MATCH` and with nested `XLOOKUP`.
- Join on a composite key using either a concatenated helper column or boolean-multiplication `MATCH`, and state the trade-off between them.
- Use approximate match correctly for banding, and recognise the sort-order requirement that makes it dangerous when violated.
- Distinguish inner-join from left-join thinking, and handle the null `ProductSubcategoryKey` case without silently losing rows.
- Diagnose the duplicate-key explosion that turns a many-to-one join into a many-to-many one, and prove a join did not inflate your totals.
- Decide which of two conflicting paths from a sale to a country is the right one, on commercial rather than technical grounds.
- Recognise the point at which formula-based joining should stop and Power Query or a data model should take over.

## The business question

The AdventureWorks VP of Sales wants a single view: reseller revenue and margin by product category, by reseller business type, and by sales territory group, for the last full year. She wants it by Thursday.

Every number she wants exists. None of it exists in one place. Revenue and cost sit in `FactResellerSales`. Product category sits two tables away from the fact, through `DimProduct` to `DimProductSubcategory` to `DimProductCategory`. Business type sits in `DimReseller`. Territory group sits in `DimSalesTerritory`. The analysis is therefore not a calculation problem at all — the arithmetic is a subtraction and a sum. It is a **join** problem, and the quality of the answer depends entirely on whether the joins are correct.

This is the shape of nearly all real analytical work, and it is where the errors live. A join can fail in two directions and both are quiet. It can **drop** rows, by failing to match and being wrapped in something that hides the failure, which understates your total. Or it can **duplicate** rows, by matching a key that turns out not to be unique on the lookup side, which overstates your total. Neither throws an error. Both produce a number you can put on a slide.

The discipline this chapter teaches, more than any individual function, is this: **after every join, reconcile the total against the unjoined source.** If enriching a table changed its revenue total, the join is wrong. That check takes one formula and it is the difference between an analyst who joins tables and one who can be trusted to.

## Data you will use

From [the AdventureWorks data dictionary](../data/adventureworks-data-dictionary.md):

- **`FactResellerSales`** — the fact table being enriched. `ProductKey`, `ResellerKey`, `SalesTerritoryKey`, `OrderDateKey`, `OrderQuantity`, `SalesAmount`, `TotalProductCost`, `SalesOrderNumber`, `SalesOrderLineNumber`.
- **`DimProduct`** — `ProductKey`, `ProductAlternateKey`, `ProductSubcategoryKey`, `EnglishProductName`, `StandardCost`, `ListPrice`, `StartDate`, `EndDate`.
- **`DimProductSubcategory`** — `ProductSubcategoryKey`, `EnglishProductSubcategoryName`, `ProductCategoryKey`.
- **`DimProductCategory`** — `ProductCategoryKey`, `EnglishProductCategoryName`.
- **`DimReseller`** — `ResellerKey`, `ResellerName`, `BusinessType`, `GeographyKey`, `AnnualSales`.
- **`DimSalesTerritory`** — `SalesTerritoryKey`, `SalesTerritoryRegion`, `SalesTerritoryCountry`, `SalesTerritoryGroup`.
- **`DimGeography`** — `GeographyKey`, `City`, `EnglishCountryRegionName`, `SalesTerritoryKey`.
- **`DimDate`** — `DateKey`, `FullDateAlternateKey`, `CalendarYear`, `MonthNumberOfYear`.
- **`DimCustomer`** — `CustomerKey`, `GeographyKey`, `YearlyIncome`, `CommuteDistance`, used for the banding and conflicting-path discussions.

Load each as an Excel table under the short names used in this book: `Resellers`, `Products`, `Subcategories`, `Categories`, `Territories`, `Geography`, `Calendar`, and `ResellerSales` for the fact table.

## Core concepts

### What a join actually is

A join takes two tables and produces one, matching rows by a shared key. In relational terms you are following a foreign key to a primary key. In Excel, every lookup function is a join executed one row at a time.

That one-row-at-a-time execution is the source of both Excel's flexibility and its danger. A database engine knows the cardinality of a relationship — it knows whether `ProductKey` is unique in `DimProduct` — and will complain if you join in a way that violates it. Excel knows nothing. `XLOOKUP` returns the first match it finds and says nothing about whether there were others. If `DimProduct` has two rows for `ProductKey` 310 because of versioning, you get one of them, chosen by position, and no indication that a choice was made.

So you must supply the cardinality check yourself. Before any join, establish that the key is unique on the lookup side:

```
=ROWS(Products) = ROWS(UNIQUE(Products[ProductKey]))
```

If that returns `TRUE`, the join is many-to-one and safe: many sales rows, one product row each. If it returns `FALSE`, stop. You have a many-to-many situation and you must decide which row is correct before you join, not after.

This check takes five seconds. Make it a habit on every lookup table you touch.

### `XLOOKUP`, properly

`XLOOKUP` has six arguments and most people use three. The other three are where the value is.

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

The basic reseller-name lookup:

```
=XLOOKUP(ResellerSales[@ResellerKey], Resellers[ResellerKey], Resellers[ResellerName])
```

Three things make this immediately better than `VLOOKUP`. The lookup array and return array are named separately, so inserting a column anywhere breaks nothing. The return array can sit to the left of the lookup array, so you never build a helper column just to satisfy argument order. And the default is exact match, which is the behaviour you want ninety percent of the time and the behaviour `VLOOKUP` does not give you by default.

**`if_not_found`** is the fourth argument and the most important one to use deliberately. Supplying it converts an `#N/A` into a value of your choosing:

```
=XLOOKUP(ResellerSales[@ProductKey], Products[ProductKey], Products[EnglishProductName], "** UNMATCHED **")
```

Note the choice of fallback. `"** UNMATCHED **"` is deliberately ugly because it must be visible in a PivotTable and impossible to mistake for real data. Compare that to a fallback of `""` or `0`, which disappears into the exhibit. The fallback is a design decision about how loudly a failure announces itself, and the default should be loud.

When you are counting rather than labelling, leave the `#N/A` in place and count it:

```
=SUMPRODUCT(--ISNA(XLOOKUP(ResellerSales[ProductKey], Products[ProductKey], Products[ProductKey])))
```

That number — unmatched fact rows — belongs in your workbook control block and in a footnote if it is non-zero.

**`match_mode`** takes 0 for exact (the default), -1 for exact-or-next-smaller, 1 for exact-or-next-larger, and 2 for wildcard. The -1 and 1 modes are how you do banding, covered below.

**`search_mode`** takes 1 for first-to-last (default), -1 for last-to-first, and 2 or -2 for binary search on sorted data. The `-1` reverse search has a specific and valuable use in AdventureWorks: `DimProduct` is versioned, so one `ProductAlternateKey` may have several rows with different `StartDate` values. If the table is in date order, reverse search returns the most recent version:

```
=XLOOKUP(Products[@ProductAlternateKey], Products[ProductAlternateKey], Products[ListPrice], , 0, -1)
```

Be careful what you have asserted when you write that. You have assumed the table is sorted such that the last matching row is the current one. That assumption is true of standard AdventureWorks extracts and will not be true of an extract someone re-sorted. A safer formulation that does not depend on sort order filters explicitly on the version validity:

```
=LET(
  k, ResellerSales[@ProductKey],
  rows, FILTER(Products[ListPrice], (Products[ProductKey]=k)*(Products[Status]="Current")),
  IF(ROWS(rows)=0, NA(), INDEX(rows,1))
)
```

The general rule: **prefer a formula that states its condition over one that relies on physical row order.** Row order is not a property of the data; it is a property of the last person who opened the file.

### `INDEX`/`MATCH`, and why it survives

```
=INDEX(Resellers[BusinessType], MATCH(ResellerSales[@ResellerKey], Resellers[ResellerKey], 0))
```

`MATCH` returns a position; `INDEX` returns the value at that position. Decomposing the lookup into those two steps is more verbose than `XLOOKUP` and still worth knowing for three reasons.

First, **one `MATCH`, many `INDEX` calls.** When you are pulling eight columns from the same lookup table on the same key, computing the position once into a helper column and running eight cheap `INDEX` calls against it is substantially faster than eight independent `XLOOKUP` calls, each of which re-scans the lookup array. On a 60,000-row fact table against a 600-row dimension this is the difference between a workbook that recalculates instantly and one that pauses.

Second, **compatibility.** `XLOOKUP` requires a reasonably current Excel. A workbook destined for a client on an older build, or for a locked-down corporate environment, needs `INDEX`/`MATCH`. This constraint is real and it recurs.

Third, **`MATCH` composes.** The boolean-multiplication composite join below is written with `MATCH` because `MATCH` accepts a computed array as its lookup array. That is a capability `XLOOKUP` shares but which reads more naturally in the `MATCH` form.

Use `XLOOKUP` by default. Reach for `INDEX`/`MATCH` when you are pulling many columns on one key, when compatibility demands it, or when the match logic is computed.

### `VLOOKUP`'s failure modes

`VLOOKUP` should be retired from new work. The reasons are specific, not stylistic.

**The fourth argument defaults to approximate match.** `=VLOOKUP(key, range, 3)` with the argument omitted performs an approximate match on unsorted data and returns a plausible, arbitrary, wrong value. Not an error — a wrong value. This is the single most costly default in Excel's history. `XLOOKUP` defaults to exact.

**The column index is positional and brittle.** `=VLOOKUP(key, Products, 7, FALSE)` returns whatever is currently in the seventh column. Insert a column at position three and it silently returns the wrong field, with no error, because the seventh column still exists and still has values in it. An analyst who inherits that workbook has no way to see the breakage.

**It cannot look leftward.** The return column must be to the right of the key, which forces either a helper column or a rearrangement of the source data. Both are work done to satisfy a function rather than to serve the analysis.

**It rescans from the left edge.** The whole range must start at the key column, which on wide tables means dragging in columns you do not need and slowing recalculation.

You will still read `VLOOKUP` in inherited workbooks, so you must be able to audit it. The audit question is always: *is the fourth argument present and `FALSE`, and is the column index still pointing where its author intended?* If you cannot answer both, treat every number downstream as unverified.

### Two-dimensional lookup

When the value you want sits at the intersection of a row and a column — a price grid, a margin matrix, a scenario table — you match on both axes.

```
=INDEX(MarginGrid, MATCH("Mountain Bikes", GridRowLabels, 0), MATCH("Europe", GridColLabels, 0))
```

Or with nested `XLOOKUP`, where the outer call returns a row and the inner call picks from it:

```
=XLOOKUP("Europe", GridColLabels, XLOOKUP("Mountain Bikes", GridRowLabels, MarginGrid))
```

The nested form reads better once you see it: the inner `XLOOKUP` returns a one-row array, and the outer one indexes into that array. Both are exact-match on both axes, which is what you want — an approximate match on a label axis is nonsense.

### Composite keys, two ways

Some joins need more than one column to identify a row. The grain of `FactResellerSales` is order plus line number, so identifying a specific line requires both `SalesOrderNumber` and `SalesOrderLineNumber`.

**The concatenated helper column** builds a single text key on both sides and joins on it:

```
=ResellerSales[@SalesOrderNumber] & "|" & ResellerSales[@SalesOrderLineNumber]
```

It is fast, readable, and has one real hazard: **delimiter collision**. Without a delimiter, `"SO43"` with line `"71"` and `"SO437"` with line `"1"` both produce `"SO43711"` — two different lines, one key. Always use a delimiter, and use one that cannot appear in either field. A pipe is safe for AdventureWorks keys. A hyphen is not, because `ProductAlternateKey` values contain hyphens.

**Boolean multiplication** avoids the helper column entirely:

```
=INDEX(ResellerSales[SalesAmount],
       MATCH(1, (ResellerSales[SalesOrderNumber]=$A2)*(ResellerSales[SalesOrderLineNumber]=$B2), 0))
```

Each comparison produces an array of `TRUE`/`FALSE`; multiplying coerces to 1 and 0; the product is 1 only where both conditions hold. `MATCH(1, ..., 0)` finds that position. It is exact, has no collision risk, and needs no extra columns.

The trade-off is cost. Boolean multiplication evaluates the full array for every row it is used in, so on a large fact table used at scale it is materially slower than a concatenated key. The practical guidance: concatenate with a safe delimiter for bulk joins on big tables, and use boolean multiplication for a handful of targeted lookups, for ad-hoc checks, and wherever you cannot guarantee the delimiter is safe.

### Approximate match and banding

Converting a continuous measure into bands — income into income brackets, order size into size tiers — is an approximate-match lookup against a threshold table.

Build the threshold table with the **lower bound** of each band, sorted ascending:

| LowerBound | Band |
|---|---|
| 0 | Under $30k |
| 30000 | $30k–$60k |
| 60000 | $90k–$120k |
| 120000 | Over $120k |

Then:

```
=XLOOKUP(Customers[@YearlyIncome], Bands[LowerBound], Bands[Band], "out of range", -1)
```

Match mode `-1` means exact-or-next-smaller, which is exactly the banding semantic: an income of 45,000 finds the 30,000 threshold and returns that band.

Two requirements, both of which bite. The thresholds must be **sorted ascending** — violate it and you get wrong bands with no error. And the first threshold must be at or below the minimum possible value, or low values fall off the bottom and return your `if_not_found` value. Supply a visible `if_not_found` so that a gap announces itself, and verify with:

```
=COUNTIFS(Customers[IncomeBand], "out of range")
```

Note the contrast with `CommuteDistance`, which is already banded but as unordered text. That is a mapping problem, not a banding problem: an exact-match lookup against a sort-index table, as in Chapter 3. Knowing which of the two problems you have is the first step.

### Inner join, left join, and the null subcategory

`DimProduct.ProductSubcategoryKey` is **null for manufacturing components that were never sold.** This single fact is a complete lesson in join semantics.

If you join products to subcategories with a plain lookup, those rows return `#N/A`. What you do next determines whether you have performed an inner join or a left join, and the difference is not cosmetic.

**Inner-join thinking** keeps only matched rows. Filter out the `#N/A` values, or let a PivotTable ignore them, and your product count drops. If you were counting the catalogue, you have now undercounted it and will not know by how much.

**Left-join thinking** keeps every row from the left table and labels the unmatched ones:

```
=LET(
  sub, XLOOKUP(Products[@ProductSubcategoryKey], Subcategories[ProductSubcategoryKey],
               Subcategories[ProductCategoryKey], NA()),
  IF(ISNA(sub), "No category (component)",
     XLOOKUP(sub, Categories[ProductCategoryKey], Categories[EnglishProductCategoryName], "** UNMATCHED **"))
)
```

Now every product has a category label, and the unmatched ones carry a label that explains why. Your category breakdown sums to the full catalogue, and a reader can see the residual bucket and ask about it.

Which is correct depends on the question. "What share of revenue comes from Bikes?" is an inner-join question about sold products, and the unsold components are legitimately irrelevant — but you should still state that they were excluded and how many there were. "How many products do we carry?" is a left-join question and excluding them is simply wrong.

**The rule: default to the left join with a visible label.** It is harder to produce a misleading number from a table where every row is present and labelled than from one where the inconvenient rows have quietly vanished. You can always filter later, deliberately, and say so.

### The duplicate-key explosion

This is the error that inflates totals, and it is worth dwelling on because it is the one that survives review.

Suppose you join `ResellerSales` to a lookup table on `ResellerKey`, and that lookup table has two rows for reseller 480 — perhaps because someone built it from a query that itself joined to geography and produced two rows for a reseller with two addresses. A formula-based `XLOOKUP` will not duplicate anything; it returns the first match. But if you build the same join in Power Query with a merge-and-expand, or in a data model relationship, the many-to-many cardinality **multiplies rows**, and every sale to reseller 480 now appears twice. Revenue for that reseller doubles. Company revenue rises by that reseller's full contribution.

Nothing errors. The total is simply larger, and larger in a way that looks like growth.

The defence is the reconciliation, and it is one formula:

```
=SUM(ResellerSalesEnriched[SalesAmount]) - SUM(ResellerSales[SalesAmount])
```

That must be exactly zero. Not approximately zero — exactly. Any non-zero value means the join changed the fact table, and a join must never change the fact table. Pair it with a row-count check:

```
=ROWS(ResellerSalesEnriched) - ROWS(ResellerSales)
```

Also exactly zero. Put both in the control block of every workbook that joins anything. This is the single highest-value habit in this chapter.

### Two paths to a country, and how to choose

A reseller sale can be attributed to a country two ways.

**Through the reseller's geography:** `ResellerKey` → `DimReseller.GeographyKey` → `DimGeography.EnglishCountryRegionName`. This is *where the reseller is located*.

**Through the order's territory:** `SalesTerritoryKey` → `DimSalesTerritory.SalesTerritoryCountry`. This is *which sales territory booked the order*.

They do not always agree, and the disagreement is not a data error. A reseller physically in Canada may be served by the Northwest US territory. Both facts are true; they answer different questions.

Choose on the basis of the decision the client will make:

- **Sales management, quota, compensation, territory performance** — use the territory path. The territory is the organisational unit being measured, and a salesperson's number must match the one in their compensation plan or the exhibit will be rejected on sight.
- **Market sizing, demand geography, logistics, pricing by market** — use the geography path. You want to know where demand physically is, regardless of which internal team booked it.

Then do two things. **Quantify the disagreement**, so you know whether it matters:

```
=SUMPRODUCT(--(XLOOKUP(ResellerSales[SalesTerritoryKey], Territories[SalesTerritoryKey], Territories[SalesTerritoryCountry])
             <> XLOOKUP(XLOOKUP(ResellerSales[ResellerKey], Resellers[ResellerKey], Resellers[GeographyKey]),
                        Geography[GeographyKey], Geography[EnglishCountryRegionName])))
```

And **label the exhibit with the path you chose.** "Revenue by territory country" and "Revenue by reseller location" are different chart titles, and using the precise one pre-empts the question that would otherwise derail your meeting.

The `NA` territory row — key 11 — deserves separate mention. It absorbs unassigned sales and will appear in any unfiltered territory exhibit as a category called `NA`. Exclude it deliberately, with a filter you can point to, and report what it contained. Excluding it accidentally, by a filter someone set and forgot, is how a reconciliation fails six weeks later.

### When to stop using formulas

Formula joins are right for exploration, for small enrichments, and for anything a reader needs to be able to trace by clicking a cell. They stop being right at identifiable points:

- **More than roughly three or four lookup columns on a 50,000-plus row fact table.** Recalculation becomes slow enough to change how you work, which changes how carefully you work.
- **Chained hops.** Product to subcategory to category is two hops. Reseller to geography to territory to group is three. Each hop nests another function and the formula stops being readable, which means it stops being auditable.
- **The join must be repeated on a schedule.** A monthly refresh wants Power Query, where the transformation is recorded once and re-executed, not a column of formulas someone must remember to drag down.
- **The relationship is genuinely many-to-many**, or you need measures that respect filter context across several tables. That is a Power Pivot data model problem; Chapters 7 and 8 cover both.

The transition is a judgment call and the signals are practical: the file is slow, the formula no longer fits on one screen, or you are about to do this again next month. Any one of those is sufficient. Formula joins are not a beginner's tool to be outgrown — they remain the right answer for traceable, one-off enrichment — but insisting on them past these points is a choice to work harder for a worse result.

## Walkthrough

Build the VP's view: reseller revenue and margin by product category, business type and territory group, with a reconciliation proving the join is clean.

### Step 1 — Verify every key before joining

On a `00_Control` sheet, build the cardinality block. This comes first, always.

| Cell | Label | Formula |
|---|---|---|
| `B3` | ProductKey unique in Products | `=ROWS(Products)=ROWS(UNIQUE(Products[ProductKey]))` |
| `B4` | ResellerKey unique in Resellers | `=ROWS(Resellers)=ROWS(UNIQUE(Resellers[ResellerKey]))` |
| `B5` | SalesTerritoryKey unique in Territories | `=ROWS(Territories)=ROWS(UNIQUE(Territories[SalesTerritoryKey]))` |
| `B6` | ProductSubcategoryKey unique in Subcategories | `=ROWS(Subcategories)=ROWS(UNIQUE(Subcategories[ProductSubcategoryKey]))` |
| `B7` | Fact row count | `=ROWS(ResellerSales)` |
| `B8` | Fact revenue | `=SUM(ResellerSales[SalesAmount])` |
| `B9` | Fact cost | `=SUM(ResellerSales[TotalProductCost])` |

All four boolean cells must read `TRUE`. If `B3` reads `FALSE` because of product versioning in your release, you have a decision to make before proceeding: either restrict `Products` to current versions, or accept that the lookup returns the first version and state that assumption. Do not proceed without choosing.

`B8` and `B9` are your reconciliation anchors. Everything downstream is checked against them.

### Step 2 — Build the enriched table

On `03_Enriched`, create `ResellerSalesEnriched`. Start by referencing the fact columns you need, then add one enrichment column at a time. One column, one decision, one visible result.

```
=XLOOKUP(ResellerSales[@ProductKey], Products[ProductKey], Products[EnglishProductName], "** NO PRODUCT **")
=XLOOKUP(ResellerSales[@ResellerKey], Resellers[ResellerKey], Resellers[BusinessType], "** NO RESELLER **")
=XLOOKUP(ResellerSales[@SalesTerritoryKey], Territories[SalesTerritoryKey], Territories[SalesTerritoryGroup], "** NO TERRITORY **")
=XLOOKUP(ResellerSales[@OrderDateKey], Calendar[DateKey], Calendar[CalendarYear], NA())
```

Now the two-hop category lookup. Compute the subcategory key first, in its own column, rather than nesting everything into one unreadable formula:

```
=XLOOKUP(ResellerSales[@ProductKey], Products[ProductKey], Products[ProductSubcategoryKey], NA())
```

Then the category, handling the null case as a left join:

```
=IF(ISNA([@SubcatKey]), "No category (component)",
    XLOOKUP(XLOOKUP([@SubcatKey], Subcategories[ProductSubcategoryKey], Subcategories[ProductCategoryKey], NA()),
            Categories[ProductCategoryKey], Categories[EnglishProductCategoryName], "** UNMATCHED **"))
```

Splitting the hops into separate columns costs two columns of width and buys auditability: when a category looks wrong you can see at a glance whether the subcategory hop or the category hop failed. On a client engagement that difference is an hour.

Finally the derived margin, which does not exist in the source:

```
=[@SalesAmount] - [@TotalProductCost]
```

### Step 3 — Reconcile, before looking at any result

This is the step that must not be skipped, and it must happen **before** you build the PivotTable. Looking at results first biases you toward believing them.

Back on `00_Control`:

| Cell | Label | Formula | Must be |
|---|---|---|---|
| `D7` | Row delta | `=ROWS(ResellerSalesEnriched)-B7` | 0 |
| `D8` | Revenue delta | `=SUM(ResellerSalesEnriched[SalesAmount])-B8` | 0 |
| `D9` | Cost delta | `=SUM(ResellerSalesEnriched[TotalProductCost])-B9` | 0 |
| `D10` | Unmatched products | `=COUNTIFS(ResellerSalesEnriched[ProductName],"** NO PRODUCT **")` | 0 |
| `D11` | Unmatched resellers | `=COUNTIFS(ResellerSalesEnriched[BusinessType],"** NO RESELLER **")` | 0 |
| `D12` | Unmatched territories | `=COUNTIFS(ResellerSalesEnriched[TerritoryGroup],"** NO TERRITORY **")` | 0 |
| `D13` | Products with no category | `=COUNTIFS(ResellerSalesEnriched[Category],"No category (component)")` | report it |
| `D14` | Revenue in no-category rows | `=SUMIFS(ResellerSalesEnriched[SalesAmount],ResellerSalesEnriched[Category],"No category (component)")` | report it |

`D7` through `D9` at zero is your proof that the join neither dropped nor duplicated anything. On a formula-based enrichment they should be zero by construction, since each formula adds a column without adding rows — which is precisely why formula joins are safe for this kind of work. The check still belongs there, because the next person to edit the workbook may convert it to a Power Query merge, where the guarantee disappears.

`D13` and `D14` are findings, not errors. A non-zero `D14` means components with no subcategory were actually sold to resellers, which is commercially interesting: it means the catalogue hierarchy is incomplete for products that generate revenue. That is a sentence for your findings page, and it is the kind of observation that comes only from doing the join carefully.

### Step 4 — Build the view

Only now, with the reconciliation green, insert a PivotTable on `ResellerSalesEnriched`:

- Rows: `Category`, then `BusinessType`
- Columns: `TerritoryGroup`
- Values: `Sum of SalesAmount`, `Sum of Margin`
- Filter: `CalendarYear`, set to the last full year in your release

Add a grand-total check beside it: the PivotTable grand total for the unfiltered view must equal `B8`. If it does not, something in your filter is excluding rows you did not intend to exclude — most commonly the `NA` territory, or a year filter picking up a partial period.

### Step 5 — Attribute the territory choice

You used `SalesTerritoryGroup`, the territory path, which is the right choice for a VP of Sales looking at territory performance. Document it, and size the alternative so you can answer the question if it is asked:

```
=SUMPRODUCT(--(XLOOKUP(ResellerSales[SalesTerritoryKey], Territories[SalesTerritoryKey], Territories[SalesTerritoryCountry])
             <> XLOOKUP(XLOOKUP(ResellerSales[ResellerKey], Resellers[ResellerKey], Resellers[GeographyKey]),
                        Geography[GeographyKey], Geography[EnglishCountryRegionName])))
```

Title the exhibit **"Reseller revenue and margin by territory group"**, not "by region". The precision in the title is doing work: it tells the reader which of two defensible definitions you used, before they have to ask.

## Common pitfalls

**Joining before checking key uniqueness.** The five-second `ROWS(UNIQUE(...))` check prevents the most expensive error in this chapter. Omitting it means you do not know whether your join is many-to-one or many-to-many, which means you do not know whether your total is right.

**Not reconciling after the join.** A join must never change the fact table's row count or revenue total. If you have not checked, you do not know that it did not. This is the single most important habit here.

**`IFERROR` or a blank `if_not_found` on a lookup.** It converts unmatched rows into invisible ones. Your total understates and the exhibit looks fine. Use a loud fallback and count the fallbacks.

**Assuming `#N/A` means the data is broken.** Often it means the relationship is genuinely optional — the null `ProductSubcategoryKey` is correct data, not corrupt data. The right response is a label, not a repair.

**Concatenating a composite key without a delimiter.** `"SO43"&"71"` and `"SO437"&"1"` collide. Always use a delimiter that cannot appear in either field, and never use a hyphen on AdventureWorks product keys.

**Approximate match on unsorted thresholds.** Returns wrong bands with no error at all. Sort ascending, include a bottom threshold at or below the data minimum, and check the out-of-range count.

**Omitting `VLOOKUP`'s fourth argument.** Defaults to approximate match on unsorted data and returns a confident wrong answer. If you are auditing inherited work, this is the first thing to look for.

**Nesting three hops into one formula.** It works and nobody can audit it, including you in two weeks. Split the hops into columns. Width is cheap; opacity is not.

**Forgetting the `NA` territory row.** It will appear in unfiltered territory exhibits and it will quietly change your totals when someone sets a filter. Exclude it explicitly and report what it held.

**Mixing the two country paths within one analysis.** Using reseller geography on one exhibit and territory on the next, in the same deck, guarantees two slides that do not tie. Pick one path per analysis, state it, and stay with it.

## Exercises

1. **Cardinality audit.** For `DimProduct`, `DimReseller`, `DimSalesTerritory`, `DimProductSubcategory` and `DimCustomer`, test whether the surrogate key is unique using `ROWS(UNIQUE(...))`. Then test whether `DimProduct.ProductAlternateKey` is unique. Explain the difference between the two results for `DimProduct` and state what it means for any join keyed on the business code rather than the surrogate key.

2. **Deliberate explosion.** Build a small lookup table that deliberately contains two rows for one `ResellerKey`. Join it to `FactResellerSales` twice: once with `XLOOKUP`, once with a Power Query merge-and-expand. Compare the resulting row counts and revenue totals against the raw fact table. Explain in three sentences why the two methods behave differently and which behaviour is more dangerous in practice.

3. **Left join versus inner join, sized.** Count the products in `DimProduct` with a null `ProductSubcategoryKey`. Then compute the reseller revenue attributable to those products. Produce two versions of a revenue-by-category exhibit, one inner-join and one left-join, and state the percentage of revenue that differs between them. Write the footnote you would put under the inner-join version.

4. **Composite key join.** Using `SalesOrderNumber` and `SalesOrderLineNumber`, look up `SalesAmount` for ten specific order lines, first with a concatenated key and then with boolean-multiplication `MATCH`. Confirm the two agree. Then construct a case where an unsafe delimiter makes the concatenated version return the wrong row, and state the general rule your case demonstrates.

5. **Income banding.** Build a five-band income threshold table and band every customer with approximate-match `XLOOKUP`. Verify that no customer falls out of range. Then deliberately re-sort the threshold table descending, observe what happens to the bands, and explain why Excel reports no error.

6. **Latest version lookup.** For a product with multiple versions in `DimProduct`, retrieve the current `ListPrice` three ways: reverse-search `XLOOKUP`, a `FILTER` on `Status="Current"`, and a `FILTER` on `ISBLANK(EndDate)`. Compare the results and state which you would ship to a client and why, with reference to what each one assumes about the source data.

7. **The two country paths.** Build reseller revenue by country twice, once through reseller geography and once through order territory. Produce a reconciliation showing the rows and revenue where the two disagree, ranked by revenue. Then write the two-sentence note you would give the VP of Sales explaining which figure appears in her pack and why.

## Solutions

**Exercise 1.** `ProductKey`, `ResellerKey`, `SalesTerritoryKey`, `ProductSubcategoryKey` and `CustomerKey` are all unique in their dimensions — they are surrogate keys and uniqueness is their purpose. `ProductAlternateKey` is **not** unique, because AdventureWorks versions products: one business code can have several rows with different `StartDate` and `EndDate` values and different costs and prices. The consequence is direct and commercially important. Any join keyed on `ProductAlternateKey` is many-to-many and will either silently pick an arbitrary version — giving you last year's cost on this year's sale — or multiply rows if executed as a merge. Facts carry `ProductKey`, the versioned surrogate, precisely so that each transaction is matched to the product *as it was priced at the time*. That is the whole point of the surrogate key, and understanding it is the difference between a margin number that reflects the economics of each sale and one that applies a single current cost retroactively across four years.

**Exercise 2.** `XLOOKUP` returns the first matching row and ignores the second, so row count and revenue total are unchanged — but the business type you pulled is one of two possible values, chosen by physical position, with no indication a choice was made. The Power Query merge expands to two rows per matching sale, so row count rises and revenue for that reseller doubles. The merge behaviour is more dangerous because it changes the number, and a number that is too large looks like a finding rather than a defect; someone may well present it as growth. The `XLOOKUP` behaviour is dangerous in a subtler way — it is silently arbitrary — but at least your totals still tie, so the reconciliation in Step 3 will not catch it. The lesson is that the reconciliation catches duplication, not arbitrary selection, and only the cardinality check catches both. You need them both.

**Exercise 3.** Expect on the order of a few hundred `DimProduct` rows with a null `ProductSubcategoryKey`, and expect the reseller revenue attributable to them to be a small share of the total — low single digits at most, since most are manufacturing components never sold externally. The percentage is the point of the exercise, not the formula. A difference of 0.2% between the inner-join and left-join exhibits is immaterial and you proceed with either, noting the exclusion. A difference of 8% is material and the inner-join version is actively misleading, because a reader summing the visible categories gets a total that does not match company revenue and will notice. The footnote should read along the lines of: *"Excludes 312 order lines (0.4% of reseller revenue) for products with no assigned subcategory; these are manufacturing components outside the standard catalogue hierarchy."* Specific count, specific share, specific reason. That footnote is what lets a reader decide whether to care, which is their job rather than yours.

**Exercise 4.** Both methods agree on all ten lines when the delimiter is safe. The failure case: use an empty delimiter, then find an order number ending in a digit where a shorter order number plus a longer line number produces the same concatenation — `"SO43"` with line `71` and `"SO437"` with line `1` both yield `SO43711`. The lookup returns whichever appears first, which is a different sales line with a different amount, and no error is raised. The general rule the case demonstrates: **a concatenated key is only as safe as the delimiter, and a delimiter is only safe if it cannot occur in any component field.** This is why boolean multiplication is the right choice when you cannot verify the component fields, and why a pipe is a better habit than a hyphen on a dataset where hyphens appear inside product codes.

**Exercise 5.** With thresholds sorted ascending, every customer bands correctly and the out-of-range count is zero provided your lowest threshold is zero or below. Re-sorted descending, `XLOOKUP` with match mode `-1` returns wrong bands for most rows — and reports nothing, because approximate match does not validate sort order; it assumes it. That assumption is documented in the function reference and invisible in the spreadsheet. This is the structural reason approximate match is the most under-appreciated hazard in Excel: the function cannot distinguish "this data is sorted and the answer is correct" from "this data is unsorted and the answer is arbitrary". Your defence is procedural rather than technical. Keep threshold tables in a dedicated, clearly-labelled range, sorted ascending, and add a check: `=SUMPRODUCT(--(Bands[LowerBound]>=INDEX(Bands[LowerBound],1)))=ROWS(Bands)` is a crude ascending test you can refine, but any check beats none.

**Exercise 6.** All three return the same price for a well-formed extract. They differ in what they assume. Reverse-search assumes the physical row order puts the current version last — an assumption about the file, not the data, and one that any re-sort destroys silently. `Status="Current"` assumes the status column is populated and maintained, which it is in standard AdventureWorks but is a column that goes stale in real source systems. `ISBLANK(EndDate)` assumes the versioning convention that an open-ended validity window marks the current row, which is the most widely used convention in dimensional modelling and the most self-documenting. Ship the `EndDate` version, because it depends on a modelling convention rather than on file state or on a maintained flag, and because a reader can see from the formula what "current" means. Then state the assumption next to it. The deeper point: three formulas returning the same number today can diverge tomorrow, and choosing between them is a judgment about which assumption is most likely to survive, not a question of which is shortest.

**Exercise 7.** The disagreement will be concentrated in a modest number of resellers whose physical location sits outside the territory serving them — Canada served from US territories is the classic pattern in this dataset. Rank the disagreeing rows by revenue, because the share of revenue affected is the only thing that determines whether this matters. The note to the VP: *"Revenue is attributed to the sales territory that booked the order, matching your territory P&L and quota definitions. Attributing instead to reseller physical location would move approximately $X million between North America and Canada; we can supply that view separately if you need a market-demand rather than territory-performance cut."* Two sentences, and they do three things: state the definition used, justify it by reference to the decision the client is making, and size the alternative so nobody needs to guess. Offering the alternative rather than defending the choice is what keeps the meeting on the recommendation instead of on the methodology.

## Key takeaways

- The join is the central analytical operation. Most analysis is not a calculation problem; it is a join problem wearing a calculation's clothes.
- Check key uniqueness with `ROWS(UNIQUE(...))` before every join. Five seconds, and it tells you whether your join is safe.
- Reconcile after every join: row count delta and revenue delta against the unjoined source, both exactly zero. A join must never change the fact table.
- Use `XLOOKUP` by default, with a loud `if_not_found` value and a count of how many rows received it.
- Keep `INDEX`/`MATCH` for many columns on one key, for compatibility constraints, and for computed match logic.
- Retire `VLOOKUP` from new work; its approximate-match default and positional column index are failure modes, not preferences.
- Composite keys need a delimiter that cannot appear in the component fields, or boolean-multiplication `MATCH` with no delimiter at all.
- Approximate match requires ascending thresholds and silently returns wrong answers when they are not. Check the out-of-range count.
- Default to left-join thinking with a visible label for unmatched rows, so your breakdowns sum to the total and the residual is inspectable.
- When two defensible paths exist from a fact to a dimension, choose on the basis of the client's decision, label the exhibit precisely, and size the alternative.
- Move from formulas to Power Query or a data model when the file slows, the formula stops fitting on a screen, or you will have to do this again next month.
