# Chapter 2 — Reading the AdventureWorks Data Model

In Chapter 1 you summed a column and got roughly $29 million. That number is
wrong as a statement about AdventureWorks, and nothing in the formula was
incorrect. It is wrong because it answers a question about one sales channel
while appearing to answer a question about a company.

This chapter is about reading a data model well enough that you know what your
own totals mean. It is the least glamorous skill in analytics and the one that
most reliably distinguishes people whose numbers survive scrutiny.

## Learning objectives

By the end of this chapter you will be able to:

- Distinguish fact tables from dimension tables and say what each is for.
- State the **grain** of any table, and explain why it is the first question to
  ask and the most common thing nobody asks.
- Tell a surrogate key from a business key, and recognise when a slowly
  changing dimension has turned "count the products" into a trap.
- Convert the integer `DateKey` into a real Excel date and explain why the
  conversion is not optional.
- Identify fan-out and double counting before they appear in an exhibit.
- Profile a table — rows, distinct counts, date range, null rates — in under
  ten minutes, and refuse to analyse one you have not profiled.
- Produce a defensible company-wide revenue figure for AdventureWorks, split by
  channel, and justify every inclusion and exclusion.

## The business question

The sales director's deck says company revenue in calendar 2013 was about $43
million. Your Chapter 1 workbook says about $16 million for the same year.
Neither of you has made an arithmetic error.

> **What is AdventureWorks' actual revenue, how many orders generated it, and
> why do two competent analysts get different answers?**

In practice this is the first real test on an engagement. The client is not
asking for analysis; they are asking whether your numbers can be trusted. Get
the reconciliation right and everything afterwards is easier. Get it wrong once
and every subsequent exhibit gets re-checked by someone else.

## Data you will use

Both fact tables, plus the date dimension and the product dimension. Full column
lists in
[../data/adventureworks-data-dictionary.md](../data/adventureworks-data-dictionary.md).

- **`FactInternetSales`** — loaded in Chapter 1 as `Sales`. Columns:
  `SalesOrderNumber`, `SalesOrderLineNumber`, `ProductKey`, `CustomerKey`,
  `OrderDateKey`, `OrderQuantity`, `SalesAmount`, `TotalProductCost`,
  `SalesTerritoryKey`.
- **`FactResellerSales`** — load it now as `ResellerSales`. Columns:
  `SalesOrderNumber`, `SalesOrderLineNumber`, `ProductKey`, `ResellerKey`,
  `EmployeeKey`, `OrderDateKey`, `OrderQuantity`, `SalesAmount`,
  `TotalProductCost`, `SalesTerritoryKey`.
- **`DimDate`** — load as `Calendar`. Columns: `DateKey`,
  `FullDateAlternateKey`, `CalendarYear`, `MonthNumberOfYear`,
  `EnglishMonthName`, `FiscalYear`.
- **`DimProduct`** — loaded in Chapter 1 as `Products`. Columns used here:
  `ProductKey`, `ProductAlternateKey`, `EnglishProductName`,
  `ProductSubcategoryKey`, `StartDate`, `EndDate`, `Status`.

Assumption carried forward from Chapter 1: revenue is `SalesAmount`; margin is
`SalesAmount − TotalProductCost`. New assumption for this chapter: **"company
revenue" means internet plus reseller `SalesAmount`**, with no elimination
entries, because the dataset contains no intercompany sales.

## Core concepts

### Facts and dimensions

A star schema separates two kinds of thing.

**Facts are events that happened and have a size.** A sale of one Road-150 on
15 July 2013 for $3,578.27. Fact tables are long, narrow, numeric and
append-only. Everything you sum, count or average lives in a fact table.

**Dimensions are the nouns you group by.** The product, the customer, the day,
the territory. Dimension tables are short, wide, mostly textual, and they
change slowly. Everything that appears on an axis or in a row label lives in a
dimension.

The practical consequence: **if you want to slice revenue by something, that
something must exist as a dimension attribute, or you must create it.** Half of
all analytical work is realising that the question requires an attribute nobody
built — "revenue by customer age band" needs an age band, and AdventureWorks
gives you `BirthDate`.

### Grain is the first question

The grain of a table is what exactly one row represents. Say it out loud as a
sentence starting "one row is…".

- `FactInternetSales`: *one row is one product line on one sales order.*
- `FactResellerSales`: *one row is one product line on one reseller order.*
- `DimDate`: *one row is one calendar day.*
- `DimProduct`: *one row is one version of one product.*
- `DimCustomer`: *one row is one customer.*

Almost every serious error in a data-driven deck traces back to a grain
misunderstanding. If one row is an order *line* and you count rows, you have
counted lines and called them orders. In `FactInternetSales` that inflates the
count by roughly a fifth, because a meaningful share of orders contain more than
one product. In `FactResellerSales` the distortion is far larger, because
wholesale orders routinely carry many lines.

Grain also determines what you may safely average. The average of
`Sales[UnitPrice]` is the average price per order line, which weights a $4
water bottle the same as a $3,578 road bike. It is not the average price a
customer paid, and it is not the average selling price of the catalogue. It is
a number with no business meaning at all, and it appears in decks constantly.

### Surrogate keys, business keys and slowly changing dimensions

`DimProduct` has two identifiers and they do different jobs.

`ProductKey` is a **surrogate key**: a meaningless integer invented by the
warehouse, used to join to facts. `ProductAlternateKey` is the **business key**:
`BK-R93R-62`, the code printed on the box, stable over time and recognised by
the client.

They diverge because `DimProduct` is a **slowly changing dimension**. When a
product's `StandardCost` or `ListPrice` changes, the warehouse does not
overwrite the old row. It closes it off with an `EndDate` and inserts a new row
with a new `ProductKey` and the same `ProductAlternateKey`. This is what
preserves history: a sale in 2012 joins to the 2012 version of the product, with
the 2012 cost.

The trap is immediate:

```
=COUNTA(Products[ProductKey])              → number of product VERSIONS
=COUNTA(UNIQUE(Products[ProductAlternateKey]))   → number of PRODUCTS
```

In the 2019 release the first figure is around 600 and the second is materially
lower. If a slide says "AdventureWorks sells over 600 products", it is wrong,
and the error came from counting the wrong key. When the client asks how many
products they sell, they mean distinct `ProductAlternateKey` among sellable
items — which also means excluding rows where `ProductSubcategoryKey` is null,
since those are manufacturing components that were never sold.

Three rules follow. Join on the surrogate key. Count and report on the business
key. Never present a count of dimension rows as a count of business entities.

### The integer DateKey trap

`OrderDateKey` is `20130715`. It looks like a date, sorts like a date, and is an
integer.

```
=20130715-20130630    → 85
```

The real gap is 15 days. Nothing errors. You get a plausible-looking number that
is wrong by a factor of five or six, and the magnitude of the error varies
depending on whether the period crosses a month boundary — so it will not even
be consistently wrong, which makes it harder to spot.

Two correct conversions:

```
=DATE(LEFT(Sales[@OrderDateKey],4), MID(Sales[@OrderDateKey],5,2), RIGHT(Sales[@OrderDateKey],2))
=XLOOKUP(Sales[@OrderDateKey], Calendar[DateKey], Calendar[FullDateAlternateKey])
```

The first parses the integer. The second looks it up in the date dimension. Both
work; the second is better practice, because it inherits every other attribute
`DimDate` already carries — fiscal year, month name, week number, quarter — and
guarantees consistency with anyone else joining through the calendar. It also
fails loudly if a `DateKey` is not in the calendar, which is information you
want.

A derived date column belongs in the **transform layer** (`20_TRN_Sales`), as a
calculated column on the Table. Add it once, and every downstream formula uses a
real date instead of integer arithmetic. This is precisely the debt Chapter 1
flagged and the reason the transform layer exists.

Note also that `CalendarYear` and `FiscalYear` disagree in `DimDate` —
AdventureWorks' fiscal year starts in July. "2013" means two different periods
to the sales director and the controller, and stating which you used is part of
the deliverable, not a footnote.

### Fan-out and double counting

Joining a fact table to a dimension is safe: many facts, one dimension row, each
fact row matches at most one dimension row. The row count of the result equals
the row count of the fact.

Joining a fact table to another fact table is almost never safe. If you join
`FactInternetSales` to `FactResellerSales` on `ProductKey`, every internet line
for a given product matches every reseller line for that product. One hundred
internet lines and fifty reseller lines for the same bike become five thousand
rows, and summing `SalesAmount` over them produces a number with no meaning and
no obvious signature of error. It is simply very large.

This is **fan-out**, and the general rule is: *you may only join tables on a
column that is unique in at least one of them.* `ProductKey` is unique in
`DimProduct` and not in either fact. So fact-to-dimension is fine, and
fact-to-fact on `ProductKey` is not.

The correct way to combine two facts of the same grain is not a join but a
**union**: stack them, with a `Channel` column marking the source. That is
Chapter 7's work in Power Query. For now you aggregate each separately and add
the totals, which is arithmetically identical and much harder to get wrong.

Fan-out has a subtler cousin. Joining `FactInternetSales` to `DimCustomer` and
then to `DimGeography` is safe, but there are now two routes from a sale to a
country: via the customer's `GeographyKey`, and via the order's
`SalesTerritoryKey`. They do not always agree, because a customer can order
while assigned to a different territory. Pick one route, state it, and use it
consistently. An exhibit that mixes both is internally inconsistent and the
inconsistency is invisible.

### The real primary key of a sales fact

Neither `SalesOrderNumber` nor `SalesOrderLineNumber` is unique on its own. The
primary key is the pair. This gives you the two counts the business actually
cares about:

```
=ROWS(Sales)                                        → order LINES
=COUNTA(UNIQUE(Sales[SalesOrderNumber]))            → ORDERS
=ROWS(Sales)/COUNTA(UNIQUE(Sales[SalesOrderNumber]))    → lines per order
```

Lines per order is a diagnostic worth computing on any transactional table. For
internet sales it sits a little above 1 — consumers buy one bike. For reseller
sales it is much higher. That single ratio tells you the two channels are
structurally different businesses, before you have analysed anything, and it is
the kind of observation that earns credibility in a first meeting.

### Profile before you analyse

Ten minutes per table, every time, written down on the transform sheet:

1. **Row count.** `=ROWS(Sales)`
2. **Distinct count of the business key.** `=COUNTA(UNIQUE(Sales[SalesOrderNumber]))`
3. **Date range.** `=MIN(Sales[OrderDate])` and `=MAX(Sales[OrderDate])`
4. **Null and blank rates** on every column you intend to use.
5. **Min, max and sum** of each numeric column you intend to use.
6. **Distinct values** of each categorical column you intend to group by.

Step 6 catches more problems than the rest combined. A column you expected to
hold five values holding seven — because two are mis-spelled variants — would
otherwise show up as two extra rows in a pivot that nobody notices.

Step 3 catches the truncation quirk. The AdventureWorks sales window ends
mid-month, so the final period looks like a collapse. It is a partial month, not
a trend, and an analyst who plots monthly revenue without profiling the date
range will present a cliff that does not exist.

## Walkthrough

Continue in the Part I workbook from Chapter 1.

**Step 1 — Load the remaining tables.** Add sheets `12_RAW_ResellerSales` and
`13_RAW_Calendar`. Import `FactResellerSales` as a Table named `ResellerSales`,
and `DimDate` as a Table named `Calendar`.

**Step 2 — Profile both facts.** On `20_TRN_Sales`, build a profile block with
one column per fact table:

```
=ROWS(Sales)                                      =ROWS(ResellerSales)
=COUNTA(UNIQUE(Sales[SalesOrderNumber]))          =COUNTA(UNIQUE(ResellerSales[SalesOrderNumber]))
=SUM(Sales[OrderQuantity])                        =SUM(ResellerSales[OrderQuantity])
=SUM(Sales[SalesAmount])                          =SUM(ResellerSales[SalesAmount])
=MIN(Sales[OrderDateKey])                         =MIN(ResellerSales[OrderDateKey])
=MAX(Sales[OrderDateKey])                         =MAX(ResellerSales[OrderDateKey])
```

In the 2019 release both facts hold roughly 60,000 lines, but reseller revenue
is several times internet revenue. Write that observation down on the sheet in
words. The row counts being similar while the revenue differs by a multiple is
the single most important structural fact about this business, and it is already
visible.

**Step 3 — Add real dates.** On the `Sales` Table in the transform layer, add a
calculated column `OrderDate_Real`:

```
=XLOOKUP([@OrderDateKey], Calendar[DateKey], Calendar[FullDateAlternateKey], "NOT IN CALENDAR")
```

Add the same to `ResellerSales`. Then verify no orphans:

```
=COUNTIF(Sales[OrderDate_Real],"NOT IN CALENDAR")
```

It should be zero. If it is not, you have fact rows whose dates are outside the
date dimension, which is a referential-integrity failure you must report rather
than quietly filter away.

Add two more calculated columns to each, pulled from the calendar rather than
parsed, so that everyone downstream groups the same way:

```
=XLOOKUP([@OrderDateKey], Calendar[DateKey], Calendar[CalendarYear])
=XLOOKUP([@OrderDateKey], Calendar[DateKey], Calendar[EnglishMonthName])
```

**Step 4 — Orders versus lines.** Still on `20_TRN_Sales`:

```
=ROWS(Sales)/COUNTA(UNIQUE(Sales[SalesOrderNumber]))
=ROWS(ResellerSales)/COUNTA(UNIQUE(ResellerSales[SalesOrderNumber]))
```

Internet comes out a little above 1. Reseller comes out substantially higher.
Record both, and record the implication: average order value is not comparable
across the two channels without saying whether you mean an order or a line.

**Step 5 — Demonstrate fan-out deliberately.** On a scratch sheet, pick one
`ProductKey` that appears in both facts. Count its lines in each, multiply the
two counts, and compare that product to what a naive product-level join would
produce. Seeing the multiplication once, with your own numbers, is what makes
the rule stick. Delete the scratch sheet afterwards; it is not part of the
deliverable.

**Step 6 — Build the defensible company figure.** On `30_ANL_Revenue`, build a
channel table for the year held in `Analysis_Year`:

```
=SUMIFS(Sales[SalesAmount], Sales[CalendarYear], Analysis_Year)
=SUMIFS(ResellerSales[SalesAmount], ResellerSales[CalendarYear], Analysis_Year)
```

Add them for company revenue. Do the same for `TotalProductCost` and derive
margin and margin percentage per channel. Now you have a figure with a stated
scope — both channels, calendar year, `SalesAmount`, no tax or freight — and the
two-channel split that explains the discrepancy in *The business question*.

**Step 7 — Reconcile explicitly.** Add a short reconciliation block in words:
your internet-only figure, the reseller figure, the total, and one line naming
the difference between calendar and fiscal year. This block is what you show
when someone's deck disagrees with you. It converts an argument about competence
into a conversation about definitions, which is a conversation you can win.

**Step 8 — Record the product count correctly.** On `30_ANL_Revenue`:

```
=COUNTA(Products[ProductKey])
=COUNTA(UNIQUE(Products[ProductAlternateKey]))
=COUNTA(UNIQUE(FILTER(Products[ProductAlternateKey], Products[ProductSubcategoryKey]<>"")))
```

Three numbers, descending. Label them *versions*, *distinct products*, and
*sellable products*. Only the third answers "how many products do we sell".

## Common pitfalls

**Counting rows and calling them orders.** The most common single error on this
dataset. Always report lines and orders separately, labelled.

**Averaging a per-line column.** `=AVERAGE(Sales[UnitPrice])` is an average over
lines, not over orders, customers or products. If the client wants average
selling price per bike, the denominator is units, not rows:
`=SUM(Sales[SalesAmount])/SUM(Sales[OrderQuantity])`.

**Joining two fact tables.** Fan-out. Union same-grain facts; never join them on
a shared dimension key.

**Arithmetic on `DateKey`.** Silently wrong, never errors. Convert first.

**Counting `ProductKey` as products.** Counts versions. Use
`ProductAlternateKey`, and exclude null `ProductSubcategoryKey`.

**Forgetting the reseller channel.** Internet-only revenue is roughly a quarter
of the company. Any "company" exhibit built on `FactInternetSales` alone is
wrong by a factor of about four, and it looks entirely plausible.

**Letting the `NA` territory in.** `SalesTerritoryKey` 11 absorbs unassigned
sales and will appear as a row in any unfiltered territory exhibit. Exclude it
deliberately and say so, or include it labelled "unassigned". Do not let it
appear as if it were a region.

**Mixing calendar and fiscal year in one exhibit.** Both exist in `DimDate`,
both are called a year, and they cover different months. Pick one, label it.

## Exercises

1. Write the grain of each of these tables as a sentence beginning "one row
   is…": `FactResellerSales`, `DimReseller`, `DimEmployee`, `DimPromotion`,
   `DimGeography`. For each, name one metric that is safe to average at that
   grain and one that is not.

2. Build the full profile block from *Core concepts* for `ResellerSales`,
   including null rates for `CarrierTrackingNumber` and `CustomerPONumber` and
   distinct values of `SalesTerritoryKey`. Identify which columns are too sparse
   to analyse, and state the threshold you used to decide.

3. Convert `OrderDateKey` on `ResellerSales` to a real date by both routes — the
   `DATE/LEFT/MID/RIGHT` parse and the `XLOOKUP` against `Calendar`. Confirm the
   two agree for every row with a single formula that returns one number, and
   explain why you would still prefer the lookup.

4. Compute company revenue for the same twelve months under both the calendar
   and the fiscal definition of a year. Quantify the difference, and write the
   one sentence you would put under an exhibit to make the basis unambiguous.

5. A colleague's slide reads: *"AdventureWorks processed 121,000 orders across
   613 products, averaging $460 per order."* Using what you have built,
   establish which of the three figures are wrong, what each one actually
   measures, and what the corrected sentence should say.

6. Demonstrate fan-out numerically. Join `Sales` to `ResellerSales` on
   `ProductKey` for a single product using `FILTER`, report the resulting row
   count against the two input counts, and state the arithmetic relationship
   between them.

## Solutions

**1.** `FactResellerSales`: one row is one product line on one reseller order —
`SalesAmount` is safe to sum, `UnitPrice` is not meaningfully averageable
because lines carry wildly different quantities, so a quantity-weighted average
is required. `DimReseller`: one row is one reseller — `AnnualSales` can be
averaged across resellers, but note the data dictionary's warning that this
dimension attribute disagrees with summed fact revenue, because it is survey
data rather than transactions. `DimPromotion`: one row is one promotion —
`DiscountPct` averages across promotions, which is almost never what anyone
wants, since promotion key 1 (`No Discount`) covers most rows and a
revenue-weighted figure is the business-relevant one.

**2.** `CarrierTrackingNumber` and `CustomerPONumber` are largely blank. A
reasonable working threshold is that a column above roughly 30 per cent null
cannot support a population-level claim, and above 60 per cent should not be
analysed at all without first establishing whether missingness is random. The
discipline that matters is less the threshold than stating it: *"we excluded
`CustomerPONumber` as 70 per cent null"* is defensible; silently dropping it is
not.

**3.** A single verifying formula:
`=SUMPRODUCT(--(ResellerSales[OrderDate_Parsed]<>ResellerSales[OrderDate_Real]))`
should return 0. The lookup remains preferable even though both work, because it
is one join that also delivers fiscal year, month name, quarter and week number
in the same consistent way for every analyst on the team, and because an
unmatched key surfaces as a visible failure rather than a silently parsed
integer.

**4.** Calendar 2013 covers January to December; fiscal 2013 covers July 2012 to
June 2013. On a growing business the fiscal figure is lower, because it reaches
further back into a smaller period. The sentence under the exhibit should read
something like: *"Revenue on a calendar-year basis (January–December 2013);
fiscal-year figures differ as AdventureWorks' fiscal year begins in July."*

**5.** All three figures are wrong. 121,000 is the combined *line* count of the
two facts, not orders; distinct `SalesOrderNumber` across both channels is far
lower. 613 is a count of `ProductKey`, which counts product versions and
includes non-sellable components with a null `ProductSubcategoryKey`; distinct
sellable `ProductAlternateKey` is materially smaller. And $460 is total revenue
divided by lines, which is revenue per line — a figure that averages a $4 water
bottle against a $3,578 road bike and that differs by a large factor between the
two channels, so a single blended figure conceals the only interesting thing
about it. The corrected sentence reports orders and lines separately, counts
distinct sellable business keys, and gives average order value per channel
rather than blended.

**6.** If the chosen product has *m* internet lines and *n* reseller lines, the
joined result has *m × n* rows, and summed `SalesAmount` over it is inflated by
roughly a factor of *n* on the internet side and *m* on the reseller side. The
relationship is multiplicative, which is why fan-out errors are usually large
enough to be caught — and why the dangerous case is a join that fans out only
slightly, where the inflation is small enough to look like a rounding
difference.

## Key takeaways

- Facts are measurable events; dimensions are the nouns you group by. Everything
  you sum is in a fact, everything on an axis is in a dimension.
- State the grain of every table as a sentence before you touch it. Most wrong
  numbers in decks are grain errors, not arithmetic errors.
- Join on surrogate keys; count and report on business keys. `ProductKey` counts
  versions, `ProductAlternateKey` counts products.
- `DateKey` is an integer. Arithmetic on it is silently wrong. Convert it in the
  transform layer by looking it up in `DimDate`.
- You may only join on a column that is unique in at least one table.
  Fact-to-fact joins fan out and inflate totals without erroring.
- `SalesOrderNumber` plus `SalesOrderLineNumber` is the real key. Report orders
  and lines separately and label which is which.
- A company view requires both channels. Internet sales alone understate
  AdventureWorks by roughly a factor of four.
- Profile every table before analysing it, and write the profile down. The
  distinct-values check on categorical columns is the highest-yield ten minutes
  in the work.
