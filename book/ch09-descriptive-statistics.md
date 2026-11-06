# Chapter 9 — Descriptive Statistics and Distributions

## Learning objectives

By the end of this chapter you will be able to:

- Choose between the mean, the median and the mode for a given business
  variable, and explain the choice to a client in one sentence.
- Measure spread with `STDEV.S`, the interquartile range and the coefficient of
  variation, and say which one belongs in which exhibit.
- Build an honest histogram in Excel, with `FREQUENCY` and with a PivotTable of
  binned values, and defend your bin widths.
- Detect outliers with the IQR rule and make a defensible decision about
  whether to exclude them.
- Compute weighted averages and recognise the common cases where an unweighted
  average is not merely imprecise but meaningless.
- Quantify concentration with a Pareto curve and read the result as a
  commercial fact rather than a chart.
- Use `CORREL` without overclaiming, and recognise when aggregation has
  inflated a correlation.
- Produce a confidence interval with `CONFIDENCE.T` and state accurately what
  it does and does not mean.

## The business question

AdventureWorks' Chief Commercial Officer opens the meeting with a number:
"Our average online order is about $480. So if we can nudge the average order
up by ten per cent, that's a few million dollars of revenue. What's the fastest
way to do that?"

That is a reasonable-sounding question built on an unreasonable statistic. The
average internet order value in AdventureWorks is around $480, and almost no
order is anywhere near $480. The distribution is violently bimodal: a very
large number of small accessory and clothing orders in the $5–$60 range, and a
much smaller number of bicycle orders between roughly $500 and $3,600. The
"average order" describes a transaction that essentially never happens.

Our real question is therefore the one behind the question: **what does the
distribution of AdventureWorks' revenue actually look like, across orders and
across customers, and what does its shape imply about where growth can come
from?** A ten per cent lift in a meaningless average is a meaningless target. A
shift in the mix between accessory orders and bike orders, or a change in the
concentration of revenue among customers, is a strategy.

This chapter is where consulting and statistics meet. Every technique in it
exists to answer one of two questions a client will always ask: *is this number
typical?* and *how much does it vary?*

## Data you will use

From [the data dictionary](../data/adventureworks-data-dictionary.md):

- **`FactInternetSales`** — loaded as the table `Sales`. The columns that
  matter here are `SalesOrderNumber`, `SalesAmount`, `TotalProductCost`,
  `OrderQuantity`, `UnitPrice`, `ProductKey` and `CustomerKey`.
- **`FactResellerSales`** — loaded as `ResellerSales`, for the comparison of
  spread across two channels with very different scale.
- **`DimProduct`** — loaded as `Products`, for `EnglishProductName` and
  `ProductSubcategoryKey`.
- **`DimProductCategory`** / **`DimProductSubcategory`** — for grouping to the
  four categories.
- **`DimCustomer`** — loaded as `Customers`, for `CustomerKey`,
  `YearlyIncome`, `EnglishOccupation` and `DateFirstPurchase`.

Two derived definitions, used throughout and stated once:

- **Order value** is the sum of `SalesAmount` for all lines sharing a
  `SalesOrderNumber`. Because `FactInternetSales` has one row per order *line*,
  the row-level `SalesAmount` is a line value, not an order value. Confusing the
  two is the single most common error in this chapter.
- **Customer lifetime revenue** is the sum of `SalesAmount` for all rows
  sharing a `CustomerKey`, across the whole four-year window. It is lifetime to
  date, not lifetime value: no forecast, no discounting.

Build both as helper tables before you compute a single statistic. Statistics
computed at the wrong grain are not approximately right; they are answers to a
different question.

## Core concepts

### The mean is a summary that assumes symmetry

The arithmetic mean is the balance point of a distribution. It is an excellent
summary when the distribution is roughly symmetric and single-peaked, and a
misleading one otherwise, because it is dragged by the tail.

For AdventureWorks internet orders, the mean order value sits at roughly $480
while the median sits far below it — in the low tens of dollars once the
accessory orders are counted properly as orders. That gap is the whole story.
When mean and median diverge, the distribution is skewed, and the mean has
stopped describing a typical case.

```
=AVERAGE(Orders[OrderValue])
=MEDIAN(Orders[OrderValue])
=MODE.SNGL(ROUND(Orders[OrderValue],0))
```

The mode deserves a word. For continuous money amounts it is usually
meaningless — no two orders are the same to the cent. For AdventureWorks it is
unusually informative, because the catalogue has fixed list prices and most
orders are a single unit, so order values cluster hard on specific price
points. A mode of $4.99 or $34.99 tells you the modal order is one bottle cage
or one pair of socks.

The practical rule for consulting work: **report the median as the typical
case, report the mean when you need a figure that multiplies up to the total**
(because mean × count = sum, and the median has no such property), and never
report one without knowing the other.

### Quartiles and percentiles describe shape without assuming it

Percentiles make no distributional assumption at all, which is why they survive
contact with real business data.

```
=QUARTILE.INC(Orders[OrderValue],1)
=QUARTILE.INC(Orders[OrderValue],3)
=PERCENTILE.INC(Orders[OrderValue],0.9)
=PERCENTILE.INC(Orders[OrderValue],0.99)
```

`QUARTILE.INC` and `PERCENTILE.INC` interpolate between data points and include
the extremes; the `.EXC` variants exclude them and are stricter about small
samples. With sixty thousand rows the difference is immaterial. With thirty
rows it is not, and `.EXC` is the more honest choice.

The interquartile range, `Q3 − Q1`, is the width of the middle half of the
data. For a skewed distribution it is a better spread measure than the standard
deviation precisely because it ignores the tail that is distorting the mean.

### Standard deviation: `STDEV.S` versus `STDEV.P`

```
=STDEV.S(Orders[OrderValue])
=STDEV.P(Orders[OrderValue])
=VAR.S(Orders[OrderValue])
```

`STDEV.S` divides by *n − 1* and estimates the standard deviation of a
population from a sample. `STDEV.P` divides by *n* and assumes your data *is*
the population. The difference is negligible above a few hundred rows.

The judgment is not arithmetic, it is conceptual: **is this dataset the whole
population or a sample from something larger?** Four years of AdventureWorks
transactions are the complete record of what happened, so `STDEV.P` is
defensible. But analysts almost never care about the variability of the past
for its own sake — they care about it as evidence for what varies in general,
which makes the data a sample and `STDEV.S` the right default. Use `STDEV.S`
unless you can articulate why not.

### Skewness and kurtosis, used sparingly

```
=SKEW(Orders[OrderValue])
=KURT(Orders[OrderValue])
```

`SKEW` returns zero for a symmetric distribution, positive for a right tail,
negative for a left tail. AdventureWorks order values return a strongly
positive figure. `KURT` returns excess kurtosis — zero for a normal
distribution, positive for fat tails and a sharp peak.

These belong in your working papers, not in a client deck. No executive has
ever made a decision because kurtosis was 4.2. Their job is to tell *you*, in
one number, that the mean is not safe to use — and then you say that in words.

### The coefficient of variation compares spread across different scales

A standard deviation of $400 means something very different for bikes, where
the mean line value is in the hundreds, than for accessories, where it is in
the tens. To compare variability across categories of different magnitude,
normalise:

```
=STDEV.S(FILTER(Sales[SalesAmount],Sales[Category]="Bikes"))
 /AVERAGE(FILTER(Sales[SalesAmount],Sales[Category]="Bikes"))
```

That ratio — standard deviation over mean — is the coefficient of variation. It
is dimensionless, so a CV of 0.3 for bikes and 1.4 for accessories is a
meaningful comparison: accessory line values are far more erratic relative to
their own size.

Two cautions. The CV is only interpretable for variables on a ratio scale with
a meaningful zero, so it works for revenue and not for, say, a satisfaction
score. And it explodes when the mean approaches zero, which makes it useless
for margin figures that straddle break-even.

### Histograms, and the quiet dishonesty of bin width

A histogram is the fastest way to show a client what you mean by "skewed". It
is also the easiest chart to mislead with, because bin width is a free
parameter and different widths tell different stories.

The `FREQUENCY` approach needs a column of bin upper bounds, then one dynamic
array formula:

```
=FREQUENCY(Orders[OrderValue],BinUpperBounds)
```

The PivotTable approach is often better for a consulting workflow because it
survives a data refresh: put `OrderValue` in Rows, right-click, **Group**, and
set a start, end and interval. The grouping is recorded in the pivot and
reapplies when the data changes.

Three rules for honest bins:

1. **Equal widths**, unless you label unequal widths explicitly. Unequal bins
   with a frequency y-axis exaggerate the wide bins.
2. **Round, human numbers.** Bins of $0–$50, $50–$100 are readable; bins of
   $0–$47.30 are a sign you let Excel choose.
3. **Show the tail, do not let it dominate.** AdventureWorks order values
   require either a log-scaled axis, a truncated final bin labelled "$2,000+",
   or two separate histograms — one for sub-$100 orders, one for bikes. The
   two-chart solution is usually the clearest, and the fact that you *need*
   two charts is itself the finding.

Starting-point bin width: the Freedman–Diaconis rule, `2 × IQR / n^(1/3)`,
computed in a cell so the reader can see it:

```
=2*(QUARTILE.INC(Orders[OrderValue],3)-QUARTILE.INC(Orders[OrderValue],1))
 /COUNT(Orders[OrderValue])^(1/3)
```

Then round to something human. The rule gives you a defensible starting point,
not an answer.

### Box-plot thinking and the IQR rule

You do not need to draw a box plot to think like one. The five-number summary —
minimum, Q1, median, Q3, maximum — plus the IQR fence is a complete, compact
characterisation of a distribution that takes five cells.

The standard outlier fence:

```
=QUARTILE.INC(Orders[OrderValue],3)
 +1.5*(QUARTILE.INC(Orders[OrderValue],3)-QUARTILE.INC(Orders[OrderValue],1))
```

Anything above that upper fence, or below the symmetric lower fence, is flagged
as an outlier by convention.

**The convention is a detector, not a verdict.** Run it on AdventureWorks
internet orders and it will flag every single bicycle order, because bikes sit
in the tail of a distribution dominated by accessories. A $3,500 Road-150 order
is not a data error. It is the company's core product, correctly recorded, and
it accounts for a large share of revenue.

This is the most important judgment in the chapter, so state it plainly:

- An outlier is a **data error** when it is impossible or internally
  inconsistent — a negative `OrderQuantity`, a `SalesAmount` that does not equal
  `UnitPrice × OrderQuantity` net of discount, an order dated before the
  customer's `DateFirstPurchase`. Investigate and fix or exclude, and document
  what you did.
- An outlier is a **real extreme** when it is a genuine, rare, large event. You
  do not delete it. You either keep it and use robust statistics, or you split
  the population into segments where each is internally coherent.
- An outlier is a **mixture signal** when there are many of them clustered
  together, as here. That means you have pooled two populations that should be
  analysed separately. Bikes and accessories are different businesses. Analyse
  them as such, and the outliers disappear because you stopped asking one
  distribution to describe two things.

Never silently exclude. Every exclusion is a decision the client is entitled to
see, with a row count attached.

### Weighted averages, and averages of averages

Here is a mistake that reaches client decks more often than any other in this
chapter. Suppose you want the average selling price of an AdventureWorks
product. The tempting formula averages `ListPrice` over the product dimension:

```
=AVERAGE(Products[ListPrice])
```

That number answers the question "if I picked a catalogue entry at random, what
would its list price be?" — a question nobody asked. The catalogue contains
hundreds of bicycle variants that sell in tiny volumes and a handful of
accessories that sell constantly. It also contains multiple versioned rows for
the same product, so the "average" double-counts price changes.

The business question is "what does a unit sold cost, on average?", and that is
a weighted average, weighted by units:

```
=SUMPRODUCT(Sales[UnitPrice],Sales[OrderQuantity])/SUM(Sales[OrderQuantity])
```

Equivalently, and more robustly, revenue over units:

```
=SUM(Sales[SalesAmount])/SUM(Sales[OrderQuantity])
```

The two differ when discounts are applied, and the difference between them is
itself a useful figure: list-price realisation.

The general rule: **an average of averages is almost always wrong unless the
groups are equal in size.** A margin percentage averaged across four categories
is not the company margin percentage. The company margin percentage is total
margin over total revenue:

```
=(SUM(Sales[SalesAmount])-SUM(Sales[TotalProductCost]))/SUM(Sales[SalesAmount])
```

Compute ratios from summed numerators and summed denominators. Always.

### Concentration and the Pareto curve

Concentration is where descriptive statistics turns into strategy. The
question is not "what is the average customer worth" but "how much of our
revenue depends on how few of them".

The mechanics: rank entities by revenue descending, take a running total,
divide by the grand total, and plot the cumulative share against the
cumulative count share.

```
=SUM($B$2:B2)/SUM($B:$B)
```

Or, with dynamic arrays and no helper column:

```
=LET(v,SORT(CustRev[Revenue],1,-1),
     SUM(TAKE(v,ROUND(0.1*ROWS(v),0)))/SUM(v))
```

That returns the revenue share of the top decile of customers directly.

For AdventureWorks you will find two very different concentration profiles.
Internet customers are relatively flat — roughly the top tenth of customers
account for something in the region of a quarter to a third of internet
revenue, because most customers buy one bike or a handful of accessories and
then stop. Reseller revenue is far more concentrated: a small number of
warehouse-type resellers carry a large share of wholesale turnover.

Those two facts point at opposite strategies, and that is the payoff of the
whole chapter. A flat distribution means growth comes from acquiring more
customers or from raising the typical basket. A concentrated distribution means
growth — and risk — lives in a handful of named accounts, and the right move is
account management, not a marketing campaign.

### `CORREL`, and why aggregation inflates it

```
=CORREL(CustRev[YearlyIncome],CustRev[LifetimeRevenue])
```

`CORREL` returns Pearson's *r*, between −1 and 1. It measures linear
association only, is not robust to outliers, and says nothing about causation.
All of that is standard and you know it.

The less-known trap matters more in practice. Run that correlation at customer
level on AdventureWorks and you will get something modest — individual
behaviour is noisy. Now aggregate: compute average income and average revenue
per occupation group, and correlate the six resulting points. The correlation
will leap, possibly above 0.8.

Nothing about the relationship changed. Averaging cancelled the individual
noise and left only the between-group signal, so the correlation of the
averages measures a different, stronger thing. This is the ecological fallacy,
and it is how a weak relationship gets presented as a strong one by an analyst
who is not being dishonest, merely careless.

The rule: **state the unit of observation next to every correlation.** "r =
0.82 across six occupation groups" is honest. "r = 0.82" is not, because the
reader will assume it is customer-level.

### The ToolPak, and `CONFIDENCE.T`

**Data → Data Analysis → Descriptive Statistics** produces mean, standard
error, median, mode, standard deviation, variance, kurtosis, skewness, range,
min, max, sum, count and a confidence level in one pass. It is fast, and it is
the right tool for a first look at an unfamiliar column.

Its limits are real. The output is **static values, not formulas** — it does
not refresh when the data changes, and six weeks later nobody can tell which
version of the data produced it. It cannot be grouped by a dimension. And it
quietly ignores text and blanks, so a column of numbers-stored-as-text yields a
confident, wrong answer with a low count. Always reconcile the ToolPak's
`Count` against your own `=COUNTA()` before believing anything else in the
block.

For a confidence interval on a mean:

```
=CONFIDENCE.T(0.05,STDEV.S(Orders[OrderValue]),COUNT(Orders[OrderValue]))
```

That returns the half-width. The interval is mean ± that value. Use
`CONFIDENCE.T` rather than `CONFIDENCE.NORM` unless you know the population
standard deviation, which you do not.

What it means: if we repeated this sampling process many times, about 95 per
cent of the intervals so constructed would contain the true population mean.

What it does **not** mean: that there is a 95 per cent probability the true
mean lies in this particular interval, or that 95 per cent of orders fall
inside it. The second confusion is the common one in business settings, and it
is a confusion between a confidence interval and a prediction interval. If a
client wants to know the range most orders fall in, give them percentiles, not
a confidence interval.

And one honest caveat: with sixty thousand rows, the confidence interval on the
mean is so tight as to be uninteresting. Sampling error is not your problem
here. Skew, mixture and wrong-grain errors are. Spend your scepticism where the
risk actually is.

## Walkthrough

### Step 1 — Build the order-level table

Start in a fresh sheet named `w_orders`. The `w_` prefix marks it as working
rather than output, per the convention in Chapter 1.

With the `Sales` table loaded, the order grain comes from a `UNIQUE` of order
numbers plus a `SUMIFS`:

```
=UNIQUE(Sales[SalesOrderNumber])
=SUMIFS(Sales[SalesAmount],Sales[SalesOrderNumber],$A2)
=SUMIFS(Sales[TotalProductCost],Sales[SalesOrderNumber],$A2)
=COUNTIFS(Sales[SalesOrderNumber],$A2)
```

Convert the result to a table named `Orders` with columns `SalesOrderNumber`,
`OrderValue`, `OrderCost` and `LineCount`. You will get roughly 27,600 internet
orders from roughly 60,400 lines — confirming that order grain and line grain
genuinely differ, by a factor of a bit over two.

Check the total before going further:

```
=SUM(Orders[OrderValue])-SUM(Sales[SalesAmount])
```

That must return exactly zero. If it does not, your aggregation has lost or
duplicated rows and nothing downstream is worth computing. This single cell is
the most valuable formula in the walkthrough.

### Step 2 — Characterise the order-value distribution

In a block labelled "Internet order value, FY2011–FY2014", lay out the
statistics with labels in column A and formulas in column B:

```
=COUNT(Orders[OrderValue])
=SUM(Orders[OrderValue])
=AVERAGE(Orders[OrderValue])
=MEDIAN(Orders[OrderValue])
=STDEV.S(Orders[OrderValue])
=QUARTILE.INC(Orders[OrderValue],1)
=QUARTILE.INC(Orders[OrderValue],3)
=PERCENTILE.INC(Orders[OrderValue],0.9)
=PERCENTILE.INC(Orders[OrderValue],0.99)
=SKEW(Orders[OrderValue])
```

Add a mean-to-median ratio and a CV, because those two numbers are what you
will actually quote:

```
=AVERAGE(Orders[OrderValue])/MEDIAN(Orders[OrderValue])
=STDEV.S(Orders[OrderValue])/AVERAGE(Orders[OrderValue])
```

The pattern you should see: a mean in the region of $1,000 at order grain (the
line-level $480 figure the CCO quoted was itself at the wrong grain), a median
far lower, a mean-to-median ratio well above 1, a CV above 1 and strong
positive skew. Every one of those says the same thing in a different dialect:
*do not use the mean*.

### Step 3 — Show the mixture

Add a category flag to `Orders` so you can see the two populations. The
cleanest route is a helper in `Sales` that resolves each line's category, then a
dominant-category assignment per order. For a first pass, a simpler binary
flag earns its keep:

```
=IF(SUMIFS(Sales[SalesAmount],Sales[SalesOrderNumber],$A2,
           Sales[Category],"Bikes")>0,"Bike order","Non-bike order")
```

Now compute the same statistics block separately for each flag. Two things
happen. The medians separate dramatically — bike orders in the high hundreds,
non-bike orders in the tens. And the skew and CV *within* each group fall
sharply, because you have stopped pooling two distributions.

That is your evidence for the central claim: AdventureWorks does not have one
order-value distribution. It has two, and no single summary statistic can
describe both.

### Step 4 — Build the histograms

Create bin bounds in a column: 0, 25, 50, 100, 200, 500, 1,000, 2,000, 4,000.
These are unequal, so you must either label the chart as a log-scale view or,
better, build two equal-width histograms.

Chart one, non-bike orders, bins of $10 from $0 to $150 with a "$150+"
overflow:

```
=FREQUENCY(FILTER(Orders[OrderValue],Orders[Flag]="Non-bike order"),NonBikeBins)
```

Chart two, bike orders, bins of $250 from $0 to $4,000:

```
=FREQUENCY(FILTER(Orders[OrderValue],Orders[Flag]="Bike order"),BikeBins)
```

Title them explicitly — "Non-bike internet orders, $10 bins" and "Bike internet
orders, $250 bins" — so the reader cannot mistake one scale for the other.

### Step 5 — Build customer lifetime revenue and the Pareto exhibit

New sheet, `w_customers`:

```
=UNIQUE(Sales[CustomerKey])
=SUMIFS(Sales[SalesAmount],Sales[CustomerKey],$A2)
=COUNTIFS(Sales[CustomerKey],$A2)
```

Name the result `CustRev` with columns `CustomerKey`, `LifetimeRevenue` and
`Lines`. Then sort descending and build the cumulative curve:

```
=SORT(CustRev[LifetimeRevenue],1,-1)
=SUM($D$2:D2)/SUM($D$2:$D$20000)
=ROW()-1
=(ROW()-1)/COUNT(CustRev[LifetimeRevenue])
```

Plot cumulative revenue share on the y-axis against cumulative customer share
on the x-axis, as a line chart with no markers. Add a 45-degree reference line;
the gap between your curve and that line *is* the concentration.

Pull out the decile figures to quote:

```
=LET(v,SORT(CustRev[LifetimeRevenue],1,-1),
     SUM(TAKE(v,ROUND(0.1*ROWS(v),0)))/SUM(v))
=LET(v,SORT(CustRev[LifetimeRevenue],1,-1),
     SUM(TAKE(v,ROUND(0.2*ROWS(v),0)))/SUM(v))
```

Repeat the whole exercise for products, using `ProductKey` grouped to
`EnglishProductName`, and for resellers using `ResellerSales`. The three curves
on one chart is the exhibit.

### Step 6 — Write the three sentences

This is the step students skip and consultants cannot. An exhibit without a
sentence is a decoration. Write, under the Pareto chart, the three sentences
you would say out loud:

> "Internet revenue is not concentrated — the top tenth of online customers
> drive roughly a third of online revenue, because most buy once and stop.
> Reseller revenue is the opposite: a handful of warehouse accounts carry the
> majority of wholesale turnover, which is both our largest revenue line and
> our largest single point of risk. So the online question is how to get a
> second purchase, and the wholesale question is how to protect six
> relationships — they are different problems and should not share a growth
> plan."

Note what those sentences contain: a number, a mechanism, and a decision. Note
what they do not contain: skewness, kurtosis, or the word "average".

## Common pitfalls

**Computing order statistics at line grain.** `FactInternetSales` is one row
per line. `=AVERAGE(Sales[SalesAmount])` is the average *line* value and will be
roughly half the average order value. Aggregate to `SalesOrderNumber` first,
and reconcile the total to zero.

**Letting the IQR rule make the decision.** The fence flags every bike order.
Flagging is not excluding. If the rule flags thousands of rows, you have a
mixture problem, not an outlier problem.

**Averaging percentages.** The average of four category margin percentages is
not the company margin percentage. Sum the numerator, sum the denominator,
divide once.

**Unweighted average price.** `=AVERAGE(Products[ListPrice])` answers a question
about the catalogue, not the business. It is also corrupted by SCD versioning,
which puts the same product in the table more than once.

**Reporting a correlation without its unit of observation.** The same
relationship yields *r* = 0.3 at customer level and *r* = 0.85 across six
occupation groups. Both are arithmetically correct. Only one is what the reader
will assume.

**Trusting the ToolPak's count.** It skips text and blanks silently. If
`HouseOwnerFlag`-style text-stored numbers reach your column, the ToolPak will
report statistics on a fraction of the rows with no warning. Reconcile `Count`
against `=COUNTA()` every time.

**Confusing a confidence interval with a range of outcomes.** The confidence
interval is about the precision of your estimate of the mean. The client
usually wants to know where most orders fall, which is a percentile question.

**Choosing bins that flatter the story.** If you tried three bin widths and
kept the one that looked most dramatic, say so or go back. The quickest way to
lose credibility in a review is to have a partner change your bin width and get
a different conclusion.

**Treating `DimProduct` row counts as product counts.** Products are versioned.
Use `ProductAlternateKey` for a distinct product count, per the data
dictionary.

## Exercises

1. **Order grain versus line grain.** Build the `Orders` table as in Step 1.
   Report the row count of `Sales`, the row count of `Orders`, the mean and
   median at each grain, and the reconciliation cell proving the totals match.
   In two sentences, explain to the CCO why the "$480 average order" figure
   quoted in the meeting was wrong in two separate ways.

2. **Spread across categories.** For each of the four product categories
   (Bikes, Components, Clothing, Accessories), compute line-level mean, median,
   `STDEV.S` and the coefficient of variation from `Sales`. Rank the categories
   by CV. Which category is most erratic relative to its own size, and give a
   commercial reason why that would be expected.

3. **Honest histograms.** Produce the Freedman–Diaconis bin width for non-bike
   order values, then build three histograms of the same data: one with the FD
   width, one with bins five times wider, one with bins five times narrower.
   Which would you put in a client deck, and what specifically does each of the
   other two hide or exaggerate?

4. **The outlier decision.** Apply the IQR fence to all internet order values
   and count the flagged rows and their revenue share. Then apply it separately
   within bike orders and within non-bike orders. Write the paragraph you would
   put in an appendix stating exactly which rows, if any, you excluded from the
   analysis and why.

5. **Weighted versus unweighted price.** Compute the unweighted average of
   `Products[ListPrice]`, the units-weighted average selling price from
   `Sales`, and the revenue-over-units figure. Explain the gap between the
   second and third, and state which single number you would give a client who
   asks "what does a typical AdventureWorks product sell for?" — and what you
   would say alongside it.

6. **Concentration and the ecological fallacy.** Build the Pareto curve for
   internet customers, reseller accounts and products. Report the top-decile
   revenue share for each. Then compute `CORREL` between `YearlyIncome` and
   lifetime revenue at customer level, and again on the six `EnglishOccupation`
   group averages. Report both with their units of observation and explain in
   three sentences why they differ.

## Solutions

**1.** Roughly 60,400 lines collapse to roughly 27,600 orders, so line-grain
statistics understate order values by a factor of about 2.2. The totals must
reconcile to exactly zero; if they do not, a `SUMIFS` criterion is wrong or
order numbers have trailing spaces. The "$480" figure was wrong twice over:
it was computed at line grain rather than order grain, and even at the right
grain the mean is not a typical order, because the distribution is a mixture of
accessory orders in the tens of dollars and bike orders in the hundreds to
low thousands. The honest framing is "half our online orders are under about
$40; about a fifth are bikes and those average well over $1,000".

**2.** Bikes show the lowest CV — typically well under 1 — because bicycles sell
one unit at a time at fixed, high list prices, so line values cluster. Clothing
and Accessories show CVs above 1: cheap items with occasional multi-unit lines
produce a long relative tail. Components sit in between and are the most
version-affected. The commercial reason is unit economics, not data quality: a
category where a line can be either one $5 bottle cage or eight of them has
inherently higher relative dispersion than one where every line is a single
$2,300 bicycle.

**3.** The FD width lands near the low single-digit dollars for non-bike
orders, which is too fine to read — a useful demonstration that the rule gives
a starting point, not an answer; round up to $10. The five-times-wider version
collapses the $0–$50 region into one or two bars and hides the price-point
clustering that is the actual finding. The five-times-narrower version turns
the chart into noise and creates spurious spikes at individual list prices.
Deck-worthy is the $10-bin version with an explicit "$150+" overflow bar.

**4.** Pooled, the fence flags somewhere in the region of ten to fifteen per
cent of orders but a far larger share of revenue — often well over half. That
asymmetry is the proof that these are not errors: you cannot discard a minority
of rows carrying a majority of revenue and claim to have analysed the business.
Split by flag, each subpopulation flags only a small number of genuinely
unusual rows. The appendix paragraph should say: no rows were excluded on
statistical grounds; rows were excluded only where `SalesAmount` failed the
`UnitPrice × OrderQuantity` net-of-discount identity, and the count of those is
stated. Document the count even when it is zero.

**5.** The unweighted catalogue average is pulled up by hundreds of low-volume
bicycle variants and inflated further by SCD duplicate rows. The units-weighted
`UnitPrice` average is lower and reflects what customers actually bought. The
revenue-over-units figure is lower still; the gap between the two weighted
figures is discount and price realisation, and quantifying that gap is more
useful than either number alone. To the client: give revenue over units, and say
"but the catalogue spans $2 to $3,600, so the average is a bookkeeping figure —
here is the distribution."

**6.** Internet top-decile share typically lands around a quarter to a third;
reseller concentration is much higher, often a majority in the top decile of
accounts; product concentration is high too, driven by a handful of
Mountain-200 and Road-250 variants. Customer-level income-to-revenue
correlation is weak, in the low tenths, because individual purchase behaviour
is dominated by idiosyncratic factors. The six-point occupation-group
correlation is far higher because averaging removed the individual noise and
left only between-group structure. The relationship did not strengthen; the
variance that was obscuring it was averaged away. Report both, each labelled
with its unit of observation, and never quote the second as if it described
individuals.

## Key takeaways

- **Get the grain right before you compute anything.** A statistic at the wrong
  grain is not imprecise, it is an answer to a different question. The
  reconciliation cell that proves your aggregate matches the source total is the
  cheapest insurance in analytics.
- **Mean and median together, or neither.** The gap between them is the
  information. Report the median as typical, the mean when you need it to
  multiply up to a total, and say which you used.
- **Compute ratios from sums, never average ratios.** This single rule prevents
  more wrong numbers in client decks than any other in this book.
- **Many outliers mean a mixture, not errors.** If a rule flags thousands of
  rows, segment the population instead of deleting the inconvenient part of it.
  Bikes and accessories are different businesses.
- **State the unit of observation next to every correlation.** Aggregation
  inflates correlation, and the reader will assume you worked at the individual
  level unless you tell them otherwise.
- **Concentration is strategy.** Flat distributions point to acquisition and
  basket size; concentrated ones point to account management and risk. The
  Pareto curve is the cheapest strategic exhibit you can build.
- **Every exhibit needs its three sentences** — a number, a mechanism, a
  decision. If you cannot write them, you have not finished the analysis.
