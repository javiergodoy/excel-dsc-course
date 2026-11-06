# Chapter 12 — Customer Segmentation and Cohorts

Chapter 11 ended with an uncomfortable result: the attributes AdventureWorks can
target explain only a small share of customer value. The marketing director
cannot buy her way to good customers. That makes this chapter the more
important of the pair, because if you cannot select better customers in advance,
you have to treat the ones you have differently — and to do that you need to
know who they are.

Segmentation is how. Done well it is one of the highest-return pieces of
analysis a consultant can deliver. Done badly it produces a colourful slide with
four named boxes that changes nothing, and most segmentation in the wild is done
badly.

## Learning objectives

By the end of this chapter you will be able to:

- State the test any segmentation must pass before it is worth building.
- Compute Recency, Frequency and Monetary values correctly from
  `FactInternetSales` against an explicit as-of date.
- Score customers into quintiles with `PERCENTRANK` or `QUARTILE` and understand
  what the scoring convention does and does not preserve.
- Assemble, name and — critically — **size** segments by revenue rather than by
  headcount.
- Build an ABC / Pareto classification with a running-total share, and test the
  80/20 claim rather than assuming it.
- Distinguish a descriptive segment from an actionable one.
- Build a cohort grid from `DateFirstPurchase` and read it correctly down the
  columns and across the rows.
- Report an honest retention finding when most customers buy exactly once.
- Explain why Solver-based clustering in Excel is a demonstration rather than a
  method, and what to do instead.
- Segment resellers as the B2B counterpart, where the economics differ.

## The business question

The marketing director has a retention budget as well as an acquisition budget,
and no basis for allocating it. She wants to know: which customers should we
spend money on, how much revenue is at stake in each group, and what should we
actually do differently for each?

That last clause is the whole test. **A segment is only worth having if you
would take a different action for it.** If two segments get the same email, they
are one segment with extra steps. Before you build anything, write down the
action you expect each segment to receive. If you cannot, you are building a
chart rather than an analysis, and you should stop.

The second question behind the brief is about retention itself. AdventureWorks
internet sales look, on inspection, like a business of one-time buyers. If that
is true, the retention budget has nothing to retain, and the honest finding is
that the company has a repeat-purchase problem rather than a targeting problem.
That is a far more valuable thing to tell a client than a four-box slide, and
the cohort grid in this chapter is how you establish it.

## Data you will use

From `FactInternetSales`: `CustomerKey`, `SalesOrderNumber`, `OrderDateKey`,
`SalesAmount`, `TotalProductCost`, `ProductKey`.

From `DimCustomer`: `CustomerKey`, `DateFirstPurchase`, `YearlyIncome`,
`EnglishEducation`, `EnglishOccupation`, `CommuteDistance`, `TotalChildren`,
`NumberChildrenAtHome`, `MaritalStatus`, `Gender`.

From `DimProduct` and its hierarchy: `ProductKey`, `EnglishProductName`,
`ProductSubcategoryKey`, and category via `DimProductSubcategory` →
`DimProductCategory`.

From `DimReseller`: `ResellerKey`, `ResellerName`, `BusinessType`,
`AnnualSales`, `NumberEmployees`, `YearOpened`. From `FactResellerSales`:
`ResellerKey`, `SalesAmount`, `SalesOrderNumber`, `OrderDateKey`.

From `DimDate`: `DateKey`, `FullDateAlternateKey`, `CalendarYear`,
`MonthNumberOfYear`.

Column definitions and the dataset's quirks are in
[the data dictionary](../data/adventureworks-data-dictionary.md). Three matter
here. `OrderDateKey` is an integer in `yyyymmdd` form, so you cannot subtract it
to get days. `CommuteDistance` sorts wrongly as text. And the sales window ends
in a partial month, which distorts the final cohort unless you exclude it.

**Fix an as-of date before you compute anything.** This book uses
**1 January 2014**, just past the end of the sales window. Recency is
meaningless without it, and `TODAY()` would make every number in the deck change
between the day you build it and the day you present it.

## Core concepts

### RFM, and why it survives

RFM scores customers on three behavioural facts: how recently they bought, how
often, and how much they spent. It is decades old, it uses no statistics beyond
ranking, and it still outperforms most elaborate alternatives for direct
marketing allocation.

It works because the three dimensions are nearly orthogonal and each maps to a
different decision. **Recency** predicts responsiveness — someone who bought
last month is reachable in a way that someone who bought three years ago is
not. **Frequency** measures habit. **Monetary** measures value, and tells you how
much you may spend to keep them.

Crucially, RFM is built entirely from behaviour. Chapter 11 showed that
demographics explain little; behaviour explains much more, because what someone
did is a better guide to what they will do than who they are.

### Scoring, and what ranking throws away

Convention is to split each dimension into quintiles and score 1 to 5, with 5
best. For Frequency and Monetary, 5 is the highest value. For Recency, 5 is the
*smallest* number of days since last purchase — the inversion catches people
out constantly, and a reversed Recency score produces a segmentation that
recommends spending money on the customers who have already left.

Ranking into quintiles buys robustness: outliers cannot distort a rank, and the
scores are comparable across the three dimensions. It costs you magnitude. A
Monetary score of 5 is assigned to the top fifth whether that fifth spent
$2,000 or $200,000 each. So **score to assign segments, but always size the
segments in actual dollars** afterwards. A segmentation presented only in scores
hides the economics it exists to expose.

Note also that quintiles force exactly 20% of customers into each band by
construction. If a dimension is heavily tied — and Frequency in AdventureWorks
is, because most customers have exactly one order — the quintile boundaries
collapse and the bands will not be equal. That is a property of the data
showing through, and it is the first signal of the single-purchase problem.

### Descriptive versus actionable

A descriptive segment tells you what a group looks like: "high-income
professionals aged 35 to 45 in the Pacific group". An actionable segment tells
you what to do: "customers who bought once, in the last 90 days, in
Accessories — send a bike-category cross-sell".

The difference is reachability and a decision. Demographic segments are usually
descriptive, because you often cannot identify a given customer's segment at
the moment you want to act. Behavioural segments are usually actionable,
because the behaviour is in your own data and the trigger is observable.

Use demographics to *characterise* a behavioural segment, not to define it. The
right sequence is: segment on behaviour, then profile each segment
demographically so the client can picture them and brief a creative agency.

### Pareto, tested rather than assumed

"80% of revenue comes from 20% of customers" is repeated as if it were a law.
It is an empirical regularity that holds in some businesses and not others, and
it is trivial to test with a running-total share. In AdventureWorks internet
sales the concentration is notably *weaker* than 80/20, because the business is
dominated by single-purchase bike buyers of similar ticket size. In reseller
sales it is far stronger. Reporting the actual curve, and noting that the two
channels differ sharply, is more useful than invoking the slogan — and catching
the difference is the kind of finding that earns a consultant credibility.

### Cohorts

A cohort groups customers by when they first bought, then tracks each group
forward over time. The grid has first-purchase month down the rows and
months-since-first-purchase across the columns.

Read it two ways and the two readings answer different questions:

- **Down a column** compares cohorts at the same age. If month-3 retention falls
  steadily for later cohorts, acquisition quality is deteriorating — the business
  is buying worse customers over time.
- **Across a row** follows one cohort's lifecycle: how fast it decays, and
  whether it flattens into a stable base or goes to zero.

The triangular shape is structural: recent cohorts have had less time to age, so
the bottom-right of the grid is empty. Never average down a column across
cohorts of different ages, and never read the final partial month as a decline.

## Walkthrough

### Step 1 — The RFM base table

Work on a sheet `RFM`, one row per customer, with `DimCustomer` loaded as
`Customers`, `FactInternetSales` as `Sales` and `DimDate` as `Calendar`. Put the
as-of date in a single labelled input cell — say `$B$1` holding `1-Jan-2014` —
so it is visible and changeable in one place.

First get each customer's last order date. `OrderDateKey` is an integer, so
take the maximum key and convert it:

```
B2:  =MAXIFS(Sales[OrderDateKey], Sales[CustomerKey], [@CustomerKey])
C2:  =DATE(LEFT([@LastOrderKey],4), MID([@LastOrderKey],5,2), RIGHT([@LastOrderKey],2))
```

`LEFT` on a number coerces it to text, which is exactly what you want here; the
`DATE` function then rebuilds a real serial date. Now the three measures:

```
D2:  =$B$1 - [@LastOrderDate]
E2:  =SUMPRODUCT((Sales[CustomerKey]=[@CustomerKey])
       / COUNTIFS(Sales[SalesOrderNumber], Sales[SalesOrderNumber]))
F2:  =SUMIFS(Sales[SalesAmount], Sales[CustomerKey], [@CustomerKey])
G2:  =SUMIFS(Sales[SalesAmount], Sales[CustomerKey], [@CustomerKey])
     - SUMIFS(Sales[TotalProductCost], Sales[CustomerKey], [@CustomerKey])
```

D is Recency in days. E is Frequency as **distinct orders**, not line count —
the `SUMPRODUCT` divides each matching row by how many rows share its order
number, so a three-line order contributes 1. This is the single most common RFM
error: `COUNTIFS` on `CustomerKey` counts lines and inflates Frequency for
anyone who bought a bike with two accessories. F is Monetary. G is margin,
which is the better basis for "how much may we spend to keep them" and which
Chapter 1's definitions discipline requires you to derive consistently.

On 60,000 fact rows the `SUMPRODUCT` form is slow. For anything you will
refresh, use a Power Pivot measure — `DISTINCTCOUNT(Sales[SalesOrderNumber])`,
as in Chapter 8 — and bring it in with `CUBEVALUE` or a pivot. Build it the slow
way once to understand it, then switch.

### Step 2 — Score into quintiles

```
I2:  =6 - CEILING(PERCENTRANK.INC(RFM[Recency], [@Recency]) * 5, 1)
J2:  =MAX(1, CEILING(PERCENTRANK.INC(RFM[Frequency], [@Frequency]) * 5, 1))
K2:  =MAX(1, CEILING(PERCENTRANK.INC(RFM[Monetary], [@Monetary]) * 5, 1))
```

Read the Recency line carefully: because low Recency is good, the score is
inverted with `6 -`. Get this wrong and every subsequent recommendation points
at the wrong customers. Check it immediately — sort by Recency ascending and
confirm the top rows carry score 5.

The `MAX(1, ...)` guards the lowest value, where `PERCENTRANK` returns exactly 0
and `CEILING` would give 0 rather than 1.

Verify the bands with a quick count:

`=COUNTIFS(RFM[F_Score], 5)`

For Frequency you will find the counts are badly unbalanced, because a large
majority of customers have exactly one order and ties cannot be split. Do not
fight this. It is the finding of Step 5 arriving early.

As an alternative, `QUARTILE.INC` gives explicit cut points you can display,
which some clients prefer because the boundaries are then auditable:

`=QUARTILE.INC(RFM[Monetary], 3)`

### Step 3 — Name the segments

Combine the scores into a label with rules, in priority order so each customer
lands in exactly one segment:

```
M2:  =IFS(
       AND([@R_Score]>=4, [@F_Score]>=4, [@M_Score]>=4), "Champions",
       AND([@R_Score]>=3, [@M_Score]>=4), "Loyal high-value",
       AND([@R_Score]<=2, [@M_Score]>=4), "At risk - high value",
       AND([@R_Score]<=2, [@M_Score]<=2), "Hibernating",
       [@R_Score]>=4,                     "Recent low-value",
       TRUE,                              "Everyone else")
```

`IFS` evaluates in order and stops at the first match, which is what gives you
mutually exclusive segments. Put "At risk - high value" above the generic
low-recency rule deliberately: it is the segment with money on the table, and
rule order is a business decision, not a formatting detail.

Six segments is near the practical limit. Clients act on three or four.

### Step 4 — Size the segments in money

This is the step that turns a chart into an analysis. Build a summary table,
one row per segment:

```
=COUNTIFS(RFM[Segment], [@Segment])
=SUMIFS(RFM[Monetary], RFM[Segment], [@Segment])
=SUMIFS(RFM[Margin], RFM[Segment], [@Segment])
=SUMIFS(RFM[Monetary], RFM[Segment], [@Segment]) / SUM(RFM[Monetary])
=AVERAGEIFS(RFM[Monetary], RFM[Segment], [@Segment])
=AVERAGEIFS(RFM[Recency], RFM[Segment], [@Segment])
```

Sort the table by revenue share descending, not by headcount and not
alphabetically. The ordering is an argument: it says "here is where the money
is". A segment holding 6% of customers and 22% of margin is the headline, and
sorting by headcount would bury it in the middle of the table.

Then profile each segment demographically so it becomes briefable — average
income, modal occupation, modal commute band — using `AVERAGEIFS` and a mode
over text. That profile is characterisation, not definition. The segment is
still defined by behaviour.

### Step 5 — The cohort grid

Assign each customer a cohort month from `DateFirstPurchase`:

`=EOMONTH(Customers[@DateFirstPurchase], 0)`

Using month-end as the cohort label keeps it a real date, so it sorts and
formats correctly — a text label like "2012-07" will sort fine but cannot be
used in date arithmetic later.

For each fact row, compute the months elapsed between the customer's cohort
month and the order month. Add a helper column to the sales table:

`=DATEDIF(cohort_month_of_this_customer, order_month_of_this_row, "m")`

Then build the grid with cohort months down column A and period numbers 0, 1,
2 … across row 1:

`=SUMIFS(Sales[SalesAmount], Sales[CohortMonth], $A2, Sales[PeriodIndex], B$1)`

The mixed references `$A2` and `B$1` let one formula fill the whole rectangle.
For a retention grid rather than a revenue grid, count distinct customers active
in each period and divide by the cohort's size at period 0 — a pivot with
`DISTINCTCOUNT` is cleaner than a worksheet formula for this.

Apply a colour scale across the grid, but set its minimum and maximum manually.
Conditional formatting scaled per row will make every cohort look identical by
construction, which is precisely the comparison you are trying to make.

Then exclude the final partial month. It will otherwise appear as a cliff in the
last column of every row, and someone will read it as collapsing retention.

### Step 6 — Report the single-purchase finding honestly

When you build this grid on AdventureWorks internet sales, most cohorts will
show very little activity after period 0. The great majority of internet
customers buy once.

The temptation is to bury this, because a retention analysis that finds no
retention feels like a failed analysis. It is the opposite. Write it plainly:

> Across the four-year window, the large majority of AdventureWorks internet
> customers placed exactly one order. Cohort revenue after the first month of
> purchase is a small fraction of first-month revenue, consistently across
> cohorts. The internet channel is operating as a one-time-purchase business, not
> a repeat-purchase one. A retention budget has very little existing repeat
> behaviour to amplify; the prior question is whether repeat purchase is
> achievable at all in this category and at this price point.

Then make the structural point that rescues the finding: a customer buying a
$3,000 road bike is not expected to buy another next quarter. The repeat
opportunity is not another bike, it is Accessories and Clothing against an
installed base of bike owners. That reframing turns a null result into a
specific, testable recommendation — and it is visible in the data, because you
can check how many bike buyers ever return for an accessory with a
`COUNTIFS` over product category.

### Step 7 — Pareto, computed

Sort customers by Monetary descending, then:

```
=SUM($F$2:[@Monetary]) / SUM(RFM[Monetary])
=[@CumulativeShare] - ROW() / COUNTA(RFM[CustomerKey])
```

The first is the running revenue share; the expanding range `$F$2:[@Monetary]`
is the one place in this book where a non-structured reference is the right
choice, because the cumulation genuinely depends on row position. Find the
customer share at which cumulative revenue crosses 80%:

`=MATCH(0.8, RFM[CumulativeShare], 1) / COUNTA(RFM[CustomerKey])`

Report the real number. In internet sales it will be well above 20% —
concentration is weak. Run the same calculation on `FactResellerSales` by
`ResellerKey` and you will find concentration far stronger. Two channels, two
different structures, one slide: that contrast is the finding, and it has a
direct implication, since key-account management is justified in one channel
and not in the other.

Do the same over `ProductKey` for an ABC classification — A for products up to
70% of cumulative revenue, B to 90%, C the tail — and you have the basis for the
inventory conversation.

### Step 8 — Clustering, and its limits

k-means asks a different question from RFM: instead of imposing dimensions you
chose, it finds groups that are naturally close together in whatever space you
give it. The algorithm is simple — assign each point to the nearest of k
centroids, recompute the centroids, repeat.

You can implement it in Excel. Put k centroid rows in a block, compute each
customer's squared distance to each centroid with `SUMXMY2`, assign with
`MATCH(MIN(...))`, and ask Solver to minimise total within-cluster distance.

You should treat that as a demonstration, not a deliverable. It is slow on
18,500 rows, Solver finds a local optimum that depends on where you started, it
is awkward to re-run, and nothing records which solution you accepted. The
honest framing for a client is: "I can show you the logic in Excel so you can
see there is no magic in it; for a production segmentation we should run it in
Python, where it takes seconds and is reproducible."

Standardise before clustering in any case — `=([@Monetary]-AVERAGE(...))/STDEV.P(...)` —
or Monetary in dollars will dominate Recency in days purely through units. And
remember that k-means on unscaled, skewed spend data mostly rediscovers "big
spenders versus everyone else", which `QUARTILE` already told you.

### Step 9 — Resellers, the B2B counterpart

Reseller segmentation is a different exercise because there are around 700
resellers rather than 18,500 customers, and they carry roughly three quarters
of company revenue. Each one is individually material, so you can name them.

Segment on `BusinessType` — Warehouse, Value Added Reseller, Specialty Bike
Shop — crossed with actual revenue decile:

`=SUMIFS(ResellerSales[SalesAmount], ResellerSales[ResellerKey], [@ResellerKey])`

Note the trap flagged in the data dictionary: `DimReseller[AnnualSales]` is a
dimension attribute that disagrees with summed fact revenue. Segment on the
fact, mention the discrepancy, and if the gap is large, flag it as a data-quality
issue for the client's CRM rather than silently picking one.

The actions differ by segment in a way that the B2C segments do not: a top-decile
Warehouse account justifies a named account manager, while a bottom-decile
Specialty shop justifies a self-service portal and nothing else. Those are
genuinely different decisions, which is the test from the start of the chapter.

## Common pitfalls

- **Recency scored the wrong way round.** The single most damaging error here;
  it inverts every recommendation. Verify by sorting.
- **Frequency as line count.** `COUNTIFS` on `CustomerKey` counts order lines.
  Distinct `SalesOrderNumber` is the measure you want.
- **`TODAY()` as the as-of date.** Every figure in the deck drifts. Fix it in a
  labelled cell.
- **Subtracting `OrderDateKey` values.** `20130715 − 20130630` is not 15 days.
  Convert to a real date first.
- **Sizing segments by headcount.** The point of segmentation is to allocate
  money. Sort by revenue or margin share.
- **Overlapping segment rules.** Without ordered `IFS` logic customers land in
  two segments and your totals exceed 100%.
- **Averaging down a cohort column across cohorts of different ages.** Younger
  cohorts have not had time to accumulate; the average is meaningless.
- **Reading the last partial month as a collapse.** Exclude it.
- **Per-row conditional formatting on the cohort grid.** It makes every cohort
  look the same and destroys the comparison.
- **Assuming 80/20.** Compute it. In this dataset the two channels differ
  sharply and the difference is the insight.
- **Clustering unstandardised variables.** Units decide the answer.
- **Segments with no attached action.** Four named boxes nobody can act on is a
  slide, not analysis.

## Exercises

1. Build the RFM base table against an as-of date of 1 January 2014. Compute
   Frequency both ways — `COUNTIFS` on `CustomerKey` and distinct
   `SalesOrderNumber` — and report what share of customers the two methods
   disagree on. Explain in two sentences why the distinct-order version is the
   one to use.
2. Score R, F and M into quintiles. Count the customers in each Frequency band
   and explain why the bands are not equal. State what that imbalance tells you
   about the business before you have built a single cohort.
3. Assign the six segments with ordered `IFS` logic. Produce the segment summary
   sorted by revenue share, including headcount, revenue, margin, revenue share
   and average recency. Identify the segment with the largest gap between its
   share of customers and its share of margin, and say what you would do about
   it.
4. Compute the Pareto curve for internet customers and for resellers. Report the
   customer share at which each crosses 80% of revenue. Write the two-sentence
   finding that contrasts the channels and names the implication for account
   management.
5. Build the cohort grid of revenue by first-purchase month and
   months-since-first-purchase, excluding the final partial month. Read it down
   one column and across one row, and write what each reading tells you.
6. Quantify the single-purchase problem: what share of internet customers placed
   exactly one order, and what share of revenue do repeat purchasers account
   for? Then test the reframing — of customers whose first order included a
   Bikes-category product, what share ever returned for an Accessories or
   Clothing product?
7. Profile your "At risk - high value" segment demographically with
   `AVERAGEIFS` and a mode over `EnglishOccupation`. State explicitly which
   parts of that profile define the segment and which merely describe it.
8. Name the one segment that should receive the client's next marketing dollar.
   Defend it with the revenue at stake, the action you propose, and the reason
   you expect that action to work on that segment specifically rather than on
   any other.

## Solutions

**Exercise 1.** The two Frequency methods diverge for every customer who bought
more than one product in a single order — a bike with a helmet and a tyre is one
order and three lines. In AdventureWorks internet sales that is a substantial
minority of customers, and the line-count version systematically promotes
multi-item single-order buyers into high Frequency bands, which then flows into
the "Loyal" and "Champions" segments. The error does not look like an error: it
produces a plausible segmentation that is wrong about who is loyal.

**Exercise 2.** The Frequency bands cannot be equal because the dimension is
dominated by a single tied value. `PERCENTRANK` assigns identical scores to
identical values, so the one-order majority all land in the same band and the
remaining bands hold whatever is left. The imbalance is diagnostic: before any
cohort work, it tells you Frequency carries almost no information in this
channel, which in turn tells you RFM will effectively reduce to RM here. Saying
that early is better than discovering it in Step 5.

**Exercise 3.** Expect the high-value segments to hold a small share of
customers and a disproportionate share of margin. The largest gap is usually in
"At risk - high value": few customers, high past spend, poor recency. That is
precisely the group where a retention intervention has the most to protect, and
it is the most common segment to be missed by a headcount-sorted table, because
it is small.

**Exercise 4.** Internet concentration is weak — the 80% crossing point sits far
above 20% of customers — because ticket sizes cluster and almost everyone buys
once. Reseller concentration is strong. The implication is direct: named
account management is justified in the reseller channel and unjustifiable in the
internet channel, where no individual customer is material. One number per
channel settles an organisational-design question.

**Exercise 5.** Down a column you are comparing cohorts at equal age, which
tests whether acquisition quality is changing over time. Across a row you are
watching one cohort decay, which tests whether there is a stable repeat base.
In this dataset the row reading is stark — activity after period 0 is slight —
and the column reading is fairly flat, meaning acquisition quality is not
deteriorating. The business has a repeat-purchase problem, not a worsening
acquisition problem. Those two diagnoses would lead to entirely different
recommendations, which is why both readings are required.

**Exercise 6.** A large majority of internet customers have exactly one order.
The reframing is the valuable part: bike buyers are not expected to repeat on
bikes, and the measurable question is cross-category return. If the share of
bike buyers who ever bought an accessory is low, the recommendation writes
itself — there is an untapped installed base, and the intervention is a
post-purchase cross-sell rather than a loyalty programme. Report the share, not
an adjective.

**Exercise 7.** Income, occupation and commute band describe the segment; they
do not define it. The definition is behavioural — low recency, high past spend.
Keeping that distinction explicit prevents the classic slide-deck error of
re-defining the segment demographically in the next meeting, at which point it
stops being reachable and the campaign cannot be targeted at all.

**Exercise 8.** The defensible answer is "At risk - high value", on three
grounds: the revenue at stake per customer is the highest of any reachable
group, the group is small enough for a budget to reach it meaningfully, and the
action — a targeted cross-category offer against a known purchase history — is
specific to what distinguishes them. An answer of "Champions" is weaker, because
those customers are already buying and the incremental return on spending there
is lowest. Whatever you choose, the mark is in the action and the sizing, not in
the segment name.

## Key takeaways

- Build a segment only if you can name a different action for it. Write the
  action down first.
- RFM uses behaviour, which Chapter 11 showed predicts far more than
  demographics. Use demographics to profile segments, never to define them.
- Fix and display the as-of date. Recency without one is not a number.
- Frequency means distinct orders. Line counts quietly promote multi-item
  buyers into your loyal segment.
- Score to classify, then size in dollars and margin. Sorting the summary by
  revenue share is itself the argument.
- Compute the Pareto curve rather than assuming 80/20 — and compare channels,
  because the contrast between them is often the finding.
- Read cohort grids down the columns and across the rows; they answer different
  questions. Drop the final partial period.
- When the data says most customers buy once, say so. A null retention finding,
  reframed as a cross-sell opportunity against an installed base, is a stronger
  deliverable than a four-box chart.
- Solver-based clustering in Excel is a teaching aid. Standardise your variables,
  show the logic, and move to Python for anything you will run twice.
- Resellers are few and individually material. Segment them on fact revenue, not
  on the dimension's self-reported `AnnualSales`.
