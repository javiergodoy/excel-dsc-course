# Chapter 10 — Time Series, Seasonality and Forecasting

## Learning objectives

By the end of this chapter you will be able to:

- Build a clean, gap-free monthly revenue series from daily AdventureWorks
  transactions, and recognise a truncated final period before it misleads you.
- State whether you used calendar or fiscal periods, and know why it changes
  the answer.
- Compute period-over-period growth, year-over-year growth and CAGR correctly,
  and choose the right one for a seasonal business.
- Smooth a series with simple and centred moving averages, and say what each
  one costs you.
- Decompose a series into trend, seasonal index and residual in worksheet
  cells, so every step is auditable.
- Forecast with `FORECAST.LINEAR`, `TREND`, `FORECAST.ETS` and the Forecast
  Sheet, and choose a horizon you can defend.
- Backtest a forecast with a holdout sample and report MAPE and RMSE.
- Separate a forecast from a target in conversation with a client, and detect
  mix shift masquerading as growth.

## The business question

The AdventureWorks board has asked for a revenue plan for the coming year. The
Head of Sales has already circulated a view: "We grew 35 per cent last year.
Call it 30 per cent next year to be prudent."

Two things are wrong with that, and neither is arithmetic. First, the growth
rate depends entirely on which periods you compare and whether you corrected
for the fact that the dataset's final month is incomplete — the raw data
contains a month that looks like an 80 per cent collapse and is nothing of the
kind. Second, and more seriously, a single aggregate growth rate hides the
mechanism. AdventureWorks' revenue has been growing because the *mix* has been
shifting toward a channel and a category with very different economics. A
growth rate extrapolated without understanding the mix shift will miss both the
revenue number and, far more expensively, the margin number.

So the real question: **what is the underlying trend in AdventureWorks revenue,
how much of the observed pattern is seasonal, and what is a defensible
twelve-month forecast with an honest uncertainty band?** And then the question
the client has not asked but needs: **is the forecast you have produced a
forecast, or is it a target?**

## Data you will use

From [the data dictionary](../data/adventureworks-data-dictionary.md):

- **`FactInternetSales`** as `Sales` — `OrderDateKey`, `SalesAmount`,
  `TotalProductCost`, `OrderQuantity`, `ProductKey`, `SalesTerritoryKey`.
- **`FactResellerSales`** as `ResellerSales` — the same columns plus
  `ResellerKey`. **Both channels are required.** Reseller sales are around three
  quarters of company turnover; a trend built on internet sales alone is a trend
  for a quarter of the business.
- **`DimDate`** as `Calendar` — `DateKey`, `FullDateAlternateKey`,
  `CalendarYear`, `MonthNumberOfYear`, `EnglishMonthName`, `CalendarQuarter`,
  `FiscalYear`, `FiscalQuarter`.
- **`DimProductCategory`** / **`DimProductSubcategory`** / **`DimProduct`** —
  for the mix analysis.

Two definitions stated once:

- **Company revenue** in this chapter is `SUM(Sales[SalesAmount]) +
  SUM(ResellerSales[SalesAmount])`, both channels, in US dollars, assigned to
  the month of `OrderDateKey`.
- A **month** is a calendar month derived from `Calendar`, not from
  `=MONTH()` applied to an integer. `OrderDateKey` is an integer in `yyyymmdd`
  form; arithmetic on it is meaningless. Join to `Calendar` and take the
  attributes from there.

## Core concepts

### A time series needs a spine, not a GROUP BY

The instinct is to pivot the fact table by month and take whatever rows come
back. That produces a series with holes in it — any month with no orders simply
does not appear — and a chart drawn on holes lies about the shape.

Build a **date spine** first: every month in the window, present whether or not
it has revenue, in order. `DimDate` already is that spine; you need its months.

```
=UNIQUE(Calendar[CalendarYear]&"-"&TEXT(Calendar[MonthNumberOfYear],"00"))
```

Then attach revenue to the spine with `SUMIFS`, which returns zero rather than
nothing for an empty month:

```
=SUMIFS(Sales[SalesAmount],Sales[OrderDateKey],">="&$B2,Sales[OrderDateKey],"<="&$C2)
 +SUMIFS(ResellerSales[SalesAmount],ResellerSales[OrderDateKey],">="&$B2,
         ResellerSales[OrderDateKey],"<="&$C2)
```

where `$B2` and `$C2` hold the first and last `DateKey` of the month as
integers — for example 20130701 and 20130731. Working in `DateKey` integers
rather than converting to dates avoids an entire class of regional
date-interpretation bugs, and the `yyyymmdd` format sorts correctly as a
number, which is precisely why the warehouse uses it.

A zero in a time series means "this month happened and sold nothing". A missing
row means "we lost this month". They look identical on a chart and mean
opposite things.

### The truncated final period

Plot AdventureWorks company revenue monthly and the last month falls off a
cliff. In the 2019 release the sales window ends in January 2014 with only a few
days of orders recorded. The chart shows a catastrophic collapse. There was no
collapse; there is no data.

This is the single most common error in business time-series work, and it does
not only happen with sample datasets. It happens every time an analyst pulls a
live extract mid-month and plots it next to complete months.

Three defences, in order of preference:

1. **Drop the incomplete period** and label the chart with the window you
   actually used: "monthly company revenue, Jan 2011 – Dec 2013 (complete
   months only)".
2. **Annualise or run-rate it** and mark it visually as an estimate — a dashed
   line segment, a hollow marker, a footnote. Never a solid bar identical to the
   complete months.
3. **Use a trailing-twelve-month series** instead of monthly, which absorbs the
   problem and is usually what a board wants anyway.

Detect it programmatically rather than by eye, because by eye you will miss it
on a busy dashboard:

```
=MAX(Sales[OrderDateKey])
=COUNTIFS(Sales[OrderDateKey],">="&20140101,Sales[OrderDateKey],"<="&20140131)
```

Compare the second figure to the median monthly order count. If it is a small
fraction, flag the month.

### Calendar versus fiscal, and saying which

`DimDate` carries both `CalendarYear` and `FiscalYear`, and the AdventureWorks
fiscal year starts in July. So fiscal 2013 runs July 2012 to June 2013. A chart
labelled "2013" means two different twelve-month windows to the controller and
the sales director, and they will both be certain they are right.

```
=SUMIFS(Sales[SalesAmount],Sales[CalendarYear],2013)
=SUMIFS(Sales[SalesAmount],Sales[FiscalYear],2013)
```

Those return materially different numbers. Neither is wrong. Presenting either
without saying which is wrong.

Pick based on the audience: fiscal for anything the finance function will
reconcile against the accounts, calendar for anything compared to external
market data. Then put the basis in the chart subtitle, every time, in words:
"fiscal years, FY begins 1 July".

### Growth arithmetic, and why YoY wins for seasonal businesses

Period-over-period growth:

```
=C3/C2-1
```

Year-over-year growth at monthly grain — the same month, twelve rows back:

```
=C14/C2-1
```

For a seasonal business, month-over-month growth is close to useless. It mostly
measures the season. If AdventureWorks sells more bikes in June than in
February, a month-over-month figure of +40 per cent in spring and −30 per cent
in autumn tells you about the weather, not the business. Year-over-year
compares like with like: June against June, with the seasonal effect appearing
on both sides of the ratio and cancelling.

CAGR, two ways:

```
=(C49/C1)^(12/48)-1
=RATE(48,0,-C1,C49)*12
```

The first is the power formula: ending over beginning, raised to the
reciprocal of the number of periods, annualised. The second uses `RATE`, which
is algebraically identical for this case and sometimes easier to read in a
model. Note the sign convention — `RATE` needs the present value negative.

CAGR's weakness is that it depends on exactly two observations, the first and
the last, and ignores everything between. If either endpoint is unusual, the
CAGR is unusual. For AdventureWorks the first months of 2011 are genuinely
small — the internet channel was ramping — so a CAGR anchored on them overstates
the sustainable rate substantially. Anchor on a full fiscal year, or fit a
trend to all the points instead:

```
=SLOPE(C2:C49,A2:A49)
```

That gives dollars per month of trend growth, using all forty-eight
observations rather than two.

### Moving averages and the centring problem

A simple trailing moving average:

```
=AVERAGE(C2:C13)
```

A twelve-month trailing average removes seasonality completely, because every
twelve-month window contains each season once. That is its purpose, and it is
why trailing-twelve-month revenue is the standard board metric.

The cost is **lag**. A trailing average is centred six months in the past, so it
turns late. If you plot it against the raw series and read the crossing point
as a turning point, you are reading a turning point that happened half a year
ago.

A centred moving average fixes the lag. For an even period like twelve months
you need the 2×12 trick, because a twelve-month window has no middle month:

```
=(0.5*C2+SUM(C3:C13)+0.5*C14)/12
```

That is the standard centred twelve-month average — half weight on the first
and last months, full weight on the ten between, divided by twelve. It is
centred on the seventh row of the window and it is the correct trend estimate
for monthly decomposition. It costs you six months at each end of the series,
which with forty-eight months is a real price: you lose a quarter of your
observations to the smoother.

### Decomposition in cells

The classical multiplicative model:

> observed = trend × seasonal × residual

Multiplicative rather than additive because AdventureWorks' seasonal swing
grows with the level of the business. A $200k December uplift on a $1m base and
on a $3m base are not the same phenomenon; a 20 per cent December uplift on both
might be.

Build it in five columns and never in your head:

**Column C** — observed monthly revenue.
**Column D** — centred 12-month trend:

```
=(0.5*C2+SUM(C3:C13)+0.5*C14)/12
```

**Column E** — detrended ratio:

```
=C8/D8
```

**Column F** — seasonal index, the average ratio for that calendar month across
all years:

```
=AVERAGEIFS($E$8:$E$43,$B$8:$B$43,B8)
```

Then normalise the twelve indices so they average exactly 1, or your
seasonally adjusted series will drift:

```
=F8/AVERAGE($F$8:$F$19)*1
```

**Column G** — residual:

```
=C8/(D8*F8)
```

**Column H** — seasonally adjusted revenue, which is the series you should show
a client who asks "are we actually growing?":

```
=C8/F8
```

The discipline here is that every intermediate quantity is a visible cell.
When a partner asks "where does the 1.18 December index come from?", you point
at a column and a formula. If the decomposition lived inside a single opaque
function you would have to say "Excel did it", which is not an answer.

For AdventureWorks you should find a pronounced seasonal pattern with
northern-hemisphere summer and pre-Christmas strength, and genuine weakness in
the early calendar year. Treat the exact index values as what your data
produces, not as figures to memorise — they shift between releases.

### Fitting and projecting a trend

`FORECAST.LINEAR` and `TREND` both fit an ordinary least-squares line. `TREND`
handles arrays and multiple future points in one formula; `FORECAST.LINEAR`
takes one x at a time.

```
=FORECAST.LINEAR(50,$C$2:$C$49,$A$2:$A$49)
=TREND($C$2:$C$49,$A$2:$A$49,A50:A61)
=SLOPE($C$2:$C$49,$A$2:$A$49)
=INTERCEPT($C$2:$C$49,$A$2:$A$49)
```

Fit the trend on the **seasonally adjusted** series, then re-apply the seasonal
index to the projection. Fitting a straight line to a raw seasonal series and
projecting it gives you a forecast that is wrong in a predictable, embarrassing
way: too low in peak months, too high in troughs, by exactly the seasonal
amplitude.

```
=(INTERCEPT($H$8:$H$43,$A$8:$A$43)+SLOPE($H$8:$H$43,$A$8:$A$43)*A50)*F50
```

That is the whole forecast in one cell: adjusted trend projected forward,
multiplied back by the seasonal index for the target month. Write it as a
`LET` if you prefer it readable:

```
=LET(b,INTERCEPT($H$8:$H$43,$A$8:$A$43),
     m,SLOPE($H$8:$H$43,$A$8:$A$43),
     (b+m*A50)*F50)
```

### `FORECAST.ETS` and the Forecast Sheet

`FORECAST.ETS` implements exponential triple smoothing — AAA
error-trend-season — and does the decomposition internally.

```
=FORECAST.ETS(A50,$C$8:$C$49,$D$8:$D$49,12,1)
=FORECAST.ETS.CONFINT(A50,$C$8:$C$49,$D$8:$D$49,0.95,12,1)
=FORECAST.ETS.SEASONALITY($C$8:$C$49,$D$8:$D$49)
=FORECAST.ETS.STAT($C$8:$C$49,$D$8:$D$49,3,12,1)
```

The arguments: target date, values, timeline, seasonality (12 for monthly, or 1
to disable, or 0 to let Excel detect), and data completion. Note that the
timeline must be **dates**, not `DateKey` integers — this is the one place in the
chapter where you must convert. `FORECAST.ETS.CONFINT` returns the half-width of
the interval, so the band is forecast ± that value.

`FORECAST.ETS.SEASONALITY` is worth running before you trust anything: if Excel
detects a seasonality of 12 on your series, good; if it detects 4 or 6, your
series is doing something you have not understood yet.

**Data → Forecast Sheet** wraps all of this in a dialog and produces a chart
plus a table of forecast, lower bound and upper bound. It is genuinely useful
for a first look, and it has one disqualifying property for consulting work: the
output is values and a chart, and the method is a black box to your reader. Use
it to orient yourself, then rebuild the forecast in visible cells for anything
that goes in front of a client.

The honest comparison to make: run the cell-based decomposition forecast and
`FORECAST.ETS` side by side. If they agree within a few per cent, your
confidence in both rises and you can present the transparent one. If they
disagree materially, something in your decomposition is wrong, or the series
has structure neither method is capturing. Either way you have learned
something before the meeting rather than during it.

### Choosing a horizon you can defend

AdventureWorks gives you roughly four complete years of monthly data —
forty-eight observations, which is three or four observations of each calendar
month's seasonal behaviour.

What that supports:

- A **twelve-month forecast**: defensible. One seasonal cycle ahead, on three to
  four cycles of history.
- A **twenty-four-month forecast**: stretching. State prominently that year two
  is an extrapolation of a trend, not a forecast of a season.
- A **sixty-month forecast**: not a forecast. It is a financial model with a
  growth assumption in it, which is a legitimate thing to build — see
  Chapter 13 — but it should not be called a forecast or given a confidence
  band.

The rule of thumb: **forecast no further than a third of your history, and no
further than one seasonal cycle beyond it without flagging the change in
character.** Then say the limit out loud, because the client will otherwise
extend your line with a ruler.

The deeper point is that the confidence band from `FORECAST.ETS` quantifies
*statistical* uncertainty under the assumption that the process continues
unchanged. It does not and cannot quantify the risk that AdventureWorks opens a
new territory, loses a warehouse reseller, or changes its pricing. Those risks
dominate at a twelve-month horizon, and they belong in scenarios, not in a
band.

### Backtesting: the only evidence you have

A forecast with no backtest is an opinion with a chart attached. Hold out the
last twelve months, fit on the first thirty-six, forecast the twelve, and
measure the error.

Absolute percentage error per month:

```
=ABS(J38-C38)/C38
```

MAPE, the headline accuracy figure:

```
=AVERAGE(K38:K49)
```

RMSE, in dollars, which punishes large misses harder:

```
=SQRT(SUMXMY2(J38:J49,C38:C49)/COUNT(C38:C49))
```

Bias, which almost nobody computes and which matters most:

```
=AVERAGE(J38:J49)/AVERAGE(C38:C49)-1
```

A forecast with a 9 per cent MAPE and zero bias is useful. A forecast with a
9 per cent MAPE and a +9 per cent bias is systematically too high every single
month, which is a different and worse failure — it will be wrong in the same
direction for four consecutive quarters, and the client will notice in quarter
two.

Also compute the **naive benchmark**: last year's same month, unchanged.

```
=ABS(C26-C38)/C38
```

If your sophisticated model cannot beat the seasonal naive forecast, present
the naive one. This happens more often than forecasting textbooks admit, and
saying so is a sign of competence, not failure.

### Forecast versus target

This is the conversation that separates an analyst from a consultant.

A **forecast** is your best estimate of what will happen if nothing changes. It
should be unbiased: as likely to be too high as too low.

A **target** is a statement of intent, chosen to be achievable but demanding. It
is deliberately above the forecast, and the gap is the amount of improvement the
business is committing to find.

A **plan** is a target plus the named initiatives that close the gap, each with
an owner and an expected contribution.

When the Head of Sales says "call it 30 per cent", they are proposing a target
and calling it a forecast. The useful response is not to argue about the number.
It is to separate the two on one slide: "the trend forecast is roughly X per
cent with a band of plus or minus Y; your 30 per cent implies an additional Z
million dollars, which has to come from somewhere specific — here are the three
candidates and what each would need to deliver."

That reframing turns an argument about a number into a conversation about
initiatives, which is the conversation that was needed.

### Mix shift masquerading as growth

The last concept, and the one that will make you useful.

Total revenue can grow while every single thing you care about deteriorates.
Suppose AdventureWorks' reseller channel grows faster than internet, and
resellers carry materially larger discounts. Total revenue rises. Blended gross
margin percentage falls. Revenue per unit falls. A revenue-only forecast, even
an accurate one, will have missed the most important thing in the data.

Detect it by forecasting the components and the ratio, not just the total:

```
=SUMIFS(Sales[SalesAmount],Sales[OrderDateKey],">="&$B2,Sales[OrderDateKey],"<="&$C2)
 /C2
=(C2-D2)/C2
```

The first is internet share of company revenue by month; the second is blended
margin percentage. Plot both as thin lines beneath the revenue chart. If either
trends, your revenue forecast needs a mix assumption stated explicitly, and
your margin forecast cannot be revenue times a historical average percentage.

The general rule: **always forecast at the grain at which the economics
differ.** For AdventureWorks that means two channels and four categories at
minimum — eight series, summed — not one.

## Walkthrough

### Step 1 — Build the monthly spine

New sheet, `w_series`. Columns: `MonthIndex` (1…n), `YearMonth` label,
`FirstDateKey`, `LastDateKey`, `CalendarYear`, `MonthNumberOfYear`,
`FiscalYear`.

Generate the month bounds from `Calendar` rather than by hand:

```
=MIN(FILTER(Calendar[DateKey],(Calendar[CalendarYear]=$E2)*(Calendar[MonthNumberOfYear]=$F2)))
=MAX(FILTER(Calendar[DateKey],(Calendar[CalendarYear]=$E2)*(Calendar[MonthNumberOfYear]=$F2)))
```

Restrict the spine to months that fall inside the union of the two fact tables'
date ranges, then check the count. You should get forty-nine or fifty months
depending on release, with the last one incomplete.

### Step 2 — Attach revenue, cost and units

```
=SUMIFS(Sales[SalesAmount],Sales[OrderDateKey],">="&$C2,Sales[OrderDateKey],"<="&$D2)
 +SUMIFS(ResellerSales[SalesAmount],ResellerSales[OrderDateKey],">="&$C2,
         ResellerSales[OrderDateKey],"<="&$D2)
```

Do the same for `TotalProductCost` and `OrderQuantity`. Add the reconciliation
cell immediately:

```
=SUM(Series[Revenue])-SUM(Sales[SalesAmount])-SUM(ResellerSales[SalesAmount])
```

Zero, or stop and fix it. If it is non-zero, the usual cause is months in the
fact tables that fall outside your spine — which is exactly the gap the spine
exists to reveal.

### Step 3 — Find and handle the truncated month

```
=COUNTIFS(Sales[OrderDateKey],">="&$C2,Sales[OrderDateKey],"<="&$D2)
 +COUNTIFS(ResellerSales[OrderDateKey],">="&$C2,ResellerSales[OrderDateKey],"<="&$D2)
=IF(I2<0.4*MEDIAN($I$2:$I$50),"INCOMPLETE","complete")
```

That flag column is the thing you will be glad of in six weeks. Filter the
series to complete months for all analysis, and keep the incomplete row visible
in the working sheet with its flag so nobody re-adds it.

### Step 4 — Growth table

Build a small block, labelled with its basis:

| Measure | Formula |
|---|---|
| Revenue by calendar year | `=SUMIFS(Series[Revenue],Series[CalendarYear],2013)` |
| Revenue by fiscal year | `=SUMIFS(Series[Revenue],Series[FiscalYear],2013)` |
| YoY by month | `=C14/C2-1` |
| CAGR, monthly series | `=(C49/C1)^(12/48)-1` |
| Trend slope, $/month | `=SLOPE($C$2:$C$49,$A$2:$A$49)` |

Put calendar and fiscal side by side deliberately. The difference between them,
on the same page, is the fastest way to make a client understand why you keep
asking which basis they want.

### Step 5 — Decompose

Follow the five columns from the concepts section: centred trend, detrended
ratio, seasonal index by month, normalised index, residual, seasonally adjusted
revenue. Then run two checks.

Indices average to one:

```
=AVERAGE(SeasonalIndices)-1
```

Residuals have no structure:

```
=CORREL(G9:G43,G8:G42)
```

A residual autocorrelation near zero means the decomposition captured the
structure. A strongly positive one means there is a pattern left over — often a
level shift, a promotion, or a territory launch — and you should find out what
it was before forecasting through it.

### Step 6 — Forecast twelve months, two ways

Extend the spine twelve months. Build the transparent forecast:

```
=LET(b,INTERCEPT($H$8:$H$43,$A$8:$A$43),
     m,SLOPE($H$8:$H$43,$A$8:$A$43),
     (b+m*A50)*F50)
```

Build the ETS forecast and band alongside, on a date timeline:

```
=FORECAST.ETS(DateCol50,$C$8:$C$49,$DateCol$8:$DateCol$49,12,1)
=FORECAST.ETS.CONFINT(DateCol50,$C$8:$C$49,$DateCol$8:$DateCol$49,0.95,12,1)
```

Chart: actuals as a solid line, forecast as a dashed line in the same colour,
the band as a light shaded area. Do not use a different colour for the
forecast — it is the same quantity, and colour should encode the series, not
your confidence. Dashing encodes the confidence.

### Step 7 — Backtest

Refit on months 1–36 and forecast 37–48, which you already know. Compute MAPE,
RMSE, bias and the seasonal-naive MAPE as laid out above. Put all four in a
labelled block directly under the chart.

Then write the sentence. Not a paragraph — the sentence:

> "On a twelve-month holdout the seasonally adjusted trend model achieved a
> MAPE of roughly 11 per cent with a small positive bias, against 16 per cent
> for a seasonal-naive benchmark; we therefore present the trend model, and
> treat month-level figures as indicative and quarter-level figures as
> plannable."

That last clause is the valuable one. It tells the client what resolution the
forecast actually supports, which stops them from managing to a monthly number
the model cannot deliver.

### Step 8 — State the assumptions in writing

A text box, on the exhibit, not in an appendix. Six lines:

1. Both channels included; US dollars; order date basis, not ship date.
2. Calendar months; the final partial month is excluded and the window is
   stated.
3. Multiplicative seasonality, indices from three complete years, normalised to
   one.
4. Trend fitted by OLS on the seasonally adjusted series over the full
   complete-month window.
5. No change assumed in pricing, territory footprint, product range or the
   reseller base.
6. The band is statistical uncertainty under those assumptions; it excludes
   business risk, which is handled in the scenarios.

That list is what makes the number defensible. Without it the forecast is
a line on a chart that someone will quote in a board paper for two years,
stripped of every caveat you gave verbally.

## Common pitfalls

**Doing arithmetic on `DateKey`.** `20130715 − 20130630` is 85, not 15. Join to
`Calendar` and use real attributes.

**Plotting the incomplete final month as if it were complete.** It looks like a
collapse. It is an artefact. Build the detection formula once and keep it.

**Internet sales only.** Reseller sales are the majority of turnover. A trend on
`FactInternetSales` alone is a trend for a minority of the business, and it has
a different shape.

**Mixing calendar and fiscal within one exhibit.** It happens when one
`SUMIFS` uses `CalendarYear` and the next uses `FiscalYear` because you copied a
formula. The totals will not reconcile and you will lose an hour.

**Month-over-month growth on a seasonal series.** You are measuring the
calendar. Use year-over-year, or seasonally adjust first.

**CAGR anchored on a ramp-up period.** AdventureWorks' earliest months are
small because the channel was launching. A CAGR from there flatters the
business materially. Anchor on complete years or fit a trend to all points.

**Fitting a trend to the raw seasonal series.** The projection will be wrong by
the seasonal amplitude in every month. Adjust, fit, then re-seasonalise.

**Forgetting to normalise the seasonal indices.** If the twelve indices average
1.04, your seasonally adjusted series has a 4 per cent artificial trend baked
into it and your forecast inherits it.

**Reading a trailing moving average as current.** A trailing 12-month average is
centred six months ago. For turning points use the centred average and accept
the lost tail.

**Presenting a Forecast Sheet output as the deliverable.** It is values and a
black box. Rebuild in cells.

**Quoting a confidence band as total uncertainty.** It covers statistical
variation under the no-change assumption. It does not cover losing a major
reseller, which is the thing most likely to make you wrong.

**Forecasting the total when the economics differ by component.** Mix shift will
make an accurate revenue forecast useless for margin. Forecast channels and
categories, then sum.

## Exercises

1. **Build the clean series.** Construct the monthly company revenue series
   from both fact tables on the `DimDate` spine, with the reconciliation cell
   and the completeness flag. Report the number of months, the number flagged
   incomplete, and the reconciliation value. State in one sentence what would
   have gone wrong had you pivoted the fact tables directly instead.

2. **Calendar versus fiscal.** Produce company revenue and year-over-year growth
   on both bases for every year in the window. Present them as one table.
   Identify the year where the two bases disagree most about the growth rate,
   and explain mechanically why.

3. **Decompose.** Build the full five-column decomposition. Report the twelve
   normalised seasonal indices, the proof that they average one, and the
   residual autocorrelation. Name the two strongest and two weakest months and
   give a plausible business explanation for each, consistent with
   AdventureWorks selling bicycles across three hemispheric groups.

4. **Forecast and band.** Produce a twelve-month forecast two ways — the
   cell-based seasonally adjusted trend, and `FORECAST.ETS` with a 95 per cent
   band. Chart both against actuals. Report the annual total each method
   implies and the percentage difference between them. Which would you present,
   and what single sentence justifies the choice?

5. **Backtest.** Hold out the final twelve complete months. Fit on the earlier
   period and compute MAPE, RMSE and bias for your model and for a
   seasonal-naive benchmark. Report all six figures. If your model does not beat
   naive, say so explicitly and state what you would present instead.

6. **Mix shift.** Compute internet share of company revenue and blended gross
   margin percentage by month across the window. Does either trend materially?
   Produce a revenue forecast built as two channel forecasts summed, and compare
   its implied margin to the margin you would get by applying the historical
   average margin percentage to a single total-revenue forecast. Quantify the
   gap in dollars and write the two sentences you would say to the CFO about it.

## Solutions

**1.** Expect forty-nine or fifty months depending on release, one flagged
incomplete, and a reconciliation of exactly zero. Pivoting the fact tables
directly would have produced a series missing any zero-revenue month and, more
importantly, would have silently included the truncated final month as a
complete observation — which both distorts the chart and biases every trend fit
downward, because OLS weights the last point like any other.

**2.** Calendar and fiscal totals differ by the amount of revenue in the
half-year offset, which for a seasonal business is substantial. The largest
disagreement in growth rate falls where a strong half-year sits on opposite
sides of the two year boundaries — typically the year containing the sharpest
acceleration. The mechanism is arithmetic, not mysterious: fiscal 2013 contains
July–December 2012, so a strong second half of calendar 2012 lifts fiscal 2013's
*base* while lifting calendar 2012's *comparator*, pushing the two growth rates
in opposite directions.

**3.** Indices must average 1.000 after normalisation; if `=AVERAGE(indices)-1`
is not essentially zero you have normalised against the wrong range. Expect
pronounced strength in the northern summer and in the run-up to December, and
weakness in the early calendar year. Residual autocorrelation should be small; a
strongly positive value typically traces to a territory launch or a promotional
period, which is a level shift the multiplicative model cannot absorb. The
business explanation: cycling is a warm-weather purchase in the northern
hemisphere, which dominates the revenue base, and the Pacific group's opposite
season partially offsets but does not cancel it — which is itself a finding
worth a sentence.

**4.** The two methods will typically land within five to ten per cent of each
other on the annual total. Present the cell-based decomposition, because the
client can audit every step of it, and quote the ETS figure as corroboration:
"an independent exponential-smoothing model gives a figure within X per cent,
which raises our confidence in both." If they diverge by more than about fifteen
per cent, do not pick one — go back and find out why, because one of them is
mis-specified and presenting either is a coin flip.

**5.** A MAPE in the region of ten to fifteen per cent at monthly grain is a
realistic result for this series, with the seasonal-naive benchmark somewhat
worse. Report bias separately and prominently: a model with low MAPE and
persistent positive bias will be wrong in the same direction every quarter,
which destroys credibility faster than a larger but unbiased error. If naive
wins, present naive and say why — "with four years of history and a stable
seasonal pattern, last year's same month plus the trend is as good as anything
more elaborate, so we are using it and spending our time on the mix question
instead." That answer demonstrates judgment; defending a losing model does not.

**6.** You should find internet share rising across the window and blended
margin percentage moving with the mix, because the two channels carry different
discount levels. Applying a historical average margin percentage to a single
revenue forecast will therefore be wrong, and the gap compounds over twelve
months into a material dollar figure. To the CFO: "Revenue grows roughly X per
cent on our forecast, but the channel mix keeps shifting, so blended margin
drifts by Y points over the plan year — applying last year's average margin to
next year's revenue overstates gross profit by about Z million. We have built
the forecast by channel so the margin falls out of the mix rather than being
assumed."

## Key takeaways

- **Build a date spine before you build a series.** A missing month and a
  zero month look identical on a chart and mean opposite things.
- **Find the truncated period programmatically.** The last partial month is the
  most reliable way to be wrong about a trend, and it is detectable in one
  formula.
- **Say calendar or fiscal, every time, in the subtitle.** Two defensible bases
  give two different numbers, and the argument that follows is pure waste.
- **Year-over-year for seasonal businesses; CAGR with care.** CAGR uses two
  points and inherits the oddness of both. A trend slope uses all of them.
- **Decompose in visible cells.** Trend, index, residual, adjusted series — five
  columns you can point at under questioning. Normalise the indices.
- **Adjust, fit, re-seasonalise.** Never fit a line to a raw seasonal series.
- **Forecast no further than a third of your history**, and no further than one
  seasonal cycle ahead without flagging the change in character.
- **A forecast with no backtest is an opinion.** Report MAPE, RMSE, bias and the
  naive benchmark, and present the naive model when it wins.
- **Separate forecast, target and plan explicitly.** The gap between forecast
  and target is the work the business is committing to do, and naming it turns
  an argument into a plan.
- **Forecast at the grain where the economics differ.** Mix shift makes an
  accurate revenue forecast useless for margin, and margin is what the board
  actually cares about.
