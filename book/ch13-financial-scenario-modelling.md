# Chapter 13 — Financial and Scenario Modelling

Up to this point the book has been about describing what happened. This chapter
is about the harder half of consulting work: saying what will happen if the
client does something, and being able to defend that number when a finance
director pushes back on it.

A model is not a forecast. A forecast is a number. A model is an argument about
how a business works, written in arithmetic, with the debatable parts exposed
so that someone can disagree with an assumption without having to attack the
whole structure. That distinction is the point of the chapter. If your client
can only say "I don't believe 18% growth", you have built a spreadsheet. If
they can say "I don't believe accessories attach rate goes from 0.9 to 1.3
units per bike", you have built a model.

## Learning objectives

By the end of this chapter you will be able to:

- Lay out a model so inputs, calculations and outputs are physically separate,
  and explain to a reviewer why that layout is not cosmetic.
- Build a driver-based revenue forecast from volume, price and mix instead of
  applying a single growth rate to last year's total.
- Decompose an actual revenue change into price, volume and mix effects that
  reconcile exactly to the total, with no residual line.
- Calculate contribution margin and operating leverage, and use them to say
  what a given revenue change does to profit.
- Build a scenario switch with `CHOOSE` or `INDEX` and explain why it beats
  Scenario Manager for work that someone else will review.
- Run one- and two-variable Data Tables, and turn a set of one-variable tables
  into a tornado chart.
- Use Goal Seek for break-even questions and Solver for a constrained
  allocation.
- Appraise an investment with `NPV`, `IRR`, `XNPV` and `XIRR`, and hold a
  sensible conversation about the discount rate.
- Recognise a circular reference as a design decision rather than an accident,
  and know when iterative calculation is acceptable.

## The business question

AdventureWorks management has approved a growth investment and wants to know
where to put it. The specific brief:

> We have $4.0 million of incremental marketing and sales investment available
> over the next three years. Build us a three-year revenue and margin
> projection. Tell us what has to be true for it to work, how sensitive the
> answer is to price and volume, and whether the investment clears our hurdle
> rate. Also explain last year's revenue change — the board was told volume
> was up, and margin still fell, and nobody has explained why.

Four distinct deliverables are hiding in that paragraph, and they come in a
deliberate order:

1. **A diagnosis.** Why did margin fall while volume rose? This is a
   price-volume-mix question, and it has to come first, because the drivers it
   exposes are the drivers the forecast should use.
2. **A driver-based projection.** Three years, by category, built up from
   things management can actually influence.
3. **A sensitivity view.** Which assumptions matter, and how wrong can we be
   before the recommendation flips.
4. **An investment appraisal.** Does the $4.0m clear the hurdle rate, and over
   what horizon.

Notice that the diagnosis feeds the forecast. Analysts who skip step 1 build
step 2 on a growth rate, and then cannot answer the one question the board
actually asked.

## Data you will use

From [the data dictionary](../data/adventureworks-data-dictionary.md):

- **`FactInternetSales`** and **`FactResellerSales`**, unioned into a single
  `Sales` table as built in Chapter 7. Columns used: `SalesAmount`,
  `OrderQuantity`, `TotalProductCost`, `DiscountAmount`, `ProductKey`,
  `OrderDateKey`, `SalesTerritoryKey`, and the `Channel` column you added
  during the union.
- **`DimProduct`** → **`DimProductSubcategory`** → **`DimProductCategory`**,
  for the four categories: Bikes, Components, Clothing, Accessories.
- **`DimDate`**, for `CalendarYear`. Use calendar year throughout this chapter
  and say so on the output sheet — AdventureWorks' fiscal year starts in July,
  and a projection labelled "Year 1" that silently mixes the two conventions is
  a reconciliation problem waiting to happen.
- **`DimSalesTerritory`**, for the Solver allocation. Exclude the `NA`
  territory (key 11) explicitly.
- **`FactFinance`** with `DimAccount`, `DimScenario` and `DimDepartmentGroup`,
  for the fixed-cost base in the three-statement-lite projection. Scenario 1 is
  Actual.

Two definitions, fixed for the whole chapter:

- **Revenue** = `SalesAmount` (already net of `DiscountAmount` in
  AdventureWorks; check this on your release before you assume it).
- **Gross margin** = `SalesAmount − TotalProductCost`. There is no margin
  column; you derive it, which means you must define it once and reference that
  definition everywhere.

The last full calendar year in the dataset is the comparison year. The final
partial month in the data is truncated, as the dictionary warns, so a projection
anchored on a trailing-twelve-month figure that runs into that stub will start
from a base that is too low. Anchor on complete calendar years.

## Core concepts

### The three-sheet rule

Every model in this chapter obeys one structural rule:

| Sheet | Contains | Never contains |
|---|---|---|
| `Inputs` | every assumption, hard-coded, blue font | any formula |
| `Calc` | all the arithmetic, black font | any hard-coded number |
| `Output` | exhibits and summary tables, green links | any new calculation |

This is not tidiness for its own sake. It buys you three specific things.

First, **auditability**. A reviewer can read `Inputs` in two minutes and know
every number they are being asked to accept. If an assumption is embedded in
`=Calc!D14*1.04`, the reviewer has to read the whole model to find out what
they are agreeing to, and so they will not read it at all, and they will sign
off on something they do not understand. That is how models cause damage.

Second, **a single point of change**. Scenario work means changing assumptions
repeatedly. If the growth rate appears in fourteen formulas, it is not an
assumption, it is fourteen assumptions that happen to agree today.

Third, **honest sensitivity**. A Data Table can only vary a cell. If the thing
you want to vary is not in a cell, you cannot test it.

The one hard-coded number allowed in a calculation is a true constant — 12
months in a year, 365 days. Everything else goes on `Inputs`.

### Driver-based build-up, not a growth rate

Here is the lazy forecast:

```
Revenue_Year1 = Revenue_Year0 * (1 + growth_rate)
```

It has one assumption, and that assumption is unfalsifiable. Nobody can tell
you whether 8% is right, so the conversation about it is a conversation about
nothing.

Here is the same forecast, driver-based, per category:

```
Revenue = Units × Average price
Units   = Customers × Orders per customer × Units per order
Margin  = Revenue − Units × Unit cost
```

Now the assumptions are about things with owners. "Orders per customer" belongs
to the CRM team. "Units per order" is the attach-rate question the accessories
business cares about. "Average price" belongs to pricing. A category manager
can look at `Inputs` and say "our attach rate is nowhere near 1.3" — which is a
useful argument, and the kind of argument you want to provoke.

For AdventureWorks the natural driver tree differs by channel, and that is a
feature, not an inconsistency:

- **Internet sales**: active customers × orders per customer × units per order ×
  price per unit. `OrderQuantity` is almost always 1, so units per order is
  effectively basket size in lines.
- **Reseller sales**: active resellers × orders per reseller × units per order ×
  price per unit × (1 − discount). Reseller discounts are large and are a
  genuine lever, so model them explicitly rather than burying them in net price.

### Price-volume-mix, with no residual

This is the technique that answers the board's question, and it is worth being
precise, because most spreadsheet versions of it leave an unexplained residual
line and the analyst calls it "other".

Let category $i$ have base-year units $q_{0i}$ and price $p_{0i}$, and
current-year units $q_{1i}$ and price $p_{1i}$. Total units $Q_0 = \sum q_{0i}$,
and the blended base price $\bar p_0 = R_0 / Q_0$.

Three effects:

- **Volume**: $(Q_1 - Q_0) \times \bar p_0$ — what revenue would have changed
  by if you sold more units at the old average price with the old mix.
- **Mix**: $Q_1 \times \sum (m_{1i} - m_{0i}) \times p_{0i}$, where
  $m_i = q_i / Q$ — the effect of selling a different *blend* of products at
  old prices.
- **Price**: $Q_1 \times \sum m_{1i} \times (p_{1i} - p_{0i})$ — the effect of
  charging different prices on the new blend.

These three sum exactly to $R_1 - R_0$. Not approximately. If your three
effects do not reconcile to the penny, you have made an arithmetic error, and
the reconciliation check is therefore the most valuable cell in the schedule.

The AdventureWorks answer to "volume up, margin down" is almost always mix. The
bike categories carry high prices and thin percentage margins; accessories
carry low prices and fat percentage margins. Shift units toward accessories and
unit volume rises while revenue per unit collapses. Shift units toward bikes and
revenue rises while percentage margin falls. Either direction can produce a
board paper that reads like a contradiction, and the mix effect is what resolves
it.

### Contribution margin and operating leverage

Gross margin is a statement about products. Contribution margin is a statement
about decisions: revenue minus the costs that actually vary with the next unit
sold. For AdventureWorks, `TotalProductCost` is the variable block; the fixed
block comes from `FactFinance`.

Operating leverage is the multiplier between a revenue change and a profit
change:

$$\text{Operating leverage} = \frac{\text{Contribution margin}}{\text{Operating profit}}$$

A business with contribution of $30m and operating profit of $10m has leverage
of 3.0: a 10% revenue increase lifts profit by roughly 30%, and a 10% decrease
cuts it by roughly 30%. The symmetry is the part clients forget, and it is
usually the most important sentence in the risk section of the deck.

### Scenario switch beats Scenario Manager

Excel's Scenario Manager stores sets of input values and writes them into cells
on demand. It works, and I avoid it on client work for three reasons: the stored
values are invisible on the sheet, so a reviewer cannot see what Downside
actually assumes; loading a scenario overwrites your inputs, so there is no way
to see two scenarios side by side; and nothing in the workbook records which
scenario is currently loaded, so a printed page is unidentifiable.

The alternative is a visible switch. Put every scenario in a labelled block on
`Inputs`, put a single scenario selector cell at the top, and pull the active
value with `INDEX` or `CHOOSE`:

`=INDEX(Inputs!$C$12:$E$12, $C$4)`

where `C4` holds 1, 2 or 3 and the three columns are Base, Upside and Downside.
Everything is on the page. The selector cell appears in the exhibit header, so
every printout says which scenario it is. And a Data Table can vary the
selector itself, which Scenario Manager cannot do.

### Circular references are a design decision

The classic circularity: interest expense depends on average debt, average debt
depends on the closing balance, the closing balance depends on cash flow, cash
flow depends on interest expense. Excel's iterative calculation setting will
resolve it, and it will also happily converge on a wrong answer, hide a genuine
error as a tiny residual, and make the model non-deterministic across recalcs.

Break the loop instead: calculate interest on the *opening* balance, or run one
explicit manual iteration in a visible column. Both are defensible choices you
can explain in a footnote. "I turned on iterative calculation" is not.

## Walkthrough

### Step 1 — Build the historical base by category and year

On `Calc`, create the base grid. With the Chapter 8 data model in place:

```
Revenue:   =CALCULATE(SUM(Sales[SalesAmount]))
Units:     =CALCULATE(SUM(Sales[OrderQuantity]))
Cost:      =CALCULATE(SUM(Sales[TotalProductCost]))
Margin:    =[Revenue] - [Cost]
Avg price: =DIVIDE([Revenue], [Units])
```

Without the data model, the worksheet-formula equivalents, with category and
year in the row and column headers:

`=SUMIFS(Sales[SalesAmount], Sales[Category], $A5, Sales[CalendarYear], B$4)`

`=SUMIFS(Sales[OrderQuantity], Sales[Category], $A5, Sales[CalendarYear], B$4)`

Lay this out as four categories × the complete calendar years in your release.
Add a total row and a check that the total equals
`=SUM(Sales[SalesAmount])` filtered to those years. If it does not, you have a
category mapping problem — most likely the null
`DimProduct.ProductSubcategoryKey` rows flagged in the dictionary. Decide
whether those belong in the analysis, label the decision, and move on.

### Step 2 — Decompose the revenue change

Take the last two complete calendar years. Build this schedule on `Calc`, one
row per category:

| Column | Formula |
|---|---|
| `q0`, `q1` | units, base and current year |
| `p0`, `p1` | `=IFERROR(R0/q0, 0)` and `=IFERROR(R1/q1, 0)` |
| `m0`, `m1` | `=q0/SUM($q0$ range)` and `=q1/SUM($q1$ range)` |
| Volume effect | `=($Q$1-$Q$0)*$P$0*m0` (allocating the blended volume effect by base mix) |
| Mix effect | `=$Q$1*(m1-m0)*p0` |
| Price effect | `=$Q$1*m1*(p1-p0)` |

where `$Q$0`, `$Q$1` are total units and `$P$0` is the blended base price
`=R0_total/Q0_total`.

Then the cell that matters:

`=ROUND(SUM(VolumeEffects)+SUM(MixEffects)+SUM(PriceEffects)-(R1_total-R0_total), 2)`

That must return 0. Format it so a non-zero value turns red. This single check
is what lets you put the exhibit in front of a finance director.

A typical AdventureWorks shape — these are illustrative magnitudes, not figures
to reproduce exactly — is: total revenue up, volume effect strongly positive,
mix effect substantially negative, price effect small. The story: unit growth
came from accessories and clothing, which carry a fraction of a bike's price,
so revenue grew far less than units, and because accessories carry different
margin economics the margin line moved in the opposite direction to volume.
That is the board's answer, in one exhibit, reconciled.

If you want a fourth effect, split price into list-price and discount effects
using `DiscountAmount`. The same reconciliation discipline applies: four
effects, zero residual.

### Step 3 — Lay out the driver inputs

On `Inputs`, build one block per category with three scenario columns:

```
                          Base    Upside   Downside
Customers growth %        4.0%     7.0%      1.0%
Orders per customer       1.15     1.25      1.05
Units per order           2.10     2.35      2.00
Price change %            1.5%     2.5%     -1.0%
Unit cost change %        1.0%     0.5%      3.0%
```

Scenario selector in `C4`, with Data Validation offering 1, 2, 3 and a label
beside it: `=CHOOSE($C$4,"Base","Upside","Downside")`.

Then on `Calc`, every driver is pulled through the switch:

`=INDEX(Inputs!$D$12:$F$12, 1, Inputs!$C$4)`

### Step 4 — Project revenue and margin

Three forecast years, per category:

```
Units_t   = Units_{t-1} * (1 + CustGrowth) * OrdersPerCust_factor * UnitsPerOrder_factor
Price_t   = Price_{t-1} * (1 + PriceChange)
Revenue_t = Units_t * Price_t
Cost_t    = Units_t * UnitCost_{t-1} * (1 + CostChange)
Margin_t  = Revenue_t - Cost_t
```

In cells, with drivers on the `Inputs` sheet and the prior year to the left:

`=C10*(1+$B$5)*(1+$B$6)`

`=C12*(1+$B$8)`

`=D10*D12`

Add the fixed-cost base from `FactFinance`:

`=SUMIFS(Finance[Amount], Finance[ScenarioKey], 1, Finance[AccountType], "Expense", Finance[CalendarYear], $B$3)`

Grow it at an `Inputs` inflation assumption, subtract the incremental $4.0m
investment in the years it is spent, and you have operating profit. That is the
"three-statement-lite" layer: revenue, variable cost, contribution, fixed cost,
operating profit. No balance sheet, no working capital — and say so in a note,
because a reader who expects a cash flow statement and finds a P&L will assume
the model is incomplete rather than deliberately scoped.

Compute leverage on the output sheet:

`=Contribution_Year1/OperatingProfit_Year1`

### Step 5 — Two-variable sensitivity on price and volume

The question management actually asks is "how wrong can we be?"

Put the output you care about — cumulative three-year operating profit, say — in
a single cell, `Calc!$B$40`. On a clean area of `Output`, build the table:

- Top-left corner cell: `=Calc!$B$40`
- Across the top row: price change values, −3% to +4% in 1% steps
- Down the left column: volume growth values, 0% to 8% in 1% steps

Select the whole block including the headers, then **Data → What-If Analysis →
Data Table**. Row input cell = the price-change input on `Inputs`. Column input
cell = the volume-growth input. Excel substitutes every pair and fills the grid.

Three things to know. Data Tables are array formulas (`{=TABLE(r,c)}`) and you
cannot edit a single cell. They recalculate on every workbook recalc, which on a
large model is slow enough to be worth switching to **Formulas → Calculation
Options → Automatic Except for Data Tables**. And the row/column input cells
must be on the *same sheet as the table* in older behaviour — the reliable
pattern is to keep a mirror of the two inputs next to the table and have
`Inputs` read from the mirror.

Now read the grid properly: shade the cells where the investment fails its
hurdle. The useful sentence is not "profit ranges from $X to $Y", it is "the
investment stops clearing the hurdle rate below about 2% volume growth, unless
price holds at +1.5% or better". That is a decision rule, and it is what the
client takes away.

### Step 6 — Break-even with Goal Seek, allocation with Solver

**Goal Seek** answers one-unknown questions. What volume growth makes
cumulative profit exactly equal the $4.0m spend? **Data → What-If Analysis →
Goal Seek**: Set cell = cumulative incremental profit, To value = 4000000, By
changing cell = the volume-growth input. Goal Seek overwrites that input, so
note the original value first, or run it on a copy. It also finds only one root
and needs the target cell to depend on the changing cell through formulas — not
through a Data Table or a PivotTable.

**Solver** answers constrained-allocation questions. Allocate the $4.0m across
ten territories to maximise incremental margin, where each territory has a
diminishing-returns response curve:

`=$B$20*Territory_Spend^$B$21`

with the exponent below 1 on `Inputs` — a square-root-ish response, which is
the honest default when you have no measured elasticity, and which you must
label as an assumption because it drives the entire answer.

Set objective: total incremental margin, maximise. By changing: the ten spend
cells. Constraints: sum of spend = 4000000; each spend ≥ 0; each spend ≤ a
sensible cap. Because the response curve is nonlinear, use the **GRG Nonlinear**
engine, and tick *Make Unconstrained Variables Non-Negative*. GRG finds local
optima, so run it from two or three different starting allocations and confirm
you land in the same place.

The output is only as good as the elasticity exponent. Say that on the slide.

### Step 7 — Appraise the investment

Lay the incremental cash flows in a row: the spend as negatives, the
incremental margin as positives, with an `Inputs` discount rate.

`=NPV($B$30, C45:H45) + B45`

Note the convention: `NPV` discounts the *first* argument of the range by one
period, so a Year 0 flow must sit outside the function. Getting this wrong by
one period is the single most common appraisal error.

`=IRR(B45:H45)`

For irregular dates — and real investment cases always have irregular dates —
use the date-aware pair:

`=XNPV($B$30, B45:H45, B44:H44)`

`=XIRR(B45:H45, B44:H44)`

`XNPV` discounts on actual days and does *not* have the offset quirk, which is
one more reason to prefer it.

Then the discount rate conversation, which is the part that gets skipped. The
rate is not a fact you look up; it is a claim about the risk of these cash
flows. A marketing investment inside an existing business is usually discounted
at the company's WACC. A new geography or a new channel is riskier than the
existing business and deserves a premium. Whatever you choose, put it on
`Inputs`, put the sensitivity of NPV to it in the deck, and expect the CFO to
have a different number — which is fine, because they can now change one cell
and see the answer move.

Watch IRR's failure modes: it does not exist for all cash-flow patterns, it can
have multiple roots when signs alternate, and it flatters short projects. When
IRR and NPV disagree about ranking, NPV is right.

### Step 8 — Tornado chart

Which driver matters most? Run a separate one-variable Data Table for each
driver, flexing it to a consistent low and high — plus or minus 20% of the base
value, or the Downside and Upside scenario values. Record, per driver, the low
and high outcome and the swing:

`=ABS(HighOutcome-LowOutcome)`

Sort by swing descending, then plot low and high as a stacked bar chart centred
on the base case with the series order reversed on the category axis. The result
is a funnel: the widest bar at the top is the assumption that deserves the most
scrutiny, and the narrow bars at the bottom are the ones you can stop arguing
about. On AdventureWorks, units per order and price change typically dominate;
cost inflation usually does not.

This exhibit does something politically useful as well as analytically useful:
it moves the meeting off the drivers that do not matter.

## Common pitfalls

- **Hard-coding an assumption inside a calculation.** It escapes every
  sensitivity test and every review. If a number is debatable, it lives on
  `Inputs`.
- **A residual line in the PVM schedule.** "Other: $1.2m" means your arithmetic
  is wrong. The three effects reconcile exactly; build the check cell first.
- **Price computed as a simple average of `UnitPrice`.** Average price must be
  revenue divided by units (`=Revenue/Units`), weighted by what actually sold.
  An unweighted average of list prices answers a different question.
- **Anchoring on a trailing-twelve-month base that includes the truncated final
  month.** Your forecast starts low and every subsequent year inherits the
  error.
- **Growth applied to revenue and to units at the same time.** If units grow 5%
  and price grows 2%, revenue grows 7.1%, not 5% or 7%. Build revenue as
  units × price and let it fall out.
- **Mixing calendar and fiscal years** between the sales build and the
  `FactFinance` cost base. AdventureWorks' fiscal year starts in July, so this
  silently shifts six months of cost.
- **Iterative calculation left on** to hide a circularity. Break the loop
  explicitly and document the choice.
- **Goal Seek run on the live model**, overwriting a base assumption nobody
  notices for three days.
- **Solver reporting a local optimum as the answer.** Restart from several
  points.
- **`NPV` with the Year 0 flow inside the range.** One period of over-discounting
  on the largest cash flow in the model.
- **A scenario printout with no scenario label.** Three versions of the same
  page circulate and nobody can tell them apart.
- **Reseller `AnnualSales` from `DimReseller` used as a forecast base.** It is
  survey data and it disagrees with the fact table. Facts win.

## Exercises

1. **Build the base grid.** Produce revenue, units, cost, margin and average
   price by product category for every complete calendar year in your release.
   Include a check cell proving the category totals reconcile to unfiltered
   total `SalesAmount` for those years, and state in one line how you handled
   products with a null `ProductSubcategoryKey`.

2. **Decompose and reconcile.** For the last two complete calendar years, build
   the price-volume-mix schedule by category with volume, mix and price effects
   and a reconciliation cell that must equal zero. Then write the three
   sentences you would say to the board to explain why volume rose and margin
   did not.

3. **Four-way decomposition.** Extend exercise 2 by splitting the price effect
   into a list-price effect and a discount effect using `DiscountAmount`. Keep
   the residual at zero. Which channel — internet or reseller — contributes more
   of the discount effect, and what does that imply about where pricing
   discipline is weakest?

4. **Driver model with a scenario switch.** Build the three-year projection on
   the `Inputs`/`Calc`/`Output` structure, with Base, Upside and Downside
   columns and an `INDEX`-driven switch. Every driver must come through the
   switch; no hard-coded assumption may appear on `Calc`. Report cumulative
   three-year operating profit under each scenario and the operating leverage
   in Year 1.

5. **Two-variable sensitivity and a decision rule.** Build a Data Table of
   cumulative three-year operating profit against price change (−3% to +4%) and
   volume growth (0% to 8%). Shade the region where the $4.0m investment fails
   a 12% hurdle rate. Write the decision rule as a single sentence a non-analyst
   could repeat.

6. **Tornado and appraisal.** Produce a tornado chart ranking at least five
   drivers by swing in cumulative operating profit. Then appraise the $4.0m
   investment with `XNPV` and `XIRR` on dated cash flows, test discount rates
   from 8% to 16%, and state the rate at which the investment breaks even.

7. **Constrained allocation.** Use Solver to allocate the $4.0m across sales
   territories (excluding `NA`) to maximise incremental margin under a
   diminishing-returns response curve. Run it from three different starting
   allocations. Report the allocation, whether the three runs agreed, and the
   one assumption that most affects the answer.

## Solutions

**Exercise 1.** Your reconciliation will usually fail on the first attempt, and
the cause is almost always the null `ProductSubcategoryKey` rows — manufacturing
components never sold through either channel. Because they have no sales, they
contribute nothing to revenue, so the gap is more often the territory `NA` row
or a year boundary. The correct handling is to keep all sold rows and label any
exclusion explicitly on the output sheet. An exclusion a reader cannot see is
indistinguishable from an error.

**Exercise 2.** The expected shape for AdventureWorks is a large positive
volume effect, a large negative mix effect, and a comparatively small price
effect. The reason is structural: unit growth is concentrated in accessories
and clothing, where average selling price is a small fraction of a bike's, so
each additional unit adds far less revenue than the base-year blended price
implies. The three board sentences should be something close to: units grew
strongly, but the growth was in low-price categories, so revenue grew much more
slowly than volume; percentage margin moved because the category blend changed,
not because any product's own economics deteriorated; and pricing was broadly
neutral, so this is a mix story and the lever is category strategy, not price.
The reconciliation cell is what makes those sentences defensible — without it
you are asserting a decomposition rather than demonstrating one.

**Exercise 3.** Reseller sales carry materially larger discounts, so the
discount effect concentrates there. The implication is not "stop discounting" —
wholesale discounts buy volume and shelf space. It is that reseller discounting
is where a small percentage of recovered price is worth more than anything
achievable on list price, and it is the part of the P&L with the least visible
governance.

**Exercise 4.** Cumulative profit should be meaningfully more sensitive to the
unit drivers than to cost inflation, because the contribution margin
multiplies volume changes while cost inflation hits only the variable block.
If your Upside and Downside come out nearly symmetric around Base, check
whether you have accidentally applied the same multiplier twice —
compounding `OrdersPerCustomer` and `UnitsPerOrder` growth when both describe
basket size is the usual culprit, and it double-counts.

**Exercise 5.** The honest answer is a frontier, not a point: the investment
clears the hurdle across most of the grid when price holds flat or better, and
fails in the bottom-left corner where low volume growth meets price erosion.
Expect to find that price change matters more than it feels like it should —
a one-point price movement on the whole revenue base outweighs a one-point
volume movement, because price carries no incremental cost.

**Exercise 6.** Tornado order typically puts units per order and price change
at the top. `XIRR` will exceed `IRR` on the same flows when the positive cash
flows land earlier in the year than the annual-period assumption implies, and
the dated version is the one to quote. If `XIRR` returns `#NUM!`, supply a
guess argument or check that the flows contain at least one sign change.

**Exercise 7.** With a concave response curve and no caps, Solver spreads spend
across territories rather than concentrating it, and equalises marginal return.
With the same exponent for every territory, the allocation is driven entirely
by the scale coefficient, which is to say by base-year margin — so the model is
really saying "spend more where you already make money". That may be right, but
it is a tautology dressed as an optimisation, and the honest report says so and
flags the per-territory elasticity as the thing worth measuring next.

## Key takeaways

- A model is an argument, not a number. Structure it so a reviewer can disagree
  with an assumption instead of with the whole thing.
- Three sheets: `Inputs` has no formulas, `Calc` has no hard-coded numbers,
  `Output` has no new calculations. This is the discipline everything else
  depends on.
- Build revenue from units and price, per category. A single growth rate is
  unfalsifiable and therefore useless in a meeting.
- Price-volume-mix reconciles exactly. Build the zero-residual check cell before
  you build the schedule.
- Average price is revenue divided by units, never an average of list prices.
- Use a visible `INDEX` or `CHOOSE` scenario switch, not Scenario Manager, and
  put the active scenario name in every exhibit header.
- Data Tables tell you how wrong you can be; tornado charts tell you which
  assumption to argue about. Both turn a forecast into a decision rule.
- Keep the Year 0 cash flow out of `NPV`, and prefer `XNPV`/`XIRR` on real
  dates.
- The discount rate is an assumption, not a lookup. Show the NPV's sensitivity
  to it.
- Solver's answer is only as good as the response curve you assumed. Name the
  assumption on the slide, before the client finds it.
