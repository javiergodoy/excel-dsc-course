# Chapter 11 — Regression and Driver Analysis

Up to this point you have described what happened. Revenue by category, trend by
month, variance against last year. Description is the larger part of analytical
work and it is often enough. But sooner or later a client asks a question that
description cannot answer: *why* do some customers spend more than others, and
*which* of the things we can observe about them actually matters?

That is a driver question, and regression is the standard tool for it. It is
also the point in a course where things go wrong most often, because regression
produces an impressive table of numbers whether or not the analysis means
anything. This chapter is as much about what you may claim as about which
function to type.

## Learning objectives

By the end of this chapter you will be able to:

- State the business question regression answers, and the two it does not.
- Fit a simple linear regression with `SLOPE`, `INTERCEPT` and `RSQ`, and read
  the full diagnostic output from `LINEST`.
- Run the Data Analysis ToolPak regression and interpret every line of its
  output, including the lines most people skip.
- Build a customer-level analysis table from `FactInternetSales` and
  `DimCustomer` suitable for multiple regression.
- Encode categorical attributes as dummy variables without walking into the
  dummy-variable trap.
- Translate a coefficient into a sentence in business units that a client can
  act on.
- Test the assumptions that make the output trustworthy: linearity, constant
  variance, independence and limited multicollinearity.
- Distinguish correlation, association and causation, and say precisely what a
  regression on observational data licenses you to put on a slide.
- Recognise when statistical significance is meaningless because the sample is
  large, and when Excel is the wrong tool and you should say so.

## The business question

AdventureWorks' marketing director has a budget for customer acquisition and a
list of audience attributes she can actually target: income band, life stage,
car ownership, commute distance, age. She wants to know which of those
attributes is associated with a higher-value customer, and how much of the
variation in customer value her targetable attributes explain at all.

Note the shape of that question carefully, because it is the only one the data
can support. She is *not* asking whether giving a customer a higher income
would make them spend more — that is a causal question about an intervention
nobody can perform. She is asking: among the people we could choose to target,
which ones tend to be worth more, and how confident should we be?

The honest answer has three parts: a ranked list of associations with their
magnitudes, a statement of how much variation remains unexplained, and the
caveat that these are observational associations rather than causal levers. A
regression gives you the first two. Your judgment supplies the third, and the
third is the part that protects the client from acting on nonsense.

There is also a secondary question lurking, and it is the more valuable one:
*how much of customer value can we predict at all from what we know before the
first purchase?* If the answer is "almost none", that is a finding. It tells the
marketing director to stop trying to buy the right customer and start trying to
develop the customer she gets.

## Data you will use

From `FactInternetSales`: `CustomerKey`, `SalesAmount`, `TotalProductCost`,
`SalesOrderNumber`, `OrderDateKey`.

From `DimCustomer`: `CustomerKey`, `BirthDate`, `YearlyIncome`, `TotalChildren`,
`NumberChildrenAtHome`, `NumberCarsOwned`, `EnglishEducation`,
`EnglishOccupation`, `CommuteDistance`, `MaritalStatus`, `Gender`,
`HouseOwnerFlag`, `DateFirstPurchase`.

Full column definitions and the known quirks are in
[the data dictionary](../data/adventureworks-data-dictionary.md). Two of those
quirks bite here: `HouseOwnerFlag` is stored as text, so it needs converting
before it can be a numeric predictor, and `CommuteDistance` is a banded text
field whose alphabetical order is wrong, so it needs explicit encoding.

The unit of analysis is **one row per customer**, not one row per order line.
`FactInternetSales` is at line grain, so the first task is an aggregation up to
customer level. Regressing at line grain would answer a different and largely
meaningless question — what drives the value of an individual line item — and
would weight heavy buyers many times over.

Enable the ToolPak before you start: **File → Options → Add-ins → Manage: Excel
Add-ins → Go → Analysis ToolPak**. It appears as **Data → Data Analysis**.

## Core concepts

### What a regression coefficient is

A linear regression finds the straight line (or, with several predictors, the
flat surface) that minimises the sum of squared vertical distances between the
observed points and the fitted values. The output is a set of coefficients, one
per predictor plus an intercept.

A coefficient answers exactly one question: **holding the other predictors in
the model fixed, how much does the outcome change on average per one-unit
change in this predictor?** Every word of that sentence carries weight. "On
average" — individual customers will deviate wildly. "Holding the others
fixed" — the coefficient is conditional on the model's other variables, and
changes when you add or drop one. "In the model" — a variable you did not
include is not held fixed; it is absorbed into the coefficients of whatever
correlates with it.

### R-squared, and why it is the number to lead with

R-squared is the share of variance in the outcome that the model accounts for.
It runs from 0 to 1. In cross-sectional consumer data, an R-squared of 0.10 is
ordinary and 0.30 is good; anyone showing you 0.95 on customer-level data has
almost certainly included something that is a restatement of the outcome.

Lead with R-squared rather than with coefficients, because it sets the honest
frame. "These attributes together explain roughly a tenth of the variation in
customer value" tells the marketing director the ceiling on what targeting can
achieve before she hears which attribute ranked first. Analysts who present
coefficients first and R-squared never are, deliberately or not, overstating
what they know.

Adjusted R-squared penalises the model for each additional predictor. Use it
when comparing models with different numbers of variables, because plain
R-squared can only rise when you add a variable, including a column of random
numbers.

### Significance, and its collapse at scale

A p-value is the probability of seeing an association at least this strong if
the true coefficient were zero. By convention, below 0.05 is called
significant.

AdventureWorks has roughly 18,500 internet customers. At that sample size,
standard errors are small and almost any non-zero association clears p < 0.05.
You will produce a regression where every predictor is "highly significant" and
the model explains a tenth of the variance. Both statements are true and the
combination is not a contradiction: significance says the association is
probably not exactly zero, and says nothing whatever about whether it is large
enough to matter.

So separate the two explicitly in your reporting. **Statistical significance**
is a question about sampling noise. **Practical significance** is a question
about magnitude in business units: does a coefficient of $0.004 per dollar of
income move any decision? Report the magnitude, in money, and let the client
judge. Leading with stars and p-values when the magnitudes are trivial is one
of the most common ways a competent analysis misleads.

### Correlation, association, causation

Three distinct claims, in increasing strength:

1. **Correlation.** These two variables move together. Makes no claim about
   direction or mechanism. `CORREL` gives it to you.
2. **Association, conditional on covariates.** They move together even after
   accounting for the other variables in the model. This is what a regression
   coefficient on observational data gives you. It is stronger than raw
   correlation because it rules out confounding by the variables you included —
   and only those.
3. **Causation.** Changing X changes Y. Requires either an experiment, or a
   design (instrumental variables, difference-in-differences, regression
   discontinuity) that approximates one, plus a defensible argument that nothing
   unobserved drives both.

AdventureWorks' data supports claim 2. In a deck you may write "higher-income
customers spend more, controlling for age and household composition". You may
not write "income drives spend", and you certainly may not write "raising
targeting toward high-income segments will increase revenue by X" without
flagging it as a projection resting on an assumption of unchanged behaviour.

The practical test: if your sentence implies someone pulling a lever, you are
making a causal claim and need to justify it or soften it.

## Walkthrough

### Step 1 — Build the customer-level analysis table

On a new sheet `Reg_Data`, start from the customer list and aggregate the facts
onto it. With `DimCustomer` loaded as `Customers` and `FactInternetSales` as
`Sales`:

```
A2:  =Customers[@CustomerKey]
B2:  =SUMIFS(Sales[SalesAmount], Sales[CustomerKey], [@CustomerKey])
C2:  =SUMIFS(Sales[SalesAmount], Sales[CustomerKey], [@CustomerKey])
     - SUMIFS(Sales[TotalProductCost], Sales[CustomerKey], [@CustomerKey])
D2:  =COUNTIFS(Sales[CustomerKey], [@CustomerKey])
```

Column B is lifetime revenue, C lifetime margin, D line count. For order count
you need distinct `SalesOrderNumber`, which `COUNTIFS` cannot do; use

`=SUMPRODUCT((Sales[CustomerKey]=[@CustomerKey])/COUNTIFS(Sales[SalesOrderNumber],Sales[SalesOrderNumber]))`

on a modest range, or get it from a Power Pivot `DISTINCTCOUNT` measure as in
Chapter 8. On 60,000 rows the `SUMPRODUCT` form is slow; prefer the Power Pivot
route for anything you will refresh.

Now the predictors:

```
F2:  =Customers[@YearlyIncome]
G2:  =DATEDIF(Customers[@BirthDate], DATE(2014,1,1), "y")
H2:  =Customers[@TotalChildren]
I2:  =Customers[@NumberCarsOwned]
J2:  =VALUE(Customers[@HouseOwnerFlag])
```

Two decisions are embedded in that block and both must be written down. Age is
computed against a **fixed as-of date** of 1 January 2014, near the end of the
sales window, rather than `TODAY()`. If you use `TODAY()`, every customer ages
a year each January and your model changes without the data changing — an
irreproducibility that will eventually embarrass you in front of a client.
And `HouseOwnerFlag` is wrapped in `VALUE` because it arrives as text; without
that, Excel silently treats it as zero in the regression and the coefficient
comes back as a meaningless 0 or an error.

Encode `CommuteDistance` as an ordered numeric using the band midpoint in
miles, which preserves the ordering the text destroys:

```
K2:  =SWITCH(Customers[@CommuteDistance],
        "0-1 Miles", 0.5, "1-2 Miles", 1.5, "2-5 Miles", 3.5,
        "5-10 Miles", 7.5, "10+ Miles", 12, NA())
```

Returning `NA()` on an unmatched value is deliberate: it makes a new or
misspelled band visible as an error rather than quietly recoding it to zero.

### Step 2 — One predictor first

Never run the full model first. Fit income alone, so you have a reference point
and can see what adding variables does.

```
=SLOPE(Reg_Data[Revenue], Reg_Data[YearlyIncome])
=INTERCEPT(Reg_Data[Revenue], Reg_Data[YearlyIncome])
=RSQ(Reg_Data[Revenue], Reg_Data[YearlyIncome])
=CORREL(Reg_Data[Revenue], Reg_Data[YearlyIncome])
```

Note that `RSQ` equals `CORREL` squared in the one-predictor case — a useful
sanity check on your ranges.

For the full diagnostic set from a worksheet function, use `LINEST` with both
optional flags set to `TRUE`:

`=LINEST(Reg_Data[Revenue], Reg_Data[YearlyIncome], TRUE, TRUE)`

It returns a 5-row by 2-column block: coefficient and intercept on row 1,
their standard errors on row 2, R-squared and the standard error of the estimate
on row 3, the F-statistic and residual degrees of freedom on row 4, and the
regression and residual sums of squares on row 5. In current Excel it spills
automatically; in older builds select the block and commit with
Ctrl+Shift+Enter.

Divide a coefficient by its standard error to get the t-statistic. As a rough
rule, |t| above about 2 corresponds to p below 0.05 at these sample sizes.

### Step 3 — Multiple regression through the ToolPak

`LINEST` handles multiple predictors, but the ToolPak output is easier to read
and gives you residuals for free.

**Data → Data Analysis → Regression.** Input Y Range: the revenue column with
its header. Input X Range: the contiguous block of predictor columns —
income, age, children, cars, home ownership, commute — with headers. Tick
**Labels**, tick **Residuals** and **Residual Plots**, and send the output to a
new worksheet.

The X range must be contiguous, which is why you built the predictors as
adjacent columns. If you need to drop a variable, copy the block and delete the
column rather than trying to feed the ToolPak a multi-area range.

Read the output top to bottom:

- **Multiple R / R Square / Adjusted R Square.** On this model expect an
  R-squared in the range of roughly 0.05–0.15 — illustrative, your exact figure
  depends on release and on how you handled customers with no purchases. That is
  the headline: targetable demographics explain a small minority of customer
  value.
- **Standard Error.** The typical size of a prediction miss, in dollars. Compare
  it to mean revenue. If the standard error is of the same order as the mean,
  the model cannot usefully predict any individual customer, whatever its
  p-values say. Say that out loud.
- **Observations.** Check this number. If it is far below your customer count,
  rows were dropped for blanks or text in a numeric column — go back and find
  out which, rather than accepting a silently smaller sample.
- **ANOVA table.** The F-statistic and its Significance F test whether the model
  as a whole beats an intercept-only model. At n = 18,500 it will pass. It is a
  floor, not an achievement.
- **Coefficients table.** Per predictor: the coefficient, its standard error,
  t-statistic, p-value and 95% confidence interval. The interval is the most
  honest single column; quote it rather than the point estimate when the client
  is about to size something off your number.

### Step 4 — Interpret in business units

Suppose income carries a coefficient of about 0.004 with a tight confidence
interval. The units are dollars of lifetime revenue per dollar of yearly
income. Multiply by a decision-sized change to make it legible: a customer
earning $20,000 more is associated with roughly $80 more lifetime revenue,
holding age, children, cars, home ownership and commute fixed.

Now ask the practical-significance question. Against an average customer
revenue in the high hundreds of dollars, $80 across a $20,000 income gap is
real but modest — it will not by itself justify rebuilding the targeting
strategy. Comparing standardised coefficients (each coefficient multiplied by
its predictor's standard deviation) puts predictors on a common footing and
gives you a defensible ranking of drivers.

Write the finding as one sentence with its caveat attached, and do not let the
caveat migrate to an appendix:

> Across 18,500 internet customers, the demographic attributes we can target —
> income, age, household composition, car ownership and commute — jointly
> explain roughly a tenth of the variation in lifetime customer value. Income
> is the strongest single association, at about $80 of additional lifetime
> revenue per $20,000 of yearly income, controlling for the others. These are
> observational associations, not demonstrated effects of targeting, and the
> remaining nine-tenths of variation suggests post-acquisition development
> matters more than audience selection.

That paragraph is the deliverable. The regression output is the evidence behind
it.

### Step 5 — Dummy variables and the trap

Categorical predictors need numeric encoding. `EnglishOccupation` has five
levels. Create four 0/1 columns and leave one level out as the reference:

```
=--(Customers[@EnglishOccupation]="Management")
=--(Customers[@EnglishOccupation]="Professional")
=--(Customers[@EnglishOccupation]="Skilled Manual")
=--(Customers[@EnglishOccupation]="Clerical")
```

The double negative coerces TRUE/FALSE to 1/0. "Manual" is the omitted
reference level, and each coefficient now reads as the average difference in
revenue between that occupation and Manual, holding everything else fixed.

**The dummy-variable trap:** include all five dummies alongside an intercept and
they sum to exactly 1 for every row, which is perfect collinearity with the
intercept. The design matrix is singular. Excel will either return `#NUM!` or,
worse, drop something without telling you. For k categories, use k − 1 dummies.
Always. State which level is the reference in your output, because a coefficient
is uninterpretable without it.

Do not dummy-encode an ordered band like `CommuteDistance` and a midpoint
version of the same field in the same model — you have entered the same
information twice.

### Step 6 — Test the assumptions

Four assumptions matter, and each has a cheap check.

**Linearity.** Scatter the outcome against each continuous predictor. If the
cloud bends, a straight line is the wrong functional form; try a log transform
of the outcome, `=LN([@Revenue]+1)`, which also tames the right skew typical of
spend data. Interpretation changes to approximate percentage effects, so say so.

**Constant variance (homoscedasticity).** Plot residuals against fitted values —
the ToolPak's residual plots do this. If the spread widens as fitted values
rise, a classic funnel, your standard errors are understated and your confidence
intervals too narrow. A log outcome usually fixes it. If it does not, report the
coefficients and stop quoting narrow intervals.

**Independence.** Each row must be one customer, which is why Step 1 aggregated.
If the same customer appears twice, residuals are correlated and the effective
sample is smaller than n.

**Multicollinearity.** Build a correlation matrix of the predictors with
`=CORREL()` across each pair, or use **Data Analysis → Correlation**. Pairs above
roughly 0.8 are trouble: coefficients become unstable and signs can flip when
you add a variable. In AdventureWorks, `TotalChildren` and
`NumberChildrenAtHome` are closely related and should not both be in the model.

For a VIF by hand, regress one predictor on all the others and compute
`=1/(1-RSQ_of_that_auxiliary_regression)`. Above 5 is a warning, above 10 means
act. The fix is to drop one of the pair, or combine them into one index, not to
hope the reader does not notice.

### Step 7 — Hold data out

A model fitted and evaluated on the same rows flatters itself. Assign a random
split once, then freeze it:

`=IF(RAND()<0.7, "Train", "Test")`

then immediately copy the column and paste as values so it stops recalculating.
Fit on Train, apply the coefficients to Test with `SUMPRODUCT`, and compare
R-squared out of sample. With six predictors on 18,500 rows the gap will be
small — overfitting is a modest risk here. It becomes a serious one the moment
you add dozens of dummies or go hunting for the specification with the best fit.

Which raises the real discipline: if you try twenty models and report the best
one, your p-values are fiction. Decide your specification from the business
question before you look at the output, and if you do explore, say how many
models you tried.

### Step 8 — Know when to leave Excel

The marketing director's next question will be "which customers will buy
again?" That is a binary outcome, and linear regression on a 0/1 outcome is the
wrong model: it predicts probabilities outside 0 to 1, its errors are
structurally heteroscedastic, and its coefficients do not mean what they
appear to. The right tool is logistic regression, and Excel has no logistic
regression. You can hand-roll maximum likelihood with Solver as a teaching
exercise; you should not ship it.

Say so plainly. "This question needs logistic regression, which Excel does not
support properly. I can get you a defensible answer in Python or R in a day, or
an approximate answer now by comparing repeat rates across segments." Clients
respect a clear statement of tool limits far more than they respect a forced
answer, and a wrong model delivered confidently is the one failure mode nobody
forgives. Chapter 12's segment comparison is the honest Excel-native
alternative.

## Common pitfalls

- **Regressing at line grain.** Heavy buyers get counted dozens of times and the
  question silently changes. Aggregate to customer first.
- **`TODAY()` in a derived age.** The model changes when the calendar does. Fix
  the as-of date and state it.
- **Text columns treated as numbers.** `HouseOwnerFlag` arrives as text; the
  ToolPak will either drop rows or read zeros. Watch the Observations count.
- **All k dummies plus an intercept.** Singular matrix, `#NUM!`, or a silent
  drop. Use k − 1 and name the reference level.
- **Leading with p-values at n = 18,500.** Everything is significant. Lead with
  R-squared and magnitudes in dollars.
- **Causal language.** "Drives", "impact of", "will increase". On observational
  data, write "is associated with".
- **Ignoring the residual plot.** The funnel is visible in ten seconds and
  invalidates every confidence interval you are about to quote.
- **Two versions of the same variable.** `TotalChildren` and
  `NumberChildrenAtHome` together inflate standard errors and destabilise signs.
- **Reporting the best of twenty models as if it were the only one.** The
  p-values no longer mean anything. Pre-specify, or disclose the search.
- **Extrapolating beyond the data.** The model says nothing about a customer
  earning $500,000 when your sample tops out far below that.

## Exercises

1. Build the customer-level analysis table described in Step 1: one row per
   customer from `DimCustomer`, with lifetime revenue, lifetime margin, order
   line count, and the predictors income, age as of 1 January 2014,
   `TotalChildren`, `NumberCarsOwned`, numeric `HouseOwnerFlag`, and
   `CommuteDistance` encoded as band midpoints. Report how many customers have
   zero revenue and decide, in writing, whether to include them. Justify the
   choice — it changes the question being answered.
2. Fit the simple regression of revenue on income using `SLOPE`, `INTERCEPT` and
   `RSQ`. Then reproduce all five rows of `LINEST` output and verify that the
   t-statistic you compute by hand from the coefficient and its standard error
   matches what the ToolPak reports for the same model.
3. Run the full multiple regression through the ToolPak with residual plots.
   Write a one-paragraph interpretation that leads with R-squared, ranks the
   predictors by standardised coefficient, and states each magnitude in dollars
   for a decision-sized change. No stars, no p-value tables.
4. Add `EnglishOccupation` as k − 1 dummy variables with Manual as the
   reference. Report how R-squared and adjusted R-squared each change, and
   interpret the largest occupation coefficient as a sentence. Then
   deliberately add the fifth dummy, record exactly what Excel does, and explain
   why.
5. Build the predictor correlation matrix and compute a by-hand VIF for
   `TotalChildren`. Decide whether to keep both child-related variables and
   defend the decision in two sentences.
6. Split the data 70/30, frozen as values, fit on Train and evaluate on Test.
   Report both R-squared figures and state whether overfitting is a material
   concern for this model and why.
7. Write the half-page the marketing director will read: the finding, the
   magnitudes, the confidence interval on the strongest predictor, the caveat
   about observational data, and one recommendation that follows from the low
   R-squared rather than from the coefficient ranking.

## Solutions

**Exercise 1.** Customers with no internet purchases exist in `DimCustomer`
(the dimension includes people who have not transacted in this fact table).
Including them with revenue 0 answers "what predicts value across everyone we
know of"; excluding them answers "among customers who bought, what predicts how
much". Both are legitimate and they give different coefficients. The marketing
director's question is about acquisition targeting, so the inclusive version is
the better match — but the decision must be stated in the output, because a
reader who assumes the other definition will misread every number.

**Exercise 2.** `RSQ` equals `CORREL` squared; if they disagree your ranges are
misaligned. The t-statistic is coefficient divided by standard error, and
matching it against the ToolPak confirms you have read the `LINEST` block's
layout correctly — the most common error is transposing rows 2 and 3.

**Exercise 3.** The paragraph should open with the share of variance explained,
typically around a tenth, and only then rank drivers. Income usually leads on
standardised coefficient, with age and household composition behind it.
Expressing each as dollars per decision-sized change — per $20,000 of income,
per additional child — is what makes it usable. A reader who gets p-values
first and R-squared last will overestimate what the model knows, which is the
failure this exercise exists to prevent.

**Exercise 4.** Adjusted R-squared is the honest comparison: plain R-squared
cannot fall when you add variables, so only the adjusted figure tells you
whether occupation earned its place. Adding the fifth dummy creates perfect
collinearity with the intercept; Excel typically returns `#NUM!` from `LINEST`,
and the ToolPak may produce output with a dropped or wildly inflated
coefficient. The lesson is that the failure is sometimes silent, so the k − 1
rule must be habit rather than something you check for afterwards.

**Exercise 5.** `TotalChildren` and `NumberChildrenAtHome` correlate strongly;
VIF on either typically lands well above the warning threshold. Keep one. If
life stage matters to the client, `NumberChildrenAtHome` is the more
behaviourally relevant of the two, since it describes the household now rather
than ever.

**Exercise 6.** Expect Train and Test R-squared within a point or two of each
other. With six predictors and 18,500 rows there is little to overfit; the
exercise matters because the habit, not this result, is what protects you when
the model grows.

**Exercise 7.** The strongest answer does not recommend tilting the targeting
mix. It observes that roughly nine-tenths of customer value is unexplained by
targetable demographics, and recommends shifting budget from audience selection
toward post-first-purchase development — which Chapter 12's cohort and segment
work can size. The regression's most valuable output here is its low R-squared.

## Key takeaways

- Regression answers "which observable characteristics are associated with
  higher value, and how much do they explain". It does not answer "what will
  happen if we change something".
- Aggregate to the unit of analysis before you fit. Customer questions need
  customer rows.
- Lead with R-squared. It frames how much the model knows before anyone hears
  which variable ranked first.
- At 18,500 observations everything is statistically significant. Report
  magnitudes in business units and let the client judge practical significance.
- Categorical variables need k − 1 dummies and a named reference level.
- Check the residual plot and the predictor correlation matrix every time. Both
  take a minute and both can invalidate the output you were about to present.
- Write "is associated with", not "drives", unless you have a design that earns
  the stronger verb.
- A low R-squared is a finding, not a failure. "We cannot predict this from what
  we know" redirects a client's budget more usefully than a weak coefficient
  ranking.
- Binary outcomes need logistic regression. Excel does not do it properly.
  Saying so is better analysis than forcing it.
