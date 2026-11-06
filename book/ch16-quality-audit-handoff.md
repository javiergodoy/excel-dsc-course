# Chapter 16 — Workbook Quality, Audit and Handoff

A workbook nobody can audit is worthless, and this is true whether or not it is
correct. That sounds like an overstatement until you have watched it happen: a
model that produces the right answer, challenged in a meeting, cannot be
defended in the room, and the recommendation dies. The analyst was right. It
made no difference. Correctness that cannot be demonstrated is indistinguishable
from luck, and experienced clients treat it that way.

This chapter is about the habits that make a workbook defensible: tie-outs that
fail loudly, a checks sheet that goes red before the client does, documentation
the next person can actually use, and a review process a colleague can execute
in twenty minutes. It closes the book because it is the part that transfers
furthest beyond Excel.

## Learning objectives

By the end of this chapter you will be able to:

- Classify errors as mechanical, logical or interpretive, and recognise why the
  dangerous ones produce no error message at all.
- Build a dedicated checks sheet of pass/fail reconciliations that turn red
  automatically when an assumption breaks.
- Reconcile every derived total back to a source total, and cross-foot a table
  in both directions.
- Carry control totals through each transformation step so you know exactly
  where rows went missing.
- Use Trace Precedents and Dependents, Evaluate Formula, the Watch Window, Show
  Formulas and Error Checking as investigative tools rather than curiosities.
- Find hard-coded constants buried inside formulas, and external links you did
  not know you had.
- Protect inputs without encrypting logic, and version a workbook so the
  history is recoverable.
- Write a README sheet that tells the next person what the workbook is, what it
  assumes, where the data came from and what needs refreshing.
- Manage calculation mode and performance on a large model without creating a
  silent-staleness risk.
- Run a twenty-minute peer review against a checklist, and decide whether to
  hand the workbook to the client at all.

## The business question

The AdventureWorks growth recommendation from Chapter 15 was accepted. Two
things now happen, and both are tests of workbook quality rather than analysis.

First, the CFO's team wants to re-run your sizing with their own attach-rate
assumptions before the budget is committed. They will open your workbook
without you in the room.

Second, the engagement ends and a different analyst picks up the pilot
measurement in three months. They will inherit your files with no handover
conversation, because the person who would have given it has rolled off.

So the question is not "is my analysis right". It is:

> *Can someone who has never seen this workbook change one assumption, get a
> defensible answer, and know immediately if they have broken something?*

If the answer is no, the analysis has a shelf life of one meeting.

## Data you will use

The Part III analysis workbook you have been building: the unioned sales table
from Chapter 7, the derived margin columns, the category and territory
dimensions, the segmentation output from Chapter 12, and the sizing model from
Chapter 15.

The source totals that every check reconciles against are the raw loaded
tables, before any transformation:

- Row count and `SUM(SalesAmount)` of `FactInternetSales` as loaded.
- Row count and `SUM(SalesAmount)` of `FactResellerSales` as loaded.
- Row count of `DimProduct`, `DimCustomer`, `DimReseller`.
- `SUM(TotalProductCost)` for both fact tables.

Record these once, on the checks sheet, as the immutable reference. Everything
downstream ties back to them. Margin remains `SalesAmount − TotalProductCost`
per [the data dictionary](../data/adventureworks-data-dictionary.md), and
Chapter 15's conclusions rest on
[the storyboard numbers](ch15-analysis-to-recommendation.md), which is exactly
why they need tying out.

## Core concepts

### The error taxonomy, and why silence is the problem

Three kinds of error, in increasing order of danger.

**Mechanical.** A range that should be `Sales[SalesAmount]` and is
`Sales[ExtendedAmount]`. A `SUMIFS` whose criteria range is off by a row. A
division by a cell that is empty. These announce themselves — wrong magnitude,
`#DIV/0!`, `#N/A` — and are usually caught within minutes.

**Logical.** Double-counting because you joined at the wrong grain. Summing
`HouseOwnerFlag` and getting zero because it is text. Averaging an average.
Filtering out the territory `NA` row in one exhibit and not another, so two
slides disagree by 0.3% and someone notices. These produce plausible numbers.
Nothing goes red. You find them by reconciliation or not at all.

**Interpretive.** The arithmetic is right and the meaning is wrong. Presenting
the truncated final month as a demand collapse. Using `DimReseller.AnnualSales`
— survey data — as if it were transactional. Describing calendar-year growth to
a controller who thinks in fiscal years. No formula can catch these. Only
someone who understands the business can, which is the argument for the README
sheet and for peer review by a human.

The distribution matters: mechanical errors are common and cheap, logical
errors are less common and expensive, interpretive errors are rare and
occasionally career-defining. Spending your review time proportional to
frequency is backwards. Spend it proportional to cost.

### Reconciliation and tie-out as a habit

A tie-out is a formula that compares a derived number to an independent source
number and reports pass or fail. Not a number you eyeball. A verdict.

The rule: **every derived total ties back to a source total.** If you unioned
two fact tables, the union's row count equals the sum of the two source counts,
and its revenue equals the sum of the two source revenues:

```excel
=IF(AND(ROWS(Sales)=Internet_Rows+Reseller_Rows,
        ROUND(SUM(Sales[SalesAmount]),2)
         =ROUND(Internet_Revenue+Reseller_Revenue,2)),
    "OK","FAIL: union does not tie to sources")
```

Two properties make this useful rather than decorative. It is a single cell
with a text verdict, so conditional formatting can turn it red. And it compares
against numbers captured *before* the transformation, so it detects the failure
mode that matters — a transformation that quietly drops rows.

Always `ROUND` before comparing currency. Floating-point arithmetic will
produce differences of 0.0000001 that are not errors, and a check that cries
wolf gets ignored, which is worse than having no check.

### Cross-footing

Cross-footing is reconciling a table in both directions: the sum of the row
totals must equal the sum of the column totals must equal the grand total.

For margin by category and territory group:

```excel
=IF(ROUND(SUM(RowTotals),2)=ROUND(SUM(ColTotals),2),
    IF(ROUND(SUM(RowTotals),2)=ROUND(GrandTotal,2),"OK",
       "FAIL: margins do not tie to grand total"),
    "FAIL: rows and columns disagree")
```

This catches the single most common PivotTable deception: a table that looks
complete but whose grand total exceeds the visible parts because something is
filtered out of the display and not out of the total — or, worse, the reverse.
In AdventureWorks the usual culprit is `DimProduct.ProductSubcategoryKey` being
null for non-sellable components, so category-grouped exhibits silently exclude
rows that the grand total includes.

### Control totals through transformations

Power Query makes it easy to build a chain of a dozen steps and impossible to
see what each one removed. So measure.

Keep a small table with one row per transformation step, columns for row count
in, row count out, and expected delta, with a verdict:

| Step | Rows in | Rows out | Expected | Verdict |
|---|---|---|---|---|
| Load FactInternetSales | — | 60,398 | as source | OK |
| Remove null ProductKey | 60,398 | 60,398 | 0 removed | OK |
| Merge DimProduct | 60,398 | 60,398 | no fan-out | OK |
| Filter out territory NA | 60,398 | 60,214 | ~0.3% removed | OK |

```excel
=IF(Rows_Out=Rows_In-Expected_Removed,"OK",
    "FAIL: step removed "&Rows_In-Rows_Out&" rows, expected "&Expected_Removed)
```

The "no fan-out" row is the important one. A merge that should be
many-to-one and is accidentally many-to-many duplicates rows and inflates
revenue, and nothing about the result looks wrong. The row count is the only
signal, and it is a decisive one.

### Investigating a formula you did not write

Five tools, each for a different question.

**Trace Precedents / Dependents** (`Formulas → Trace Precedents`, or Ctrl+[ to
jump to a precedent). Precedents answer "what feeds this cell". Dependents
answer the far more important question before you change anything: "what breaks
if I touch this". Always trace dependents before editing someone else's model.

**Evaluate Formula.** Steps through a nested formula one calculation at a time.
This is how you debug a `LET` with four variables, or find which branch of a
nested `IF` is firing. Faster than decomposing the formula into helper cells,
and it does not alter the workbook.

**Watch Window.** Pins chosen cells and shows their values live while you work
elsewhere. Put your checks-sheet verdicts and your three headline numbers in
it. Now when you change an assumption on another sheet, you see the checks go
red immediately rather than discovering it later.

**Show Formulas** (Ctrl+`). Displays formulas instead of results across the
sheet. This is the fastest way to spot a column where one cell has been
overtyped with a constant — the pattern breaks visibly.

**Error Checking** and, where your installation has it, **Inquire**. Error
Checking finds inconsistent formulas, formulas omitting adjacent cells, and
text-formatted numbers. Inquire — available in some Professional Plus
installations — produces a workbook relationship diagram and a formula-by-sheet
inventory, and compares two versions of a file cell by cell. If you have it,
the version comparison is worth the whole feature.

### Hard-coded numbers inside formulas

The error that survives every review done by reading results rather than
formulas:

```excel
=SUM(Sales[SalesAmount])*0.82
```

Why 0.82? Nobody knows. It was right in March. It is wrong now and will stay
wrong forever, because the number is invisible from the output.

Find them. Select the sheet, then `Home → Find & Select → Formulas` to select
every formula cell, then scan with Show Formulas on. For a systematic sweep,
a helper column flags suspects:

```excel
=IF(AND(ISFORMULA(C5),OR(ISNUMBER(SEARCH("*0.",FORMULATEXT(C5))),
   ISNUMBER(SEARCH("/1000",FORMULATEXT(C5))))),"REVIEW","")
```

The rule that prevents them: **every constant that is not 0, 1, 12 or 100 lives
in a named input cell on an inputs sheet.** Then `=SUM(Sales[SalesAmount])*
Attach_Conversion_Rate` is self-documenting, and the CFO's team can change it
without editing a formula — which is exactly what they came to do.

### External links

A workbook that references another file is a workbook that can break when
someone else moves a folder, and that will prompt the client with a security
warning you are not there to explain.

`Data → Queries & Connections → Edit Links` lists them. `Find` for `[` or `.xls`
in formulas catches ones that hide inside nested functions. Decide
deliberately: either break the link and paste values with a note recording
where they came from and as of when, or document the dependency prominently on
the README sheet. Both are defensible. An undocumented link is not.

### Protecting inputs without encrypting logic

The goal is to stop accidental overwrites, not to hide your work. Hiding your
work in a consulting deliverable is a reputational problem, and a client who
discovers a password-protected calculation sheet reasonably concludes you did
not want it examined.

The pattern: unlock the input cells, lock everything else, protect the sheet
with no password or a documented one, and leave structure visible.

1. Select all, `Format Cells → Protection → Locked` on.
2. Select the input cells, turn `Locked` off.
3. `Review → Protect Sheet`, allowing select-locked-cells so people can still
   read and trace formulas.

Never protect with an undocumented password on a deliverable. You will leave
the project, and the workbook becomes unmaintainable at exactly the moment
nobody can ask you.

### Naming, versioning, change log

File names sort chronologically and carry meaning:
`AW_Growth_Sizing_v04_2026-03-18.xlsx`. Not `final`. Not `final_v2`. Not
`final_FINAL_use_this`. Every analyst has met that folder and nobody can
reconstruct which file produced the number in the deck.

A change log sheet, four columns: date, version, what changed, who. Three lines
per version is enough. Its value appears when a number moves between two
meetings and someone asks why — and that question always comes, usually from
the person least inclined to accept "I must have refined something".

### The README sheet

The first sheet in the workbook, before any data. It answers what the next
person will ask, in this order:

- **Purpose.** One sentence: what question this workbook answers.
- **Owner and date.** Who built it, when, who to ask.
- **Data sources.** Which AdventureWorks release, which tables, extracted when,
  from where. "AdventureWorksDW2019, FactInternetSales and FactResellerSales,
  extracted 2026-03-02 via Power Query from the CSV release."
- **Definitions.** Every derived metric, stated once. "Gross margin =
  SalesAmount − TotalProductCost. Not a source column."
- **Assumptions.** Every judgment call with its basis, and who owns it.
- **Exclusions.** Territory `NA` excluded; partial final month excluded;
  products with null subcategory excluded from category views. Each with a
  reason.
- **How to refresh.** The exact sequence, in order, including what to check
  afterwards: "Data → Refresh All, then confirm checks sheet is all OK."
- **Known limitations.** Promotion analysis not possible — nearly all rows are
  `No Discount`. Reseller dimension `AnnualSales` is survey data and disagrees
  with the fact table; the facts are used.
- **Sheet map.** One line per sheet.

Twenty minutes to write. It is the difference between a workbook that outlives
the engagement and one that is rebuilt from scratch by the next analyst, who
will quietly assume yours was wrong.

### Refresh and recalculation discipline

On a large model, automatic calculation becomes unworkable and you switch to
`Formulas → Calculation Options → Manual`. This is a legitimate performance
decision and it introduces a genuine hazard: the workbook can be saved and sent
while displaying stale results. The displayed numbers are from the previous
assumptions and nothing indicates it.

Three defences. Press F9 before every save, until it is reflex. Put a
calculation-mode indicator on the checks sheet:

```excel
=IF(INFO("recalc")="Automatic","Auto","MANUAL - press F9 before trusting")
```

And set calculation back to automatic before the final save of a deliverable,
so the recipient cannot inherit the hazard.

Refresh order matters too. Power Query loads, then Power Pivot measures
recalculate, then PivotTables need refreshing, then formulas referring to those
PivotTables recalculate. `Refresh All` usually handles the chain, but a
GETPIVOTDATA formula pointing at a Pivot that has not refreshed returns the old
number without complaint. Check the checks sheet after every refresh; that is
what it is for.

### Performance and file-size hygiene

Large workbooks are not only slow, they are fragile, and the causes are
consistent:

- **Whole-column references.** `=SUMIFS(A:A,B:B,...)` over a million rows,
  recalculated constantly. Structured table references size themselves.
- **Volatile functions.** `OFFSET`, `INDIRECT`, `TODAY`, `NOW`, `RAND`
  recalculate on every change and force their dependents to as well.
- **Formatting applied to the used range by accident.** Ctrl+End jumping to
  `XFD1048576` is the symptom. Delete the unused rows and columns, save, reopen.
- **Duplicated raw data.** The same fact table pasted into three sheets. Load
  once into the data model and reference it.
- **Images and stale shapes** from pasted charts nobody deleted.

Prefer Power Pivot measures over hundreds of thousands of calculated cells on
large models. A DAX measure computes on demand over compressed columns; a
helper column materialises a value for every row whether anyone looks at it or
not.

### The twenty-minute peer review

A review that takes a day does not happen. Make it executable.

1. **Open the README.** Can you state what the workbook does, from the sheet
   alone? (2 min)
2. **Checks sheet.** All green? Any check that is permanently green regardless
   of input is not a check — break something deliberately and confirm it goes
   red. (3 min)
3. **Tie the headline number** to one source total independently, with your own
   formula in a scratch cell. (4 min)
4. **Show Formulas on the two main calculation sheets.** Look for broken
   patterns and hard-coded constants. (3 min)
5. **Trace precedents on the three biggest numbers.** Do they come from where
   the README says? (3 min)
6. **Edit Links** and `INFO("recalc")`. External dependencies and calculation
   mode. (2 min)
7. **Change one input** and confirm the output moves in the expected direction
   and magnitude. (3 min)

Step 2's sub-clause is the one people skip and the one that catches the most.
A checks sheet full of green cells that would stay green if you deleted half
the data is theatre. Break something on purpose, every time.

### Should you hand the workbook to the client?

Not automatically. Consider three things.

**Can it be maintained without you?** If it needs a verbal explanation to run,
handing it over creates a liability rather than an asset. Either document it to
the point where it does not, or hand over results and keep the model.

**What is in it that should not leave?** Benchmark data under licence, other
clients' figures in a comparison tab, intermediate scratch work with
speculative labels. Audit every sheet, including hidden ones — `Right-click →
Unhide` and check — before anything leaves.

**Will it be used as a decision tool or read as a report?** If the CFO's team
will change assumptions and act on the output, the inputs must be protected,
labelled and sensitivity-tested, and the README must say which assumptions are
theirs to change and which are structural.

When you do hand it over, hand over a version with calculation set to
automatic, all checks green, the change log current, and a short covering note
naming the one number that matters and the one assumption it is most sensitive
to.

## Walkthrough

### Step 1 — Capture immutable source totals

New sheet, first position after the README, named `Checks`. Top block, labelled
"Source totals captured at load — do not edit":

```excel
=ROWS(FactInternetSales)
=ROUND(SUM(FactInternetSales[SalesAmount]),2)
=ROUND(SUM(FactInternetSales[TotalProductCost]),2)
=ROWS(FactResellerSales)
=ROUND(SUM(FactResellerSales[SalesAmount]),2)
```

Name them `Internet_Rows`, `Internet_Revenue`, `Internet_Cost`,
`Reseller_Rows`, `Reseller_Revenue`. Named ranges, not cell addresses, because
the checks below should read as sentences.

### Step 2 — Write checks that fail loudly

Second block, one row per check, with a `Verdict` column. Six checks earn their
place in this workbook.

**Union ties to sources:**

```excel
=IF(AND(ROWS(Sales)=Internet_Rows+Reseller_Rows,
        ROUND(SUM(Sales[SalesAmount]),2)=ROUND(Internet_Revenue+Reseller_Revenue,2)),
    "OK","FAIL: union lost rows or revenue")
```

**Margin is internally consistent:**

```excel
=IF(ROUND(SUM(Sales[Margin]),2)
     =ROUND(SUM(Sales[SalesAmount])-SUM(Sales[TotalProductCost]),2),
    "OK","FAIL: margin column does not equal revenue less cost")
```

**No orphaned product keys** — the join that silently drops rows:

```excel
=IF(SUMPRODUCT(--ISNA(MATCH(Sales[ProductKey],Products[ProductKey],0)))=0,
    "OK","FAIL: sales rows with no matching product")
```

**Category exhibit ties to the grand total** — catches the null-subcategory
exclusion:

```excel
=IF(ROUND(SUM(CategoryMargin[Margin]),2)=ROUND(SUM(Sales[Margin]),2),
    "OK","FAIL: category view excludes "
     &TEXT(SUM(Sales[Margin])-SUM(CategoryMargin[Margin]),"#,##0")&" of margin")
```

**Segmentation covers every customer exactly once:**

```excel
=IF(AND(SUM(Segments[CustomerCount])=ROWS(Customers),
        COUNTA(UNIQUE(SegmentAssignment[CustomerKey]))=ROWS(Customers)),
    "OK","FAIL: segments do not partition the customer base")
```

**Calculation mode:**

```excel
=IF(INFO("recalc")="Automatic","Auto","MANUAL - press F9 before trusting")
```

Then a master verdict at the very top, which is the only cell anyone needs to
look at:

```excel
=IF(COUNTIF(Verdicts,"OK")=ROWS(Verdicts),"ALL CHECKS PASS",
    COUNTIF(Verdicts,"FAIL*")&" CHECK(S) FAILING")
```

Conditional format it: green on "ALL CHECKS PASS", red otherwise, bold, large.
Add the master cell to the Watch Window.

### Step 3 — Break things on purpose

A check you have never seen fail is a check you do not know works. Test each
one:

- Delete a row from `Sales` → the union check must fail.
- Overtype one `Margin` cell with a constant → the margin check must fail.
- Change the category view's filter to exclude Accessories → the cross-foot
  check must fail and report the missing amount.
- Switch calculation to manual → the mode indicator must change.

Undo each. If any check stayed green, it is wrong — most often because it
compares a derived number to another derived number that moves with it, which
is self-fulfilling and useless.

### Step 4 — Run the peer-review checklist and fix what it finds

Executed against a real Part III workbook, the review typically surfaces four
things. All four are representative rather than hypothetical, and all four are
fixes you should expect to make.

*No README sheet.* Write it, using the structure above. Fifteen minutes.

*A hard-coded constant.* `=SUM(Sales[SalesAmount])/1000000` labelled "$m" in
one place and a different scaling elsewhere, so two exhibits disagree by a
factor of a thousand. Move the scalar to a named input `Scale_Millions` and
reference it.

*A category exhibit that does not cross-foot*, short by the margin of products
with null `ProductSubcategoryKey`. Decide explicitly: either add an
"Unclassified" category so the total ties, or keep the exclusion and state it
on the exhibit. Add the check either way.

*Manual calculation left on.* Set it back to automatic, add the indicator
check, press F9, save.

Re-run the checklist. The point of a checklist is that the second pass is
cheap, and a fix that breaks something else shows up immediately.

## Common pitfalls

**Checks that cannot fail.** Comparing a derived figure to another derived
figure from the same chain. It will always agree, and it will always agree
while being wrong.

**Comparing unrounded currency.** Floating-point noise produces failures of
0.0000001. The check gets ignored within a week, and so do the real failures
next to it.

**Eyeballing instead of asserting.** A sheet of numbers you "check" by looking
at them is not reconciled. If there is no formula producing a verdict, there is
no check.

**Reconciling only at the end.** By then you cannot tell which of twelve
transformation steps lost the rows. Control totals at each step cost minutes
and save hours.

**Password-protecting logic.** It reads as concealment, and it makes the
workbook dead on the day you leave.

**The `final_final` folder.** If you cannot tell which file produced the number
in the deck, the number cannot be defended.

**Documenting data sources in your head.** Three months later you will not
remember which AdventureWorks release you used, and the row counts differ
between releases, so the next person's numbers will not match yours and neither
of you will know why.

**Treating performance work as cosmetic.** A model that takes ninety seconds to
recalculate does not get tested properly, because nobody tests what is painful.
Slowness causes errors.

**Assuming refresh is atomic.** Query loads, measure recalculation, Pivot
refresh and formula recalculation are separate. GETPIVOTDATA will hand you a
stale number with total confidence.

## Exercises

1. **Build the checks sheet.** Implement all six checks from the walkthrough on
   your own Part III workbook, with named source totals and the master verdict
   cell. Add the master cell to the Watch Window.

2. **Break every check.** For each of the six, cause a failure deliberately,
   record the exact failure message produced, and undo it. Report any check
   that stayed green and explain why it was self-fulfilling.

3. **Control totals through Power Query.** Document your query chain as a
   step-in/step-out table with expected deltas and verdicts. Identify which
   step has a fan-out risk and write the check that would detect it.

4. **Hunt hard-coded constants.** Sweep your workbook with Show Formulas and
   the `FORMULATEXT` flag helper. List every constant that is not 0, 1, 12 or
   100, state what it was meant to represent, and move it to a named input.

5. **Write the README sheet.** All nine blocks, for your own workbook. Then
   give the workbook and nothing else to a colleague and ask them to refresh it
   and reproduce your headline number. Record every question they had to ask
   you, and add each answer to the README.

6. **Run the twenty-minute review on someone else's workbook.** Time yourself.
   Produce a findings list classified as mechanical, logical or interpretive.
   Note which step of the checklist found each one.

7. **Make the handoff decision.** For your Part III workbook, write a short
   memo: would you hand it to the AdventureWorks CFO's team, what would you
   remove first, which assumptions are theirs to change, and what covering note
   would accompany it.

## Solutions

**Exercise 2.** The check that most often survives being broken is the margin
consistency check, when `Sales[Margin]` and the comparison are both computed
from the same column at the same moment — the two sides move together, so the
test cannot discriminate. The fix is to compare against the *captured source
totals* rather than the live columns: `Internet_Revenue - Internet_Cost` plus
the reseller equivalent, recorded at load and never recomputed. The general
principle is that a check needs an independent reference point, and anything
inside the same calculation chain is not independent. This is also why the
source-totals block carries the instruction "do not edit"; the moment someone
helpfully "fixes" a source total to make a check pass, the workbook has lost its
only external anchor and every verdict below becomes decorative.

**Exercise 3.** The fan-out risk is the merge against `DimProduct`. Because
AdventureWorks versions products — the same `ProductAlternateKey` appearing with
several `ProductKey` values and different validity windows — a merge joined on
the alternate key rather than the surrogate key produces multiple matches per
sales row, duplicating revenue invisibly. The row count is the only symptom and
it is decisive: if rows out exceeds rows in on a merge that should be
many-to-one, you have fanned out. The check is simply
`=IF(Rows_Out=Rows_In,"OK","FAIL: merge fanned out")` on that step, and it
should be written before the merge is trusted rather than after a number looks
surprising.

**Exercise 4.** The commonest finds are a units scalar applied inconsistently,
a tax or uplift percentage from a conversation nobody recorded, and a row-limit
constant left over from testing on a sample — the last being the dangerous one,
because it produces a plausible answer from part of the data. For each, the
remediation is the same: named cell, inputs sheet, basis documented. The test of
whether you have finished is that the CFO's team can change every assumption
the recommendation depends on without opening a single formula.

**Exercise 5.** Expect three to five questions, and the pattern is consistent:
which release of the data, what exactly was excluded, and what to press to
refresh. Every question is a gap in the README, and the exercise is really a
measurement of that document rather than of your colleague. A README that
survives this test unamended is rare and is the actual deliverable of this
chapter.

**Exercise 7.** The defensible answer is usually a qualified yes: hand over a
protected-input version with the scratch sheets removed, all checks green,
calculation set to automatic, and a covering note naming the attach-rate
assumption as the single sensitivity. The things to remove first are hidden
sheets and any comparison tab containing figures that are not AdventureWorks'
own. If the honest answer is that the workbook cannot be refreshed without you
on the phone, the correct decision is to hand over the results and keep the
model until it can be.

## Key takeaways

- A workbook that cannot be audited is worthless whether or not it is correct.
  Correctness you cannot demonstrate is indistinguishable from luck.
- Mechanical errors shout, logical errors stay silent, interpretive errors
  require a human who understands the business. Allocate review effort by cost,
  not by frequency.
- Every derived total ties back to a source total captured before
  transformation, through a formula that returns a verdict, not a number you
  look at.
- A check you have never seen fail is not a check. Break each one on purpose.
- Round currency before comparing. A check that cries wolf gets ignored, which
  is worse than no check at all.
- Carry control totals through every transformation step. Fan-out on a merge is
  invisible in the values and obvious in the row count.
- Every constant that is not 0, 1, 12 or 100 belongs in a named input cell.
- Protect inputs, never conceal logic, and never use an undocumented password
  on something you will hand over.
- The README sheet — purpose, sources, definitions, assumptions, exclusions,
  refresh steps, limitations — is what makes the work outlive the engagement.
- Manual calculation mode is a legitimate performance choice and a silent
  staleness risk. Indicate it, press F9, and restore automatic before the final
  save.

### What transfers

The functions in this book will date. Dynamic arrays replaced a generation of
array formulas, Power Query replaced a generation of manual cleaning, and
something will replace both. The habits will not date, because they are not
about Excel.

Separating inputs from logic is the same discipline as parameterising a SQL
query instead of editing literals into the `WHERE` clause. Defining a metric
once and referencing it everywhere is a DAX measure, a dbt model, a semantic
layer. Tie-outs are unit tests and data-quality assertions. The README sheet is
a model's documentation block. Control totals through transformations are
row-count assertions in a pipeline. Versioning and a change log are, more or
less, source control.

Which is the real argument of the course. The analyst who has internalised these
habits moves to SQL, Python or a BI tool and carries the judgment across, needing
only new syntax. The one who memorised two hundred functions starts again from
nothing, and — more consequentially — never learned the part that was actually
hard: making a number you can stand behind when someone who outranks you asks
where it came from.
