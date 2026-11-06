# Chapter 3 — Data Hygiene and Cleaning

## Learning objectives

By the end of this chapter you will be able to:

- Diagnose what type a column actually is, as opposed to what it looks like, and explain why the difference silently breaks aggregations.
- Repair text that carries invisible damage — leading spaces, non-printing characters, inconsistent case — using `TRIM`, `CLEAN`, `PROPER` and `SUBSTITUTE`.
- Convert text-that-should-be-numbers and integers-that-should-be-dates into real numbers and real dates, and know when the conversion is lossy.
- Split and recombine text with `LEFT`, `RIGHT`, `MID`, `TEXTSPLIT`, `TEXTBEFORE` and `TEXTAFTER`, handling the blank-middle-name case without producing double spaces.
- Distinguish blanks, zeros and nulls, and choose the right treatment for each in a consulting context where the distinction changes the answer.
- Find duplicates before deleting anything, using `COUNTIFS` and `UNIQUE`, and justify every row you remove.
- Use Data Validation to stop bad input at the source rather than cleaning it downstream forever.
- Use `IFERROR` responsibly, and recognise the far more common irresponsible use that hides a real defect.
- Derive age from a birth date and state the as-of date that makes the number meaningful.
- Build an error-audit column that flags every row failing a stated business rule, so cleaning becomes testable rather than anecdotal.

## The business question

The AdventureWorks commercial director wants to know which customer segments buy the most profitable mix of products, and she has asked for a view by home ownership, income band and age group. Three days of analysis sit behind that request. Roughly four hours of it is cleaning, and if you skip those four hours the entire analysis is wrong in ways that will not announce themselves.

Here is the specific failure that waits for you. `DimCustomer.HouseOwnerFlag` holds the values `"0"` and `"1"` — as text. Sum it and Excel returns zero, because text has no numeric value to add. A PivotTable grouping by home ownership will still work, because grouping does not require arithmetic, so you will get a table that looks entirely reasonable and reports that AdventureWorks has no homeowners when you try to count them numerically. Nobody in the room will catch it, because the number is plausible-looking and the underlying mistake is invisible.

The consulting question for this chapter is therefore not "how do I use `TRIM`". It is: **how do I establish that the data I am about to analyse can bear the weight of the conclusion I am about to build on it?** Cleaning is not clerical work that precedes the analysis. It is the first analysis, and the one on which every subsequent number depends.

The deliverable is a cleaned, analysis-ready customer table with a validation column, plus a reconciliation that accounts for every row between the raw extract and the cleaned output. If you cannot explain where each missing row went, you have not cleaned the data. You have changed it.

## Data you will use

From [the AdventureWorks data dictionary](../data/adventureworks-data-dictionary.md):

- **`DimCustomer`** — the main subject of this chapter. Specifically `CustomerKey`, `CustomerAlternateKey`, `FirstName`, `MiddleName`, `LastName`, `BirthDate`, `MaritalStatus`, `Gender`, `YearlyIncome`, `TotalChildren`, `NumberChildrenAtHome`, `EnglishEducation`, `EnglishOccupation`, `HouseOwnerFlag`, `NumberCarsOwned`, `AddressLine1`, `AddressLine2`, `Phone`, `DateFirstPurchase`, `CommuteDistance`.
- **`FactInternetSales`** — `CustomerKey`, `OrderDateKey`, `SalesAmount`, used to verify that cleaning did not orphan any customer who actually transacted.
- **`DimGeography`** — `GeographyKey`, `City`, `EnglishCountryRegionName`, used lightly to show how a cleaning rule must respect international formats.

Load `DimCustomer` as an Excel table named `Customers`. The walkthrough builds a second table, `CustomersClean`, alongside it. **Never clean in place.** The raw extract stays untouched so that any number in the final deck can be traced back to source, and so that a cleaning rule you later decide was wrong can be revised rather than archaeologically reconstructed.

Assume for the whole chapter that the raw extract has roughly 18,500 rows. Your release may differ by a few hundred; that is exactly why the walkthrough reconciles counts with formulas rather than quoting a number.

## Core concepts

### A column has a type, and the type is not what you see

Excel stores every cell as one of a small number of types: number, text, logical, error, or empty. Dates and currency are numbers wearing a format. This matters because the type determines what arithmetic is possible, and Excel will not warn you when arithmetic silently excludes most of your data.

The fastest diagnostic is to count the cells that are numeric and compare that to the count of cells that are populated:

```
=ROWS(Customers)
=COUNT(Customers[HouseOwnerFlag])
=COUNTA(Customers[HouseOwnerFlag])
```

`COUNT` counts numbers. `COUNTA` counts anything non-empty. On `HouseOwnerFlag`, `COUNTA` returns roughly 18,500 and `COUNT` returns 0. That gap is the entire diagnosis: the column is fully populated and entirely non-numeric.

Run this pair on every column you intend to do arithmetic on, before you do any. It takes two minutes and it is the single highest-return habit in this book.

A second diagnostic, useful when you suspect a mixed column — some rows numeric, some text, which is far nastier than a uniformly text column — is:

```
=SUMPRODUCT(--ISTEXT(Customers[YearlyIncome]))
```

A non-zero result on a column that should be numeric means some rows will be excluded from every `SUM`, `AVERAGE` and `SUMIFS` you write, and the result will look like a plausible number rather than an error. Mixed-type columns are the most dangerous data defect in Excel precisely because they fail partially.

### Dates that are integers, and integers that are not dates

`DateKey` columns in AdventureWorks are integers in `yyyymmdd` form: `20130715`. They are genuinely integers, not dates with an odd format. Subtracting them produces arithmetic nonsense — `20130715 - 20130630` returns 85, not 15 — because the gap crosses a month boundary that the integer encoding does not understand.

Convert deliberately:

```
=DATE(LEFT(Sales[@OrderDateKey],4), MID(Sales[@OrderDateKey],5,2), RIGHT(Sales[@OrderDateKey],2))
```

Note the implicit coercion: `LEFT` on a number returns text, and `DATE` coerces the text back to a number. It works, but it works by accident of Excel's tolerance, and it is clearer to be explicit with `VALUE` if your team reads formulas closely. The better long-term answer is to join to `DimDate` and take `FullDateAlternateKey`, which is already a real date — covered in Chapter 4. Convert in a formula only when a join is genuinely unavailable.

The reverse error also occurs. `BirthDate` and `DateFirstPurchase` arrive as real dates in most releases, but a CSV import can land them as text, particularly if the file was written in a locale whose day/month order differs from your Excel's. `=COUNT(Customers[BirthDate])` returning zero is your signal. Re-import with the column type set explicitly at import time rather than patching afterwards; a locale-mangled date column cannot always be recovered, because `03/04/2013` is irreversibly ambiguous once the source is gone.

### Text carries invisible damage

Three functions handle almost all of it.

`TRIM` removes leading and trailing spaces and collapses internal runs of spaces to one. `CLEAN` removes non-printing characters, which arrive routinely from systems that export fixed-width files or from copy-paste out of a web page. `SUBSTITUTE` removes anything else you can name, including the non-breaking space (character 160) that `TRIM` does not touch — a genuinely common defect that makes two apparently identical labels fail to match.

The standard combined repair:

```
=TRIM(CLEAN(SUBSTITUTE(Customers[@EnglishOccupation], CHAR(160), " ")))
```

Why this matters commercially: an untrimmed `"Professional "` and a clean `"Professional"` are two different values to a PivotTable, to `COUNTIFS`, and to any lookup. Your occupation breakdown will show two Professional rows, one large and one small, and the small one will look like a rounding artefact rather than a data defect. Exhibits have gone to clients with that error in them.

`PROPER` capitalises the first letter of each word. Use it with caution and never on anything you did not inspect: it will turn `AW00011000` into `Aw00011000`, breaking the business key, and it mangles names like `McDonald` into `Mcdonald` and `O'Brien` into `O'brien`. Case normalisation is for display labels, not for keys and not for names you will show to a customer.

### Splitting and recombining text

`LEFT`, `RIGHT` and `MID` take fixed positions and are right when the structure is genuinely fixed — extracting the category prefix from `ProductAlternateKey` values like `BK-R93R-62`, where the first two characters are always the product line code.

When the structure is delimited rather than fixed, modern Excel is far better:

```
=TEXTBEFORE(Products[@ProductAlternateKey], "-")
=TEXTAFTER(Products[@ProductAlternateKey], "-", -1)
=TEXTSPLIT(Customers[@AddressLine1], " ")
```

`TEXTBEFORE` and `TEXTAFTER` take an instance number, and a negative instance counts from the end — `-1` means "after the last hyphen". That single feature replaces a tangle of nested `FIND` and `LEN` calls that generations of analysts have written and mis-written.

The recombination case that trips people is the full name. `MiddleName` is blank for most AdventureWorks customers. The naive formula produces a double space:

```
=Customers[@FirstName] & " " & Customers[@MiddleName] & " " & Customers[@LastName]
```

For a customer with no middle name that yields `"Jon  Yang"`, which will not match `"Jon Yang"` in any lookup and looks sloppy in a mail merge. Two robust fixes. The first wraps the result in `TRIM`, which collapses the double space:

```
=TRIM(Customers[@FirstName] & " " & Customers[@MiddleName] & " " & Customers[@LastName])
```

The second uses `TEXTJOIN` with its ignore-empty argument set to `TRUE`, which is the clearer statement of intent:

```
=TEXTJOIN(" ", TRUE, Customers[@FirstName], Customers[@MiddleName], Customers[@LastName])
```

Prefer the second. It says what you mean — join these, skipping the empties — rather than producing a defect and then cleaning it up.

### Blanks, zeros and nulls are three different things

This is a conceptual point with direct commercial consequences, and it is where analysts most often substitute convenience for thought.

- A **zero** is a measurement. The customer owns zero cars. That is information.
- A **blank** is an absence of measurement. We do not know how many cars this customer owns. That is a different fact.
- A **null** in a relational source means the same as a blank, but it arrives through an export as either an empty cell, the literal text `"NULL"`, or a sentinel value like `-1` or `9999`. Sentinels are the dangerous case: they are numeric, so they aggregate, and an average that includes a few `9999` sentinels is badly wrong in a direction you will not notice.

`AVERAGE` ignores blanks and includes zeros, which is almost always the behaviour you want. But if a source has coded "unknown" as zero, `AVERAGE` will pull your mean toward zero and you will under-report. Check before you trust:

```
=COUNTIFS(Customers[NumberCarsOwned], 0)
=COUNTBLANK(Customers[NumberCarsOwned])
```

If the zero count is implausibly high relative to the population, suspect that zero is being used as a sentinel for unknown, and ask the source system owner. In `DimCustomer`, `AddressLine2` is genuinely blank for most rows — most addresses have no second line — and that blankness is correct and should be left alone. `TotalChildren` of zero is a real measurement. Treating both the same way is the error.

The practical rule: **decide and document the treatment of missing values per column, before analysis, and record the decision in the workbook.** "Rows with blank `BirthDate` excluded from age analysis; 74 rows, 0.4% of population" is a sentence that belongs in your appendix. It is also the sentence that protects you when someone recomputes your number and gets a different one.

### Find duplicates before you delete them

Excel's **Data → Remove Duplicates** is destructive and immediate. It deletes rows, reports a count, and gives you no record of what went. Never run it on data you have not first inspected, and never run it as the first step.

Count first:

```
=COUNTIFS(Customers[CustomerAlternateKey], Customers[@CustomerAlternateKey])
```

Added as a column, any value above 1 flags a row whose business key appears more than once. Then look at those rows. The question is not "how do I remove them" but "why do they exist", and the answer determines the treatment:

- **Genuine duplicate load** — the same row imported twice. Delete one. Safe.
- **Slowly changing dimension versions** — the same business entity at two points in time, as in `DimProduct` where one `ProductAlternateKey` has several `ProductKey` values. **Do not delete.** You need all versions; you need to pick the right one per transaction date, which is a lookup problem, not a cleaning problem.
- **Two genuinely different entities that collide on a bad key** — two customers with the same name and city. The key is inadequate; find a better one.

`UNIQUE` is the non-destructive counterpart and should be your default:

```
=ROWS(UNIQUE(Customers[CustomerAlternateKey]))
```

Compare that to `=ROWS(Customers)`. If they match, there are no duplicates on that key and you can stop worrying. If they do not, the difference is the number of excess rows, and you now have a number to investigate rather than a vague suspicion.

### Data Validation stops the problem upstream

Everything above is remedial. Data Validation is preventive, and in any workbook where a human types values — a scenario input sheet, an assumption block, a manually maintained mapping table — it is the cheapest quality control available.

Select the input range, then **Data → Data Validation**. Use a list for categorical inputs (pointing at a named range so the list is maintainable, not typed into the dialog), whole-number or decimal with bounds for numeric inputs, and a date range for dates. Write a meaningful input message and a meaningful error alert; "Invalid entry" teaches nothing, while "Discount must be between 0 and 0.40; higher discounts need commercial approval" teaches the business rule.

One limitation to know: validation applies to typed entry, not to pasted values. A user who pastes a block over a validated range will defeat it entirely. Validation reduces error rates; it does not guarantee correctness. That is why the error-audit column below exists as well.

### IFERROR, used responsibly and used badly

`IFERROR` returns an alternative when a formula errors. The responsible use is where the error is **expected, understood, and the alternative is correct**:

```
=IFERROR(XLOOKUP(Sales[@CustomerKey], Customers[CustomerKey], Customers[EnglishOccupation]), "Unknown")
```

Here you know some sales rows may reference customers absent from your filtered dimension, you have decided the right label is `"Unknown"`, and you will report how many rows got it.

The irresponsible use is the reflex wrapper:

```
=IFERROR(SUMIFS(...), 0)
```

This converts every possible failure — wrong range, misspelled column, broken reference, division by zero — into a clean, confident, wrong zero. Your total will be understated, nothing will look broken, and the defect will survive into the final deck. `IFERROR(x, 0)` on a revenue figure is one of the most reliable ways to mislead a client without intending to.

The discipline: **wrap errors only after you have reproduced them and understood their cause, and always count how many rows you suppressed.** If you do not know the count, you do not know the size of your blind spot.

```
=SUMPRODUCT(--ISNA(XLOOKUP(Sales[CustomerKey], Customers[CustomerKey], Customers[CustomerKey])))
```

That counts the unmatched rows explicitly. Report it. A footnote reading "312 of 60,398 order lines (0.5%) could not be matched to a customer record and are excluded from the segment view" is a mark of competence, not of weakness.

### Age, and the as-of date

`DimCustomer` stores `BirthDate`, not age, and that is correct design: age changes, birth date does not. Deriving age requires a decision you must state.

```
=DATEDIF(Customers[@BirthDate], $B$1, "Y")
=INT(YEARFRAC(Customers[@BirthDate], $B$1, 1))
```

Both return completed years. `DATEDIF` is an undocumented legacy function that nonetheless works reliably and reads clearly. `YEARFRAC` with basis 1 (actual/actual) is documented and composes better with other calculations. Either is fine; be consistent.

The critical detail is `$B$1`, an explicit as-of date cell, rather than `TODAY()`. If you use `TODAY()`, your age bands shift every time the workbook recalculates, and the deck you printed in March will not reproduce in June. Worse, a workbook sent to a client and opened months later shows different numbers than the ones you presented. **Pin the as-of date in a labelled input cell and reference it everywhere.**

There is a deeper point. The AdventureWorks sales window ends around early 2014. Computing customer age as of today makes every customer about a decade older than they were when they transacted, which distorts any analysis linking age to purchase behaviour. The correct as-of date for a behavioural analysis is a date inside the data window — typically the end of the analysis period, or the customer's own `DateFirstPurchase` if you want age at acquisition. Choose on analytical grounds and write the choice down.

### The error-audit column

The pattern that makes cleaning testable rather than anecdotal: one column that evaluates every stated business rule per row and reports which ones failed.

```
=LET(
  issues, TEXTJOIN("; ", TRUE,
    IF(LEN(TRIM(Customers[@LastName]))=0, "missing surname", ""),
    IF(NOT(ISNUMBER(Customers[@BirthDate])), "bad birth date", ""),
    IF(OR(Customers[@HouseOwnerFlag]<>"0", Customers[@HouseOwnerFlag]<>"1")=FALSE, "", "bad owner flag"),
    IF(Customers[@NumberChildrenAtHome] > Customers[@TotalChildren], "children at home exceeds total", "")
  ),
  IF(issues="", "OK", issues)
)
```

Three properties make this worth the effort. It is **per-row**, so you can filter to the failures and look at them. It is **cumulative**, so a row failing three rules reports all three rather than only the first. And it is **named**, so the output is a readable finding rather than `FALSE`.

The fourth rule above is a genuine cross-field consistency check: `NumberChildrenAtHome` cannot exceed `TotalChildren`. Rules of that kind catch defects that no single-column type check will ever find, and they are where domain knowledge enters the cleaning process. The type checks are mechanical; the consistency rules require you to understand the business.

## Walkthrough

### Step 1 — Load and baseline

Load `DimCustomer` into a sheet named `01_Raw` as a table named `Customers`. Do not sort it, do not delete columns, do not reformat. This sheet is the audit trail.

On a new sheet `00_Control`, build a baseline block. Label every cell.

| Cell | Label | Formula |
|---|---|---|
| `B1` | As-of date | `2014-01-31` (typed, formatted as a date) |
| `B3` | Raw row count | `=ROWS(Customers)` |
| `B4` | Distinct business keys | `=ROWS(UNIQUE(Customers[CustomerAlternateKey]))` |
| `B5` | Duplicate rows on business key | `=B3-B4` |
| `B6` | Rows with blank surname | `=COUNTIFS(Customers[LastName],"")` |
| `B7` | Rows with non-numeric birth date | `=ROWS(Customers)-COUNT(Customers[BirthDate])` |
| `B8` | Numeric cells in HouseOwnerFlag | `=COUNT(Customers[HouseOwnerFlag])` |

`B8` returning 0 against a populated column is the text-flag trap, now documented rather than discovered later. Set `B1` to a date inside the data window, for the reason given above.

### Step 2 — Type repairs

On a sheet `02_Clean`, build `CustomersClean`. Start with the keys and the repairs, one column per decision.

```
=Customers[@CustomerKey]
=TRIM(Customers[@CustomerAlternateKey])
=TEXTJOIN(" ", TRUE, TRIM(Customers[@FirstName]), TRIM(Customers[@MiddleName]), TRIM(Customers[@LastName]))
=IF(Customers[@HouseOwnerFlag]="1", TRUE, IF(Customers[@HouseOwnerFlag]="0", FALSE, NA()))
=DATEDIF(Customers[@BirthDate], Control!$B$1, "Y")
```

Note the home-ownership repair. It converts to a genuine logical, and anything that is neither `"0"` nor `"1"` becomes `#N/A` rather than quietly falling into one bucket. The error is deliberate: it will show up, and you will investigate it, which is the whole point. A `TRUE`/`FALSE` logical also sums correctly with `SUMPRODUCT(--...)` and groups correctly in a PivotTable, so one repair solves both problems.

Do not use `=VALUE(Customers[@HouseOwnerFlag])` here even though it would work. The column is semantically a flag, not a quantity, and `TRUE`/`FALSE` says so. Types should carry meaning.

### Step 3 — Normalise the phone column

`Phone` arrives in several formats — `1 (11) 500 555-0162`, `545-555-0173`, `1 (11) 500 555-0190` — because it was captured across countries and over years. For any matching or dialling use, normalise to digits only.

```
=LET(
  raw, Customers[@Phone],
  d, SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(raw)," ",""),"(",""),")",""),"-",""),".",""),
  IF(LEN(d)=0, "", d)
)
```

This is deliberately conservative: it strips formatting characters and nothing else. It does not attempt to add or infer country codes, because inferring a country code from a local number is guesswork, and guesswork embedded in a cleaning step is indistinguishable from data once two weeks have passed.

Then check what you produced, rather than assuming:

```
=COUNTIFS(CustomersClean[PhoneDigits], "<>", CustomersClean[PhoneLen], "<7")
```

Any normalised number shorter than seven digits is suspect. Flag it; do not fix it silently.

### Step 4 — The validation column

Add `ValidationFlag` to `CustomersClean` using the `LET` pattern from Core concepts, with the rules written for this table:

```
=LET(
  f, TEXTJOIN("; ", TRUE,
      IF(LEN(TRIM(Customers[@LastName]))=0,"missing surname",""),
      IF(COUNT(Customers[@BirthDate])=0,"bad birth date",""),
      IF(ISNA([@HomeOwner]),"unparseable owner flag",""),
      IF(Customers[@NumberChildrenAtHome]>Customers[@TotalChildren],"children inconsistency",""),
      IF(COUNTIFS(Customers[CustomerAlternateKey],Customers[@CustomerAlternateKey])>1,"duplicate business key","")
     ),
  IF(f="","OK",f)
)
```

### Step 5 — Reconcile

Back on `00_Control`, build the reconciliation. This is the step that distinguishes cleaning from editing.

| Cell | Label | Formula |
|---|---|---|
| `D3` | Clean row count | `=ROWS(CustomersClean)` |
| `D4` | Rows flagged OK | `=COUNTIFS(CustomersClean[ValidationFlag],"OK")` |
| `D5` | Rows flagged with issues | `=D3-D4` |
| `D6` | Issue rate | `=D5/D3` |
| `D7` | Rows dropped vs raw | `=B3-D3` |
| `D8` | Flagged rows that actually transacted | `=SUMPRODUCT((CustomersClean[ValidationFlag]<>"OK")*(COUNTIFS(Sales[CustomerKey],CustomersClean[CustomerKey])>0))` |

`D7` should be zero. The clean table has the same number of rows as the raw table, because cleaning repaired values and flagged problems — it did not delete anything. Deletion, if it happens at all, is a later and separate decision, taken with the flag column in front of you and recorded in writing.

`D8` is the commercially important number. A flagged row for a customer who never bought anything is a nuisance. A flagged row for a customer with real revenue attached is a problem, because that revenue will be misattributed or lost in your segment analysis. Sort by that and work the list in revenue order; it is the fastest route from "the data is dirty" to "here is the subset of dirt that changes the answer".

### Step 6 — Decide, document, and only then exclude

With the flags in place, write the exclusion rules explicitly in a block on `00_Control`, as sentences:

> Excluded from age-band analysis: rows with unparseable `BirthDate`. Count in `B7`, currently 0. Blank `BirthDate` would be excluded by `DATEDIF` returning an error; none present in this release.
>
> Retained with `Unknown` label: rows with unparseable `HouseOwnerFlag`. Count in `D5` detail. These rows keep their revenue in company totals but appear as a separate `Unknown` category in home-ownership splits, so the splits sum to the total.
>
> Not de-duplicated: `CustomerAlternateKey` duplicates, if any, are retained pending confirmation from the source system owner that they are not slowly changing versions.

That last sentence is the professional move. "I found duplicates and did not remove them because removal requires a decision I am not authorised to make, and here is the question I need answered" is a far stronger position than a workbook that is quietly 40 rows lighter than its source.

## Common pitfalls

**Cleaning in place.** The most common and most expensive error. Once the raw extract is overwritten you cannot answer "where did this number come from", you cannot revise a cleaning rule, and you cannot reconcile. Always keep an untouched raw sheet, even when the workbook feels small enough not to need one. It will not stay small.

**Assuming `TRIM` fixed it.** `TRIM` does not remove non-breaking spaces (character 160), which are extremely common in data pasted from web pages or Word documents. Two visually identical labels will fail to match and you will conclude the data is inconsistent when it is actually your cleaning that is incomplete. Test with `=LEN(cell)` against the length you expect, or `=CODE(RIGHT(cell,1))`, before concluding anything.

**Applying `PROPER` to keys.** It will convert `AW00011000` to `Aw00011000` and silently break every join that uses it. Case-normalise display labels; never touch identifiers.

**Treating zero and blank as interchangeable.** They encode different facts. Replacing blanks with zeros to "make the formulas work" fabricates measurements that were never taken, and will bias every average computed afterwards toward zero.

**Reflexive `IFERROR(x, 0)`.** Discussed above and worth repeating because it is so common. It converts diagnosable failures into confident wrong answers. If you must suppress an error, count the suppressions and report the count.

**Using `TODAY()` in derived ages or tenures.** Your numbers will not reproduce. A deck that cannot be reproduced cannot be defended. Pin the as-of date in a labelled cell.

**Running Remove Duplicates before investigating.** It is destructive, immediate, and gives no audit trail. Count with `COUNTIFS`, inspect the flagged rows, understand why they exist, then decide. Slowly changing dimension rows look exactly like duplicates and must not be deleted.

**Cleaning columns nobody will use.** Cleaning is expensive. `AddressLine2` is blank for most customers and irrelevant to a segment analysis; normalising it is wasted effort. Clean what the analysis touches, flag what it does not, and say which is which. Judgment about scope is part of the skill.

**Not re-running the checks after the next data refresh.** Cleaning is not a one-time act. Build the control block as formulas so that a refreshed extract re-runs every diagnostic automatically. A cleaning process that only works once is a cleaning anecdote.

## Exercises

1. **Type audit.** Build a diagnostic table with one row per column of `DimCustomer` and columns for `COUNTA`, `COUNT`, `COUNTBLANK` and the count of text cells via `SUMPRODUCT(--ISTEXT(...))`. Identify every column whose type is not what its name implies. There are at least two beyond `HouseOwnerFlag`; name them and state what arithmetic each one would break.

2. **Full name, three ways.** Build the customer full name using (a) naive concatenation, (b) `TRIM` around naive concatenation, and (c) `TEXTJOIN` with ignore-empty. Then count how many rows differ between (a) and (c) using a `SUMPRODUCT` comparison. Explain which you would ship and why, in two sentences, to a colleague who thinks all three are equivalent.

3. **Commute distance ordering.** `CommuteDistance` is text and sorts alphabetically, placing `10+ Miles` before `2-5 Miles`. Build a two-column mapping table assigning a sort index to each distinct value, then add a numeric `CommuteOrder` column to `CustomersClean` with `XLOOKUP`. Verify with `UNIQUE` that you have mapped every distinct value and that none returns an error — a mapping that silently misses a category is worse than no mapping.

4. **Age sensitivity.** Compute customer age three ways: as of `TODAY()`, as of 2014-01-31, and as of each customer's own `DateFirstPurchase`. Build a small table comparing the mean age and the share of customers under 35 under each definition. Write two sentences on which definition you would use for an analysis of age versus purchase behaviour, and why the choice is not merely cosmetic.

5. **Duplicate investigation.** Use `COUNTIFS` to flag duplicate `CustomerAlternateKey` values, and separately flag rows duplicated on the combination of `FirstName`, `LastName` and `GeographyKey`. Compare the two counts. Explain what a difference between them tells you about the adequacy of each key, and state which rows — if any — you would delete, and what you would need to confirm first.

6. **Validation that bites.** Add Data Validation to a manual mapping sheet so that the sort index in exercise 3 must be a whole number between 1 and 10. Then paste a block of invalid values over the validated range. Document what happens, and add a formula-based check that catches what validation missed.

7. **Error-audit extension.** Extend the `ValidationFlag` formula with two additional cross-field business rules of your own choosing, drawn from the data dictionary rather than invented. For each, state the rule in plain English, report the count of rows failing it, and state the share of total internet revenue attached to those rows.

## Solutions

**Exercise 1.** Beyond `HouseOwnerFlag`, the two columns most likely to surprise you are `YearlyIncome` and `NumberCarsOwned` — not in every release, but in CSV extracts where currency formatting or thousands separators survive the export, landing the column as mixed text and number. The diagnostic that reveals it is the gap between `COUNT` and `COUNTA` being non-zero but not equal to the row count: a partially text column. That partial failure is the interpretive point of the exercise. A fully text column breaks loudly, returning zero, and you catch it immediately. A column that is 94% numeric and 6% text returns a plausible `SUM` that is 6% too low, and nothing anywhere indicates a problem. `CommuteDistance` is also text, but correctly so — it is a band label, not a quantity — which is the distinction worth articulating: the question is never "is this text" but "is this the right type for what this column means".

**Exercise 2.** Expect the naive and `TEXTJOIN` versions to differ on the large majority of rows, because `MiddleName` is blank for most AdventureWorks customers and the naive formula inserts a double space on every one of them. `TRIM` and `TEXTJOIN` produce identical output here. Ship `TEXTJOIN` with ignore-empty. The reasoning to give your colleague: `TRIM` works by creating a defect and then repairing it, which means the formula does not express the business rule, while `TEXTJOIN(" ", TRUE, ...)` states the rule directly — join these parts, skip the empty ones. When the requirement later changes to include a suffix or a title, the `TEXTJOIN` version extends by adding an argument and the `TRIM` version needs re-reasoning. Formulas that say what they mean survive maintenance; formulas that happen to produce the right output do not.

**Exercise 3.** The mapping should run `0-1 Miles` → 1, `1-2 Miles` → 2, `2-5 Miles` → 3, `5-10 Miles` → 4, `10+ Miles` → 5. The verification step matters more than the mapping. Wrap the lookup in `XLOOKUP(..., ..., ..., NA())` rather than a default, run `=SUMPRODUCT(--ISNA(CustomersClean[CommuteOrder]))`, and confirm it returns zero. If you had supplied a default of 0 or 99 instead, an unmapped category would sort silently to one end of every chart and you would present a commute-distance exhibit with a category in the wrong place. The general principle: when building a mapping, make the unmapped case loud. A deliberate `#N/A` you must resolve is cheaper than a plausible number you never question.

**Exercise 4.** Age as of `TODAY()` will be roughly twelve years higher than age as of 2014-01-31, because the data window closed in early 2014. The share of customers under 35 will accordingly be far smaller under the `TODAY()` definition — potentially a few percent against something closer to a quarter. For analysing age against purchase behaviour, use age at the time of the behaviour: either the end of the analysis window, or age at `DateFirstPurchase` if you are studying acquisition specifically. The choice is not cosmetic because the conclusion inverts. Under `TODAY()` ageing, a segment that was young and high-spending when it transacted appears as a middle-aged high-spending segment, and a recommendation to target under-35s becomes a recommendation to target over-45s. Same data, opposite advice, and the only difference is an undocumented as-of date.

**Exercise 5.** In a clean AdventureWorks release, `CustomerAlternateKey` duplicates should be zero or near zero, while duplicates on name-plus-geography will be non-zero — there are genuinely different customers who share a common name within the same city. The difference tells you that `CustomerAlternateKey` is an adequate key and the name composite is not. The correct action is to delete nothing. Name collisions are not duplicates; they are the key being too weak to distinguish real entities. Had the result been reversed — business key duplicated, name composite unique — you would be looking at either a double load or slowly changing versions, and you would need the source system owner to tell you which before touching a row. Either way the deliverable from this exercise is a question for a named person, not a deletion.

**Exercise 6.** Validation blocks typed entry and reports your error message. Pasting over the range defeats it entirely and silently — the invalid values land and the validation rule remains attached to the cells, which is the worst of both worlds, because the sheet still appears validated. The formula-based catch is a check in the control block: `=SUMPRODUCT((SortIndex<1)+(SortIndex>10)+(SortIndex<>INT(SortIndex)))`, which should be zero. The lesson generalises: Data Validation is a usability feature that reduces error rates, not a control that guarantees correctness. Any rule that matters needs a formula check as well, somewhere a reviewer will look.

**Exercise 7.** Good candidate rules from the data dictionary: `DateFirstPurchase` must not precede the start of the sales window, and must not precede `BirthDate` by any amount; `YearlyIncome` must be positive; a customer with `NumberChildrenAtHome` above zero must have `TotalChildren` above zero. The second half of the exercise is the part that matters professionally. Reporting "118 rows fail validation" invites a shrug. Reporting "118 rows fail validation, carrying 0.3% of internet revenue, so the segment view is materially unaffected" closes the issue. Reporting "41 rows fail validation, carrying 11% of internet revenue, concentrated in three large accounts" escalates it, correctly, before anyone builds a recommendation on top of it. Attaching revenue to a data-quality finding is what converts it from a technical observation into a business judgment, and it is the difference between an analyst who reports problems and one who sizes them.

## Key takeaways

- Run `COUNT` against `COUNTA` on every column before you do arithmetic on it. Two minutes; catches the most expensive class of silent error.
- Mixed-type columns are more dangerous than uniformly wrong ones, because they fail partially and return plausible numbers.
- `TRIM` plus `CLEAN` plus a `SUBSTITUTE` for character 160 handles nearly all invisible text damage. `PROPER` is for display labels only, never for keys.
- Use `TEXTJOIN(" ", TRUE, ...)` to combine name parts. Write formulas that state the rule rather than formulas that create and then repair a defect.
- Blanks, zeros and sentinels encode three different facts. Decide the treatment per column, document it, and watch for zeros being used to mean unknown.
- Count duplicates with `COUNTIFS` and `UNIQUE` before deleting anything. Slowly changing dimension versions look identical to duplicates and must be kept.
- `IFERROR` is a decision, not a reflex. Reproduce the error, understand it, choose the fallback deliberately, and always report how many rows you suppressed.
- Pin the as-of date in a labelled input cell. Never derive ages or tenures from `TODAY()` in work that must reproduce.
- A per-row, cumulative, named validation column turns cleaning from an anecdote into a test that survives the next data refresh.
- Never clean in place, and always reconcile: same row count in and out, with every flagged row accounted for and sized in revenue terms.
