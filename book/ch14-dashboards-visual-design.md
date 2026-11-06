# Chapter 14 — Dashboards and Visual Design

Everything in the previous thirteen chapters produced numbers. This chapter is
about the last few inches between a correct number and a decision, and those
inches are where most analytical work dies. A client does not act on your
analysis. They act on what they understood of your analysis, in the ninety
seconds they spent looking at your chart, in a meeting where two other things
were also on the agenda.

That is not a complaint about clients. It is the design constraint. Treat it
as one and your work gets used.

## Learning objectives

By the end of this chapter you will be able to:

- Distinguish a dashboard, an exhibit and a report, and build the right one for
  the request you were actually given.
- Choose a chart form from the question being asked, using a short list of forms
  that survive client scrutiny.
- Explain precisely what goes wrong with pie charts, dual axes, 3-D effects and
  truncated axes — and name the rare case where a broken axis is honest.
- Replace one overloaded chart with small multiples, and legends with direct
  labels.
- Use colour to carry meaning, with palettes that survive colour-blind viewers
  and a greyscale printer.
- Strip an Excel default chart down to its data, step by step.
- Build KPI tiles that carry a comparison and a trend rather than a bare number.
- Wire slicers and a timeline to several PivotTables so one control drives the
  whole dashboard.
- Write a chart title that states the finding.
- Export to PowerPoint without the formatting disintegrating.

## The business question

The AdventureWorks leadership team meets monthly. The brief:

> We get a forty-page pack every month and nobody reads it. Give us one screen
> we can look at in the first five minutes of the meeting that tells us how the
> business is doing, where the problem is this month, and whether last month's
> problem got better. We want to be able to look at it by territory ourselves
> without asking you.

Read that carefully, because it contains three different artefacts and the
request is only for one of them:

- **A dashboard** answers "how are we doing?" repeatedly, over time, for a
  user who will interact with it. It is a monitoring instrument. It has no
  argument; it has state.
- **An exhibit** answers one question once, as part of an argument. It belongs
  in a deck, it has a headline that states a finding, and it is designed to be
  understood with no interaction at all.
- **A report** documents. It is complete, it is reference material, and nobody
  reads it front to back, which is fine — that is not what it is for.

The forty-page pack is a report doing a dashboard's job. The fix is not a
shorter report. It is one dashboard, plus the discipline to let the report stay
long and unread in the appendix where it belongs.

The last clause — "without asking you" — is the one that dictates the build. It
means slicers, not screenshots, and it means the thing has to be robust enough
to be used by someone who will not call you when a filter looks odd.

## Data you will use

From [the data dictionary](../data/adventureworks-data-dictionary.md):

- The unioned **`Sales`** table (internet + reseller) from Chapter 7, with
  `SalesAmount`, `TotalProductCost`, `OrderQuantity`, `SalesOrderNumber`,
  `OrderDateKey`, `SalesTerritoryKey`, `ProductKey` and `Channel`.
- **`DimDate`** for `CalendarYear`, `MonthNumberOfYear`, `EnglishMonthName`
  and `FullDateAlternateKey`. The timeline slicer needs a real date column,
  which is why `FullDateAlternateKey` matters and `OrderDateKey` alone will not
  do.
- **`DimProductCategory`** via the subcategory bridge, for the four categories.
- **`DimSalesTerritory`** for `SalesTerritoryRegion` and
  `SalesTerritoryGroup`. Exclude the `NA` row (key 11) in the model, once, so
  it cannot leak into a filtered view later.

Measures, defined once in the data model as in Chapter 8, because a dashboard
with its metrics redefined in nine different pivots is a dashboard that will
eventually contradict itself:

```
Revenue        = SUM(Sales[SalesAmount])
Cost           = SUM(Sales[TotalProductCost])
Margin         = [Revenue] - [Cost]
Margin %       = DIVIDE([Margin], [Revenue])
Orders         = DISTINCTCOUNT(Sales[SalesOrderNumber])
Revenue PY     = CALCULATE([Revenue], SAMEPERIODLASTYEAR(Calendar[Date]))
Revenue YoY %  = DIVIDE([Revenue] - [Revenue PY], [Revenue PY])
```

The final month in the dataset is truncated. On a dashboard that matters more
than anywhere else, because a trend line that falls off a cliff at the right
edge will be read as a collapse by everyone who looks at it for five seconds.
Either exclude incomplete periods in the model or mark the last point
explicitly as partial. Do not rely on a footnote; nobody reads the footnote.

## Core concepts

### Choose the form from the question

There are five questions charts answer, and for each a short list of forms that
hold up in front of a client:

| The question | Forms that work | Avoid |
|---|---|---|
| How has it changed over time? | line, column (for discrete periods), slope chart | stacked area with many series |
| What is it made of? | stacked bar for 2–4 parts, simple bar of parts, waterfall for a bridge | pie, donut, treemap |
| How do these compare? | horizontal bar, sorted; dot plot | radar, 3-D column |
| How is it distributed? | histogram, box plot, strip plot | a mean with no spread |
| Are these two things related? | scatter, with a fitted line if you mean it | dual-axis line pair |

Two rules sit behind the table. **Length is the most accurately read visual
channel; angle and area are among the worst** — which is the whole case against
pies. And **comparison needs a shared baseline**: bars from zero compare
cleanly, bars floating in a stack beyond the first segment do not.

For AdventureWorks specifically: revenue over time is a line; category
contribution is a sorted horizontal bar; the margin bridge between two years is
a waterfall; the territory view is a sorted bar, not a map, unless geography is
genuinely the point.

### The four things that cause arguments

**Pie charts.** Readers cannot compare angles, cannot rank slices of similar
size, and cannot read a slice against last year. Two categories is fine, and
even then a bar is no worse. Four AdventureWorks categories in a pie means
nobody can tell whether Clothing beat Accessories. Use a sorted bar and the
question answers itself.

**Dual axes.** Two series, two scales, two y-axes — and the crossing point
means nothing, because you chose it when you set the scales. Anyone who wants
to tell a different story can tell it by rescaling. The honest alternatives:
index both series to 100 at a base period and use one axis, or stack two charts
with a shared x-axis. The only dual axis I will defend is a count and a rate
where the rate is plotted as a line over columns and both axes start at zero.

**3-D effects, shadows, gradients.** They add ink that encodes nothing and they
actively distort: in a 3-D pie the front slice looks bigger, and in a 3-D column
chart the plotted value sits at the back of the column while the eye reads the
front. There is no analytical case for any of it.

**Truncated axes.** Starting a bar chart's axis at 90 instead of 0 makes a 2%
difference look like a 10x difference, and bar length is the thing readers trust
most. For bars, always zero. For *lines*, truncation is normal and often
necessary — a line encodes change through slope, not length, and forcing a
margin-percentage line to start at zero can flatten a genuine movement into
invisibility. The rule, then: **bars from zero, lines from wherever the data
lives, and label the axis clearly either way.** If you truncate a line axis on a
sensitive metric, say so in the subtitle.

### Small multiples instead of one overloaded chart

A line chart with seven series becomes a nest. The reader spends their attention
matching colours to a legend and has none left for the data.

Small multiples solve it: the same chart repeated once per category, same axes,
same scale, arranged in a grid. The eye compares shapes across panels in a way
it cannot compare tangled lines. Four AdventureWorks categories as four small
line charts, identical y-axis, is strictly more readable than one chart with
four lines — and it stays readable at ten categories, where the single chart
died at five.

In Excel: build one chart properly, copy it three times, repoint each copy's
series, and then — the step people skip — **fix the y-axis maximum to the same
value on all four**. Panels with auto-scaled axes look comparable and are not,
which is worse than an obviously messy chart.

### Direct labelling beats legends

A legend asks the reader to do a lookup: colour → name → back to the line. Every
lookup costs attention. Put the series name at the end of its line, in the
line's colour, and the lookup disappears.

Excel does this through data labels on the last point only: add data labels,
delete all but the final one, set it to show Series Name, and colour the text to
match. It takes a minute per series and it is the single highest-return
formatting change in this chapter.

The same logic kills most gridlines and most axis labels. If you label the last
point of each line and the first and last x-values, you often need no gridlines
at all.

### Colour carries meaning or it carries nothing

Three rules.

**Grey is a colour, and usually the right one.** Plot the context in grey and
the one series you are making a point about in a strong colour. A chart where
one line is blue and the other six are grey has already made its argument.

**One meaning per colour, across the whole deck.** If Bikes is dark blue on
slide 4, it is dark blue on slide 11. Inconsistent category colours force the
reader to re-learn the chart every page, and they will not.

**Use a scale for ordered data and a palette for unordered data.** Revenue
intensity across months is sequential: one hue, light to dark. Over- and
under-performance against target is diverging: two hues meeting at a neutral
midpoint. Four product categories are unordered: four distinguishable hues, no
implied ranking.

Then two accessibility checks that take thirty seconds each. Roughly one man in
twelve has some form of red-green colour deficiency, so **never encode
good/bad as red versus green alone** — add a shape, a sign, a position, or use
blue/orange, which survives nearly every form of deficiency. And print the page
in greyscale before the meeting: if two series become the same grey, you needed
a difference in lightness, not just in hue. Client decks get printed, and they
get printed in black and white.

### Data-ink, applied concretely

Every mark on a chart either encodes data or competes with it. The Excel default
is full of marks that compete. Here is the sequence, in order, on a default
clustered column chart:

1. Delete the chart border and the plot-area fill. Grey backgrounds exist to be
   deleted.
2. Delete horizontal gridlines, or reduce them to very light grey if the reader
   must read values off the axis.
3. Delete the y-axis entirely and add data labels, if there are fewer than about
   eight bars. Labels are more precise than an axis and use less ink.
4. Delete the tick marks.
5. Set the gap width to around 50% so bars are wider than the gaps. Excel's
   default makes thin bars that underplay the data.
6. Sort the bars by value unless the category order is inherently meaningful
   (months are; product categories are not).
7. Recolour: everything grey, the point of interest in one accent colour.
8. Replace the title. This is step 8 only because the formatting is quick; it is
   the most important step.
9. Add a source line in small grey text: table, filter, date of extract.

That sequence turns the default into something publishable in about two
minutes, and it is worth doing by hand enough times that you can then save it
as a chart template (right-click → Save as Template) and apply it in one click
afterwards.

### Conditional formatting, sparingly

Data bars inside a table are the best use of conditional formatting in Excel:
they give a bar chart's comparison without leaving the table. Set the minimum
to 0 explicitly rather than Automatic, or the shortest bar becomes invisible and
the comparison is a lie.

Colour scales work on a dense grid — a month-by-category heatmap of YoY growth,
diverging around zero, reads instantly. They do not work on a single column of
twelve numbers, where they just make the table loud.

Icon sets are the weakest tool here. Three traffic lights throw away magnitude,
and the thresholds are invisible to the reader. If you use them, state the
thresholds on the page.

Sparklines (Insert → Sparklines) put a trend in a single cell, which is exactly
what a KPI table wants: the number, the comparison, and the shape, in one row.

### KPI tiles need a comparison and a trend

A tile that says `$29.4M` is nearly useless. The reader's immediate question is
"compared to what?" — and if they have to ask, the tile failed.

A usable tile has four parts: the metric name, the value, a comparison with
direction, and a trend. In cells:

`="Revenue  " & TEXT([Revenue], "$#,##0.0,,") & "M"`

`="vs PY  " & TEXT([Revenue YoY %], "+0.0%;-0.0%") & IF([Revenue YoY %]>=0, " ▲", " ▼")`

with a sparkline in the cell beside it. Colour the comparison text with a
conditional rule — but pair the colour with the arrow, so it still reads in
greyscale and for a colour-blind viewer. Four to six tiles across the top of the
dashboard; more than that and they stop being a summary.

### Layout: a grid and a reading order

Western readers scan top-left first, then right, then down. So the most
important thing goes top-left, and that is almost always the headline finding or
the KPI row — not the company logo and not the filter panel.

Build on a grid. Set all column widths in the dashboard area to something
uniform and narrow (say 20 pixels), then size every chart to span a whole number
of columns. The result lines up, and misalignment is the main thing that makes
an otherwise competent dashboard look amateur.

Practical mechanics: one `Dashboard` sheet holding only objects, pivots living
on a hidden `PivotData` sheet, gridlines turned off (View → Gridlines), freeze
panes if it scrolls at all — and ideally it does not, because "one screen" was
the brief.

### Write the title as the finding

This is the highest-leverage sentence in the chapter.

A title like "Margin by Category" names the axes, which the reader can already
see. A title like "Accessories drove all of 2013's margin growth" tells them
what to conclude. The first makes the reader do your job; the second makes your
job checkable, which is better for both of you.

For a *dashboard*, titles stay neutral and descriptive — the finding changes
every month and the chart cannot know it. For an *exhibit* in a deck, the title
is the finding, always, and the body of the chart is the evidence for it. That
is the real difference between the two artefacts, and it is why you cannot
screenshot a dashboard into a deck and call it an exhibit.

### Performance, refresh and export

A dashboard that takes forty seconds to refresh will not be used. Keep the data
model lean: drop columns you do not need in Power Query rather than loading and
hiding them, prefer measures to calculated columns, and avoid volatile functions
(`NOW`, `OFFSET`, `INDIRECT`) in anything that feeds a chart. Set pivots to
refresh on open (PivotTable Options → Data → Refresh data when opening the
file) so the user never sees stale numbers without knowing it.

For PowerPoint, paste as a **picture (enhanced metafile)** for anything final:
it cannot break, cannot re-theme itself, and cannot show a client your source
data. Paste-link only when the deck is genuinely live and you control both
files — and expect the link to break the first time someone renames the
workbook. Either way set the chart's font to match the deck's before export, and
check it at presentation size, not at 100% zoom on your monitor. Text that is
legible on a laptop is frequently illegible projected.

## Walkthrough

### Step 1 — Prepare the engine, not the picture

Create a hidden `PivotData` sheet. Build the pivots the dashboard needs, each
one minimal:

- Revenue and margin by `CalendarYear` and `MonthNumberOfYear` (for the trend).
- Revenue, margin and margin % by product category (for the breakdown).
- Revenue and YoY % by `SalesTerritoryRegion` (for the territory view).
- A single-cell pivot per KPI tile.

Build every pivot from the same data model connection. This is what makes one
slicer able to drive all of them.

### Step 2 — Build the KPI row

In `B2:I4` of the `Dashboard` sheet, lay out five tiles: Revenue, Margin,
Margin %, Orders, Revenue YoY %. Each tile reads from its pivot via
`GETPIVOTDATA` or a direct cube formula, and each shows value, comparison and
sparkline:

`=GETPIVOTDATA("Revenue", PivotData!$A$3)`

`=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Revenue]", Slicer_Territory)`

The second form is worth the trouble: `CUBEVALUE` respects the slicer by name,
needs no pivot on screen, and does not shift when a pivot resizes. Then the
comparison line:

`=TEXT(C2, "$#,##0,,\"M\"") & "   " & TEXT(C3, "+0.0%;-0.0%") & " vs PY"`

and a sparkline over the twelve monthly revenue cells on `PivotData`.

Format the tiles as cells with a light fill and a thin border, not as text
boxes. Cells align to the grid; floating text boxes do not, and they drift every
time someone inserts a row.

### Step 3 — The trend chart

Insert a line chart on the monthly revenue pivot. Then apply the data-ink
sequence: delete the border, delete the fill, keep very light gridlines (a
twelve-point monthly series does need them), delete the legend, label the final
point of each year's line directly, and set the y-axis to start at zero — this
is revenue, a magnitude, and the reader will compare heights.

Handle the truncated final month explicitly: either stop the series at the last
complete month, or plot the partial point as a hollow marker on a dashed segment
and label it "partial month". Choose one and be consistent.

Title it descriptively for the dashboard: "Monthly revenue, last 24 months".

### Step 4 — The category breakdown

A horizontal bar chart, sorted descending by revenue. Add margin % as data
labels on each bar rather than as a second series on a second axis — the
comparison the reader wants is "big but thin versus small but fat", and two
numbers on one bar deliver it without a dual axis.

Colour all bars grey except the one with the worst margin %, in your accent
colour. The chart now has an opinion without needing a sentence.

### Step 5 — The territory view

Sorted horizontal bars of revenue, with a small second panel of YoY % as a
diverging bar around zero. Resist the map. A choropleth of ten territories
across three continents wastes most of its area on empty ocean and encodes
magnitude in a channel — area of an irregular region — that readers cannot
compare. Use a map only when the question is genuinely spatial, such as
proximity or clustering.

### Step 6 — Wire one slicer to everything

Insert a slicer on `SalesTerritoryGroup` and a timeline on
`Calendar[FullDateAlternateKey]`. Then the step that makes it a dashboard rather
than four charts: right-click the slicer → **Report Connections**, and tick
every pivot. Repeat for the timeline.

Test it properly. Click each slicer value and check that all five KPI tiles and
all three charts move together. A dashboard where one pivot is unconnected is
worse than no dashboard: it shows two inconsistent numbers on the same screen
and destroys trust in everything else on the page.

Place the slicers top-right or in a left-hand rail — out of the primary reading
path, but visibly available. Then protect the sheet, leaving the slicers
unlocked, so a user cannot accidentally overwrite a tile formula.

### Step 7 — Critique and fix the ugly version

Build the default version deliberately: select the category revenue data, press
`Alt+F1`, and look at what Excel gives you. Typically: a 3-D-ish clustered
column chart, a grey plot area, horizontal gridlines, a legend for a single
series, a title reading "Total", unsorted categories, thin bars, and a palette
of six colours for four bars.

Now list the faults and fix them in order: single-series legend (delete, it
encodes nothing), unsorted categories (sort descending), thin bars (gap width to
50%), gridlines plus axis plus labels (keep one, delete two), six colours for
one meaning (grey plus one accent), 3-D (remove), title (replace with the
finding).

Do this once by hand. Then save the result as a chart template. The exercise
below asks you to write the critique out, because being able to say *why* a
chart is bad is what lets you push back when a client insists on the pie.

## Common pitfalls

- **A pie chart with more than two slices.** Nobody can rank the slices, which
  is the only question anyone had.
- **Dual axes to force a correlation.** The crossing point is an artefact of
  your scaling choice. Index to 100 instead.
- **A bar chart axis that does not start at zero.** Bar length is the channel
  readers trust most; truncating it misleads them most.
- **Auto-scaled axes across small multiples.** The panels look comparable and
  are not.
- **A legend where direct labels would do.** Every legend lookup is attention
  taken from the data.
- **Red/green as the only encoding of good and bad.** It fails for a tenth of
  your audience and for every greyscale printout.
- **Data bars with Automatic minimum.** The smallest value gets a bar of width
  zero and vanishes.
- **A KPI tile with no comparison.** `$29.4M` answers nothing.
- **An unconnected pivot.** Two contradictory numbers on one screen, and the
  whole dashboard loses credibility.
- **Floating text boxes instead of cells.** They drift, they misalign, and they
  do not print where you left them.
- **The truncated final month left unmarked.** It reads as a crash, and someone
  will raise it in the meeting before you can.
- **Exhibit titles that name the axes.** "Margin by category" makes the reader
  do the analysis you were paid for.
- **Pasting a live-linked chart into a client deck.** It breaks, and until it
  breaks it carries your whole source workbook inside the file.
- **Refresh left manual.** The user eventually presents last quarter's numbers,
  and it will be described as your error.

## Exercises

1. **Classify the artefact.** For each of these requests, state whether the
   right deliverable is a dashboard, an exhibit or a report, and give the single
   reason: (a) the monthly leadership meeting; (b) a board paper arguing for the
   $4.0m growth investment from Chapter 13; (c) the audit trail for how revenue
   was calculated; (d) a weekly territory-performance email to ten regional
   managers.

2. **Critique the default.** Select revenue by product category and press
   `Alt+F1`. Write a numbered critique of at least eight specific faults in
   Excel's default output, each with the fix. Then produce the corrected chart
   and save it as a chart template.

3. **Small multiples.** Build four monthly-revenue line charts, one per product
   category, on a shared and explicitly fixed y-axis, arranged in a 2×2 grid
   with direct labels and no legends. Then build the single four-line version
   of the same data. Which comparison is easier to make in each, and what does
   that tell you about when the overloaded version is still the right choice?

4. **KPI tiles.** Build a five-tile KPI row for Revenue, Margin, Margin %,
   Orders and Revenue YoY %, each with value, a prior-year comparison with
   direction, and a sparkline. Tiles must be cells, not text boxes, and the
   direction must be readable in greyscale.

5. **The full dashboard.** Assemble the one-screen dashboard: KPI row, monthly
   trend, sorted category breakdown with margin % labels, and territory view —
   all on a uniform grid, gridlines off, driven by one `SalesTerritoryGroup`
   slicer and one timeline connected to every pivot. Verify connection by
   clicking through every slicer value and confirming all eight objects move.
   Handle the truncated final month visibly.

6. **Accessibility pass.** Print your dashboard to greyscale PDF and simulate a
   red-green colour deficiency (a colour-blindness simulator, or convert to
   greyscale and check lightness separation). List every place meaning was lost
   and fix each one without adding a legend.

7. **Title rewrite.** Take the three charts from your dashboard and write two
   titles for each: the neutral dashboard title, and the finding-first exhibit
   title you would use if that chart appeared in the Chapter 13 investment deck.
   Explain in one sentence per chart why the two differ.

## Solutions

**Exercise 1.** (a) Dashboard — recurring, interactive, state rather than
argument. (b) Exhibit — a single argument made once, with a finding-first
headline. (c) Report — completeness matters, readability does not. (d) This one
is the interesting case: it is an exhibit *series*, not a dashboard, because
email recipients cannot interact. Push for a dashboard link with a short
narrative email, or accept that you are now committed to producing ten tailored
exhibits every week forever.

**Exercise 2.** The expected list: single-series legend, unsorted categories,
excessive gap width, redundant gridlines plus y-axis plus labels, multiple hues
encoding nothing, grey plot-area fill, chart border, tick marks, a title naming
the variable, and no source note. The fix for each is in the data-ink sequence
above. The point of writing it out is that a client who wants a 3-D pie needs an
answer better than "it looks bad" — "readers cannot rank slices by angle, and
the 3-D projection makes the front slice look larger than it is" ends the
conversation.

**Exercise 3.** Small multiples win when the question is "what is the shape of
each category?" and lose when it is "which category is largest?" — the single
overloaded chart puts the series on a shared axis where magnitudes compare
directly. So the honest answer is that the choice follows the question, and the
overloaded version is right when there are two or three series and ranking is
the point. With four or more series, or when the shapes differ more than the
levels, go to panels. If your four panels look identical, that is itself a
finding: the categories move together, so a single total line plus a mix chart
tells the story better than either version.

**Exercise 4.** The common failure is building tiles as text boxes because they
are easier to position. They drift, they do not print in place, and they break
when rows are inserted. Build them as merged or bordered cell blocks on the
grid. The second common failure is colouring the comparison red or green with
no arrow — pair every colour with a glyph and the tile survives greyscale and
colour deficiency at no cost.

**Exercise 5.** Expect the slicer-connection test to fail at least once; a pivot
created from a different connection, or added after you set Report Connections,
will not be listed. This is the single most important test on the page, because
an unconnected pivot produces two contradictory numbers on one screen and the
client will remember that longer than any insight. On the truncated month, both
treatments are defensible — stopping the series is cleaner, marking it as
partial is more informative — but a dashboard that silently shows a cliff will
generate a question about data quality in every meeting until someone fixes it.

**Exercise 6.** The usual losses are a red-versus-green KPI arrow pair, two
similar-lightness category hues that collapse to the same grey, and a diverging
colour scale whose midpoint disappears. Fixes: pair colour with glyph or sign;
vary lightness as well as hue across categories; and for the diverging scale,
rely on the bar's direction from zero rather than on its colour.

**Exercise 7.** Dashboard: "Monthly revenue, last 24 months". Exhibit: "Revenue
growth stalled from Q3 and has not recovered". The difference is that the
dashboard cannot know next month's finding, so a finding-first title there would
be wrong half the time — while an exhibit exists precisely to assert one thing,
and a neutral title there wastes the most-read line on the page.

## Key takeaways

- Decide which artefact you are building first. A dashboard monitors, an exhibit
  argues, a report documents, and they are not interchangeable.
- Pick the chart form from the question. Length encodes magnitude well; angle
  and area do not.
- Bars start at zero, always. Lines may be truncated — label the axis and say so
  on anything sensitive.
- Small multiples with a fixed shared axis beat one overloaded chart past three
  or four series.
- Direct labels beat legends. Grey plus one accent colour beats a palette.
- Never encode good and bad in red and green alone, and check the page in
  greyscale before the meeting.
- A KPI tile needs a comparison and a trend, not just a number.
- One slicer, connected to every pivot, tested by clicking through every value.
  An unconnected pivot costs more credibility than the dashboard earns.
- Build on a grid, in cells, with gridlines off, and put the most important
  thing top-left.
- Write exhibit titles as the finding. It is the most-read text in your work and
  usually the least considered.
