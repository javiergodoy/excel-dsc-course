# Chapter 15 — From Analysis to Recommendation

Everything in the previous fourteen chapters was instrumental. You built a
clean data model, joined it correctly, aggregated it, forecast it, regressed it
and segmented it. None of that is the work. The work is the sentence you say to
the AdventureWorks management team when they ask where to put the next dollar,
and the fact that the sentence survives being argued with.

This chapter is about the conversion: analysis into recommendation. It is the
step where most technically strong analysts lose the room, because the skills
are different. Analysis rewards thoroughness. Recommendation rewards
compression, sequence and nerve.

## Learning objectives

By the end of this chapter you will be able to:

- Frame a client question into a decision before touching the data, and
  recognise when the question you were asked is not the question that needs
  answering.
- Decompose a vague growth question into a MECE issue tree whose branches are
  separately testable.
- Run hypothesis-driven analysis, and tell it apart from exploratory fishing
  that happens to produce charts.
- Write a day-one answer and then spend your effort trying to break it.
- Apply the "so what, how much" test to every finding, and discard findings
  that fail it however interesting they are.
- Build a storyboard whose slide titles, read in sequence and with no exhibits
  at all, constitute the argument.
- Make an exhibit presentable: source, units, period, filters, sample size.
- Answer "where did that number come from" in a live meeting without losing
  control of the discussion.
- Triangulate a finding from two independent angles before presenting it.
- Convert a recommendation into an action, an owner, a cost and an expected
  return, with the risks and the conditions under which it fails.

## The business question

AdventureWorks' management team has approved an incremental growth investment
budget. The sponsor's words, in the kickoff meeting, were:

> "We have budget for one real push next year. Sales thinks we should hire
> reseller account managers in Europe. Marketing wants to spend it on the
> website. Product wants a new mountain bike line. Tell us where to put it."

That is a decision, not a research topic, and it has three properties that
should shape everything you do:

1. **It is a choice among alternatives, not an open question.** You are not
   being asked "how is the business doing". You are being asked to rank three
   named options, and you are permitted to propose a fourth if you can defend
   it.
2. **It is budget-constrained and single-shot.** "One real push" means the
   answer cannot be "do all three, a bit". Recommending a portfolio when the
   client has told you they can only execute one thing is a failure to listen.
3. **Each advocate has a stake.** Sales, Marketing and Product are not neutral
   parties. Your analysis will be read as taking someone's side, which means
   the arithmetic has to be beyond dispute even where the judgment is not.

Before any data work, write the decision down in one sentence, in the client's
language, and get it confirmed:

> *Which of reseller coverage in Europe, direct-channel investment, or new
> product development will return the most incremental gross margin per dollar
> invested over the next three years?*

Notice what that sentence did. It set the metric (incremental gross margin, not
revenue), the horizon (three years), and the comparison basis (per dollar
invested). Half the arguments you will have later are pre-empted by having
agreed those three things in week one. If the sponsor pushes back on any of
them — "we care about revenue growth, not margin" — you have learned something
important for free, and far more cheaply than learning it in the final meeting.

## Data you will use

This chapter consumes the outputs of Parts II and III rather than returning to
raw tables. Specifically:

- The unioned company sales table from Chapter 7, combining
  `FactInternetSales` and `FactResellerSales` with a `Channel` column. Revenue
  is `SalesAmount`; margin is `SalesAmount − TotalProductCost`, as fixed in
  [the data dictionary](../data/adventureworks-data-dictionary.md).
- Category and subcategory from `DimProductCategory` and
  `DimProductSubcategory` via `DimProduct`.
- `DimSalesTerritory.SalesTerritoryGroup` and `SalesTerritoryRegion` for
  geography, with the `NA` territory row excluded explicitly.
- `DimDate.CalendarYear` and `MonthNumberOfYear` for the trend and seasonality
  work from Chapter 10.
- The customer segmentation and cohort output from Chapter 12, keyed on
  `DimCustomer.CustomerKey`.
- The driver regression from Chapter 11.

One discipline carries over from Chapter 16's argument and is worth stating
here: every number that reaches a slide must be reproducible from a named cell
in a named sheet. If you cannot point at it, do not present it.

## Core concepts

### Frame the decision, not the topic

The single most common engagement failure is answering the question as asked
when the question as asked is wrong. The sponsor said "where should we put the
budget". Decomposed, that contains a prior question: *is there a growth problem
at all, and if so, where does it live?* If AdventureWorks' direct channel is
growing at 40% and reseller sales are flat, the three options are not
symmetrical, and ranking them without saying so would be technically correct and
practically useless.

Reframing is legitimate. Silently substituting your own question is not. The
move is to say, in the first meeting: "To rank those three, I need to establish
where growth currently comes from and where it has stalled. That is the first
exhibit, and it may change the shortlist."

### The issue tree, and what MECE actually buys you

MECE — mutually exclusive, collectively exhaustive — is taught as a tidiness
rule. Its real function is to make a negative result informative. If your
branches overlap, ruling one out tells you nothing, because the effect might
live in the overlap. If they do not cover the space, ruling all of them out
leaves you with nothing to say.

For AdventureWorks:

```
Where should the next growth dollar go?
├── Sell more to existing customers
│   ├── Higher order value (attach rate, basket size)
│   └── Higher order frequency (repeat purchase, retention)
├── Sell to new customers
│   ├── In existing geographies (penetration)
│   └── In new geographies (expansion)
└── Sell different things
    ├── New products in existing categories
    └── New categories
```

Each leaf is separately testable with the data you have, and the leaves do not
overlap: a sale is either to a customer you had or one you did not; it is either
in a geography you served or not. Note that the three sponsor options map onto
leaves — reseller coverage in Europe is *new customers in existing
geographies*, a new bike line is *new products in existing categories* — which
lets you rank them against alternatives they did not think to propose.

### Hypothesis-driven, not exploratory

Exploratory analysis asks "what is in this data". Hypothesis-driven analysis
asks "if X were true, what would I expect to see, and do I see it". The
difference is not rigour; it is that the second one terminates.

Write the hypothesis as a falsifiable statement with a threshold:

> *H1: AdventureWorks' direct channel has a basket-size problem — more than
> 90% of internet orders contain exactly one line — and fixing attach rate is
> worth more than any geographic expansion.*

Now the test is mechanical:

```excel
=LET(
   lines, COUNTIFS(Sales[Channel],"Internet"),
   orders, COUNTA(UNIQUE(FILTER(Sales[SalesOrderNumber],Sales[Channel]="Internet"))),
   lines/orders
)
```

If average lines per order comes back near 1.0, H1's premise holds and the
opportunity is real. If it comes back at 2.8, H1 is dead and you have spent
twenty minutes instead of two weeks. Killing your own hypothesis quickly is the
highest-return activity in consulting, and the hardest to make yourself do.

### The day-one answer

On day one, with no analysis, write down what you think the answer is. Then
spend the engagement trying to break it.

This feels backwards and it is not. The day-one answer does three things. It
exposes your priors so you can test them rather than smuggle them in. It gives
the team a hypothesis to organise around, so analysis is targeted instead of
comprehensive. And it means that if the engagement is cut short — which happens
— you have an answer, labelled with its confidence, rather than a folder of
clean data.

A day-one answer for AdventureWorks might be: *"The push should go to
accessories attach rate in the direct channel, because bikes are a
once-every-five-years purchase and accessory margin percentage is far higher
than bike margin percentage."* That is a guess. It is a useful guess because
it is specific enough to be wrong in a way you can detect.

### The "so what, how much" test

Every finding gets two questions, in order.

**So what?** What decision changes if this is true? If no decision changes, it
is trivia. "Thursday is our strongest order day" is trivia unless someone can
act on day-of-week.

**How much?** Put a number on it. Not a direction, a magnitude. "Attach rate is
low" fails. "Raising average lines per internet order from roughly 1.0 to 1.3
would add on the order of several million dollars of gross margin annually at
current traffic" passes, because now it can be compared against the cost of
doing it.

Findings that pass both go in the deck. Findings that pass one go in the
appendix. Findings that pass neither get deleted, however much work they took.
Sunk cost is not an argument for slide space.

### The pyramid principle

Structure the deck as a pyramid: the answer at the top, the two to four reasons
that support it on the next level, the evidence for each reason below that.
Every level answers the question raised by the level above.

Most analysts build the inverted pyramid — context, then method, then data, then
conclusion — because that is the order in which they did the work. It is the
wrong order for a reader who may only get through three slides, and it is the
wrong order for a meeting that may be cut to fifteen minutes. Lead with the
answer. The reasoning is what you say when they push back, not what you say
first.

### The storyboard

A storyboard is one line per slide, written before any exhibit is built. Each
line is the slide's title, and each title is a full sentence that makes a claim.

Titles that are labels — "Revenue by Category", "Customer Segmentation" — force
the reader to do your interpretive work. Titles that are sentences — "Bikes
drive 70% of revenue but under half of gross margin" — do the work for them,
and have a second property: read in sequence, with the exhibits removed, they
form a coherent argument. That is the test. If the titles alone do not argue
the case, the deck does not either, and no amount of chart polish will fix it.

### What makes an exhibit presentable

An exhibit is presentable when a stranger can read it unaided and a skeptic
cannot undermine it with a procedural question. Minimum furniture:

- **A sentence title** making the claim.
- **Units**, explicitly: "$m", "% of gross margin", "units".
- **Period**, explicitly: "FY2013 (calendar)". Chapter 2 warned you that
  calendar and fiscal years differ in `DimDate`; the exhibit must say which.
- **Filters applied**, stated: "Excludes territory NA (unassigned, ~0.3% of
  revenue). Excludes partial final month."
- **Sample size** where it is a sample: "n = 412 repeat customers".
- **Source note** naming the tables: "Source: FactInternetSales,
  FactResellerSales, DimProduct."

The source note is not bureaucracy. It is the thing that lets you answer the
hardest question in the meeting in one sentence.

### "Where did that number come from?"

When a CFO asks this, they are not usually attacking the number. They are
testing whether you know it. The answer has three parts, delivered in about
fifteen seconds:

1. **The source.** "Reseller and internet sales fact tables, FY2013."
2. **The definition.** "Gross margin is sales amount less total product cost;
   it is not in the data, we derive it."
3. **The offer.** "The cell is on the Margin_by_Category sheet, I can show the
   trace."

What you must not do is improvise a reconstruction out loud. If you cannot
answer in those three parts, say "I will come back to you with the trace this
afternoon" and write it down visibly. One deferred answer costs you nothing. One
wrong answer defended confidently costs you the whole deck, because from that
point every number is suspect.

### Triangulation

Before a finding goes on a slide, derive it a second way, independently. Two
methods agreeing is not proof, but two methods disagreeing is a guaranteed
problem you would otherwise have presented.

For "accessories carry much higher percentage margin than bikes":

```excel
=SUMIFS(Sales[SalesAmount],Sales[Category],"Accessories")
 -SUMIFS(Sales[TotalProductCost],Sales[Category],"Accessories")
```

as a share of the same denominator, against a product-level check that avoids
the fact table entirely:

```excel
=AVERAGE(IF(Products[Category]="Accessories",
   (Products[ListPrice]-Products[StandardCost])/Products[ListPrice]))
```

The first is realised margin on transactions; the second is theoretical margin
on the catalogue. They will not match — discounting, mix and the versioned
`DimProduct` rows guarantee a gap — but they should agree on direction and
rough magnitude. If realised margin percentage is far below catalogue margin
percentage, you have found discounting, which is itself a finding.

### From recommendation to action

A recommendation that cannot be executed is an opinion. Convert it into four
fields:

- **Action**: what someone does on Monday.
- **Owner**: a named role, not "the business".
- **Cost**: investment required, with its basis.
- **Expected return**: incremental gross margin, with the arithmetic shown and
  the sensitivity attached.

Then write the two paragraphs most analysts skip.

**What would have to be true.** State the conditions your number depends on:
that traffic does not fall, that attach-rate uplift of the assumed size is
achievable, that accessory supply can meet it. Stating these is not hedging. It
converts an argument about your conclusion into an argument about specific
assumptions, which is a far better argument to have, and one the client is
better placed than you to resolve.

**What would make this wrong.** Name the failure modes and the leading
indicators that would show them early. A recommendation with a named kill
criterion is more credible than one without, not less.

### The number you choose not to show

Every deck involves selection, and selection is where the ethics live. Some
rules that have survived contact with real engagements:

- If a number materially weakens your recommendation, it goes in the deck, not
  the appendix. Burying it does not make it go away; it makes its eventual
  discovery your credibility problem.
- Never choose a period, a filter or an axis range because it makes the chart
  more persuasive. Choose it for an analytical reason you can state, and state
  it on the exhibit.
- Small-sample findings get their n shown. The AdventureWorks promotion data is
  the live example: almost every row is `No Discount`, so promotion-effect
  findings rest on a handful of transactions. That is a finding — "we cannot
  evaluate promotion effectiveness from this data" — and presenting it honestly
  is more valuable than a confident number built on forty rows.
- If a sponsor asks you to remove a number because of what it implies, that is
  a conversation to have explicitly, not a silent edit.

## Walkthrough

### Step 1 — Write the storyboard before building anything

Open a blank sheet, one row per slide, columns: `Slide`, `Title sentence`,
`Exhibit type`, `Source sheet`, `Status`. Six exhibits.

The titles, which are the deliverable at this stage:

1. **"AdventureWorks grew 24% last year, but all of the growth came from
   resellers in Europe — the direct channel was flat."**
2. **"Bikes generate roughly 70% of revenue and under half of gross margin;
   accessories are the inverse."**
3. **"Nine in ten internet orders contain a single line, so almost no
   accessory margin is being attached to bike sales."**
4. **"Customers who buy an accessory with their first bike are materially more
   likely to buy again within 12 months."**
5. **"Closing the attach-rate gap is worth more incremental gross margin than
   either proposed alternative, at a fraction of the cost."**
6. **"Recommendation: fund direct-channel attach rate, with a European
   reseller pilot held as the fallback."**

Read those six sentences in order with no charts. They argue the case: here is
where growth is and is not, here is where margin actually lives, here is a
specific unexploited mechanism, here is evidence the mechanism compounds, here
is the size, here is the ask. That is the test passing.

Now notice what the storyboard did to your workplan. You need exactly four
analyses, and you know which. You are not going to build the geographic
heat map that nobody asked for.

### Step 2 — Build exhibit 3's number, properly

The attach-rate claim is load-bearing, so it gets the most care. Lines per
order, by year, excluding the truncated final month:

```excel
=LET(
   ch, "Internet",
   yr, $A5,
   rows, FILTER(Sales[SalesOrderNumber],
          (Sales[Channel]=ch)*(Sales[CalendarYear]=yr)*(Sales[IsPartialMonth]=0)),
   COUNTA(rows)/COUNTA(UNIQUE(rows))
)
```

Then the share of single-line orders, which is the more arresting statistic and
the one that belongs in the title:

```excel
=LET(
   ords, UNIQUE(FILTER(Sales[SalesOrderNumber],Sales[Channel]="Internet")),
   lines, COUNTIFS(Sales[SalesOrderNumber],ords),
   SUM(--(lines=1))/COUNTA(ords)
)
```

Triangulate against a PivotTable: order number in rows, count of
`SalesOrderLineNumber` in values, grouped. Two independent routes to the same
proportion, and a one-cell check that tells you loudly if they diverge:

```excel
=IF(ROUND(SinglelineShare_Formula,4)=ROUND(SinglelineShare_Pivot,4),"OK","FAIL")
```

Two independent routes to the same proportion. If they differ, you have a grain error — probably counting lines
where you meant orders, which is the error the data dictionary warned about in
its first paragraph on `FactInternetSales`.

### Step 3 — Size the recommendation

Sizing is where credibility is won or lost, because it is the only part the
client will arithmetic-check in the room. Build it as a visible input block,
blue for inputs, black for formulas, per Chapter 1. Accessory margin per line,
the one input that comes from data rather than judgment, is derived once and
referenced everywhere:

```excel
=LET(
   acc, FILTER(Sales[SalesAmount]-Sales[TotalProductCost],
         (Sales[Category]="Accessories")*(Sales[Channel]="Internet")),
   AVERAGE(acc)
)
```

| Input | Value | Basis |
|---|---|---|
| Internet orders, FY2013 | *from data* | `COUNTA(UNIQUE(...))`, Orders sheet |
| Current lines per order | *from data* | Exhibit 3 |
| Target lines per order | 1.30 | Assumption: benchmark attach uplift, stated |
| Average accessory line margin | *from data* | Margin sheet, accessories only |
| Addressable share of orders | 0.70 | Assumption: bike and clothing orders only |

```excel
=Orders*(Target_LPO-Current_LPO)*Addressable_Share*Accessory_Line_Margin
```

Then immediately build the sensitivity, because the first question will be
"what if 1.30 is optimistic":

```excel
=DataTable over Target_LPO = {1.10, 1.15, 1.20, 1.30, 1.40}
```

Present the range, not the point. "Between roughly X and Y million of
incremental annual gross margin, depending on how much of the attach gap we
close, against an implementation cost in the low hundreds of thousands." A
range with stated drivers is more persuasive than a single number, because the
single number invites the audience to attack its precision instead of engaging
with its magnitude.

Compare against the alternatives on the same basis — incremental gross margin
per dollar invested over three years — and show the comparison as a table, not
prose. Like-for-like comparison is the whole point; if you size your preferred
option carefully and the alternatives casually, you have rigged it and a sharp
reader will notice.

### Step 4 — Write the one-page executive summary

One page, four blocks, in this order.

**The answer (2 sentences).** "Fund direct-channel attach rate. It returns an
estimated 4–6x the gross margin per dollar of either alternative, and can be
tested within one quarter."

**Why (3 bullets, each a finding with a number).** Where growth is and is not;
where margin actually lives; the size of the attach gap.

**The ask (action, owner, cost, return).** "Approve $X for checkout
cross-sell and bundle pricing, owned by the Director of eCommerce, targeting
$Y incremental gross margin in year one."

**What would have to be true, and what would make this wrong (4 bullets).**
Traffic holds; attach uplift of the assumed magnitude is achievable; accessory
inventory can meet demand; discounting does not absorb the margin. Kill
criterion: if attach rate has not moved by 10 basis points after one quarter of
the pilot, stop and redirect to the reseller option.

If the summary page runs to two pages, the problem is almost never that there
is too much to say. It is that the answer has not been decided firmly enough to
compress.

## Common pitfalls

**Presenting the analysis rather than the answer.** The chronological order of
your work is the worst possible order for a deck. Nobody wants the journey.

**Label titles.** "Revenue by Channel" is a caption. If every title in the deck
could sit above a different chart without changing, you have written captions.

**Unsized findings.** "There is an opportunity in accessories" cannot be
prioritised against anything. If you cannot size it even roughly, say so and
say what it would take to size it.

**Asymmetric rigour.** Sizing your favoured option to three decimals and the
alternatives with a shrug. Reviewers detect this instantly and it reads as
advocacy, not analysis.

**Hiding the inconvenient number.** The flat direct channel in exhibit 1 is
awkward for a recommendation to invest in the direct channel. Lead with it and
explain why it strengthens the case — the channel is flat *because* nothing is
attached, not because demand is absent. A weakness you raise first is an
argument; one the client finds is a problem.

**Spurious precision.** "$4,732,118 of incremental margin" from a model with a
judgment-call attach-rate assumption. Round to the precision your weakest input
supports and no further.

**No kill criterion.** A recommendation that cannot fail cannot be managed. It
also signals that you have not thought about implementation.

**Confusing revenue and margin in the ask.** The sponsor's advocates will argue
in revenue because revenue is bigger. The metric was agreed in week one as
gross margin. Hold the line, and show both if you must, labelled.

## Exercises

1. **Reframe the question.** Write the AdventureWorks growth question as a
   one-sentence decision, specifying metric, horizon and comparison basis. Then
   write the prior question that must be answered first, and one sentence
   explaining to the sponsor why you are answering it.

2. **Build the issue tree.** Produce a MECE decomposition of "where should the
   next growth dollar go" to two levels. Map each of the three sponsor options
   onto a leaf, and identify one leaf nobody proposed that the data can test.

3. **State and kill a hypothesis.** Write three falsifiable hypotheses with
   numeric thresholds, each testable from the unioned sales table. Test all
   three. Report which died, with the formula and the result, and what each
   death tells you about where to focus.

4. **Apply "so what, how much".** Take five findings from your Part II and
   Part III work. For each, write the decision it would change and a magnitude
   estimate. Classify each as deck, appendix or delete, and justify every
   delete.

5. **Write the storyboard.** Produce six slide titles as full sentences for the
   growth recommendation. Then give the titles alone — no exhibits — to someone
   who has not seen your analysis and ask them to state your recommendation and
   its two main supports. Record where they misread you and fix the titles, not
   the charts.

6. **Size it, and size the alternatives.** Build the sizing block with a
   visible input table and a two-variable sensitivity. Size all three options
   on the same per-dollar basis. Write the three sentences you would say if the
   CFO asked "what is the single assumption this most depends on".

7. **Survive the question.** For the three largest numbers in your deck, write
   the fifteen-second source-definition-offer answer. Then have a colleague
   pick one at random and deliver it out loud without notes.

## Solutions

**Exercise 1.** A good reframe reads close to: *"Which of reseller coverage in
Europe, direct-channel investment, or new product development returns the most
incremental gross margin per dollar invested over three years?"* The prior
question is *"where is growth currently coming from, and where has it
stalled?"* The sentence to the sponsor is the honest one: ranking three options
requires knowing which parts of the business are already growing without
investment, since those may not need the budget. Choosing gross margin over
revenue is the substantive move, and it needs saying out loud — in
AdventureWorks the two metrics point at different categories, because bikes
dominate revenue while accessories dominate percentage margin. An engagement
that never agrees its metric will discover the disagreement in the final
meeting.

**Exercise 3.** The attach-rate hypothesis survives: internet orders in
AdventureWorks are overwhelmingly single-line, so average lines per order sits
close to 1.0 and the attach opportunity is real rather than assumed. A
"European customers are lower value" hypothesis typically dies once you compare
margin per order rather than revenue per order — the apparent gap is mix, not
value. A "promotions drive the growth" hypothesis dies on sample size:
`DimPromotion` key 1 is `No Discount` and covers nearly all rows, so there is
no analysable promotional variation. That third death is the instructive one.
It is not a failed analysis, it is a finding — the business cannot currently
measure promotion effectiveness — and it belongs in the deck as a
recommendation about instrumentation.

**Exercise 4.** The common error is classifying by effort rather than by
decision. A finding that took three days and changes nothing is a delete. The
day-of-week pattern is the canonical example: real, robust, and actionable by
nobody at AdventureWorks. The reseller-versus-direct margin gap is a deck
finding because it reframes which channel deserves investment. Category
seasonality is appendix: true, useful for planning timing, not load-bearing for
the recommendation.

**Exercise 5.** The usual failure is that readers can state the recommendation
but not its supports, which means the middle titles are descriptive rather than
argumentative. Compare "Internet orders contain few lines" with "Nine in ten
internet orders contain a single line, so almost no accessory margin attaches
to bike sales". Both are true; only the second carries the inference the next
slide depends on. A title that states a fact without its implication forces the
reader to do the inferential work, and in a fifteen-minute meeting they will
not do it.

**Exercise 6.** Sizing should land as a range with named drivers, not a point
estimate, and the dominant assumption is almost always the target attach rate
rather than the margin per accessory line — which the sensitivity table makes
visible rather than asserted. The discipline that matters most here is
symmetry: the reseller option must get a genuine three-year ramp, hiring cost
and territory productivity assumption, not a placeholder number that loses by
construction. If your recommendation only wins because the alternatives were
sized carelessly, you have not done the analysis, you have decorated a prior.

## Key takeaways

- Frame the decision before you touch the data, and get the metric, horizon and
  comparison basis agreed in writing. Most final-meeting arguments are
  week-one omissions.
- MECE exists so that a negative result is informative. Overlapping branches
  make "we ruled it out" meaningless.
- Write the day-one answer, then attack it. Killing your own hypothesis fast is
  the highest-return activity available to you.
- Every finding must pass "so what" and "how much". Unsized findings cannot be
  prioritised, and unactionable findings are trivia regardless of how hard they
  were to produce.
- The storyboard is the deliverable before any exhibit exists. If the titles
  read in sequence do not argue the case, the deck does not either.
- Triangulate load-bearing numbers two independent ways before presenting them.
  Disagreement between methods is a gift received privately rather than
  publicly.
- Answer "where did that number come from" in three parts — source, definition,
  offer to trace — in fifteen seconds, or defer visibly. Never improvise.
- A recommendation is an action, an owner, a cost, an expected return, the
  conditions it depends on, and the criterion that would kill it.
- Lead with the number that weakens your case. The weakness you raise is an
  argument; the weakness they find is a credibility problem.
- Round to the precision your weakest assumption supports. Spurious precision
  invites an argument about decimals instead of magnitude.
