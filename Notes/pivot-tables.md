# PivotTables

## What it is and why it exists

A PivotTable summarises a table by dragging column names into boxes. No formulas are
written and none can be got wrong.

The problem it solves is not "adding up". It is that the question keeps changing. Spend per
client. Actually, per client per month. Actually, only cash, and only high-risk clients,
and as a chart. Written as formulas, each of those is a different sheet, twenty minutes,
and a fresh opportunity for a range error. As a pivot, each is ten seconds of dragging, and
the underlying data is never touched.

The reason it can do that is a property of the **source**, not of the pivot. Data in long
format - one row per event, one column per attribute, nothing summarised - can be
rearranged into any summary you can describe. Data that has already been arranged into a
report cannot be rearranged into a different one.

This produces the most important rule about pivots, and it is a rule about your source
data: **keep the source long, ugly and one-row-per-thing**. The moment you type totals into
it, merge cells, or split it into a sheet per month, you have destroyed the raw material.
It is a mistake made from good intentions, by people tidying their data into the shape of
whichever report they happen to need first.

## How it behaves

### Building one

Click inside the source table, Insert > PivotTable > From Table/Range, and put it on a new
worksheet. Always a new worksheet: a pivot grows and shrinks as you change it, and next to
your data it will eventually collide with it.

The Fields pane has four boxes.

| Box | Holds | Appears |
|---|---|---|
| **Rows** | what you group by | down the left |
| **Columns** | a second grouping | across the top |
| **Values** | what gets aggregated | in the middle |
| **Filters** | a control over the whole thing | above the pivot |

Drop a text field into Values and Excel chooses `Count`. Drop a number and it chooses
`Sum` - unless the column contains even one blank or text cell, in which case it chooses
`Count` and your "total" is a row count. It is guessing, and its guess is right often
enough to be dangerous. When a Sum looks far too small, check the aggregation before the
data.

### The three things worth knowing after the basics

**Summarize Values By.** The same field can go into Values more than once, answering a
different question each time. Sum tells you size. Count tells you activity. Average and Max
tell you *behaviour*, and behaviour is where the interesting things hide. A counterparty
with a large total and a large average is a large counterparty. One with a large total and
a small maximum is doing something else entirely, and no amount of staring at totals will
reveal it.

**Show Values As.** The second tab in the same dialog, which almost nobody clicks. It turns
the displayed number into a percentage of the total, a running total, a rank, or a
difference from the previous period. Every one of those is a formula you did not have to
write.

**Grouping.** Right-click a date in the pivot and choose Group to collapse dates into days,
months, quarters or years. Numbers can be grouped into bands the same way. This only works
on **real** dates and numbers; text dates group as one label each, in alphabetical order,
which puts March before January.

### Behaviour, not size: reading a threshold

Most summaries are read for size. The biggest client, the biggest month, the biggest total.
Size is usually the least interesting thing in the table, because everybody already knows
it.

Put the same numeric field into Values three times, as **Sum**, **Count** and **Max**, and
a different question becomes answerable: not how much, but *how*. Sum tells you scale.
Count tells you frequency. Max tells you the shape of the largest single event.

The combination matters where a rule has a threshold in it. Regulated reporting is full of
these: a transaction at or above some amount has to be reported, and below it does not.
Anyone who wants to move value without generating that report has one obvious option,
which is to split it into several smaller pieces, each of them individually unremarkable
and entirely legal-looking.

That pattern is invisible one record at a time. It only appears when you aggregate by
counterparty and look at the Max column: an entity with many events and a maximum that
never once crosses the line is behaving differently from every other row in the table.

Three disciplines go with it.

**Aggregate, then read the outlier, then drill in.** Double-click the cell to get the
source rows, and read the actual records before you say anything.

**The finding is a pattern, not a conclusion.** There is often a mundane explanation, and
deciding is not your job. Escalate it accurately to whoever owns that decision.

**Write it as evidence.** "Seven payments totalling 64,300 between the 4th and the 21st, none
exceeding 9,850, no payment from this counterparty has crossed the threshold" is what a
professional writes. Naming a crime in an email is how you end up in a deposition.

### PivotCharts

`PivotTable Analyze > PivotChart` gives you a chart wired to the pivot rather than a picture
of it. Filter the pivot and the chart follows; filter the chart and the pivot follows. They
are two views of one object.

Four practical rules.

- **Bar for categories, column for time.** Long category names read left to right, so they
  belong on the vertical axis. Time reads left to right, so it belongs on the horizontal.
- **No pie chart with more than about four slices.** Nobody can compare angles.
- **Turn off the field buttons before you screenshot it.** They are useful while you build
  and pure noise in a deck: `PivotChart Analyze > Field Buttons > Hide All`.
- **The chart inherits the pivot's filters,** including the one you forgot about. A chart
  captioned "total spend" that silently excludes most of the data looks exactly like
  one that does not.

### Refresh, which is the thing that catches everybody

A PivotTable holds its own cached copy of the source, taken when you built it or last
refreshed it. That cache is what makes it fast, and it is what makes it lie.

Change the source and the pivot does not move. There is no warning, no colour, no error. It
simply keeps showing you last week.

- `Alt+F5` refreshes the selected pivot.
- `Ctrl+Alt+F5` refreshes every pivot in the workbook.

Make it a habit: refresh before you screenshot, refresh all before you send.

If the source is an Excel Table, added rows are inside the source range automatically and a
refresh picks them up. If it is a plain range, you have to change the range by hand, and
almost nobody remembers to.

### Two traps in reading one

**A filtered pivot totals what is visible.** Filter to the top three and the Grand Total is
the total of those three, not of everything. Screenshot that into a deck captioned "total
spend" and you are wrong by whatever you filtered out, with a number that looks entirely
plausible. When you filter, say so in the title of whatever you paste it into.

**A slicer is a filter people can see.** It is a panel of buttons wired to one or more
pivots. The speed is nice; the real reason to prefer it over a dropdown is that a dropdown
filter is invisible from three feet away. Somebody opens your workbook, reads a number, and
has no idea it excludes most of the data. A slicer sits there in bold saying CASH.

### Drill-through

Double-click any number inside a pivot and Excel creates a new sheet containing the source
rows behind it. It is the fastest route from a pattern to the evidence, and the loop it
completes - aggregate, find the thing that behaves differently, drill in, read the actual
records - is what most investigative analysis consists of.

Delete the sheets it creates as you go, or the workbook fills with them.

## Cost and trade-offs

- **Stale by default.** The single biggest source of wrong numbers in circulated workbooks.
- **Calculated fields and calculated items** exist inside pivots and are a reliable way to
  confuse yourself, particularly around how they interact with filters. Do the calculation
  as a column in the source instead, where it is visible and testable.
- **`GETPIVOTDATA`.** Point a formula at a pivot cell and Excel writes this instead of a
  cell reference. It is robust and unreadable, and it breaks when the pivot changes shape.
  Turn it off: PivotTable Analyze > Options > Generate GetPivotData.
- **One source per pivot.** Combining two tables needs the Data Model, a relationship
  between them, or a join done beforehand with a lookup.
- **Formatting inside the pivot, not on the cells.** Format cells directly and the
  formatting stays attached to positions rather than to fields, so it lands on the wrong
  numbers the moment the pivot changes shape. Right-click > Number Format.
- **Blank category values** silently drop rows out of a grouping. Fill them with a visible
  label first.
- **Cache size.** Each pivot stores its own copy of the source, so several pivots on a large
  table make a large file. Building them from the same source lets Excel share one cache.

## Recognize it on sight

- A summary block whose cells cannot be edited, with `Row Labels` or a field name in the
  top-left corner.
- A **PivotTable Analyze** and **Design** pair of ribbon tabs appearing when you click it.
- Floating panels of buttons: slicers. If any button is highlighted, the data is filtered
  and the totals are partial.
- `GETPIVOTDATA` in a formula elsewhere on the sheet.
- A sheet named `Sheet12` containing a few dozen raw rows: somebody drilled through and
  forgot to delete it.

## Adjacent question

*"You have three hundred thousand rows and three related tables. Still PivotTables?"*

Not in this form. At that scale you load the tables into the Data Model, define
relationships between them, and pivot across all three, which is Power Pivot. Beyond that
the answer is a database and SQL, with Excel as the presentation layer or not involved at
all. The thinking transfers exactly: group by, aggregate, filter. A PivotTable is a `GROUP
BY` with a mouse.

## Say it in an interview

"A PivotTable answers a summary question by dragging fields rather than writing formulas,
which means the question can change without the work being redone. But it only works if the
source is long format, one row per event with nothing pre-summarised, so most of the
discipline is in how the data is stored rather than in the pivot itself."

"The thing to know about pivots is that they cache. Change the source and the pivot shows
you the old numbers with no warning at all. I refresh all before sending anything."

"A filtered pivot totals only what is visible, so I put the filter in the title of anything
I paste it into. And I prefer slicers to dropdown filters, because a slicer is visible to
whoever opens the file afterwards."

"Putting the same field into Values more than once, as a sum and a count and a maximum, is
how you find behaviour rather than size. The interesting number is rarely the biggest one;
it is the one behaving differently from its neighbours."

## Check yourself

1. You change a value in the source and the pivot does not move. Is it broken?
2. Why must source data be one row per event with no totals inside it?
3. A pivot filtered to the top three clients shows a Grand Total. What is that total?
4. Why does grouping by month fail on some date columns?
5. What does double-clicking a value inside a pivot do?
6. Give one reason to prefer a slicer over the built-in filter dropdown.

<details>
<summary>Answers</summary>

1. No. It holds a cached copy of the source. `Alt+F5` refreshes one, `Ctrl+Alt+F5`
   refreshes all.
2. Because the pivot summarises. Totals already inside the source are treated as events and
   counted again, so everything double-counts.
3. The total of the three visible clients only. Filtered pivots total what is visible, with
   no warning.
4. Because those dates are text, not real dates. Grouping only works on genuine date and
   number types.
5. Drill-through: it creates a new sheet containing the source rows behind that number.
6. It is visible. A dropdown filter cannot be seen from a distance, so a reader has no idea
   the totals are partial.

</details>

## Resources

- Create a PivotTable to analyze worksheet data: https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576
- Group or ungroup data in a PivotTable: https://support.microsoft.com/en-us/office/group-or-ungroup-data-in-a-pivottable-c9d1ddd0-6580-47d1-82bc-c84a5a340725
- Use slicers to filter data: https://support.microsoft.com/en-us/office/use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d
- Show different calculations in PivotTable value fields: https://support.microsoft.com/en-us/office/show-different-calculations-in-pivottable-value-fields-014d2777-baaf-480b-a32b-98431f48bfec
