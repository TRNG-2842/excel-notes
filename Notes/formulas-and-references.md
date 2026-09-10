# Formulas, references and the difference between a value and its costume

## What it is and why it exists

A formula is anything typed into a cell that starts with `=`. Excel evaluates it instead of
storing it. The result is displayed; the formula itself lives underneath and is visible in
the formula bar.

The point of a formula is that it refers to **cells**, not to values. `=B2-C2` does not
mean "12400 minus 9750". It means "whatever is in B2, minus whatever is in C2". Change
either input and the answer follows without anybody touching it. That is the entire reason
spreadsheets exist, and it produces the first working rule: **never type a number that
could be calculated, and never calculate a number twice**.

Two ideas sit underneath every formula and cause most of the confusion.

**References move when you copy them.** Excel does not store the address you typed. It
stores the direction and distance from the formula to its input. Copy the formula somewhere
else and the direction and distance are preserved, so the address changes. This is right
almost always, and catastrophic in the one case where the input is a fixed constant.

**A number and its display are different things.** Formatting changes what you see and
never what is stored. A cell showing `12,400` and a cell showing `12400.00` and a cell
showing `1.24E+04` can hold the identical number. A cell showing a date holds an integer.

## How it behaves

### Relative, absolute and the `F4` key

| Written | Called | When copied |
|---|---|---|
| `B2` | relative | both column and row shift |
| `$B$2` | absolute | never moves |
| `B$2` | mixed | column shifts, row is pinned |
| `$B2` | mixed | row shifts, column is pinned |

While editing a formula, put the cursor inside a reference and press `F4` to cycle through
all four. On some laptops this needs `Fn+F4`, and the cell must be in edit mode, not merely
selected.

The rule of thumb: anything that is a **single constant** - a rate, a threshold, an FX
rate, a tax percentage - belongs in one labelled cell, referred to absolutely. A model with
the same rate typed into forty formulas is not a model; it is forty chances to be wrong
when the rate changes.

Worked example. `D2` holds `0.05`, labelled "handling rate". In `F2`:

```
=B2*(1+$D$2)
```

Fill down. The `B` walks down the rows; the `$D$2` stays put.

| Row | `B` | `F` with `$D$2` | `F` written as `=B2*(1+D2)` and filled down |
|---|---|---|---|
| 2 | 1000 | 1050 | 1050 |
| 3 | 2500 | 2625 | 2500 (reads `D3`, which is empty, so multiplies by 1) |
| 4 | 800 | 840 | 800 |

Change `D2` to `0.08` and the middle column becomes 1080, 2700, 864 at once. The right-hand
column does not move for rows 3 and 4 - no error, just plausible numbers that mean
nothing.

### The counting functions, and the gap between them

| Function | Counts |
|---|---|
| `COUNT(range)` | cells holding a **number** (dates count; they are numbers) |
| `COUNTA(range)` | cells holding **anything at all**, text included |
| `COUNTBLANK(range)` | genuinely empty cells |

`COUNTA` minus `COUNT` over the same column is the cheapest data-quality check that exists.
If a column of amounts has 120 entries and only 113 of them are numbers, seven of them are
blank or - worse and more common - text that looks like a number. Either way, seven rows are
not what you assumed, and you found that out in under a minute.

Note the asymmetry that catches people: `AVERAGE` **skips** blanks, so it divides by the
number of numeric cells. Subtraction **treats a blank as zero**. So `=AVERAGE(A1:A10)` and
`=SUM(A1:A10)/10` disagree the moment there is a gap, and both are behaving correctly.

### Numbers and their costumes

Excel stores a date as the number of days since 30 December 1899. `45292` is 1 January
2024. It only looks like a date when a date number format is applied. This is why a date
arriving from a formula appears as a five-digit number, and why sorting text dates puts
March before January.

Percentages work the same way. A cell holding `0.25` and formatted as a percentage displays
`25.0%`. Multiplying by 100 *and* formatting as a percentage gives `2500%`.

The practical consequence: when a number surprises you, look at the formula bar, not the
cell. The bar shows what is there. The cell shows what somebody decided you should see.

### Reading a nested formula

Excel evaluates from the inside out. Read it the same way.

```
=IF(ISBLANK(C2),"NOT RECORDED YET",D2/B2)
```

Innermost first: is `C2` empty? If yes, the answer is the text `NOT RECORDED YET`. If no,
the answer is `D2` divided by `B2`. Reading outside-in produces nothing but confusion.

## Cost and trade-offs

- **Mixed types in one column.** A column that returns a number sometimes and text at other
  times is honest and slightly awkward: anything that averages it will ignore the text
  rows. Usually the right trade, and always worth naming.
- **Floating point.** Excel stores decimals in binary and cannot represent numbers like 0.1
  exactly. Two amounts that genuinely match can subtract to `0.0000000001`. Never test
  money for equality with zero; test whether the difference is smaller than a tolerance,
  usually a cent.
- **Displayed rounding is not rounding.** Formatting a column to zero decimals does not
  change the values, so a column of numbers displayed as `100` can total `1004.7`. If you
  need the value rounded, use `ROUND`.
- **Long formulas are unmaintainable.** Four nested functions is near the practical limit
  for something a colleague has to read. Beyond that, use helper columns. They cost nothing
  and can be deleted at the end.
- **Whole-column references** like `A:A` are convenient and make Excel evaluate a million
  rows. Fine occasionally, slow everywhere.

## Recognize it on sight

- Dollar signs scattered through a formula: somebody understood absolute references. All
  dollar signs everywhere: somebody pressed `F4` until it worked.
- A five-digit number in the low forty-thousands where a date should be: a date with no
  date format.
- A column of numbers hard against the **left** of their cells: text, not numbers. `SUM`
  will return zero and not complain.
- `=SUM(B2:B41)` in a workbook whose data grows monthly: a range error waiting to happen.
- Hard-coded rates inside formulas rather than a labelled assumptions block.

## Adjacent question

*"Your model has a rate typed into forty formulas and finance changes the rate. What do you
do, and what do you do differently next time?"*

The immediate fix is Find and Replace, carefully, then a check that the total moved by the
amount you expected. The real answer is the second half: one labelled cell, referred to
absolutely, and an assumptions block at the top of the sheet where a reviewer can see every
input to the model in one place without reading a formula.

## Say it in an interview

"References are relative by default, which means Excel stores the direction and distance
to the input rather than the address. That is what makes filling down work. It is also why
anything that is a fixed constant has to be locked with dollar signs, or it walks away from
the cell you meant."

"I put every assumption in one labelled cell and point at it absolutely. If a rate changes,
that is one edit, not forty, and a reviewer can see every input without reading a formula."

"`COUNTA` minus `COUNT` on the same column tells you how many entries are not numbers. It
is the first check I run on any file I have been handed, because it costs nothing and it
finds the two failures that matter: blanks, and numbers that arrived as text."

"Formatting changes the display and never the value. Most of the arguments I have seen
about a spreadsheet being wrong were really about rounding in the display."

## Check yourself

1. `D2` contains `=B2-C2`. You copy it to `D7`. What does `D7` contain, and why?
2. When must a reference be absolute?
3. `COUNTA` returns 120 and `COUNT` returns 113 on the same column. What do you now know?
4. Why should you never test whether two amounts differ by exactly zero?
5. A cell displays `25.0%`. What is stored in it?
6. `AVERAGE` over ten cells, two of them blank. What does it divide by?

<details>
<summary>Answers</summary>

1. `=B7-C7`. References are relative: Excel preserves the offset, so moving the formula
   down five rows moves both inputs down five rows.
2. When the input is a single fixed cell that must not move as the formula is copied:
   a rate, a threshold, an exchange rate, a lookup range.
3. Seven of the cells are not numbers. They are blank, or they are text that looks like a
   number. Either way seven rows are not what you assumed.
4. Binary floating point means values that genuinely agree can subtract to a tiny non-zero
   remainder. Compare against a tolerance, typically 0.01.
5. `0.25`. The percent sign is a number format, not part of the value.
6. Eight. `AVERAGE` ignores blanks. Subtraction and `SUM`, by contrast, treat a blank as
   zero, which is why the two can disagree.

</details>

## Resources

- Overview of formulas in Excel: https://support.microsoft.com/en-us/office/overview-of-formulas-in-excel-ecfdc708-9162-49e8-b993-c311f47ca173
- Switch between relative, absolute and mixed references: https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9
- Format numbers as dates or times: https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309
