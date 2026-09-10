# Excel Tables and the shape of good data

## What it is and why it exists

An Excel Table is a rectangle of cells that you have told Excel is one thing. It has a
name, a header row, and a boundary that Excel maintains for you.

Without one, a spreadsheet is a grid of independent cells that happen to sit near each
other. Excel does not know that row 250 is the last row of anything. Add row 251 and nothing
notices: not your totals, not your charts, not the summary on the next sheet. Every one of
them is pointing at a fixed range like `A1:H250`, and every one of them is now silently
wrong.

A Table fixes that by making the boundary a fact Excel tracks rather than a number you
typed. It exists because most spreadsheet errors are range errors, and a
range error never announces itself.

Underneath the Table is a second idea that matters more: **the shape the data is in**.
Data in *long* format has one row per event and one column per attribute. Nothing is
summarised, nothing is merged, nothing is arranged for a human to read. It looks
deliberately boring. Boring data can be turned into any report on demand. Data that has
already been arranged into a report cannot be turned back.

## How it behaves

Select any cell inside a block of data and press `Ctrl+T`. Excel guesses the boundary and
whether there is a header row, and shows you both. Accept, and you get:

- **A name.** Set it in the Table Design ribbon tab. `Table1` tells nobody anything; use
  something like `tblPayments`. Names cannot contain spaces or start with a digit, and no
  two tables in one workbook may share a name.
- **Automatic growth.** Type in the first cell to the right of the last column, or the
  first row under the last row, and the Table absorbs it. Formatting, filters and
  everything referring to the Table follow.
- **Calculated columns.** Put a formula in one cell of an empty column inside a Table and
  Excel fills the entire column immediately. Edit the formula in any row and it offers to
  update all of them.
- **A sticky header.** Scroll down inside a Table and the column letters in the grey bar
  are replaced by the column names.
- **Structured references.** `=SUM(tblPayments[amount])` instead of `=SUM(F2:F250)`. It
  means "the amount column, however long it is". Add rows and the sum includes them. Inside
  the Table, `[@amount]` means "the amount on this row".
- **Filters,** on every column, without asking.

A worked example. A table named `tblPayments` with columns `amount` and `fee`:

| Formula | Where | Result |
|---|---|---|
| `=SUM(tblPayments[amount])` | anywhere in the workbook | total of that column, at whatever length |
| `=COUNTA(tblPayments[payment_id])` | anywhere | number of rows |
| `=[@amount]-[@fee]` | in a column inside the Table | that row's amount minus that row's fee, filled down the whole column automatically |

### Moving around, sorting and filtering without breaking anything

**Navigation.** `Ctrl+Down` jumps to the last filled cell in the column; `Ctrl+Up` comes
back. `Ctrl+Shift+Down` selects from here to there. `Ctrl+Home` returns to `A1`,
`Ctrl+End` to the last used cell on the sheet. On a large file these replace scrolling
entirely, and `Ctrl+End` landing far below the data tells you there are stray cells
somewhere.

**Freeze Panes.** View > Freeze Panes > Freeze Top Row keeps the headers on screen while
you scroll. Select `B2` and choose Freeze Panes instead to pin both the top row and the
first column. A Table gives you a sticky header without this; a plain range does not.

**Sorting.** Click the filter arrow on any header and choose Sort A to Z, or Data > Sort
for several levels at once. Inside a Table every row moves as one unit. In a plain range,
selecting one column and sorting it sorts *only that column*: Excel warns with "Expand the
selection?" and if you decline, every amount now sits beside the wrong client, with no way
back except undo. This is the single most destructive thing a beginner does to a dataset.

**Filtering.** The filter arrow hides rows that do not match. Two things to know. The
status bar reads "n of m records found", which is the count you want. And `SUM` over a
filtered column still adds the hidden rows: `=SUM(tblPayments[amount])` returns the total
of everything whether or not a filter is on. Use `=SUBTOTAL(109, tblPayments[amount])`
for visible rows only, or turn on the Table's Total Row (Table Design > Total Row), which
does that for you. Clear filters with Data > Clear before you hand the file on; a saved
filter is invisible to the next reader.

**Number formats.** Home > Number, or `Ctrl+Shift+1` for two decimals and a thousands
separator. Formatting changes what is shown and never what is stored, so `1,250.00` and
`1250` can be the same cell.

## Cost and trade-offs

- **Structured references get long.** `tblPayments[amount]` is more to type and read than
  `F:F`. Autocomplete helps: type `tblP` and press Tab.
- **Auto-expansion can surprise you.** Type a note in the cell beside a Table and it
  becomes a Table column called something like `Column14`. Leave a genuinely blank spacer
  column if you want a boundary.
- **Two Tables cannot touch.** A Table will not expand into another Table's cells, and it
  will refuse the edit rather than doing something clever.
- **Names are workbook-wide.** If you copy a sheet, Excel silently renames the copied
  Table to `tblPayments2`, and any formula you copy alongside it may still point at the
  original.
- **Merged cells and Tables are incompatible,** which is a feature. So are Tables and
  multi-row headers.
- **On very large data**, a Table adds a little overhead. It is not the reason your
  workbook is slow; volatile functions and conditional formatting usually are.

## Recognize it on sight

- Banded rows in a colour scheme you did not choose, with filter arrows on every header.
- A **Table Design** tab appearing in the ribbon when the cursor is inside the data.
- Formulas containing square brackets: `tblSomething[column]` or `[@column]`.
- In the Name Box dropdown, or under Formulas > Name Manager, the table names are listed.

The opposite tells you something too. A sheet with merged title cells, a blank row under
the header, subtotals scattered mid-data and a different layout per month is a report
somebody typed, not a dataset. Before you can analyse it you have to reconstruct the long
form, and that is usually the bulk of the work.

## Adjacent question

*"You have twelve monthly files with the same columns. How do you get one Table?"*

The honest answer is that you do not do it by copying and pasting twelve times, because
you will have to do it again next month. This is what Power Query exists for: point it at
a folder, tell it the shape, and it appends every file into one Table that refreshes with
a button. Knowing that the answer is "a repeatable import, not manual assembly" matters
more than knowing the clicks.

## Say it in an interview

"An Excel Table gives the data a name and a boundary that Excel maintains, so anything
built on it follows it when it grows. Without one, every formula points at a fixed range
and quietly stops covering the data the first time a row is added."

"I keep source data in long format, one row per event, no totals and no merged cells. Then
anything that has to present it does so downstream. The moment you tidy the source into
the shape of the report, you have made every other report impossible."

"Structured references also make formulas readable. `SUM(tblPayments[amount])` says what
it does. `SUM(F2:F250)` says where it is, which is the less useful fact."

## Check yourself

1. You add a row to the bottom of a plain range that a `SUM` above it covers. What happens
   to the total?
2. What is the keyboard shortcut to create a Table, and what two things does the dialog
   ask you to confirm?
3. Why can two Tables in the same workbook not share a name?
4. What does `[@amount]` mean, and where is it valid?
5. Give two signs that a spreadsheet you have been sent is a report rather than a dataset.
6. A filter is on and `=SUM` over the column has not changed. Why, and what do you use
   instead?

<details>
<summary>Answers</summary>

1. Nothing. The total is anchored to a fixed range and now excludes the new row. Nothing
   errors and the number just becomes wrong.
2. `Ctrl+T`. It asks you to confirm the range it detected and whether the first row is a
   header row.
3. Table names act as workbook-level names, the same namespace as named ranges, so they
   must be unique across the file rather than per sheet.
4. The value in the `amount` column on the current row. It is only valid inside a formula
   that sits within that Table.
5. Any two of: merged cells; a title above the header; blank rows inside the data;
   subtotal rows mixed in with detail rows; one sheet per month with an identical layout;
   numbers stored with their currency symbol as text.
6. `SUM` ignores filters and adds hidden rows. `SUBTOTAL(109, range)` or the Table's
   Total Row sums only the visible ones.

</details>

## Resources

- Overview of Excel tables: https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
- Using structured references with Excel tables: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Create or delete an Excel table: https://support.microsoft.com/en-us/office/create-or-delete-an-excel-table-e81aa349-b006-4f8a-9806-5af9df0ac664
- Rename an Excel table: https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114
