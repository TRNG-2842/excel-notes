# Catch-up: sessions 01 to 04, on your own

You are joining this course at session 05, PivotTables. Sessions 01 to 04 were delivered
live to an earlier group, and 05 assumes you sat through them. This document lets you do
those four sessions by yourself, in the same workbook, with the same numbers to check
against, in about four hours. Recordings of the live sessions are being arranged; this is
what you can do while you wait for them.

Nothing here is marked and nothing is submitted. Every step tells you what you should see.
If you see it, move on. If you do not, the "If it goes wrong" table under each part covers
the common causes.

## How the workbook is built, and why you cannot fall behind

The workbook is `ops-primer.xlsx`. It has one sheet per session, and **every sheet
already contains the finished version of the sheet before it.** `02_Formulas` is
`01_Foundations` with session 01 done. `03_Checks` is `02_Formulas` with session 02 done.
And so on. So:

- if a part goes badly, click the next tab and you are level again;
- you never have to redo an earlier part to make a later one work;
- session 05 starts on `05_Pivots`, which is already the clean result of session 04.

One exception: `04_Cleaning` starts a new dataset, payments rather than engagements. The
engagement tracker's story ends on `03_Checks`.

Do the parts in order anyway. Each one uses ideas from the one before, and 05 uses all of
them.

## Time

| Part | Sheet | Alone, roughly |
| --- | --- | --- |
| 0. Set up | any | 5 min |
| 1. Foundations: read, sort, filter, make a Table | `01_Foundations` | 45 min |
| 2. Formulas: calculated columns, dollar signs, a summary box | `02_Formulas` | 40 min |
| 3. Checks: the five errors, `XLOOKUP`, validation | `03_Checks` | 50 min |
| 4. Cleaning: 312 dirty rows to 300 clean | `04_Cleaning` | 70 min |

Part 4 is the long one and the one session 05 leans on most. If you only have two hours
before 05, do Part 0, skim Parts 1 and 2 reading the "What you should know afterwards"
boxes, do Part 3 Steps 3.3, 3.4 and 3.6 (the `XLOOKUP`, its named fallback, and the
`NO COST YET` fix), and do all of Part 4. Session 05 refers to all three.

## Four things before you start

1. **Desktop Excel, Microsoft 365.** Every formula here works in the current Microsoft
   365 desktop app. `XLOOKUP` and `UNIQUE` do not exist in Excel 2019 or earlier.
2. **Save your own copy first.** File > Save As, into a local folder, under a name that
   is yours. Never work in the copy that was sent round. `Ctrl+S` often.
3. **Copying a formula down a column.** In an Excel Table it happens by itself. In a plain
   range, click the formula cell, `Ctrl+C`, click the Name Box (the small box left of the
   formula bar), type the range *below* it such as `P3:P301`, Enter, `Ctrl+V`. This
   document says "copy down to row N" to mean exactly that.
4. **Three keys worth knowing today.** `F2` edits the active cell in place. `F4`, while
   editing, cycles the dollar signs on a reference. `Ctrl+1` opens Format Cells. On some
   laptops the function keys are media keys and you hold `Fn` as well.

---

# Part 0 - Set up (5 minutes)

**Step 0.1 - Open it properly.** Open `ops-primer.xlsx`. If a yellow bar reads
**Protected View**, click **Enable Editing**; nothing you type is kept until you do.
File > Save As, pick a local folder, name it `ops-primer-mine.xlsx`, Save.

**Step 0.2 - Check the title bar.** It reads `ops-primer-mine`. The tab strip at the
bottom shows `00_Agenda`, `01_Foundations`, `02_Formulas`, `03_Checks`, `04_Cleaning`,
`05_Pivots`, `06_Reconcile`, `07_Lab_Tasks`, `07_Lab_Data`.

**Step 0.3 - Two quick checks.** Click any cell with data, press `F2`: the cell opens for
editing and the cursor sits at the end of its contents. `Escape`. Press `Ctrl+1`: Format
Cells opens. `Escape`. If `F2` did nothing, try `Fn+F2`, and use `Fn` with `F4` later too.

---

# Part 1 - Foundations (`01_Foundations`, about 45 minutes)

**What is on the sheet.** Columns A to J, rows 1 to 41: 40 consulting engagements as they
came out of a system. Columns too narrow, dates showing as five-digit numbers, money with
no separators. Column P holds a grey note. Columns K, L and M are empty and stay empty
until the last step.

**What this part is for.** Nothing in this part changes a single value. It is all about
making a file readable, because every mistake you will ever catch, you catch by looking,
and nobody checks what they cannot read. It ends by turning the block into an **Excel
Table**, which is the thing every later session stands on.

### Step 1.1 - Move without the mouse (5 min)

Click any cell in the middle of the data, then try each of these:

```
Ctrl+Home          jump to A1
Ctrl+End           jump to the last used cell
Ctrl+Down Arrow    fall to the bottom of the current block of data
Ctrl+Up Arrow      back to the top
Ctrl+Right Arrow   run to the last column with data
```

`Ctrl+Right` from inside the block stops at column J, the edge of the data. `Ctrl+End`
lands on `P41`, not `J41`, because the grey note in column P is part of the used range.

Click the **Name Box** (left of the formula bar), type `H41`, Enter. You are there. Any
address, any sheet, Enter. You will use it constantly.

Click `A1`, then `Ctrl+Shift+Down`, then `Ctrl+Shift+Right`. `A1:J41` is selected and the
**status bar** at the bottom of the window shows a Count. Shift turns moving into
selecting.

### Step 1.2 - Make it readable (4 min)

Click the small triangle above row 1 and left of column A to select everything
(shortcut: `Ctrl+A` twice). Home tab > Cells group > **Format** > **AutoFit Column
Width** (keyboard: `Alt`, then `H`, `O`, `I`). Every column snaps to its widest entry.

Column P is now enormous. Right-click column P's header letter > **Column Width**, type
`92`, OK.

Select `A1:J1`, press `Ctrl+B`. Header row goes bold.

If a column shows `########`, it is not an error. It is a number too wide for its column.
Widen the column.

### Step 1.3 - Number formats: the same value, dressed differently (10 min)

Click cell `F2` (the cell, not the key). It shows `46061`, and so does the formula bar.
That is a date. Excel stores a date as the number of days since 1 January 1900; what you
see is a costume, and right now the costume is missing.

Select `F2:G41`. Home tab > Number group > the dropdown that reads `General` > **Short
Date**. Cell `F2` shows `2/8/2026`, and the formula bar now reads it as a date too.

Same selection, press `Ctrl+1`. Choose **Date**, and in the Type list `03/14/2012` (on a
US-region machine; elsewhere pick the two-digit-day equivalent, or Custom `mm/dd/yyyy`).
OK. Cell `F2` shows `02/08/2026`. `Ctrl+1` is Format Cells, the one shortcut worth memorising
today. Every appearance question in Excel lives behind it.

Select `H2:I41`. Home > Number > the **Comma Style** button (`,`). `H2` shows
`478,500.00`. Click **Decrease Decimal** (the `.00` to `.0` button) twice. `H2` shows
`478,500`, `I2` shows `358,900`.

Still on `H2:I41`, press `Ctrl+1` and in the left-hand list click **Currency**, then
**Accounting**, watching the preview. Currency puts the symbol against the digits.
Accounting pushes it to the left edge, lines up the decimal points, and shows zero as a
dash, which is why finance uses it. Click **Cancel**; the columns stay at `478,500`.

**The point of this step:** `H2` has been `478500` the whole time. Formatting changes the
display and never the value. The day you forget that, you will see a column of cells
showing `100`, sum them, get `100.4`, and think Excel is broken.

`I5` is blank. Leave it blank. It is blank in the source system and Part 2 uses it.

### Step 1.4 - Freeze the header (2 min)

Click `A2` first. View tab > Window group > **Freeze Panes** > **Freeze Panes**. A thin
line appears under row 1. `Ctrl+Down`; row 1 stays. `Ctrl+Home` back.

Excel freezes everything above and left of the selected cell. `A2` freezes row 1 only. If
you had `A1` selected, nothing visible happens: Freeze Panes > Unfreeze Panes, click `A2`,
try again.

### Step 1.5 - Sort: this one changes your file (5 min)

Click one cell inside the data, for example `D5`. Data tab > Sort & Filter group >
**Sort**. "My data has headers" is ticked. Sort by `service_line`, A to Z. **Add Level**.
Then by `budget`, Largest to Smallest. OK.

**See:** all eight Advisory rows first, biggest budget at the top. Row 2 is `ENG-1028`,
budget 459,000. Row 3 is `ENG-1003`, 448,500.

`Ctrl+Z` once. `ENG-1001` is back at row 2.

**The rule:** always click **one** cell and let Excel find the edges. If you select a
single column and sort, Excel offers "Continue with the current selection". Take that and
the one column is reordered while the rest of every row stays put. Nothing errors. Every
row is now wrong. If you sort, save and close, the old order is gone.

### Step 1.6 - Filter: this one does not change your file (5 min)

Click any cell in the data. Data > Sort & Filter > **Filter** (shortcut `Ctrl+Shift+L`).
Dropdown arrows appear on row 1.

Arrow on `status` (column J) > untick Select All > tick `Active` > OK. Row numbers on the
left turn blue and skip, a funnel icon sits on the `status` header, and the status bar at
bottom left reads `21 of 40 records found`. Those are the tells that a filter is on.

Arrow on `service_line` (column D) > tick `Audit` only > OK. `4 of 40 records found`.
Filters stack.

Data > Sort & Filter > **Clear**. All 40 rows return.

A file that "has rows missing" almost always has a filter left on. First move on any file
somebody hands you: `Ctrl+Shift+L` twice, off and on. And clear your filters before you
copy anything: a `Ctrl+C` on a whole column takes the hidden rows too.

### Step 1.7 - Turn it into an Excel Table (10 min)

Right now this is a block of cells that looks like a table. Excel does not know it is
one. An Excel Table is a named object: it grows by itself when you type next to it, it
has a name you can refer to in formulas, and everything built on it follows it.

Click one cell inside the data, for example `C4`. Press `Ctrl+T`. The Create Table dialog
shows `=$A$1:$J$41` with "My table has headers" ticked. Do not touch either. OK.

**See:** blue banded rows, filter arrows on every header, and a new **Table Design** tab
on the ribbon.

Table Design > Properties group > **Table Name** box reads `Table1`. Select that text,
type `tblEng01`, Enter.

Table names cannot contain spaces or start with a number. The sheet number is in the name
because Excel refuses two tables with one name in one workbook, and this tracker appears
on three sheets: `tblEng01`, `tblEng02`, `tblEng03`.

`Ctrl+Down` to row 41, then scroll back up slowly. If the window is short enough that
row 1 would scroll off, the column headers replace the column letters in the grey bar at
the top while you are inside the table. A Table does that for you without freeze panes.
On a tall window all 41 rows fit and there is nothing to see; move on.

Table Design > Table Style Options > tick **Total Row**. Row 42 appears. Click `H42`,
use its dropdown, choose `Sum`. **See:** `11,645,000`. Remember that number. Untick Total
Row.

If the Create Table dialog offered a smaller range, there is a blank row or column inside
the data: Escape, select `A1:J41` by hand or through the Name Box, `Ctrl+T`.

### Step 1.8 - Three empty columns (2 min)

Click `K1`, type `variance`, Enter. **See:** column K turns blue and joins the table on
its own. `L1`: `margin_pct`. `M1`: `budget_with_uplift`. The table is now `A1:M41` with
three empty columns.

Type the headers carefully with no trailing spaces; a trailing space in a header haunts
every formula that names the column.

### If it goes wrong

| You see | Cause | Fix |
| --- | --- | --- |
| `Ctrl+Down` jumps to row 1048576 | you started on an empty cell | click a cell with data first |
| Dates still show numbers after Short Date | cells are text, not dates | not a problem today; Part 4 fixes text dates |
| Table Design tab missing | cursor not inside the table | click a cell in it |
| Table name refused | space in the name, or starts with a digit | `tblEng01` exactly |
| Excel invented `Column1`, `Column2` | "My table has headers" was unticked | `Ctrl+Z`, `Ctrl+T` again, tick it |

### What you should know afterwards

- Number formatting changes what you see, never what is stored.
- Sort rewrites rows. Filter only hides rows.
- Click one cell before sorting. Never select a column and sort it.
- An Excel Table has a name, grows by itself, keeps its filters, and is what formulas and
  PivotTables point at.
- `########` is a narrow column, not an error.

<details>
<summary>Check yourself (answers)</summary>

1. *I change H2 from showing 478500 to 478,500. What is stored in H2?* 478500. Unchanged.
2. *Rows seem to be missing from a file. First check?* A filter: blue, skipping row
   numbers, a funnel on a header. Data > Clear.
3. *Name two things an Excel Table gives you.* Any of: grows automatically, has a name,
   comes with filters, things built on it follow it.
4. *A cell shows `########`. Error?* No. Widen the column.

</details>

---

# Part 2 - Formulas (`02_Formulas`, about 40 minutes)

**What is on the sheet.** Part 1 finished: `A1:M41` is a Table named **`tblEng02`**, with
`K`, `L`, `M` headed and empty. Column S holds a small `PARAMETERS` box with `Fee uplift`
in `S2` and `7.5%` in `T2`, and a `SUMMARY` block with seven labels in `S5:S11` and
empty cells beside them in `T5:T11`.

**What this part is for.** Three ideas that carry the rest of the course. A formula starts
with `=`. A formula refers to cells, not values, so it updates itself. And when you copy a
formula, its references move with it, which is usually right and occasionally a disaster;
the dollar sign is how you say which.

### Step 2.1 - Your first formula, and the Table fills itself (8 min)

Click `K2`. Type exactly, then Enter:

```
=H2-I2
```

**See:** `K2` shows `119,600`, and `K2` to `K41` all fill at the same instant. You did
not copy anything. In an empty column of a Table, one formula becomes the whole column.
That is a **calculated column**. A small lightning-bolt icon may appear beside the cell;
it offers to undo the fill. Ignore it.

Click `K3`. The formula bar reads `=H3-I3`, not `=H2-I2`. Excel stored the instruction
("three columns left minus two columns left"), not the address. That is a **relative
reference**, and it is the default.

Click `K5`. **See:** `59,500`, which is that row's entire budget, because `I5` is blank.
Excel treats a blank as zero in arithmetic and has cheerfully reported this engagement
as 59,500 under budget. It is not under budget; nobody has entered the cost. No error, no
warning. Remember row 5.

### Step 2.2 - Percentages (5 min)

Click `L2`. Type, Enter:

```
=K2/H2
```

**See:** `L2` shows `25.0%` and the column fills. No multiplying by 100; the column is
already formatted as a percentage. Excel stores 0.25 and prints 25.0%.

Click `L5`. **See:** `100.0%`. Rows 11, 20, 29, 36 and 41 say the same. Six rows out of
forty are reporting a 100% margin because their cost is missing. Arithmetically perfect,
completely false. This is the shape of almost every Excel disaster: a correct formula fed
a fact that is not there.

### Step 2.3 - Absolute references: the dollar sign (12 min)

`T2` holds 7.5%, a rate-card assumption. Build a column that applies it, so that when
finance changes 7.5 to 8, you change one cell and forty rows follow.

Click `M2`. Type, Enter:

```
=H2*(1+T2)
```

**See:** `M2` shows `514,388`. Every row below it is **wrong**: `M3` shows `298,500`,
which is exactly `H3`. Click `M3`: the formula bar reads `=H3*(1+T3)`. The reference to
`T2` slid down with the formula, `T3` is empty, empty is zero, budget times one is the
budget. No error. A column of plausible numbers that mean nothing.

Click `M2`, press the `F2` key to edit. Put the cursor inside the `T2` and press the
**`F4`** key once.
The reference becomes `$T$2`:

```
=H2*(1+$T$2)
```

Enter. The column refills. **See:** `M3` shows `320,888`. Nothing is zero.

`F4` cycles: `T2`, then `$T$2` locked both ways, then `T$2`, then `$T2`, then back.
Today you want the second one. `$` in front of the column letter locks the column; in
front of the row number locks the row. Both, and the reference does not move when
copied. Dollar signs mean *do not move*. If `F4` does nothing, hold `Fn` with it, or
type the two dollar signs by hand. `F4` only works while the cell is in edit mode; on a
selected cell it repeats your last action instead.

Click `T2`, type `0.10`, Enter. Every cell in M jumps; `M2` becomes `526,350`. `Ctrl+Z`.
Back to `514,388`. One cell, forty rows.

`M2` displays `514,388` but stores `514,387.50`; the column has no decimals. Part 1's
lesson.

### Step 2.4 - The summary box, referring to columns by name (12 min)

Click `T5`. Type exactly, Enter:

```
=SUM(tblEng02[budget])
```

**See:** `11,645,000`. The same number the Total Row gave in Part 1, by a different route.

`tblEng02[budget]` is a **structured reference**: the table's name, then the column's
name in brackets. It means "the budget column, however long the table is". `H2:H41`
would silently ignore rows added tomorrow. As you type `tblE`, Excel offers to finish the
name; press `Tab` to accept, then `[` lists the columns.

The mouse route: type `=SUM(`, hover the **top edge** of header cell `H1` until the
pointer is a small black down arrow, click once, type `)`. If the formula reads
`tblEng02[[#Headers],[budget]]` you clicked inside the header instead; that sums the word
"budget" and shows `0` with no error. Escape and type it.

Now the rest, one per cell:

| Cell | Type exactly | See |
| --- | --- | --- |
| T6 | `=SUM(tblEng02[actual_cost])` | `8,847,400` |
| T7 | `=COUNTA(tblEng02[engagement_id])` | `40` |
| T8 | `=COUNT(tblEng02[actual_cost])` | `34` |
| T9 | `=AVERAGE(tblEng02[budget])` | `291,125` |
| T10 | `=SUM(tblEng02[variance])` | `2,797,600` |
| T11 | `=T10/T5` | `24.0%` |

**Stop at 40 and 34.** `COUNTA` counts anything not empty. `COUNT` counts numbers only.
The six-row gap is the six blank costs, and now the workbook knows it, not just you.
**COUNTA minus COUNT is your blank detector.** Session 05 refers back to this.

`AVERAGE` ignores blanks rather than treating them as zero, so an average of
`actual_cost` would be over 34 rows, not 40. Different functions, different rules.

### Step 2.5 - Prove the chain, then undo (3 min)

Click `I5`, type `59500`, Enter. **See:** `K5` drops to `0`, `L5` to `0.0%`, `T6` rises
to `8,906,900`, `T8` goes to `35`, `T10` to `2,738,100`, `T11` to `23.5%`. Six cells you
did not touch, all moved.
`Ctrl+Z` once. `T8` back to `34`. If it is not, click `I5` and press Delete.

### If it goes wrong

| You see | Cause | Fix |
| --- | --- | --- |
| The cell shows the text `H2-I2` left-aligned | no `=` | retype starting with `=` |
| Only `K2` filled, rest of column empty | calculated-column fill is switched off (File > Options > Proofing > AutoCorrect Options > AutoFormat As You Type > Fill formulas in tables) | `K2`, `Ctrl+C`, Name Box `K3:K41`, `Ctrl+V` |
| `#NAME?` in the summary box | table name misspelled, or you are on a sheet where it is `tblEng03` | type `tblE` and take the autocomplete |
| `#VALUE!` in column K | a letter typed into H or I by accident | `Ctrl+Z` |
| `F4` did nothing | media-key laptop, or cell not in edit mode | `F2` first; `Fn+F4`; or type `$` by hand |
| `T5` shows `0` | reference is `[[#Headers],[budget]]` | retype `=SUM(tblEng02[budget])` |

### What you should know afterwards

- One formula in an empty Table column fills the whole column.
- Relative references move when copied. `$T$2` does not. `F4` toggles it.
- `tblEng02[budget]` is a structured reference and follows the table as it grows.
- `COUNTA` minus `COUNT` finds the blanks.
- A blank treated as zero produces a confident, wrong number. Nothing errors.

<details>
<summary>Check yourself (answers)</summary>

1. *`K2` contains `=H2-I2`. Copy it to `K7`. What is in `K7`?* `=H7-I7`.
2. *What do the dollar signs in `$T$2` do?* Lock the reference so it does not shift when
   copied.
3. *`COUNTA` says 40, `COUNT` says 34. What have you learned?* Six cells are not numbers.
   Here, blank.
4. *Why `=SUM(tblEng02[budget])` rather than `=SUM(H2:H41)`?* It follows the table when
   rows are added.
5. *Row 5 shows 100% margin. Is the formula wrong?* No. The data is missing.

</details>

---

# Part 3 - Errors, lookups and checks (`03_Checks`, about 50 minutes)

**What is on the sheet.** Parts 1 and 2 finished: `A1:M41` is a Table named
**`tblEng03`**, summary box filled. Two new things. At `X1:AC26` a second Table,
**`tblClients`**: 25 clients with `client_id`, `client_name`, `industry`, `country`,
`risk_rating`, `relationship_manager`. This is a **reference table**, the thing the `CLI-`
codes in column B have been pointing at all along. And at `AF1` an amber block headed
`SCRATCH - deliberate errors`. Columns N to Q are empty; that is where this part's work
goes.

**What this part is for.** You cannot read forty thousand rows, so you do not. You build
a small set of questions the file answers about itself and read the answers. It starts by
causing all five Excel errors on purpose, because they are the most honest thing in the
product, and most people panic at a hash symbol instead of reading it.

### Step 3.1 - Cause all five errors on purpose (10 min)

Everything typed in the amber block is meant to be wrong and stays wrong. Do not fix it.
The first column of each pair is a plain label; the second is the formula. `AG9` uses
`XLOOKUP`, which Step 3.3 explains; type it as written for now.

| Cell | Type | Cell | Type exactly | See |
| --- | --- | --- | --- | --- |
| AF6 | `DIV/0` | AG6 | `=1/0` | `#DIV/0!` |
| AF7 | `VALUE` | AG7 | `="apple"*2` | `#VALUE!` |
| AF8 | `NAME` | AG8 | `=SUME(H2:H10)` | `#NAME?` |
| AF9 | `N/A` | AG9 | `=XLOOKUP("CLI-999",tblClients[client_id],tblClients[client_name])` | `#N/A` |
| AF10 | `REF` | AG10 | `=AH10*2` | `0` for now |

Use the Name Box to reach `AF6`.

- `#DIV/0!`: divided by zero or an empty cell. You will meet it whenever a percentage's
  denominator is blank.
- `#VALUE!`: arithmetic on something that is not a number. This one dominates Part 4.
- `#NAME?`: a word Excel does not recognise. A misspelled function or table name,
  ninety-nine times in a hundred.
- `#N/A`: a lookup ran correctly and found nothing. The formula is perfect; there is no
  client CLI-999. **This one is a finding, not a mistake.** Two systems disagree about
  what exists.

Now for `#REF!`: right-click the column letter **`AH`** at the top of the grid >
**Delete**. **See:** `AG10` changes to `#REF!`. The formula points at a cell that no
longer exists. Leave column AH deleted; it was empty scratch. This is what happens when
somebody tidies a workbook by deleting a column they think is unused.

Say the letters to yourself twice before you delete. If you delete a column inside the
tracker by mistake, `Ctrl+Z` immediately.

### Step 3.2 - `IFERROR` (3 min)

`AG11`, no label needed, type, Enter:

```
=IFERROR(1/0,"Check the denominator")
```

**See:** the message. `IFERROR` tries the first thing; if it is any error at all, shows
the second. Two warnings. It catches everything, so write a message that says what
happened. And it does not fix anything; it hides it. Never use it to make a red thing go
away. The most common way a reconciliation break gets buried is an `IFERROR` around a
lookup with the message set to `""`.

### Step 3.3 - `XLOOKUP`: pull the client name in (12 min)

Column B holds codes like `CLI-001`. Column X holds the same codes beside names. Two
tables that share a key. Joining them is the single most common task in business data.

Click `N1`, type `client_name`, Enter. Column N turns blue and joins `tblEng03`.

Click `N2`. Type exactly (type `tblC` and `Tab` to autocomplete the table name), Enter:

```
=XLOOKUP(B2,tblClients[client_id],tblClients[client_name])
```

**See:** `N2` shows `Aldridge Manufacturing` and the column fills.

Read it as a sentence. `B2`: **what** am I looking for. `tblClients[client_id]`: **where**
do I look, one column. `tblClients[client_name]`: **what** do I bring back, one column of
the same height. Find this, in there, give me that.

Scroll to row 13. **See:** `#N/A`. Row 13 is `ENG-1012`, client `CLI-099`. Scroll the
client table; it stops at `CLI-025`. An engagement with 434,500 of budget booked to a
client that does not exist in the client master. That is the real one.

Click `O1`, type `risk_rating`. Click `O2`, type, Enter:

```
=XLOOKUP(B2,tblClients[client_id],tblClients[risk_rating])
```

**See:** `O2` shows `Low`, `O13` shows `#N/A`. Same key, different column brought back.

**This is exactly what session 05's sheet has already done for you.** Columns L and M of
`05_Pivots`, `client_name` and `risk_rating`, are this `XLOOKUP` applied to the
transactions table. Nobody typed them.

Somebody will mention `VLOOKUP`. It is the old one: it counts columns instead of naming
them, breaks when a column is inserted, and cannot look left. Read it in legacy files; do
not write it.

### Step 3.4 - Name the miss (4 min)

Click `N2`, press the `F2` key to edit, change it to:

```
=XLOOKUP(B2,tblClients[client_id],tblClients[client_name],"CHECK CLIENT ID")
```

Enter. **See:** `N13` reads `CHECK CLIENT ID`. Then `O2`, `F2` key, change it to:

```
=XLOOKUP(B2,tblClients[client_id],tblClients[risk_rating],"CHECK CLIENT ID")
```

Enter. **See:** `O13` reads `CHECK CLIENT ID`.

The fourth argument is `XLOOKUP`'s built-in if-not-found. Not blank, not "error", not a
dash: an instruction to a human, in capitals. That is the difference between hiding an
error and reporting one. The quotes are mandatory; without them, `#NAME?`.

### Step 3.5 - Two checks the file runs on itself (10 min)

`P1`: `id_count`. `P2`, type, Enter:

```
=COUNTIF(tblEng03[engagement_id],A2)
```

**See:** `1`, and the column fills, mostly with 1s. Each row asks "how many times does my
own ID appear in the whole column". Use the filter arrow on `id_count`, tick only `2`.
**See:** rows 18 and 34, both `ENG-1017`, with different clients, service lines and
budgets. Two pieces of work given one reference number. Every report that groups by ID
will silently add them together. Data > Sort & Filter > Clear.

`Q1`: `id_ok`. `Q2`, type, Enter:

```
=IF(LEN(A2)=8,"OK","CHECK LENGTH")
```

**See:** `OK`, column fills. `IF` is a question, an answer if yes, an answer if no. `LEN`
counts characters; every ID should be `ENG-` plus four digits, eight characters. Filter
`id_ok` to `CHECK LENGTH`. **See:** row 27, `ENG-104`, seven characters, budget 221,000.
Clear the filter.

`LEN` counts spaces, so an ID with a trailing space is nine characters and fails. That is
correct, and it sets up Part 4.

### Step 3.6 - Fix Part 2's lie (5 min)

Click `L2`, `F2` key, change it to:

```
=IF(ISBLANK(I2),"NO COST YET",K2/H2)
```

Enter. **See:** `L2` still `25.0%`. `L5`, `L11`, `L20`, `L29`, `L36`, `L41` now read
`NO COST YET`. Six cells that were confidently wrong now say "I do not know", which is the
most valuable thing a spreadsheet can say. `T11` is unchanged at `24.0%` because it sums
the variance column, not this one; the overall margin is still overstated and you now
know exactly why.

Column L is now a mix of numbers and text. Session 05 warns you not to drag it into a
pivot for that reason.

### Step 3.7 - Stop bad data getting in (5 min)

Select `J2:J41` (Name Box). Data tab > Data Tools group > **Data Validation** > Data
Validation. Allow: `List`. Source, type exactly:

```
Active,Complete,On Hold,Planned
```

Leave "In-cell dropdown" ticked. OK. Click `J2`; a dropdown arrow appears. Pick
`Complete`, then `Ctrl+Z`. Click `J3`, type `Actve`, Enter. **See:** a dialog, *"This
value doesn't match the data validation restrictions defined for this cell."* Cancel.

Think about what that typo would have done: every filter for Active misses the row, every
count is one short, every pivot grows an extra category. Validation only checks new
typing; anything already wrong in the column stays wrong. On a US-configured machine the
list separator is a comma; on some European ones it is a semicolon.

### If it goes wrong

| You see | Cause | Fix |
| --- | --- | --- |
| `AG10` still shows `0` after the delete | a different column was deleted | `Ctrl+Z`, then right-click the letter `AH` itself |
| Tracker columns shifted or `T5` broke | a column inside the data was deleted | `Ctrl+Z` immediately |
| `AG7` shows `#NAME?` instead of `#VALUE!` | quotes left out, so Excel read `apple` as a name | retype with the quotes |
| `#NAME?` instead of a client name | `tblClients` misspelled | type `tblC`, Tab |
| Only `N2` filled | not inside the Table | `N2`, `Ctrl+C`, Name Box `N3:N41`, `Ctrl+V` |
| `ISBLANK` row still shows a percentage | `I5` still holds `59500` from Step 2.5, or a space | click `I5`, Delete |
| `Actve` was accepted | the rule was applied to a different range | reselect `J2:J41`, redo |

### What you should know afterwards

- The five errors and what each says. `#N/A` is usually information, not a defect.
- `IFERROR` hides; a named fallback like `CHECK CLIENT ID` reports.
- `XLOOKUP(what, where, bring_back, [if_not_found])`. Session 05's `client_name` and
  `risk_rating` columns are this.
- `COUNTIF` finds duplicates. `LEN` finds malformed IDs. `ISBLANK` names the gaps.
- Data validation is the cheapest control there is and checks new input only.

<details>
<summary>Check yourself (answers)</summary>

1. *Difference between `#N/A` and the other four?* The others mean the formula is wrong.
   `#N/A` usually means the data is missing.
2. *Why is `IFERROR(...,"")` dangerous?* Every failure becomes a blank. Breaks vanish.
3. *Three `XLOOKUP` arguments in English?* What am I looking for; where; what do I bring
   back.
4. *You add validation to a column that already holds a typo. What happens to the typo?*
   Nothing.

</details>

---

# Part 4 - Data cleaning (`04_Cleaning`, about 70 minutes)

**What is on the sheet.** Columns A to J, rows 1 to 313: **312** transactions exported
from a payments system, every cell formatted as Text on purpose. Column K is empty and
stays empty until Step 4.3 puts the country there. L is a spacer. M to Q are where the
clean versions go. R is empty. Column S holds a grey list of every defect in the data.
Scratch cells are `S30:S32`, under that note.

**What this part is for.** Analysts spend half to three quarters of their time on this.
None of it errors: a duplicate row is a valid row, a number stored as text is valid text,
and Excel will sum a column, ignore the two hundred text cells, and hand you a wrong total
with total confidence. The order is always the same: remove what should not be there,
split what is glued together, mark what is missing, standardise shapes, turn text into
numbers, then throw away the workings.

**The defects, all real:**

| Defect | Where | How many |
| --- | --- | --- |
| Exact duplicate rows | anywhere | 8 |
| Near duplicates, same `txn_id`, one field changed | anywhere | 4 |
| `txn_id` in four shapes: `TX0004`, `tx0001`, `TX-0002`, ` TX0003 ` | A | most rows |
| Three date formats, all text | B | all rows |
| `client_id` sometimes lower case | C | about a quarter |
| `txn_type` blank | D | 20 rows as loaded, 18 once the duplicates are gone |
| Amounts as text: `$7,477.00`, `17,413.00`, `2496.00` | F | all rows |
| Currency as ` usd ` | G | about a ninth |
| Counterparty and country in one column, split by a pipe character | J | all rows |
| Counterparty in random case, with double spaces | J | most rows |
| Stray tab characters at the end of counterparty names | J | some rows |
| Non-breaking spaces inside counterparty names | J | some rows |

**Copying down.** Every helper column in this part is a plain range, not a Table, so
nothing fills by itself. "Copy down to row 301" means: click the formula cell, `Ctrl+C`,
Name Box, type the range below it such as `P3:P301`, Enter, `Ctrl+V`. Do not use
`Ctrl+Shift+Down` from the formula cell: the column below is empty, so it runs to the
bottom of the sheet.

### Step 4.1 - Look at the damage, prove the amounts are fake (5 min)

Click `A2` and read across. `A2` is `tx0001`, lower case. `B2` is `01/01/2026` as text.
`D2` is empty. `F2` is `$7,477.00` as text. `G2` is ` usd ` with spaces round it. `J2` is
`ALPINE  FREIGHT SERVICES`, then an invisible tab, then `|Canada`: two facts in one cell,
and a mess. Columns C, E, H and I on this row are fine. One row, seven problems, and it
is just row two.

`Ctrl+End`. Row 313, so 312 data rows.

Name Box `S30`, Enter, type, Enter:

```
=SUM(F2:F313)
```

**See:** `0`. Three hundred and twelve payments add up to nothing, because every one is
text and `SUM` ignores text. Not an error. Zero.

Look at column F. Everything is jammed to the **left**. Left means text; numbers and
dates sit on the right on their own. That alignment is the fastest diagnosis in Excel.

Clear `S30` with **Home > Editing > Clear > Clear All**, not the Delete key. Column F is
Text, so `S30` picked up Text format from it; Delete keeps that format, and the next
formula you type there would be stored as words. Clear All removes both.

### Step 4.2 - Duplicates, in two passes (10 min)

**Warning:** Remove Duplicates deletes rows permanently, with no preview. `Ctrl+Z` is
the only way back. On real data, save a copy first.

Click one cell in the data, for example `C5`. Data > Data Tools > **Remove Duplicates**.
A dialog lists every column, all ticked, "My data has headers" ticked. Check that tick
every time; unticked, Excel treats row 1 as data and may delete it. All ten columns
ticked means: only delete a row if every field matches another. OK.

**See:** *"8 duplicate values found and removed; 304 unique values remain."* OK. Click
`A2`, `Ctrl+Down`: row 305.

Click a data cell again. Data > Remove Duplicates. Click **Unselect All**, then tick
**only `txn_id`**. OK. **See:** *"4 duplicate values found and removed; 300 unique
values remain."* `A2`, `Ctrl+Down`: row 301.

The second pass catches the re-keyed rows: same ID, one field different, so not identical,
but the same payment. Ticking only `txn_id` says "this column is the identity of the
row". Excel keeps the **first** occurrence; here the original came first, which is luck.
On a real job, sort by an entry timestamp first so "first" means something.

If you select one column before running it, Excel offers "Continue with the current
selection" or "Expand the selection". Always expand. The first deletes cells from one
column and shifts it up, wrecking every row.

### Step 4.3 - Split two facts in one column (8 min)

Click `J2`, look at the formula bar: name, pipe, country. You cannot group by country
while it is glued to a name.

Click `J2`, then `Ctrl+Shift+Down`. `J2:J301` is selected. Data > Data Tools > **Text to
Columns**.

1. Step 1 of 3: **Delimited**. Next.
2. Step 2 of 3: **untick `Tab`**. Tick **Other** and type a pipe `|` in its box (`Shift`
   and backslash on a US keyboard). The preview splits into two columns. Next.
3. Step 3 of 3: leave **General**. Finish.

**See:** column J holds names only, still messy: `J2` is `ALPINE  FREIGHT SERVICES`.
Column K holds countries: `K2` is `Canada`.

Click `J1`, type `counterparty`. Click `K1`, type `country`.

Text to Columns writes into the columns to the right. That is why K was left empty; if it
had held data, Excel would have asked to replace it. And in step 3, if you ever split
ID-like data, set the column to **Text**, or Excel turns things like `01-234` into dates.

### Step 4.4 - Fill the blanks in one move (6 min)

Name Box `D2:D301`, Enter. Do not use `Ctrl+Shift+Down` from `D2`: `D2` is itself blank,
so the shortcut stops after two cells.

`S30`: `=COUNTBLANK(D2:D301)`. **See:** `18`. (Twenty when the sheet opened; two sat on
duplicate rows.) Clear All on `S30`.

Name Box `D2:D301` again. Press `F5` (`Fn+F5` on some laptops, or `Ctrl+G`) > **Special**
> **Blanks** > OK.
**See:** only the empty cells in D are selected, scattered down the column. The status
bar shows no count for an all-blank selection, which is why you counted first.

Without clicking anything, type:

```
Unclassified
```

Press **`Ctrl+Enter`**, not Enter. **See:** all 18 blanks fill at once. Plain Enter fills
one cell and wastes the selection. Proof if you want it: `S30`,
`=COUNTIF(D2:D301,"Unclassified")` gives `18`. Clear All afterwards.

Marking a gap is honest. Filling it with the row above is a guess, and on a regulated
file a guess is fabrication. Go To Special will also select every formula on a sheet, or
every constant. Worth remembering.

### Step 4.5 - The counterparty column, one function at a time (12 min)

Four things are wrong with this column at once. Watch each layer fail before adding the
next. Every edit to `P2` needs its own copy down afterwards; this is a plain range.

`P1`: `clean_counterparty`. `P2`, type, Enter, copy down to `P301`:

```
=TRIM(J2)
```

**See:** `ALPINE FREIGHT SERVICES` with the double space gone. Still an invisible
character in front and a tab at the end. `TRIM` removes ordinary spaces from the ends and
collapses runs inside. Only ordinary spaces. You cannot see either leftover, so prove
them: in `S30` type `=CODE(P2)` and get `160`, the code of the first character, which is
not the letter A. In `S31` type `=LEN(P2)` and get `26`, three more than the 23 letters
and spaces you can see: characters `TRIM` could not remove. Leave both in place; they
update as you go.

`P2`, `F2` key, change to, Enter, copy down to `P301`:

```
=TRIM(CLEAN(J2))
```

**See:** `S31` drops to `24`. The tab and its neighbour are gone. `CLEAN` strips non-printing characters. `CLEAN` sits **inside** `TRIM`
because Excel works inside out: `CLEAN` runs first, `TRIM` tidies what it leaves. The
other way round leaves a trailing space nobody can see.

`P2` again, Enter, copy down:

```
=TRIM(CLEAN(SUBSTITUTE(J2,CHAR(160)," ")))
```

**See:** exactly `ALPINE FREIGHT SERVICES`, nothing in front. `S30` now reads `65`, the
letter A, and `S31` reads `23`. `CHAR(160)` is the
**non-breaking space**. It looks like a space, prints like a space, is not a space, and
`TRIM` will not touch it. Two cells that read identically and refuse to match in a lookup:
this is the first thing to suspect. Anything through a PDF or a web page is full of them.

`P2`, final version, Enter, copy down to `P301`:

```
=PROPER(TRIM(CLEAN(SUBSTITUTE(J2,CHAR(160)," "))))
```

**See:** `P2` reads `Alpine Freight Services`, `P3` `Petra Energy Partners`. Clear `S30`
and `S31` with Home > Clear > Clear All. Read a nested
formula from the inside out: swap fake spaces for real, clean the junk, trim the edges,
set the case. `PROPER` is not always right; it turns `KPMG LLP` into `Kpmg Llp`.

### Step 4.6 - Transaction ID (4 min)

`M1`: `clean_txn_id`. `M2`, type, Enter, copy down to `M301`:

```
=UPPER(SUBSTITUTE(TRIM(A2),"-",""))
```

**See:** `TX0001`, `TX0002`, `TX0003`. Trim, delete hyphens, force upper case. `TX0002`,
`TX-0002` and ` tx0002 ` are three different strings to a computer, and a lookup between
them returns nothing. Every reconciliation failure starts as a formatting difference in a
key field. The `""` is two double quotes, meaning "replace with nothing".

### Step 4.7 - Dates (6 min)

`N1`: `clean_date`. `N2`, type, Enter, copy down to `N301`:

```
=DATEVALUE(B2)
```

**See:** `46023`. A real date without its costume. Select `N2:N301`, Home > Number >
dropdown > **Short Date**. **See:** `1/1/2026`, right-aligned.

`DATEVALUE` read three formats in column B without being told which was which, using the
machine's regional settings. On a US machine `03/04/2026` is 4 March; on a UK machine it
is 3 April. Same formula, two answers, no error. If you receive a file from another
office, check a date you know before you trust the column.

### Step 4.8 - Amounts, and the payoff (10 min)

`O1`: `clean_amount`. `O2`, type, Enter, copy down to `O301`:

```
=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(F2),"$",""),",",""))
```

**See:** `7477`, and the column jumps to the **right** on its own. That movement is your
receipt. Trim, strip dollars, strip commas, then `VALUE` turns text into a number. Two
`SUBSTITUTE`s because each handles one character.

Now the evidence, in the scratch cells:

| Cell | Type | See |
| --- | --- | --- |
| S30 | `=SUM(O2:O301)` | `1865697` (Step 4.1 gave 0 for the same money) |
| S31 | `=COUNT(O2:O301)` | `300`, so every amount converted |
| S32 | `=COUNTA(UNIQUE(P2:P301))` | `20` distinct counterparties |

**1,865,697 is the number session 05 opens with.** Your first pivot's grand total is it.
`Ctrl+1` > Number > 2 decimals > Use 1000 Separator if you want `S30` to read
`1,865,697.00`.

`UNIQUE` spills its answers into the cells below unless wrapped in `COUNTA`; if twenty
names appear down column S, you typed it without the `COUNTA`, and Clear All on `S32`
removes them all at once. Then select `S30:S32`, Home > Clear > Clear All.

### Step 4.9 - Client ID (2 min)

`Q1`: `clean_client_id`. `Q2`, Enter, copy down to `Q301`:

```
=UPPER(TRIM(C2))
```

**See:** `CLI-001`, all upper case. Same pattern: trim, force case, stop.

### Step 4.10 - Freeze the result, throw away the workings (6 min)

Right now M to Q are formulas pointing at A to J. Delete a source column, or sort half
the sheet, and they turn to hash. A cleaned dataset should not depend on the mess it came
from.

Click `M1`, then `Ctrl+Shift+Right`, then `Ctrl+Shift+Down`. `M1:Q301` is selected.
`Ctrl+C`. Without moving the selection, right-click inside it and under **Paste Options**
click the **Values** icon (the clipboard with `123`). Keyboard: `Ctrl+Alt+V`, `V`, Enter.
Nothing visibly changes; click `O2` and the formula bar shows `7477`, not a formula.

Proof: click `A2`, press Delete, look at `M2`. Still `TX0001`. `Ctrl+Z`.

Only ever do this at the end. There is no way back except undo. The most common mistake
is copying, clicking somewhere else, then pasting: paste over the same selection.

### Step 4.11 - Where the clean data went (1 min)

Click `05_Pivots` and look. Same 300 rows, tidied into a Table called `tblTxn` in
`A1:M301`, with `client_name` and `risk_rating` joined on the end by the `XLOOKUP` from
Part 3. Do not do any work on it yet. That is where session 05 starts.

### If it goes wrong

| You see | Cause | Fix |
| --- | --- | --- |
| A formula in `S30` shows itself as text | you used Delete instead of Clear All in Step 4.1 | Clear All on `S30`, retype |
| A formula in M to Q shows itself as text | that cell picked up Text format somehow; the sheet ships them as General | Clear All on the cell, retype |
| Dedup counts not 8 and 4 | passes run in the wrong order, or one skipped | `Ctrl+Z` back to 312 rows, start Step 4.2 again |
| `K2` blank and `Canada` landed in `L2` | `Tab` left ticked in the wizard | `Ctrl+Z`, clear L, redo with Tab unticked |
| "There's already data here" from the wizard | something typed into K | Cancel, clear K, redo |
| Go To Special says "No cells were found" | wrong selection | Name Box `D2:D301`, try again |
| Only one cell filled with `Unclassified` | plain Enter instead of `Ctrl+Enter` | `Ctrl+Z`, reselect, redo |
| `P3` still shows an older version of the name | copy down skipped after editing `P2` | copy down again |
| `#VALUE!` from `DATEVALUE` on some rows | machine set to a day-first region; the third date shape does not parse | Windows Settings > Time & Language > Region > Regional format: United States, then reopen Excel |
| `#VALUE!` from `DATEVALUE` on every row | typo in the formula | retype |
| `Ctrl+Alt+V` did nothing | machine intercepts it | right-click > Paste Options > Values |

### What you should know afterwards

- Left-aligned means text. `SUM` ignores text and returns zero without complaint.
- Duplicates take two passes: all columns, then the key column only. Remove Duplicates is
  permanent.
- Text to Columns splits on a delimiter and writes to the right. Untick Tab.
- Go To Special > Blanks, then `Ctrl+Enter`, fills every gap with one label. Do not guess.
- `PROPER(TRIM(CLEAN(SUBSTITUTE(...,CHAR(160)," "))))`, read inside out. Nesting order
  matters.
- `VALUE` and `DATEVALUE` turn text into numbers and dates. `DATEVALUE` trusts the
  machine's region.
- Paste Values at the end, and only at the end.

<details>
<summary>Check yourself (answers)</summary>

1. *You sum 300 amounts and get zero. First check?* Which side of the cell they sit on.
2. *Remove Duplicates with every column ticked removed 8, but 4 bad rows survived. Why?*
   They shared an ID and differed in one field. Second pass on the ID column.
3. *`TRIM` did not remove a leading space. What is it?* A non-breaking space, `CHAR(160)`.
   `SUBSTITUTE` first, then `TRIM`.
4. *Why `Unclassified` rather than copying the row above?* The row above is a guess.
5. *Why Paste Values at the end?* So the clean data no longer depends on the messy source.

</details>

---

# Before session 05: what it assumes you know

Session 05 refers back to sessions 01 to 04 without re-explaining. Here is the list, so
nothing comes as a surprise.

| It says | It means | Where you did it |
| --- | --- | --- |
| "Click inside `tblTxn`" | an Excel Table, named on the Table Design tab | Part 1, Step 1.7 |
| "the fourth thing Tables buy you" | it grows by itself, keeps its header on screen, can be referred to by name, and now: a pivot finds its range unaided | Part 1, Step 1.7 |
| "the same number you calculated yesterday with a SUM", 1,865,697 | your `S30` in Part 4 Step 4.8 | Part 4 |
| "Columns L and M were joined on for you" | `XLOOKUP` against `tblClients` | Part 3, Step 3.3 |
| "we turned three formats of text into real dates with DATEVALUE" | Part 4 Step 4.7 | Part 4 |
| "the COUNT versus COUNTA check" | blank detector, 40 against 34 | Part 2, Step 2.4 |
| "the row-5 problem" | a blank cost treated as zero, reported as a 100% margin | Part 2, Step 2.2 |
| "yesterday's summary box", 11,645,000 / 8,847,400 / 2,797,600 | `T5:T11` on `02_Formulas` | Part 2, Step 2.4 |
| "the `PROPER` step" | collapsing every spelling of a counterparty into one; the count of 20 in Step 4.8 is the proof | Part 4, Steps 4.5 and 4.8 |
| "writing CHECK CLIENT ID instead of leaving a blank" | a named fallback in `XLOOKUP` | Part 3, Step 3.4 |
| "a filter left on from yesterday" | filter arrows on a Table, Data > Clear | Part 1, Step 1.6 |
| "do not drag `margin_pct` into Values" | column L on `03_Checks` is now a mix of numbers and text | Part 3, Step 3.6 |
| "right-click, Number Format, every time" | format inside the pivot so it survives the pivot changing shape; the same Number formats as Part 1, reached a different way | Part 1, Step 1.3 |
| "`Alt+F5` to refresh" | new in 05; nothing to do beforehand | - |

Five numbers to have in your head walking in:

| Number | What it is |
| --- | --- |
| 300 | rows in `tblTxn` |
| 1,865,697 | total of all amounts |
| 40 and 34 | `COUNTA` and `COUNT` of the engagement tracker |
| 11,645,000 / 8,847,400 / 2,797,600 | budget, actual, variance totals |
| 20 | distinct counterparties after cleaning |

## Reading, if you want it

Ask your trainer for the concept notes: one per topic, ten to twenty minutes each,
written to stand on their own: `tables-and-structure.md`,
`formulas-and-references.md`, `errors-and-validation.md`, `data-cleaning.md`. Read them
after the matching part, not before. `pivot-tables.md` is the one for session 05.
