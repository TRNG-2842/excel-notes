# Data cleaning

## What it is and why it exists

Cleaning is turning a file that a system produced into a file you can calculate on. Ask
any working analyst and they will tell you it takes more of their time than the analysis
itself, and that it always has.

The reason is structural rather than fixable. Data is entered by people, or by systems
built by people, in different countries, at different times, into fields nobody validated.
By the time three systems have exchanged a record, the same counterparty is spelled four
ways, half the amounts are text, and two facts have been jammed into one column to save
space.

What makes it dangerous rather than merely tedious is that **none of it errors**. A
duplicate row is a valid row. A number stored as text is valid text. Excel will sum three
hundred amounts, silently ignore the two hundred that are text, and hand you a total with
no warning at all. The output is not an error message. It is a wrong number that looks
exactly like a right one.

## How it behaves

### The order, which does not change

1. **Remove what should not be there.** Duplicates first, because everything after is
   cheaper on fewer rows and because deduplicating cleaned data is harder than it looks.
2. **Split anything holding two facts.** One column, one fact.
3. **Fill or mark what is missing.** Mark, generally. See below.
4. **Standardise shapes.** One date format, one case, one way of writing a key.
5. **Convert text that means a number into a number.**
6. **Freeze the result and throw away the workings.**

### Duplicates, in two passes

Data > Remove Duplicates. You choose which columns count as "the same".

- **All columns ticked** deletes a row only if every field matches another row. This is the
  strict test and where you always start.
- **Only the key column ticked** deletes a row if the identifier matches, whatever else
  differs. This catches the re-keyed near-duplicate: same reference, amount a few dollars
  out, entered by somebody correcting a mistake and creating a second one.

Two things to know before you run it. Excel keeps the **first** occurrence and deletes
those below it, so on real data you sort by an entry timestamp first, and you know which
"first" you mean. And it is permanent, with no preview and no confirmation beyond the
dialog: save a copy before running it on anything that matters.

Deduplicate on the **cleaned** key, never the raw one. Two references differing only by a
trailing space are not duplicates as far as Excel is concerned.

### Splitting one column into two

Data > Text to Columns. Choose Delimited, tick the character the fields are separated by,
and finish.

It writes into the columns to the **right** of the one you split. If those cells already
hold anything, Excel asks "There's already data here. Do you want to replace it?", and Yes
overwrites them. Insert enough blank columns first, so the question never comes up.

One trap worth carrying: in the wizard's third step you can set each resulting column's
type. Leave it as General and Excel applies its usual guesses, which will turn an
identifier like `03-118` into a date. When splitting anything ID-shaped, set the column
type to **Text**.

### The text-tidying functions

| Function | Does | Does not |
|---|---|---|
| `TRIM(t)` | removes leading and trailing spaces, and collapses internal runs of spaces to one | remove non-breaking spaces |
| `CLEAN(t)` | removes non-printing characters, codes 0 to 31: line feeds, tabs, control codes | remove spaces of any kind |
| `PROPER(t)` | Capitalises Each Word | know that an acronym like `ACME LLP` is not `Acme Llp` |
| `UPPER(t)` / `LOWER(t)` | force case | anything else |
| `SUBSTITUTE(t, find, replace)` | swap every occurrence | handle more than one target per call |

**The non-breaking space.** Character 160. It looks exactly like a space, prints like a
space, and `TRIM` will not touch it. Two cells reading the same company name can look
identical and refuse to match, and you can lose an afternoon to it. Anything that has been
through a web page or a PDF is full of them. The fix:

```
=SUBSTITUTE(text, CHAR(160), " ")
```

then `TRIM`. Whenever two values that are obviously the same will not match, suspect this
first.

Before and after, with `.` standing for a space and `<LF>` for a line feed so they can
be seen:

| Formula | Input | Result |
|---|---|---|
| `=TRIM(A2)` | `..Acme...Ltd.` | `Acme.Ltd` |
| `=CLEAN(A2)` | `Acme.Ltd<LF>` | `Acme.Ltd` |
| `=PROPER(A2)` | `ACME LLP` | `Acme Llp` |
| `=UPPER(A2)` | `acme ltd` | `ACME LTD` |
| `=SUBSTITUTE(A2,",","")` | `1,250` | `1250` |
| `=LEN(TRIM(A2))` | `Acme.Ltd.` where the last space is character 160 | `9`, because `TRIM` left it |

**Nesting order matters, and not in the way people expect.** Excel evaluates inside out.

```
=PROPER(TRIM(CLEAN(SUBSTITUTE(t, CHAR(160), " "))))
```

Substitute the fake spaces for real ones, strip the invisible junk, tidy the edges and the
doubles, then set the case. Put `TRIM` inside `CLEAN` instead and you get a subtly wrong
answer: `TRIM` runs first, `CLEAN` then removes a trailing tab, and the space that
was in front of it is left behind as a trailing space nobody can see. Every value on those
rows now fails to match. **`TRIM` goes outside `CLEAN`.**

### Converting text to numbers and dates

```
=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(t),"$",""),",",""))
=DATEVALUE(t)
```

`VALUE` needs the text stripped of currency symbols and thousands separators first; each
`SUBSTITUTE` handles one character, so they nest.

`DATEVALUE` copes with several formats at once, including `03/14/2026`, `14-Mar-2026` and
`2026-03-14`. It returns a serial number, so apply a date format afterwards.

| Formula | Input | Result |
|---|---|---|
| `=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""))` | `.$1,250.00.` | `1250` |
| `=VALUE(A2)` | `$1,250.00` | `#VALUE!` - the symbols were not stripped |
| `=DATEVALUE(A2)` | `2024-03-14` | `45365`, which shows as `14-Mar-2024` once formatted |
| `=DATEVALUE(A2)` | `14 March 2024` | `45365` |

The caution that matters if a file crosses borders: `DATEVALUE` trusts the machine's
regional settings. `03/04/2026` is 4 March on a US-configured machine and 3 April on a UK
one. Same file, same formula, two answers, no error either time. Check a date you know the
answer to before trusting a converted column.

**The alignment tell.** This is the fastest diagnosis in Excel and it is free. Text sits
against the **left** of a cell. Numbers and dates sit against the **right**, without being
told. Before doing anything else to a file you have been handed, look at which way the
columns lean. A column of amounts hard against the left is text, and `SUM` over it returns
zero. When a conversion works, the column visibly jumps to the right. That movement is your
receipt.

### Blanks

Select the column, press `F5`, click Special, choose Blanks. Only the empty cells are
selected. Type a value and press `Ctrl+Enter` - not Enter - to fill all of them at once.

Then the judgement call, which is not a technical one. **Mark the gap; do not guess it.**
Filling blanks with the value from the row above is the instinct and it is almost always
wrong: you have invented data, and on a regulated file that is fabrication. `Unknown`
is a fact. A copied value is a claim you cannot support.

### Freezing the result

At the end, select the cleaned columns, copy, and Paste Values. The formulas are replaced
by their answers, and the cleaned data stops depending on the mess it came from. Delete or
re-sort the source afterwards and nothing breaks.

Do it **only** at the end. There is no way back except undo.

Finish with three evidence cells, because somebody will ask how you know it worked:

```
=COUNT(cleaned_amounts)          should equal your row count
=SUM(cleaned_amounts)            compare against the source system's total
=COUNTA(UNIQUE(cleaned_names))   distinct values, which should have fallen sharply
```

## Cost and trade-offs

- **It is manual and it is not repeatable.** Everything here has to be done again next
  month, by hand, the same way. If the job recurs, this is the wrong tool: use Power Query,
  which records the steps as a recipe you refresh with a button.
- **Paste Values is irreversible** once saved.
- **`PROPER` damages acronyms.** Any name that is an initialism comes back title-cased.
  Check your data before reaching for it.
- **Remove Duplicates is irreversible** and silently destructive if run with the wrong
  columns ticked.
- **Flash Fill (`Ctrl+E`)** guesses a pattern from examples. It is fast, impressive, and
  unauditable: nobody can review a guess. Avoid it on anything that matters.
- **Helper columns multiply.** Five working columns on a twenty-column file is normal and
  should all disappear at the end.
- **`UNIQUE` and the other dynamic array functions are Microsoft 365 only.** On older
  builds they return `#NAME?`.

## Recognize it on sight

- Numbers hard against the left of their cells. Text pretending to be money.
- A total of exactly `0` on a column that visibly contains numbers.
- Green triangles in the top-left corner of cells: Excel's own "number stored as text"
  warning, which people switch off.
- The same name appearing in three cases, or twice with different spacing.
- Two facts separated by a punctuation mark inside one column.
- A column of dates sorting with March before January: text, not dates.
- `########`: a number in a column too narrow to display it. Not an error.

## Adjacent question

*"This file arrives every month. How do you avoid doing all of this again?"*

Power Query. Data > Get Data, load the file, perform the same steps in its editor, and it
records them as a query. Next month you drop in the new file and press Refresh. It handles
splitting, trimming, type conversion, deduplication and appending many files from a folder,
and unlike a column of formulas the steps are listed, named and reviewable. If you can only
learn one more Excel feature after the basics, this is it.

The second half of the answer is upstream: if you own the input, validation at the point of
entry removes the need for most of this.

## Say it in an interview

"The first thing I do with a file is look at which way the columns lean. Text aligns left,
numbers align right. A column of amounts sitting on the left means `SUM` will return zero
without any error at all, and that is the most common silent failure in finance."

"I deduplicate in two passes: once on every column to catch exact repeats, then once on the
key alone to catch re-keyed rows that differ in one field. And always on the cleaned key,
because two references differing by a trailing space are not duplicates to Excel."

"I mark missing data rather than guessing it. Copying the value from the row above invents
data, and nobody downstream can tell the difference between a number I calculated and one I
made up."

"`TRIM` does not remove non-breaking spaces, character 160. That single fact has saved me
more time than any function I know, because it is the usual reason two values that are
obviously identical refuse to match."

"If the file arrives monthly I would not do any of this by hand. Power Query records the
steps and refreshes."

## Check yourself

1. You `SUM` three hundred amounts and get zero. What is the first thing you check, and
   why is there no error?
2. Remove Duplicates with all columns ticked removed eight rows, and four bad rows
   survived. Why, and what do you do?
3. `TRIM` has not removed what is clearly a leading space. What is it and what removes it?
4. Why must `TRIM` sit outside `CLEAN` rather than inside it?
5. Why fill blank categories with a label rather than the value from the row above?
6. What does Paste Values buy you at the end of a clean-up?
7. Your colleague's converted date column is a day out from yours on the same file. What
   happened?

<details>
<summary>Answers</summary>

1. The alignment. Left-aligned means text, and `SUM` ignores text entirely, returning zero
   rather than an error because ignoring text is documented, correct behaviour.
2. Those four were not identical: they share a key and differ in one field. Run a second
   pass with only the key column ticked.
3. A non-breaking space, character 160. `SUBSTITUTE(text,CHAR(160)," ")` first, then `TRIM`.
4. Excel evaluates inside out. `TRIM` inside means it runs before `CLEAN` removes a trailing
   tab, leaving the space that preceded it stranded at the end of the string.
5. Because the row above is a guess. A label records that the value is unknown, which is a
   fact; copying invents data you cannot support.
6. The cleaned data stops depending on the source. Delete, re-sort or overwrite the messy
   original and the clean version is unaffected.
7. `DATEVALUE` reads ambiguous dates using the machine's regional settings, so `03/04/2026`
   is 4 March on one and 3 April on another. Neither errors.

</details>

## Resources

- TRIM function: https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9
- CLEAN function: https://support.microsoft.com/en-us/office/clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41
- SUBSTITUTE function: https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332
- Split text into different columns with the Convert Text to Columns Wizard: https://support.microsoft.com/en-us/office/split-text-into-different-columns-with-the-convert-text-to-columns-wizard-30b14928-5550-41f5-97ca-7a3e9c363ed7
- Filter for unique values or remove duplicate values: https://support.microsoft.com/en-us/office/filter-for-unique-values-or-remove-duplicate-values-ccf664b0-81d6-449b-bbe1-8daaec1e83c2
- About Power Query in Excel: https://support.microsoft.com/en-us/office/about-power-query-in-excel-7104fbee-9e62-4cb9-a02e-5bfb1a6c536a
