# Errors, and building checks a file runs on itself

## What it is and why it exists

Excel has five error values. They are not crashes and they are not the product failing.
They are the most honest thing in the application: each one tells you, in a single word,
exactly what could not be done.

Most people see a hash symbol and reach for a way to make it disappear. That instinct is
the problem. An error is information, and one of the five is usually not an error at all
but a finding about your data.

The larger idea is **validation**: instead of reading a file to see whether it is correct,
you build a small set of questions the file answers about itself, and you read the answers.
On a hundred rows you could read the data. On a hundred thousand you cannot, and the questions are
the only thing that scales. Checking whether a file is trustworthy is most of what a junior
analyst is paid for.

## How it behaves

### The five errors

| Error | Means | Usual cause |
|---|---|---|
| `#DIV/0!` | division by zero or by an empty cell | a percentage whose denominator is missing |
| `#VALUE!` | arithmetic on something that is not a number | a number that arrived as text; a stray letter |
| `#NAME?` | Excel does not recognise a word | misspelled function, misspelled table name, or a function this version does not have |
| `#N/A` | a lookup ran and found nothing | the key genuinely is not in the other table |
| `#REF!` | a formula points at a cell that no longer exists | somebody deleted a row, column or sheet |

Excel has a few more - `#NUM!` for a calculation that cannot be done, `#NULL!` for a
mistyped range, `#SPILL!` when a dynamic array has no room to land - but these five are the
ones you will meet every day.

`#N/A` is the odd one out and deserves separate thought. The other four mean *the formula
is wrong*. `#N/A` usually means *the formula is right and the data is missing*. When a
lookup between two systems returns `#N/A`, you have discovered that the systems disagree
about what exists. That is a finding to escalate, not a defect to suppress.

### `IFERROR`, and why it is dangerous

```
=IFERROR(something, fallback)
```

Try `something`. If it comes back as **any** error, show `fallback` instead.

Two things follow from "any".

First, it is indiscriminate. A typo, a missing lookup, a deleted column and a division by
zero are all caught by the same net, so a genuine structural break can be hidden by a
wrapper you added for a cosmetic reason.

Second, and this is the part that matters professionally: **`IFERROR` does not fix
anything. It hides it.** `=IFERROR(XLOOKUP(...),"")` produces a file that looks clean and
has lost its most important information. Somebody will make a decision on that file.

The working rule: use it where the error is **expected and meaningful**, and make the
fallback say what happened. `"CHECK CUSTOMER ID"` is a fallback. `""` is a cover-up.

Where a function offers a built-in not-found argument, prefer it. `XLOOKUP` has one as its
fourth argument, and unlike `IFERROR` it catches only the not-found case, leaving real
errors visible.

### Checks a file runs on itself

Four questions cover most of what goes wrong with tabular data. Each is one formula.

**Is anything recorded twice?**

```
=COUNTIF(tblData[id], A2)
```

Returns how many times this row's own id appears in the whole column. The honest answer is
`1`. Filter for anything else.

| `A` | Result |
|---|---|
| `C-0417` | `1` |
| `C-0552` | `2` |
| `C-0552` | `2` | Note that a duplicated key is often worse than a duplicated
row: two genuinely different records sharing a reference will be silently added together by
everything downstream.

**Is any key the wrong shape?**

```
=IF(LEN(A2)=10,"OK","CHECK LENGTH")
```

`LEN` counts characters, including spaces, which is a feature: a value with a trailing
space fails, and it should, because it will not match anything.

| `A` | `LEN` | Result |
|---|---|---|
| `AB12345678` | 10 | `OK` |
| `AB1234567` | 9 | `CHECK LENGTH` |
| `AB12345678 ` (trailing space) | 11 | `CHECK LENGTH` |

**Is anything missing?**

```
=IF(ISBLANK(C2),"NOT RECORDED YET",D2/B2)
```

`ISBLANK` is true only for a genuinely empty cell. Not zero, not a space, not an empty
string returned by a formula.

| `B` | `C` | `D` | Result |
|---|---|---|---|
| 4000 | 2024-03-14 | 3200 | `0.8` |
| 4000 | (empty) | 3200 | `NOT RECORDED YET` | Writing "NOT RECORDED YET" instead of calculating is the
most valuable thing a spreadsheet can do, because a reader cannot otherwise distinguish a
number you calculated from a number you invented.

**Does every code point at something real?**

```
=XLOOKUP(B2, tblRef[id], tblRef[name], "CHECK ID")
```

A fallback that reads as an instruction to a human being.

| `B` | Result |
|---|---|
| `C-0417` | `Halden Foods` |
| `C-0999` | `CHECK ID` |

### Stopping bad data at the door

Data Validation puts a rule on a cell and refuses input that breaks it. Select the range,
Data > Data Validation, Allow: List, and type the permitted values separated by your
machine's list separator:

```
Open,In progress,Closed,On hold
```

Type `Clsoed` into a validated cell and Excel refuses it outright.

Consider what that misspelling would otherwise do. Every filter for `Closed` misses the
row. Every count is one short. Every PivotTable grows a category called `Clsoed` with one
row in it. Nothing is red and nobody notices.

Three things to know about it:

- It only applies to **future** typing. Bad values already in the column are untouched, and
  the rule sits there looking useful while doing nothing.
- It applies only to the cells you selected. New rows added to an Excel Table inherit it; a
  plain range does not.
- The list separator is regional: a comma on a US-configured machine, often a semicolon
  elsewhere.

## Cost and trade-offs

- Every check is a column, and columns cost screen space and attention. Build the four that
  match the failures your data actually has, not a wall of them.
- `COUNTIF` over a large column is O(n) per row, so on a hundred thousand rows a
  self-referential `COUNTIF` is genuinely slow.
- A check column mixing text and numbers cannot be summed. That is usually correct and
  occasionally annoying.
- Data validation is a convention, not a security control. It is trivially defeated by
  pasting, which bypasses it entirely and silently.
- Checks tell you a file is *internally* consistent. They cannot tell you it is *true*. A
  file can pass every check and describe the wrong quarter.

## Recognize it on sight

- A column of `#N/A` beside an otherwise clean lookup: two systems disagree about what
  exists. Read it before you touch it.
- `IFERROR(...,"")` wrapped around a lookup: assume information has been destroyed, and go
  and find out what.
- A column of `1`s with a few `2`s: somebody built a duplicate check. Look at the `2`s.
- Small dropdown arrows appearing only when a cell is selected: data validation.
- Words in capitals inside a data column - `CHECK ID`, `MISSING`, `NOT RECORDED` -
  are somebody deliberately surfacing a gap. Treat them as findings, not as untidiness.

## Adjacent question

*"Your validation finds sixty rows carrying a customer code that does not exist in the
customer master. What do you do?"*

Not fix them. You do not own that data and a plausible guess is fabrication. You quantify
it - how many rows, what value they carry, whether they cluster in one period or one source
system - and you send it to the data owner with the evidence attached. The analyst's job is
to make the problem precise and hand it to whoever can decide. Silently mapping an unknown
code to the nearest-looking record is how an incident starts.

## Say it in an interview

"Excel's five errors each tell you exactly what went wrong. Four of them mean the formula
is wrong. `#N/A` usually means the formula is right and the data is missing, which makes it
a finding rather than a defect."

"I try not to wrap lookups in `IFERROR` with a blank message. It makes the file look clean
and destroys the information. If a lookup can legitimately miss, I give it a fallback that
says so in words, so whoever opens the file knows what to do next."

"Before I report on a file I build four checks on it: duplicate keys with `COUNTIF`, key
shape with `LEN`, blanks with `ISBLANK` on the row and `COUNTA` against `COUNT` on the
column, and referential integrity with a lookup against the master. It takes about ten minutes and it has never once found nothing."

"Where I control the input, a data validation list on the category fields is the cheapest
control there is. It costs ten seconds and prevents the class of error that is invisible
afterwards."

## Check yourself

1. Which of the five errors is most likely to be information rather than a mistake, and
   why?
2. What is wrong with `=IFERROR(XLOOKUP(...),"")`?
3. You need a lookup to say `CHECK ID` when it finds nothing. What is the better way, and
   why is it better than `IFERROR`?
4. What does `COUNTIF` on a column, comparing each row against its own id, tell you?
5. You add a validation dropdown to a column that already contains a typo. What happens to
   the typo?
6. `ISBLANK` returns FALSE on a cell that looks empty. Give two reasons.

<details>
<summary>Answers</summary>

1. `#N/A`. The formula ran correctly and reported that the value is not there. On a lookup
   between two systems, that is a genuine discrepancy worth escalating.
2. It replaces every possible failure with a blank, so a real break becomes invisible and
   the file looks clean. Whoever reads it cannot tell the difference between "no problem"
   and "problem hidden".
3. `XLOOKUP`'s fourth argument. It catches only the not-found case, so genuine errors such
   as a misspelled table name still surface instead of being swallowed.
4. How many times each id appears. Anything other than 1 is a duplicate key, which is worse
   than a duplicate row because downstream reports will merge two different things.
5. Nothing. Validation only checks new input. Existing values are untouched.
6. It contains a space, or it contains a formula returning an empty string. Both look empty
   and neither is.

</details>

## Resources

- Detect errors in formulas: https://support.microsoft.com/en-us/office/detect-errors-in-formulas-3a8acca5-1d61-4702-80e0-99a36a2822c1
- IFERROR function: https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611
- Apply data validation to cells: https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249
- COUNTIF function: https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34
