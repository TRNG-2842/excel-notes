# Lookups and reconciliation

## What it is and why it exists

A lookup goes to another table, matches on a shared key, and brings a value back. It is the
most common thing anybody does in a spreadsheet, because business data is stored as codes
and read by humans who want names.

Reconciliation is what you do with lookups once there are two systems that are supposed to
agree. It is proving that two independent records of the same events match, and explaining
every place they do not.

The part people get wrong: **reconciling is not making the difference zero.** The difference
is almost never zero and is not supposed to be. Reconciling means *explaining* the
difference, item by item, with a reason and an owner against each. A reconciliation is
finished when the listed breaks add up to the difference you started with, not when the
difference disappears.

## How it behaves

### `XLOOKUP`

```
=XLOOKUP(what, where_to_look, what_to_bring_back, [if_not_found])
```

Read it as a sentence. Find *this*, in *there*, give me *that*, and if you cannot, say
*this instead*.

- `where_to_look` and `what_to_bring_back` are single columns of the same height.
- The return column can be to the left of the search column. This matters more than it
  sounds.
- The fourth argument is optional and you should almost always supply it. Without it a miss
  returns `#N/A`.

```
=XLOOKUP(B2, tblRef[customer_id], tblRef[customer_name], "CHECK CUSTOMER ID")
```

| `B` | Result |
|---|---|
| `C-0417` | `Halden Foods` |
| `C-0999` | `CHECK CUSTOMER ID` |

Write the fallback as an instruction to a person. `"CHECK CUSTOMER ID"` tells somebody what
to do. `""` tells them nothing and hides the fact that anything happened.

Prefer the fourth argument to wrapping the whole thing in `IFERROR`. `IFERROR` catches every
possible failure including a misspelled table name, so a structural mistake gets disguised as
a missing record. The fourth argument catches only the not-found case.

### `VLOOKUP`, which you will read but should not write

`VLOOKUP` identifies the return column by **counting** columns rather than naming one, so
inserting a column anywhere in the range silently changes what it returns. It also cannot
look to the left of its key. It is in every legacy workbook you will inherit; read it,
understand it, and write `XLOOKUP` instead. `INDEX`/`MATCH` is the older way of doing what
`XLOOKUP` now does in one function, and is still perfectly good.

### The five shapes a break takes

Given two systems that share a reference, four shapes are visible to a lookup:

| Shape | Means |
|---|---|
| In A, missing from B | it happened and nobody recorded it |
| In B, missing from A | it was recorded and never happened; often a cancellation never reversed |
| Both, amounts differ | somebody typed it wrong, or a fee was deducted |
| Both, dates differ | usually a cut-off or overnight posting; often not an error at all, and still has to be explained |

And a fifth that hides from all four: **the same reference recorded twice on one side**.
Both entries are individually valid. Every matching test passes, because the lookup finds
the first one and it agrees perfectly. It is invisible to matching and can only be found by
**counting**:

```
=COUNTIF(tblEntries[reference], D2)
```

Anything greater than 1 is a duplicate posting. Add this to every reconciliation you build.

### Run it in both directions

A lookup from A to B can only find things A knows about. Entries that exist only in B are
invisible from that side, and they are usually the ones that flatter you. Two lookups, one
each way. Every time.

### Tolerance

Never test money for equality with zero.

```
=IF(ABS(delta) >= 0.01, "AMOUNT DIFFERS", "OK")
```

Excel stores decimals in binary and cannot represent values like 0.1 exactly, so two amounts
that genuinely match can subtract to 0.0000000001. Compared against zero, that row is a
break, and somebody spends an afternoon on a difference that does not exist. Every
reconciliation has a tolerance; a cent is a normal one.

`ABS` also makes the test catch differences in both directions, which a plain `>` does not.

| `delta` | Result |
|---|---|
| `0.004` | `OK` |
| `12.50` | `AMOUNT DIFFERS` |
| `-0.03` | `AMOUNT DIFFERS` |

### A status column that reads like English

One row per bank entry. `B` is the bank amount, `C` the bank date, `D` the ledger amount
brought back by a lookup with `"NOT IN LEDGER"` as its fallback, `E` the ledger date, and
`F` is `=B2-D2`.

```
=IF(D2="NOT IN LEDGER","MISSING IN LEDGER",
   IF(ABS(F2)>=0.01,"AMOUNT DIFFERS",
      IF(E2<>C2,"DATE DIFFERS","OK")))
```

| `B` | `C` | `D` | `E` | `F` | Result |
|---|---|---|---|---|---|
| 1250.00 | 03-Mar | 1250.00 | 03-Mar | 0.00 | `OK` |
| 1250.00 | 03-Mar | `NOT IN LEDGER` | | `#VALUE!` | `MISSING IN LEDGER` |
| 1250.00 | 03-Mar | 1237.50 | 03-Mar | 12.50 | `AMOUNT DIFFERS` |
| 1250.00 | 03-Mar | 1250.00 | 04-Mar | 0.00 | `DATE DIFFERS` |

The `#VALUE!` in `F` on the missing row is fine: the status test reads `D` first and never
reaches `F`. If it bothers you, make `F` conditional on `D` being a number.

The order is a decision, not an accident. Ask the most serious question first: is it missing
entirely. Only if it is present do the amounts matter. Only if the amounts agree do the
dates matter. Put the date test first and a missing item gets labelled a date difference,
which is worse than useless.

### The deliverable is the tracker, not the formulas

One row per break:

`issue_no` | `reference` | `found_in` | `issue_type` | `amount_a` | `amount_b` | `owner` |
`status`

The last two are the ones that matter. A break with no owner is not being worked on. A list
with no status column is a list, not a control. Make `status` a data validation dropdown so
nobody invents a fourth state.

Everything else - the lookups, the deltas, the conditional formatting - is your working.
Nobody outside the team will ever look at it.

### Making the breaks visible

Conditional formatting colours a cell according to a rule, and the rule re-evaluates every
time the data changes. It is how a status column becomes readable from across the room.

Select the status column, Home > Conditional Formatting > Highlight Cells Rules > Text that
Contains, type `DIFFERS`, pick a fill. Add a second rule for `MISSING`. Every row that is
not `OK` now stands out, and stays standing out as the lookups update.

To colour the **whole row** rather than one cell: select the data, Conditional Formatting >
New Rule > Use a formula, and enter a formula anchored to the status column, such as

```
=$H2<>"OK"
```

The `$` pins the column so every cell in the row reads the same status; the `2` is
relative so each row reads its own. Inside a Table the rule grows with the data.

Three cautions. Rules pile up silently - Manage Rules shows what is there, and a sheet that
has been edited for months usually carries dead ones. Many rules over many rows are the
usual reason a workbook is slow. And a colour is not a filter: if you need the breaks as a
list, filter on the status column or build the tracker. The colour is for reading, never
for counting.

### Proving the list is complete

This is the step that separates a reconciliation from a list of things somebody noticed.

Every break either adds to the difference between the two totals or subtracts from it. So if
the list is complete, the breaks must reconcile to the difference exactly:

```
(sum of A-only items)
  - (sum of B-only items)
  +/- (each amount difference, signed)
  - (each duplicate posting's extra copy)
  = total A - total B
```

Items that differ only in date contribute nothing, because the amounts are equal on both
sides.

If the two figures agree, the list is complete. If they do not, there is a break you have
not found, and you know that with certainty rather than hoping. Being able to state which
version you reconciled to - before or after removing duplicates - is part of the same skill.

## Cost and trade-offs

- **Exact matching only.** Everything here assumes a shared, clean reference. When
  references are missing you are into fuzzy matching on amount and date, which is a much
  larger problem and not a spreadsheet one.
- **Many-to-one.** One payment covering several invoices breaks the one-row-to-one-row
  assumption completely. Common in practice and beyond what a lookup can do.
- **Keys must be clean first.** `INV0002`, `INV-0002` and ` inv0002 ` are three different
  strings. Most reconciliation failures are formatting differences in the key, not genuine
  breaks. Standardise before you match.
- **`XLOOKUP` returns the first match** and says nothing about the second. This is exactly
  why the count is a separate check.
- **Multi-currency** adds an FX movement to every delta and needs a rate table and a
  decision about which date's rate applies.
- **It does not scale by hand.** Above a few thousand rows a month, this belongs in Power
  Query or a database.
- **`XLOOKUP` is Microsoft 365 only.** On older builds, `INDEX`/`MATCH`.

## Recognize it on sight

- Two blocks of data side by side with a narrow column of lookups between them.
- A column of `#N/A` beside an otherwise clean join: a real disagreement between systems,
  not a broken formula.
- `IFERROR(VLOOKUP(...),"")`: assume information has been destroyed and go and find out
  what.
- A "difference" cell at the top of a sheet with a list underneath: somebody was taught to
  reconcile properly.
- A delta column tested against `=0` rather than a tolerance: expect spurious breaks.
- Only one lookup, in one direction: half a reconciliation.

## Adjacent question

*"Your reconciliation shows a difference of zero. Are you done?"*

No, and the question is a trap. Zero can mean everything matches, or it can mean two errors
in opposite directions cancelled. A payment recorded twice and a payment missed entirely can
net to nothing while both are real problems. The totals agreeing is necessary and not
sufficient; you still run the item-level match in both directions and the duplicate count. A
net-zero difference with a dozen breaks underneath it is a completely normal outcome.

## Say it in an interview

"Reconciling is not making the difference zero. It is explaining the difference. I finish by
checking that the breaks I listed add up to the difference I started with, and if they do
not, I know there is one I have not found."

"I always run it in both directions. A lookup from the bank to the ledger only finds things
the bank knows about, and the entries that exist only in the ledger tend to be the ones that
flatter you."

"I add a count as well as a match, because a duplicate posting passes every matching test.
The lookup finds the first entry, it agrees perfectly, and the second copy is invisible.
`COUNTIF` on the reference is the only thing that catches it."

"I compare deltas against a tolerance rather than zero, because floating point means amounts
that genuinely match can subtract to a tiny non-zero remainder."

"And the deliverable is the tracker, not the formulas. One row per break with an owner and a
status, so somebody who was not involved can pick it up."

## Check yourself

1. What are the three required arguments of `XLOOKUP`, in English?
2. Why supply the fourth argument rather than wrapping the lookup in `IFERROR`?
3. Why run a reconciliation in both directions?
4. A duplicate posting matches perfectly on amount and date. How do you find it?
5. Why test `ABS(delta) >= 0.01` instead of `delta <> 0`?
6. Two totals agree exactly. Is the reconciliation complete?
7. What proves that your list of breaks is complete?

<details>
<summary>Answers</summary>

1. What am I looking for; where do I look for it; what do I bring back.
2. `IFERROR` catches every error, so a misspelled table name is disguised as a missing
   record. The fourth argument catches only the not-found case and leaves real errors
   visible.
3. A lookup from A to B can only find records A has. Anything existing only in B is
   invisible from that side.
4. By counting, not matching. `COUNTIF` on the reference column, looking for more than one.
5. Binary floating point means matching amounts can subtract to a tiny non-zero value.
   Comparing to zero produces breaks that do not exist. `ABS` also catches both directions.
6. Not necessarily. Two errors in opposite directions cancel. You still match at item level,
   both ways, and count for duplicates.
7. The listed breaks reconcile to the difference between the two totals, exactly. If they do
   not, there is one you have not found.

</details>

## Resources

- XLOOKUP function: https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
- VLOOKUP function: https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
- INDEX function: https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
- Use conditional formatting to highlight information: https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f
