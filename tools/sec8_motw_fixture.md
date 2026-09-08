<!-- SEC.8 live-test fixture recipe. Companion to tools/sec13_word_fixture.md. -->

# SEC.8 fixture: a workbook that Windows thinks came from the internet

Frazaro's provenance gate reads the `Zone.Identifier` alternate data stream —
the Mark-of-the-Web. To test it you need a workbook that carries one. You do
**not** need to actually download or email anything: the mark is a plain
alternate data stream and PowerShell can write it directly, which is exactly
what a browser or Outlook does when it saves a file.

Everything here is ordinary PowerShell. Nothing launches Office; you open
Excel yourself.

## Two things about the order of operations

Both of these are easy to get wrong in a way that makes a test *look* like it
passed while proving nothing. They shape the whole recipe, so they come first.

**1. The mark lives on the file, not in the workbook — so mark it after the
workbook exists, not before.** The `Zone.Identifier` stream is attached to the
file on disk, so the sequence is: build the workbook → save → **close Excel**
→ mark the file → open it → run.

*Measured on this machine, 2026-09-08:* the mark **survives an Excel save and
reopen** — a marked fixture that was saved from Excel still came back with the
yellow PROTECTED VIEW banner ("files from the Internet can contain viruses").
So editing and re-saving a fixture does not silently disarm it, and you do not
have to rebuild one from scratch after a change. This was worth measuring
rather than assuming: a save that dropped the stream would leave you testing
an unmarked file, and an unmarked file is *allowed* by design — the test would
pass for the wrong reason and look like a success. If you ever get a result
you doubt, §*Check the mark* settles it in one line.

**2. Interpret runs the WHOLE program, and the first refusal stops the rest.**
There is no way to run one sentence in isolation. Frazaro reads every cell in
column B from row 1 down and interprets them as one program
(`ProgramText` → `VlaInterpret`), and a `SEC.8` refusal aborts the run at
Frazaro's own error handler — every sentence below it is never reached. Two
consequences the fixture is built around:

  * An *allowed* sentence placed **above** a refusing one still runs. That is
    useful: one file can demonstrate "ordinary work is unaffected" and "the
    external effect is refused" in a single Interpret.
  * A second thing you want to see refused needs its **own file**, because
    the first refusal would hide it.

Hence two fixture workbooks, not one.

## Build the workbooks

What they must be is specific: a **`.xlsx`**, with **no VBA of its own**,
carrying the program in cells. That combination is the whole point of `SEC.8`.
An `.xlsm` would be a different test — Office's own macro block would apply to
it, which is exactly the case this item is *not* about.

For each of the two, in Excel:

1. **New blank workbook.**
2. Rename `Sheet1` to **`Frazaro`**. (Frazaro's own *New Program* command
   builds this sheet for you with formatting and a hidden column A; either
   way, all that matters is the sheet name and column B.)
3. Type the sentences into **column B, starting at row 1**, one per row.
   Column B is the program; column C is where results are written back.
   Every sentence below is copied from the grammar's own `test-success`
   examples in `scripts/polyglotta/english.vla`, so all are known to parse.
4. **Save As** → *Save as type* → **Excel Workbook (`.xlsx`)** → save to your
   Desktop under the name given. Excel will not warn; there are no macros to
   lose.

**Workbook A — `sec8_egress.xlsx`** (covers live tests 2, 3 and 4):

| Cell | Sentence |
|------|----------|
| `B1` | `Put 1 in cell A2.` |
| `B2` | `Email "boss@co.com" with subject "Report" and message "Attached.".` |

`B1` is above `B2` deliberately. On a marked file you should see A2 become
`1` **and** the run stop with a refusal naming *"email something out of
Excel"* — ordinary work unaffected, external effect refused, one run.

Note the doubled full stop in `B2`: the message string ends in a period, and
the sentence needs its own.

**Workbook B — `sec8_protect.xlsx`** (covers live tests 5 and 6):

| Cell | Sentence |
|------|----------|
| `B1` | `Protect this sheet with password "abc".` |

Alone in its own file precisely because workbook A's refusal would otherwise
stop the run before reaching it.

This workbook is used **twice**, marked and unmarked, and that pair is the
most informative test in the set:

  * marked (live test 6) — the run is refused, naming *"protect a sheet with
    a password"*;
  * unmarked, via the `sec8_protect_local.xlsx` copy (live test 5) — the run
    **succeeds and the sheet really is protected**. Excel then says "The cell
    or chart you're trying to change is on a protected sheet" if you try to
    type in it, and the Review tab reads *Unprotect Sheet*. Unprotect with
    password `abc` when you are done.

Same sentence, same workbook contents, one bit of provenance different.

*Why this sentence and not the email one for test 5:* the email verb has no
observable success on a machine without Outlook — it fails with "Could not
start Outlook to create the email", so the only evidence the gate ALLOWED it
is the absence of the SEC.8 refusal plus the presence of an unrelated error.
That is a test whose pass condition a person has to reason about, and an
earlier draft of this file asked exactly that; the owner rightly got stuck on
it. `protect` leaves a state you can see. Every other test in this set shows
Frazaro saying *no* — this is the only one that shows a GATED verb running to
completion, which is what distinguishes a working gate from one stuck shut.

5. **Close Excel entirely** once both are saved.

## Mark them

```powershell
$dir = "$env:USERPROFILE\Desktop"
$marked = @("$dir\sec8_egress.xlsx", "$dir\sec8_protect.xlsx")

# Workbook A also gets an UNMARKED twin for live test 5.
#
# The Unblock-File is NOT redundant, and this block is written to be safe
# to re-run. Copy-Item PRESERVES alternate data streams on NTFS, so on any
# run after the first - when sec8_egress.xlsx is still marked from last
# time - the copy arrives already carrying ZoneId=3. An earlier version of
# this recipe just ordered the copy before the marking and relied on that;
# it was correct exactly once. Unblocking the twin explicitly is correct
# whatever state the source is in.
#
# This is the failure worth engineering against: a silently MARKED twin
# makes test 5 show a refusal, which reads as a false positive in the gate
# rather than as a broken fixture. Owner-caught live, 2026-09-08, by the
# verification loop below.
Copy-Item "$dir\sec8_protect.xlsx" "$dir\sec8_protect_local.xlsx" -Force
Unblock-File -LiteralPath "$dir\sec8_protect_local.xlsx"

foreach ($f in $marked) {
    Set-Content -LiteralPath $f -Stream 'Zone.Identifier' -Value @"
[ZoneTransfer]
ZoneId=3
ReferrerUrl=https://example.invalid/
HostUrl=https://example.invalid/$(Split-Path $f -Leaf)
"@
}

# Confirm all three are in the state you expect before running anything.
foreach ($f in @($marked + "$dir\sec8_protect_local.xlsx")) {
    $has = (Get-Item -LiteralPath $f -Stream * ).Stream -contains 'Zone.Identifier'
    "{0,-22} marked: {1}" -f (Split-Path $f -Leaf), $has
}
```

Expected: `sec8_egress.xlsx` and `sec8_protect.xlsx` marked `True`,
`sec8_protect_local.xlsx` marked `False`.

**Read that output before running anything in Excel, every time.** It is the
only thing standing between a mis-built fixture and a wrong conclusion about
the gate, and the two failure directions are not symmetric: a fixture that
should be marked and isn't makes a *working* gate look permissive, and a
fixture that shouldn't be marked and is makes a *working* gate look broken.
Neither shows up as an error — both just quietly produce the wrong result in
Excel a minute later.

`ZoneId=3` is the internet zone — what a browser download or a mailed
attachment gets. `ZoneId=4` is the restricted zone and is refused the same
way; `ZoneId=0`/`1`/`2` (local machine / intranet / trusted sites) are
allowed.

## Run them

Open one workbook. Excel will show the yellow **PROTECTED VIEW** banner —
"files from the Internet can contain viruses" — which is itself confirmation
the mark took. Click **Enable Editing**, run Frazaro's **Interpret** command
once, and note what happened. Repeat per workbook.

Seeing that banner is a precondition, not a nuisance: no banner means no mark,
and a run without it proves nothing. If it does not appear, check the mark
(below) before going further.

> **Live test 5 protects a sheet for real.** Running workbook B's sentence on
> an *unmarked* file genuinely password-protects the sheet with `abc`. Since
> you close without saving, the protection dies with the file — but if you do
> save, unprotect via *Review → Unprotect Sheet*, password `abc`.

## Check the mark

```powershell
Get-Item -LiteralPath $p -Stream * | Select-Object Stream, Length
```

`Zone.Identifier` in that list means the file is marked. To read it back:

```powershell
Get-Content -LiteralPath $p -Stream 'Zone.Identifier'
```

## Remove the mark (this is the "grant" path)

```powershell
Unblock-File -LiteralPath "$dir\sec8_egress.xlsx"
```

This is the command-line equivalent of right-click → Properties → tick
*Unblock*, and it is what live test 4 uses. It is the ONLY way to grant a
marked workbook external effect — SEC.8 deliberately keeps no trust store of
its own, so there is no Frazaro setting, dialog, or checkbox that does this.
Use the Explorer route at least once when testing, since that is what the
refusal message tells a person to do and it is worth confirming the
instruction is accurate.

## The one that decides the item

**Does Excel strip the mark when you click "Enable Editing"?** If it does, the
gate has nothing to read by the time Frazaro runs and the whole item is inert
— so run this before anything else, on a fresh marked copy:

```powershell
$p = "$dir\sec8_egress.xlsx"

# 1. confirm the stream is present
Get-Item -LiteralPath $p -Stream * | Select-Object Stream, Length

# 2. open $p in Excel, click "Enable Editing", LEAVE IT OPEN, then:
Get-Item -LiteralPath $p -Stream * | Select-Object Stream, Length

# 3. close WITHOUT saving, then once more:
Get-Item -LiteralPath $p -Stream * | Select-Object Stream, Length
```

Step 3 deliberately does not save. The save question is already answered —
the mark survives one (see §*Two things about the order of operations*) — and
mixing the two would make a failure ambiguous. This step asks one thing only:
does *Enable Editing* strip the mark?

**Step 2 is the whole test.** Step 1 only proves the fixture was built right,
and the Protected View banner already tells you that much. If
`Zone.Identifier` is still listed at step 2 — with the workbook open and
editing enabled, which is exactly the state Frazaro runs in — the gate reads a
live signal and every other live test is meaningful. If it disappears there,
stop and say so: the design, the policy, the memo, the guards and the pin all
stay, and only `ReadZoneStream`'s *source* has to change (Protected View
state, or the Trusted Documents registry).

## Cleanup

```powershell
Remove-Item -LiteralPath "$dir\sec8_egress.xlsx", "$dir\sec8_protect.xlsx", "$dir\sec8_protect_local.xlsx" -Force
```

## A note on where you put the fixture

Keep it on a local fixed drive (`C:`). A UNC share or a folder that syncs over
WebDAV cannot reliably carry an alternate data stream at all, and SEC.8 treats
an unreadable mark in those locations as unknown provenance and refuses —
correct behaviour, but it would make a fixture there prove the wrong thing. A
OneDrive folder that syncs to local disk is an ordinary local path and is fine.
