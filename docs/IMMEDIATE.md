# The Immediate window

*For someone who knows VBA and wants to look underneath. Nothing here is
needed to use Frazaro — the ribbon does all of it — and no message Frazaro
shows will ever send a user here: `tools/check_no_vba_advice.ps1` refuses
any message that recommends VBA (the owner's rule, 2026-09-26).*

## Where to type them

In the development workbook, with `src/` imported as
[TESTING.md](TESTING.md) describes, type a command into the VBA editor's
Immediate window (**Ctrl+G**). A line starting with `?` prints what the
command returns.

An installed add-in's code is locked for viewing, but its public procedures
still run, through `Application.Run` and the add-in's file name:

```
?Application.Run("Frazaro_English.xlam!EnglishListPhrases")
```

The grammar these read is whatever the last button loaded, so press
**Validate Instructions** once first.

## What Frazaro understands — in the add-in too

| Command | What it does |
|---|---|
| `?EnglishListPhrases` | Every sentence form loaded right now, grouped by first word: the list **What can I say?** shows. |
| `EnglishExplain "Fit all columns."` | Prints how one sentence is read: its words, then the rule and the phrasebook that matched it. |
| `?EnglishToVla("Fit all columns.")` | The VLA a sentence translates to. |
| `?EnglishLoadedSourcesReport` | Which phrasebook files are loaded, and how many rules each one supplies. |
| `?EnglishVocabStats` | Rules, macros and phrasebook tests, counted at the last load. |
| `EnglishResetGrammar` | Empties the grammar. |
| `?EnglishLoadVocabulary("C:\...\english.vla")` | Loads one phrasebook on top of the grammar and returns its rule count. Loading adds, so the same file twice without a reset is refused rather than defining every word twice. Every ribbon button resets and reloads for you. |
| `EnglishIdeClearLog` | Empties the hidden log of sentences Check could not read, the log **Copy Diagnostic Report** collects. |

## In the development workbook only

These live in modules the add-in does not ship (`VLA_Tests*`, `VLA_DevRig`,
`VLA_Build`).

| Command | What it does |
|---|---|
| `VlaDevReload` | Imports every module from `src/`; compile afterwards (**Debug > Compile VBAProject**). |
| `?VlaSelfTest` | The suite that needs no workbook: PASS or FAIL per test, and the summary line last. |
| `?VlaSelfTestHost` | The tests that need a live workbook. |
| `?VerifyReports` | Checks the regression corpus, `scripts/instructions.txt`: the report its last compiled Run left, then an interpreted run of its own ([TESTING.md](TESTING.md), pass 4). |
| `VlaTryValue "(+ 2 3)"` | Compiles and runs one VLA expression and prints its value. It resets the grammar, so press **Validate Instructions** before the next Check. |
| `VlaDiagnostics` | Prints the version of every module and the state around them. |
| `VlaBuildAddin` | Builds the add-ins; [DEPLOY.md](DEPLOY.md) has the release steps around it. |
