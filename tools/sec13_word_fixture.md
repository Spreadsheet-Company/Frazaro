# SEC.13 live-test fixture — a Word document that tries to run a macro

`SEC.13` closed the hole where `ReadWordFile` (`src/VLA_IDE.bas`) called
`Documents.Open` with no `Application.AutomationSecurity`, so Word's
automation default (`msoAutomationSecurityLow`) ran an opened document's
auto-macro the instant *Import Program File…* touched the file.

`tools/check_word_automation_security.ps1` pins that the guard is still
*present*. It cannot prove the guard *works* — only Word can, and the pure
suite does no COM. So the proof is one live test, against a document that
genuinely carries a macro.

**Why this file exists instead of a committed `.docm`:** a macro-bearing
binary is not a thing to commit to a public repository, and a script cannot
build one either — injecting a VBA project by automation needs *Trust access
to the VBA project object model* switched on in Word, which is a security
setting we are not going to ask anyone to weaken in order to run a security
test. So the fixture is a recipe, and this file is the recipe.

Build it once; keep it outside the repo (Desktop is fine).

## Building `sec13-fixture.docm`

1. Open Word → **Blank document**.
2. Type exactly these three lines into the body:

   ```text
   Work on sheet Output.
   Put "SEC13 import worked" into cell A1.
   Make cell A1 bold.
   ```

3. Press **Alt+F11** to open Word's VBA editor.
4. In the Project pane, double-click **ThisDocument** and paste:

   ```vba
   Private Sub Document_Open()
       MsgBox "SEC.13 FIXTURE: Document_Open RAN", vbExclamation
   End Sub
   ```

5. **Insert → Module**, and paste into the new module:

   ```vba
   Sub AutoOpen()
       MsgBox "SEC.13 FIXTURE: AutoOpen RAN", vbExclamation
   End Sub
   ```

   Both hooks, not one, on purpose: `AutoOpen` (a classic auto-macro) and
   `Document_Open` (the document's own event) are different mechanisms with
   different histories under automation. Carrying only one risks a false
   pass, where the guard is never really tested because the hook chosen was
   the one that would not have fired anyway.

6. **Alt+Q** to return to Word. **File → Save As**, set *Save as type* to
   **Word Macro-Enabled Document (\*.docm)**, name it `sec13-fixture.docm`.
7. Close Word completely. If Word asks about macros when you later reopen
   the file *by hand*, that is Word's own Trust Center prompt and is
   expected — it is not the path under test. The path under test is Frazaro
   opening it through automation, where no prompt is shown at all and the
   macro simply runs or does not.

## The second fixture: a document Word refuses to open

Two of the tests below need a file that *exists* but that Word cannot open.
It cannot be a missing file: `GetOpenFilename` will not return a nonexistent
path, and `EnglishIdeReload` pre-checks with `Dir$` and raises
`ide-program-file-moved` before `ImportFromPath` is reached — so
`Documents.Open` cannot be made to fail with a missing file at all.

Truncating a valid `.docm` past its ZIP central directory (which lives at the
*end* of the file) is the reliable way. It keeps the `PK` header, so Word
cannot sniff the file as plain text and "succeed" with garbage:

```powershell
$src = "$env:USERPROFILE\Desktop\sec13-fixture.docm"
$dst = "$env:USERPROFILE\Desktop\sec13-corrupt.docm"
$b = [IO.File]::ReadAllBytes($src)
[IO.File]::WriteAllBytes($dst, $b[0..([int]($b.Length/2))])
```

## The live tests

Re-run these whenever `ReadWordFile` or its callers are touched. None is
reachable without launching Word, so none is automatable; the static pin
(`tools/check_word_automation_security.ps1`) proves only that the guard is
*present*.

Run test 3 before test 4b: in 3 the Word instance is visible, so if Word
does put up a dialog it is clickable.

1. **Macro suppressed, text still imports, on an instance Frazaro started.**
   Close Word entirely. *Import Program File…* → `sec13-fixture.docm`.
   **Expect:** no message box of any kind; the three program lines land in
   the sheet and Check runs clean. **Fail =** either fixture MsgBox appears.

2. **Plain Word documents unaffected.** Save a normal `.docx` holding the
   same three lines; import it. **Expect:** an ordinary successful import.

3. **Macro suppressed on your own attached Word, and its security setting
   comes back.** Open Word by hand. Alt+F11 → Ctrl+G →
   `?Application.AutomationSecurity` → Enter; **write the number down**.
   Leave Word open; import `sec13-fixture.docm`. **Expect:** no message box,
   text imports. Re-run `?Application.AutomationSecurity`. **Expect the same
   number.** **Fail =** a MsgBox appeared, or the value is now `3`.

4. **The error path restores it too.** This is the test that exercises the
   restore line in `cleanup:`, and it only works with Word **attached** —
   that line is guarded by `Not createdNew`, so with Word closed it is a
   deliberate no-op and the test would pass without proving anything.

   - **4a (attached).** With Word open and its `AutomationSecurity` noted,
     import `sec13-corrupt.docm`. **Expect:** a refusal box beginning
     *"Could not read the Word document (is Word installed?):"* — the text
     after the colon is Word's own wording and varies. Then re-read
     `?Application.AutomationSecurity`: **expect the noted number**, not `3`.
     Your Word must still be open — Frazaro must not quit what it did not
     start.
   - **4b (created).** Close Word entirely, import `sec13-corrupt.docm`.
     **Expect:** the same refusal, and **no `WINWORD.EXE`** left in Task
     Manager afterwards.

5. **Your Word's alert level is never touched.** Open Word by hand.
   Alt+F11 → Ctrl+G → `?Application.DisplayAlerts`; **note the number.**
   `WdAlertLevel` is `wdAlertsNone = 0`, `wdAlertsAll = -1`,
   `wdAlertsMessageBox = -2`; all three are legitimate and the ambient value
   varies between machines (`-2` observed on the owner's, 2026-09-08). Leave
   Word open, import `sec13-fixture.docm`, re-read it. **Expect the same
   number you noted** — the assertion is *unchanged*, not any particular
   value. **Fail =** it now reads `0`, meaning the `DisplayAlerts` suppression
   leaked out of the created-instance branch into an attached instance.
