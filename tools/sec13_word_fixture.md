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

## What a correct result looks like

Frazaro imports the three lines of text into the sheet, and **neither
message box ever appears**. Before the fix, one of them would.

The numbered live-test steps that use this fixture are in the handoff for
`SEC.13` (and in `docs/RELEASES.md`'s `0.5.3` notes).
