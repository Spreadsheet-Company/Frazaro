Attribute VB_Name = "VLA_IDE"
Option Explicit
Public Const VLA_IDE_VERSION As String = "U.32"
' U.32: Undo puts a workbook's sheets back as they were - names, tab
' order, visibility and the active sheet. Before a Run, TakeRunSnapshot
' records every worksheet's name and visibility in tab order and the
' active sheet on the Run's own marker sheet, VLAs_<tag>_Sheets (the
' roster, a fourth snapshot kind beside the copies, the tombstones and the
' staging; staged, swapped and swept like the others). At Undo, and when a
' Run stops, PutBackLastRun restores the copies IN PLACE (the copy goes
' Before the sheet set aside, so a restored sheet no longer lands at the
' end of the tab order) and, once the copies and tombstones are done,
' applies the roster: a sheet the run hid is shown again, the roster's
' sheets go back into their relative order, a sheet the run showed is
' hidden again, and the sheet that was active is active again - each act
' decided purely (VlaIdeRosterPlan) and named in the dialog
' (VlaIdeRosterWords), and nothing when nothing differs. The snapshot's
' scan also reads the helper rows a raw program can hold (ScanHelperRows:
' rename-sheet, clear-sheet, delete-sheet, add-sheet-at, copy-sheet-named
' and the runtime calls under them), so a renamed sheet's old name is
' copied and its new name tombstoned, as for the sentence. The owner's
' calls, 2026-10-07: only a tombstoned name is removed (a bare copy
' stands); the roster stays with the copies, so a second Undo puts the
' same state back; the active sheet is the one before the Run. A copy
' comes back under its sheet's own spelling, the roster's
' (VlaIdeRosterSpelling): a sheet a sentence names by a bare name is
' scanned, and so copied, in lower case, and Undo used to rename "Sales"
' to "sales" (found 2026-10-09, preparing the close).
' LE.11: "What can I say?" writes three columns - Category, Template,
' Example - with no header rows (EnglishPhraseRows's categories), so the
' Category filter shows one word's sentences, and its cells wrap, so a
' long template or example grows its row instead of spilling past the
' table's edge.
' LX.14 (call 2): Check writes the translation's notes (EnglishLastNotes)
' beside OK, each on its own row, in yellow (VlaIdeMarkNote) - a program's
' own "To <word> of ...:" masking a phrasebook's word. Nothing is refused.
' U.29: a Run gives back Excel's settings as it found them (the audit's
' C62). Before, a Run put back only screen updating, always to True: a
' program that turned calculation off, or put words in the status bar, and
' then stopped left every open workbook in manual, or the words showing,
' for the rest of the Excel session, and Undo Last Run, which puts back
' sheets, reached neither. VlaIdeRecordExcel records calculation, the
' status bar, alerts, events, screen updating and whether cut or copy mode
' was on, before the Run's first change; VlaIdeGiveBackExcel puts each back
' on every exit of RunProgram and InterpretProgram, finished or stopped,
' before any stop's put-back and message. The program sheet is protected
' only when it had no protection, and unprotected only when the Run made
' it. Public, with VlaIdeStatusBarBack (pure), for the suites.
' U.28: the workspace sheet is made after the snapshot, not before, in
' both Run procedures, so a Run that creates it can have it removed: a
' stop or Undo Last Run takes it away, as every other sheet a run creates,
' and says so, where it used to put back an empty sheet the person never
' had and call it "as they were before the run".
' SOP.6: the <Frazaro> tag. Load and Reload read only the part of a
' document between <Frazaro> and </Frazaro> when it has one, with values
' from the opening tag filled into the section's sentences
' (VlaIdeFrazaroSections, beside NormalizeProgramText); and a document
' with no tag that fails its first Check says how to add one, on the red
' row itself (DoCheck's failHint).
' U23.0: three follow-ups to U.19. U.21 - Undo finds the copy it restores
' by its new name, not Worksheets(snap.Index + 1) (Index counts chart
' sheets, Worksheets does not), and a restore that fails part-way puts
' that sheet back as the Run left it (RollBackRestore). U.22 - program
' tags compare without case everywhere (VlaIdeTagsMatch), as Excel sheet
' names do: the Add Program and Run guards, the snapshot sweep, Undo's
' match. U.23 - a Run's new copies are staged as VLAn_<tag>_ and replace
' the previous Run's only once all exist, so a Run refused on the way
' keeps the Undo that was there; the sweep's decision is the pure
' VlaIdeSweepsSnapshot.
' U19.0: a failed Undo snapshot leaves nothing behind (U.19).
' TakeRunSnapshot records every sheet it makes the moment it exists and,
' if any step fails, removes them all and refuses through
' ide-undo-snapshot-failed; RunProgram and InterpretProgram now STOP
' there instead of running on without Undo. The copy is found by its new
' name (VlaIdeAddedSheetNames), not by tab position, and a scanned name
' no sheet can have is not a target (VlaIdeCanBeSheetName). The module's
' last raw Err.Raise - the LX2.0 note's 16th, TakeRunSnapshot's cleanup
' re-raise - is gone with it.
' TER2.0: ClearMarks clears column C down to its own last mark as well as
' the program's last row, so a mark on a sentence whose text was deleted
' - the last one, or the whole program - no longer survives the next
' Check (TER-2). Found on TER-2's live pass: the empty-program guard on
' Check, Compile, Interpret, Export and Show VBA could never fire; they
' use ProgramIsBlank now. And, owner request: a newly created workspace
' tab opens with B1 active (SelectFirstSentenceCell).
' U20.0: Lint VLA can no longer break a file it rewrites (U.20). Both
' flows decide through VLA_Lint.VlaLintFileBytes before anything is
' written - refusing a generated artifact, invalid UTF-8, and any
' reformatting that would move a non-whitespace byte, by name - and
' write only through WriteLintedFile (a verified sibling copy, then a
' rename). The folder flow shows its whole plan before the first
' overwrite. WriteTextFileVlaIde is UTF-8 now, which also fixes Export
' Expanded Phrasebook and Phrasebook Test Coverage for non-ASCII text.
' EDITIONMANIFEST.1: IdeVocabPath's four candidate filenames were
' hardcoded to "english.vla" for every edition - a real gap this
' session's own scoping pass found (a stray external english.vla could
' silently outrank a Spanish edition's embedded default). Now derived
' from IdeVocabFileName (new), which reads a small embedded marker
' VLA_Build.bas now writes per edition (VLAe_Name), falling back to
' "english.vla" when absent - the dev workbook and any pre-existing
' build both keep behaving exactly as before. IdeLoadVocab's embedded-
' fallback branch now loops over a CHAIN of sheets (VLAe_Source,
' VLAe_Source2, ...) instead of assuming exactly one - see its own
' header for why (a non-English edition's own overlay phrasebook).
' LX2.0: this module's 15 of 16 raw Err.Raise refusal sites now route
' through VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2. The 16th
' (TakeRunSnapshot's cleanup, ~line 2343 - NOT ReadWordFile's, which this
' note misnamed until SEC.13's pass corrected it) is a bare re-raise of an
' already-caught error, not an origination of new English text, and stays
' raw by design. Rendered text and Err.Number/Err.Source are unchanged.
' DI3.2: DI.3c - the standalone path has no installer AppVersion for a
' user to compare, so the running add-in now shows VLA.VLA_RELEASE_VERSION
' itself: the Known Sentences sheet's own header (EnglishIdeShowPhrases,
' the closest thing this product has to an About surface already) now
' reads "Frazaro <version> understands:" - owner-revised from a separate
' "You can say:" / "Frazaro <version>" two-line header, one line reading
' cleaner - and "Copy Feedback"'s header gains its own "Frazaro <version>"
' line alongside the existing engine/vocabulary version line, which named
' the wrong thing for this purpose (VLA_ENGLISH_VERSION is a per-module
' provenance tag, not a release number - kept as-is, since it's still
' genuinely useful for a maintainer diagnosing a phrasebook issue; the
' release version is added beside it, not instead of it).
' LX3.0: every identity-deciding LCase$ site now folds through
' VLA_Identity.Fold (invariant ASCII, never locale-dependent) - R6/SD-8.
' IO6.0: generated-module namespacing. OUT_MODULE was "EN_Sheet", now
' "Frazaro_EN_Sheet"; the named-program path was "EN_" & tag, now
' "Frazaro_EN_" & tag. Prevents, not just catches (IO.1), a collision
' with a foreign module of the same name. No external pilot workbook
' exists yet with the old name in its history, so this is a clean
' cutover: a workbook run under the old name gets a fresh module on its
' next Run, and any old EN_Sheet/EN_<tag> module is left behind,
' orphaned. Application.Run targets for the injected runtime (EN_Runtime
' -> Frazaro_EN_Runtime) updated to match VLA_Runtime.bas's RT_MODULE.

' =====================================================================
'  VLA_IDE - the worksheet workspace for English programs.
'
'  One-time setup (Immediate window):   VlaIdeSetup
'  This prepares a sheet named "Frazaro": write the program in column
'  A, one instruction per row - blank rows end blocks, exactly like
'  blank lines in a file, and # starts a comment.
'
'  V2: a workbook holds SEVERAL programs. Each lives on its own
'  workspace sheet - the default "Frazaro" plus any number of
'  "Frazaro (<name>)" sheets (menu: Add Program...). Every command
'  binds to the program whose sheet you invoked it on (a button
'  click IS its sheet; menu commands take the active sheet), falling
'  back to the workbook's only program when there is exactly one.
'  A compiled program compiles to its own module (Frazaro_EN_Sheet for
'  the default, Frazaro_EN_<tag> for named ones); both backends
'  snapshot under the same per-program prefix (VLAu_<tag>_...), so Run
'  and Undo are scoped per program and two programs' undos cannot
'  collide. Buttons:
'
'    Check Instructions      translate + validate only; nothing runs.
'                            Every good row gets a green OK; the first
'                            bad row gets the error message, in red,
'                            next to the instruction.
'    Interpret Instructions  IN.6: checks first, then runs the checked
'                            instructions directly - no VBA is
'                            generated, no VBProject access is needed,
'                            so this works even where Compile cannot.
'    Compile Instructions    checks first; only an all-green program
'                            compiles (into module Frazaro_EN_Sheet)
'                            and runs - starting on the "Output" sheet,
'                            with the program sheet protected so a
'                            program cannot overwrite its own
'                            instructions. Needs VBProject trust.
'    Show Compiled VBA       IN.4: checks first, then writes the
'                            generated VBA (the same text Compile would
'                            inject) to a "Generated VBA" sheet, plain
'                            text, gridlines off. No VBProject write -
'                            works without trust, like Check.
'    Known Sentences         writes the vocabulary's own supported
'                            phrasings, as a real Excel Table named
'                            "Phrasebook" (LE.1), to a sheet tab of the
'                            same name for parity - the language's
'                            catalog, not this program's instructions.
'                            Category, Template, Example (LE.11).
'
'  The vocabulary comes from IdeVocabPath below - edit it to point at
'  your .vocab file(s).
' =====================================================================

Private Const IDE_SHEET As String = "Frazaro"
Private Const IDE_SHEET_LEGACY As String = "English"   ' pre-naming workspaces migrate in place
Private Const PHRASEBOOK_SHEET As String = "Phrasebook"
Private Const VBA_SHEET As String = "Generated VBA"   ' IN.4: "Show me the VBA"
' IO.6: was "EN_Sheet" - renamed to prevent, not just catch (IO.1),
' a collision with a foreign module of the same name. No external
' pilot workbook exists yet with the old name in its history, so this
' is a clean cutover, not a migration: a workbook run under the old
' name gets a fresh Frazaro_EN_Sheet module on its next Run, and the
' old EN_Sheet module (if any) is simply left behind, orphaned.
Private Const OUT_MODULE As String = "Frazaro_EN_Sheet"
Private Const OUT_SHEET As String = "Output"
Private Const UNDO_PREFIX As String = "VLAu_"
Private Const DEL_PREFIX As String = "VLAd_"       ' D1.2: tombstone markers -
                                                   ' "this sheet did not exist
                                                   ' before the run; Undo
                                                   ' removes it"
Private Const STAGE_PREFIX As String = "VLAn_"     ' U.23: a Run's new copies and
                                                   ' markers while they are made;
                                                   ' renamed VLAu_/VLAd_ only once
                                                   ' every one of them exists
Private Const ROSTER_PREFIX As String = "VLAs_"    ' U.32: the Run's roster - every
                                                   ' sheet's name, visibility and
                                                   ' the active sheet, as they were
                                                   ' before the Run; one per Run,
                                                   ' VLAs_<tag>_Sheets. The four
                                                   ' prefixes are one length by
                                                   ' design (the swap in
                                                   ' TakeRunSnapshot relies on it)
Private Const ROSTER_WORD As String = "Sheets"     ' the roster's target part
Private Const ROSTER_HEAD As String = "Frazaro.Undo:1"   ' its first line, versioned
Private Const LOG_SHEET As String = "VLA_Log"      ' deliberately not VLAu_/VLAd_:
                                                   ' the snapshot sweeper owns
                                                   ' those prefixes
Private Const FIRST_ROW As Long = 1   ' S3.3 (owner revision): row 1
                                      ' belongs to the program - real
                                      ' IDEs don't steal line 1. Sheet
                                      ' row N IS program line N now
                                      ' (file/sheet parity); the old
                                      ' guidance line lives in a cell
                                      ' comment on A1

' SOP.1: a password Word will never match, handed to every
' Documents.Open ReadWordFile makes. Word's password box is a DIALOG,
' not an alert, so DisplayAlerts does not cover it: a protected
' document would otherwise stop to ask for a password on a window
' nobody can see. With a wrong one supplied up front Word refuses
' instead, and ide-word-password says so in words.
Private Const WORD_NO_PASSWORD As String = "frazaro-no-password"

' D1: the host workbook handle, captured once per public command by
' CaptureHost (defined beside HostBook below). Declared HERE because
' VBA requires module-level declarations to precede the first
' procedure - placing this next to its machinery drew "Only comments
' may appear after End Sub..." on D1.0's maiden compile.
Private mHost As Workbook
Private mHostSheet As Object       ' V2: the sheet active at invocation
                                   ' (Object - could be a chart sheet)

' GO.6: every persisted phrasebook path, vbLf-joined, one workbook-
' scoped CustomDocumentProperty - see PersistPhrasebookPath's own
' header, further down, for the full design.
Private Const PHRASEBOOK_LIST_PROP As String = "VLA_LoadedPhrasebooks"

' SEC.9: the registry section holding this DEVICE's answer to "may
' that phrasebook load?" - one value per path, "<digest>|granted" or
' "<digest>|denied". The workbook property above is the REQUEST; this
' is the ANSWER, and only the answer is trusted. Full design in the
' SEC.9 header block further down, beside the gate itself.
Private Const PHRASEBOOK_CONSENT_SECTION As String = "SEC9PhrasebookPath"

' U.29: Excel's settings as a Run found them - VlaIdeRecordExcel fills
' one before the Run's first change, VlaIdeGiveBackExcel puts it back.
' Public so the suites and VerifyReportInterpreter can bracket a run the
' same way. The status bar is kept as text: an idle one reads back as
' the text FALSE (VlaIdeStatusBarBack).
Public Type VlaExcelSettings
    Calculation As Long
    StatusBar As String
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    CopyModeOff As Boolean
End Type

' U.32: what one put-back did, for the dialog - PutBackLastRun fills one.
' Each list is comma-separated sheet names, "" when nothing of the kind
' happened. restoring names the sheet a failed put-back was on ("" when it
' stopped before changing anything, or finished). hadRoster says whether
' the last Run left a roster at all (a Run before U.32 did not).
Public Type VlaUndoReport
    restored As String       ' sheets put back from their copies
    removed As String        ' tombstoned sheets the run had created, gone again
    restoring As String
    shown As String          ' roster: hidden by the run, shown again
    hidden As String         ' roster: shown by the run, hidden again
    moved As String          ' roster: put back in their place
    activated As String      ' roster: the sheet active before the Run, active again
    missing As String        ' roster: gone, and no copy brought them back
    rosterProblem As String  ' the roster could not be read: the words, else ""
    hadRoster As Boolean
End Type

' EDITIONMANIFEST.7 (owner-caught, live, the real root cause behind
' three straight "Bad file name or number" rounds - the first two
' fixes were real bugs too, just not THIS one): Dir$ is a local-
' filesystem call; it raises 52 outright on anything it cannot even
' parse as a local path, confirmed live via VBE break: hb.Path for a
' OneDrive-stored workbook came back as "https://d.docs.live.net/...",
' not a drive letter. IdeVocabPath's own candidate search had no
' reason to assume every host workbook lives on a local drive, and
' nothing before this wrapped Dir$ defensively - a candidate that
' merely doesn't exist and a candidate Dir$ cannot even evaluate were
' never distinguished. Both now mean the same thing: try the next
' candidate. Not URL-specific by design - any other reason Dir$ might
' choke on a path (a dead network share, a permission error, a future
' cloud-storage shape neither of us has hit yet) is handled the same
' way, on purpose, rather than growing a special case per cause.
Private Function SafeFileExists(ByVal p As String) As Boolean
    On Error Resume Next
    SafeFileExists = (Len(Dir$(p)) > 0)
    If Err.Number <> 0 Then SafeFileExists = False
    On Error GoTo 0
End Function

' EDITIONMANIFEST.1: which filename an external override must carry to
' be found at all - "english.vla" unless this edition's own build
' embedded a different one (VLA_Build.EmbedTextAsSheet's "VLAe_Name",
' read back the same way VLAe_Source/VLAp_Source already are). A real
' gap this session's own EDITION-MANIFEST scoping pass found, not
' hypothetical: before this, EVERY edition's IdeVocabPath searched for
' a file literally named "english.vla" - a stray one left beside a
' Spanish-edition workbook (a leftover from also evaluating the
' English edition, say) would silently outrank the embedded Spanish
' default with English vocabulary, in the Spanish product. Absent in
' the dev workbook (no embedded sheets exist there at all) and in any
' add-in built before this pass - both correctly fall back to the one
' name every edition understood until now, so nothing already built or
' installed changes behavior.
' EDITIONMANIFEST.4 (owner-caught, live): VlaEmbeddedText always
' appends vbCrLf after every line it reads back, including the last -
' correct for its usual multi-line-source-text callers (VLAp_Source,
' VLAe_Source), wrong here, where VLAe_Name holds exactly one bare
' filename, not parseable text. Trim$ does NOT strip Chr(13)/Chr(10) -
' it only strips spaces - so this was silently returning
' "espanol.vla" & vbCrLf, and every candidate IdeVocabPath built ended
' in two literal control characters: "...\espanol.vla<CR><LF>", the
' exact shape Dir$ raises "Bad file name or number" on. Never caught
' before now because VLAe_Name only exists inside a BUILT .xlam - the
' dev workbook has no such sheet, so this function always took the
' "english.vla" default branch in every dev-side test this session.
Private Function IdeVocabFileName() As String
    Dim nm As String
    nm = VLA.VlaEmbeddedText("VLAe_Name")
    nm = Replace(nm, vbCrLf, "")
    nm = Replace(nm, vbCr, "")
    nm = Replace(nm, vbLf, "")
    If Len(Trim$(nm)) = 0 Then nm = "english.vla"
    IdeVocabFileName = Trim$(nm)
End Function

' Where the grammar lives: looked for next to the user's workbook
' first (scripts\ subfolder, then alongside), then next to the add-in.
' EDITIONMANIFEST.5 (owner-caught, live - a real regression, not a
' close call): a scripts\polyglotta\ candidate was added HERE for
' ThisWorkbook.Path in EDITIONMANIFEST.3, meant only to help the DEV
' WORKBOOK find its own companion files. It broke a BUILT edition
' instead: VlaBuildAddin always saves an edition's own .xlam right
' beside the dev workbook ("next to this workbook"), so inside a built
' Frazaro_Espanol.xlam, ThisWorkbook.Path IS the repo root - the
' candidate found the real scripts\polyglotta\espanol.vla sitting
' there and IdeLoadVocab took it as a genuine external override,
' loading espanol.vla ALONE and skipping english.vla - which
' espanol.vla's own overlay rules call macros from but never defines
' itself. Every one of those macro calls then compiled as a literal,
' unexpanded "Call whatever(...)" - "Sub or Function not defined" at
' Compile, the first live proof this whole embedded-chain mechanism
' got, not caught by Check (test-success compares translated VLA TEXT
' structurally; it never needs a macro to actually be expandable to
' "pass"). The polyglotta candidate moved to IdeLoadVocab below, as a
' true last resort AFTER the embedded chain - it must never be able to
' outrank an edition's own embedded phrasebook again.
'
' EDITION-VOCABPATH (2026-09-07): that move left three ribbon buttons
' with nothing to find. Translate to VLA/VBA (TranslateViaRibbon), Export
' Expanded Phrasebook and Phrasebook Test Coverage each need a REAL disk
' path to point at and never go through IdeLoadVocab's embedded chain, so
' in the dev workbook - no embedded sheets, and nothing flat at
' scripts\english.vla since the polyglotta reorg - all three reported
' "vocabulary file on disk... none found" from the day of that reorg
' until this pass (found live, P-TOK's own verification click-through).
' Re-adding the polyglotta candidate to the shared list would reopen
' EDITIONMANIFEST.5's regression, since IdeLoadVocab calls this same
' function first. So the two kinds of caller see different candidate
' sets: includePolyglotta:=True appends the polyglotta path as a FIFTH
' candidate, after the four override locations; it defaults to False, so
' IdeLoadVocab's own pre-check is byte-for-byte unchanged. The path
' itself lives in IdeDevPolyglottaPath, shared with IdeLoadVocab's last
' resort, so the two can never drift apart.
' SEC.9: mayPrompt = False means "tell me which file you WOULD use, but
' do not stop to ask anyone about it." Two callers need that and neither
' loads any grammar: VlaIdeInfo, which builds a one-line diagnostics
' string, and the "vocabulary not found" message, which is assembling
' the text of an error that is already being raised. A status report
' that pops a consent dialog is a defect however good the consent is -
' found by reading the call sites after the owner's first live pass, not
' by the checker, which pins the two LOADING sites rather than these.
Private Function IdeVocabPath(Optional ByVal includePolyglotta As Boolean = False, _
                              Optional ByVal mayPrompt As Boolean = True) As String
    ' V5 (Mac spike): paths built with Application.PathSeparator -
    ' the one hard-coded-backslash site in the codebase, found by
    ' the portability inventory. Costless on Windows, correct on
    ' Mac, where Check (which loads the vocabulary) already works.
    Dim sep As String
    sep = Application.PathSeparator
    Dim fn As String
    fn = IdeVocabFileName()
    Dim cands(1 To 4) As String
    Dim hb As Workbook
    Set hb = HostBook()
    cands(1) = hb.Path & sep & "scripts" & sep & fn
    cands(2) = hb.Path & sep & fn
    cands(3) = ThisWorkbook.Path & sep & "scripts" & sep & fn
    cands(4) = ThisWorkbook.Path & sep & fn
    ' EDITIONMANIFEST.6 (owner-caught, live, second time this shape has
    ' bitten): the old guard (Len(cands(i)) > Len(sep & fn)) only
    ' protects the FLAT candidates (2 and 4) from a blank base path -
    ' it degenerates to exactly sep & fn there, correctly skipped. The
    ' "\scripts\" candidates (1 and 3) still pass it even with a blank
    ' base (hb.Path = "" - an unsaved "Book1", the same real scenario
    ' VlaIdeExportStandalone already guards explicitly, Len(hb.Path) = 0,
    ' a few dozen lines below), producing a driveless path like
    ' "\scripts\espanol.vla" - exactly the shape Dir$ raises "Bad file
    ' name or number" on, the same error this whole area already cost
    ' two live rounds to diagnose once. Guard the BASE path directly
    ' now, not the string that happens to result from it.
    Dim hasHostPath As Boolean, hasAddinPath As Boolean
    hasHostPath = (Len(hb.Path) > 0)
    hasAddinPath = (Len(ThisWorkbook.Path) > 0)
    Dim i As Long
    For i = 1 To 4
        If (i <= 2 And hasHostPath) Or (i >= 3 And hasAddinPath) Then
            If SafeFileExists(cands(i)) Then
                ' SEC.9: candidates 1 and 2 live in the HOST WORKBOOK's
                ' own directory and are chosen ahead of the add-in's own
                ' copy - the search-order hijack, and the half of SEC.9
                ' that was observed happening. They now need this
                ' device's approval; 3 and 4 are the add-in's own files
                ' and are trusted by construction.
                '
                ' A decline FALLS THROUGH to the next candidate rather
                ' than refusing, and that is the right shape here: the
                ' question being answered is "should the file sitting
                ' next to this workbook win over the built-in grammar?",
                ' and "no" has a complete, safe answer sitting one line
                ' further down the list. Nothing is lost by declining
                ' except the override.
                '
                ' Existence is probed BEFORE the gate here, unlike the
                ' replay loop, and the difference is deliberate: this
                ' path is the workbook's own folder, which Excel already
                ' opened the workbook from, so looking at it discloses
                ' nothing that opening the file has not already
                ' disclosed. The replay loop's paths are attacker-chosen
                ' strings pointing anywhere at all, which is a different
                ' question with a different answer.
                Dim useIt As Boolean
                useIt = True
                If i <= 2 And mayPrompt Then
                    useIt = PhrasebookPathApproved(cands(i), _
                        "This file sits next to the workbook you are working in, and Frazaro would use it INSTEAD of its own built-in grammar.")
                End If
                If useIt Then
                    IdeVocabPath = cands(i)
                    Exit Function
                End If
            End If
        End If
    Next
    ' EDITION-VOCABPATH: the dev workbook's own source file, for the
    ' ribbon callers that need a real file and have no embedded chain to
    ' fall back through (see the header). Never reached by IdeLoadVocab.
    If includePolyglotta And hasAddinPath Then
        If SafeFileExists(IdeDevPolyglottaPath()) Then
            IdeVocabPath = IdeDevPolyglottaPath()
            Exit Function
        End If
    End If
    ' Best candidate for the error message only (never passed to Dir$
    ' again) - prefer one with a real base path over cands(1), which is
    ' a driveless "\scripts\<fn>" and actively misleading when hb.Path
    ' is blank (an unsaved workbook - the message should point at a
    ' real-looking location, not this one).
    If hasHostPath Then
        IdeVocabPath = cands(1)
    ElseIf hasAddinPath Then
        IdeVocabPath = cands(3)
    Else
        IdeVocabPath = cands(1)
    End If
End Function

' The dev workbook's own phrasebook source, scripts\polyglotta\<edition
' file>, beside the add-in (ThisWorkbook - the dev workbook itself, or a
' built .xlam saved beside the repo). One definition, two readers:
' IdeLoadVocab's last resort after the embedded chain (EDITIONMANIFEST.5)
' and IdeVocabPath's fifth candidate for the ribbon buttons that need a
' real file (EDITION-VOCABPATH).
Private Function IdeDevPolyglottaPath() As String
    Dim sep As String
    sep = Application.PathSeparator
    IdeDevPolyglottaPath = ThisWorkbook.Path & sep & "scripts" & sep & "polyglotta" & sep & IdeVocabFileName()
End Function

' =====================================================================
'  GO.6: persisted, workbook-scoped phrasebook list. EnglishIdeLoadPhrasebook
'  ADDS a phrasebook the user picks on top of whatever IdeVocabPath/the
'  embedded chain already loaded (GO.1's ratified "last-loaded wins,
'  provenance always visible" multi-source precedence, via G3's existing
'  cross-file override mechanism in VLA_SentenceEngine.AddPhraseRule -
'  nothing new needed there), then remembers the file's path here so
'  IdeLoadVocab replays it on every future command in THIS workbook -
'  the owner's own call on GO.6's "does it persist" open question.
'
'  UNBOUNDED by design (owner correction, same session as the first
'  draft): this list is the same kind of thing as a source file's own
'  import statements - "base" alone is already prelude+english+espanol,
'  "org" could mean company-wide/department-wide/team-wide phrasebooks
'  stacked, and "community" is bottomless (pirate.vla, alien.vla, ...).
'  No programming language caps how many imports a file may declare;
'  this list does not either. One CustomDocumentProperty
'  (VLA_LoadedPhrasebooks) holds every path, vbLf-joined, in load
'  order - not numbered slots with a ceiling (that first draft copied
'  IdeLoadVocab's embedded-chain loop's bound without noticing that
'  loop is bounded for its OWN reason, a "does the next one exist"
'  probe over a small fixed set of workbook-embedded sheets, which does
'  not apply to a user's own open-ended, explicitly-chosen list).
'
'  Same CustomDocumentProperties mechanism as SEC.2's own workbook-
'  scope raw-consent record (VLA_SentenceEngine.VlaRawConsentRecordWorkbook).
'  Public, not a test-bypass toggle: VlaIdeProgramTag/VlaIdeModuleFor/
'  VlaIdeScanTargets already set the precedent of a real, reusable IDE
'  helper that is also directly unit-tested with pre-seeded state, same
'  as SEC.2's own workbook-scope test does for
'  ActiveWorkbook.CustomDocumentProperties.
' =====================================================================

' =====================================================================
'  SEC.9: which phrasebook paths this DEVICE has agreed to load.
'
'  THE HOLE THIS CLOSES, and it is not hypothetical. Two loaders
'  trusted the workbook's own account of what grammar to use:
'  IdeVocabPath put <host workbook dir>\scripts\<edition>.vla and
'  <host workbook dir>\<edition>.vla AHEAD of the add-in's own copy - a
'  DLL-search-order hijack wearing a .vla extension - and
'  ReplayPersistedPhrasebooks replayed absolute paths out of the
'  workbook's VLA_LoadedPhrasebooks property with no prompt at all.
'  Grammar decides what every sentence MEANS, so whoever supplies it
'  decides what the program does. OBSERVED LIVE 2026-09-08, benignly
'  and by accident: a stale english.vla in the owner's Downloads
'  folder was loaded ahead of the add-in's own copy, because the
'  workbook happened to be there too. It failed loudly only by luck -
'  that copy was old enough not to know "put", so it errored. A merely
'  DIFFERENT grammar would have loaded in silence. No attacker, no
'  crafted file, no privilege: an ordinary Downloads folder.
'
'  WHY THE RECORD LIVES ON THE DEVICE, not in the workbook. The
'  workbook is precisely the thing not being trusted here, so a
'  permission slip carried inside it would be written by the same hand
'  that wrote the request - SEC.10's whole finding about SEC.2's
'  workbook-scope grant, not repeated here by choice. The workbook's
'  VLA_LoadedPhrasebooks property keeps its old job, which is a
'  REQUEST ("this program wants these phrasebooks"); this registry
'  section is the ANSWER, and only the answer is trusted.
'
'  KEYED BY (path, digest), both halves load-bearing. Path alone would
'  let an approved location be swapped underneath the approval; digest
'  alone would let bytes approved in one directory authorize the same
'  bytes appearing somewhere else later. The digest is SEC.11's
'  EnglishSourceHash - which is why SEC.11 was built first: a grant
'  keyed to a forgeable fingerprint is worth about nothing, since an
'  attacker who can drop a file beside a workbook could also make it
'  collide under the old 32-bit polynomial.
'
'  THE ORDER OF OPERATIONS IS THE SECURITY PROPERTY, and it is easy to
'  get backwards. A path out of the workbook is an arbitrary
'  attacker-chosen string, and \\attacker\share\x.vla leaks this
'  machine's Windows credentials to that server the moment ANYTHING
'  touches it - Dir$ included. So a path with no recorded decision is
'  never touched at all: not hashed, not probed for existence, not
'  even asked about by name-that-exists. It is described to the person
'  as text and nothing more. Only once a decision exists for that path
'  does reading it become authorized, and only then is it hashed. That
'  is also why the lookup cannot simply be "hash it and see" - the
'  hash is exactly the thing that requires permission first.
'
'  DECLINING IS RECORDED TOO, deliberately, and this is a departure
'  from SEC.2's shape rather than an oversight. SEC.2 raises when raw
'  consent is refused, because a half-loaded phrasebook is
'  indistinguishable from a bug. Here a decline has a safe, complete,
'  obvious answer - the built-in grammar - so refusing loudly on every
'  command afterwards would trap a person in an error with no UI to
'  clear it (nothing in this codebase removes an entry from
'  VLA_LoadedPhrasebooks). A recorded "denied" is skipped quietly
'  thereafter and reported through the ordinary loaded-sources channel.
'  Changing the file re-opens the question either way, because the
'  record is content-keyed.
' =====================================================================
' (PHRASEBOOK_CONSENT_SECTION itself is declared at the top of this
' module, with the other module-level Consts - VBA requires every bare
' module-level declaration to sit in that one block before the first
' procedure, a rule VLA_SentenceEngine.bas's own SEC.2 constants
' record three live compiles learning.)

' PURE, and Public for exactly the reason LoadedPhrasebookPaths and
' VlaIdeProgramTag are: VlaSelfTest pins this directly, with no file,
' no workbook and no dialog. It is a string predicate and nothing else
' - deliberately, because the whole point is to reach a verdict on a
' path WITHOUT touching it.
'
' "Remote" means UNC (\\server\share, and the \\?\ / \\.\ prefixed
' forms, which start the same way) or anything carrying a URL scheme.
' The "://" test is the general one: no Windows local path can contain
' it, because a drive letter is a bare colon with no slashes. Forward
' slashes are treated like backslashes throughout, since Windows
' accepts // as a UNC introducer too and a stored path is whatever the
' workbook chose to write.
'
' NOT detected, and stated rather than implied: a mapped drive letter
' (Z:\x.vla) that resolves to a share. Telling those apart needs a
' host call, which would defeat the purpose of a pure predicate, and
' the mapping was made by this user on this machine rather than named
' by the workbook. It still passes through the consent gate below like
' any other path.
Public Function VlaPhrasebookPathIsRemote(ByVal path As String) As Boolean
    Dim p As String
    p = Replace(Trim$(path), "/", "\")
    If Len(p) = 0 Then Exit Function
    If Left$(p, 2) = "\\" Then
        VlaPhrasebookPathIsRemote = True
        Exit Function
    End If
    ' Scheme test on the ORIGINAL text: "://" survives the slash
    ' normalization above as ":\\", so look for either spelling rather
    ' than depending on which one ran first.
    If InStr(path, "://") > 0 Or InStr(p, ":\\") > 0 Then
        VlaPhrasebookPathIsRemote = True
    End If
End Function

' PURE. Is path inside dir - the add-in's own folder, whose contents
' ship with Frazaro and are trusted by construction, exactly as
' SEC.16's own accepted-risk note assumes (anything able to write
' there is already running as this user).
'
' Case-insensitive because Windows paths are, and the separator is
' appended before comparing so that a sibling directory whose name
' merely STARTS with the trusted one - C:\Frazaro-evil next to
' C:\Frazaro - cannot pass as being inside it. That prefix trap is the
' reason this is a named, tested function rather than an inline
' InStr.
Public Function VlaPhrasebookPathIsUnderDir(ByVal path As String, ByVal trustedDir As String) As Boolean
    ' trustedDir rather than the obvious "dir": Dir is a VBA intrinsic,
    ' and a parameter of that name shadows it for the whole procedure -
    ' harmless here only because nothing below calls it, which is
    ' exactly how that trap stays invisible until someone edits this.
    If Len(trustedDir) = 0 Or Len(path) = 0 Then Exit Function
    Dim d As String, p As String
    d = Replace(Trim$(trustedDir), "/", "\")
    p = Replace(Trim$(path), "/", "\")
    If Right$(d, 1) <> "\" Then d = d & "\"
    If Len(p) <= Len(d) Then Exit Function
    VlaPhrasebookPathIsUnderDir = (StrComp(Left$(p, Len(d)), d, vbTextCompare) = 0)
End Function

' The gate. True means "load it"; False means "do not", and the caller
' decides whether that is a fallback or a refusal.
'
' whyAsked names the situation in the person's own terms - the two
' callers are genuinely different questions ("this file sits next to
' your workbook" vs "your workbook remembers this path"), and a prompt
' that cannot say which is a prompt nobody can answer well.
Private Function PhrasebookPathApproved(ByVal path As String, ByVal whyAsked As String) As Boolean
    ' Trusted by construction: anything inside the add-in's own folder
    ' ships with Frazaro, and something able to write there is already
    ' running as this user - SEC.16's accepted-risk reasoning, applied
    ' rather than restated. Checked FIRST so the dev workbook's own
    ' scripts\polyglotta\english.vla, and any add-in-folder phrasebook a
    ' person once picked deliberately, never turn into a dialog on
    ' every Check.
    If VlaPhrasebookPathIsUnderDir(path, ThisWorkbook.Path) Then
        PhrasebookPathApproved = True
        Exit Function
    End If

    Dim rec As String
    rec = GetSetting("Frazaro", PHRASEBOOK_CONSENT_SECTION, LCase$(Trim$(path)), "")

    If Len(rec) = 0 Then
        ' No decision has ever been made about this path, so nothing
        ' here may touch it - see the header's order-of-operations
        ' note. Ask about the text alone.
        PhrasebookPathApproved = AskAndRecordPhrasebookPath(path, whyAsked)
        Exit Function
    End If

    ' A decision exists, so reading this path is authorized now.
    Dim storedDigest As String, decision As String
    Dim bar As Long
    bar = InStr(rec, "|")
    If bar > 0 Then
        storedDigest = Left$(rec, bar - 1)
        decision = Mid$(rec, bar + 1)
    Else
        decision = rec
    End If

    ' An empty stored digest means the decision was reached WITHOUT ever
    ' reading the file - the only honest outcome for a path that was
    ' declined, or for a remote one that was never contacted. There are
    ' no bytes it can be compared against, so the decision stands as
    ' given. Getting this wrong is not academic: the first version
    ' hashed only on approval and then compared that empty digest
    ' against a real one, so a "no" never matched itself and the gate
    ' asked again on every single command. Live-caught by the owner on
    ' the first pass - "denying the sibling english.vla both times".
    If Len(storedDigest) = 0 Then
        PhrasebookPathApproved = (decision = "granted")
        Exit Function
    End If

    If Not SafeFileExists(path) Then Exit Function   ' gone: nothing to load, nothing to ask

    Dim current As String
    current = EnglishSourceHash(path)
    If current = storedDigest Then
        PhrasebookPathApproved = (decision = "granted")
        Exit Function
    End If

    ' Same path, different contents. The old answer was about bytes
    ' that are no longer there, so it does not carry over - in either
    ' direction.
    PhrasebookPathApproved = AskAndRecordPhrasebookPath(path, whyAsked & vbCrLf & vbCrLf & _
        "(This file has changed since you last answered for it.)")
End Function

Private Function AskAndRecordPhrasebookPath(ByVal path As String, ByVal whyAsked As String) As Boolean
    Dim prompt As String
    prompt = "Load grammar rules from this file?" & vbCrLf & vbCrLf & _
             path & vbCrLf & vbCrLf & _
             whyAsked & vbCrLf & vbCrLf & _
             "A phrasebook defines what your sentences MEAN, so a different one can silently change what a program does. Only load one you trust."
    If VlaPhrasebookPathIsRemote(path) Then
        prompt = prompt & vbCrLf & vbCrLf & _
                 "This is a network or web location. Frazaro has not contacted it - opening it would hand your Windows sign-in to that server."
    End If

    Dim answer As Long
    answer = MsgBox(prompt, vbYesNo Or vbExclamation Or vbDefaultButton2, "Frazaro - load this phrasebook?")

    ' Record the answer keyed to the bytes it was given about, so that
    ' editing the file re-opens the question. The digest is computed
    ' whenever reading the path is legitimate: after a yes (they just
    ' authorized it), or for any LOCAL path, since reading local bytes
    ' to fingerprint them is not the hazard - the hazard is contacting a
    ' server the workbook named, which is why a declined REMOTE path is
    ' the one case that stays unhashed and records an empty digest.
    '
    ' Hashing only on yes, which is what the first version did, meant a
    ' "no" stored an empty digest that could never match the real one,
    ' so the gate re-asked on every command forever. Owner-caught live.
    Dim digest As String
    If answer = vbYes Or Not VlaPhrasebookPathIsRemote(path) Then
        On Error Resume Next
        digest = EnglishSourceHash(path)
        On Error GoTo 0
    End If

    SaveSetting "Frazaro", PHRASEBOOK_CONSENT_SECTION, LCase$(Trim$(path)), _
                digest & "|" & IIf(answer = vbYes, "granted", "denied")
    AskAndRecordPhrasebookPath = (answer = vbYes)
End Function

' Records a grant for a path the person chose themselves in a file
' dialog. Picking a file in an Open dialog IS the decision the gate
' above would otherwise stop to ask for, so asking again immediately
' would be theatre - but the RECORD still has to exist, or the replay
' on the next Check would ask about a file they just deliberately
' opened.
Private Sub GrantPhrasebookPath(ByVal path As String)
    Dim digest As String
    On Error Resume Next
    digest = EnglishSourceHash(path)
    On Error GoTo 0
    SaveSetting "Frazaro", PHRASEBOOK_CONSENT_SECTION, LCase$(Trim$(path)), digest & "|granted"
End Sub

' Every persisted phrasebook path for hb, in the order they were
' loaded.
Public Function LoadedPhrasebookPaths(hb As Workbook) As Collection
    Dim r As New Collection
    Dim raw As String
    raw = ""
    On Error Resume Next
    raw = CStr(hb.CustomDocumentProperties(PHRASEBOOK_LIST_PROP).Value)
    On Error GoTo 0
    If Len(raw) = 0 Then
        Set LoadedPhrasebookPaths = r
        Exit Function
    End If
    Dim parts() As String
    parts = Split(raw, vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then r.Add parts(i)
    Next i
    Set LoadedPhrasebookPaths = r
End Function

' Case-insensitive: the same real file reached through two differently-
' cased path strings (Windows paths are not case-sensitive) must count
' as the one already-loaded phrasebook, not a second one.
Public Function PhrasebookAlreadyLoaded(hb As Workbook, ByVal path As String) As Boolean
    Dim existing As Collection
    Set existing = LoadedPhrasebookPaths(hb)
    Dim e As Variant
    For Each e In existing
        If StrComp(CStr(e), path, vbTextCompare) = 0 Then
            PhrasebookAlreadyLoaded = True
            Exit Function
        End If
    Next
End Function

' Remembers path for hb, unless it is already remembered (a no-op, not
' a duplicate entry - EnglishIdeLoadPhrasebook's own caller already
' checked PhrasebookAlreadyLoaded before ever calling this, but a
' second, independent guard here costs nothing and keeps this function
' safe to call on its own).
Public Sub PersistPhrasebookPath(hb As Workbook, ByVal path As String)
    If PhrasebookAlreadyLoaded(hb, path) Then Exit Sub
    Dim existing As Collection
    Set existing = LoadedPhrasebookPaths(hb)
    Dim joined As String
    Dim e As Variant
    For Each e In existing
        joined = joined & CStr(e) & vbLf
    Next
    joined = joined & path
    On Error Resume Next
    hb.CustomDocumentProperties(PHRASEBOOK_LIST_PROP).Value = joined
    If Err.Number <> 0 Then
        Err.Clear
        hb.CustomDocumentProperties.Add Name:=PHRASEBOOK_LIST_PROP, LinkToContent:=False, Type:=4, Value:=joined   ' 4 = msoPropertyTypeString, the literal - see VlaRawConsentRecordWorkbook's own identical note
    End If
    On Error GoTo 0
End Sub

' Called from IdeLoadVocab, on every command, right after the base
' corpus is settled (whichever of its three sources supplied it) -
' layers every remembered phrasebook on top, ADD semantics, in the
' order they were originally loaded. A moved or deleted file is a
' benign environmental problem, not a content bug - skipped with a
' Debug.Print note (IdeLoadVocab's own trailing diagnostics are
' already Debug.Print, never a live dialog, since this runs inside
' every Check/Interpret/Run, not just at file open) rather than a hard
' raise, so losing one phrasebook file never blocks every other
' command in the workbook. A genuine CONTENT collision - an override
' that no longer matches after the base corpus changed, a same-shape
' rule with no override: marker - still raises exactly as it always
' has; that is real, human-attention-worthy grammar state, not an
' environmental hiccup, and swallowing it here would hide the exact
' bug G3/GO.1 exist to surface.
Public Sub ReplayPersistedPhrasebooks()
    Dim hb As Workbook
    Set hb = HostBook()
    Dim paths As Collection
    Set paths = LoadedPhrasebookPaths(hb)
    Dim p As Variant
    For Each p In paths
        ' SEC.9: the gate comes FIRST, before SafeFileExists - and that
        ' order is the whole point, not a style preference. These paths
        ' are arbitrary strings out of a document property the workbook's
        ' author wrote, and Dir$ on \\attacker\share\x.vla hands this
        ' machine's Windows credentials to that server before it ever
        ' returns an answer. A "does it exist?" check placed above this
        ' line would have already lost. PhrasebookPathApproved touches
        ' nothing until a decision for the path exists.
        If PhrasebookPathApproved(CStr(p), _
               "This workbook remembers this phrasebook and is asking to load it again.") Then
            If SafeFileExists(CStr(p)) Then
                EnglishLoadVocabulary CStr(p)
            Else
                Debug.Print "GO.6: remembered phrasebook not found, skipped: " & CStr(p)
            End If
        Else
            ' Skipped, not raised. The built-in grammar is a complete and
            ' safe answer here, and nothing in this codebase can remove an
            ' entry from VLA_LoadedPhrasebooks - so refusing loudly on
            ' every command would trap a person in an error they have no
            ' way to clear. The decision is recorded, so this is silent on
            ' the second pass rather than a prompt per command.
            Debug.Print "SEC.9: phrasebook not approved on this device, skipped: " & CStr(p)
        End If
    Next
End Sub

' =====================================================================
'  D1: ambient-state capture. The workbook the person is working in
'  is read from ActiveWorkbook exactly ONCE, at the entry of each
'  public command, and every helper in this module sees that same
'  handle for the command's duration. Before D1, HostBook() re-read
'  ActiveWorkbook on every call, so a focus change mid-command -
'  a third-party add-in's event handler activating something during
'  our own sheet writes, a program's "Open workbook ..." leaving a
'  different book active, a click landing in a yield - could bind
'  different halves of one command to different workbooks. Object
'  references, once Set, are immune to focus changes; the capture
'  turns every ambient read into an object reference at the earliest
'  possible moment. Command semantics are unchanged and deliberate:
'  a menu command binds to whichever workbook is active AT INVOCATION
'  (that is what invoking it there means), then stays bound.
' =====================================================================

' Read the ambient world once. Every public entry point calls this
' first; nothing else in the module may touch ActiveWorkbook.
' V2: the ACTIVE SHEET is captured too - a workspace button runs on
' the sheet that owns it, so the sheet active at invocation IS the
' program the person meant. Declared Object, not Worksheet: a chart
' sheet can be active, and the resolver treats it as no-sheet.
Private Sub CaptureHost()
    Set mHost = ActiveWorkbook     ' add-ins never appear here, which
                                   ' is why this was never ThisWorkbook
    If mHost Is Nothing Then
        VLA_Messages.RaiseMsg "ide-no-workbook"
    End If
    Set mHostSheet = Nothing
    On Error Resume Next
    Set mHostSheet = mHost.ActiveSheet
    On Error GoTo 0

    ' SEC.8: read the host workbook's Mark-of-the-Web ONCE per command,
    ' here, for exactly D1's own reason - this is the moment the ambient
    ' world becomes a held reference, and provenance is ambient state
    ' like any other. Every external-effect dispatch site then consults
    ' the memo instead of re-reading. Deliberately does NOT refuse
    ' anything: capturing is not gating, a program that never reaches
    ' outside the workbook must run unimpeded no matter where the file
    ' came from, and the refusal belongs at the site that names the verb.
    ' FullName on a never-saved workbook is the bare name with no path;
    ' VlaPathIsDemonstrablyLocal reads that as local, which it is.
    On Error Resume Next
    VLA_Provenance.VlaProvenanceCapture HostFullPath()
    On Error GoTo 0
End Sub

' SEC.8: the host workbook's path as the provenance gate needs it -
' "" when the workbook has never been saved (Path is empty, and
' FullName is then just the display name, which is not a path and must
' not be treated as one). Kept beside CaptureHost so the one place
' that reads this ambient fact is the one place that captures it.
Private Function HostFullPath() As String
    On Error Resume Next
    If mHost Is Nothing Then Exit Function
    If Len(mHost.Path) = 0 Then Exit Function
    HostFullPath = mHost.FullName
End Function

' The captured handle. The lazy path serves Immediate-window calls
' that arrive without a command entry, and re-captures if the held
' workbook has been closed since (a dead handle errors on .Name).
' V5 (Mac spike): compile-time platform truth, one place.
Private Function VlaOnMac() As Boolean
#If Mac Then
    VlaOnMac = True
#End If
End Function

Private Function HostBook() As Workbook
    If Not mHost Is Nothing Then
        Dim nm As String
        On Error Resume Next
        nm = mHost.Name
        If Err.Number <> 0 Then Set mHost = Nothing
        On Error GoTo 0
    End If
    If mHost Is Nothing Then CaptureHost
    Set HostBook = mHost
End Function

' ---------------------------------------------------------------------
'  Buttons
' ---------------------------------------------------------------------

Public Sub VlaIdeSetup()
    On Error GoTo failed
    CaptureHost
    MigrateLegacySheet
    Dim isNew As Boolean
    isNew = Not SheetExists(HostBook(), IDE_SHEET)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(IDE_SHEET)
    BuildWorkspace ws
    ws.Activate
    If isNew Then SelectFirstSentenceCell ws
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' V2: a second (third, ...) program in the same workbook. The sheet
' is "Frazaro (<name>)"; identical buttons, its own module, its own
' snapshots. Name rules are checked HERE, at birth, so the Run-time
' guards rarely fire: length (the tab must fit), Excel's forbidden
' characters, and tag collision with existing programs.
Public Sub EnglishIdeAddProgram()
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Set hb = HostBook()
    Dim nm As String
    nm = Trim$(InputBox("Name the new program - its sheet will be '" & IDE_SHEET & " (<name>)'." & vbCrLf & vbCrLf & _
                        "Short names work best: letters and numbers, up to 20 characters.", "Frazaro"))
    If Len(nm) = 0 Then Exit Sub
    If Len(nm) > 20 Then
        VLA_Messages.RaiseMsg "ide-program-name-too-long"
    End If
    Dim i As Long
    For i = 1 To Len(nm)
        If InStr(":\/?*[]", Mid$(nm, i, 1)) > 0 Then
            VLA_Messages.RaiseMsg "ide-sheet-name-bad-char", "char", Mid$(nm, i, 1)
        End If
    Next
    Dim full As String
    full = IDE_SHEET & " (" & nm & ")"
    Dim tag As String
    tag = VlaIdeProgramTag(full)
    Dim wsV As Variant
    For Each wsV In WorkspaceSheets(hb)
        If VLA_Identity.Fold(wsV.Name) <> VLA_Identity.Fold(full) Then
            If VlaIdeTagsMatch(VlaIdeProgramTag(wsV.Name), tag) Then
                VLA_Messages.RaiseMsg "ide-program-name-too-similar", "new", full, "existing", wsV.Name, "tag", tag
            End If
        End If
    Next
    Dim isNew As Boolean
    isNew = Not SheetExists(hb, full)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(full)
    BuildWorkspace ws
    ws.Activate
    If isNew Then SelectFirstSentenceCell ws
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' The workspace furniture - shared verbatim by the default program
' and every named one, so all programs feel identical.
Private Sub BuildWorkspace(ws As Worksheet)
    ' Owner request: the IDE pair shifted right one column - B is now
    ' the sentence column, C the result column - so A is free to be a
    ' file-reference helper and D/E/F are plain scratch columns. The
    ' guidance that used to occupy A1 is a cell comment now, anchored
    ' on B1 (the sentence column's own first cell) so it stays where a
    ' user is actually looking.
    ' Hiding a large multi-thousand-column range was tried and
    ' abandoned here (owner call, three straight "Unable to set the
    ' Hidden property of the Range class" failures across three
    ' different approaches - literal "D:XFD", ws.Columns.Count, and
    ' hiding all-but-the-last-column separately - not worth chasing
    ' further for a cosmetic). D, E, F and everything past them are
    ' left visible; A alone still hides fine below - a single-column
    ' hide is the one form of this that has never failed.
    EnsureModernLayout ws
    On Error Resume Next
    ws.Cells(1, 2).Comment.Delete
    On Error GoTo 0
    ws.Cells(1, 2).AddComment _
        "Write the program here, one sentence per row - not strictly required, but keeps Validate's marks meaningful; blank rows end blocks; # starts a comment." & vbCrLf & _
        "Validate marks each row in column C; results go to the Output sheet unless the program says: Work on sheet <name>." & vbCrLf & _
        "A1 remembers the file Reload Instructions should re-import. Columns D, E and F are free for your own values or imports - a sentence can reference a cell (e.g. D5) instead of spelling out a calculation."
    ws.Cells(1, 2).Comment.Shape.TextFrame.AutoSize = True
    ws.Columns(2).ColumnWidth = 72
    ws.Columns(2).NumberFormat = "@"          ' sentences are text, never formulas
    ws.Columns(2).Interior.Color = RGB(247, 244, 252)   ' owner request: a whisper of the logo's lavender
    ws.Columns(3).ColumnWidth = 60
    ws.Columns(3).WrapText = True
    ' Owner request: a very slight gray base for the result column, so
    ' it reads as its own lane next to white column D even on rows
    ' Check has not marked yet. MarkOK/MarkErr paint over this per row
    ' once there is a verdict; ClearMarks (below) restores this same
    ' gray, not colorless, when a fresh Check clears the previous
    ' run's marks.
    ws.Columns(3).Interior.Color = RGB(242, 242, 242)
    ' A holds only the file-reference helper (below); small grey font
    ' keeps it visually de-emphasized since it is not part of the
    ' program a user is writing.
    ws.Cells(1, 1).Font.Size = 8
    ws.Cells(1, 1).Font.Color = RGB(140, 140, 140)
    ' Delete BEFORE hiding: on a sheet that already has furniture from
    ' an earlier Setup (GetOrCreateSheet fetched it, not created it),
    ' the old Buttons are still anchored starting at G1 - Excel refuses
    ' to collapse a column with a Shape anchored in it ("Unable to set
    ' the Hidden property of the Range class"), so the delete must run
    ' first, not after.
    ' Owner request, this pass: the G1-down sheet-button stack that
    ' used to live right here (Check/Interpret/Compile/Show VBA/Undo/
    ' Reload/Known Sentences/Copy Feedback/Add Program) is gone - every
    ' one of those commands is now a ribbon button (the regroup two
    ' commits back), so the stack was pure duplicate clutter on every
    ' workspace sheet. `ws.Buttons.Delete` stays: it is what clears an
    ' ALREADY-EXISTING sheet's old stack the next time Setup/Add Program
    ' touches it, not just a guard for buttons this pass still creates.
    ws.Buttons.Delete
    ws.Columns(1).Hidden = True   ' a lone single-column hide - the only kind that has ever actually worked here
    ws.Activate
    ActiveWindow.DisplayGridlines = False   ' see RenderVbaSource's own note on this being per-sheet, not per-workbook
End Sub

' Owner request, with TER-2's live pass: a NEW workspace tab, named or
' not, opens with B1 - the first sentence cell - active. Excel starts a
' new sheet on A1, and BuildWorkspace hides column A, so the cursor sat
' in a hidden cell and had to be moved one right before typing. Only on
' creation: re-running Setup on an existing tab leaves its selection
' alone. Both callers activate the sheet first, which Select requires.
Private Sub SelectFirstSentenceCell(ws As Worksheet)
    ws.Cells(FIRST_ROW, 2).Select
End Sub

Public Sub EnglishIdeCheck()
    On Error GoTo failed
    CaptureHost
    DoCheck IdeSheet()
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Public Sub EnglishIdeRun()
    RunProgram False
End Sub

' IN.6: the interpreter-backed entry points - "Interpret Instructions"/
' "Interpret and Trace", the Interpreter ribbon section's own pair,
' mirroring EnglishIdeRun/EnglishIdeRunTrace exactly, one call deeper.
Public Sub EnglishIdeInterpret()
    InterpretProgram False
End Sub

Public Sub EnglishIdeInterpretTrace()
    InterpretProgram True
End Sub

' S5.3 (owner design): Run & Trace - the trace's real home. The flag
' is armed AFTER compile and read in the SAME execution chain as the
' run, so no state has to survive an idle boundary or a project reset
' (the S5.0/S5.1 failures both lived in that gap). The transcript
' pops in a window a user can copy and mail to support - no Immediate
' window, no VBA developer required.
Public Sub EnglishIdeRunTrace()
    RunProgram True
End Sub

' The shared Run core. wantTrace arms VlaTrace around main only; every
' other step - Check, the S4.3 gates, inject, compile, snapshot,
' protect, the failure handler - is identical, so tracing can never
' change what a program DOES, only whether its steps are recorded.
Private Sub RunProgram(ByVal wantTrace As Boolean)
    Dim ws As Worksheet
    Dim hb As Workbook
    Dim running As Boolean
    ' U.29: Excel's settings as this run found them, whether they have
    ' been recorded yet, and whether the run protected the program sheet.
    Dim borrowed As VlaExcelSettings
    Dim recorded As Boolean
    Dim runProtected As Boolean
    On Error GoTo failed
    CaptureHost
    ' V5 (Mac spike): the transpiler road runs through VBProject
    ' automation, which Mac Excel does not permit - so Run is
    ' honestly Windows-only today, said up front in one sentence
    ' instead of discovered as a cryptic automation error mid-pilot.
    ' Check still works on Mac (translation touches no VBProject),
    ' which makes write-and-check-on-Mac, run-on-Windows a real
    ' workflow. Mac RUNNING arrives with interpreter mode (F1).
    If VlaOnMac() Then
        VLA_Messages.RaiseMsg "ide-compile-needs-windows"
    End If
    ' IN.6: Compile writes real VBA into the workbook's VBProject, the
    ' one Excel setting managed IT departments routinely disable
    ' (BETA_REARVIEW.md's own "gate to every enterprise conversation").
    ' Refuse in words, naming the reason and the alternative (SD-10),
    ' now that the alternative - Interpret Instructions, which touches
    ' no VBProject at all - sits right beside this button instead of
    ' being a roadmap decision with no button behind it yet.
    If Not VlaHasVbProjectTrust() Then
        VLA_Messages.RaiseMsg "ide-compile-needs-trust"
    End If
    Set hb = HostBook()
    Set ws = IdeSheet()
    ' V1: Check hands back the VLA of its one translation, and the
    ' compile consumes it directly - a Run now translates the program
    ' exactly once (it used to translate twice: once checking, once
    ' inside EnglishCompileToModule).
    ' S5.3 (owner catch): an empty program from the Run button used to
    ' borrow Check's "Nothing to check" wording (wrong verb) or worse.
    ' Say the true thing, in Run's voice, and touch nothing.
    If ProgramIsBlank(ProgramText(ws, FIRST_ROW, IdeLastRow(ws))) Then
        VlaShowInfo "No instructions to compile - write a program in column B of the '" & ws.Name & "' sheet first."
        Exit Sub
    End If
    Dim vla As String
    If Not DoCheck(ws, vla) Then Exit Sub     ' bad row is already marked
    Dim progText As String
    progText = ProgramText(ws, FIRST_ROW, IdeLastRow(ws))
    ' V2: everything below is scoped to THIS program - its module,
    ' its snapshot prefix - after a loud guard against two programs
    ' sharing one tag.
    GuardTagCollision hb, ws
    Dim outMod As String
    outMod = VlaIdeModuleFor(ws.Name)
    ' V5.3: ship the runtime helper zoo into the user's workbook as
    ' Frazaro_EN_Runtime FIRST, so the generated module's helper calls
    ' (vlacolor, vlacount, the VlaDict family...) resolve in-project.
    ' Before this, they resolved only because the dev workbook - which
    ' happened to hold the helpers - was open; the first standalone
    ' run failed to compile at vlacolor the moment it was closed.
    ' Best-effort by soft-failure doctrine: if injection somehow
    ' fails, the program compile below will surface it with the
    ' same "Sub or Function not defined" - no worse than before, and
    ' the message names the missing helper.
    VlaInjectRuntime hb
    ' D1: the compile injects into the CAPTURED host, not whichever
    ' workbook happens to be active by the time the VBE call lands.
    VlaCompileToModule vla, outMod, hb
    ' (S4.1: the pre-run compile probe was removed after the owner's
    ' deliberate-defect test disproved its mechanism - see the module
    ' note in VLA_English's BuildStepInfra and the Alpha 3 roadmap.)

    ' Safety net: snapshot the sheets this run can touch, so "Undo
    ' last run" can put everything back. U.19: a snapshot that cannot be
    ' made now STOPS the Run, in words, before its first sentence - it
    ' used to go on with only an Immediate-window note, Undo silently
    ' gone for that Run and a half-made copy left standing. Stopping
    ' keeps the promise the Run button makes, that what it changes can
    ' be undone: a warning AFTER the run would arrive once the sheets had
    ' already changed, with nothing left to decide, and a question BEFORE
    ' it would be a mid-run dialog. The failures are ones the person can
    ' clear (a protected workbook, a sheet already holding a name
    ' Frazaro keeps for its copies), and the refusal says how.
    ' TakeRunSnapshot removes everything it made before it raises.
    TakeRunSnapshot hb, progText, VlaIdeProgramTag(ws.Name), ws.Name

    ' The workspace sheet: "Output" by default, or whatever the
    ' program declares with "Work on sheet <name>." U.28: made AFTER
    ' the snapshot, so a sheet this makes is one the snapshot has
    ' marked for removal, like every other sheet a run creates, and a
    ' stop or Undo Last Run removes it and says so. It used to be made
    ' before, so that "Undo restores to blank" - which put back an
    ' empty sheet the person never had, and called it "as they were
    ' before the run". A workspace sheet that already existed is
    ' copied and put back, as ever.
    Dim outName As String
    outName = DeclaredOutputSheet(progText)
    If Len(outName) = 0 Then outName = OUT_SHEET
    Dim outWs As Worksheet
    Set outWs = GetOrCreateSheet(outName)

    ' U.29: Excel's settings as the run found them, recorded before its
    ' first change and given back on every exit below (VlaIdeRecordExcel
    ' says which, and why each goes back to what it was).
    VlaIdeRecordExcel borrowed
    recorded = True

    ' Sheet context: unqualified cell references in the vocabulary
    ' mean "the current sheet" (VBA's ActiveSheet), and "Go to sheet
    ' X." is the context switch. Programs therefore START on the
    ' Output sheet, and the program sheet is protected while they run,
    ' so a program cannot overwrite its own sentences - an attempted
    ' write there fails at its step with a clear "protected" message.
    ' U.29: protected only when it has no protection already, and
    ' unprotected only when the run protected it - a protection the
    ' person made is theirs. (It used to be protected and unprotected
    ' whatever it was, which would have taken a person's protection
    ' away; Check's first write to column C refuses a protected sheet
    ' in Excel's own words, so no Run reached that far until now.)
    outWs.Activate
    runProtected = Not ws.ProtectContents
    If runProtected Then ws.Protect
    ' V1: the run bracket - screen updating off for the duration,
    ' given back on BOTH exits (U.29: as the run found it, which for a
    ' Run a person starts is on; it was set True, Excel's default).
    ' IDE-side and screen-updating ONLY;
    ' two roads deliberately not taken, reasons on the Alpha 2
    ' roadmap: wrapping main's body in (with-fast-excel ...) would
    ' displace the step handler (its restore-then-re-raise fires
    ' inside an active handler, so vla-fail can never catch it and
    ' step attribution dies), and turning calculation off changes
    ' SEMANTICS - a program that writes a formula and reads its
    ' result back (instructions.txt does, B3 sum-of-formula) would read
    ' stale values. Screen updating is pure speed. The restore also
    ' un-freezes the screen for a program that turned updating off
    ' and then failed before turning it back on - a real hole today.
    ' S5.3: arm the trace AFTER compile (the compile's project-state
    ' reset is now upstream of the flag) and in this same stack, so
    ' the run's VlaTraceOn reads a live flag - the durable Name still
    ' backs it, but nothing has to cross an idle boundary anymore.
    ' S5.5: defense in depth. The top-of-path empty guard already
    ' refuses an empty program, but Application.Run against a module
    ' with no runnable main is the ONE failure that reaches the user
    ' as Excel's raw "Cannot run the macro" box (the owner's
    ' screenshot). Confirm main exists before calling it - so even a
    ' future path that reaches here with a bodyless module refuses in
    ' Frazaro's voice, having touched nothing.
    If Not ModuleHasMain(hb, outMod) Then
        GiveBackRun ws, borrowed, runProtected
        If wantTrace Then VlaTrace False
        VlaShowInfo "No instructions to compile - write a program in column B of the '" & ws.Name & "' sheet first."
        Exit Sub
    End If
    ' S5.6 (standing-override field report): arm the copy the run
    ' will ACTUALLY consult. In a user workbook that is the injected
    ' Frazaro_EN_Runtime - main resolves vlatraceon in-project - while this
    ' Sub executes in the add-in; arming the add-in's own instance
    ' armed nothing the steps could see (the dev workbook masked
    ' this: same-project guard skips injection there, so one
    ' instance served both roles). Host-qualified arm, local
    ' fallback for the dev topology.
    If wantTrace Then ArmTrace hb, True
    ' U.25: the compiled program's own handler catches its stop, so this
    ' Sub learns of it only by asking: armed, VlaReportStop records the
    ' stop instead of showing it, and the stop is handled below exactly
    ' as the interpreter's is.
    ArmStopReport hb, True
    running = True
    Application.ScreenUpdating = False
    Application.Run "'" & hb.Name & "'!" & outMod & ".main"
    running = False
    ' U.29: finished or stopped, Excel's settings go back before a stop's
    ' put-back and its message, so both happen in Excel as it was.
    GiveBackRun ws, borrowed, runProtected
    Dim stopRep As String
    stopRep = ReadStopReport(hb)
    ArmStopReport hb, False
    If wantTrace Then ArmTrace hb, False
    If Len(stopRep) > 0 Then
        Dim stopLine As Long, stopCut As Long, stopSentence As String
        stopCut = InStr(stopRep, vbLf)
        If stopCut = 0 Then stopCut = Len(stopRep) + 1
        stopSentence = VlaIdeSplitStepText(Left$(stopRep, stopCut - 1), stopLine)
        ReportStoppedRun hb, ws, stopLine, stopSentence, Mid$(stopRep, stopCut + 1)
        If wantTrace Then ShowTraceWindow hb
        Exit Sub
    End If
    If wantTrace Then ShowTraceWindow hb
    ' VerifyReports stale-read fix (VLA_Runtime's own header note has
    ' the full incident): only stamped on a run that reached HERE
    ' without hitting "failed" below, so a crashed run leaves whatever
    ' provenance was already there rather than falsely claiming success.
    On Error Resume Next
    VLA_Runtime.VlaStampRunProvenance "emitter"
    On Error GoTo 0
    Exit Sub
failed:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    ' U.29: a failure before the record changed nothing of Excel's, and
    ' left no protection of the run's to take away.
    If recorded Then
        GiveBackRun ws, borrowed, runProtected
    Else
        Application.ScreenUpdating = True
    End If
    ' S5.3: a traced run that failed still shows the steps it reached -
    ' the trace up to the failure is often the whole diagnosis.
    ' (S5.6: wantTrace alone decides; the old wasTracing check read
    ' the LOCAL instance, which standalone is the wrong one.)
    If wantTrace Then ArmTrace hb, False
    ArmStopReport hb, False
    On Error GoTo 0
    ' U.25: a failure that escaped the program's own handler while it
    ' ran is a stop too, with no step to name.
    If running Then
        ReportStoppedRun hb, ws, 0, "", d
    Else
        VlaShowError d
    End If
    If wantTrace Then ShowTraceWindow hb
End Sub

' IN.6: does this Excel grant "Trust access to the VBA project object
' model"? The standard probe - touch VBProject under On Error Resume
' Next and see whether it complained - the same technique
' VlaCompileToModule's own callers already rely on implicitly, made
' explicit here so Compile can refuse in words BEFORE spending several
' steps discovering the same thing the hard way. The setting is
' application-wide, not per-workbook, so ThisWorkbook (the add-in
' itself, always open while any of this runs) is as good a probe
' target as the host.
Private Function VlaHasVbProjectTrust() As Boolean
    On Error Resume Next
    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count
    VlaHasVbProjectTrust = (Err.Number = 0)
    On Error GoTo 0
End Function

' IN.6: the interpreter-backed twin of RunProgram, same shape, one
' real difference - no VlaInjectRuntime, no VlaCompileToModule, no
' Application.Run against a compiled module, because there is no
' module. VLA_Interpreter.VlaInterpret runs the SAME translated VLA
' text directly, in-project, against whatever ActiveWorkbook/
' ActiveSheet already are (outWs.Activate below still does that job -
' the interpreter's own dotted-global dispatch reads ActiveWorkbook/
' ActiveSheet exactly like the emitted VBA does). No VBProject touch
' anywhere in this Sub, which is IN.6's whole point: no Mac guard (the
' ONLY reason RunProgram refuses on Mac is the VBProject automation
' Compile needs), no trust probe, nothing for a locked-down enterprise
' machine to refuse. GuardTagCollision and TakeRunSnapshot stay -
' both are pure worksheet operations (checked directly, no VBProject
' calls in either), and Undo has to work identically regardless of
' which backend wrote the sheets it is restoring.
Private Sub InterpretProgram(ByVal wantTrace As Boolean)
    Dim ws As Worksheet
    Dim hb As Workbook
    Dim running As Boolean
    ' U.29: as RunProgram's.
    Dim borrowed As VlaExcelSettings
    Dim recorded As Boolean
    Dim runProtected As Boolean
    On Error GoTo failed
    CaptureHost
    Set hb = HostBook()
    Set ws = IdeSheet()
    If ProgramIsBlank(ProgramText(ws, FIRST_ROW, IdeLastRow(ws))) Then
        VlaShowInfo "No instructions to interpret - write a program in column B of the '" & ws.Name & "' sheet first."
        Exit Sub
    End If
    ' IN.6: the interpreter has never been proven against step-tracked
    ' VLA (VerifyReportInterpreter's own IN.3 comment already names why
    ' - it always turns tracking off first) - found the hard way, live,
    ' the first time this Sub ran for real: DoCheck's translation
    ' carries EmitStmt's vla_step/vlatraceon/at-line harness by default
    ' (mStepTracking starts True), built for the emitter's OWN "Compile
    ' and Trace" and never exercised against the interpreter's
    ' dispatch, which doesn't understand it. The interpreter has its
    ' own trace mechanism (the effect log) and no use for this one, so
    ' it runs against the same clean translation VerifyReportInterpreter
    ' already proved out - On Error Resume Next here only to guarantee
    ' the True restore still runs if DoCheck itself raises, matching
    ' VerifyReportInterpreter's own shape exactly.
    ' TER-10: the line marks alone ride along (EnglishLineMarks), so a
    ' stop can name its sentence and row.
    Dim vla As String
    Dim checkOk As Boolean
    On Error Resume Next
    EnglishStepTracking False
    EnglishLineMarks True
    checkOk = DoCheck(ws, vla)
    EnglishLineMarks False
    EnglishStepTracking True
    On Error GoTo failed
    If Not checkOk Then Exit Sub
    Dim progText As String
    progText = ProgramText(ws, FIRST_ROW, IdeLastRow(ws))
    GuardTagCollision hb, ws

    ' U.19: a snapshot that cannot be made stops this run too, in words,
    ' before its first sentence - RunProgram's call says why.
    TakeRunSnapshot hb, progText, VlaIdeProgramTag(ws.Name), ws.Name

    ' U.28: the workspace sheet after the snapshot - RunProgram's says why.
    Dim outName As String
    outName = DeclaredOutputSheet(progText)
    If Len(outName) = 0 Then outName = OUT_SHEET
    Dim outWs As Worksheet
    Set outWs = GetOrCreateSheet(outName)

    ' U.29: Excel's settings recorded before the run's first change, and
    ' the program sheet protected only when it has no protection - as in
    ' RunProgram, which says why.
    VlaIdeRecordExcel borrowed
    recorded = True
    outWs.Activate
    runProtected = Not ws.ProtectContents
    If runProtected Then ws.Protect
    Application.ScreenUpdating = False
    ' IN.6: pass hb explicitly rather than relying on VlaInterpret's
    ' ActiveWorkbook default - "thisworkbook" (save-workbook-as/
    ' close-workbook) should mean the workbook this program is running
    ' in even if something mid-run changes which workbook is active.
    ' U.25: running brackets the sentences alone - a failure inside it
    ' is a stop, and the handler puts the sheets back; one before it
    ' (Check, the snapshot) changed nothing, and one after it (the event
    ' registrations) comes once every sentence has run.
    running = True
    VLA_Interpreter.VlaInterpret vla, hb
    running = False
    ' IN.7: a program declaring "When the sheet changes:" arms its
    ' handler the moment it is interpreted - re-running it (this same
    ' sub, "Interpret Instructions"/"Interpret and Trace") replaces any
    ' earlier registration for this workbook outright, so an edited
    ' handler body takes over immediately rather than the old one
    ' lingering alongside it.
    If VLA_Interpreter.VlaHasProc("on:sheet-change") Then
        VLA_Events.VlaRegisterSheetChangeHandler hb, vla
    End If
    ' IN.7 (button-click half): one registration per declared
    ' 'When "<caption>" is clicked:' handler - EnglishClickHandlerNames/
    ' Procs name whatever the DoCheck translation above just declared,
    ' the same "read it right after this run's own translation" timing
    ' VlaHasProc/VlaRegisterSheetChangeHandler already use just above.
    Dim clkNames As Collection, clkProcs As Collection
    Set clkNames = EnglishClickHandlerNames()
    Set clkProcs = EnglishClickHandlerProcs()
    Dim ci As Long
    For ci = 1 To clkNames.Count
        VLA_Events.VlaRegisterButtonClickHandler hb, CStr(clkNames.Item(ci)), vla, CStr(clkProcs.Item(ci))
    Next
    GiveBackRun ws, borrowed, runProtected
    If wantTrace Then ShowInterpreterTraceWindow
    ' VerifyReports stale-read fix - same reasoning as RunProgram's own
    ' stamp, "interpreter" instead of "emitter". Covers the button-
    ' triggered path; VerifyReportInterpreter (VLA_Tests_Host.bas)
    ' stamps this independently for its own internal interpret call.
    On Error Resume Next
    VLA_Runtime.VlaStampRunProvenance "interpreter"
    On Error GoTo 0
    Exit Sub
failed:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    ' U.29: given back before a stop's put-back and message, as in
    ' RunProgram; a failure before the record changed nothing of Excel's.
    If recorded Then
        GiveBackRun ws, borrowed, runProtected
    Else
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
    If running Then
        ReportStoppedRun hb, ws, VLA_Interpreter.VlaInterpreterLine(), "", d
    Else
        VlaShowError d
    End If
    If wantTrace Then ShowInterpreterTraceWindow
End Sub

' U.29 (the audit's C62): a Run borrows Excel's settings and gives them
' back as it found them. Before, a Run put back only screen updating, and
' always to True: a program that turned calculation off, or put words in
' the status bar, and then stopped left every open workbook in manual
' calculation, or the words showing, for the rest of the Excel session -
' and calculation is saved into any workbook saved meanwhile - while Undo
' Last Run, which puts back sheets, reached neither. Recorded before a
' Run's first change and given back on every exit, finished or stopped,
' so a Run's lasting effects are its sheets and what it saves or sends,
' never Excel's settings. The settings: every Application setting a
' phrasebook sentence, a prelude macro or the interpreter can change (the
' interpreter's DynamicSet reaches exactly calculation, cutcopymode,
' displayalerts, screenupdating and statusbar), and EnableEvents, which
' Frazaro's own sheet-change and button handlers need and a compiled
' hand-written program can turn off. check_run_gives_back.ps1 holds the
' list to the phrasebooks and the prelude. Each goes back to what it was,
' not to Excel's default: a person who works in manual calculation keeps
' it, where a default would switch every open workbook to automatic, and
' recalculate them all, at the end of every Run. (Excel sets alerts back
' itself when code finishes; they are here so one rule covers them all.)
' Cut or copy mode cannot be put back, only cancelled: it is cancelled
' when it was off before the run, since the marching ants of a paste that
' failed make the next Enter paste. Not bracketed, on purpose: the console
' and a sheet's buttons and change handlers, which nothing can undo, and
' where a setting a person gives is meant to last.
Public Sub VlaIdeRecordExcel(ByRef st As VlaExcelSettings)
    st.Calculation = Application.Calculation
    st.StatusBar = CStr(Application.StatusBar)
    st.DisplayAlerts = Application.DisplayAlerts
    st.EnableEvents = Application.EnableEvents
    st.ScreenUpdating = Application.ScreenUpdating
    st.CopyModeOff = (Application.CutCopyMode = 0)
End Sub

' U.29: put back what VlaIdeRecordExcel recorded. Calculation first, while
' the screen is still off, since going back to automatic recalculates
' every open workbook; screen updating last, so the screen is drawn once,
' as it will stay. Never raises: it runs inside a Run's failure handling.
Public Sub VlaIdeGiveBackExcel(ByRef st As VlaExcelSettings)
    On Error Resume Next
    If Application.Calculation <> st.Calculation Then Application.Calculation = st.Calculation
    If st.CopyModeOff Then Application.CutCopyMode = False
    Application.StatusBar = VlaIdeStatusBarBack(st.StatusBar)
    Application.EnableEvents = st.EnableEvents
    Application.DisplayAlerts = st.DisplayAlerts
    Application.ScreenUpdating = st.ScreenUpdating
    On Error GoTo 0
End Sub

' U.29: what the status bar is given back as. An idle status bar reads
' back as the text FALSE (OPTIMIZE's live catch, its PutStatusBarBack),
' and handing that text back would show the word: FALSE, in any case,
' gives the bar back to Excel, and other text - another add-in's words -
' is shown again. Pure, for the self-test.
Public Function VlaIdeStatusBarBack(ByVal recorded As String) As Variant
    If VLA_Identity.Fold(recorded) = "false" Then
        VlaIdeStatusBarBack = False
    Else
        VlaIdeStatusBarBack = recorded
    End If
End Function

' U.29: one exit of a Run - Excel's settings, then the program sheet's
' protection when the run made it. Never raises.
Private Sub GiveBackRun(ws As Worksheet, ByRef borrowed As VlaExcelSettings, ByVal runProtected As Boolean)
    VlaIdeGiveBackExcel borrowed
    On Error Resume Next
    If runProtected Then ws.Unprotect
    On Error GoTo 0
End Sub

' TER-10/U.25: a Run that stopped part-way. Its sheets are put back as
' they were before it (PutBackLastRun, Undo Last Run's own restore, from
' the snapshot the Run took before its first sentence), the row it
' stopped on is marked and shown, and one message says all of it. Both
' backends end here: the interpreter's handler with the line its
' at-line wrappers left, a compiled Run with the step text its program
' handed VlaReportStop. sentence may be "", and is then read from the
' row. Never raises - it runs inside a Run's failure handling, where a
' raise would reach the user as VBA's own dialog.
Private Sub ReportStoppedRun(hb As Workbook, ws As Worksheet, ByVal lineNo As Long, _
                             ByVal sentence As String, ByVal what As String)
    Dim rep As VlaUndoReport
    Dim why As String
    On Error Resume Next
    why = PutBackLastRun(hb, VlaIdeProgramTag(ws.Name), rep)
    Dim r As Long
    r = FIRST_ROW + lineNo - 1
    If lineNo > 0 And r <= IdeLastRow(ws) Then
        If Len(sentence) = 0 Then sentence = Trim$(CStr(ws.Cells(r, 2).Value))
        ws.Cells(r, 3).Value = "Stopped here: " & Replace(Replace(what, vbCrLf, " "), vbLf, " ")
        ws.Cells(r, 3).Interior.Color = RGB(247, 215, 215)
        Application.Goto ws.Cells(r, 2)
    End If
    On Error GoTo 0
    ' U.32: the roster's lines ride along, without the active sheet's -
    ' the stop goes to the row, and that decides what is in view.
    VlaShowError VlaIdeStopMessage(lineNo, sentence, what, why, rep.restoring, rep.restored, rep.removed, _
                                   VlaIdeRosterWords(rep.shown, rep.hidden, rep.moved, "", rep.missing, rep.rosterProblem))
End Sub

' TER-10/U.25: the words a stopped Run ends with. Pure, for the
' self-test. lineNo is 0 when the line is not known; why is "" when the
' sheets went back, and otherwise PutBackLastRun's reason, restoring
' naming the sheet it was on. It says what was put back and never
' claims more: a file saved or an email drafted is not in any sheet.
' U.32: rosterWords, the roster's own lines (VlaIdeRosterWords), follow
' the two lists when there are any.
Public Function VlaIdeStopMessage(ByVal lineNo As Long, ByVal sentence As String, ByVal what As String, _
                                  ByVal why As String, ByVal restoring As String, _
                                  ByVal restored As String, ByVal removed As String, _
                                  Optional ByVal rosterWords As String = "") As String
    Dim m As String
    If lineNo > 0 Then
        m = "The run stopped at line " & lineNo & ":" & vbCrLf & sentence
    Else
        m = "The run stopped."
    End If
    m = m & vbCrLf & vbCrLf & what & vbCrLf & vbCrLf
    If Len(restored) > 0 Then m = m & "Put back as they were before the run: " & restored & "." & vbCrLf
    If Len(removed) > 0 Then
        m = m & "Removed the sheet" & IIf(InStr(removed, ",") > 0, "s", "") & " the run had created: " & removed & "." & vbCrLf
    End If
    If Len(rosterWords) > 0 Then m = m & rosterWords & vbCrLf
    If Len(why) > 0 Then
        If Len(restoring) > 0 Then
            m = m & "The sheet '" & restoring & "' could not be put back (" & why & ") - Undo Last Run can try again." & vbCrLf
        Else
            m = m & "The sheets could not be put back (" & why & ") - Undo Last Run can try again." & vbCrLf
        End If
    ElseIf Len(restored) = 0 And Len(removed) = 0 And Len(rosterWords) = 0 Then
        m = m & "There was nothing to put back." & vbCrLf
    End If
    VlaIdeStopMessage = m & "Anything it did anywhere else, like a file saved or an email drafted, stays as it is."
End Function

' U.25: a compiled Run's step text is its sentence, then " [line N]"
' (ParseTracked's step table). Returns the sentence and sets lineNo;
' with no line on the end, the whole text and 0. Pure, for the self-test.
Public Function VlaIdeSplitStepText(ByVal stepText As String, ByRef lineNo As Long) As String
    lineNo = 0
    VlaIdeSplitStepText = stepText
    Dim p As Long
    p = InStrRev(stepText, " [line ")
    If p = 0 Or Right$(stepText, 1) <> "]" Then Exit Function
    Dim n As String
    n = Mid$(stepText, p + 7, Len(stepText) - p - 7)
    If Len(n) = 0 Or Len(n) > 9 Or n Like "*[!0-9]*" Then Exit Function
    lineNo = CLng(n)
    VlaIdeSplitStepText = Left$(stepText, p - 1)
End Function

' U.25: arm the stop record on the copy a compiled program calls - the
' host's Frazaro_EN_Runtime, or this project's own in the dev topology,
' ArmTrace's shape for ArmTrace's reasons.
Private Sub ArmStopReport(hb As Workbook, ByVal onOff As Boolean)
    On Error GoTo localOnly
    Application.Run "'" & hb.Name & "'!Frazaro_EN_Runtime.VlaArmStopReport", onOff
    Exit Sub
localOnly:
    VlaArmStopReport onOff
End Sub

' U.25: the stop the compiled program recorded, "" when it ran through.
Private Function ReadStopReport(hb As Workbook) As String
    On Error GoTo localOnly
    ReadStopReport = CStr(Application.Run("'" & hb.Name & "'!Frazaro_EN_Runtime.VlaStopReport"))
    Exit Function
localOnly:
    ReadStopReport = VlaStopReport()
End Function

' =====================================================================
'  The CLI (Ctrl+Shift+`): a modeless frmCLI carrying one multi-line
'  textarea, fed straight through the SAME English-or-VLA pipeline
'  InterpretProgram uses (TryTranslate -> VLA_Interpreter.VlaInterpret)
'  - one entry point for whatever a program file or english.vla would
'  also accept, just without the workspace sheet around it. Modeless on
'  purpose: the whole point of a REPL-style entry point is clicking a
'  cell to reference it while composing the next command, which a modal
'  dialog would block.
'
'  Deliberately narrower than InterpretProgram: no row marking (there is
'  no sheet backing the text to mark), no Output-sheet activation (a
'  one-off command runs against whatever sheet is already active, same
'  as "unqualified means current sheet" already means everywhere else in
'  the vocabulary), no Undo snapshot (TakeRunSnapshot is keyed to a
'  workspace sheet's own tag; a CLI command has none - known gap, not an
'  oversight). One consequence worth naming: a command that says "Work
'  on sheet Frazaro." is NOT refused here the way a workspace program's
'  own Check would refuse it (ForbiddenSheetTarget scans a sheet's
'  cells, and there is no sheet here to scan) - unlikely to type by
'  accident, but real.
' =====================================================================

Public Sub VlaOpenCli()
    frmCLI.Show vbModeless
    frmCLI.FocusBox                      ' CLI.5: and, the first time, the sample selected
End Sub

' CLI.5 (the owner's request): VlaCli - six letters for the Immediate
' window, where the CLI is opened far more often than from the ribbon or
' by Ctrl+Shift+`. An alias, not a second way in: it is VlaOpenCli. The
' name was checked free of every procedure and every module first - a
' bare name that matches a module resolves to the module ("Expected
' variable or procedure, not module"; VLA_DevRig's header records the
' two renames that cost).
Public Sub VlaCli()
    VlaOpenCli
End Sub

' OPTIMIZE.3 slice 5: Frazaro > Logic Engines > Optimize Selected Cell, and
' the legacy menu's item of the same name - the active cell's own
' =OPTIMIZE(...) call run again as a command, its answer on a new sheet
' that is left showing (VLA_Optimize.VlaOptimizeCell has the whole
' account; the sentence "Optimize cell C1." reaches the same procedure).
' A refusal is shown in words here, since nothing called it to catch one.
Public Sub VlaOptimizeSelectedCell()
    On Error GoTo failed
    If TypeName(Selection) <> "Range" Then
        VLA_Messages.RaiseMsg "optimize-command-no-cell", "what", TypeName(Selection)
    End If
    VLA_Optimize.VlaOptimizeCell ActiveCell, True
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' The Run button/Ctrl+Enter's actual work. Returns a short status line
' - CLI.5: OK, Failed or Translation failed, and the time, which ends
' the command's entry in the transcript (the dialog lost its own status
' label to it) - while full failure detail goes out through
' VlaShowError, same voice as every other refusal in this product,
' rather than shrinking it to fit a label.
' CLI.5: and now, beside the status, everything else the command
' produced, for the CLI's transcript. said is every dialog it would
' have raised - a refusal, a helper's own note - captured through
' VlaShowError's own seam (VlaMessageCapture) instead of shown; a
' program's own (msgbox ...) is not a Frazaro dialog and still appears.
' printed is what its debug-print statements printed; runValue and
' hasValue are what it came to (VLA_Interpreter.VlaInterpreterValue).
' consoleResults are the CLI's last three results, bound as *, ** and
' ***. Capture is always switched off again before this returns, so no
' later dialog anywhere in Frazaro can be swallowed by a CLI command.
Public Function VlaCliRun(ByVal text As String, Optional ByVal consoleResults As Collection = Nothing, _
                          Optional ByRef said As String, Optional ByRef printed As Collection, _
                          Optional ByRef runValue As Variant, Optional ByRef hasValue As Boolean) As String
    Dim status As String, broke As String
    Set printed = New Collection
    hasValue = False
    VLA_Runtime.VlaMessageCapture True
    On Error Resume Next
    status = CliRunCaptured(text, consoleResults, printed, runValue, hasValue)
    If Err.Number <> 0 Then
        broke = Err.Description
        status = CliStatus("Failed")
    End If
    On Error GoTo 0
    said = VLA_Runtime.VlaCapturedMessages()
    VLA_Runtime.VlaMessageCapture False
    If Len(broke) > 0 Then
        If Len(said) > 0 Then said = said & vbCrLf
        said = said & broke
    End If
    VlaCliRun = status
End Function

' CLI.5: whether the CLI hands text straight to the interpreter rather
' than to the English reader. Text whose first character, past any
' spaces, tabs or blank lines, is "(" - a form - or ";" - a VLA comment,
' like the three the CLI's sample now opens with, and a character the
' English reader refuses outright (its comments are #) - or one of the
' console's own result names alone, *, ** or ***, so that typing *
' shows the last result the way it does at every Lisp listener's
' prompt. No English sentence is a lone run of asterisks. (Before CLI.5
' a blank line above a form sent it to the English reader, which
' refused it.)
Public Function VlaCliIsVla(ByVal text As String) As Boolean
    Dim t As String
    t = Trim$(Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " "))
    VlaCliIsVla = (Left$(t, 1) = "(") Or (Left$(t, 1) = ";") Or t = "*" Or t = "**" Or t = "***"
End Function

' CLI.5: a status as the transcript writes it - how the run ended, and
' when. The time tells one run of a command from the next.
Private Function CliStatus(ByVal ending As String) As String
    CliStatus = ending & " - " & Format$(Now, "hh:mm:ss")
End Function

' CLI.5: VlaCliRun's own body from before the transcript, now run with
' every Frazaro dialog captured - it still refuses through VlaShowError,
' which lands in the transcript instead of on the screen.
Private Function CliRunCaptured(ByVal text As String, ByVal consoleResults As Collection, _
                                ByRef printed As Collection, ByRef runValue As Variant, _
                                ByRef hasValue As Boolean) As String
    Dim ran As Boolean
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Set hb = HostBook()

    ' Owner finding (live): a textarea is not a workspace-sheet row, and
    ' feeding raw multi-line VLA (a `defmacro` spanning several lines,
    ' exactly the shape english.vla/prelude.vla themselves are full of)
    ' through EnglishToVla hit its tokenizer's own real, pre-existing
    ' limit - EnTokenize's raw-VLA-line passthrough reads balanced parens
    ' but stops at the first newline regardless of depth ("Keep one
    ' complete form per row (multi-row forms wait on the Check rework,
    ' V1)"), because that reader is fundamentally married to one-row-per-
    ' cell sheet semantics. english.vla/prelude.vla are never read that
    ' way - VLA.VlaCompileToForms (VlaInterpret's own first step) is a
    ' real paren-balancing reader with no row limit at all. So: text that
    ' LOOKS like VLA (starts with a form, same test a human uses to tell
    ' the two apart) skips EnglishToVla entirely and goes straight to the
    ' interpreter, unmodified - "parsed the way english.vla is parsed",
    ' literally the same reader. Anything else is English, unchanged.
    Dim vla As String
    Dim viaEnglish As Boolean
    If VlaCliIsVla(text) Then            ' CLI.5: a form, a ; comment, or *, ** or *** alone
        vla = text
    Else
        viaEnglish = True
        IdeLoadVocab

        ' IN.6's own gotcha (InterpretProgram's identical bracket, same
        ' reason): DoCheck's translation carries EmitStmt's vla_step/
        ' vlatraceon/at-line harness by default (mStepTracking starts
        ' True) - built for the emitter's own "Compile and Trace", never
        ' exercised against the interpreter's dispatch, which evaluates a
        ' step-wrapped statement's inner form as an EXPRESSION and chokes
        ' resolving its head (e.g. 'set!') as a generic call. On Error
        ' Resume Next only to guarantee the True restore still runs if
        ' TryTranslate itself raises.
        Dim errMsg As String
        On Error Resume Next
        EnglishStepTracking False
        Dim translated As Boolean
        translated = TryTranslate(text, errMsg, vla)
        EnglishStepTracking True
        On Error GoTo failed
        If Not translated Then
            VlaShowError errMsg
            CliRunCaptured = CliStatus("Translation failed")
            Exit Function
        End If
    End If

    ran = True
    VLA_Interpreter.VlaInterpret vla, hb, consoleResults
    Set printed = VLA_Interpreter.VlaInterpreterPrinted()
    hasValue = VLA_Interpreter.VlaInterpreterValue(runValue)

    ' Same post-run registration InterpretProgram does - a CLI command
    ' declaring "When the sheet changes:" or a button-click handler
    ' should arm exactly like it would from the workspace sheet. Sheet-
    ' change detection reads real runtime proc registration either way;
    ' the click-handler names/procs below come from EnglishToVla's own
    ' per-translation state, so they only mean anything for the English
    ' branch - raw VLA declaring its own on:click:N proc by hand is a
    ' real but narrow gap, left for when a real sentence wants it.
    If VLA_Interpreter.VlaHasProc("on:sheet-change") Then
        VLA_Events.VlaRegisterSheetChangeHandler hb, vla
    End If
    If viaEnglish Then
        Dim clkNames As Collection, clkProcs As Collection
        Set clkNames = EnglishClickHandlerNames()
        Set clkProcs = EnglishClickHandlerProcs()
        Dim ci As Long
        For ci = 1 To clkNames.Count
            VLA_Events.VlaRegisterButtonClickHandler hb, CStr(clkNames.Item(ci)), vla, CStr(clkProcs.Item(ci))
        Next
    End If

    CliRunCaptured = CliStatus("OK")
    Exit Function
failed:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    ' CLI.5: what it printed before it was refused belongs in the record.
    If ran Then Set printed = VLA_Interpreter.VlaInterpreterPrinted()
    hasValue = False
    VlaShowError d
    On Error GoTo 0
    CliRunCaptured = CliStatus("Failed")
End Function

' IN.4 (part two, walking skeleton): the export mechanism, proven
' before any UX decision is made - IN.0.5's own precedent, applied
' here on purpose. The real unproven claim isn't the three open ribbon/
' dialog/run-button questions still pending owner sign-off; it's
' whether a copy of this workbook, with the runtime and a compiled
' module injected into IT rather than the live session, actually runs
' with the add-in closed. Nothing here is wired into the ribbon, the
' sheet buttons, or the legacy menu - IN.0.5's walking skeleton was not
' either, for the same reason: a throwaway measurement exercise, not a
' thing to keep polishing before the mechanism it measures is trusted.
' Deliberately minimal and hardcoded where the real version will not
' be: one fixed naming scheme (silently overwrites a same-named file
' on rerun), no save dialog, no worksheet run-button in the export -
' every one of those is cheap, reversible plumbing on TOP of a proven
' mechanism, not a reason to block proving the mechanism first.
'
' Mechanics: hb.SaveCopyAs never touches the live session (the open
' workbook's own in-memory state, saved or not, is untouched) - only
' AFTER the copy exists on disk does Workbooks.Open give a live handle
' on it, and VlaInjectRuntime/VlaCompileToModule already accept an
' arbitrary target workbook (D1's own design), so no new injection
' logic exists here at all, only a new caller of what IN.6/emitter
' already proved.
Public Sub EnglishIdeExportSkeleton()
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Dim ws As Worksheet
    Set hb = HostBook()
    Set ws = IdeSheet()
    If ProgramIsBlank(ProgramText(ws, FIRST_ROW, IdeLastRow(ws))) Then
        VlaShowInfo "No instructions to export - write a program in column B of the '" & ws.Name & "' sheet first."
        Exit Sub
    End If
    If Len(hb.Path) = 0 Then
        VlaShowInfo "Save this workbook first - export needs a real file location to save the standalone copy next to."
        Exit Sub
    End If
    If Not VlaHasVbProjectTrust() Then
        VLA_Messages.RaiseMsg "ide-export-needs-trust"
    End If
    Dim vla As String
    If Not DoCheck(ws, vla) Then Exit Sub

    ' Walking-skeleton naming: <host name> (standalone).xlsm, beside the
    ' host, always .xlsm regardless of the host's own extension - Excel
    ' silently drops VBA on save to .xlsx, and exporting FROM a
    ' no-trust .xlsx host is the flagship case (IN.9's own "email it to
    ' a colleague on a machine without the add-in").
    Dim baseName As String
    baseName = hb.Name
    Dim dotAt As Long
    dotAt = InStrRev(baseName, ".")
    If dotAt > 0 Then baseName = Left$(baseName, dotAt - 1)
    Dim exportPath As String
    exportPath = hb.Path & Application.PathSeparator & baseName & " (standalone).xlsm"

    Application.DisplayAlerts = False
    hb.SaveCopyAs exportPath
    Application.DisplayAlerts = True

    Dim copyWb As Workbook
    Set copyWb = Workbooks.Open(exportPath)
    Dim outMod As String
    outMod = VlaIdeModuleFor(ws.Name)
    VlaInjectRuntime copyWb
    VlaCompileToModule vla, outMod, copyWb
    copyWb.Save
    copyWb.Close SaveChanges:=False

    VlaShowInfo "Standalone copy exported (walking skeleton - fixed name, overwrites on rerun):" & vbCrLf & vbCrLf & exportPath & _
                 vbCrLf & vbCrLf & "To prove it really is standalone: close Excel entirely (add-in included), reopen only that file, and run its " & outMod & ".main."
    Exit Sub
failed:
    Application.DisplayAlerts = True
    VlaShowError Err.Description
End Sub

' S5.6: flip the trace on the instance the run consults - the host
' workbook's injected Frazaro_EN_Runtime. In the dev topology no
' Frazaro_EN_Runtime exists (same-project guard), so the qualified call
' fails and the local instance IS the right one; the fallback is the
' design, not an apology. Application.Run needs no VBProject trust.
Private Sub ArmTrace(hb As Workbook, ByVal onOff As Boolean)
    On Error GoTo localOnly
    Application.Run "'" & hb.Name & "'!Frazaro_EN_Runtime.VlaTrace", onOff
    Exit Sub
localOnly:
    VlaTrace onOff
End Sub

' S5.5: does the freshly compiled module actually define main? A
' program that is only comments/blank lines translates to a module
' with no runnable body, and calling its main is Excel's untrappable-
' looking "Cannot run the macro" box. Component/CodeModule access is
' the same VBProject road the compile already travelled, so it needs
' no new trust; best-effort - if the check itself cannot see the
' project, it returns True and lets the run proceed (no worse than
' before this guard existed).
Private Function ModuleHasMain(hb As Workbook, ByVal moduleName As String) As Boolean
    On Error GoTo cannotTell
    Dim cm As Object
    Set cm = hb.VBProject.VBComponents(moduleName).CodeModule
    ModuleHasMain = (cm.ProcStartLine("main", 0) > 0)   ' 0 = vbext_pk_Proc
    Exit Function
cannotTell:
    ModuleHasMain = True
End Function

' S5.3: pour the run transcript onto a "Trace" sheet and show it -
' a place a user can read it, Ctrl+A / Ctrl+C, and paste into an
' email. Reads the trace through Frazaro_EN_Runtime (the topology-
' correct copy the run actually wrote), so it works standalone too.
' Best-effort: a trace-window hiccup never turns a successful run into
' a failure.
Private Sub ShowTraceWindow(hb As Workbook)
    On Error Resume Next
    Dim report As String
    report = CStr(Application.Run("'" & hb.Name & "'!Frazaro_EN_Runtime.VlaTraceReport"))
    If Len(report) = 0 Then report = VlaTraceReport()   ' dev-topology fallback
    On Error GoTo 0
    RenderTraceReport report
End Sub

' IN.6: the interpreter's own "Interpret and Trace" - its effect log
' (IN.3.5), not the emitter's step-numbered vlatracestep/VlaTraceReport
' machinery, which only exists inside a compiled module and therefore
' means nothing on the no-VBProject path. Same window, same rendering,
' a different source of truth - reading what a program did should not
' require knowing which backend ran it. No Application.Run needed: the
' interpreter runs in-project, so its effect log is a direct call.
Private Sub ShowInterpreterTraceWindow()
    RenderTraceReport VLA_Interpreter.VlaInterpreterEffectLog()
End Sub

Private Sub RenderTraceReport(ByVal report As String)
    On Error Resume Next
    Dim tw As Worksheet
    Set tw = GetOrCreateSheet("Trace")
    tw.Cells.Clear
    tw.Columns(1).ColumnWidth = 90
    Dim lines() As String
    lines = Split(Replace(report, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        tw.Cells(i + 1, 1).Value = lines(i)
    Next
    tw.Activate
    tw.Cells(1, 1).Select
    On Error GoTo 0
End Sub

' IN.4: "Show me the VBA" - the audit story (the procedure she wrote is
' the procedure the auditor reads) must not depend on which backend
' ran. VlaTranspile already returns the generated VBA as plain text -
' no new compile machinery, no VBProject write, Mac-safe exactly like
' Check (DoCheck's own G12 note already established this same
' transpile-for-text pattern). Shows the code WITH step-tracking
' instrumentation intact, deliberately not the same EnglishStepTracking
' False bracket InterpretProgram uses - an auditor should see the real
' thing Compile would inject, not a cleaned-up stand-in for it.
Public Sub EnglishIdeShowVba()
    On Error GoTo failed
    CaptureHost
    Dim ws As Worksheet
    Set ws = IdeSheet()
    If ProgramIsBlank(ProgramText(ws, FIRST_ROW, IdeLastRow(ws))) Then
        VlaShowInfo "No instructions to show VBA for - write a program in column B of the '" & ws.Name & "' sheet first."
        Exit Sub
    End If
    Dim vla As String
    If Not DoCheck(ws, vla) Then Exit Sub
    Dim code As String
    code = VlaTranspile(vla)
    RenderVbaSource code
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' Translate to VLA/VBA: the FILE-based counterpart to Show Compiled VBA
' above - the first ribbon buttons operating on a file on disk rather
' than the IDE workspace sheet. Named "Translate", not "Export" (owner
' correction, renamed from EnglishIdeExportVla/EnglishIdeExportVba):
' this pipeline never touches workbook content - a picked SOURCE file
' becomes an OUTPUT file, purely external, unlike Export Expanded
' Vocabulary below (which genuinely exports the live, currently-loaded
' grammar state) - "Export" would misdescribe a pure file-to-file
' translation.
' Owner-reported gap, fixed: the first cut only ever showed an OPEN
' dialog (to pick the SOURCE program) and then silently auto-wrote the
' output with no dialog at all - reasonable for a workbook's own button
' wired directly to EnglishTranslateToVla/EnglishTranslateToVba (still
' true, still dialog-free), but wrong for THIS interactive ribbon flow:
' an OPEN dialog enforces the typed name already exist - typing a
' genuinely new output name (the natural thing to try) hit Windows' own
' native "file not found" refusal, not anything this code raised. Now:
' pick the SOURCE via GetOpenFilename (unchanged), THEN a real Save-As
' dialog for the OUTPUT (suggesting the same auto-derived name
' TranslatedPathFor always computed, so accepting the default
' reproduces the old behavior exactly), with an explicit overwrite
' confirmation before writing - GetSaveAsFilename only returns a chosen
' path, it does not save anything itself and does not reliably prompt
' on an existing file, so ConfirmOverwrite (below) checks Dir$ and asks
' directly.
Public Sub EnglishIdeTranslateVla()
    TranslateViaRibbon "vla"
End Sub

Public Sub EnglishIdeTranslateVba()
    TranslateViaRibbon "vba"
End Sub

Private Sub TranslateViaRibbon(ByVal kind As String)
    On Error GoTo failed
    CaptureHost

    Dim f As Variant
    f = Application.GetOpenFilename( _
        "Programs (*.txt;*.en;*.vla),*.txt;*.en;*.vla,All files (*.*),*.*", _
        , "Translate to " & UCase$(kind) & " - choose the SOURCE program")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled

    ' IdeVocabPath always returns ITS OWN best-guess candidate, real or
    ' not (by design - "for the error message", per its own comment) -
    ' every other caller checks it exists before trusting it, same here.
    ' includePolyglotta: this button needs a real file and has no
    ' embedded chain to fall back through (EDITION-VOCABPATH).
    Dim vocab As String
    vocab = IdeVocabPath(includePolyglotta:=True)
    If Not SafeFileExists(vocab) Then
        VlaShowError "Translate needs a vocabulary file on disk - Check/Compile still work from " & _
            "this add-in's own built-in copy, but Translate to VLA/VBA needs a real file to point at."
        Exit Sub
    End If

    Dim ext As String
    ext = "." & kind
    Dim outPath As Variant
    outPath = Application.GetSaveAsFilename( _
        InitialFileName:=TranslatedPathFor(CStr(f), ext), _
        FileFilter:=UCase$(kind) & " files (*" & ext & "),*" & ext & ",All files (*.*),*.*", _
        Title:="Translate to " & UCase$(kind) & " - choose where to save")
    If VarType(outPath) = vbBoolean Then Exit Sub   ' cancelled
    If Not ConfirmOverwrite(CStr(outPath), "Translate to " & UCase$(kind)) Then Exit Sub

    ' EnglishTranslateToVla/EnglishTranslateToVba are self-contained
    ' (matching EnglishRunProgram's own "safe front door" contract) - a
    ' failure already showed its own error box and returned False; only
    ' show THIS confirmation on a real success, or a failure would be
    ' followed by a bogus "Translated" message on top of its own error.
    Dim ok As Boolean
    If kind = "vla" Then
        ok = EnglishTranslateToVla(CStr(f), CStr(outPath), vocab)
    Else
        ok = EnglishTranslateToVba(CStr(f), CStr(outPath), vocab)
    End If
    If ok Then VlaShowInfo "Translated: " & CStr(outPath)
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' Owner request: VLA_Lint.VlaLintFormat already runs automatically over
' english_expanded.vla (GEXPANDERLINT.0) and over Translate to VLA/VBA's
' own output (ReformatForm/PpForm are the same pretty-printer both paths
' share) - but the two files a human actually hand-types, english.vla
' and prelude.vla, stay whatever shape their author left them in, since
' nothing ever runs the formatter over them. This closes that gap for
' ANY .vla file, not just those two by name.
' Second pass, owner request: lint a whole FOLDER at once, not just one
' file. GetOpenFilename can only ever return a file - there is no
' single native dialog that lets one control pick "a file, or a
' folder" - so this asks once, up front, which the click means, then
' routes to the matching Windows dialog type.
Public Sub EnglishIdeLintVla()
    On Error GoTo failed
    CaptureHost

    Dim wantFolder As VbMsgBoxResult
    wantFolder = MsgBox("Lint every .vla file in a folder?" & vbCrLf & vbCrLf & _
                         "Yes - pick a folder; every .vla file directly in it relints in place." & vbCrLf & _
                         "No - pick one file.", vbYesNoCancel Or vbQuestion, "Lint VLA")
    If wantFolder = vbCancel Then Exit Sub

    If wantFolder = vbYes Then
        LintVlaFolder
    Else
        LintVlaOneFile
    End If
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' The original single-file flow: Open-then-Save-As, same shape as
' TranslateViaRibbon above, with one difference on purpose - the
' Save-As dialog's InitialFileName is the SOURCE file itself, so
' accepting the default overwrites the original in place (the common
' case - lint this file), while typing a different name keeps the
' original untouched and writes the linted copy alongside it.
' ConfirmOverwrite still fires unconditionally on whatever path comes
' back, matching every other Save-As flow here - accepting the default
' means confirming the overwrite it obviously is, not skipping it.
Private Sub LintVlaOneFile()
    Dim f As Variant
    f = Application.GetOpenFilename( _
        "VLA files (*.vla),*.vla,All files (*.*),*.*", _
        , "Lint VLA - choose the file to lint")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled

    ' U.20: the whole decision is made BEFORE the Save-As dialog, so a
    ' file Lint VLA will not rewrite - generated, not UTF-8, not readable
    ' as VLA, or one whose reformatting would change more than whitespace
    ' - is refused while nothing has been chosen or touched.
    Dim src() As Byte, srcN As Long
    Dim outB() As Byte, outN As Long
    Dim at As Long, detail As String, status As String
    srcN = VLA_Loader.VlaReadFileBytes(CStr(f), src)
    status = VLA_Lint.VlaLintFileBytes(src, srcN, outB, outN, at, detail)
    If status <> "rewrite" And status <> "unchanged" Then RaiseLintRefusal status, CStr(f), at, detail

    Dim outPath As Variant
    outPath = Application.GetSaveAsFilename( _
        InitialFileName:=CStr(f), _
        FileFilter:="VLA files (*.vla),*.vla,All files (*.*),*.*", _
        Title:="Lint VLA - save linted copy as (defaults to overwriting the original)")
    If VarType(outPath) = vbBoolean Then Exit Sub   ' cancelled

    Dim inPlace As Boolean
    inPlace = (StrComp(CStr(outPath), CStr(f), vbTextCompare) = 0)
    If inPlace And status = "unchanged" Then
        VlaShowInfo "Already in house style - nothing to change: " & CStr(f)
        Exit Sub
    End If
    If Not inPlace Then RefuseGeneratedTarget CStr(outPath)
    If Not ConfirmOverwrite(CStr(outPath), "Lint VLA") Then Exit Sub

    WriteLintedFile CStr(f), src, srcN, CStr(outPath), outB, outN
    VlaShowInfo "Linted: " & CStr(outPath)
End Sub

' Owner request: relint every .vla file directly in a chosen folder,
' one call. An overwrite prompt still fires per file - owner's own
' call ("clicking multiple times in the same spot is a quick-enough
' operation for directory rewriting for now") - so this reuses the
' single-file flow's ConfirmOverwrite per name rather than a bulk
' skip-all-prompts flag. Non-recursive on purpose: only the files
' directly inside the chosen folder, matching "select A directory,"
' not a whole tree nobody asked to touch.
' Application.FileDialog is a Windows-only Office API - Mac Excel has
' no folder-picker dialog at all, so this refuses in words rather than
' failing on a missing member; the single-file flow (GetOpenFilename/
' GetSaveAsFilename, already used for Translate to VLA/VBA and Export
' Expanded Phrasebook with no Mac gating anywhere in this codebase)
' stays the Mac-safe path.
Private Sub LintVlaFolder()
    If VlaOnMac() Then
        VLA_Messages.RaiseMsg "ide-lint-folder-needs-windows"
    End If

    ' 4 = msoFileDialogFolderPicker, referenced by its literal value so
    ' this project takes on no Office object library reference for one
    ' constant - VLA_Build.bas's own vbext_ct_ClassModule literal is the
    ' same precedent.
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    fd.Title = "Lint VLA - choose a folder"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub   ' cancelled

    Dim folderPath As String
    folderPath = fd.SelectedItems(1)
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    ' Drain Dir$'s OWN internal "find next" cursor into an array FIRST.
    ' Dir$ is not reentrant - one process-wide search state - and
    ' ConfirmOverwrite below calls Dir$ itself (its own existence
    ' check); interleaving the two would silently corrupt this
    ' enumeration mid-loop, skipping or repeating names at random.
    Dim names() As String, nameCount As Long
    Dim fn As String
    fn = Dir$(folderPath & "\*.vla")
    Do While Len(fn) > 0
        nameCount = nameCount + 1
        ReDim Preserve names(1 To nameCount)
        names(nameCount) = fn
        fn = Dir$()
    Loop

    If nameCount = 0 Then
        VlaShowInfo "No .vla files found directly in " & folderPath
        Exit Sub
    End If

    ' U.20: PLAN every file before writing any - read it, lint it in
    ' memory, decide - so one dialog says what will be rewritten, what
    ' was refused and why, what is generated and never linted, and how
    ' many are already clean, BEFORE the first overwrite. A refusal in
    ' one file no longer stops the folder part-way, with earlier files
    ' already rewritten.
    Dim statuses() As String
    ReDim statuses(1 To nameCount)
    Dim plannedSrc As New Collection, plannedSrcN As New Collection
    Dim plannedOut As New Collection, plannedOutN As New Collection
    Dim src() As Byte, srcN As Long, outB() As Byte, outN As Long
    Dim at As Long, detail As String
    Dim listWrite As String, listRefused As String, listGenerated As String
    Dim nWrite As Long, nRefused As Long, nGenerated As Long, nClean As Long
    Dim i As Long, fp As String
    For i = 1 To nameCount
        fp = folderPath & "\" & names(i)
        srcN = VLA_Loader.VlaReadFileBytes(fp, src)
        statuses(i) = VLA_Lint.VlaLintFileBytes(src, srcN, outB, outN, at, detail)
        plannedSrc.Add src
        plannedSrcN.Add srcN
        If statuses(i) = "rewrite" Then
            plannedOut.Add outB
        Else
            plannedOut.Add Empty
        End If
        plannedOutN.Add outN
        Select Case statuses(i)
            Case "rewrite"
                nWrite = nWrite + 1
                listWrite = listWrite & "  " & names(i) & vbCrLf
            Case "unchanged"
                nClean = nClean + 1
            Case "generated"
                nGenerated = nGenerated + 1
                listGenerated = listGenerated & "  " & names(i) & vbCrLf
            Case Else
                nRefused = nRefused + 1
                ' The bare name, not fp: the plan already names the folder,
                ' and two full paths could push the plan past MsgBox's limit.
                listRefused = listRefused & "  " & LintRefusalText(statuses(i), names(i), at, detail) & vbCrLf
        End Select
    Next

    Dim plan As String
    plan = "Nothing has been changed yet. In " & folderPath & ":" & vbCrLf & vbCrLf
    If nWrite > 0 Then plan = plan & "To rewrite (" & nWrite & "):" & vbCrLf & listWrite & vbCrLf
    If nRefused > 0 Then plan = plan & "Refused, left alone (" & nRefused & "):" & vbCrLf & listRefused & vbCrLf
    If nGenerated > 0 Then plan = plan & "Generated, so never linted (" & nGenerated & "):" & vbCrLf & listGenerated & vbCrLf
    If nClean > 0 Then plan = plan & "Already in house style, left alone: " & nClean & vbCrLf
    plan = ClipForDialog(plan)

    If nWrite = 0 Then
        VlaShowInfo plan & vbCrLf & "Nothing to rewrite.", "Lint VLA"
        Exit Sub
    End If
    If MsgBox(plan & vbCrLf & "Rewrite the " & nWrite & " file(s) under ""To rewrite""?", _
              vbYesNo Or vbQuestion, "Lint VLA") <> vbYes Then Exit Sub

    ' The per-file overwrite prompt stays - the owner's call, recorded in
    ' this Sub's header - now only for files the plan says will change.
    Dim linted As Long, skipped As Long, nFailed As Long, listFailed As String
    Dim tb() As Byte, tbN As Long, ob() As Byte, obN As Long
    Dim errNum As Long, errText As String
    For i = 1 To nameCount
        If statuses(i) = "rewrite" Then
            fp = folderPath & "\" & names(i)
            If ConfirmOverwrite(fp, "Lint VLA - " & names(i)) Then
                tb = plannedSrc(i)
                tbN = plannedSrcN(i)
                ob = plannedOut(i)
                obN = plannedOutN(i)
                On Error Resume Next
                WriteLintedFile fp, tb, tbN, fp, ob, obN
                errNum = Err.Number
                errText = Err.Description
                On Error GoTo 0
                If errNum = 0 Then
                    linted = linted + 1
                Else
                    nFailed = nFailed + 1
                    listFailed = listFailed & "  " & errText & vbCrLf
                End If
            Else
                skipped = skipped + 1
            End If
        End If
    Next

    Dim summary As String
    summary = "Linted " & linted & " of the " & nWrite & " file(s) that needed it in " & folderPath & _
              IIf(skipped > 0, " (" & skipped & " skipped)", "")
    If nFailed > 0 Then summary = summary & vbCrLf & vbCrLf & "Not rewritten (" & nFailed & "):" & vbCrLf & listFailed
    VlaShowInfo ClipForDialog(summary)
End Sub

' U.20: the refusal, in words (SD-2), for each status
' VLA_Lint.VlaLintFileBytes returns other than "rewrite"/"unchanged".
Private Sub RaiseLintRefusal(ByVal status As String, ByVal path As String, ByVal at As Long, ByVal detail As String)
    Select Case status
        Case "generated"
            VLA_Messages.RaiseMsg "ide-lint-generated-file", "path", path
        Case "not-utf8"
            VLA_Messages.RaiseMsg "ide-lint-not-utf8", "path", path, "offset", CStr(at + 1), "line", detail
        Case "unparseable"
            VLA_Messages.RaiseMsg "ide-lint-unreadable-vla", "path", path, "reason", detail
        Case "meaning"
            VLA_Messages.RaiseMsg "ide-lint-would-change-text", "path", path, "line", CStr(at), "text", detail
    End Select
End Sub

' The same refusal's text without raising it, for the folder plan's list.
Private Function LintRefusalText(ByVal status As String, ByVal path As String, ByVal at As Long, ByVal detail As String) As String
    On Error Resume Next
    RaiseLintRefusal status, path, at, detail
    LintRefusalText = Err.Description
    Err.Clear
End Function

' U.20: Save-As can aim a linted copy at a DIFFERENT, existing file;
' when that file is a generated artifact it is refused exactly as
' linting one would be.
Private Sub RefuseGeneratedTarget(ByVal outPath As String)
    If Len(Dir$(outPath)) = 0 Then Exit Sub
    Dim b() As Byte, n As Long
    n = VLA_Loader.VlaReadFileBytes(outPath, b)
    If VLA_Lint.VlaLintIsGenerated(b, n) Then
        VLA_Messages.RaiseMsg "ide-lint-generated-file", "path", outPath
    End If
End Sub

' U.20: the only place Lint VLA writes a file, and why it cannot leave
' one broken. (1) If the source no longer holds the bytes that were
' planned - another program, or another session in the same folder,
' saved it in between - nothing is written. (2) The new bytes go to a
' sibling "<name>.lint-tmp" first and are read back; unless they match
' byte for byte, that copy is deleted and the target was never touched.
' (3) Only then is the target replaced, by a delete and a rename, so the
' worst interruption leaves the finished text in the .lint-tmp file,
' never a half-written phrasebook. (".lint-tmp" never matches the folder
' flow's "*.vla", not even through an 8.3 short name.)
Private Sub WriteLintedFile(ByVal srcPath As String, ByRef planned() As Byte, ByVal plannedN As Long, _
                            ByVal outPath As String, ByRef outB() As Byte, ByVal outN As Long)
    Dim cur() As Byte, curN As Long
    curN = VLA_Loader.VlaReadFileBytes(srcPath, cur)
    If Not BytesEqual(cur, curN, planned, plannedN) Then
        VLA_Messages.RaiseMsg "ide-lint-file-changed", "path", srcPath
    End If

    Dim tmpPath As String
    tmpPath = outPath & ".lint-tmp"
    VLA_Loader.VlaWriteFileBytes tmpPath, outB, outN
    Dim back() As Byte, backN As Long
    backN = VLA_Loader.VlaReadFileBytes(tmpPath, back)
    If Not BytesEqual(back, backN, outB, outN) Then
        Kill tmpPath
        VLA_Messages.RaiseMsg "ide-lint-write-not-verified", "path", outPath
    End If

    Dim killErr As Long
    On Error Resume Next
    Kill outPath                             ' error 53, no such file, is fine: a new Save-As name
    killErr = Err.Number
    On Error GoTo 0
    If killErr <> 0 And killErr <> 53 Then
        Kill tmpPath
        VLA_Messages.RaiseMsg "ide-lint-cannot-replace", "path", outPath
    End If
    Name tmpPath As outPath
End Sub

Private Function BytesEqual(ByRef a() As Byte, ByVal aN As Long, ByRef b() As Byte, ByVal bN As Long) As Boolean
    If aN <> bN Then Exit Function
    Dim i As Long
    For i = 0 To aN - 1
        If a(i) <> b(i) Then Exit Function
    Next
    BytesEqual = True
End Function

' MsgBox shows roughly 1,024 characters; a long plan is cut with a note
' saying so, rather than silently by Windows. 900 leaves room for the
' note and the folder flow's closing question inside that limit.
Private Function ClipForDialog(ByVal s As String) As String
    Const LIMIT As Long = 900
    If Len(s) <= LIMIT Then
        ClipForDialog = s
    Else
        ClipForDialog = Left$(s, LIMIT) & vbCrLf & "... (the rest is not shown)"
    End If
End Function

' Export Expanded Vocabulary (GEXPANDER.1): a Save-As dialog on EVERY
' click, by design - this is a deliberate, occasional, human-triggered
' action now (owner's own reversal of the original "every load writes
' one automatically" design), not something a self-test or
' VlaWriteGoldens run could ever trip over, so prompting every time is
' exactly the expected "Save As" behavior, not a UX problem the way it
' would be on an automatic path.
Public Sub EnglishIdeExportExpandedVocabulary()
    On Error GoTo failed
    CaptureHost
    Dim vocab As String
    vocab = IdeVocabPath(includePolyglotta:=True)   ' EDITION-VOCABPATH: needs a real file, has no embedded chain
    If Not SafeFileExists(vocab) Then
        VlaShowError "Export Expanded Phrasebook needs a vocabulary file on disk - none found."
        Exit Sub
    End If

    Dim f As Variant
    f = Application.GetSaveAsFilename( _
        InitialFileName:=ExpandedSiblingPath(vocab), _
        FileFilter:="VLA files (*.vla),*.vla,All files (*.*),*.*", _
        Title:="Export Expanded Phrasebook")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled
    If Not ConfirmOverwrite(CStr(f), "Export Expanded Phrasebook") Then Exit Sub

    Dim text As String
    text = EnglishExpandedVocabularyText(vocab)
    WriteTextFileVlaIde CStr(f), text
    VlaShowInfo "Exported: " & CStr(f)
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' Phrasebook Test Coverage (AS.1/RULECOVERAGE.0): same Save-As-on-every-
' click shape as Export Expanded Phrasebook, immediately above - a
' deliberate, occasional, human-triggered report, not an automatic
' side effect. EnglishRuleCoverageReport does its own fresh
' EnglishResetGrammar+EnglishLoadVocabulary, so nothing here needs to
' guess whether the currently-loaded grammar is the right one.
Public Sub EnglishIdeRuleCoverageReport()
    On Error GoTo failed
    CaptureHost
    Dim vocab As String
    vocab = IdeVocabPath(includePolyglotta:=True)   ' EDITION-VOCABPATH: needs a real file, has no embedded chain
    If Not SafeFileExists(vocab) Then
        VlaShowError "Phrasebook Test Coverage needs a vocabulary file on disk - none found."
        Exit Sub
    End If

    Dim f As Variant
    f = Application.GetSaveAsFilename( _
        InitialFileName:=CoverageReportPath(vocab), _
        FileFilter:="Text files (*.txt),*.txt,All files (*.*),*.*", _
        Title:="Phrasebook Test Coverage")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled
    If Not ConfirmOverwrite(CStr(f), "Phrasebook Test Coverage") Then Exit Sub

    Dim text As String
    text = EnglishRuleCoverageReport(vocab)
    WriteTextFileVlaIde CStr(f), text
    VlaShowInfo "Exported: " & CStr(f)
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' Owner-reported gap: GetSaveAsFilename only returns a chosen path - it
' never saves anything itself, and does not reliably show a native
' "already exists, overwrite?" prompt (empirically confirmed: it did
' not, exporting over an existing english_expanded.vla). Shared by both
' Export flows above - True to proceed (a fresh name, or the human said
' yes), False to abort silently (the human said no; a Cancel here is
' not a failure worth an error box).
Private Function ConfirmOverwrite(ByVal path As String, ByVal title As String) As Boolean
    If Len(Dir$(path)) = 0 Then
        ConfirmOverwrite = True
        Exit Function
    End If
    ConfirmOverwrite = (MsgBox(path & " already exists." & vbCrLf & "Overwrite it?", _
                                vbYesNo Or vbExclamation, title) = vbYes)
End Function

' Export Expanded Phrasebook's and Phrasebook Test Coverage's writer -
' named with an "Ide" suffix only to avoid confusion with
' VLA_SentenceEngine.bas's identically-named-in-spirit WriteTextFileVla.
' U.20: this was Open ... For Output + Print #, VBA's ANSI write, which
' sent a pound sign back to disk as the lone byte A3 (invalid UTF-8) and
' would send anything outside the system code page as "?". Now UTF-8
' with no BOM (no file this project ships carries one, and every reader
' here takes either), still ending with the one line break Print #
' appended, so an ASCII export is byte-identical to the old writer's.
' Not yet converted, and outside U.20's Lint VLA scope: WriteTextFileVla
' (Translate to VLA/VBA) and VLA_Build.bas's WriteTextFile.
Private Sub WriteTextFileVlaIde(ByVal filePath As String, ByVal content As String)
    Dim b() As Byte, n As Long
    n = VLA_Loader.VlaUtf8Encode(content & vbCrLf, b)
    VLA_Loader.VlaWriteFileBytes filePath, b, n
End Sub

' Owner request: gridlines off, so the generated code reads like code
' instead of a spreadsheet with code typed into it. Monospace and
' NumberFormat "@" for the same "text, never formulas" reason
' BuildWorkspace's own program column already gets - a generated line
' is never meant to be typed by a user, but nothing stops a VBA line
' from starting with a character Excel would otherwise try to parse.
' ActiveWindow.DisplayGridlines is a per-ACTIVE-SHEET setting in
' practice (checked, not assumed) despite living on Window in the
' object model, so this only ever affects this one sheet's own view.
Private Sub RenderVbaSource(ByVal code As String)
    Dim tw As Worksheet
    Set tw = GetOrCreateSheet(VBA_SHEET)
    tw.Cells.Clear
    tw.Columns(1).ColumnWidth = 110
    tw.Columns(1).NumberFormat = "@"
    tw.Columns(1).Font.Name = "Consolas"
    tw.Columns(1).Font.Size = 10
    Dim lines() As String
    lines = Split(Replace(code, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        tw.Cells(i + 1, 1).Value = lines(i)
    Next
    tw.Activate
    ActiveWindow.DisplayGridlines = False
    tw.Cells(1, 1).Select
End Sub

' LE.1 (thin slice): a real Excel Table - Template, Example - named
' "Phrasebook", generated from the live grammar (EnglishPhraseRows,
' VLA_English.bas), not the old single-column text dump. A ListObject
' rather than plain cells so each column gets a real filter/search
' dropdown for free (owner request) - "searchable" needs nothing
' beyond what Excel's own AutoFilter already does once this is a real
' table, not a hand-rolled search box. Rule rows carry a REAL, proof-
' verified worked example when the rule has a test: line, blank
' otherwise - never a synthesized one.
' LE.11 (owner request): the first column is each row's category - the
' word its sentence starts with, or "value" for a function word or
' phrase - where a bold header row of its own stood before, so the
' Category filter shows one word's sentences and the table carries no
' row that is not a template. The table's cells wrap and sit at the top
' of their rows, so a long template or worked example grows its row
' instead of spilling past the table's edge; the rows are fitted to
' their text, down to the old table's last row, so a row a longer
' table left tall is short again. Row 2 says what a category is.
Public Sub EnglishIdeShowPhrases()
    On Error GoTo failed
    CaptureHost
    IdeLoadVocab
    Dim ps As Worksheet
    Set ps = GetOrCreateSheet(PHRASEBOOK_SHEET)
    ' LE.11: the old table's last row, read before the table goes, since
    ' deleting it can shrink UsedRange - its rows are fitted below.
    Dim oldLast As Long
    oldLast = ps.UsedRange.Row + ps.UsedRange.Rows.Count - 1
    ' A ListObject must be torn down explicitly before Cells.Clear -
    ' clearing cells out from under a live table is the same class of
    ' refusal BuildWorkspace's own "delete before hiding" comment
    ' already documents for Buttons anchored in a column about to
    ' collapse.
    Dim oldLo As ListObject
    For Each oldLo In ps.ListObjects
        oldLo.Delete
    Next
    ps.Cells.Clear
    ps.Cells(1, 1).Value = "Frazaro " & VLA_RELEASE_VERSION & " understands:"
    ps.Cells(1, 1).Font.Bold = True
    ps.Cells(2, 1).Value = "Category is the word a sentence starts with. A value fits anywhere a value goes."
    ps.Cells(2, 1).Font.Italic = True

    Dim categories As Collection, templates As Collection, examples As Collection
    EnglishPhraseRows categories, templates, examples

    Const headerRow As Long = 3
    ps.Cells(headerRow, 1).Value = "Category"
    ps.Cells(headerRow, 2).Value = "Template"
    ps.Cells(headerRow, 3).Value = "Example"

    Dim r As Long
    r = headerRow + 1
    Dim i As Long
    For i = 1 To templates.Count
        ps.Cells(r, 1).Value = CStr(categories.Item(i))
        ps.Cells(r, 2).Value = CStr(templates.Item(i))
        ps.Cells(r, 3).Value = CStr(examples.Item(i))
        r = r + 1
    Next

    Dim tbl As ListObject
    Set tbl = ps.ListObjects.Add(xlSrcRange, ps.Range(ps.Cells(headerRow, 1), ps.Cells(r - 1, 3)), , xlYes)
    tbl.Name = "Phrasebook"

    ps.Columns(1).ColumnWidth = 14
    ps.Columns(2).ColumnWidth = 60
    ps.Columns(3).ColumnWidth = 60
    ps.Cells.WrapText = False
    tbl.Range.WrapText = True
    tbl.Range.VerticalAlignment = xlTop
    If oldLast < r - 1 Then oldLast = r - 1
    ps.Range(ps.Cells(1, 1), ps.Cells(oldLast, 1)).EntireRow.AutoFit
    ps.Activate
    ps.Rows(headerRow + 1).Select
    ActiveWindow.FreezePanes = True
    ps.Cells(1, 1).Select
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' ---------------------------------------------------------------------
'  The check: ONE translation of the whole program (V1). The old
'  fallback re-translated progressive prefixes to find the failing
'  row - up to n full translations for an n-sentence program, each
'  itself a linear rule scan, so a 500-row pilot program would make
'  Check feel broken, and slow feedback kills the check-before-run
'  habit the safety story rests on. Attribution is now a ladder over
'  the single translation:
'    1. the error's own token line (the A1 LineTag machinery - nearly
'       every user-reachable error carries one; exact knowledge);
'    2. an error message quoting exactly one row's sentence (call-
'       check errors quote the call, which may sit far from the
'       definition that exposed it);
'    3. the last row (the message text still names the problem).
'  The one refusal that used to NEED the prefix loop - assigning a
'  Defined alias, raised after parsing during sub assembly - now
'  carries its line from the English layer (V1's engine half), so
'  the ladder gives up nothing the loop had.
'  Each Check prints a timing line to the Immediate window: the
'  measured number that decides when first-word rule indexing gets
'  promoted from the backlog (measured, not assumed - the V1 deal).
' ---------------------------------------------------------------------

' S3.3: pre-parity workspaces carry the old stolen header row (A1
' guidance text, B1 "Result"). Migrate in place exactly once, keyed
' on the B1 marker: keep the ancient sheet's own C1 file-path cell,
' delete row 1 - the program and its column-B marks slide up into
' parity together - and restore the path into today's A1. Idempotent:
' after migration B1 belongs to line 1's result and never reads
' "Result" again (Check never writes that word).
' Note: this only strips the ancient header ROW: it does not also
' relocate a pre-existing program's text into today's B/C/A columns -
' that sheet generation predates every column-position change this
' function has ever needed to know about, and (per SD-doctrine
' elsewhere in this codebase) no external pilot workbook has ever
' existed to carry one forward. The path preservation below targets
' A1, today's file-reference home, so at least Reload keeps working
' for anything that does migrate through here.
Private Sub EnsureModernLayout(ws As Worksheet)
    If CStr(ws.Cells(1, 2).Value) <> "Result" Then Exit Sub
    Dim savedPath As String
    savedPath = Trim$(CStr(ws.Cells(1, 3).Value))   ' the ancient sheet's own path cell
    ws.Rows(1).Delete
    If Len(savedPath) > 0 Then ws.Cells(1, 1).Value = savedPath   ' restored to today's location, not the old one
    ws.Cells(1, 1).Font.Size = 8
    ws.Cells(1, 1).Font.Color = RGB(140, 140, 140)
End Sub

Private Function DoCheck(ws As Worksheet, Optional ByRef vlaOut As String, _
                         Optional ByVal failHint As String = "") As Boolean
    EnsureModernLayout ws
    Dim lastRow As Long
    lastRow = IdeLastRow(ws)
    ClearMarks ws, lastRow

    Dim t0 As Single, msLoad As Double
    t0 = Timer
    IdeLoadVocab
    msLoad = (Timer - t0) * 1000#

    Dim full As String
    full = ProgramText(ws, FIRST_ROW, lastRow)
    If ProgramIsBlank(full) Then
        VlaShowInfo "Nothing to check - write the program in column B of the '" & ws.Name & "' sheet."
        Exit Function
    End If

    Dim errMsg As String
    Dim ok As Boolean
    t0 = Timer
    ok = TryTranslate(full, errMsg, vlaOut)
    Debug.Print "VLA-IDE: Check timing - vocabulary " & Format$(msLoad, "0") & _
                " ms, translation " & Format$((Timer - t0) * 1000#, "0") & " ms, " & _
                (lastRow - FIRST_ROW + 1) & " row" & IIf(lastRow - FIRST_ROW + 1 = 1, "", "s")

    Dim r As Long
    If ok Then
        ' S4.3 (owner incident): a program may not work ON a Frazaro
        ' workspace or reserved sheet. "Work on sheet Frazaro." sent
        ' the snapshot machinery after the program sheet itself: a
        ' stray visible copy, no Output sheet, and an Undo dying on
        ' the wreckage with "Automation error". One Check-time gate
        ' ends the whole class - covering the declaration AND mid-
        ' program "Go to sheet" hops.
        Dim gRow As Long
        Dim gName As String
        If ForbiddenSheetTarget(ws, lastRow, gRow, gName) Then
            For r = FIRST_ROW To gRow - 1
                If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then MarkOK ws, r
            Next
            MarkErr ws, gRow, "'" & gName & "' is a Frazaro sheet - programs can't write where instructions or Frazaro's bookkeeping live"
            VlaShowError "This program tries to work on '" & gName & "', but that sheet belongs to Frazaro - it holds a program's instructions or Frazaro's own bookkeeping, and writing results there would destroy it." & vbCrLf & vbCrLf & _
                         "Results go to the Output sheet on their own, or to any sheet of yours by name, e.g.: Work on sheet Results."
            Exit Function
        End If
        ' S4.2: the resolve check - every vla-helper the generated
        ' code calls must exist (this program's own definitions, or
        ' the runtime's manifest). A dialect template's typo'd helper
        ' refuses HERE, named, at its row - not at VBA's untrappable
        ' compile modal (the S4.0/S4.1 verdict).
        Dim missing As String
        Dim missLine As Long
        missing = EnglishResolveCheck(vlaOut, missLine)
        If Len(missing) > 0 Then
            Dim mRow As Long
            mRow = missLine
            If mRow < FIRST_ROW Or mRow > lastRow Then mRow = lastRow
            For r = FIRST_ROW To mRow - 1
                If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then MarkOK ws, r
            Next
            MarkErr ws, mRow, "this instruction needs a helper named '" & missing & "' that this Frazaro doesn't provide"
            VlaShowError "This program needs a helper named '" & missing & "' that this Frazaro doesn't provide, so nothing ran." & vbCrLf & vbCrLf & _
                         "The instruction most likely comes from a vocabulary file whose template misspells the helper, or expects a newer Frazaro - whoever maintains that vocabulary will want to know."
            Exit Function
        End If
        ' G12 (folding in the PL6.2 rider): Check now TRANSPILES its
        ' own translation - pure text, Mac-safe (no VBProject touch),
        ' cheap at Check's scale. Two things this buys, both owed to
        ' earlier records: include errors (a missing library, a
        ' broken 'lib.vla line N') surface HERE at Check where they
        ' belong, not at Run; and the transpile populates the P.L6
        ' procedure-doc table, so an imported library's procedures -
        ' and the program's own - are apropos-browsable right after
        ' Check. A transpile failure is not row-attributable in
        ' general (it may live inside an included file), so the
        ' message carries the transpiler's own words - which since
        ' P.L7 name the file and line - and no row is marked wrong.
        Dim buildErr As String
        Dim ignored As String
        On Error Resume Next
        ignored = VlaTranspile(vlaOut)
        ' IN.7: VLA_ERR_INTERPRETER_ONLY is an EXPECTED refusal, not a
        ' real build problem - "When the sheet changes:" transpile-
        ' refuses by design (VLA.bas's own EmitProc guard; "When ...
        ' is clicked:"/"Make a button ..." compile for real now,
        ' button-click's own compiled-parity follow-up), on purpose,
        ' so Compile/Compile and Trace fail with words instead of
        ' emitting VBA that would hit an untrappable compile modal.
        ' Check's own transpile probe exists to catch REAL build
        ' problems (include errors, missing helpers) - this one is
        ' neither, so it must not fail CHECK too, or the already-
        ' shipped, live-verified sheet-change half would stop passing
        ' Check the moment this guard shipped.
        If Err.Number <> 0 And Err.Number <> VLA_ERR_INTERPRETER_ONLY Then buildErr = Err.Description
        On Error GoTo 0
        If Len(buildErr) > 0 Then
            VlaShowError "The program translated, but its code could not be built:" & vbCrLf & vbCrLf & _
                         buildErr & vbCrLf & vbCrLf & _
                         "If the message names a library file, the problem lives in that file, not in these instructions."
            Exit Function
        End If
        For r = FIRST_ROW To lastRow
            If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then MarkOK ws, r
        Next
        ' LX.14 (call 2): the translation's notes, each on its own line's
        ' row, beside OK - a program's own word masking a phrasebook's.
        Dim note As Variant, bar As Long, noteRow As Long
        For Each note In EnglishLastNotes()
            bar = InStr(CStr(note), "|")
            If bar > 1 Then
                noteRow = FIRST_ROW + CLng(Left$(CStr(note), bar - 1)) - 1
                If noteRow >= FIRST_ROW And noteRow <= lastRow Then
                    VlaIdeMarkNote ws, noteRow, Mid$(CStr(note), bar + 1)
                End If
            End If
        Next
        DoCheck = True
        Exit Function
    End If

    ' The attribution ladder (see the block comment above).
    Dim badRow As Long
    Dim errLine As Long
    errLine = EnglishLastErrorLine()
    If errLine > 0 Then
        badRow = FIRST_ROW + errLine - 1
        If badRow > lastRow Then badRow = lastRow
        If badRow < FIRST_ROW Then badRow = FIRST_ROW
    Else
        badRow = RowFromMessage(ws, lastRow, errMsg)
        If badRow = 0 Then badRow = lastRow
    End If

    For r = FIRST_ROW To badRow - 1
        If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then MarkOK ws, r
    Next
    MarkErr ws, badRow, errMsg
    ' SOP.6: ImportFromPath hands in a hint when the document it just
    ' poured had no <Frazaro> tag. It joins the red row's own text only
    ' after MarkErr has logged the failure, so VLA_Log keeps just what
    ' Frazaro could not read.
    If Len(failHint) > 0 Then ws.Cells(badRow, 3).Value = VlaIdeCellText(CStr(ws.Cells(badRow, 3).Value) & " " & failHint)
End Function

' S4.3: does any sentence aim the program at a sheet Frazaro owns?
' Scans the sheet rows directly (row IS line since S3.3), so the
' refusal marks the exact offending row. Names are compared after
' stripping the trailing period and optional quotes, case-blind.
Private Function ForbiddenSheetTarget(ws As Worksheet, ByVal lastRow As Long, _
                                      ByRef rowOut As Long, ByRef nameOut As String) As Boolean
    Dim hb As Workbook
    Set hb = HostBook()
    Dim r As Long
    For r = FIRST_ROW To lastRow
        Dim raw As String
        raw = Trim$(CStr(ws.Cells(r, 2).Value))
        Dim t As String
        t = VLA_Identity.Fold(raw)
        Dim nm As String
        nm = ""
        If Left$(t, 14) = "work on sheet " Then
            nm = Mid$(raw, 15)
        ElseIf Left$(t, 12) = "go to sheet " Then
            nm = Mid$(raw, 13)
        End If
        nm = Trim$(nm)
        If Len(nm) > 0 Then
            If Right$(nm, 1) = "." Then nm = Trim$(Left$(nm, Len(nm) - 1))
            If Len(nm) > 1 Then
                If Left$(nm, 1) = "'" And Right$(nm, 1) = "'" Then nm = Mid$(nm, 2, Len(nm) - 2)
            End If
            If Len(nm) > 1 Then
                If Left$(nm, 1) = """" And Right$(nm, 1) = """" Then nm = Mid$(nm, 2, Len(nm) - 2)
            End If
            If Len(nm) > 0 And IsFrazaroSheetName(hb, nm) Then
                rowOut = r
                nameOut = nm
                ForbiddenSheetTarget = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function IsFrazaroSheetName(hb As Workbook, ByVal nm As String) As Boolean
    Dim k As String
    k = VLA_Identity.Fold(nm)
    If Left$(k, Len(UNDO_PREFIX)) = VLA_Identity.Fold(UNDO_PREFIX) Then IsFrazaroSheetName = True: Exit Function
    If Left$(k, Len(DEL_PREFIX)) = VLA_Identity.Fold(DEL_PREFIX) Then IsFrazaroSheetName = True: Exit Function
    ' U.32: the staging and the roster are Frazaro's too (the runtime's
    ' twin, VlaIsFrazaroSheetName, has had the staging since L-SHEET-HELPERS).
    If Left$(k, Len(STAGE_PREFIX)) = VLA_Identity.Fold(STAGE_PREFIX) Then IsFrazaroSheetName = True: Exit Function
    If Left$(k, Len(ROSTER_PREFIX)) = VLA_Identity.Fold(ROSTER_PREFIX) Then IsFrazaroSheetName = True: Exit Function
    ' IN.4: VBA_SHEET added alongside the other Frazaro-owned display
    ' sheets. "Trace" (ShowTraceWindow) has this same exposure and
    ' always has - a pre-existing gap, not introduced here and not
    ' fixed here, since fixing it only for the sheet this pass happens
    ' to add would be an arbitrary asymmetry, not a real fix.
    If k = VLA_Identity.Fold(LOG_SHEET) Or k = VLA_Identity.Fold(PHRASEBOOK_SHEET) _
        Or k = VLA_Identity.Fold(VBA_SHEET) Or k = "feedback" Then IsFrazaroSheetName = True: Exit Function
    Dim wsx As Variant
    For Each wsx In WorkspaceSheets(hb)
        If VLA_Identity.Fold(wsx.Name) = k Then IsFrazaroSheetName = True: Exit Function
    Next
End Function

' V1: the translated VLA rides out through the optional vla parameter,
' so the Run path can compile what Check already translated instead of
' translating a second time.
Private Function TryTranslate(ByVal text As String, ByRef errMsg As String, _
                              Optional ByRef vla As String) As Boolean
    On Error GoTo bad
    vla = EnglishToVla(text)
    TryTranslate = True
    Exit Function
bad:
    errMsg = Err.Description
    vla = ""
End Function

' DI2.1: an external phrasebook still wins when present, in the exact
' same four-candidate order IdeVocabPath already searched - unchanged
' from before this pass, so a user's or org's own phrasebook always
' overrides the built-in one (now matched by this EDITION's own
' filename - IdeVocabFileName - rather than always "english.vla").
' Only when no external candidate exists does this fall back to the
' add-in's own embedded CHAIN: VLAe_Source, then VLAe_Source2,
' VLAe_Source3, ... in order, stopping at the first sheet that doesn't
' exist (VLA.VlaEmbeddedText's own documented "" for a missing sheet).
' EDITIONMANIFEST.1: an edition's own overlay phrasebook (espanol.vla
' on top of english.vla, say) embeds as a second sheet in this chain -
' the identical "one EnglishLoadVocabulary call per file" shape that
' file's own header already documents as ordinary multi-file loading,
' just read back from sheets instead of disk. The English edition
' still embeds exactly one sheet, so this loop makes exactly one call,
' same as before this pass.
Private Sub IdeLoadVocab()
    EnglishResetGrammar
    Dim vocabPathFound As String
    vocabPathFound = IdeVocabPath()
    If SafeFileExists(vocabPathFound) Then
        EnglishLoadVocabulary vocabPathFound
    Else
        Dim anyLoaded As Boolean
        Dim n As Long
        For n = 1 To 8   ' generous ceiling - no edition needs more than a couple of overlay files today
            Dim sheetName As String
            sheetName = "VLAe_Source"
            If n > 1 Then sheetName = sheetName & CStr(n)
            Dim embedded As String
            embedded = VLA.VlaEmbeddedText(sheetName)
            If Len(embedded) = 0 Then Exit For
            EnglishLoadVocabularyText embedded, "built-in vocabulary" & IIf(n > 1, " (" & n & ")", "")
            anyLoaded = True
        Next n
        If Not anyLoaded Then
            ' EDITIONMANIFEST.5: true last resort, AFTER the embedded
            ' chain has already come up empty - a dev workbook has no
            ' embedded sheets at all, so this is what actually finds
            ' english.vla for it now that the file lives in
            ' scripts\polyglotta\. Deliberately never checked ahead of
            ' the embedded chain (see IdeVocabPath's own header) - a
            ' BUILT edition's .xlam sits beside the dev repo too, and
            ' this path would otherwise "find" the repo's own source
            ' file and silently skip the edition's real, embedded,
            ' correctly-chained phrasebook.
            Dim devFallback As String
            devFallback = IdeDevPolyglottaPath()
            If SafeFileExists(devFallback) Then
                EnglishLoadVocabulary devFallback
                anyLoaded = True
            End If
        End If
        If Not anyLoaded Then
            VLA_Messages.RaiseMsg "ide-vocab-not-found", "path", IdeVocabPath(mayPrompt:=False)
        End If
    End If
    ReplayPersistedPhrasebooks   ' GO.6: every user-loaded phrasebook, ADDED on top
    Debug.Print EnglishLintReport()
    Debug.Print EnglishVocabStats()       ' U.10: the counters line
End Sub

' ---------------------------------------------------------------------
'  Sheet plumbing
' ---------------------------------------------------------------------

' V2: resolve WHICH program a command means. Priority: (1) the sheet
' the command was invoked on, when it is a workspace - a button click
' is its sheet, which is what clicking there means (the D1 capture
' doctrine extended one level down); (2) the workbook's only
' workspace, when there is exactly one - the pre-V2 behavior, so
' single-program workbooks feel no change; (3) otherwise refuse,
' naming the programs, because guessing between programs is exactly
' the ambiguity-with-a-dangerous-branch shape.
Private Function IdeSheet() As Worksheet
    MigrateLegacySheet
    Dim hb As Workbook
    Set hb = HostBook()
    If Not mHostSheet Is Nothing Then
        Dim an As String
        an = ""
        On Error Resume Next
        an = mHostSheet.Name
        On Error GoTo 0
        If VlaIdeIsWorkspaceName(an) Then
            ' Fetch by name through Worksheets: a CHART sheet named
            ' like a workspace comes back Nothing and falls through.
            On Error Resume Next
            Set IdeSheet = hb.Worksheets(an)
            On Error GoTo 0
            If Not IdeSheet Is Nothing Then Exit Function
        End If
    End If
    Dim list As Collection
    Set list = WorkspaceSheets(hb)
    If list.Count = 1 Then
        Set IdeSheet = list.Item(1)
        Exit Function
    End If
    If list.Count = 0 Then
        VLA_Messages.RaiseMsg "ide-no-workspace-sheet", "sheet", IDE_SHEET
    End If
    VLA_Messages.RaiseMsg "ide-workspace-ambiguous", "count", list.Count, "names", WorkspaceNamesOf(list)
End Function

' All workspace sheets in the book, in tab order.
Private Function WorkspaceSheets(hb As Workbook) As Collection
    Dim r As New Collection
    Dim i As Long
    For i = 1 To hb.Worksheets.Count
        If VlaIdeIsWorkspaceName(hb.Worksheets(i).Name) Then r.Add hb.Worksheets(i)
    Next
    Set WorkspaceSheets = r
End Function

Private Function WorkspaceNamesOf(list As Collection) As String
    Dim r As String
    Dim wsV As Variant
    For Each wsV In list
        If Len(r) > 0 Then r = r & ", "
        r = r & wsV.Name
    Next
    WorkspaceNamesOf = r
End Function

' =====================================================================
'  V2: program identity, as three PURE name functions - public so the
'  self-test can pin them without a workbook in hand.
'
'  A workspace sheet is "Frazaro" (the default program) or
'  "Frazaro (<name>)" (a named one). Each program derives a short
'  TAG - "Main" for the default, else the name reduced to letters
'  and digits, capped at 10 - which scopes everything per-program:
'  the module is Frazaro_EN_Sheet (default) or Frazaro_EN_<tag>, and
'  snapshots wear VLAu_<tag>_/VLAd_<tag>_ prefixes.
'  Tags contain no underscore BY CONSTRUCTION, so the first "_"
'  after the prefix always ends the tag - the parse cannot be
'  ambiguous. Two programs whose names reduce to one tag ("Q1 Data"
'  and "Q1-Data") are refused loudly at Run/Undo/Add rather than
'  silently sharing a module and an undo history.
' =====================================================================

Public Function VlaIdeIsWorkspaceName(ByVal sheetName As String) As Boolean
    If VLA_Identity.Fold(sheetName) = VLA_Identity.Fold(IDE_SHEET) Then
        VlaIdeIsWorkspaceName = True
        Exit Function
    End If
    If Len(sheetName) > Len(IDE_SHEET) + 3 Then
        If VLA_Identity.Fold(Left$(sheetName, Len(IDE_SHEET) + 2)) = VLA_Identity.Fold(IDE_SHEET) & " (" _
           And Right$(sheetName, 1) = ")" Then
            VlaIdeIsWorkspaceName = True
        End If
    End If
End Function

Public Function VlaIdeProgramTag(ByVal sheetName As String) As String
    If Not VlaIdeIsWorkspaceName(sheetName) Then Exit Function
    If VLA_Identity.Fold(sheetName) = VLA_Identity.Fold(IDE_SHEET) Then
        VlaIdeProgramTag = "Main"
        Exit Function
    End If
    Dim inner As String
    inner = Mid$(sheetName, Len(IDE_SHEET) + 3)
    inner = Left$(inner, Len(inner) - 1)
    Dim i As Long
    Dim c As String
    Dim r As String
    For i = 1 To Len(inner)
        c = Mid$(inner, i, 1)
        If c Like "[A-Za-z0-9]" Then r = r & c
    Next
    If Len(r) = 0 Then r = "Prog"    ' degenerate hand-made names;
                                     ' the collision guard still
                                     ' catches two of them
    If Len(r) > 10 Then r = Left$(r, 10)
    VlaIdeProgramTag = r
End Function

Public Function VlaIdeModuleFor(ByVal sheetName As String) As String
    Dim tag As String
    tag = VlaIdeProgramTag(sheetName)
    If tag = "Main" Or Len(tag) = 0 Then
        VlaIdeModuleFor = OUT_MODULE     ' the default program keeps
                                         ' OUT_MODULE - one constant,
                                         ' one place to rename it again
    Else
        VlaIdeModuleFor = "Frazaro_EN_" & tag
    End If
End Function

' Refuse when another workspace's tag matches this one - shared tag
' would mean a shared module and a shared undo history, silently.
Private Sub GuardTagCollision(hb As Workbook, ws As Worksheet)
    Dim tag As String
    tag = VlaIdeProgramTag(ws.Name)
    Dim wsV As Variant
    For Each wsV In WorkspaceSheets(hb)
        If Not wsV Is ws Then
            If VlaIdeTagsMatch(VlaIdeProgramTag(wsV.Name), tag) Then
                VLA_Messages.RaiseMsg "ide-programs-share-short-name", "a", ws.Name, "b", wsV.Name, "tag", tag
            End If
        End If
    Next
End Sub

' The workspace sheet was named "English" before the product took its
' own name. An existing legacy workspace is renamed in place - the
' program text and buttons ride along with the sheet - so old
' workbooks keep working untouched. Only a sheet carrying the
' workspace buttons is migrated: a data sheet that merely happens to
' be named "English" is left alone. Soft-failure doctrine applies -
' any hiccup here degrades to the normal missing-sheet story.
Private Sub MigrateLegacySheet()
    On Error GoTo softly
    Dim hb As Workbook
    Set hb = HostBook()
    Dim cur As Worksheet, leg As Worksheet
    On Error Resume Next
    Set cur = hb.Worksheets(IDE_SHEET)
    Set leg = hb.Worksheets(IDE_SHEET_LEGACY)
    On Error GoTo softly
    If Not cur Is Nothing Then Exit Sub       ' already named, or both exist: hands off
    If leg Is Nothing Then Exit Sub           ' nothing to migrate
    Dim b As Button
    For Each b In leg.Buttons
        If b.Caption = "Check Sentences" Then
            leg.Name = IDE_SHEET
            Exit Sub
        End If
    Next
softly:
End Sub

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim hb As Workbook
    Set hb = HostBook()
    On Error Resume Next
    Set GetOrCreateSheet = hb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = hb.Worksheets.Add( _
            After:=hb.Worksheets(hb.Worksheets.Count))
        GetOrCreateSheet.Name = name
    End If
End Function

Private Function IdeLastRow(ws As Worksheet) As Long
    IdeLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
End Function

' Rows joined with line breaks: a blank row becomes a blank line,
' which is a paragraph break, which closes blocks - the grid and the
' language share one structure.
Private Function ProgramText(ws As Worksheet, ByVal fromRow As Long, ByVal toRow As Long) As String
    Dim r As Long, s As String
    For r = fromRow To toRow
        s = s & CStr(ws.Cells(r, 2).Value) & vbCrLf
    Next
    ProgramText = s
End Function

' TER-2's live pass: True when a program's text is nothing but spaces,
' tabs and line breaks. ProgramText ends EVERY row with a line break, an
' empty row included, and Trim$ strips only spaces - so the guard that
' used to read Len(Trim$(text)) = 0 could never be true. Clearing a whole
' program and clicking Check sent a bare line break to the translator,
' which came back as "Subscript out of range" marked on row 1; Compile,
' Interpret, Export and Show VBA carried the same dead guard.
Private Function ProgramIsBlank(ByVal text As String) As Boolean
    text = Replace(Replace(Replace(text, vbCr, ""), vbLf, ""), vbTab, "")
    ProgramIsBlank = (Len(Trim$(text)) = 0)
End Function

' TER-2: clear down to whichever is further, the program (lastRow, read
' from column B) or the marks themselves (column C). Callers compute
' lastRow AFTER the user's edit, so when the last sentence's text was
' deleted - or the whole program - its old mark sat below lastRow and
' survived the Check. Column C holds nothing but these marks
' (BuildWorkspace formats the whole column for them), so clearing to its
' own last value is safe. PourProgram used to work this out inline for
' itself; every caller gets it here now.
Private Sub ClearMarks(ws As Worksheet, ByVal lastRow As Long)
    Dim lastMark As Long
    lastMark = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastMark > lastRow Then lastRow = lastMark
    If lastRow < FIRST_ROW Then Exit Sub
    With ws.Range(ws.Cells(FIRST_ROW, 3), ws.Cells(lastRow, 3))
        .ClearContents
        .Interior.Color = RGB(242, 242, 242)   ' BuildWorkspace's own result-column gray, not colorless
    End With
End Sub

Private Sub MarkOK(ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 3).Value = "OK"
    ws.Cells(r, 3).Interior.Color = RGB(221, 235, 221)
End Sub

Private Sub MarkErr(ws As Worksheet, ByVal r As Long, ByVal msg As String)
    ws.Cells(r, 3).Value = VlaIdeCellText(msg)
    ws.Cells(r, 3).Interior.Color = RGB(247, 215, 215)
    LogParseFailure ws, r, msg
End Sub

' LX.14 (call 2): a note on a row that checked, in yellow, in place of its
' OK - the one the translation leaves when a program's own "To <word> of
' ...:" masks a phrasebook's word. Nothing is refused and nothing is
' logged: the row is OK, and the note says what it means.
Public Sub VlaIdeMarkNote(ws As Worksheet, ByVal r As Long, ByVal note As String)
    ws.Cells(r, 3).Value = VlaIdeCellText(note)
    ws.Cells(r, 3).Interior.Color = RGB(255, 242, 204)
End Sub

' U.26: text to put in a cell so the cell shows exactly that text. Excel
' reads a leading apostrophe as its own mark for "this is text" and hides
' it, so a refusal that opens by quoting a word - 'a2' is shaped like a
' cell..., 'seek' is a reserved word... - showed without its first
' character; and a leading = + - or @ would be read as a formula. One
' apostrophe in front takes Excel's mark, and the text shows whole.
Public Function VlaIdeCellText(ByVal s As String) As String
    Select Case Left$(s, 1)
        Case "'", "=", "+", "-", "@"
            VlaIdeCellText = "'" & s
        Case Else
            VlaIdeCellText = s
    End Select
End Function

' Every misunderstood sentence is remembered: timestamp, row, the
' sentence, the message - on a very-hidden sheet in the user's own
' workbook. Local only; nothing is sent anywhere. Best-effort by
' doctrine: a logging problem must never break a Check.
Private Sub LogParseFailure(ws As Worksheet, ByVal r As Long, ByVal msg As String)
    On Error Resume Next
    Dim hb As Workbook
    Set hb = ws.Parent
    Dim lg As Worksheet
    Set lg = Nothing
    Set lg = hb.Worksheets(LOG_SHEET)
    If lg Is Nothing Then
        Set lg = hb.Worksheets.Add(After:=hb.Worksheets(hb.Worksheets.Count))
        lg.Name = LOG_SHEET
        lg.Cells(1, 1).Value = "when"
        lg.Cells(1, 2).Value = "row"
        lg.Cells(1, 3).Value = "instruction"
        lg.Cells(1, 4).Value = "problem"
        lg.Rows(1).Font.Bold = True
        lg.Columns(3).NumberFormat = "@"
    End If
    lg.Visible = xlSheetVeryHidden
    Dim nr As Long
    nr = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row + 1
    If nr < 2 Then nr = 2
    lg.Cells(nr, 1).Value = Now
    lg.Cells(nr, 2).Value = r
    lg.Cells(nr, 3).Value = VlaIdeCellText(CStr(ws.Cells(r, 2).Value))
    lg.Cells(nr, 4).Value = VlaIdeCellText(msg)       ' U.26: the report reads back what was said
    If nr > 2000 Then lg.Rows("2:1001").Delete      ' keep the newest thousand
    On Error GoTo 0
End Sub

' If the error message contains exactly one row's text, use that row
' (e.g. an error quoting the offending sentence, which may sit far
' from where the problem surfaced). Since V1 this is the ladder's
' second rung, consulted only when the error carried no token line.
Private Function RowFromMessage(ws As Worksheet, ByVal lastRow As Long, ByVal msg As String) As Long
    Dim m As String
    m = NormText(msg)
    Dim r As Long, hits As Long, hitRow As Long
    Dim t As String
    For r = FIRST_ROW To lastRow
        t = NormText(CStr(ws.Cells(r, 2).Value))
        If Len(t) >= 6 Then
            If InStr(m, t) > 0 Then
                hits = hits + 1
                hitRow = r
            End If
        End If
    Next
    If hits = 1 Then RowFromMessage = hitRow
End Function

Private Function NormText(ByVal s As String) As String
    s = VLA_Identity.Fold(s)
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormText = Trim$(s)
End Function

' =====================================================================
'  D1 recovery-path truth table (audited with the capture pass; every
'  row re-checked against the captured-host rule and the soft-failure
'  doctrine - auxiliary machinery fails softer than what it assists):
'
'  path                        on failure                      verdict
'  Word reader (import)        doc closed, a Word WE started   sound
'                              is quit, a Word the user had
'                              open is left alone; wrapped
'                              message names Word
'  Reload after file moved     friendly "has moved" message;   fixed:
'                              Dir$ raising on a dead drive    was raw
'                              now takes the same road         error
'  Snapshot failure            Run proceeds, note in the       sound
'                              Immediate window, Undo simply
'                              unavailable for that run
'  Undo mid-restore failure    transactional per sheet (set    sound
'                              aside, copy, only then delete);
'                              message names the sheet it was
'                              restoring
'  Long sheet name             snapshot skipped with a named   sound
'                              note; other sheets still
'                              covered
'  Run failure                 program sheet unprotected in    sound
'                              the handler via the CAPTURED
'                              sheet object, so a mid-run
'                              focus change cannot misdirect
'                              the unprotect - U.29: only when
'                              the run protected it
'  Excel's settings (U.29)     recorded before a Run's first   sound
'                              change, given back on every
'                              exit, before a stop's put-back
'  Run stop (U.25)             the snapshot is put back before sound
'                              anything is said, the row it
'                              stopped on is marked; a put-back
'                              that fails names its sheet and
'                              Undo Last Run, which can retry
'  Parse-failure logging       best-effort throughout; a log   sound
'                              problem never breaks a Check
'                              (book derived from ws.Parent -
'                              the object-based pattern D1
'                              generalizes)
'
'  Undo for runs. Before a program executes, the sheets it can touch
'  (Output, plus any sheet named after the word "sheet" in the program
'  text) are copied to very-hidden snapshots. "Undo Last Run" copies
'  them back. Sheets with names too long to snapshot losslessly are
'  skipped with a note in the Immediate window. U.19: a snapshot is all
'  or nothing - if any copy or marker cannot be made, everything made so
'  far is removed and the Run stops, in words, before its first sentence.
'  A sheet that does not exist yet gets a removal marker instead of a
'  copy, so a sheet the run creates is removed again - its workspace
'  sheet too, since U.28 made that sheet after the snapshot.
' =====================================================================

' V2: snapshot names carry the program's tag - VLAu_<tag>_<sheet> -
' so each program's Undo sees only its own history. Tags contain no
' underscore by construction, so the name never parses ambiguously.
'
' U.19: all or nothing. Every sheet this makes - a copy or a tombstone
' marker - joins `made` the moment it exists, BEFORE its rename and its
' hide, so a failure at any later step removes every one of them: a
' visible "<sheet> (2)" can no longer outlive a failed snapshot (G-PROLOG
' slice 1's stray "gp1 (2)"). The failure then refuses in words through
' ide-undo-snapshot-failed, and the Run stops before its first sentence -
' RunProgram's call says why it stops rather than going on without Undo.
' The copy is found by what is NEW among the sheet names
' (VlaIdeAddedSheetNames), never by position. Measured live (U.19's M1):
' a Copy made After a very-hidden last sheet lands BEFORE it, so the old
' Worksheets(Count) read renamed that hidden sheet - the Output tombstone
' made a moment earlier - and raised nothing, leaving the real copy
' visible as "<sheet> (2)" and a blank snapshot for Undo to put back. A scanned name no sheet can have is not a target at all
' (VlaIdeCanBeSheetName): its tombstone's rename was certain to fail.
Private Sub TakeRunSnapshot(hb As Workbook, ByVal programText As String, ByVal tag As String, ByVal programName As String)
    Dim scr As Boolean, da As Boolean
    scr = Application.ScreenUpdating
    da = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim made As Collection
    Set made = New Collection
    Dim curTarget As String
    Dim failWhy As String
    Dim staged As Collection
    Set staged = New Collection
    Dim finals As Collection
    Set finals = New Collection
    On Error GoTo failed

    ' U.23: this first sweep keeps the previous Run's Undo - only orphans
    ' and this program's own interrupted staging go. The previous copies
    ' are replaced below, once every new one exists, so a Run refused on
    ' the way leaves Undo Last Run exactly as it was.
    DeleteSnapshots hb, tag, False, made
    Dim uPre As String, dPre As String, sPre As String, rPre As String
    uPre = UNDO_PREFIX & tag & "_"
    dPre = DEL_PREFIX & tag & "_"
    sPre = STAGE_PREFIX & tag & "_"
    rPre = ROSTER_PREFIX & tag & "_"

    ' U.32: the roster, read before anything is made (Worksheets.Add and
    ' Copy activate what they make) and written to the Run's own marker
    ' sheet, staged and swapped like every other snapshot sheet, so a Run
    ' refused on the way keeps the previous roster with the previous
    ' copies. Made first, so a workbook that refuses a new sheet refuses
    ' here, before a copy exists.
    Dim rosterRows As Collection
    Dim rosterActive As String
    Set rosterRows = ReadRosterRows(hb, rosterActive)
    Dim snap As Worksheet
    curTarget = rPre & ROSTER_WORD
    Set snap = hb.Worksheets.Add(After:=hb.Worksheets(hb.Worksheets.Count))
    made.Add snap
    snap.Name = sPre & ROSTER_WORD
    WriteRosterRows snap, rosterRows, rosterActive
    snap.Visible = xlSheetVeryHidden
    staged.Add snap
    finals.Add rPre & ROSTER_WORD

    Dim targets As New Collection
    AddTarget targets, OUT_SHEET
    ScanSheetNames programText, targets

    Dim n As Variant
    Dim priorNames As Collection
    Dim added As Collection
    Dim addedV As Variant
    For Each n In targets
        curTarget = CStr(n)
        If Not VlaIdeCanBeSheetName(curTarget) Then
            ' U.19: the scan reads quoted text too, so it can offer a
            ' name like Q1/Q2 that no sheet can have - no Run can reach
            ' or create that sheet, so there is nothing to put back.
        ElseIf SheetExists(hb, curTarget) Then
            If Len(curTarget) + Len(uPre) <= 31 Then
                Set priorNames = SheetNamesOf(hb)
                hb.Worksheets(curTarget).Copy After:=hb.Worksheets(hb.Worksheets.Count)
                Set added = VlaIdeAddedSheetNames(priorNames, SheetNamesOf(hb))
                For Each addedV In added
                    made.Add hb.Worksheets(CStr(addedV))
                Next
                If added.Count <> 1 Then
                    VLA_Messages.RaiseMsg "ide-undo-snapshot-copy-not-found", "count", CStr(added.Count)
                End If
                Set snap = hb.Worksheets(CStr(added.Item(1)))
                snap.Name = sPre & curTarget
                snap.Visible = xlSheetVeryHidden
                staged.Add snap
                finals.Add uPre & curTarget
            Else
                Debug.Print "VLA-IDE: sheet name too long to snapshot: " & n
            End If
        Else
            ' D1.2 (first-user report): this named target does not
            ' exist yet, so the run may CREATE it - and "the state
            ' before the last run" is a state where it did not exist.
            ' Leave a tombstone marker; Undo removes the sheet and
            ' the marker together. (Before this, Undo restored every
            ' snapshot but left run-created sheets standing - the
            ' documented gap a real Undo click finally raised.)
            If Len(curTarget) + Len(dPre) <= 31 Then
                Set snap = hb.Worksheets.Add(After:=hb.Worksheets(hb.Worksheets.Count))
                made.Add snap
                snap.Name = sPre & curTarget
                snap.Visible = xlSheetVeryHidden
                staged.Add snap
                finals.Add dPre & curTarget
            Else
                Debug.Print "VLA-IDE: sheet name too long for an undo marker: " & n
            End If
        End If
    Next

    ' U.23: the swap. A final name no worksheet holds is taken at once -
    ' so a chart sheet sitting on it is refused HERE, while the previous
    ' Undo is still whole. A name the previous Run's copy holds waits
    ' until the second sweep below has removed that copy. The three
    ' prefixes are one length, so the target is the same slice of every
    ' final name.
    Dim deferred As Collection
    Set deferred = New Collection
    Dim k As Long
    Dim idxV As Variant
    For k = 1 To staged.Count
        curTarget = Mid$(CStr(finals.Item(k)), Len(uPre) + 1)
        If SheetExists(hb, CStr(finals.Item(k))) Then
            deferred.Add k
        Else
            Set snap = staged.Item(k)
            snap.Name = CStr(finals.Item(k))
        End If
    Next
    DeleteSnapshots hb, tag, True, made
    For Each idxV In deferred
        k = CLng(idxV)
        curTarget = Mid$(CStr(finals.Item(k)), Len(uPre) + 1)
        Set snap = staged.Item(k)
        snap.Name = CStr(finals.Item(k))
    Next

    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    Exit Sub
failed:
    ' Keep Excel's words before anything below can reset Err. Resume
    ' ends the handler, so the roll-back runs as ordinary code and the
    ' refusal raised after it reaches the caller, not this handler.
    failWhy = Err.Description
    Resume rollBack
rollBack:
    On Error GoTo 0
    Dim leftBehind As String
    leftBehind = RemoveSnapshotSheets(made)
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    If Len(leftBehind) = 0 Then
        VLA_Messages.RaiseMsg "ide-undo-snapshot-failed", "program", programName, "sheet", curTarget, "reason", failWhy
    End If
    VLA_Messages.RaiseMsg "ide-undo-snapshot-failed-copies-left", "program", programName, "sheet", curTarget, "reason", failWhy, "left", leftBehind
End Sub

' U.19: remove every sheet a failed snapshot made, whatever name or
' visibility it had reached by then. Returns the ones it could NOT
' remove, quoted and comma-separated - "" when every one went. Its own
' error scope, so one stubborn sheet never stops the rest being tried.
Private Function RemoveSnapshotSheets(made As Collection) As String
    Dim da As Boolean
    da = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim r As String
    Dim sV As Variant
    Dim nm As String
    On Error Resume Next
    For Each sV In made
        nm = ""
        nm = sV.Name
        sV.Visible = xlSheetVisible
        Err.Clear
        sV.Delete
        If Err.Number <> 0 Then
            If Len(r) > 0 Then r = r & ", "
            r = r & "'" & nm & "'"
        End If
        Err.Clear
    Next
    On Error GoTo 0
    Application.DisplayAlerts = da
    RemoveSnapshotSheets = r
End Function

' U.19: the names of a book's worksheets, in tab order.
Private Function SheetNamesOf(hb As Workbook) As Collection
    Dim r As Collection
    Set r = New Collection
    Dim i As Long
    For i = 1 To hb.Worksheets.Count
        r.Add hb.Worksheets(i).Name
    Next
    Set SheetNamesOf = r
End Function

' U.19: the names in afterNames that were not in priorNames - how
' TakeRunSnapshot finds the copy Excel just made, instead of trusting
' where it landed. Compared through Fold, as Excel compares sheet names
' without regard to case. Pure, for the self-test; the caller refuses
' unless there is exactly one.
Public Function VlaIdeAddedSheetNames(priorNames As Collection, afterNames As Collection) As Collection
    Dim seen As Collection
    Set seen = New Collection
    Dim v As Variant
    For Each v In priorNames
        AddTarget seen, CStr(v)
    Next
    Dim r As Collection
    Set r = New Collection
    For Each v In afterNames
        If Not CollHasKeyIde(seen, VLA_Identity.Fold(CStr(v))) Then r.Add CStr(v)
    Next
    Set VlaIdeAddedSheetNames = r
End Function

' U.19: could an Excel sheet carry this name? 1 to 31 characters, none
' of : \ / ? * [ ], no apostrophe first or last, and not Excel's
' reserved "History". Pure, for the self-test.
Public Function VlaIdeCanBeSheetName(ByVal nm As String) As Boolean
    If Len(nm) = 0 Or Len(nm) > 31 Then Exit Function
    If Left$(nm, 1) = "'" Or Right$(nm, 1) = "'" Then Exit Function
    If VLA_Identity.Fold(nm) = "history" Then Exit Function
    Dim i As Long
    For i = 1 To Len(nm)
        If InStr(":\/?*[]", Mid$(nm, i, 1)) > 0 Then Exit Function
    Next
    VlaIdeCanBeSheetName = True
End Function

Public Sub EnglishIdeUndo()
    Dim d As String
    Dim scr As Boolean, da As Boolean
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Set hb = HostBook()
    ' Capture the restore state FIRST: the V2 resolution below can
    ' raise legitimately (ambiguity, tag collision), and the failed
    ' handler restores scr/da - which must hold the real prior
    ' values by then, not Boolean defaults.
    scr = Application.ScreenUpdating
    da = Application.DisplayAlerts
    ' V2: Undo is scoped to ONE program - the one this command
    ' resolved to (its button's sheet, or the book's only program) -
    ' by matching only ITS snapshot prefix. Guard first: colliding
    ' tags would mean a shared history, refused rather than mixed.
    Dim ws As Worksheet
    Set ws = IdeSheet()
    GuardTagCollision hb, ws
    Dim tag As String
    tag = VlaIdeProgramTag(ws.Name)

    ' U.25: the restore itself is PutBackLastRun, shared with a Run that
    ' stops; its failure comes back as words and the sheet it was on.
    Dim rep As VlaUndoReport
    d = PutBackLastRun(hb, tag, rep)
    If Len(d) > 0 Then GoTo reportFailure

    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    ' U.32: the roster's lines, one per kind of act it had to do, after
    ' today's two; nothing when it had nothing to do.
    Dim rosterWords As String
    rosterWords = VlaIdeRosterWords(rep.shown, rep.hidden, rep.moved, rep.activated, rep.missing, rep.rosterProblem)
    If Len(rep.restored) = 0 And Len(rep.removed) = 0 And Len(rosterWords) = 0 Then
        ' U.32: a Run that left a roster has been undone already, or changed
        ' no sheet; "yet" would say no Run happened (the owner's second Undo
        ' of a Run with no copies, 2026-10-08).
        If rep.hadRoster Then
            VlaShowInfo "The sheets are as they were before the last Run of '" & ws.Name & "' - there is nothing to undo."
        Else
            VlaShowInfo "Nothing to undo yet for '" & ws.Name & "' - Undo covers that program's most recent Run."
        End If
    Else
        ' U.32: with a roster, the sheet active before the Run is active
        ' again (the owner's call, 2026-10-07); a Run from before the
        ' roster keeps the older way, the Output sheet brought into view.
        If Not rep.hadRoster Then
            If SheetExists(hb, OUT_SHEET) Then hb.Worksheets(OUT_SHEET).Activate
        End If
        Dim m As String
        m = "Put back the way it was before the last Run of '" & ws.Name & "'"
        If Len(rep.restored) > 0 Then m = m & ": " & rep.restored
        If Len(rep.removed) > 0 Then m = m & vbCrLf & "Removed the sheet" & IIf(InStr(rep.removed, ",") > 0, "s", "") & " the run had created: " & rep.removed
        If Len(rosterWords) > 0 Then m = m & vbCrLf & rosterWords
        VlaShowInfo m
    End If
    Exit Sub
failed:
    d = Err.Description
reportFailure:
    On Error Resume Next
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    On Error GoTo 0
    If Len(rep.restoring) > 0 Then
        VlaShowError "Undo couldn't put back (or remove) the sheet '" & rep.restoring & "'." & vbCrLf & vbCrLf & _
               "Excel says: " & d
    Else
        ' S4.3: Excel's bare description ("Automation error") taught
        ' the owner nothing - say WHERE it died and what usually
        ' clears it. Deeper phase forensics wait for demand.
        VlaShowError "Undo stopped before it could change anything - Excel says: " & d & vbCrLf & vbCrLf & _
               "A leftover sheet from an interrupted Run (a name starting with VLAu_, or a stray copy of a program sheet) can cause this; deleting it clears the way."
    End If
End Sub

' U.25: put back what the last Run of the program tagged `tag` changed -
' Undo Last Run's restore, moved here unchanged so a Run that stops can
' make the same one. Returns "" when it is done, rep.restored and
' rep.removed naming the sheets put back and the sheets the run had
' created that are gone again. Otherwise returns why it stopped, with
' rep.restoring naming the sheet it was on ("" when it stopped before
' changing anything); that sheet is left as the run left it (U.21), and
' the sheets already put back stay put back.
' U.32: a copy is restored IN PLACE - Before the sheet set aside, where a
' restored sheet used to land at the end of the tab order - under its
' sheet's own spelling (the roster's, else the set-aside sheet's, since a
' bare name is scanned in lower case), and takes the visibility the run
' left its sheet; then, once the copies and tombstones
' are done, the Run's roster (VLAs_<tag>_Sheets) is applied: the sheet
' that was active is given back first (a Copy activates what it makes),
' and ApplyRoster shows, moves, hides and activates what the roster says,
' filling rep's lists for the dialog. A Run from before the roster has
' none, and rep.hadRoster says so.
Private Function PutBackLastRun(hb As Workbook, ByVal tag As String, ByRef rep As VlaUndoReport) As String
    Dim scr As Boolean, da As Boolean
    scr = Application.ScreenUpdating
    da = Application.DisplayAlerts
    Dim restored As String, removed As String, restoring As String
    On Error GoTo failed
    Dim uPre As String, dPre As String
    uPre = UNDO_PREFIX & tag & "_"
    dPre = DEL_PREFIX & tag & "_"
    Dim activeAtStart As String
    activeAtStart = ActiveSheetNameOf(hb)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Clean any temp left by a previously interrupted Undo.
    On Error Resume Next
    hb.Worksheets(UNDO_PREFIX & "old").Delete
    On Error GoTo failed

    ' Collect the snapshots FIRST: restoring inserts sheets mid-loop,
    ' which an index-based walk would mis-count.
    ' U.22: matched through the snapshot's own tag, without case - a
    ' program whose tab was re-cased still finds its copies.
    Dim snaps As New Collection
    Dim tombs As New Collection
    Dim rost As Worksheet
    Dim i As Long
    Dim nmI As String
    For i = 1 To hb.Worksheets.Count
        nmI = hb.Worksheets(i).Name
        If VlaIdeTagsMatch(VlaIdeSnapshotTag(nmI), tag) Then
            If VlaIdeSnapshotKind(nmI) = "u" Then
                snaps.Add hb.Worksheets(i)
            ElseIf VlaIdeSnapshotKind(nmI) = "d" Then
                tombs.Add hb.Worksheets(i)
            ElseIf VlaIdeSnapshotKind(nmI) = "s" Then
                Set rost = hb.Worksheets(i)
            End If
        End If
    Next

    ' U.32: the sheets' own spellings, off the roster. A sheet a sentence
    ' names by a bare name is scanned in lower case, so its copy is
    ' "VLAu_<tag>_sales" for a sheet called "Sales"; the copy takes the
    ' recorded name back. A roster that cannot be read gives none, and
    ' ApplyRoster says why below.
    Dim spellRows As Collection
    Dim spellActive As String
    Set spellRows = New Collection
    If Not rost Is Nothing Then
        If Len(VlaIdeRosterParse(ReadRosterText(rost), spellRows, spellActive)) > 0 Then Set spellRows = New Collection
    End If

    Dim snapV As Variant
    Dim snap As Worksheet, cpy As Worksheet, oldWs As Worksheet
    Dim orig As String
    Dim restName As String, asideName As String
    ' U.21: what one sheet's restore has made so far, and whether the
    ' current sheet is set aside - so a failure part-way can be undone
    ' (RollBackRestore, from the handler below).
    Dim made As Collection
    Dim setAside As Boolean
    Dim oldVis As Long
    Dim inRestore As Boolean
    Dim priorNames As Collection
    Dim added As Collection
    Dim addedV As Variant
    For Each snapV In snaps
        Set snap = snapV
        orig = Mid$(snap.Name, Len(uPre) + 1)
        ' U.32: the name the copy comes back under - the roster's spelling,
        ' else the sheet's own as it stands, else the scanned name.
        asideName = ""
        If SheetExists(hb, orig) Then asideName = hb.Worksheets(orig).Name
        restName = VlaIdeRosterSpelling(spellRows, orig)
        If Len(restName) = 0 Then restName = asideName
        If Len(restName) = 0 Then restName = orig
        restoring = restName
        Set made = New Collection
        setAside = False
        oldVis = xlSheetVisible
        inRestore = True
        ' Excel cannot Copy a very-hidden sheet, so unhide the
        ' snapshot for the duration (screen updating is off).
        snap.Visible = xlSheetVisible
        ' Transactional order: set the current sheet ASIDE (rename),
        ' copy the snapshot in, and only then delete the old one -
        ' if the copy fails, nothing has been destroyed.
        If Len(asideName) > 0 Then
            Set oldWs = hb.Worksheets(asideName)
            oldWs.Name = UNDO_PREFIX & "old"
            setAside = True
            ' U.32: the copy goes where the sheet is, Before the one set
            ' aside, which is made visible for the moment: Excel places a
            ' sheet beside a hidden one unreliably, it is deleted a moment
            ' later, and a roll-back gives its visibility back.
            oldVis = oldWs.Visible
            oldWs.Visible = xlSheetVisible
        End If
        ' U.21: the copy is the one NEW sheet name, never
        ' Worksheets(snap.Index + 1) - Index counts chart sheets and
        ' Worksheets does not, so one chart sheet before the snapshot
        ' made that read rename the wrong sheet.
        Set priorNames = SheetNamesOf(hb)
        If setAside Then
            snap.Copy Before:=hb.Worksheets(UNDO_PREFIX & "old")
        Else
            snap.Copy After:=snap
        End If
        Set added = VlaIdeAddedSheetNames(priorNames, SheetNamesOf(hb))
        For Each addedV In added
            made.Add hb.Worksheets(CStr(addedV))
        Next
        If added.Count <> 1 Then
            VLA_Messages.RaiseMsg "ide-undo-snapshot-copy-not-found", "count", CStr(added.Count)
        End If
        Set cpy = hb.Worksheets(CStr(added.Item(1)))
        cpy.Name = restName
        ' U.32: the copy takes the visibility the run left the sheet,
        ' hidden or visible, so the roster can put it back and say so;
        ' very hidden is never written (Frazaro's own storage).
        If setAside And oldVis = xlSheetHidden Then
            cpy.Visible = xlSheetHidden
        Else
            cpy.Visible = xlSheetVisible
        End If
        ' The copy IS the sheet now: a later failure must not remove it.
        Set made = New Collection
        If SheetExists(hb, UNDO_PREFIX & "old") Then
            hb.Worksheets(UNDO_PREFIX & "old").Delete
        End If
        setAside = False
        snap.Visible = xlSheetVeryHidden
        inRestore = False
        restoring = ""
        If Len(restored) > 0 Then restored = restored & ", "
        restored = restored & restName
    Next

    ' D1.2: tombstones - sheets the last run CREATED. Remove each
    ' (with its marker); before the run they did not exist, so after
    ' Undo they must not either. Removal consumes the marker, so a
    ' second Undo correctly reports nothing left to do.
    Dim existed As Boolean
    For Each snapV In tombs
        Set snap = snapV
        orig = Mid$(snap.Name, Len(dPre) + 1)
        restoring = orig
        ' V5.2 (field report): a tombstone marks "this did not exist
        ' before the run" - but the run may never have CREATED it (a
        ' failed "Go to sheet Nowhere-Land." inside a Try). The
        ' marker is still consumed; the dialog names only sheets
        ' actually deleted, because reporting the removal of a sheet
        ' that never existed reads as nonsense to the person - and
        ' it did, in the first real screenshot this project received.
        existed = SheetExists(hb, orig)
        If existed Then hb.Worksheets(orig).Delete
        snap.Visible = xlSheetVisible
        snap.Delete
        restoring = ""
        If existed Then
            If Len(removed) > 0 Then removed = removed & ", "
            removed = removed & orig
        End If
    Next

    ' U.32: the sheet that was active when this began, back (a Copy
    ' activates what it makes), so the roster's own decision about the
    ' active sheet is made against the truth; then the roster, now that
    ' every sheet it orders exists.
    ActivateQuietly hb, activeAtStart
    rep.restored = restored
    rep.removed = removed
    If Not rost Is Nothing Then
        rep.hadRoster = True
        ApplyRoster hb, rost, rep, restoring
    End If
    rep.restoring = restoring

    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    Exit Function
failed:
    PutBackLastRun = Err.Description
    If Len(PutBackLastRun) = 0 Then PutBackLastRun = "error " & Err.Number
    ' U.21: a sheet whose restore failed part-way goes back to how the
    ' Run left it - its stray copy removed, its own name returned - so
    ' the caller's words are true and nothing is left named VLAu_old.
    If inRestore Then RollBackRestore hb, made, setAside, asideName, snap, oldVis
    rep.restored = restored
    rep.removed = removed
    rep.restoring = restoring
    On Error Resume Next
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    On Error GoTo 0
End Function

' V2: a program's snapshot sweep deletes ITS OWN snapshots (prefix
' match on its tag) plus ORPHANS - snapshot sheets whose tag part
' matches no existing workspace, which covers pre-V2 legacy names
' (their "tag" reads as the whole target) and snapshots of programs
' whose sheets were deleted. Another live program's snapshots are
' never touched: that is the whole point of the tag.
'
' U.23: two sweeps now. With includeOwnSnapshots False - before a Run's
' new copies exist - this program's finished copies and tombstones (the
' previous Run's Undo) stay, and only orphans and its own interrupted
' staging go. With True - once the new copies all exist - the previous
' ones go too. `keep` holds the sheets this Run made, which no sweep
' removes. The decision itself is the pure VlaIdeSweepsSnapshot.
Private Sub DeleteSnapshots(hb As Workbook, ByVal tag As String, ByVal includeOwnSnapshots As Boolean, keep As Collection)
    Dim knownTags As Collection
    Set knownTags = New Collection
    Dim wsV As Variant
    For Each wsV In WorkspaceSheets(hb)
        On Error Resume Next
        knownTags.Add "1", VlaIdeProgramTag(wsV.Name)
        On Error GoTo 0
    Next
    Dim i As Long
    Dim da As Boolean
    Dim sweepSheet As Worksheet
    da = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = hb.Worksheets.Count To 1 Step -1
        Set sweepSheet = hb.Worksheets(i)
        If VlaIdeSweepsSnapshot(sweepSheet.Name, tag, knownTags, includeOwnSnapshots) Then
            If Not CollHoldsSheet(keep, sweepSheet) Then
                ' A snapshot is very-hidden; if it would be the last
                ' sheet Excel allows deleting, make it visible first
                ' so the delete can't trip the one-sheet-must-remain
                ' rule.
                On Error Resume Next
                sweepSheet.Visible = xlSheetVisible
                sweepSheet.Delete
                On Error GoTo 0
            End If
        End If
    Next
    Application.DisplayAlerts = da
End Sub

' U.22/U.23: what kind of Frazaro snapshot sheet a name is - "u" an Undo
' copy, "d" a tombstone marker, "n" a copy or marker still being staged,
' "s" a Run's roster (U.32) - or "" for any other sheet. Excel compares
' sheet names without case, so this does too. Pure, for the self-test.
Public Function VlaIdeSnapshotKind(ByVal sheetName As String) As String
    Dim pre As String
    pre = VLA_Identity.Fold(Left$(sheetName, Len(UNDO_PREFIX)))
    If pre = VLA_Identity.Fold(UNDO_PREFIX) Then
        VlaIdeSnapshotKind = "u"
    ElseIf pre = VLA_Identity.Fold(DEL_PREFIX) Then
        VlaIdeSnapshotKind = "d"
    ElseIf pre = VLA_Identity.Fold(STAGE_PREFIX) Then
        VlaIdeSnapshotKind = "n"
    ElseIf pre = VLA_Identity.Fold(ROSTER_PREFIX) Then
        VlaIdeSnapshotKind = "s"
    End If
End Function

' U.22: the program tag inside a snapshot name - the text between the
' prefix and the next underscore - or "" when the name is no snapshot or
' carries no tag (a pre-V2 legacy name). All three prefixes are one
' length, and tags hold no underscore by construction. Pure.
Public Function VlaIdeSnapshotTag(ByVal sheetName As String) As String
    If Len(VlaIdeSnapshotKind(sheetName)) = 0 Then Exit Function
    Dim post As String
    post = Mid$(sheetName, Len(UNDO_PREFIX) + 1)
    Dim cut As Long
    cut = InStr(post, "_")
    If cut > 1 Then VlaIdeSnapshotTag = Left$(post, cut - 1)
End Function

' U.22: do two program tags name the same program? Excel sheet names
' and VBA module names ignore case, so "Gp1" and "gp1" must be one tag,
' or their snapshot names collide. An empty tag matches nothing. Pure.
Public Function VlaIdeTagsMatch(ByVal a As String, ByVal b As String) As Boolean
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    VlaIdeTagsMatch = (VLA_Identity.Fold(a) = VLA_Identity.Fold(b))
End Function

' U.23: does a program's sweep remove this sheet? Not a snapshot: never.
' This program's own: a staged sheet always (only an interrupted Run
' leaves one, and it is never an Undo state), a finished copy or marker
' only when includeOwnSnapshots. Otherwise an orphan - no tag, or a tag
' no live program has - goes; another live program's sheets never do.
' Pure, for the self-test.
Public Function VlaIdeSweepsSnapshot(ByVal sheetName As String, ByVal tag As String, knownTags As Collection, ByVal includeOwnSnapshots As Boolean) As Boolean
    Dim snapKind As String
    snapKind = VlaIdeSnapshotKind(sheetName)
    If Len(snapKind) = 0 Then Exit Function
    Dim snapTag As String
    snapTag = VlaIdeSnapshotTag(sheetName)
    If VlaIdeTagsMatch(snapTag, tag) Then
        VlaIdeSweepsSnapshot = (snapKind = "n" Or includeOwnSnapshots)
    ElseIf Len(snapTag) = 0 Then
        VlaIdeSweepsSnapshot = True
    Else
        VlaIdeSweepsSnapshot = Not CollHasKeyIde(knownTags, snapTag)
    End If
End Function

' Is this very sheet (by object, not by name) one of the collection's?
Private Function CollHoldsSheet(col As Collection, sh As Worksheet) As Boolean
    Dim v As Variant
    For Each v In col
        If v Is sh Then
            CollHoldsSheet = True
            Exit Function
        End If
    Next
End Function

' U.21: undo the part of one sheet's restore that happened before it
' failed - remove the copy it made, give the set-aside sheet its name
' back, hide the snapshot again - so that sheet stays as the Run left
' it. Called from EnglishIdeUndo's handler, so it keeps its own error
' scope: each step is tried whatever the one before it did.
Private Sub RollBackRestore(hb As Workbook, made As Collection, ByVal wasSetAside As Boolean, ByVal orig As String, snap As Worksheet, ByVal oldVis As Long)
    RemoveSnapshotSheets made
    On Error Resume Next
    If wasSetAside Then
        ' U.32: the set-aside sheet was shown for the copy; its
        ' visibility goes back with its name.
        If SheetExists(hb, UNDO_PREFIX & "old") Then hb.Worksheets(UNDO_PREFIX & "old").Visible = oldVis
        If Not SheetExists(hb, orig) Then hb.Worksheets(UNDO_PREFIX & "old").Name = orig
    End If
    If Not snap Is Nothing Then snap.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' =====================================================================
'  U.32: the roster - a workbook's sheets as they were before a Run, and
'  the acts that put them back. The record is text, one row a line: the
'  first line ROSTER_HEAD (versioned), then "active:<name>", then one
'  "<visibility>:<name>" row per worksheet in tab order, the visibility
'  being visible, hidden or very hidden. A colon cannot be in a sheet
'  name (VlaIdeCanBeSheetName), so the first colon ends the kind. The
'  parse, the plan and the dialog's words are pure, for the self-test;
'  the reads and writes of a live workbook sit below them.
' =====================================================================

' U.32: the roster text read into rows ("<visibility>:<name>", in tab
' order) and the active sheet. Returns "" when the text is a roster, or
' why it is not - a first line that is not Frazaro's, a row of a shape it
' does not know - so a sheet carrying the roster's name but not its
' record is reported, never misread. Pure.
Public Function VlaIdeRosterParse(ByVal text As String, ByRef rosterRows As Collection, ByRef activeName As String) As String
    Set rosterRows = New Collection
    activeName = ""
    If Len(text) = 0 Then
        VlaIdeRosterParse = "it is empty"
        Exit Function
    End If
    Dim lines() As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    If lines(0) <> ROSTER_HEAD Then
        VlaIdeRosterParse = "its first line is '" & lines(0) & "', not " & ROSTER_HEAD
        Exit Function
    End If
    If UBound(lines) < 1 Then
        VlaIdeRosterParse = "it has no second line"
        Exit Function
    End If
    If Left$(lines(1), 7) <> "active:" Then
        VlaIdeRosterParse = "its second line is '" & lines(1) & "', not the active sheet"
        Exit Function
    End If
    activeName = Mid$(lines(1), 8)
    Dim i As Long
    Dim visWord As String
    For i = 2 To UBound(lines)
        If Len(lines(i)) > 0 Then
            visWord = RosterKind(lines(i))
            If visWord <> "visible" And visWord <> "hidden" And visWord <> "very hidden" Then
                VlaIdeRosterParse = "line " & (i + 1) & " is '" & lines(i) & "'"
                Exit Function
            End If
            If Len(RosterName(lines(i))) = 0 Then
                VlaIdeRosterParse = "line " & (i + 1) & " names no sheet"
                Exit Function
            End If
            rosterRows.Add lines(i)
        End If
    Next
End Function

' U.32: a sheet's own spelling on the roster - its name as it was before
' the Run - for a name compared without case; "" when the roster has no
' such sheet. A sheet a sentence names by a bare name is scanned, and so
' copied, in lower case; its copy takes this name back. Pure.
Public Function VlaIdeRosterSpelling(recRows As Collection, ByVal nm As String) As String
    Dim k As String
    k = VLA_Identity.Fold(nm)
    Dim v As Variant
    For Each v In recRows
        If VLA_Identity.Fold(RosterName(CStr(v))) = k Then
            VlaIdeRosterSpelling = RosterName(CStr(v))
            Exit Function
        End If
    Next
End Function

' A roster row's two halves: the visibility before the first colon, the
' sheet's name after it.
Private Function RosterKind(ByVal rec As String) As String
    Dim p As Long
    p = InStr(rec, ":")
    If p > 0 Then RosterKind = Left$(rec, p - 1)
End Function

Private Function RosterName(ByVal rec As String) As String
    Dim p As Long
    p = InStr(rec, ":")
    If p > 0 Then RosterName = Mid$(rec, p + 1)
End Function

' The rows' visibilities keyed by folded name.
Private Function RosterKinds(rosterRows As Collection) As Collection
    Dim r As Collection
    Set r = New Collection
    Dim v As Variant
    On Error Resume Next
    For Each v In rosterRows
        r.Add RosterKind(CStr(v)), VLA_Identity.Fold(RosterName(CStr(v)))
    Next
    On Error GoTo 0
    Set RosterKinds = r
End Function

' Where a name sits in a list of names, without case; 0 when absent.
Private Function PosInList(col As Collection, ByVal nm As String) As Long
    Dim i As Long
    Dim k As String
    k = VLA_Identity.Fold(nm)
    For i = 1 To col.Count
        If VLA_Identity.Fold(CStr(col.Item(i))) = k Then
            PosInList = i
            Exit Function
        End If
    Next
End Function

' A name moved to right after position afterIdx (0: first) in a list of
' names keyed by folded name - the plan's own picture of the tab order
' as each move leaves it.
Private Sub ListMoveAfter(col As Collection, ByVal nm As String, ByVal afterIdx As Long)
    Dim k As String
    k = VLA_Identity.Fold(nm)
    Dim j As Long
    j = PosInList(col, nm)
    col.Remove j
    If afterIdx >= j Then afterIdx = afterIdx - 1
    If col.Count = 0 Then
        col.Add nm, k
    ElseIf afterIdx = 0 Then
        col.Add nm, k, Before:=1
    Else
        col.Add nm, k, After:=afterIdx
    End If
End Sub

' A roster sheet hidden now, while there are moves to make: lifted for
' the moves and dropped after them, unless a show act shows it anyway.
Private Sub LiftIfHidden(lifts As Collection, curKind As Collection, shows As Collection, ByVal nm As String)
    Dim k As String
    k = VLA_Identity.Fold(nm)
    If Not CollHasKeyIde(curKind, k) Then Exit Sub
    If curKind.Item(k) <> "hidden" Then Exit Sub
    If CollHasKeyIde(shows, k) Or CollHasKeyIde(lifts, k) Then Exit Sub
    lifts.Add nm, k
End Sub

' U.32: the acts that put the roster's sheets back, decided from the
' roster (its rows in tab order and its active sheet) and the workbook
' now (the same, read the same way), as "<act>:<sheet>[:...]" rows in
' the order they are applied:
'   missing:X            on the roster, gone, and no copy brought it
'                        back - reported, since Undo cannot make it;
'   show:X               recorded visible, hidden now;
'   lift:X               hidden now, with a move to apply: shown for the
'                        moves and hidden again after them (drop:X),
'                        unreported. Excel places a moved sheet by the
'                        visible tabs - before the first visible sheet
'                        after the anchor, or after the last visible one
'                        when none follows - so a hidden roster sheet
'                        beside a move would be stepped over; shown, each
'                        move lands exactly where the plan has it;
'   move:X:after:P       X is not right after P, the roster's sheet
'   move:X:before:F      before it, among the roster's own sheets; or X
'                        should be first and F is first now;
'   drop:X               see lift;
'   hide:X               recorded hidden, visible now;
'   activate:X           the sheet active before the Run, when any act
'                        above was decided or another sheet is active
'                        now (a move or a show can change the active
'                        sheet); only a roster sheet recorded visible.
' Over the roster's sheets that exist now and are not very hidden, now or
' then: Undo never moves or writes a very-hidden sheet, Frazaro's own
' storage, and leaves one the run showed or hid where it is. Sheets not
' on the roster are left alone (the owner's call, 2026-10-07: only a
' tombstoned name is removed). Each move is decided against the order as
' the moves before it leave it, so a rerun with nothing changed decides
' nothing. Names compare without case, as Excel's do. Pure.
Public Function VlaIdeRosterPlan(recRows As Collection, ByVal recActive As String, _
                                 curRows As Collection, ByVal curActive As String) As Collection
    Dim plan As Collection
    Set plan = New Collection
    Dim recKind As Collection, curKind As Collection
    Set recKind = RosterKinds(recRows)
    Set curKind = RosterKinds(curRows)
    Dim v As Variant
    Dim nm As String, k As String
    Dim shows As Collection, hides As Collection
    Set shows = New Collection
    Set hides = New Collection
    For Each v In recRows
        nm = RosterName(CStr(v))
        k = VLA_Identity.Fold(nm)
        If Not CollHasKeyIde(curKind, k) Then
            plan.Add "missing:" & nm
        ElseIf RosterKind(CStr(v)) = "visible" And curKind.Item(k) = "hidden" Then
            shows.Add nm, k
        ElseIf RosterKind(CStr(v)) = "hidden" And curKind.Item(k) = "visible" Then
            hides.Add nm, k
        End If
    Next
    For Each v In shows
        plan.Add "show:" & CStr(v)
    Next

    ' The order, over the roster's sheets that exist and are not very hidden.
    Dim wanted As Collection
    Set wanted = New Collection
    For Each v In recRows
        nm = RosterName(CStr(v))
        k = VLA_Identity.Fold(nm)
        If CollHasKeyIde(curKind, k) Then
            If RosterKind(CStr(v)) <> "very hidden" And curKind.Item(k) <> "very hidden" Then wanted.Add nm, k
        End If
    Next
    Dim have As Collection
    Set have = New Collection
    For Each v In curRows
        nm = RosterName(CStr(v))
        k = VLA_Identity.Fold(nm)
        If CollHasKeyIde(wanted, k) Then have.Add nm, k
    Next
    Dim moves As Collection, lifts As Collection
    Set moves = New Collection
    Set lifts = New Collection
    Dim i As Long, j As Long
    Dim prev As String
    For i = 1 To wanted.Count
        nm = CStr(wanted.Item(i))
        j = PosInList(have, nm)
        If i = 1 Then
            If j <> 1 Then
                moves.Add "move:" & nm & ":before:" & CStr(have.Item(1))
                ListMoveAfter have, nm, 0
            End If
        Else
            prev = CStr(wanted.Item(i - 1))
            If j <> PosInList(have, prev) + 1 Then
                moves.Add "move:" & nm & ":after:" & prev
                ListMoveAfter have, nm, PosInList(have, prev)
            End If
        End If
    Next
    ' Every hidden roster sheet is shown for the moves (the owner's host
    ' run, 2026-10-08: a copy made After a visible sheet landed past the
    ' hidden sheet beside it, Excel placing by the visible tabs); a sheet a
    ' show act shows needs no lift.
    If moves.Count > 0 Then
        For Each v In wanted
            LiftIfHidden lifts, curKind, shows, CStr(v)
        Next
    End If
    For Each v In lifts
        plan.Add "lift:" & CStr(v)
    Next
    For Each v In moves
        plan.Add CStr(v)
    Next
    For Each v In lifts
        plan.Add "drop:" & CStr(v)
    Next
    For Each v In hides
        plan.Add "hide:" & CStr(v)
    Next

    ' The active sheet.
    k = VLA_Identity.Fold(recActive)
    If Len(recActive) > 0 Then
        If CollHasKeyIde(recKind, k) And CollHasKeyIde(curKind, k) Then
            If recKind.Item(k) = "visible" Then
                If plan.Count > 0 Or VLA_Identity.Fold(curActive) <> k Then plan.Add "activate:" & recActive
            End If
        End If
    End If
    Set VlaIdeRosterPlan = plan
End Function

' U.32: the dialog's lines for what the roster did, one per kind of act,
' the sheets named, and the words when its record could not be read; ""
' when it did nothing. Pure.
Public Function VlaIdeRosterWords(ByVal shownList As String, ByVal hiddenList As String, ByVal movedList As String, _
                                  ByVal activeBack As String, ByVal missingList As String, ByVal problem As String) As String
    Dim m As String
    If Len(shownList) > 0 Then m = m & "Shown again: " & shownList & vbCrLf
    If Len(hiddenList) > 0 Then m = m & "Hidden again: " & hiddenList & vbCrLf
    If Len(movedList) > 0 Then m = m & "Put back in their place: " & movedList & vbCrLf
    If Len(activeBack) > 0 Then m = m & "Back on sheet: " & activeBack & vbCrLf
    If Len(missingList) > 0 Then m = m & "Not put back, the run removed them and Undo had no copy: " & missingList & vbCrLf
    If Len(problem) > 0 Then m = m & problem & vbCrLf
    If Len(m) > 0 Then m = Left$(m, Len(m) - 2)
    VlaIdeRosterWords = m
End Function

' U.32: a worksheet's visibility as the roster writes it.
Private Function VisWordOf(ws As Worksheet) As String
    Select Case ws.Visible
        Case xlSheetHidden: VisWordOf = "hidden"
        Case xlSheetVeryHidden: VisWordOf = "very hidden"
        Case Else: VisWordOf = "visible"
    End Select
End Function

' U.32: the roster's rows read off the workbook - every worksheet that is
' no snapshot sheet, in tab order - and the active sheet's name ("" when
' Excel has none to give).
Private Function ReadRosterRows(hb As Workbook, ByRef activeName As String) As Collection
    Dim r As Collection
    Set r = New Collection
    Dim i As Long
    For i = 1 To hb.Worksheets.Count
        If Len(VlaIdeSnapshotKind(hb.Worksheets(i).Name)) = 0 Then
            r.Add VisWordOf(hb.Worksheets(i)) & ":" & hb.Worksheets(i).Name
        End If
    Next
    activeName = ActiveSheetNameOf(hb)
    Set ReadRosterRows = r
End Function

Private Function ActiveSheetNameOf(hb As Workbook) As String
    On Error Resume Next
    ActiveSheetNameOf = hb.ActiveSheet.Name
    On Error GoTo 0
End Function

' U.32: the roster written to its sheet, one line a row from A1, each
' behind an apostrophe so a name shaped like a number or a formula stays
' text.
Private Sub WriteRosterRows(ws As Worksheet, rosterRows As Collection, ByVal activeName As String)
    ws.Cells(1, 1).Value = "'" & ROSTER_HEAD
    ws.Cells(2, 1).Value = "'active:" & activeName
    Dim i As Long
    For i = 1 To rosterRows.Count
        ws.Cells(2 + i, 1).Value = "'" & CStr(rosterRows.Item(i))
    Next
End Sub

' U.32: the roster's text read back, the rows from A1 down to the first
' empty cell, one line each.
Private Function ReadRosterText(ws As Worksheet) As String
    Dim r As Long
    Dim t As String
    r = 1
    Do While Len(CStr(ws.Cells(r, 1).Value)) > 0
        If r > 1 Then t = t & vbLf
        t = t & CStr(ws.Cells(r, 1).Value)
        r = r + 1
    Loop
    ReadRosterText = t
End Function

' U.32: a sheet made active again, quietly - it may be gone or hidden,
' and the roster decides what to do about that.
Private Sub ActivateQuietly(hb As Workbook, ByVal nm As String)
    If Len(nm) = 0 Then Exit Sub
    On Error Resume Next
    hb.Sheets(nm).Activate
    On Error GoTo 0
End Sub

Private Sub AppendName(ByRef list As String, ByVal nm As String)
    If Len(list) > 0 Then list = list & ", "
    list = list & nm
End Sub

' U.32: the roster applied. Its record is parsed first - a sheet carrying
' the roster's name but not its record is said in the dialog, not
' misread, and the sheets already put back stay put back - the workbook
' is read the same way, and the plan's acts are applied in order, each
' naming its sheet in restoring while it runs, so a failure says which.
' rep's lists name what was shown, hidden, moved, made active and found
' missing; a sheet lifted for a move and dropped after it is no act of
' the roster's and goes unreported.
Private Sub ApplyRoster(hb As Workbook, rost As Worksheet, ByRef rep As VlaUndoReport, ByRef restoring As String)
    Dim recRows As Collection
    Dim recActive As String
    Dim why As String
    why = VlaIdeRosterParse(ReadRosterText(rost), recRows, recActive)
    If Len(why) > 0 Then
        On Error Resume Next
        VLA_Messages.RaiseMsg "ide-undo-roster-unreadable", "sheet", rost.Name, "why", why
        rep.rosterProblem = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Dim curRows As Collection
    Dim curActive As String
    Set curRows = ReadRosterRows(hb, curActive)
    Dim plan As Collection
    Set plan = VlaIdeRosterPlan(recRows, recActive, curRows, curActive)
    Dim act As Variant
    Dim parts() As String
    For Each act In plan
        parts = Split(CStr(act), ":")
        restoring = parts(1)
        Select Case parts(0)
            Case "missing"
                AppendName rep.missing, parts(1)
            Case "show", "lift"
                hb.Worksheets(parts(1)).Visible = xlSheetVisible
                If parts(0) = "show" Then AppendName rep.shown, parts(1)
            Case "move"
                If parts(2) = "after" Then
                    hb.Worksheets(parts(1)).Move After:=hb.Worksheets(parts(3))
                Else
                    hb.Worksheets(parts(1)).Move Before:=hb.Worksheets(parts(3))
                End If
                AppendName rep.moved, parts(1)
            Case "drop", "hide"
                hb.Worksheets(parts(1)).Visible = xlSheetHidden
                If parts(0) = "hide" Then AppendName rep.hidden, parts(1)
            Case "activate"
                hb.Worksheets(parts(1)).Activate
                If VLA_Identity.Fold(curActive) <> VLA_Identity.Fold(parts(1)) Then rep.activated = parts(1)
        End Select
        restoring = ""
    Next
End Sub

' U.32: the snapshot and the put-back with a public door each, so the
' host suite can take a Run's snapshot on a scratch workbook and put it
' back without a Run, a button or a dialog.
Public Sub VlaIdeTakeRunSnapshot(hb As Workbook, ByVal programText As String, ByVal tag As String, ByVal programName As String)
    TakeRunSnapshot hb, programText, tag, programName
End Sub

Public Function VlaIdePutBackLastRun(hb As Workbook, ByVal tag As String, ByRef rep As VlaUndoReport) As String
    VlaIdePutBackLastRun = PutBackLastRun(hb, tag, rep)
End Function

' Module-local keyed probe (modules are self-contained by rule 12).
Private Function CollHasKeyIde(col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col.Item(key)
    CollHasKeyIde = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function SheetExists(hb As Workbook, ByVal name As String) As Boolean
    On Error Resume Next
    SheetExists = Not hb.Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Sub AddTarget(col As Collection, ByVal name As String)
    On Error Resume Next
    col.Add name, VLA_Identity.Fold(name)
    On Error GoTo 0
End Sub

' Collect names appearing after the word "sheet" in the program text:
' quoted names verbatim, bare names as their (lowercased) word. "Add
' sheet called X" chains through "called" to reach the name.
Private Sub ScanSheetNames(ByVal text As String, targets As Collection)
    ' V5.2 (first field report - the standing override invoked): the
    ' scan read the RAW program text, so prose in # comment lines
    ' ("the missing sheet and...", "the sheet already exists...")
    ' minted phantom targets named "and", "already", "the" - each
    ' quietly minting a tombstone sheet, then paraded in the Undo
    ' dialog as sheets "the run had created". Comments are prose;
    ' the scan now drops them first (DeclaredOutputSheet always did).
    ' Deliberately still scanned: quoted text values - a sheet name
    ' inside a string can be a real reach, and over-snapshotting is
    ' the safe direction (restoring an untouched sheet is a no-op).
    text = StripCommentLines(text)
    Dim lc As String
    lc = VLA_Identity.Fold(text)
    Dim p As Long, q As Long
    Dim nm As String
    p = 1
    Do
        p = InStr(p, lc, "sheet")
        If p = 0 Then Exit Do
        ' word boundaries around "sheet"
        If (p = 1 Or Not IsNameChar(Mid$(lc, p - 1, 1))) And _
           Not IsNameChar(Mid$(lc, p + 5, 1)) Then
            q = p + 5
            nm = ReadNameAt(lc, text, q)
            If nm = "called" Then nm = ReadNameAt(lc, text, q)
            If Len(nm) > 0 Then AddTarget targets, nm
        End If
        p = p + 5
    Loop

    ' D1 audit find: since B5's quoting liberation, a program can
    ' reach a sheet with no "sheet" keyword at all - the qualified
    ' reference (cell Data!B2, range 'Q1 Data'!A1:B10). Those sheets
    ' must be snapshotted too, or Undo silently would not cover them.
    ' Scan for the two ! shapes: 'quoted name'! keeps its case; a
    ' bare word! is its lowercased word (matching how references
    ' travel). Nonexistent names are harmless - the snapshot loop
    ' only copies sheets that exist - so a "bang!" inside a text
    ' string merely adds a name that matches nothing.
    p = 1
    Dim s0 As Long
    Do
        p = InStr(p, lc, "!")
        If p = 0 Then Exit Do
        nm = ""
        ' V5.2: the qualifier must be word-flanked on BOTH sides -
        ' the tokenizer's own rule, finally applied here too. The
        ' left-only check read the raw VLA row "(set! (range ...))"
        ' as sheet "set"; a "!" not followed by a name character
        ' (set! ( / bang!" / end of text) is punctuation, not a
        ' sheet qualifier.
        If Not IsNameChar(Mid$(lc, p + 1, 1)) Then
            p = p + 1
        Else
        If p > 1 Then
            If Mid$(text, p - 1, 1) = "'" Then
                ' 'Quoted Name'! - walk back to the opening quote;
                ' the inner text keeps its original case, exactly as
                ' the reference travels.
                s0 = InStrRev(text, "'", p - 2)
                If s0 > 0 And p - 1 - s0 > 1 Then nm = Mid$(text, s0 + 1, p - 2 - s0)
            ElseIf IsNameChar(Mid$(lc, p - 1, 1)) Then
                ' word! - walk back over name characters (lowercased,
                ' matching how bare references travel).
                s0 = p - 1
                Do While s0 > 1
                    If Not IsNameChar(Mid$(lc, s0 - 1, 1)) Then Exit Do
                    s0 = s0 - 1
                Loop
                nm = Mid$(lc, s0, p - s0)
            End If
        End If
        If Len(nm) > 0 Then AddTarget targets, nm
        p = p + 1
        End If
    Loop

    ' U.32: the helper rows a raw program can hold name the sheets they
    ' change, as a sentence's "sheet" would.
    ScanHelperRows text, targets
End Sub

' U.32: the sheet names a raw helper row changes, added to the targets as
' a sentence's would be: the sheet a row clears, deletes or renames, so it
' is copied, and the name a row gives, so it is tombstoned. Per head, the
' argument positions that name such a sheet. A copy's source, a moved
' sheet and the sheets hide and show touch are the roster's, since their
' content does not change, and a bare copy's name is Excel's own. Only a
' quoted name counts: a name held in a variable is not known before the
' run. Heads compare without case, as symbols do; the runtime calls under
' the macros are read too, since a raw row can call them. The rows are
' read in the order the program has them, as the sentence scan reads its
' sentences, so the targets follow the program.
Private Sub ScanHelperRows(ByVal text As String, targets As Collection)
    Dim heads As Variant
    heads = Array("add-sheet-called", "1", "add-sheet-at", "1", "copy-sheet-named", "2", _
                  "rename-sheet", "1,2", "clear-sheet", "1", "delete-sheet", "1", _
                  "vlaaddsheetat", "1", "vlacopysheet", "4", "vlarenamesheet", "1,2", _
                  "vlaclearsheet", "1", "vladeletesheet", "1")
    Dim lc As String
    lc = VLA_Identity.Fold(text)
    Dim h As Long, p As Long, q As Long
    Dim hd As String
    p = 1
    Do
        p = InStr(p, lc, "(")
        If p = 0 Then Exit Do
        For h = LBound(heads) To UBound(heads) Step 2
            hd = CStr(heads(h))
            If Mid$(lc, p + 1, Len(hd)) = hd Then
                q = p + 1 + Len(hd)
                ' the whole head: "(copy-sheet" inside "(copy-sheet-named" is not it
                If q <= Len(lc) Then
                    If InStr(" " & vbTab & vbCr & vbLf & ")", Mid$(lc, q, 1)) > 0 Then
                        ReadHelperArgs text, q, CStr(heads(h + 1)), targets
                        Exit For
                    End If
                End If
            End If
        Next
        p = p + 1
    Loop
End Sub

' U.32: the arguments after a helper row's head, from position q, counted
' from 1; each quoted one at a wanted position (a comma-separated list of
' positions) becomes a target. A string reads as the reader reads it, \"
' and \\ being escapes. An argument that is no string - a name, a number,
' a nested form - is stepped over and counted.
Private Sub ReadHelperArgs(ByVal text As String, ByVal q As Long, ByVal wanted As String, targets As Collection)
    Dim argNo As Long
    Dim c As String
    Dim depth As Long
    Dim s As String
    Do While q <= Len(text)
        c = Mid$(text, q, 1)
        If c = ")" Then Exit Do
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            q = q + 1
        ElseIf c = """" Then
            argNo = argNo + 1
            s = ""
            q = q + 1
            Do While q <= Len(text)
                c = Mid$(text, q, 1)
                If c = "\" Then
                    If q < Len(text) Then s = s & Mid$(text, q + 1, 1)
                    q = q + 2
                ElseIf c = """" Then
                    q = q + 1
                    Exit Do
                Else
                    s = s & c
                    q = q + 1
                End If
            Loop
            If InStr("," & wanted & ",", "," & CStr(argNo) & ",") > 0 Then
                If Len(s) > 0 Then AddTarget targets, s
            End If
        ElseIf c = "(" Then
            argNo = argNo + 1
            depth = 0
            Do While q <= Len(text)
                c = Mid$(text, q, 1)
                If c = "(" Then depth = depth + 1
                If c = ")" Then depth = depth - 1
                q = q + 1
                If depth = 0 Then Exit Do
            Loop
        Else
            argNo = argNo + 1
            Do While q <= Len(text)
                c = Mid$(text, q, 1)
                If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Or c = ")" Or c = "(" Then Exit Do
                q = q + 1
            Loop
        End If
    Loop
End Sub

' Comment lines are prose: drop them before the target scan. Same
' rule DeclaredOutputSheet has always used.
Private Function StripCommentLines(ByVal t As String) As String
    Dim lines() As String
    lines = Split(Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim i As Long
    Dim r As String
    For i = LBound(lines) To UBound(lines)
        If Left$(Trim$(lines(i)), 1) <> "#" Then r = r & lines(i) & vbLf
    Next
    StripCommentLines = r
End Function

' V5.2: the scan, exposed pure for the self-test - the phantom
' targets the field report surfaced are pinned through this.
Public Function VlaIdeScanTargets(ByVal programText As String) As String
    Dim c As New Collection
    ScanSheetNames programText, c
    Dim v As Variant
    Dim r As String
    For Each v In c
        If Len(r) > 0 Then r = r & ","
        r = r & CStr(v)
    Next
    VlaIdeScanTargets = r
End Function

Private Function IsNameChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    IsNameChar = (c Like "[A-Za-z0-9_-]")
End Function

' Read a quoted name (original case) or a bare word (lowercase) at
' position q, skipping leading spaces; advances q past it.
Private Function ReadNameAt(ByVal lc As String, ByVal orig As String, ByRef q As Long) As String
    Do While q <= Len(lc)
        If Mid$(lc, q, 1) <> " " Then Exit Do
        q = q + 1
    Loop
    If q > Len(lc) Then Exit Function
    Dim r As String
    If Mid$(orig, q, 1) = """" Then
        q = q + 1
        Do While q <= Len(orig)
            If Mid$(orig, q, 1) = """" Then
                q = q + 1
                Exit Do
            End If
            r = r & Mid$(orig, q, 1)
            q = q + 1
        Loop
    Else
        Do While q <= Len(lc)
            If Not IsNameChar(Mid$(lc, q, 1)) Then Exit Do
            r = r & Mid$(lc, q, 1)
            q = q + 1
        Loop
    End If
    ReadNameAt = r
End Function

' =====================================================================
'  DI.2 Pass 3: self-registration with Excel's own Add-ins list -
'  writes this add-in's path into the OPEN/OPENn registry slots under
'  HKCU\...\Excel\Options, the exact mechanism the Add-Ins dialog
'  itself writes when a user ticks a box (and the one
'  installer/Frazaro.iss's own Pascal Script already proves live end
'  to end: written, confirmed via Application.AddIns picking it up
'  with Installed=True after a fresh Excel launch, removed cleanly).
'  Ported here so a standalone-installed Frazaro (no exe, see
'  DEPLOY.md's "Which download should I use?") never needs the Add-Ins
'  dialog at all - one ribbon click instead.
'
'  Unlike the installer, which cannot know in advance which Excel will
'  eventually open the file and so loops over every installed Office
'  version, this code runs INSIDE the Excel that will load it -
'  Application.Version names the one version that matters; no registry
'  enumeration needed (VBA has none built in; WMI/API declarations
'  would be more machinery than this problem needs).
' =====================================================================

Private Function ExcelOptionsKey() As String
    ExcelOptionsKey = "HKCU\Software\Microsoft\Office\" & Application.Version & "\Excel\Options\"
End Function

Private Function OpenSlotName(ByVal idx As Long) As String
    If idx = 0 Then
        OpenSlotName = "OPEN"
    Else
        OpenSlotName = "OPEN" & idx
    End If
End Function

' Writes targetPath into the next free OPEN/OPENn slot, or does
' nothing if it is already registered in some slot - so calling this
' twice never mints a duplicate entry. Returns True if it made a
' change, False if it was already present. targetPath defaults to this
' add-in's own path (the real, production behavior); overridable so a
' test can register a synthetic path instead of touching this
' session's own real Excel Options.
Public Function VlaRegisterForAutoLoad(Optional ByVal targetPath As String = "") As Boolean
    If Len(targetPath) = 0 Then targetPath = ThisWorkbook.FullName
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim i As Long
    For i = 0 To 50
        Dim slot As String
        slot = ExcelOptionsKey() & OpenSlotName(i)
        Dim existing As String
        Dim gotErr As Boolean
        On Error Resume Next
        existing = sh.RegRead(slot)
        gotErr = (Err.Number <> 0)
        On Error GoTo 0
        If gotErr Then
            sh.RegWrite slot, targetPath, "REG_SZ"
            VlaRegisterForAutoLoad = True
            Exit Function
        ElseIf existing = targetPath Then
            VlaRegisterForAutoLoad = False
            Exit Function
        End If
    Next
    VLA_Messages.RaiseMsg "ide-autoload-no-slot", "key", ExcelOptionsKey()
End Function

' Removes exactly the slot(s) holding targetPath, leaving any other
' add-in's OPEN entry untouched - checks all 50 slots rather than
' exiting on the first match, matching Frazaro.iss's own
' UnregisterExcelAddin exactly, in case more than one somehow exists.
' Returns True if it removed anything. Not yet wired to any ribbon
' button - built alongside VlaRegisterForAutoLoad so the shared search
' logic exists once; the planned "Uninstall Frazaro" button is its
' first real caller.
Public Function VlaUnregisterAutoLoad(Optional ByVal targetPath As String = "") As Boolean
    If Len(targetPath) = 0 Then targetPath = ThisWorkbook.FullName
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim i As Long
    For i = 0 To 50
        Dim slot As String
        slot = ExcelOptionsKey() & OpenSlotName(i)
        Dim existing As String
        Dim gotErr As Boolean
        On Error Resume Next
        existing = sh.RegRead(slot)
        gotErr = (Err.Number <> 0)
        On Error GoTo 0
        If Not gotErr Then
            If existing = targetPath Then
                sh.RegDelete slot
                VlaUnregisterAutoLoad = True
            End If
        End If
    Next
End Function

' The ribbon-facing wrapper - reports what happened via
' VlaShowInfo/VlaShowError, the same convention every other ribbon
' command in this module already uses.
Public Sub VlaIdeRegisterForAutoLoad()
    On Error GoTo failed
    If VlaRegisterForAutoLoad() Then
        VlaShowInfo "Frazaro will now load automatically the next time Excel starts." & vbCrLf & vbCrLf & _
            ThisWorkbook.FullName
    Else
        VlaShowInfo "Frazaro is already registered to load automatically." & vbCrLf & vbCrLf & _
            ThisWorkbook.FullName
    End If
    Exit Sub
failed:
    VlaShowError "Could not register for auto-load: " & Err.Description
End Sub

' =====================================================================
'  DI.2 Pass 4: "Uninstall Frazaro" - the exit door DEPLOY.md's own
'  doctrine has promised since Pass 5 ("both paths get an identical
'  one-click exit inside the product itself"). Confirms, deregisters
'  (VlaUnregisterAutoLoad, already built and tested in Pass 3), hands
'  off to the real Windows uninstaller if this copy was exe-installed
'  (Frazaro.iss places unins000.exe beside Frazaro.xlam, Inno's own
'  default) - or, for a standalone install, genuinely deletes itself.
'
'  The standalone case went through three shapes, each correction
'  caught live by the owner, not found in review:
'  1. "Close and tell the user to delete the file by hand" - correctly
'     called out as not a real answer for a non-technical user asked
'     to find and delete a file in an unfamiliar folder.
'  2. A background PROCESS (not this workbook's own code) launched
'     from THIS Sub, waiting a fixed 3 seconds then deleting once -
'     sidesteps the genuinely untested question of whether code can
'     delete the very file it's running from, by moving the delete
'     into an independent process, the same shape a self-updating
'     installer uses to replace its own running executable. But a
'     fixed delay is fragile against anything that slows the close
'     down, so it became a RETRY loop (ScheduleSelfDelete, below) -
'     live-tested against a file held under an exclusive lock for 2s
'     with the retry already running, recovering cleanly the moment
'     the lock released.
'  3. Still broken, and found by deliberately trying to break it:
'     launching that retry process from THIS Sub meant its 10-second
'     window started counting down the moment the (separate,
'     informational) dialog below appeared - and a human is free to
'     sit on a dialog far longer than any retry budget. The fix:
'     ScheduleSelfDelete is no longer called here at all. It moved to
'     VlaUninstallCloseDeferred, which runs with nothing blocking
'     between it and the actual close - see that Sub's own header.
'     This Sub now only confirms, deregisters, and (for an exe
'     install) hands off to the real uninstaller; everything
'     filesystem-touching for the standalone case happens later.
'  4. 0.6.1, owner-caught at the cost of the dev workbook itself: none
'     of the above ever asked WHICH file ThisWorkbook is. The ribbon's
'     onAction is the bare name "VlaRibbonAction", and with VLA.xlsm
'     open beside an installed add-in, Excel ran the dev workbook's own
'     copy of it - so ThisWorkbook was VLA.xlsm, no unins000.exe sat
'     beside it, and ScheduleSelfDelete force-deleted it (Remove-Item
'     -Force, past the Recycle Bin). Three guards now, each sufficient:
'     VlaUninstallRefusal refuses anything but a built add-in (IsAddin
'     AND a .xlam name) here AND again in VlaUninstallCloseDeferred
'     (Public, so reachable from the Macros dialog without passing
'     through here); and ScheduleSelfDelete refuses any non-.xlam path
'     on its own. The delete itself stays a hard delete: a Recycle Bin
'     version was built and dropped the same session, since recycling
'     an add-in is not uninstalling it (owner decision). The dialog's
'     "companion 'scripts' folder" line had the same flaw - any folder
'     named scripts beside the add-in, live the repo's own - so it now
'     lists files only when every entry is Frazaro's own (see
'     UninstallCompanionFiles).
' =====================================================================

' Why Uninstall must refuse to touch this file, or "" when it may: only
' a built add-in - IsAddin (VlaBuildAddin sets it) AND a .xlam name -
' is ever Frazaro's own install. Pure over its two inputs so the self-
' test can pin it without the workbook it would otherwise delete.
Public Function VlaUninstallRefusal(ByVal isAddin As Boolean, ByVal fullName As String) As String
    If Not isAddin Then
        VlaUninstallRefusal = "this workbook is not an installed Frazaro add-in, so there is nothing to uninstall and it will not be touched"
    ElseIf LCase$(Right$(fullName, 5)) <> ".xlam" Then
        VlaUninstallRefusal = "only a Frazaro .xlam add-in can be uninstalled, and this file is not one, so it will not be touched"
    End If
End Function

' The files a standalone Frazaro keeps in a 'scripts' folder beside the
' add-in (DEPLOY.md, "Installing"): its prelude and edition phrasebooks.
Public Function VlaIsUninstallCompanion(ByVal fileName As String) As Boolean
    Select Case LCase$(fileName)
        Case "prelude.vla", "english.vla", "espanol.vla"
            VlaIsUninstallCompanion = True
    End Select
End Function

' 0.6.1: the full paths of the companion files in folder, one per line -
' or "" when the folder is missing, empty, or holds ANYTHING else. A
' folder merely named 'scripts' beside the add-in is not proof it is
' Frazaro's: live, beside a dev copy in the repo root, the old dialog
' told the owner to delete the repo's own scripts\ folder by hand.
' So the dialog names files, never the folder, and only when every
' entry in it is one of Frazaro's own.
Private Function UninstallCompanionFiles(ByVal folder As String) As String
    On Error GoTo notOurs
    Dim found As String
    Dim entry As String
    entry = Dir$(folder & "\*", vbDirectory Or vbHidden Or vbSystem)
    Do While Len(entry) > 0
        If entry <> "." And entry <> ".." Then
            If (GetAttr(folder & "\" & entry) And vbDirectory) = vbDirectory Then GoTo notOurs
            If Not VlaIsUninstallCompanion(entry) Then GoTo notOurs
            found = found & vbCrLf & folder & "\" & entry
        End If
        entry = Dir$()
    Loop
    If Len(found) > 0 Then UninstallCompanionFiles = Mid$(found, Len(vbCrLf) + 1)
    Exit Function
notOurs:
    UninstallCompanionFiles = ""
End Function

Public Sub VlaIdeUninstall()
    On Error GoTo failed
    Dim refusal As String
    refusal = VlaUninstallRefusal(ThisWorkbook.IsAddin, ThisWorkbook.FullName)
    If Len(refusal) > 0 Then
        VlaShowError "Uninstall Frazaro stopped: " & refusal & "." & vbCrLf & vbCrLf & ThisWorkbook.FullName
        Exit Sub
    End If

    Dim uninstPath As String
    uninstPath = ThisWorkbook.Path & "\unins000.exe"
    Dim exeInstalled As Boolean
    exeInstalled = (Len(Dir$(uninstPath)) > 0)

    Dim leftovers As String
    leftovers = UninstallCompanionFiles(ThisWorkbook.Path & "\scripts")

    ' Every fact the user needs goes in THIS one dialog, before they
    ' commit - owner-caught, live: an earlier version showed a SECOND,
    ' purely informational dialog after this one, and deliberately
    ' pausing on it for 30+ seconds outran ScheduleSelfDelete's own
    ' 10-second retry window entirely, because that window had already
    ' started counting down the moment the second dialog appeared -
    ' long before the workbook (still open, blocked on that dialog)
    ' ever actually closed. One dialog, nothing blocking after Yes,
    ' fixes the whole class of bug rather than tuning the window again.
    Dim confirmMsg As String
    confirmMsg = "This removes Frazaro from Excel's auto-load list and closes it." & vbCrLf & vbCrLf & _
        ThisWorkbook.FullName & vbCrLf & vbCrLf
    If exeInstalled Then
        confirmMsg = confirmMsg & "The Windows uninstaller will then remove the rest."
    Else
        confirmMsg = confirmMsg & "This file will then be deleted automatically."
        If Len(leftovers) > 0 Then confirmMsg = confirmMsg & vbCrLf & vbCrLf & _
            "Frazaro's own files beside it stay behind. If no other workbook uses them, delete them by hand:" & _
            vbCrLf & leftovers
    End If
    confirmMsg = confirmMsg & vbCrLf & vbCrLf & "Continue?"

    Dim resp As VbMsgBoxResult
    resp = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Uninstall Frazaro")
    If resp <> vbYes Then Exit Sub

    VlaUnregisterAutoLoad

    If exeInstalled Then
        ' Hand off to the real Windows uninstaller, which removes the
        ' files and the Add/Remove Programs entry - launched now,
        ' before this workbook closes, the same way an "uninstall me"
        ' command inside a running app normally works. WScript.Shell.Run,
        ' not bare VBA Shell, matching the one external-process
        ' convention this codebase already uses (VlaInjectRibbon's own
        ' PowerShell invocation). WindowStyle 1 = normal, visible
        ' window (the uninstaller has its own UI); wait:=False - this
        ' Sub must not block waiting for it.
        Dim sh As Object
        Set sh = CreateObject("WScript.Shell")
        sh.Run """" & uninstPath & """", 1, False
    End If
    ' Standalone case: the actual self-delete is scheduled from
    ' VlaUninstallCloseDeferred, not here - see that Sub's own header
    ' for why (this is the fix for the bug described above).

    ' Deferred close, not a direct ThisWorkbook.Close here - closing
    ' the workbook whose own code is mid-execution, synchronously,
    ' inside the very ribbon callback Excel just invoked on it, is
    ' exactly the kind of VBA lifecycle landmine this project has hit
    ' before in other shapes. Scheduling the close for a moment after
    ' this Sub returns lets the current call stack unwind first - and,
    ' critically, nothing blocking sits between here and there anymore.
    Application.OnTime Now + TimeSerial(0, 0, 1), _
        "'" & ThisWorkbook.Name & "'!VlaUninstallCloseDeferred"
    Exit Sub
failed:
    VlaShowError "Could not uninstall: " & Err.Description
End Sub

' Launches a detached PowerShell process that RETRIES the delete of
' targetPath every 500ms for up to 20 attempts (10s), rather than
' gambling on one fixed-delay attempt - owner-caught concern, correct:
' a single timed guess is fragile against anything that slows the
' close down (a loaded machine, antivirus scanning the file, a
' cloud-sync client touching it, Workbook_BeforeClose's own cleanup
' work taking longer than usual). Live-tested against exactly that
' failure shape before trusting it: a file held under an exclusive
' lock for 2 seconds, with the retry loop started while it was still
' locked - it kept failing silently and retrying every 500ms, then
' deleted successfully on the attempt immediately after the lock was
' released (4th attempt, ~2s in), where a fixed "wait N seconds, try
' once" design would have needed to guess N correctly in advance.
' Deliberately a SEPARATE process, not this workbook's own code, and
' writes a temp .ps1 invoked via -File rather than an inline -Command
' string, matching this codebase's own established PowerShell-
' automation shape (VlaInjectRibbon's RibbonPs1, tools/sign_installer.ps1)
' instead of introducing a new one. Returns True if the process was
' launched (not proof the delete itself later succeeded - there is no
' synchronous way to know that from here, and after 20 failed attempts
' the script simply gives up rather than retrying forever); False if
' even launching it failed, so the caller can fall back to the manual
' instruction.
' 0.6.1: refuses any path that is not a .xlam, whoever calls it. The
' delete stays a hard one - uninstall means gone, not recycled (owner
' decision); the guards, not a recoverable delete, are the safety.
Private Function ScheduleSelfDelete(ByVal targetPath As String) As Boolean
    If LCase$(Right$(targetPath, 5)) <> ".xlam" Then Exit Function
    On Error GoTo failed
    Dim tmp As String
    tmp = Environ$("TEMP")
    Dim psPath As String
    psPath = tmp & "\VlaSelfDelete.ps1"

    Dim f As Integer
    f = FreeFile
    Open psPath For Output As #f
    Print #f, "$target = '" & Replace(targetPath, "'", "''") & "'"
    Print #f, "for ($i = 0; $i -lt 20; $i++) {"
    Print #f, "    Start-Sleep -Milliseconds 500"
    Print #f, "    try {"
    Print #f, "        Remove-Item -LiteralPath $target -Force -ErrorAction Stop"
    Print #f, "        exit 0"
    Print #f, "    } catch {}"
    Print #f, "}"
    Close #f

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psPath & """", 0, False
    ScheduleSelfDelete = True
    Exit Function
failed:
    ScheduleSelfDelete = False
End Function

' The deferred half of VlaIdeUninstall - runs on its own, a moment
' later, once Excel's own ribbon-callback call stack has unwound, with
' no dialog or other blocking step between here and the actual close.
' Launches ScheduleSelfDelete HERE, immediately before closing, rather
' than back in VlaIdeUninstall before a dialog the user could sit on
' indefinitely - live-caught bug, fixed structurally rather than by
' tuning a timeout: the retry window now starts counting down right at
' the moment that matters, not whenever a human finishes reading a
' message box. Re-derives install type from ThisWorkbook rather than
' threading a parameter through Application.OnTime, which only
' supports bare procedure names, no arguments.
Public Sub VlaUninstallCloseDeferred()
    ' 0.6.1: re-checked here - Public, so reachable without the confirm.
    If Len(VlaUninstallRefusal(ThisWorkbook.IsAddin, ThisWorkbook.FullName)) > 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(ThisWorkbook.Path & "\unins000.exe")) = 0 Then
        ScheduleSelfDelete ThisWorkbook.FullName
    End If
    ThisWorkbook.Close False
End Sub

' =====================================================================
'  Add-in menu: appears under the ribbon's Add-ins tab. Wired up by
'  the built add-in's Workbook_Open; safe to call manually too.
' =====================================================================

Public Sub VlaAddinMenu()
    VlaAddinMenuRemove
    Dim pop As CommandBarPopup
    Set pop = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
        Type:=msoControlPopup, Temporary:=True)
    ' V5.1: the menu wears the product name, same rationale as the
    ' artifact rename - the ribbon tab already said Frazaro.
    pop.Caption = "Frazaro"
    AddMenuBtn pop, "Set Up Workspace", "VlaIdeSetup"
    AddMenuBtn pop, "Register for Auto-Load", "VlaIdeRegisterForAutoLoad"
    AddMenuBtn pop, "Add Program...", "EnglishIdeAddProgram"
    AddMenuBtn pop, "Import Program File...", "EnglishIdeImport"
    AddMenuBtn pop, "Reload Instructions", "EnglishIdeReload"
    AddMenuBtn pop, "Uninstall Frazaro", "VlaIdeUninstall"
    AddMenuBtn pop, "Check Instructions", "EnglishIdeCheck"
    AddMenuBtn pop, "Interpret Instructions", "EnglishIdeInterpret"
    AddMenuBtn pop, "Interpret and Trace", "EnglishIdeInterpretTrace"
    AddMenuBtn pop, "Compile Instructions", "EnglishIdeRun"
    AddMenuBtn pop, "Compile and Trace", "EnglishIdeRunTrace"
    AddMenuBtn pop, "Show Compiled VBA", "EnglishIdeShowVba"
    AddMenuBtn pop, "Translate to VLA...", "EnglishIdeTranslateVla"
    AddMenuBtn pop, "Translate to VBA...", "EnglishIdeTranslateVba"
    AddMenuBtn pop, "Undo Last Run", "EnglishIdeUndo"
    AddMenuBtn pop, "Known Sentences", "EnglishIdeShowPhrases"
    AddMenuBtn pop, "Export Expanded Phrasebook...", "EnglishIdeExportExpandedVocabulary"
    AddMenuBtn pop, "Phrasebook Test Coverage...", "EnglishIdeRuleCoverageReport"
    AddMenuBtn pop, "Lint VLA...", "EnglishIdeLintVla"
    AddMenuBtn pop, "Copy Feedback", "EnglishIdeCopyFeedback"
    AddMenuBtn pop, "Optimize Selected Cell", "VlaOptimizeSelectedCell"
End Sub

' V5: the ribbon's single callback - every Frazaro tab button
' dispatches here by id, onto the SAME public commands the sheet
' buttons and the legacy menu call, so the three surfaces cannot
' drift. IRibbonControl lives in the Office object library, which
' every Excel VBA project references by default. The Case ids are
' pinned by the self-test against VlaRibbonXml, so a renamed button
' cannot silently orphan its command.
Public Sub VlaRibbonAction(control As IRibbonControl)
    On Error GoTo failed
    Select Case control.ID
        Case "VlaSetup": VlaIdeSetup
        Case "VlaRegister": VlaIdeRegisterForAutoLoad
        Case "VlaAddProgram": EnglishIdeAddProgram
        Case "VlaImport": EnglishIdeImport
        Case "VlaReload": EnglishIdeReload
        Case "VlaUninstall": VlaIdeUninstall
        Case "VlaCheck": EnglishIdeCheck
        Case "VlaInterpret": EnglishIdeInterpret
        Case "VlaInterpretTrace": EnglishIdeInterpretTrace
        Case "VlaRun": EnglishIdeRun
        Case "VlaRunTrace": EnglishIdeRunTrace
        Case "VlaShowVba": EnglishIdeShowVba
        Case "VlaTranslateVla": EnglishIdeTranslateVla
        Case "VlaTranslateVba": EnglishIdeTranslateVba
        Case "VlaUndo": EnglishIdeUndo
        Case "VlaPhrases": EnglishIdeShowPhrases
        Case "VlaLoadPhrasebook": EnglishIdeLoadPhrasebook
        Case "VlaExportExpanded": EnglishIdeExportExpandedVocabulary
        Case "VlaRuleCoverage": EnglishIdeRuleCoverageReport
        Case "VlaLintVla": EnglishIdeLintVla
        Case "VlaFeedback": EnglishIdeCopyFeedback
        Case "VlaForgetPhrasebooks": EnglishIdeForgetPhrasebookApprovals
        Case "VlaOpenCli": VlaOpenCli
        Case "VlaOptimizeCell": VlaOptimizeSelectedCell
        Case Else
            VlaShowError "Unknown ribbon command: " & control.ID
    End Select
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Public Sub VlaAddinMenuRemove()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Frazaro").Delete
    ' V5.1: also sweep the pre-rename caption, so a session that
    ' loaded an older build cannot strand a stale menu beside the
    ' new one (Temporary:=True already clears both at restart).
    Application.CommandBars("Worksheet Menu Bar").Controls("VLA English").Delete
    On Error GoTo 0
End Sub

Private Sub AddMenuBtn(pop As CommandBarPopup, ByVal caption As String, ByVal macro As String)
    Dim b As CommandBarControl
    Set b = pop.Controls.Add(Type:=msoControlButton, Temporary:=True)
    b.Caption = caption
    b.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
End Sub

' =====================================================================
'  The file bridge: write the program in Word or any text editor;
'  Import pours it into column B (curly quotes and Word line endings
'  normalized) and Checks it immediately. The file stays the source
'  of truth: its path is remembered in A1 (hidden), and "Reload
'  Instructions" re-reads it after every edit - the daily loop is
'  edit, save, Reload, Interpret (or Compile).
' =====================================================================

' GO.6: the button an ordinary phrasebook author - an org admin, a
' community contributor - actually needs. VLA_IDE.IdeVocabPath's own
' four candidate paths were never that: undocumented anywhere a user
' would see them, built for internal edition/dev purposes
' (EDITIONMANIFEST.*), and a full REPLACEMENT of the base corpus by
' exact filename match, never an ADDITION alongside it. This button
' calls EnglishLoadVocabulary directly - the same file-path loader
' Translate to VLA/VBA already uses - so it inherits SEC.2's raw-
' consent gate and G3's override/same-shape-collision refusal for
' free, exactly as GO.6's own roadmap entry says it must.
'
' IdeLoadVocab runs FIRST, unconditionally: it settles the base corpus
' and replays every phrasebook already remembered from a prior click
' (ReplayPersistedPhrasebooks), so the file picked here is always
' loaded against the CURRENT full stack, not just the bare base corpus
' - correct override resolution and an accurate "already loaded" check
' both depend on that. If the chosen file is already remembered, the
' replay just performed already loaded it: loading it a SECOND time
' here would re-register the same rules with no override: marker and
' raise a same-shape-collision refusal, so PhrasebookAlreadyLoaded
' short-circuits before that ever happens.
Public Sub EnglishIdeLoadPhrasebook()
    On Error GoTo failed
    CaptureHost
    IdeLoadVocab

    Dim f As Variant
    f = Application.GetOpenFilename( _
        "Phrasebooks (*.vla),*.vla,All files (*.*),*.*", _
        , "Load Phrasebook")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled

    Dim path As String
    path = CStr(f)
    Dim hb As Workbook
    Set hb = HostBook()

    ' SEC.9: picking this file in an Open dialog IS the approval the gate
    ' would otherwise stop to ask for, so record it rather than asking
    ' again a half-second later. Recorded BEFORE the load, so the replay
    ' on the next Check finds a decision already there and stays silent
    ' about a file the person just deliberately opened.
    '
    ' And recorded ABOVE the already-remembered check, which is the whole
    ' point rather than a tidy-up. This dialog is the ONLY way inside the
    ' product to reverse a previous "no" - but a declined phrasebook is
    ' still in this workbook's remembered list, so PhrasebookAlreadyLoaded
    ' says True and the early exit below used to run before any approval
    ' was recorded. The one route out of a decline was closed to exactly
    ' the files that needed it. Owner-caught live, second pass.
    GrantPhrasebookPath path

    If PhrasebookAlreadyLoaded(hb, path) Then
        ' "Remembered" and "loaded" are not the same state, and saying
        ' the wrong one is what made the dead end above invisible: a
        ' declined phrasebook is remembered and NOT loaded. The
        ' loaded-sources report is the authority on which it is, and it
        ' is already being read here, so ask it rather than assume.
        Dim rpt As String
        rpt = EnglishLoadedSourcesReport()
        If InStr(1, rpt, path, vbTextCompare) > 0 Then
            VlaShowInfo "Already loaded: " & Dir$(path) & vbCrLf & vbCrLf & rpt
        Else
            VlaShowInfo "Approved: " & Dir$(path) & vbCrLf & vbCrLf & _
                "This workbook already remembered this phrasebook, but it was not loaded - it had been declined. " & _
                "Picking it here approves it, and it will load from the next Check." & vbCrLf & vbCrLf & rpt
        End If
        Exit Sub
    End If

    Dim added As Long
    added = EnglishLoadVocabulary(path)
    PersistPhrasebookPath hb, path
    VlaShowInfo "Loaded: " & Dir$(path) & " (" & added & " rule" & IIf(added = 1, "", "s") & " added)" & _
                vbCrLf & vbCrLf & EnglishLoadedSourcesReport()
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Public Sub EnglishIdeImport()
    On Error GoTo failed
    CaptureHost
    MigrateLegacySheet
    ' V2: import pours into the RESOLVED program - the sheet the
    ' command was invoked on, or the book's only one. First contact
    ' (no workspace at all) builds the default workspace; ambiguity
    ' between several programs raises IdeSheet's naming message.
    Dim ws As Worksheet
    If WorkspaceSheets(HostBook()).Count = 0 Then
        VlaIdeSetup                           ' first contact: build the workspace
    End If
    Set ws = IdeSheet()
    EnsureModernLayout ws   ' S3.3: migrate BEFORE writing at FIRST_ROW,
                            ' or a legacy header row would swallow line 1

    Dim f As Variant
    f = Application.GetOpenFilename( _
        "Programs (*.txt;*.en;*.docx;*.doc),*.txt;*.en;*.docx;*.doc,All files (*.*),*.*", _
        , "Import Program")
    If VarType(f) = vbBoolean Then Exit Sub   ' cancelled
    ImportFromPath ws, CStr(f)
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Public Sub EnglishIdeReload()
    On Error GoTo failed
    CaptureHost
    Dim ws As Worksheet
    Set ws = IdeSheet()
    Dim path As String
    path = Trim$(CStr(ws.Cells(1, 1).Value))
    If Len(path) = 0 Then
        EnglishIdeImport                      ' nothing imported yet: pick a file
        Exit Sub
    End If
    ' D1 audit find: Dir$ itself raises on a disconnected drive or
    ' dead UNC path - which is exactly a "file has moved" situation,
    ' so both roads lead to the same friendly message.
    Dim gone As Boolean
    On Error Resume Next
    gone = (Len(Dir$(path)) = 0)
    If Err.Number <> 0 Then gone = True
    On Error GoTo failed
    If gone Then
        VLA_Messages.RaiseMsg "ide-program-file-moved", "path", path
    End If
    ImportFromPath ws, path
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Private Sub ImportFromPath(ws As Worksheet, ByVal path As String)
    EnsureModernLayout ws   ' S3.3: every import route migrates first -
                            ' this is the single chokepoint both the
                            ' Import dialog and Reload Instructions use
    Dim ext As String
    ext = VLA_Identity.Fold(Mid$(path, InStrRev(path, ".") + 1))
    ' SOP.1, owner's call 2026-09-19: PDF intake is NOT shipped, and a
    ' PDF is refused BY NAME here rather than falling through to
    ' VlaReadFile, which would pour a few hundred rows of compressed
    ' bytes over the person's program. The refusal names the way out -
    ' convert it in Word first - and warns about what that conversion
    ' does, because the measurement behind this decision applies just
    ' as much when a person runs it by hand: Word's PDF conversion
    ' re-derives paragraphs from the geometry of the page, joining
    ' steps that sit close together and swallowing the blank line that
    ' ends an indented block. See SOP.1 in docs/BETA_REARVIEW.md for
    ' the measured numbers, and tools/export_example_pdfs.ps1 for the
    ' instrument that produced them.
    If ext = "pdf" Then
        VLA_Messages.RaiseMsg "ide-pdf-not-supported", "file", FileNameOnly(path)
    End If
    Dim text As String
    If VlaIdeReadsThroughWord(ext) Then
        text = ReadWordFile(path)
        ' SOP.1: a scan is a picture of text, and Word hands back
        ' nothing but its own control marks for one. Refused HERE,
        ' before PourProgram: pouring it would clear the program the
        ' person already has and leave one empty row where it was -
        ' silent loss, where a refusal in words costs them nothing. A
        ' .docx of pasted screenshots is exactly that case.
        If Not VlaIdeHasReadableText(text) Then
            VLA_Messages.RaiseMsg "ide-word-no-text", "file", FileNameOnly(path)
        End If
    Else
        text = VlaReadFile(path)
    End If
    ' SOP.6: a document with a <Frazaro> tag pours only its tagged
    ' sections. Every refusal of a malformed tag is raised in here,
    ' before PourProgram, so the program on the sheet is untouched.
    Dim hasTags As Boolean
    Dim prog As String
    prog = VlaIdeFrazaroSections(NormalizeProgramText(text), FileNameOnly(path), _
                                 VLA_RELEASE_VERSION, IdeTagLanguages(), hasTags)
    PourProgram ws, prog
    ws.Cells(1, 1).Value = path
    ' SOP.6: the point of pain teaches the tag. A document with no tag
    ' that fails its first Check has just been read as instructions from
    ' top to bottom, notes for people included - so the red row says how
    ' to fence the instructions off. On the row, never in a dialog: a
    ' person reloading a program they are still fixing is not stopped by
    ' it every time.
    If hasTags Then
        DoCheck ws                            ' immediate feedback per row
    Else
        DoCheck ws, failHint:="(Loading a whole SOP? Put <Frazaro> on a line of its own above the instructions and </Frazaro> below them, then press Reload Instructions: Frazaro reads only that part, and the rest of the document can stay exactly as it is.)"
    End If
End Sub

' SOP.1: the extensions Word reads for us - Word's own formats, and
' only those. Public and pure so the host-free suite can pin the
' routing, which is otherwise provable only by opening Word.
'
' ".pdf" is deliberately NOT here. Word can open a PDF, through its own
' conversion, and SOP.1 measured what comes back: paragraphs re-derived
' from the geometry of the page, steps joined into one line about once
' a page, and - worst - the blank line that closes an indented block
' swallowed, which re-nests a procedure while every line still passes
' Check. ImportFromPath refuses a .pdf by name instead.
Public Function VlaIdeReadsThroughWord(ByVal ext As String) As Boolean
    Select Case VLA_Identity.Fold(ext)
        Case "doc", "docx", "docm": VlaIdeReadsThroughWord = True
    End Select
End Function

' SOP.1: is there anything in what Word handed back that a person could
' read? Word's text stream is full of characters that are not text -
' Chr(1) where a picture sits, Chr(7) at a table cell, Chr(12) at a page
' break, Chr(13) at every paragraph - and a scanned page converts to a
' document made of exactly those and nothing else. Everything below 33
' counts as not-text, as does 160 (the non-breaking space), so the
' answer does not depend on which of them a given Word version emits.
Public Function VlaIdeHasReadableText(ByVal t As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If c < 0 Then c = c + 65536   ' AscW returns a SIGNED Integer - everything above &H7FFF comes back negative
        If c > 32 And c <> 160 Then
            VlaIdeHasReadableText = True
            Exit Function
        End If
    Next
End Function

' The file's own name, for a refusal to name. Deliberately not Dir$,
' which raises on a disconnected drive or a dead UNC path - the D1
' audit find EnglishIdeReload already works around - because a refusal
' must never fail while it is being written.
Private Function FileNameOnly(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p = 0 Then p = InStrRev(path, "/")
    FileNameOnly = Mid$(path, p + 1)
End Function

' Word documents via late-bound automation, always through Frazaro's
' OWN hidden Word: started for this one read, quit after it.
'
' SOP.1 replaced the earlier "reuse a running Word if there is one"
' shape. The reasons are worth keeping written down:
'   * Documents.Open on a file ALREADY open in that Word hands back
'     THEIR open document (its Revert argument defaults to "activate
'     the open document"), and the doc.Close below would then shut
'     their window and discard unsaved edits. That is the ordinary
'     state of a person editing an SOP in Word and pressing Reload
'     Instructions.
'   * GetObject can hand back an INVISIBLE Word left behind by a
'     crash, where every prompt hangs Excel exactly as a hidden one
'     would - with alerts deliberately left unsuppressed, because that
'     instance reads as "not ours".
' What it costs, stated rather than discovered later: a second or two
' of Word start-up on every read, even when Word is already open, and
' the file is read AS SAVED - unsaved edits in the user's own Word are
' not seen, so the advice is save first, then Reload.
'
' SEC.13: Word's AutomationSecurity default under automation is
' msoAutomationSecurityLow, so without the guard below Documents.Open
' runs a .docm's AutoOpen/Document_Open the instant Frazaro reads the
' file - and ImportFromPath routes .doc/.docx/.docm here straight off
' the "Import Program File..." menu item and ribbon button, so this is
' a one-click path, not a hypothetical one. Both settings are made
' AFTER "On Error GoTo cleanup" on purpose: if either cannot be set we
' refuse through cleanup, rather than open the document unguarded or on
' an instance that can still stop to ask a question.
'
' The second half of that reverses SEC.13's own "a failed DisplayAlerts
' is tolerated, this is robustness and must not refuse an import", and
' the reason it reverses is the instance question above, not PDF. When
' that line was written, an alert could still land on the user's own
' VISIBLE Word, where a person could answer it. Every read now happens
' on a hidden instance of our own, so an alert we failed to silence has
' no window anybody can click and can only appear as Excel hanging. A
' refusal in words is strictly better than that.
Private Function ReadWordFile(ByVal path As String) As String
    If VlaOnMac() Then VLA_Messages.RaiseMsg "ide-word-needs-windows"

    Dim wordApp As Object, doc As Object
    Dim ours As Boolean, mine As Boolean
    Dim d As String
    Dim n As Long

    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    d = Err.Description       ' captured before On Error GoTo 0, which clears Err
    On Error GoTo 0
    If wordApp Is Nothing Then VLA_Messages.RaiseMsg "ide-word-not-available", "detail", d

    ' Prove it is ours BEFORE changing a setting or quitting anything.
    ' CreateObject starts a NEW Word - Word registers single-use - and a
    ' Word started for automation has no window and no document open. If
    ' this one has either, it belongs to somebody else, most likely the
    ' user, and quitting it would throw away their unsaved work: refuse
    ' instead, having touched nothing. A probe that errors counts as not
    ' ours for the same reason. Belt and braces against this function's
    ' one load-bearing assumption, so that being wrong about it refuses
    ' rather than closes a Word somebody is working in.
    On Error Resume Next
    mine = (Not wordApp.Visible) And (wordApp.Documents.Count = 0)
    If Err.Number <> 0 Then mine = False
    On Error GoTo 0
    If Not mine Then
        Set wordApp = Nothing
        VLA_Messages.RaiseMsg "ide-word-not-available", "detail", _
            "Word did not open a separate copy for Frazaro to use, and Frazaro never changes or closes a Word you are working in"
    End If

    On Error GoTo cleanup
    ours = True
    wordApp.DisplayAlerts = 0   ' 0 = wdAlertsNone - the literal, not the named Word constant, so this compiles with no dependency on the Word Object Library being a checked reference
    wordApp.AutomationSecurity = 3   ' 3 = msoAutomationSecurityForceDisable - the literal, not the named Office constant, so this compiles with no dependency on the Office Object Library being a checked reference
    ' Starting a Word of our own costs a second or two even when Word
    ' is already open, and a large document costs more, so Excel's own
    ' status bar says what is happening rather than leaving a person
    ' looking at a frozen window. Cosmetic, and tolerated if it fails:
    ' unlike the two settings above, nothing hangs or opens unguarded
    ' without it.
    On Error Resume Next
    Application.StatusBar = "Frazaro: reading " & FileNameOnly(path) & " through Word..."
    On Error GoTo cleanup
    ' Every argument below stops Word asking a question on a window
    ' nobody can see: no converter dialog, no encoding dialog, and a
    ' password it will never match instead of a password box.
    Set doc = wordApp.Documents.Open(path, ConfirmConversions:=False, ReadOnly:=True, _
                                     AddToRecentFiles:=False, PasswordDocument:=WORD_NO_PASSWORD, _
                                     NoEncodingDialog:=True)
    ReadWordFile = doc.Content.Text
    doc.Close False
    Set doc = Nothing
    wordApp.Quit False
    Set wordApp = Nothing
    On Error Resume Next
    Application.StatusBar = False
    Exit Function
cleanup:
    d = Err.Description
    n = Err.Number
    On Error Resume Next
    Application.StatusBar = False
    If Not doc Is Nothing Then doc.Close False
    If ours Then wordApp.Quit False
    On Error GoTo 0
    ' 5408 is Word's own "the password is incorrect", which here means
    ' the document has one at all - nobody typed a password. The text
    ' test catches the same refusal worded differently, since a
    ' non-English Word's number is not something to rely on; when
    ' neither matches, Word's own words are reported as they are.
    If n = 5408 Or InStr(1, d, "password", vbTextCompare) > 0 Then
        VLA_Messages.RaiseMsg "ide-word-password", "file", FileNameOnly(path)
    End If
    VLA_Messages.RaiseMsg "ide-word-read-failed", "file", FileNameOnly(path), "detail", d
End Function

' Undo Word's helpful typography: curly quotes back to straight,
' typographic dashes to hyphens, non-breaking spaces to spaces, and
' every flavor of line/page/cell break to a plain line break.
Private Function NormalizeProgramText(ByVal t As String) As String
    t = Replace(t, ChrW$(8220), """")         ' left double quote
    t = Replace(t, ChrW$(8221), """")         ' right double quote
    t = Replace(t, ChrW$(8216), "'")          ' left single quote
    t = Replace(t, ChrW$(8217), "'")          ' right single quote
    t = Replace(t, ChrW$(8211), "-")          ' en dash
    t = Replace(t, ChrW$(8212), "-")          ' em dash
    t = Replace(t, ChrW$(160), " ")           ' non-breaking space
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, Chr$(11), vbLf)            ' Word manual line break
    t = Replace(t, Chr$(12), vbLf)            ' page break
    t = Replace(t, Chr$(7), vbLf)             ' table cell marker
    NormalizeProgramText = t
End Function

' =====================================================================
'  SOP.6: the <Frazaro> tag. A client's SOP stays exactly as they wrote
'  it, and Frazaro reads only the part between <Frazaro> and </Frazaro>.
'  A pre-scan, not a second importer: ImportFromPath runs it on every
'  Load and Reload, after NormalizeProgramText and before PourProgram.
'  The owner's design, 2026-09-25 (the adjudication is SOP.6 in
'  docs/BETA_REARVIEW.md):
'    - No tag line anywhere: the whole document is the program, exactly
'      as before this item. The text comes back untouched.
'    - Otherwise only the tagged sections are, any number of them, in
'      document order. Only the last may be left open, and it then runs
'      to the end of the document - "top half for people, bottom half
'      for Frazaro", with PHP's optional closing ?> as the precedent. A
'      section boundary is a blank line, so no block stays open across
'      prose the program never shows.
'    - A tag sits ALONE on its line, case ignored. A sentence that
'      merely mentions the tag is not one.
'    - Inside the opening tag, the SHAPE says whose a thing is:
'        name=value  the author's: a value the section writes as {name}
'        (any text)  the author's: a note, never read by Frazaro
'        bare token  Frazaro's: a version (0.6.2 - needs that or later)
'                    or a language (espanol - needs that phrasebook)
'      Any other bare token is refused. That keeps every bare word free
'      for a meaning added later, and makes an older Frazaro refuse a
'      newer setting out loud instead of ignoring it. An unused value is
'      fine (a coworker may use it later), and so is a note that says
'      "approved": a note is a label, never a gate.
'    - Values are filled in HERE, at import, so the sheet shows the
'      filled sentence. Every reference stays literal - Check,
'      DeclaredOutputSheet and Rehearse all see what will run - and a
'      value's scope is its own section's text, which no blank line can
'      end. A value goes in as one whole literal, never as raw text: Q3.
'      Final pasted raw would end the sentence at its full stop.
'  Pure: the release and the languages come in as arguments, so all of
'  it is pinned host-free (TestIdeTags).
' =====================================================================

' The pre-scan. text is the document after NormalizeProgramText;
' fileName names it in every refusal and in the first row; release is
' the running Frazaro's version; languages lists the phrasebooks it can
' read, separated by spaces. hasTags comes back True when the document
' has a tag line, and the result is then its tagged sections under a
' first row saying so; False, and the result is text itself.
Public Function VlaIdeFrazaroSections(ByVal text As String, ByVal fileName As String, _
                                      ByVal release As String, ByVal languages As String, _
                                      ByRef hasTags As Boolean) As String
    hasTags = False
    VlaIdeFrazaroSections = text
    Dim lines() As String
    lines = Split(text, vbLf)
    Dim n As Long
    n = UBound(lines)
    If n < 0 Then Exit Function              ' Split("") has no lines, so no tags

    ' Pass 1: which lines are tags - 1 opens a section, 2 closes one. A
    ' malformed tag refuses here, in either kind of document: a line that
    ' plainly meant to be a tag must neither run as an instruction nor be
    ' skipped as prose.
    Dim kinds() As Long
    ReDim kinds(0 To n)
    Dim i As Long
    For i = 0 To n
        kinds(i) = FzTagLineKind(lines(i), fileName, i + 1)
        If kinds(i) <> 0 Then hasTags = True
    Next i
    If Not hasTags Then Exit Function

    ' Pass 2: the sections' shape.
    Dim secOpen() As Long, secClose() As Long
    ReDim secOpen(1 To n + 1)
    ReDim secClose(1 To n + 1)
    Dim k As Long
    Dim inSec As Boolean
    For i = 0 To n
        If kinds(i) = 1 Then
            If inSec Then
                VLA_Messages.RaiseMsg "ide-tag-opener-inside-section", "file", fileName, _
                    "line", i + 1, "open", secOpen(k) + 1
            End If
            k = k + 1
            secOpen(k) = i
            inSec = True
        ElseIf kinds(i) = 2 Then
            If Not inSec Then
                VLA_Messages.RaiseMsg "ide-tag-closer-without-opener", "file", fileName, "line", i + 1
            End If
            secClose(k) = i
            inSec = False
        End If
    Next i
    If inSec Then secClose(k) = n + 1        ' left open: it runs to the end

    ' Pass 3: each section's tag, then its lines with the values filled
    ' in. Row 0 is written last, because it names every section's lines.
    Dim outp() As String
    ReDim outp(0 To n + 2 * k + 1)
    Dim m As Long
    m = 1
    Dim s As Long
    Dim vals As Collection
    Dim tagRow As String
    Dim formDepth As Long
    Dim inQuote As Boolean
    Dim textSeen As Boolean
    For s = 1 To k
        Set vals = New Collection
        tagRow = FzTagReadOpener(lines(secOpen(s)), fileName, secOpen(s) + 1, release, languages, vals)
        If s > 1 Then
            outp(m) = ""                     ' the boundary: it closes any block left open
            m = m + 1
        End If
        If Len(tagRow) > 0 Then
            outp(m) = "# The section " & FzTagRangeText(secOpen(s), secClose(s), n) & ": " & tagRow
            m = m + 1
        End If
        formDepth = 0
        inQuote = False
        For i = secOpen(s) + 1 To secClose(s) - 1
            If vals.Count > 0 Then
                outp(m) = FzTagFillLine(lines(i), vals, formDepth, inQuote, fileName, i + 1, secOpen(s) + 1)
            Else
                outp(m) = lines(i)
            End If
            If VlaIdeHasReadableText(lines(i)) Then textSeen = True
            m = m + 1
        Next i
    Next s
    ' ide-word-no-text's promise again: pouring nothing would clear the
    ' program already on the sheet and leave an empty one in its place.
    If Not textSeen Then
        VLA_Messages.RaiseMsg "ide-tag-sections-empty", "file", fileName
    End If
    outp(0) = FzTagHeaderRow(fileName, secOpen, secClose, k, n, lines)
    ReDim Preserve outp(0 To m - 1)
    VlaIdeFrazaroSections = Join(outp, vbLf)
End Function

' 0 - not a tag; 1 - opens a section; 2 - closes one. A line is tag-LIKE
' when, spaces aside, it begins "<frazaro" or "</frazaro" in any case and
' the name ends there (so <Frazaros> is prose). A tag-like line whose
' LAST character is ">" is a tag, and refuses if anything is wrong with
' it; one with no ">" anywhere is a tag with its end missing, and refuses
' too; one with a ">" somewhere else is a sentence that begins by
' mentioning the tag - "<Frazaro> marks the part that runs." - and is
' left alone.
Private Function FzTagLineKind(ByVal ln As String, ByVal fileName As String, ByVal lineNo As Long) As Long
    Dim t As String
    t = Trim$(Replace(ln, vbTab, " "))
    If Left$(t, 1) <> "<" Then Exit Function
    Dim p As Long
    p = 2
    Do While Mid$(t, p, 1) = " "
        p = p + 1
    Loop
    Dim closing As Boolean
    If Mid$(t, p, 1) = "/" Then
        closing = True
        p = p + 1
        Do While Mid$(t, p, 1) = " "
            p = p + 1
        Loop
    End If
    If VLA_Identity.Fold(Mid$(t, p, 7)) <> "frazaro" Then Exit Function
    If Mid$(t, p + 7, 1) Like "[A-Za-z0-9_-]" Then Exit Function
    If Right$(t, 1) <> ">" Then
        If InStr(p + 7, t, ">") > 0 Then Exit Function
        VLA_Messages.RaiseMsg "ide-tag-malformed", "file", fileName, "line", lineNo, _
            "problem", "it never closes with >"
    End If
    Dim body As String
    body = Trim$(Mid$(t, p + 7, Len(t) - (p + 7)))
    If closing Then
        If Len(body) > 0 Then
            VLA_Messages.RaiseMsg "ide-tag-malformed", "file", fileName, "line", lineNo, _
                "problem", "a closing tag holds nothing but its name, and this one also holds " & body
        End If
        FzTagLineKind = 2
    Else
        If Right$(body, 1) = "/" Then
            VLA_Messages.RaiseMsg "ide-tag-malformed", "file", fileName, "line", lineNo, _
                "problem", "a tag cannot close itself with />, so put </Frazaro> on a line of its own where the section ends"
        End If
        FzTagLineKind = 1
    End If
End Function

' The opening tag's contents, read by shape (see the block header):
' values into vals, keyed by folded name, each an Array(the name as
' written, its literal in an instruction, its text inside quotes); a
' version checked against release; a language against languages; notes
' kept as written. Returns the contents for the section's own row, or ""
' for a bare <Frazaro>.
Private Function FzTagReadOpener(ByVal ln As String, ByVal fileName As String, ByVal lineNo As Long, _
                                 ByVal release As String, ByVal languages As String, _
                                 vals As Collection) As String
    Dim t As String
    t = Trim$(Replace(ln, vbTab, " "))
    ' Past "<", any spaces and the name - FzTagLineKind has proved the
    ' shape, so the first "frazaro" is the name - up to the final ">".
    Dim p As Long
    p = InStr(1, t, "frazaro", vbTextCompare) + 7
    Dim body As String
    body = Mid$(t, p, Len(t) - p)
    Dim nb As Long
    nb = Len(body)
    Dim shown As String, version As String, language As String, lang As String
    Dim tok As String, c As String
    Dim i As Long, st As Long, q As Long, depth As Long
    i = 1
    Do
        Do While Mid$(body, i, 1) = " "
            i = i + 1
        Loop
        If i > nb Then Exit Do
        c = Mid$(body, i, 1)
        If c = "(" Then
            ' A note, to its matching ")" - a note may hold its own pairs.
            st = i
            depth = 0
            Do While i <= nb
                c = Mid$(body, i, 1)
                If c = "(" Then depth = depth + 1
                If c = ")" Then
                    depth = depth - 1
                    If depth = 0 Then Exit Do
                End If
                i = i + 1
            Loop
            If depth <> 0 Then FzTagMalformed fileName, lineNo, "a note that opens with ( never closes with )"
            shown = FzTagJoin(shown, Mid$(body, st, i - st + 1))
            i = i + 1
        ElseIf c = """" Then
            tok = FzTagQuoted(body, i, fileName, lineNo)
            VLA_Messages.RaiseMsg "ide-tag-quoted-name", "file", fileName, "line", lineNo, _
                "text", """" & tok & """"
        ElseIf c = ")" Or c = "=" Then
            FzTagMalformed fileName, lineNo, "a " & c & " stands where a word, a value or a note should begin"
        Else
            st = i
            Do While i <= nb
                c = Mid$(body, i, 1)
                If c = " " Or c = "=" Or c = "(" Or c = ")" Or c = """" Then Exit Do
                i = i + 1
            Loop
            tok = Mid$(body, st, i - st)
            q = i
            Do While Mid$(body, q, 1) = " "
                q = q + 1
            Loop
            If Mid$(body, q, 1) = "=" Then
                i = q + 1
                shown = FzTagJoin(shown, FzTagReadValue(tok, body, i, fileName, lineNo, vals))
            ElseIf tok Like "[0-9]*" Then
                If Len(version) > 0 Then FzTagTwice fileName, lineNo, "a Frazaro version"
                version = tok
                FzTagCheckVersion tok, release, fileName, lineNo
                shown = FzTagJoin(shown, "Frazaro " & tok)
            Else
                lang = FzTagLanguageWord(tok)
                If Len(lang) = 0 Then
                    VLA_Messages.RaiseMsg "ide-tag-unknown-word", "file", fileName, "line", lineNo, "word", tok
                End If
                If Len(language) > 0 Then FzTagTwice fileName, lineNo, "a language"
                language = lang
                If InStr(1, " " & languages & " ", " " & lang & " ") = 0 Then
                    VLA_Messages.RaiseMsg "ide-tag-language-missing", "file", fileName, "line", lineNo, "language", lang
                End If
                shown = FzTagJoin(shown, tok)
            End If
        End If
    Loop
    FzTagReadOpener = shown
End Function

' One name=value, with i just past the "=". A quoted value is text. An
' unquoted one is a single word or number - a number when it reads as
' one, text otherwise - exactly as it would be typed in a sentence, and
' anything that could start or end something else needs the quotes.
Private Function FzTagReadValue(ByVal nm As String, ByVal body As String, ByRef i As Long, _
                                ByVal fileName As String, ByVal lineNo As Long, _
                                vals As Collection) As String
    If Not FzTagIsName(nm) Then
        VLA_Messages.RaiseMsg "ide-tag-bad-name", "file", fileName, "line", lineNo, "name", nm
    End If
    Dim nb As Long
    nb = Len(body)
    Do While Mid$(body, i, 1) = " "
        i = i + 1
    Loop
    If i > nb Then FzTagMalformed fileName, lineNo, nm & "= has no value after the ="
    Dim raw As String, lit As String
    Dim st As Long, j As Long
    If Mid$(body, i, 1) = """" Then
        raw = FzTagQuoted(body, i, fileName, lineNo)
        lit = """" & Replace(raw, """", """""") & """"
    Else
        st = i
        Do While i <= nb
            If Mid$(body, i, 1) = " " Then Exit Do
            i = i + 1
        Loop
        raw = Mid$(body, st, i - st)
        For j = 1 To Len(raw)
            If InStr("""'(){}<>=", Mid$(raw, j, 1)) > 0 Then
                VLA_Messages.RaiseMsg "ide-tag-value-needs-quotes", "file", fileName, "line", lineNo, "name", nm
            End If
        Next j
        If Not FzTagNumberLiteral(raw, lit) Then lit = """" & raw & """"
    End If
    Dim key As String
    key = VLA_Identity.Fold(nm)
    If FzTagHas(vals, key) Then FzTagTwice fileName, lineNo, "the value " & nm
    vals.Add Array(nm, lit, raw), key
    FzTagReadValue = nm & " = " & lit
End Function

' A quoted value at body(i), read the way a sentence reads a string - a
' doubled quote is one quote - and i left just past its closing quote.
Private Function FzTagQuoted(ByVal body As String, ByRef i As Long, _
                             ByVal fileName As String, ByVal lineNo As Long) As String
    Dim r As String, c As String
    i = i + 1
    Do While i <= Len(body)
        c = Mid$(body, i, 1)
        If c = """" Then
            If Mid$(body, i + 1, 1) = """" Then
                r = r & """"
                i = i + 2
            Else
                i = i + 1
                FzTagQuoted = r
                Exit Function
            End If
        Else
            r = r & c
            i = i + 1
        End If
    Loop
    FzTagMalformed fileName, lineNo, "a quote that opens a value never closes"
End Function

' A bare version: three numbers (VlaVersionParse, the same reading a
' phrasebook's requires: line gets), at most the running release.
Private Sub FzTagCheckVersion(ByVal need As String, ByVal release As String, _
                              ByVal fileName As String, ByVal lineNo As Long)
    Dim a As Long, b As Long, c As Long
    If Not VLA.VlaVersionParse(need, a, b, c) Then
        VLA_Messages.RaiseMsg "ide-tag-version-malformed", "file", fileName, "line", lineNo, "version", need
    End If
    If VLA.VlaVersionCompare(release, need) < 0 Then
        VLA_Messages.RaiseMsg "ide-tag-needs-newer", "file", fileName, "line", lineNo, _
            "need", need, "have", release
    End If
End Sub

' The languages a tag can name: the phrasebooks Frazaro ships in
' scripts\polyglotta, by file name, and the spelling a native speaker
' types for three of them. A word here is permanent (SD-4): add, never
' rename. "" when the word is not a language at all.
Private Function FzTagLanguageWord(ByVal w As String) As String
    Dim f As String
    f = VLA_Identity.Fold(w)
    Select Case f
        Case "english", "espanol", "francais", "deutsche", "dansk", "latin", "esperanto", "pirate", "alien"
            FzTagLanguageWord = f
        Case "espa" & ChrW$(241) & "ol"
            FzTagLanguageWord = "espanol"
        Case "fran" & ChrW$(231) & "ais"
            FzTagLanguageWord = "francais"
        Case "deutsch"
            FzTagLanguageWord = "deutsche"
    End Select
End Function

' A value's name: a letter, then letters, digits, "-" and "_" - one of
' the language's own words, never a number.
Private Function FzTagIsName(ByVal nm As String) As Boolean
    If Len(nm) = 0 Then Exit Function
    If Not (Left$(nm, 1) Like "[A-Za-z]") Then Exit Function
    Dim i As Long
    For i = 2 To Len(nm)
        If Not (Mid$(nm, i, 1) Like "[A-Za-z0-9_-]") Then Exit Function
    Next i
    FzTagIsName = True
End Function

' Does an unquoted value read as a number? Digits, an optional leading
' "-", an optional decimal part, and thousands commas only in the strict
' shape EnTokenize swallows (1,500 - never 1,50). lit is the number as
' an instruction gets it: no commas.
Private Function FzTagNumberLiteral(ByVal v As String, ByRef lit As String) As Boolean
    Dim s As String
    s = v
    Dim sign As String
    If Left$(s, 1) = "-" Then
        sign = "-"
        s = Mid$(s, 2)
    End If
    Dim whole As String, frac As String
    Dim p As Long
    p = InStr(s, ".")
    If p > 0 Then
        whole = Left$(s, p - 1)
        frac = Mid$(s, p + 1)
        If Not FzTagAllDigits(frac) Then Exit Function
    Else
        whole = s
    End If
    If InStr(whole, ",") > 0 Then
        Dim groups() As String
        groups = Split(whole, ",")
        If Len(groups(0)) > 3 Or Not FzTagAllDigits(groups(0)) Then Exit Function
        Dim g As Long
        For g = 1 To UBound(groups)
            If Len(groups(g)) <> 3 Or Not FzTagAllDigits(groups(g)) Then Exit Function
        Next g
        whole = Replace(whole, ",", "")
    ElseIf Not FzTagAllDigits(whole) Then
        Exit Function
    End If
    lit = sign & whole
    If p > 0 Then lit = lit & "." & frac
    FzTagNumberLiteral = True
End Function

Private Function FzTagAllDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "[0-9]") Then Exit Function
    Next i
    FzTagAllDigits = True
End Function

Private Function FzTagHas(ByVal col As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col.Item(key)
    FzTagHas = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function FzTagJoin(ByVal a As String, ByVal b As String) As String
    If Len(a) = 0 Then
        FzTagJoin = b
    Else
        FzTagJoin = a & "; " & b
    End If
End Function

' "a, b and c": the names a section's tag gives, for a refusal to list.
Private Function FzTagNameList(ByVal vals As Collection) As String
    Dim r As String
    Dim k As Long
    Dim v As Variant
    For Each v In vals
        k = k + 1
        If k > 1 Then
            If k = vals.Count Then
                r = r & " and "
            Else
                r = r & ", "
            End If
        End If
        r = r & CStr(v(0))
    Next v
    FzTagNameList = r
End Function

' Where a section sits, for the rows that say so. closeIdx past the last
' line means the section was left open.
Private Function FzTagRangeText(ByVal openIdx As Long, ByVal closeIdx As Long, ByVal n As Long) As String
    If closeIdx > n Then
        FzTagRangeText = "from line " & (openIdx + 1) & " to the end"
    Else
        FzTagRangeText = "on lines " & (openIdx + 1) & "-" & (closeIdx + 1)
    End If
End Function

' The program's first row: what was read and what was left out. It is
' there for the change that is easy to miss - one tag added to a
' document that used to be read whole silences everything outside it.
' Only lines with text in them count as left out.
Private Function FzTagHeaderRow(ByVal fileName As String, secOpen() As Long, secClose() As Long, _
                                ByVal k As Long, ByVal n As Long, lines() As String) As String
    Dim covered() As Boolean
    ReDim covered(0 To n)
    Dim s As Long, i As Long, lastIdx As Long
    For s = 1 To k
        lastIdx = secClose(s)
        If lastIdx > n Then lastIdx = n
        For i = secOpen(s) To lastIdx
            covered(i) = True
        Next i
    Next s
    Dim leftOut As Long
    For i = 0 To n
        If Not covered(i) Then
            If Len(Trim$(Replace(lines(i), vbTab, ""))) > 0 Then leftOut = leftOut + 1
        End If
    Next i
    Dim r As String
    If k = 1 Then
        r = "# From """ & fileName & """: Frazaro read only the <Frazaro> section " & _
            FzTagRangeText(secOpen(1), secClose(1), n)
    Else
        r = "# From """ & fileName & """: Frazaro read only the " & k & " <Frazaro> sections ("
        For s = 1 To k
            If s > 1 Then
                If s = k Then
                    r = r & " and "
                Else
                    r = r & ", "
                End If
            End If
            r = r & FzTagRangeText(secOpen(s), secClose(s), n)
        Next s
        r = r & ")"
    End If
    If leftOut = 1 Then
        r = r & ", and left out the document's other line."
    ElseIf leftOut > 1 Then
        r = r & ", and left out the document's other " & leftOut & " lines."
    Else
        r = r & "."
    End If
    FzTagHeaderRow = r
End Function

' One line of a section whose tag gives values, with every {name} filled
' in. It follows EnTokenize's own reading of a line, so a brace is filled
' only where it would otherwise have been read: in an instruction it
' becomes one whole literal ("Q3" or 2026); inside a quoted string, the
' value's text with any quote doubled; inside a 'Sheet name'!A1
' reference, the text. After an unquoted #, the rest of the line is a
' note and is left as written. A raw VLA form - a row whose first
' character is "(" - belongs to the Lisp layer, braces and all, since
' (interpolate "{x}" :x 1) is that layer's own placeholder: it passes
' through untouched, however many rows it takes. formDepth and inQuote
' carry across rows, because both a form and a quoted string can.
Private Function FzTagFillLine(ByVal s As String, ByVal vals As Collection, ByRef formDepth As Long, _
                               ByRef inQuote As Boolean, ByVal fileName As String, _
                               ByVal lineNo As Long, ByVal openNo As Long) As String
    If Not inQuote Then
        If formDepth > 0 Or Left$(LTrim$(Replace(s, vbTab, " ")), 1) = "(" Then
            FzTagFormDepth s, formDepth
            FzTagFillLine = s
            Exit Function
        End If
    End If
    Dim r As String, c As String
    Dim i As Long, q As Long
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If inQuote Then
            If c = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    r = r & """"""
                    i = i + 2
                Else
                    r = r & c
                    inQuote = False
                    i = i + 1
                End If
            ElseIf c = "{" Or c = "}" Then
                r = r & FzTagFillBrace(s, i, vals, 1, fileName, lineNo, openNo)
            Else
                r = r & c
                i = i + 1
            End If
        ElseIf c = "#" Then
            r = r & Mid$(s, i)
            Exit Do
        ElseIf c = """" Then
            r = r & c
            inQuote = True
            i = i + 1
        ElseIf c = "'" Then
            ' Excel's 'Sheet Name'!A1, recognized exactly as EnTokenize
            ' recognizes it: closed on this row, then ! and a word.
            q = InStr(i + 1, s, "'")
            If q > 0 Then
                If Mid$(s, q + 1, 1) = "!" And Mid$(s, q + 2, 1) Like "[A-Za-z0-9_-]" Then
                    r = r & "'" & FzTagFillSheetName(Mid$(s, i + 1, q - i - 1), vals, fileName, lineNo, openNo) & "'"
                    i = q + 1
                Else
                    r = r & c
                    i = i + 1
                End If
            Else
                r = r & c
                i = i + 1
            End If
        ElseIf c = "{" Or c = "}" Then
            r = r & FzTagFillBrace(s, i, vals, 0, fileName, lineNo, openNo)
        Else
            r = r & c
            i = i + 1
        End If
    Loop
    FzTagFillLine = r
End Function

Private Function FzTagFillSheetName(ByVal inner As String, ByVal vals As Collection, _
                                    ByVal fileName As String, ByVal lineNo As Long, _
                                    ByVal openNo As Long) As String
    Dim r As String, c As String
    Dim i As Long
    i = 1
    Do While i <= Len(inner)
        c = Mid$(inner, i, 1)
        If c = "{" Or c = "}" Then
            r = r & FzTagFillBrace(inner, i, vals, 2, fileName, lineNo, openNo)
        Else
            r = r & c
            i = i + 1
        End If
    Loop
    FzTagFillSheetName = r
End Function

' One brace at s(i), and i left past it. "{{" and "}}" stand for the
' brace itself, and a lone "}" is only ever a brace - L-INTERPOLATE's
' own rule, since a "}" is ambiguous only while a placeholder is open.
' "{name}" is the value: mode 0 in an instruction, 1 inside a quoted
' string, 2 inside a 'Sheet name'! reference, which has no way to hold
' an apostrophe at all.
Private Function FzTagFillBrace(ByVal s As String, ByRef i As Long, ByVal vals As Collection, _
                                ByVal mode As Long, ByVal fileName As String, _
                                ByVal lineNo As Long, ByVal openNo As Long) As String
    Dim c As String
    c = Mid$(s, i, 1)
    If Mid$(s, i + 1, 1) = c Then
        FzTagFillBrace = c
        i = i + 2
        Exit Function
    End If
    If c = "}" Then
        FzTagFillBrace = c
        i = i + 1
        Exit Function
    End If
    Dim cl As Long
    cl = InStr(i + 1, s, "}")
    If cl = 0 Then
        VLA_Messages.RaiseMsg "ide-tag-placeholder-malformed", "file", fileName, "line", lineNo, _
            "problem", "a brace opens a placeholder here and never closes it"
    End If
    Dim nm As String
    nm = Mid$(s, i + 1, cl - i - 1)
    If Len(nm) = 0 Then
        VLA_Messages.RaiseMsg "ide-tag-placeholder-malformed", "file", fileName, "line", lineNo, _
            "problem", "the braces here have no name between them"
    End If
    If Not FzTagIsName(nm) Then
        VLA_Messages.RaiseMsg "ide-tag-placeholder-malformed", "file", fileName, "line", lineNo, _
            "problem", "{" & nm & "} holds something that is not a value's name"
    End If
    Dim key As String
    key = VLA_Identity.Fold(nm)
    If Not FzTagHas(vals, key) Then
        VLA_Messages.RaiseMsg "ide-tag-placeholder-unknown", "file", fileName, "line", lineNo, _
            "placeholder", "{" & nm & "}", "open", openNo, "names", FzTagNameList(vals)
    End If
    Dim v As Variant
    v = vals.Item(key)
    If mode = 0 Then
        FzTagFillBrace = CStr(v(1))
    ElseIf mode = 1 Then
        FzTagFillBrace = Replace(CStr(v(2)), """", """""")
    Else
        If InStr(CStr(v(2)), "'") > 0 Then
            VLA_Messages.RaiseMsg "ide-tag-apostrophe-in-sheet-name", "file", fileName, "line", lineNo, _
                "name", CStr(v(0))
        End If
        FzTagFillBrace = CStr(v(2))
    End If
    i = cl + 1
End Function

' Follows a raw VLA form's parentheses through one row the way
' EnTokenize does: a string there (backslash escapes) never spans rows,
' ";" is a comment to the end of the row, and the form is over when its
' depth comes back to zero - the rest of that row is the tokenizer's
' business, not this scan's.
Private Sub FzTagFormDepth(ByVal s As String, ByRef depth As Long)
    Dim i As Long
    Dim c As String
    Dim inS As Boolean
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If inS Then
            If c = "\" Then
                i = i + 1
            ElseIf c = """" Then
                inS = False
            End If
        ElseIf c = """" Then
            inS = True
        ElseIf c = ";" Then
            Exit Do
        ElseIf c = "(" Then
            depth = depth + 1
        ElseIf c = ")" Then
            depth = depth - 1
            If depth <= 0 Then
                depth = 0
                Exit Do
            End If
        End If
        i = i + 1
    Loop
End Sub

Private Sub FzTagMalformed(ByVal fileName As String, ByVal lineNo As Long, ByVal problem As String)
    VLA_Messages.RaiseMsg "ide-tag-malformed", "file", fileName, "line", lineNo, "problem", problem
End Sub

Private Sub FzTagTwice(ByVal fileName As String, ByVal lineNo As Long, ByVal what As String)
    VLA_Messages.RaiseMsg "ide-tag-twice", "file", fileName, "line", lineNo, "what", what
End Sub

' SOP.6: the languages a tag's language word may ask for, read from names
' alone: English always (every edition builds on english.vla), this
' edition's own phrasebook, and every phrasebook this workbook asks to
' load. No file is touched - a remembered path is an arbitrary string
' out of the workbook (SEC.9), and only its name is compared. It is the
' workbook's REQUEST, not this device's answer, and that is enough here:
' the word is a courtesy that turns "Frazaro does not understand" into
' "this needs the Spanish phrasebook", not a gate. A phrasebook this
' device declined still fails every one of its sentences at Check.
Private Function IdeTagLanguages() As String
    Dim r As String
    r = "english " & FzTagFileStem(IdeVocabFileName())
    Dim p As Variant
    For Each p In LoadedPhrasebookPaths(HostBook())
        r = r & " " & FzTagFileStem(CStr(p))
    Next p
    IdeTagLanguages = r
End Function

Private Function FzTagFileStem(ByVal path As String) As String
    Dim nm As String
    nm = FileNameOnly(path)
    If VLA_Identity.Fold(Right$(nm, 4)) = ".vla" Then nm = Left$(nm, Len(nm) - 4)
    FzTagFileStem = VLA_Identity.Fold(nm)
End Function

' Replace the program area: clear old sentences and marks, then one
' line per row from row 2 down (column B is text-formatted, so a line
' can never coerce into a formula).
Private Sub PourProgram(ws As Worksheet, ByVal text As String)
    Dim last As Long
    last = IdeLastRow(ws)
    Dim lastC As Long
    lastC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastC > last Then last = lastC
    If last >= FIRST_ROW Then
        ws.Range(ws.Cells(FIRST_ROW, 2), ws.Cells(last, 2)).ClearContents
        ClearMarks ws, last
    End If

    Dim lines() As String
    lines = Split(text, vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        ws.Cells(FIRST_ROW + i - LBound(lines), 2).Value = lines(i)
    Next
End Sub

' The sheet a program declares as its workspace via "Work on sheet
' <name>." (first declaration wins), or "" when it relies on the
' default. Comment lines don't count.
Private Function DeclaredOutputSheet(ByVal text As String) As String
    Dim lines() As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim i As Long
    Dim ln As String, lc As String
    Dim q As Long
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            lc = VLA_Identity.Fold(ln)
            If Left$(lc, 14) = "work on sheet " Then
                q = 15
                DeclaredOutputSheet = ReadNameAt(lc, ln, q)
                Exit Function
            End If
        End If
    Next
End Function

' One-line environment summary for VlaDiagnostics.
Public Function VlaIdeInfo() As String
    Dim r As String
    On Error Resume Next
    r = "vocab path: " & IdeVocabPath(mayPrompt:=False)   ' SEC.9: a diagnostics string must never stop to ask a question
    Dim hb As Workbook
    Set hb = HostBook()                       ' D1: lazy capture, never ambient
    If Not hb Is Nothing Then
        ' V2: list every program in the book, with its imported file
        ' where one is remembered.
        Dim list As Collection
        Set list = WorkspaceSheets(hb)
        If list.Count = 0 Then
            r = r & " | no workspace sheet yet"
        Else
            Dim wsV As Variant
            For Each wsV In list
                r = r & " | " & wsV.Name & " [" & VlaIdeModuleFor(CStr(wsV.Name)) & "]"
                If Len(Trim$(CStr(wsV.Cells(1, 1).Value))) > 0 Then
                    r = r & " file: " & Trim$(CStr(wsV.Cells(1, 1).Value))
                End If
            Next
        End If
    End If
    r = r & VlaPhrasebookApprovalsReport()
    On Error GoTo 0
    VlaIdeInfo = r
End Function

' SEC.9: every phrasebook decision this device has recorded, in plain
' words, for VlaDiagnostics.
'
' WHY THIS EXISTS, and it is not decoration. A recorded "no" makes a
' phrasebook stop loading, silently and on every command thereafter -
' which is the correct security behaviour and a terrible debugging
' experience, because the only symptom is a sentence that no longer
' resolves. The owner hit exactly that on SEC.9's first live pass: a
' stale "denied" left by an earlier build meant the gate never asked
' again, and from outside the product that is indistinguishable from
' the gate having vanished. A decision a person cannot SEE is a
' decision they cannot correct.
'
' Deliberately read-only. Clearing an approval is a separate action
' with its own consequences (it re-opens a security question), and
' inventing a button for it here would be scope this item did not
' scope. Naming the registry location is what lets a person - or a
' support conversation - act on it today.
Public Function VlaPhrasebookApprovalsReport() As String
    Dim s As Variant
    On Error Resume Next
    s = GetAllSettings("Frazaro", PHRASEBOOK_CONSENT_SECTION)
    On Error GoTo 0
    If IsEmpty(s) Then Exit Function

    Dim r As String, i As Long
    r = vbCrLf & "phrasebook approvals (this device; SEC.9) - registry: " & _
        "HKCU\Software\VB and VBA Program Settings\Frazaro\" & PHRASEBOOK_CONSENT_SECTION
    For i = LBound(s, 1) To UBound(s, 1)
        Dim v As String, decision As String, bar As Long
        v = CStr(s(i, 1))
        bar = InStr(v, "|")
        If bar > 0 Then decision = Mid$(v, bar + 1) Else decision = v
        r = r & vbCrLf & "  " & decision & ": " & CStr(s(i, 0))
        ' An empty digest means the answer was reached without reading
        ' the file - a declined remote path, or a record left by the
        ' pre-fix build. Those never re-ask on their own, so say so
        ' rather than leaving a person to wonder why editing the file
        ' changes nothing.
        If bar <= 1 Then r = r & "   (remembered regardless of the file's contents - editing it will not re-ask)"
    Next i
    VlaPhrasebookApprovalsReport = r
End Function

' SEC.9: the way back out of a "no".
'
' WHY A BUTTON, and why this one rather than a cleverer mechanism. A
' recorded decline is permanent and content-keyed, which is the right
' security shape - the file never loads, and there is no repeating
' dialog to train someone into clicking Yes without reading it. But the
' first version shipped that permanence with no route back, and the two
' routes it CLAIMED to have both turned out not to work for the case
' that matters. Load Phrasebook ADDS a phrasebook on top of the loaded
' base, while a sibling override REPLACES that base, so re-picking a
' denied sibling english.vla refuses with a macro-name collision rather
' than re-approving it - correctly, but uselessly. Editing the file does
' re-open the question, because the record is content-keyed, and nobody
' would ever guess that. Owner-caught across two live passes.
'
' Clears BOTH answers, deliberately, rather than only the declines. A
' person who wants to review what their machine has agreed to should not
' have to trust that this button kept the convenient half; and every
' grant it drops costs exactly one prompt to restore, on a dialog that
' names the file. Erring toward asking again is the safe direction.
Public Sub EnglishIdeForgetPhrasebookApprovals()
    On Error GoTo failed
    CaptureHost

    Dim s As Variant
    On Error Resume Next
    s = GetAllSettings("Frazaro", PHRASEBOOK_CONSENT_SECTION)
    On Error GoTo failed
    If IsEmpty(s) Then
        VlaShowInfo "No phrasebook approvals are recorded on this device yet." & vbCrLf & vbCrLf & _
            "Frazaro asks before using a phrasebook that did not come with it - a grammar file sitting " & _
            "next to your workbook, or one a workbook remembers. Your answers would be listed here."
        Exit Sub
    End If

    Dim n As Long
    n = UBound(s, 1) - LBound(s, 1) + 1
    If MsgBox("Forget every phrasebook answer recorded on this device?" & vbCrLf & vbCrLf & _
              VlaPhrasebookApprovalsReport() & vbCrLf & vbCrLf & _
              "Frazaro will ask again, naming the file, the next time any of these is used. " & _
              "Nothing is deleted from your workbooks and no phrasebook file is changed.", _
              vbYesNo Or vbQuestion Or vbDefaultButton2, _
              "Frazaro - forget phrasebook approvals?") <> vbYes Then Exit Sub

    ' DeleteSetting with only the section name removes the whole section
    ' in one call. Guarded because it raises 5 when the section is
    ' already gone - a second click, or another Excel instance having
    ' just cleared it.
    On Error Resume Next
    DeleteSetting "Frazaro", PHRASEBOOK_CONSENT_SECTION
    On Error GoTo failed

    VlaShowInfo "Forgotten " & n & " phrasebook answer" & IIf(n = 1, "", "s") & "." & vbCrLf & vbCrLf & _
        "Frazaro will ask again, naming the file, the next time one of them is used."
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

' =====================================================================
'  Feedback export: the parse-failure log as one pasteable block -
'  what real people actually tried to say, which is the raw material
'  every future vocabulary decision should be made from.
' =====================================================================

Public Sub EnglishIdeCopyFeedback()
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Set hb = HostBook()
    Dim lg As Worksheet
    On Error Resume Next
    Set lg = hb.Worksheets(LOG_SHEET)
    On Error GoTo failed
    Dim last As Long
    If Not lg Is Nothing Then last = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row
    If lg Is Nothing Or last < 2 Then
        VlaShowInfo "No unknown instructions to share yet. Click me again after you type something I don't understand (a red cell in column C)."
        Exit Sub
    End If

    Dim t As String
    t = "Frazaro diagnostic report - " & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf
    t = t & "Frazaro " & VLA_RELEASE_VERSION & vbCrLf
    t = t & "engine " & VLA_ENGLISH_VERSION & ", " & EnglishRuleCount() & " sentences loaded" & vbCrLf
    ' V2: one workbook can hold several programs; name each with its
    ' imported file where one is remembered.
    On Error Resume Next
    Dim wsV As Variant
    For Each wsV In WorkspaceSheets(hb)
        If Len(Trim$(CStr(wsV.Cells(1, 1).Value))) > 0 Then
            t = t & "program file (" & wsV.Name & "): " & Trim$(CStr(wsV.Cells(1, 1).Value)) & vbCrLf
        End If
    Next
    On Error GoTo failed
    ' SEC.9: the recorded phrasebook decisions belong in the SHIPPED
    ' report, not only in VlaDiagnostics - that Sub lives in
    ' VLA_DevRig.bas, which is not in the build's module list, so a
    ' person running a built add-in could never reach it. Owner-caught:
    ' "'visible' is a strong word". This is the right surface for a
    ' second reason: a declined phrasebook makes sentences stop
    ' resolving, which is exactly what fills the unknown-instruction log
    ' this report needs before it will run at all.
    t = t & VlaPhrasebookApprovalsReport() & vbCrLf
    t = t & "-----" & vbCrLf
    Dim i As Long
    For i = 2 To last
        t = t & Format$(lg.Cells(i, 1).Value, "yyyy-mm-dd hh:nn") & vbTab & _
                "row " & lg.Cells(i, 2).Value & vbTab & _
                lg.Cells(i, 3).Value & vbTab & _
                lg.Cells(i, 4).Value & vbCrLf
    Next

    ' Clipboard via late-bound MSForms DataObject (no reference
    ' needed); if that fails, fall back to a visible Feedback sheet.
    Dim ok As Boolean
    On Error Resume Next
    Dim dob As Object
    Set dob = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dob.SetText t
    dob.PutInClipboard
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo failed
    If ok Then
        VlaShowInfo "Copied " & (last - 1) & " log entr" & IIf(last - 1 = 1, "y", "ies") & _
               " to the clipboard - paste into an email to whoever maintains your vocabulary."
    Else
        Dim fb As Worksheet
        Set fb = GetOrCreateSheet("Feedback")
        fb.Cells.Clear
        fb.Columns(1).NumberFormat = "@"
        Dim lines() As String
        lines = Split(t, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            fb.Cells(i + 1, 1).Value = lines(i)
        Next
        fb.Columns(1).ColumnWidth = 100
        fb.Activate
        VlaShowInfo "Couldn't reach the clipboard - the feedback is on the 'Feedback' sheet; copy it from there."
    End If
    Exit Sub
failed:
    VlaShowError Err.Description
End Sub

Public Sub EnglishIdeClearLog()
    On Error GoTo failed
    CaptureHost
    Dim hb As Workbook
    Set hb = HostBook()
    Dim da As Boolean
    da = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    hb.Worksheets(LOG_SHEET).Delete
    On Error GoTo failed
    Application.DisplayAlerts = da
    VlaShowInfo "Feedback log cleared."
    Exit Sub
failed:
    Application.DisplayAlerts = True
    VlaShowError Err.Description
End Sub
