Attribute VB_Name = "VLA_Build"
Option Explicit
Public Const VLA_BUILD_VERSION As String = "EDITIONMANIFEST.2"
' EDITIONMANIFEST.2: owner correction - VlaBuildAddin's zero-arg call
' now builds EVERY known edition (VlaEditionNames), not just English,
' the same "run everything, report everything, one failure doesn't
' hide the rest" shape VlaSelfTest already uses for tests. The English
' edition's own output also moves to "Frazaro_English.xlam" (was
' "Frazaro.xlam", EDITIONMANIFEST.1's own backward-compatibility
' choice, reversed once "build everything" made a same-named-as-before
' English artifact no longer the natural default anyway). The former
' single-edition body is now VlaBuildOneEdition (Private), called once
' per edition from a loop; a bad edition no longer aborts the whole
' call, matching VlaSelfTest's own "collect every failure" doctrine.
' `VlaBuildAddin "Espanol"` still builds just one edition - a dev-
' iteration convenience, not the primary interface.
' EDITIONMANIFEST.1: VlaBuildAddin gains an Optional edition parameter
' (default "English", so the documented zero-arg release step in
' DEPLOY.md keeps working unchanged) plus VlaEditionVocabPaths/
' VlaEditionOutputName/VlaEditionVocabOverrideName below - the small,
' hardcoded, reviewable table this session's own EDITION-MANIFEST
' scoping pass argued for over a second per-edition `mods` array
' (BETA_ROADMAP2.md's own EDITION LINE text: "not another pair of
' hand-maintained arrays" - F.15's finding, not multiplied). The `mods`
' array itself is untouched and stays edition-invariant: one compiled
' engine, sibling downloads - what varies is only which phrasebook
' file(s) get audited and embedded, and the output filename.
' A non-English edition's own file is an OVERLAY, not a replacement -
' found by reading espanol.vla itself, not guessed: its own header
' ("Carga english.vla PRIMERO") calls macros english.vla defines, so
' english.vla always loads first, for every edition, and every edition
' is bilingual by construction today (every English form still works
' alongside whatever's been translated) - a deliberate tradeoff,
' favoring a native speaker seeing as much of Frazaro in their own
' tongue as the corpus currently offers, over either re-deriving the
' whole ~122-rule corpus per language up front (EDITION-VOCAB/
' MESSAGES/CHROME's own job, grown over time) or refusing English
' sentences the engine already understands correctly. Gate 1 now
' audits the full chain (EnglishAuditPhrasebook's own ParamArray, only
' ever fed one path before now) and the embed step writes one very-
' hidden sheet per file in order (VLAe_Source, VLAe_Source2, ...),
' read back the same way by VLA_IDE.IdeLoadVocab's own new loop - see
' its header there for the reader half. A new embedded marker,
' VLAe_Name, carries the edition's own external-override filename
' (VLA_IDE.IdeVocabPath's own hardcoded "english.vla" search target
' before this pass - a real gap this session's own scoping pass found:
' a stray external english.vla beside a Spanish-edition workbook would
' otherwise silently outrank the embedded Spanish default). The
' `Frazaro_EN_Sheet`/`Frazaro_EN_Runtime` compiled-program naming (IO6.0)
' is a related but separate finding, filed as its own roadmap item
' (EDITION-MODULETAG) rather than folded in here - it's a Compile-time
' concern (the user's own workbook), not a build-time one (this file).
' PROLOG.3: mods array gains VLA_Prolog (new module - the first real
' =PROLOG(...) worksheet function; docs/BETA_ROADMAP2.md's own PROLOG.3
' entry), placed after VLA_Sql, its own family's last-shipped member.
' PROLOG.1: mods array gains VLA_Unify (new module - G-RENDER's one-way
' unifier, hoisted out of VLA_English.bas; docs/BETA_ROADMAP2.md's own
' PROLOG entry), placed beside its sibling VLA_Relation - both ship in
' the add-in from day one, unlike VLA_Datalog/VLA_Sql's still-
' aspirational tranche.
' LX2.0: mods array gains VLA_Messages (SD-2/LX.2's new catalogue module).
' DI3.1: VlaWriteInstallerVersion - writes VLA.VLA_RELEASE_VERSION out to
' installer\version.iss (an #include Frazaro.iss now reads its AppVersion
' from) on every build, so the installer's version number can never again
' silently rot the way it did at a hand-typed "1.0" for the life of this
' project (DI.3's own scoping finding). Soft-failure doctrine, same as
' VlaInjectRibbon: a checkout with no installer\ folder never fails the
' add-in build. See the function's own header for the full reasoning.
' DI2.4: VlaRefreshBetaCopy corrected from opt-in (refresh an existing
' copy only) to unconditional (create it too) - owner-caught,
' immediately after DI2.3 shipped: the Uninstall Frazaro button's own
' successful self-delete removed Frazaro_Beta.xlam entirely, and the
' opt-in gate then meant this convenience died silently and permanently
' the moment that happened, recoverable only by manually redoing the
' exact copy/rename step it exists to eliminate - missing the entire
' point on the very first real use of the two features together. See
' the function's own header for the full reasoning.
' DI2.3: the ribbon gains "Uninstall Frazaro" (VLA_IDE.VlaIdeUninstall)
' - deregisters (VlaUnregisterAutoLoad, from Pass 3), hands off to the
' real Windows uninstaller if this copy was exe-installed, or - for a
' standalone install - genuinely deletes itself via a detached,
' delayed PowerShell process (VLA_IDE.ScheduleSelfDelete), live-tested
' before shipping rather than left as a "delete this by hand"
' deferral: owner pushback correctly named that asking a non-technical
' user to find and delete a file by hand was not a real answer. Closes
' itself via a deferred Application.OnTime call rather than a
' synchronous Close from inside its own running ribbon callback.
' Completes the exit-door symmetry this distribution effort has been
' building toward since Pass 5's own doctrine table.
' DI2.2: VlaRefreshBetaCopy - an opt-in dev convenience that keeps a
' Register-for-Auto-Load test copy (e.g. Frazaro_Beta.xlam) current on
' every build, closing it first if this session has it open. Owner-
' requested after hitting the build-lock trap live; see the function's
' own header for the full reasoning, including a live-tested
' correction to the "locked file" assumption the request was framed
' around - a bare file copy turns out NOT to be blocked by an open
' .xlam at all, so the actual problem it solves is stale in-memory
' state, not a copy failure.
' DI2.1: prelude.vla and english.vla now ride INSIDE the built add-in
' too, the same way V5.3 already carries VLA_Runtime's injectable text
' - very-hidden sheets (VLAp_Source, VLAe_Source), written here while
' the dev project can still read the source files, read back by
' VLA.PreludeMacros/VLA_IDE.IdeLoadVocab (via the new shared
' VLA.VlaEmbeddedText) only when no external file is found beside the
' workbook or the add-in. The external-file lookup is UNCHANGED and
' still tried first in both readers - a user's or org's own vocabulary
' still overrides the built-in one; bundling only removes the "you must
' also ship these two files" requirement for the common case, per
' DEPLOY.md's DI.2 Pass 1. prelude.vla missing at build time still
' fails the build (every compile needs it, so an add-in with neither an
' external copy nor an embedded one is broken by construction);
' english.vla missing at build time is a soft note, matching Gate 1's
' own existing tolerance for a phrasebook-less build just above.
' DI1.0 (tried here, moved elsewhere): Authenticode-signing the built
' .xlam directly was attempted in this module (the same shell-to-
' PowerShell/soft-failure shape VlaInjectRibbon already established)
' and found NOT to work, empirically - Set-AuthenticodeSignature on a
' real .xlam on this machine fails with "The form specified for the
' subject is not one supported or known by the specified trust
' provider": Office's OOXML package format has no registered
' Authenticode Subject Interface Package, a structural fact (Office
' documents use a different, XML-DSig-based signature scheme than
' PE/script Authenticode), not a missing-cert or missing-tool problem
' - confirmed by the same mechanism signing a plain .ps1 file cleanly
' on this same machine. DI.1's actual signing step now lives at the
' installer stage instead (installer/sign_installer.ps1, run after
' ISCC.exe builds FrazaroSetup.exe) - see DEPLOY.md's "Code signing"
' section for what that does and does not buy.
' F5.0: VlaFrame (F.5's new class module) added to the shipped module
' list - VLA.bas's VlaPushContext/VlaPopContext reference it by type, so
' an add-in built without it would fail Debug > Compile in the fresh
' workbook this sub assembles. Export/Import both hardcoded ".bas";
' correct for every prior entry (all standard modules) but wrong for a
' class, so both now ask the live VBComponent's own .Type instead of
' guessing from the name - vbext_ct_ClassModule is 2 in the VBIDE type
' library, hardcoded rather than adding an Extensibility reference for
' one constant (this project takes no non-Excel library references).
' LX3.0: VLA_Identity added to the shipped module list (below).
' IN1.0: VLA_HeadTable added to the shipped module list (below).
' F3.0: prelude.vla (VLA.bas's PreludeVlaPath, ThisWorkbook-relative,
' same "scripts/ or flat" shape as english.vla above) is now a
' REQUIRED companion file - every compile needs it, not just phrasebook
' translation. No new gate added here: Gate 1's own audit already
' transpiles every test: line, which already needs the prelude, so a
' missing prelude.vla surfaces through the existing gate with
' PreludeVlaPath's own clear error rather than a generic file-I/O
' crash. Ship prelude.vla beside the built add-in the same way
' english.vla already must be.

' =====================================================================
'  VLA_Build - assembles the distributable add-in from the modules in
'  THIS workbook (the development workbook). Run VlaBuildAddin from
'  the Immediate window; it produces Frazaro.xlam next to this
'  workbook. One manual step remains (locking the project) - see
'  DEPLOY.md.
'
'  Requires "Trust access to the VBA project object model".
' =====================================================================

' EDITION-MANIFEST: the edition table. Each edition names only the two
' things BETA_ROADMAP2.md's own EDITION LINE text says vary - which
' phrasebook file(s) get audited/embedded, in load order, and the
' output filename - never the module list (VlaBuildAddin's own `mods`,
' below, stays a single array for every edition). Add a new edition by
' adding one Case to each of these three functions; nothing else in
' this file names an edition by hand.
' EDITIONMANIFEST.3: english.vla moved to scripts\polyglotta\ (owner,
' this session - "to keep language files together" with its seven
' dialect siblings, which already lived there). Every OTHER path in
' this codebase that reads english.vla by name already tolerated this
' via a three-candidate FindDevFile helper (VLA_Tests.bas/
' VLA_Interpreter.bas); this one hardcoded a single flat location and
' broke the moment the file moved - found by checking, not assumed
' fine because nothing raised yet (Gate 1's own soft-tolerance for a
' missing base file would have silently shipped an edition with no
' embedded phrasebook at all, the exact failure mode this file's own
' Gate 1 exists to prevent for other cases).
Private Function VlaEditionVocabPaths(ByVal edition As String) As Variant
    Select Case LCase$(edition)
        Case "english"
            VlaEditionVocabPaths = Array(ThisWorkbook.Path & "\scripts\polyglotta\english.vla")
        Case "espanol"
            ' espanol.vla is an OVERLAY (calls macros english.vla
            ' defines - confirmed by reading the file itself, not
            ' guessed), so english.vla always loads first.
            VlaEditionVocabPaths = Array( _
                ThisWorkbook.Path & "\scripts\polyglotta\english.vla", _
                ThisWorkbook.Path & "\scripts\polyglotta\espanol.vla")
        Case Else
            Err.Raise 5, "VLA-Build", "Build refused - unknown edition """ & edition & """."
    End Select
End Function

Private Function VlaEditionOutputName(ByVal edition As String) As String
    Select Case LCase$(edition)
        Case "english"
            VlaEditionOutputName = "Frazaro_English.xlam"
        Case "espanol"
            VlaEditionOutputName = "Frazaro_Espanol.xlam"
        Case Else
            Err.Raise 5, "VLA-Build", "Build refused - unknown edition """ & edition & """."
    End Select
End Function

' EDITIONMANIFEST.2: every edition VlaBuildAddin knows how to build,
' in the order it builds them - the one place a new edition gets added
' to the DEFAULT (build-everything) run. The three functions above
' still take an edition NAME rather than reading this list themselves,
' so a single-edition dev build (VlaBuildAddin "Espanol") never has to
' touch this array.
Private Function VlaEditionNames() As Variant
    VlaEditionNames = Array("English", "Espanol")
End Function

' The bare filename VLA_IDE.IdeVocabPath searches for beside a user's
' own workbook to override this edition's embedded default - matches
' the edition's own overlay file so a customization convention that
' already works for English ("your own english.vla always wins") reads
' the same way in a second language, rather than every edition still
' watching for a file literally named "english.vla".
Private Function VlaEditionVocabOverrideName(ByVal edition As String) As String
    Select Case LCase$(edition)
        Case "english"
            VlaEditionVocabOverrideName = "english.vla"
        Case "espanol"
            VlaEditionVocabOverrideName = "espanol.vla"
        Case Else
            Err.Raise 5, "VLA-Build", "Build refused - unknown edition """ & edition & """."
    End Select
End Function

' EDITIONMANIFEST.2: builds every known edition by default - the same
' "run everything, report everything, one failure doesn't hide the
' rest" shape VlaSelfTest already uses for tests, applied to editions.
' Immediate window: `VlaBuildAddin` alone now builds
' Frazaro_English.xlam AND Frazaro_Espanol.xlam in one call.
' `VlaBuildAddin "Espanol"` still builds just that one edition, a dev-
' iteration convenience (rebuild the Spanish edition alone while
' working on espanol.vla, without waiting on English too).
Public Sub VlaBuildAddin(Optional ByVal onlyEdition As String = "")
    Dim editions As Variant
    If Len(onlyEdition) > 0 Then
        editions = Array(onlyEdition)
    Else
        editions = VlaEditionNames()
    End If

    Dim report As String
    Dim builtCount As Long, failedCount As Long
    Dim ed As Variant
    For Each ed In editions
        Dim editionNote As String
        Dim ok As Boolean
        ok = VlaBuildOneEdition(CStr(ed), editionNote)
        If ok Then
            builtCount = builtCount + 1
        Else
            failedCount = failedCount + 1
        End If
        report = report & editionNote & vbCrLf & vbCrLf
        Debug.Print "===== " & ed & ": " & IIf(ok, "built", "FAILED") & " ====="
    Next ed

    ' DI.3a: keep the installer's AppVersion synced to the real release
    ' version instead of a hand-typed literal - see VlaWriteInstallerVersion's
    ' own header for why this fixes a finding, not a hypothetical. Once
    ' per call, not per edition - it names the release, not one edition.
    Dim versionNote As String
    versionNote = VlaWriteInstallerVersion()
    Debug.Print versionNote

    MsgBox builtCount & " edition(s) built, " & failedCount & " failed." & vbCrLf & vbCrLf & _
           report & _
           versionNote & vbCrLf & vbCrLf & _
           "Remaining manual steps, per .xlam built:" & vbCrLf & _
           "  1. Open it." & vbCrLf & _
           "  2. Lock its VBA project with a password (Tools > VlaTools Properties > Protection)." & vbCrLf & _
           "  3. Save." & vbCrLf & _
           "See DEPLOY.md for details.", IIf(failedCount = 0, vbInformation, vbExclamation), "VLA Build"
End Sub

' EDITIONMANIFEST.2: one edition's entire build, pulled out of
' VlaBuildAddin so a failure in one edition (a phrasebook defect, a
' missing file) can be caught, reported, and moved past - the loop
' above keeps building the REST of the editions, the same "collect
' every failure, don't stop at the first" doctrine VlaSelfTest already
' uses for tests. Returns True with a report in `note` on success,
' False with the error description in `note` on failure; never raises
' out to its own caller.
Private Function VlaBuildOneEdition(ByVal edition As String, ByRef note As String) As Boolean
    On Error GoTo failed

    Dim vocabFiles As Variant
    vocabFiles = VlaEditionVocabPaths(edition)

    ' Gate 1: audit the full phrasebook CHAIN that will actually ship,
    ' in order - EnglishAuditPhrasebook's own ParamArray already
    ' supported this (only ever fed one path before now). Loading runs
    ' every test: line across the whole chain (a failing proof raises
    ' here) and then reports duplicate signatures or operator-word
    ' hazards - either refuses the build. The base file missing is the
    ' pre-existing soft-tolerance case (an incomplete checkout); an
    ' edition's own OVERLAY file missing is new and refuses hard -
    ' silently shipping "Frazaro_Espanol.xlam" with no Spanish content
    ' would be a worse failure than refusing to build it at all.
    Dim auditNote As String
    Dim missingAt As Long
    missingAt = -1
    Dim vi As Long
    For vi = LBound(vocabFiles) To UBound(vocabFiles)
        If Len(Dir$(CStr(vocabFiles(vi)))) = 0 Then
            missingAt = vi
            Exit For
        End If
    Next
    If missingAt = -1 Then
        Dim auditR As String
        auditR = EnglishAuditPhrasebook(vocabFiles)
        If Len(auditR) > 0 Then
            Err.Raise 5, "VLA-Build", "Build refused - the phrasebook has authoring defects:" & _
                vbCrLf & vbCrLf & auditR & vbCrLf & "Fix these and rebuild."
        End If
        auditNote = "Phrasebook audited clean (" & edition & "):" & vbCrLf & Join(vocabFiles, vbCrLf)
    ElseIf missingAt = LBound(vocabFiles) Then
        auditNote = "NOTE: no phrasebook found beside this workbook - nothing was audited."
    Else
        Err.Raise 5, "VLA-Build", "Build refused - " & edition & " edition's own overlay phrasebook not found: " & _
            CStr(vocabFiles(missingAt))
    End If
    Debug.Print auditNote

    Dim mods As Variant
    ' V5.3: VLA_Runtime ships too - the helper zoo the generated
    ' program calls. The add-in needs it so its own tooling resolves
    ' the helpers, and the compile step injects a copy into the
    ' user's workbook (VlaInjectRuntime) so generated code does too.
    ' LX3.0: VLA_Identity ships too - every other module here now
    ' folds identity through it (R6/SD-8). It never crosses the
    ' injection boundary itself (VLA_Runtime keeps its own duplicate
    ' fold for that, per R7), so it is add-in/dev-project only.
    ' F5.0: VlaFrame ships too - VLA.bas's VlaPushContext/VlaPopContext
    ' name it as a type, so a build without it fails Debug > Compile.
    ' IN.6: VLA_Interpreter ships too - left off this list since IN.0.5
    ' built it as a dev-rig-only throwaway skeleton ("not in
    ' VLA_Build.bas's shipped module list", BETA_ROADMAP.md's own words at
    ' the time), a decision nothing revisited as it grew into the real
    ' evaluator (IN.2-IN.11). IN.6 is the first shipped module
    ' (VLA_IDE's own InterpretProgram) to actually reference
    ' VLA_Interpreter.VlaInterpret by name, which is exactly what
    ' surfaced the gap: a fresh-built add-in failing Debug > Compile
    ' with "Variable not defined" on that line, the same failure shape
    ' F5.0/LX3.0/IN1.0 already document above for the same root cause.
    ' IN.7: VLA_Events and VLA_EventSink ship too - VLA_IDE's
    ' InterpretProgram now references VLA_Events.VlaRegisterSheetChangeHandler
    ' by name (the same F5.0/LX3.0/IN1.0/IN.6 failure shape this
    ' comment block already documents above: a shipped module calling
    ' an unshipped one fails Debug > Compile in the fresh workbook,
    ' not at write time). VLA_EventSink is this project's first class
    ' module besides VlaFrame - the .Type-driven export/import logic
    ' below already generalizes to it with no further changes.
    ' CLI.0: frmCLI ships too - VLA_IDE.VlaOpenCli (below) references it
    ' by name, the same shape as every prior "a shipped module calls an
    ' unshipped one" gap this comment block already documents (F5.0,
    ' LX3.0, IN1.0, IN.6, IN.7). A UserForm export/import needs no new
    ' mechanism here - VBIDE's own Export/Import already carry a form's
    ' paired .frx alongside its .frm without being told to; only the
    ' extension-by-.Type branch just below needed a third case.
    ' LX2.0: VLA_Messages ships too - a Layer-0 module (no dependencies,
    ' same as VLA_Identity beside it) so it sits with VLA_Identity ahead
    ' of everything that will come to call it.
    ' SEC.8: VLA_Provenance sits next to VLA_Loader, whose ADODB
    ' byte-reading idiom it reuses to read the Mark-of-the-Web. Standard
    ' modules have no load order in VBA, so its position here is for a
    ' reader, not the compiler - but it MUST be in this list: the shipped
    ' set is read straight from this array by check_raise_ratchet.ps1,
    ' and an add-in built without it would dispatch external effect with
    ' the provenance gate silently absent.
    ' PNTH.0 (found 2026-09-08, latent since 0.5.0's initial import):
    ' VlaSlice ships too, and always should have. It is the NINTH
    ' instance of the exact failure this comment block has been
    ' documenting since F5.0, and the first one found mechanically
    ' rather than by a person hitting a compile error - by
    ' tools/check_devrig_mods_parity.ps1, written for the SEC.8
    ' recurrence and red on its first run. VlaSlice was in
    ' VLA_DevRig.bas's array but not this one, so the dev workbook
    ' compiled it while a freshly built add-in would not have had the
    ' class at all - and VLA.bas's own ListTail names it as a type
    ' ("Dim sl As New VlaSlice", "Dim srcSlice As VlaSlice"), which is
    ' word-for-word F5.0's rationale for shipping VlaFrame. It sits
    ' beside VlaFrame here for the same reason it does there.
    mods = Array("VLA_Identity", "VLA_Messages", "VLA_HeadTable", "VLA", "VlaFrame", "VlaSlice", "VLA_Loader", "VLA_Provenance", "VLA_English", "VLA_SentenceEngine", "VLA_Runtime", "VLA_Interpreter", "VLA_Events", "VLA_EventSink", "VLA_IDE", "VLA_Lint", "VLA_Unify", "VLA_Relation", "VLA_Datalog", "VLA_Sql", "VLA_Prolog", "VLA_Browser", "frmCLI")

    ' Export the current, in-project versions - the build always ships
    ' exactly what the dev workbook contains. (VLA_English and its
    ' relocated VLA_Runtime both come from the dev project here.)
    Dim tmp As String
    tmp = Environ$("TEMP")
    Dim i As Long
    Dim fp As String
    Dim srcComp As Object
    Dim ext As String
    For i = LBound(mods) To UBound(mods)
        Set srcComp = ThisWorkbook.VBProject.VBComponents(mods(i))
        ' F5.0: every prior entry is a standard module (.bas); VlaFrame
        ' is a class, so the extension now follows the component's own
        ' .Type instead of being hardcoded - 2 is vbext_ct_ClassModule
        ' (VBIDE's own enum), a literal because this project takes no
        ' Extensibility library reference for one constant, matching
        ' the late-bound Object typing every VBComponents call here
        ' already uses.
        ext = ".bas"
        If srcComp.Type = 2 Then ext = ".cls"
        If srcComp.Type = 3 Then ext = ".frm"   ' CLI.0: vbext_ct_MSForm
        fp = tmp & "\" & mods(i) & ext
        If Len(Dir$(fp)) > 0 Then Kill fp
        srcComp.Export fp
    Next

    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.VBProject.Name = "VlaTools"           ' never ship "VBAProject"
    For i = LBound(mods) To UBound(mods)
        fp = tmp & "\" & mods(i) & ".bas"
        If Len(Dir$(fp)) = 0 Then fp = tmp & "\" & mods(i) & ".cls"   ' F5.0: VlaFrame
        If Len(Dir$(fp)) = 0 Then fp = tmp & "\" & mods(i) & ".frm"   ' CLI.0: frmCLI
        wb.VBProject.VBComponents.Import fp
    Next

    ' The add-in's own open/close handlers manage the menu. IN.7: also
    ' own VLA_EventSink's lifetime here - a module-level WithEvents
    ' object only keeps hearing Application events for as long as
    ' something holds a reference to it, so ThisWorkbook (the add-in
    ' itself, loaded for the whole Excel session) is where that
    ' reference has to live, the same way it already owns the menu's.
    ' CLI.0: Application.OnKey, not a Windows API hook - dispatch by
    ' macro NAME, which is why this keeps working no matter which
    ' workbook is active (the whole point of an always-loaded add-in
    ' owning it). Registered/unregistered in the same open/close
    ' bracket as the menu and event sink for the same reason: nothing
    ' this add-in hooks should outlive the add-in itself. Unregister
    ' passes the key with no procedure name, which restores Excel's
    ' own (unbound) behavior for Ctrl+Shift+` rather than leaving a
    ' dangling reference to a macro that's about to disappear.
    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString _
        "Private mVlaEventSink As VLA_EventSink" & vbCrLf & _
        "Private Sub Workbook_Open()" & vbCrLf & _
        "    VlaAddinMenu" & vbCrLf & _
        "    Set mVlaEventSink = New VLA_EventSink" & vbCrLf & _
        "    Application.OnKey ""^+`"", ""VlaOpenCli""" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
        "    Application.OnKey ""^+`""" & vbCrLf & _
        "    VlaAddinMenuRemove" & vbCrLf & _
        "    Set mVlaEventSink = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf

    ' V5.3: write the runtime helpers' injectable text into a very-
    ' hidden sheet of the add-in. The shipped project gets LOCKED
    ' (DEPLOY's manual step), and component access to a protected
    ' project raises - so at run time VlaInjectRuntime cannot read
    ' its own module inside the locked add-in; it falls back to this
    ' sheet, written now, while the dev project is still readable.
    Dim rtText As String
    rtText = VlaRuntimeInjectText()
    If Len(Trim$(rtText)) = 0 Then
        Err.Raise 5, "VLA-Build", "Build refused - VLA_Runtime's injectable text came back empty"
    End If
    EmbedTextAsSheet wb, VlaRuntimeSheetName(), rtText

    ' DI2.1: prelude.vla, same treatment, same reason - the shipped
    ' add-in's own copy of core language infrastructure the locked
    ' project can no longer read a file for once distributed alone.
    ' Candidate order matches VLA.PreludeVlaPath exactly (flat beside
    ' the workbook, then its scripts folder) so the embedded copy is
    ' whichever one the dev workbook itself would actually load.
    Dim preludePath As String
    preludePath = ThisWorkbook.Path & "\prelude.vla"
    If Len(Dir$(preludePath)) = 0 Then preludePath = ThisWorkbook.Path & "\scripts\prelude.vla"
    If Len(Dir$(preludePath)) = 0 Then
        Err.Raise 5, "VLA-Build", "Build refused - prelude.vla not found beside this workbook or in its scripts folder. Every VLA compile needs it, embedded or external."
    End If
    EmbedTextAsSheet wb, "VLAp_Source", VLA_Loader.VlaReadFile(preludePath)

    ' EDITIONMANIFEST.1: the phrasebook CHAIN, one very-hidden sheet per
    ' file in load order (VLAe_Source, VLAe_Source2, ...) - the same
    ' "one EnglishLoadVocabulary call per file" shape
    ' EnglishLoadVocabulary's own header already documents as ordinary
    ' multi-file loading, just embedded instead of read fresh each
    ' time. Reader side: VLA_IDE.IdeLoadVocab's own loop. A missing
    ' base file is the pre-existing soft-tolerance case; an edition's
    ' own overlay missing already refused the build above, so every
    ' path here is confirmed to exist. VLAe_Name carries the filename
    ' VLA_IDE.IdeVocabPath should search for to let a user's or org's
    ' own file override this edition's embedded default - see that
    ' function's own header for the gap this closes.
    Dim bundleNote As String
    bundleNote = "Embedded (" & edition & "):" & vbCrLf & "prelude.vla"
    If missingAt = -1 Then
        For vi = LBound(vocabFiles) To UBound(vocabFiles)
            Dim embedSheet As String
            embedSheet = "VLAe_Source"
            If vi > LBound(vocabFiles) Then embedSheet = embedSheet & CStr(vi - LBound(vocabFiles) + 1)
            EmbedTextAsSheet wb, embedSheet, VLA_Loader.VlaReadFile(CStr(vocabFiles(vi)))
            bundleNote = bundleNote & ", " & Dir$(CStr(vocabFiles(vi)))
        Next
        EmbedTextAsSheet wb, "VLAe_Name", VlaEditionVocabOverrideName(edition)
    Else
        bundleNote = bundleNote & " (phrasebook NOT embedded - none found to embed; still needs distributing separately)"
    End If

    wb.IsAddin = True
    Dim outPath As String
    ' V5.1 (owner revision): the artifact wears the PRODUCT's name -
    ' Frazaro is what users install; "VLA English" names one layer of
    ' its engine. The VBProject inside stays "VlaTools" (the VBE-
    ' facing name of the tooling, and what DEPLOY's locking step
    ' references); commands stay name-independent (OnAction binds
    ' through ThisWorkbook.Name), so only the file and its docs move.
    ' EDITIONMANIFEST.1/2: the filename itself is the one other
    ' edition-specific thing (VlaEditionOutputName) - "Frazaro_English.xlam"
    ' or "Frazaro_Espanol.xlam".
    outPath = ThisWorkbook.Path & "\" & VlaEditionOutputName(edition)
    Application.DisplayAlerts = False
    wb.SaveAs outPath, 55                    ' 55 = xlOpenXMLAddIn
    Application.DisplayAlerts = True
    wb.Close False

    ' V5: inject the Frazaro ribbon tab into the saved package (the
    ' file must be closed first - Office reads customUI parts only
    ' at open). Soft-failure doctrine: a ribbon problem never fails
    ' the build - the legacy Add-ins menu remains as the fallback
    ' surface either way.
    Dim ribbonNote As String
    ribbonNote = VlaInjectRibbon(outPath)
    Debug.Print ribbonNote

    ' Personal dev convenience, opt-in - see VlaRefreshBetaCopy's own
    ' header for the full reasoning (DEPLOY.md's "One trap, now
    ' reachable two ways" is the incident that motivated it).
    ' EDITIONMANIFEST.1: English-edition only - Frazaro_Beta.xlam is a
    ' personal English dev/auto-load shortcut; refreshing it from a
    ' non-English build would silently swap the owner's own live test
    ' copy out from under them.
    Dim betaNote As String
    If LCase$(edition) = "english" Then
        betaNote = VlaRefreshBetaCopy(outPath)
        If Len(betaNote) > 0 Then Debug.Print betaNote
    End If

    ' EDITIONMANIFEST.2: assembled into `note` (this edition's own
    ' block of the combined report) instead of its own MsgBox -
    ' VlaBuildAddin shows ONE final summary across every edition, the
    ' same "collect everything, report once" shape VlaSelfTest already
    ' uses. betaNote (the Frazaro_Beta.xlam dev-convenience refresh,
    ' English-only) folds in only when non-empty.
    note = edition & " edition built:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           auditNote & vbCrLf & vbCrLf & _
           bundleNote & vbCrLf & vbCrLf & _
           ribbonNote
    If Len(betaNote) > 0 Then note = note & vbCrLf & vbCrLf & betaNote
    VlaBuildOneEdition = True
    Exit Function
failed:
    Application.DisplayAlerts = True
    note = edition & " edition FAILED: " & Err.Description
    VlaBuildOneEdition = False
End Function

' Personal dev convenience: keeps a "Frazaro_Beta.xlam" test copy
' beside this workbook current with every build, so it can be
' registered for Excel's own auto-load (VLA_IDE's Register-for-Auto-
' Load button - see DEPLOY.md's "One trap, now reachable two ways" for
' the incident that motivated the whole convenience) and just work,
' with no manual copy/rename step ever.
'
' UNCONDITIONAL, not opt-in - the first version only refreshed an
' EXISTING copy and did nothing if the file was ever missing (the
' stated reason: "a fresh checkout should never see it appear on its
' own"). Owner-caught, correctly: that meant the file's own
' disappearance - which happens for a completely ordinary reason, DI.2
' Pass 4's own Uninstall Frazaro button successfully deleting it - left
' this convenience silently dead with no way back except manually
' redoing the exact copy/rename dance it exists to replace, which
' defeats the entire point. This module is dev-only and never ships
' (not in VlaBuildAddin's own `mods` array), so the "don't surprise a
' fresh checkout" concern priced in a cost (silent, permanent failure
' the moment the file is ever removed by anything) far higher than the
' benefit (nobody sees an extra file appear). Now creates the file if
' it's missing, for any reason, every time.
'
' Closes the currently-open copy first, if this session has it open -
' NOT because the file is locked against being overwritten (live-
' tested before writing this: FileCopy succeeds over a .xlam another
' workbook in this same Excel instance already has open - Office uses
' its own advisory ~$ lock-file convention, not an exclusive OS
' handle, so a bare file copy is invisible to it). The real reason to
' close first: overwriting the bytes on disk underneath an
' already-loaded session does NOT retroactively update that session's
' own in-memory VBA - it would keep running the stale build until
' manually closed and reopened, silently. Closing first guarantees the
' next thing that opens Frazaro_Beta.xlam - in this session or the
' next Excel launch - reads the fresh copy, not stale state. If it's
' open in a DIFFERENT Excel process instead, this cannot reach it at
' all (Workbooks() only sees this session's own open books) - that
' process's copy stays stale until closed and reopened by hand.
'
' Returns a one-line note for the build report; never fails the build.
Public Function VlaRefreshBetaCopy(ByVal builtPath As String) As String
    Const BETA_COPY_NAME As String = "Frazaro_Beta.xlam"
    Dim betaPath As String
    betaPath = ThisWorkbook.Path & "\" & BETA_COPY_NAME
    Dim existedBefore As Boolean
    existedBefore = (Len(Dir$(betaPath)) > 0)

    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(BETA_COPY_NAME)
    On Error GoTo 0
    If Not wb Is Nothing Then wb.Close False

    On Error Resume Next
    FileCopy builtPath, betaPath
    Dim failed As Boolean
    failed = (Err.Number <> 0)
    On Error GoTo 0

    If failed Then
        ' A genuine I/O failure (permissions, disk full, a read-only
        ' attribute) - being open elsewhere does NOT cause this;
        ' FileCopy overwrites an open .xlam's bytes on disk regardless
        ' (live-tested, see this function's own header note).
        VlaRefreshBetaCopy = BETA_COPY_NAME & " NOT created/refreshed (" & Err.Description & _
            "); copy by hand: " & builtPath & " -> " & betaPath
    ElseIf Not existedBefore Then
        VlaRefreshBetaCopy = BETA_COPY_NAME & " created. Click Register for Auto-Load on it " & _
            "(once) if you want it to auto-load."
    ElseIf wb Is Nothing Then
        ' Nothing was open in THIS session to close - either it
        ' genuinely wasn't open anywhere (the common case, and the
        ' copy above is simply correct), or it's open in a different
        ' Excel process this code cannot reach or detect - flagged as
        ' a possibility rather than silently assumed fine either way.
        VlaRefreshBetaCopy = BETA_COPY_NAME & " refreshed. (If it's open in a different Excel " & _
            "process, that session is now running stale code until closed and reopened.)"
    Else
        VlaRefreshBetaCopy = BETA_COPY_NAME & " refreshed from the fresh build (was open in this " & _
            "session - closed first)."
    End If
End Function

' DI.3a: keeps installer\version.iss (an Inno Setup preprocessor file,
' `#define MyAppVersion "..."`) synced to VLA.VLA_RELEASE_VERSION on every
' build. Frazaro.iss #includes it instead of carrying its own literal -
' the fix for a real, confirmed finding: AppVersion sat hardcoded at "1.0"
' and untouched for the installer's entire life, because there was never a
' single place a human had to remember to update. This function is that
' single place now; a human still decides WHEN to bump VLA_RELEASE_VERSION
' (a release is a judgment call), but once bumped, every build downstream
' of it - including the installer - picks it up with no second edit.
'
' Soft-failure doctrine, same as VlaInjectRibbon: a checkout with no
' installer\ folder (someone building just the add-in, never the
' installer) never fails the add-in build - this is a best-effort side
' effect, not a gate. Returns a one-line note for the build report.
Public Function VlaWriteInstallerVersion() As String
    On Error GoTo softly
    Dim installerDir As String
    installerDir = ThisWorkbook.Path & "\installer"
    If Len(Dir$(installerDir, vbDirectory)) = 0 Then
        VlaWriteInstallerVersion = "installer\version.iss NOT written (no installer\ folder found here)."
        Exit Function
    End If
    WriteTextFile installerDir & "\version.iss", _
        "; Auto-generated by VlaBuildAddin (DI.3a) from VLA.VLA_RELEASE_VERSION - do not edit by hand." & vbCrLf & _
        "#define MyAppVersion """ & VLA_RELEASE_VERSION & """"
    VlaWriteInstallerVersion = "installer\version.iss written:" & vbCrLf & "MyAppVersion=" & VLA_RELEASE_VERSION
    Exit Function
softly:
    VlaWriteInstallerVersion = "installer\version.iss NOT written (" & Err.Description & ")."
End Function

' =====================================================================
'  V5: the ribbon. Office reads a custom ribbon from a customUI part
'  inside the .xlam package (an OPC zip) - a build artifact, not a
'  code-model change, exactly as the audit predicted. Pure-VBA zip
'  editing (Shell32 CopyHere) is asynchronous and flaky, and a
'  hand-authored template package cannot be verified blind, so the
'  injection rides PowerShell's real zip API (System.IO.Compression),
'  generated and invoked synchronously by the build; the dev machine
'  is Windows by definition (the transpiler already is). If the
'  injection cannot run, the build still succeeds and says so - the
'  legacy CommandBars menu remains the fallback surface, so both
'  surfaces ship and neither is load-bearing alone.
' =====================================================================

' The customUI14 XML (Office 2010+). Label-only buttons by design:
' imageMso names cannot be verified blind, and a wrong one degrades
' the ribbon; labels degrade nothing. One callback, dispatch by id -
' the ids are pinned by the self-test against VLA_IDE's dispatcher.
Public Function VlaRibbonXml() As String
    Dim q As String
    q = Chr$(34)
    Dim x As String
    x = "<customUI xmlns=" & q & "http://schemas.microsoft.com/office/2009/07/customui" & q & ">" & vbLf
    x = x & "  <ribbon><tabs>" & vbLf
    x = x & "    <tab id=" & q & "FrazaroTab" & q & " label=" & q & "Frazaro" & q & ">" & vbLf
    x = x & "      <group id=" & q & "FrazaroID" & q & " label=" & q & "IDE" & q & ">" & vbLf
    x = x & RibbonBtn("VlaSetup", "New Frazaro")
    x = x & RibbonBtn("VlaAddProgram", "New Named Frazaro")
    ' Owner: the first question every new user asks is "what can I
    ' type?" - front-loaded right beside the two creation buttons, in
    ' first-person voice on purpose (curiosity-driven, not a spec doc).
    x = x & RibbonBtn("VlaPhrases", "What can I say?")
    ' GO.6: the phrasebook-loading button, grouped beside its own
    ' "what can I say?" - both phrasebook-facing, distinct from the two
    ' program-facing buttons (Load/Reload Instructions) just below.
    x = x & RibbonBtn("VlaLoadPhrasebook", "Load Phrasebook")
    x = x & RibbonBtn("VlaImport", "Load Instructions")
    x = x & RibbonBtn("VlaReload", "Reload Instructions")
    x = x & "      </group>" & vbLf
    x = x & "      <group id=" & q & "FrazaroCheck" & q & " label=" & q & "Testing" & q & ">" & vbLf
    x = x & RibbonBtn("VlaCheck", "Validate Instructions")
    x = x & RibbonBtn("VlaUndo", "Undo Last Run")
    x = x & "      </group>" & vbLf
    ' IN.6: the interpreter, the default runtime (IN.9) - no VBProject
    ' access, works where Compile cannot (locked-down IT, Mac Excel).
    x = x & "      <group id=" & q & "FrazaroInterpreter" & q & " label=" & q & "Interpreter" & q & ">" & vbLf
    x = x & RibbonBtn("VlaInterpret", "Interpret and Run")
    x = x & RibbonBtn("VlaInterpretTrace", "Interpret and Trace")
    x = x & "      </group>" & vbLf
    ' IN.6: compiling to real VBA - the export/self-contained-workbook
    ' path (IN.9), and the one that needs VBProject trust. Trimmed to
    ' just the two run actions - the file-audit buttons that used to
    ' live here moved to Auditing below.
    x = x & "      <group id=" & q & "FrazaroCompiler" & q & " label=" & q & "Compiler" & q & ">" & vbLf
    x = x & RibbonBtn("VlaRun", "Compile and Run")
    x = x & RibbonBtn("VlaRunTrace", "Compile and Trace")
    x = x & "      </group>" & vbLf
    ' Everything that makes a compiled/loaded artifact visible or
    ' checkable, grouped together regardless of which backend produced
    ' it: the IDE-facing reveal, the file-based translation pair, and
    ' the vocabulary-introspection exports.
    x = x & "      <group id=" & q & "FrazaroHelp" & q & " label=" & q & "Auditing" & q & ">" & vbLf
    x = x & RibbonBtn("VlaShowVba", "Show me the VBA")
    x = x & RibbonBtn("VlaTranslateVla", "Translate File to VLA")
    x = x & RibbonBtn("VlaTranslateVba", "Translate File to VBA")
    x = x & RibbonBtn("VlaExportExpanded", "Export Expanded Phrasebook")
    x = x & RibbonBtn("VlaRuleCoverage", "Export Phrasebook Test Coverage")
    x = x & RibbonBtn("VlaLintVla", "Lint VLA File")
    x = x & "      </group>" & vbLf
    ' Everything outside the write -> test -> run -> audit loop: the
    ' alternate CLI surface, install lifecycle, and diagnostics.
    x = x & "      <group id=" & q & "FrazaroUtilities" & q & " label=" & q & "Utilities" & q & ">" & vbLf
    x = x & RibbonBtn("VlaOpenCli", "Open CLI")
    x = x & RibbonBtn("VlaRegister", "Register for Auto-Load")
    x = x & RibbonBtn("VlaUninstall", "Uninstall Frazaro")
    x = x & RibbonBtn("VlaFeedback", "Copy Diagnostic Report")
    x = x & "      </group>" & vbLf
    x = x & "    </tab>" & vbLf
    x = x & "  </tabs></ribbon>" & vbLf
    x = x & "</customUI>" & vbLf
    VlaRibbonXml = x
End Function

Private Function RibbonBtn(ByVal id As String, ByVal label As String) As String
    Dim q As String
    q = Chr$(34)
    RibbonBtn = "        <button id=" & q & id & q & " label=" & q & label & q & _
                " size=" & q & "large" & q & " onAction=" & q & "VlaRibbonAction" & q & "/>" & vbLf
End Function

' Inject the customUI part into the saved .xlam. Returns a one-line
' note for the build report; never raises (soft-failure doctrine).
Public Function VlaInjectRibbon(ByVal xlamPath As String) As String
    On Error GoTo softly
    Dim tmp As String
    tmp = Environ$("TEMP")
    Dim xmlPath As String, psPath As String
    xmlPath = tmp & "\VlaRibbon_customUI14.xml"
    psPath = tmp & "\VlaRibbonInject.ps1"
    WriteTextFile xmlPath, VlaRibbonXml()
    WriteTextFile psPath, RibbonPs1(xlamPath, xmlPath)

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """"
    Dim code As Long
    code = sh.Run(cmd, 0, True)              ' hidden window, wait for exit
    If code = 0 Then
        VlaInjectRibbon = "Ribbon injected:" & vbCrLf & "Frazaro tab (customUI14)."
    Else
        VlaInjectRibbon = "Ribbon NOT injected (PowerShell exit " & code & ") - run " & psPath & _
                          " by hand, or use the legacy Add-ins menu; the add-in works either way."
    End If
    Exit Function
softly:
    VlaInjectRibbon = "Ribbon NOT injected (" & Err.Description & ") - the legacy Add-ins menu still works."
End Function

' The injection script. PowerShell single-quoted strings take the
' paths (no interpolation, no escaping hazards beyond doubling any
' embedded apostrophe). Steps: add/replace the customUI part, add
' the extensibility relationship to _rels/.rels if absent, and
' defensively ensure [Content_Types].xml declares the xml default
' (standard packages already do; a missing one would fail silently).
Private Function RibbonPs1(ByVal xlamPath As String, ByVal xmlPath As String) As String
    Dim s As String
    s = "$ErrorActionPreference = 'Stop'" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "  Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbCrLf
    s = s & "  $xlam = '" & PsQuote(xlamPath) & "'" & vbCrLf
    s = s & "  $xmlPath = '" & PsQuote(xmlPath) & "'" & vbCrLf
    s = s & "  $zip = [System.IO.Compression.ZipFile]::Open($xlam, 'Update')" & vbCrLf
    s = s & "  try {" & vbCrLf
    s = s & "    $old = $zip.GetEntry('customUI/customUI14.xml')" & vbCrLf
    s = s & "    if ($old) { $old.Delete() }" & vbCrLf
    s = s & "    $entry = $zip.CreateEntry('customUI/customUI14.xml')" & vbCrLf
    s = s & "    $w = New-Object System.IO.StreamWriter($entry.Open())" & vbCrLf
    s = s & "    $w.Write([System.IO.File]::ReadAllText($xmlPath))" & vbCrLf
    s = s & "    $w.Close()" & vbCrLf
    s = s & "    $rels = $zip.GetEntry('_rels/.rels')" & vbCrLf
    s = s & "    $r = New-Object System.IO.StreamReader($rels.Open())" & vbCrLf
    s = s & "    $txt = $r.ReadToEnd(); $r.Close()" & vbCrLf
    s = s & "    if ($txt -notmatch 'customUI14') {" & vbCrLf
    s = s & "      $ins = '<Relationship Id=""rIdVlaUI"" Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Target=""customUI/customUI14.xml""/>'" & vbCrLf
    s = s & "      $txt = $txt.Replace('</Relationships>', $ins + '</Relationships>')" & vbCrLf
    s = s & "      $rels.Delete()" & vbCrLf
    s = s & "      $rels2 = $zip.CreateEntry('_rels/.rels')" & vbCrLf
    s = s & "      $w2 = New-Object System.IO.StreamWriter($rels2.Open())" & vbCrLf
    s = s & "      $w2.Write($txt); $w2.Close()" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    $ct = $zip.GetEntry('[Content_Types].xml')" & vbCrLf
    s = s & "    $r3 = New-Object System.IO.StreamReader($ct.Open())" & vbCrLf
    s = s & "    $txt3 = $r3.ReadToEnd(); $r3.Close()" & vbCrLf
    s = s & "    if ($txt3 -notmatch 'Extension=""xml""') {" & vbCrLf
    s = s & "      $ins3 = '<Default Extension=""xml"" ContentType=""application/xml""/>'" & vbCrLf
    s = s & "      $txt3 = $txt3.Replace('</Types>', $ins3 + '</Types>')" & vbCrLf
    s = s & "      $ct.Delete()" & vbCrLf
    s = s & "      $ct2 = $zip.CreateEntry('[Content_Types].xml')" & vbCrLf
    s = s & "      $w3 = New-Object System.IO.StreamWriter($ct2.Open())" & vbCrLf
    s = s & "      $w3.Write($txt3); $w3.Close()" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "  } finally { $zip.Dispose() }" & vbCrLf
    s = s & "  exit 0" & vbCrLf
    s = s & "} catch { Write-Error $_; exit 1 }" & vbCrLf
    RibbonPs1 = s
End Function

Private Function PsQuote(ByVal p As String) As String
    PsQuote = Replace(p, "'", "''")          ' PS single-quote doubling
End Function

' DI2.1: the "write a plain-text asset into a very-hidden sheet, one
' line per row" trick, extracted once V5.3's own VLA_Runtime case
' gained two siblings (prelude.vla, english.vla) rather than copied a
' third time. Reader side: VLA.VlaEmbeddedText.
Private Sub EmbedTextAsSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal text As String)
    Dim sh As Worksheet
    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = sheetName
    sh.Columns(1).NumberFormat = "@"
    Dim ls() As String
    ls = Split(Replace(text, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(ls) To UBound(ls)
        sh.Cells(i - LBound(ls) + 1, 1).Value = ls(i)
    Next
    sh.Visible = xlSheetVeryHidden
End Sub

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    If Len(Dir$(path)) > 0 Then Kill path
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub
