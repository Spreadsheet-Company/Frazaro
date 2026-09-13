# Frazaro — security and architecture summary for IT reviewers

*SIG.1 · **Frazaro 0.5.6** · 2026-09-10. For later builds, see
[README.md's open items](../README.md#known-open-security-items).*

## 1. What it is, and the gate you will actually decide on

An Excel add-in (`.xlam`) written in VBA. People write programs as English
sentences in cells. **Interpret** (the default) runs them in place, writing
no code; **Compile** (opt-in) writes them out as VBA modules in the
workbook. Open source; generated code belongs to the user
([OUTPUT-EXCEPTION.md](../OUTPUT-EXCEPTION.md)). A beta from one
maintainer, without warranty.

**It is a macro add-in, so your macro policy decides whether it runs.**

- A downloaded `.xlam` carries Windows' Mark of the Web, and Excel blocks
  it. **For Excel add-ins a digital signature does not lift that block**
  [1]. Remove the mark (Properties → Unblock) or use a Trusted Location
  [1]. Files placed by the installer carry no mark (`DEPLOY.md`, tested).
- The VBA project is signed with a **self-signed** certificate
  (`CN=Frazaro VBA Signing`). `tools/release.ps1` will not publish without
  the maintainer's attestation that it is signed; VBA has no API to verify
  it. Under Microsoft's baseline ("disable all except digitally signed
  macros", "require a trusted publisher") [1], you would deploy that
  certificate as a Trusted Publisher. It is not yet published separately.
  `FrazaroSetup.exe` is Authenticode-signed, also self-signed
  (`CN=Frazaro Dev Signing`), so SmartScreen does not trust it on its own.
- **Interpret needs no VBA-project access.** Compile needs *Trust access to
  the VBA project object model*, which Office denies by default [2].
  Frazaro never changes that setting, and Compile refuses in words without
  it (`src/VLA_IDE.bas`). Leaving it off costs only Compile.

## 2. What it installs and writes

| What | Where | Removed on uninstall? |
|---|---|---|
| The add-in | Standalone: one `.xlam`, anywhere. Installer: `%AppData%\Frazaro\` (`.xlam` and two `.vla` grammar files), per-user, no admin rights, uninstall entry in HKCU [3]; no services, tasks, drivers, shortcuts or `[Run]` (`installer/Frazaro.iss`) | Yes, bar a `scripts` folder a user added (the dialog says so) |
| Excel registration | An `OPEN`/`OPENn` value under `HKCU\…\Office\<ver>\Excel\Options`, as Excel's Add-ins dialog writes | Yes (only its own) |
| Consents | `HKCU\Software\VB and VBA Program Settings\Frazaro` [4]: raw-VBA consents, phrasebook-path approvals | **No**; delete by hand |
| In the workbook | Program and output sheets; very-hidden Undo snapshots (`VLAu_*`) and an unrecognized-sentence log (`VLA_Log`); document properties. Compile only: `Frazaro_EN_Runtime`, a module per program, two `VLAt_*` names | No, by design: compiled workbooks run without Frazaro |
| Temp file | `%TEMP%\VlaSelfDelete.ps1`, from the standalone uninstall | No |

## 3. What it can reach

**Frazaro's own code makes no network connection.** It has no HTTP client,
socket or download API, and its only direct Windows API calls are four
`user32` window functions. It has no update check and no telemetry, by
standing decision (SD-13). `tools/check_no_network.ps1` pins this on every
push (CI) and before every release, and fails if a route outward appears
that this table does not name. Except through a `raw` rule or Compile's
SEC.12 (both below), programs reach outside the workbook only this way:

| Route | What happens | Control today |
|---|---|---|
| Files | Open, save-as, save-copy at a path the program names; print; PDF export; sheet passwords | Interpret: refused if internet-marked (SEC.8), else unprompted (SEC.7). Compile: not gated |
| Mail | An Outlook draft, optionally with a local file attached, shown for the user to send; never sent | As Files |
| "Refresh everything." | Refreshes the workbook's *existing* external data ranges and PivotTables [5], wherever they point; Frazaro adds no connection | **None.** Not SEC.8-gated; not yet a filed item |
| Formulas | Any formula, including `WEBSERVICE` ("returns data from a web service on the Internet or Intranet" [6]) or DDE | None (SEC.15, open) |
| Other programs | Word (macros force-disabled) to import a `.doc`/`.docx` (SEC.13). Uninstall runs the uninstaller, or for a standalone copy a hidden `powershell -ExecutionPolicy Bypass` script deleting the `.xlam` (endpoint security may flag it) | User-initiated only |

**Inside Excel**, Interpret reaches only a fixed, named list of object-model
members and refuses anything else in words (SEC.1;
[THREAT_MODEL.md](THREAT_MODEL.md) §1.1). A **phrasebook** (grammar file)
can contain a `raw` rule: arbitrary VBA with full privilege. Such a
phrasebook loads only after a consent dialog naming it (SEC.2) or a
remembered consent, and SEC.10 below is how a sent workbook can carry
one. Phrasebooks found beside, or remembered by, a workbook need this
computer's approval (SEC.9). The built-in phrasebooks contain no
`raw` rule and are audited at build time. Older text's "zero-trust
runtime" means *no trust prompt*, not least privilege (THREAT_MODEL §0).

## 4. Security items as of 0.5.6

Findings from reading the code, not exploits anyone has run; no external
review yet (SEC.6, open). Until the open items close, load phrasebooks
only from people you would accept a macro-enabled workbook from.

- **Open:** SEC.3 (generated code does not record which phrasebook
  introduced it), SEC.7 (no permission prompt on external-effect verbs),
  SEC.10 (workbook-scope consent lives in the workbook, so a sent one can
  arrive pre-consented), SEC.15 (formulas are not screened).
- **Accepted, reasons on record:** SEC.12, SEC.17 (Compile only, behind
  the §1 setting); SEC.14 (a program can hang Excel, which Ctrl+Break
  stops, or nest deep enough to crash it); SEC.16 (replacing the grammar
  files beside the add-in needs code already running as the user).
- **Closed:** SEC.1, 2, 4, 8, 9, 11, 13. File, line and fix for each:
  [BETA_ROADMAP1.md](BETA_ROADMAP1.md).

## 5. Updates, removal, reporting

**No automatic updates, by design.** A fix ships as a new release with a
security block in its notes; nothing reaches an installed copy on its own.
To hear about fixes, watch the
[GitHub repository](https://github.com/Spreadsheet-Company/Frazaro) for
releases [7]; the running version shows on the *What can I say?* sheet.

**Removal:** the ribbon's **Uninstall Frazaro** button deregisters the
add-in; a standalone copy then deletes itself, an installed one hands off
to the Windows uninstaller (also in Settings → Apps). The self-delete and
the uninstaller were each tested live (`DEPLOY.md`). What stays is in §2.
**Report** security issues to `english@spreadsheet.company`, acknowledged
within 3 business days ([SECURITY.md](SECURITY.md)).

*Sources (Microsoft unless named), accessed 2026-09-10:*
[1] [Macros from the internet are blocked by default in Office](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked) ·
[2] [Enable or disable macros in Microsoft 365 files](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6) ·
[3] [Inno Setup: Non Administrative Install Mode](https://jrsoftware.org/ishelp/topic_admininstallmode.htm) ·
[4] [SaveSetting statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/savesetting-statement) ·
[5] [Workbook.RefreshAll method](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.refreshall) ·
[6] [WEBSERVICE function](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4) ·
[7] [GitHub: About notifications](https://docs.github.com/en/subscriptions-and-notifications/concepts/about-notifications)
