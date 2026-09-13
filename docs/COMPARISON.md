# Frazaro beside Copilot in Excel, Office Scripts and Python in Excel

*SIG.2 · Frazaro 0.5.6 · 2026-09-10. Other products' cells cite Microsoft's
documentation; Frazaro's point at this repository.*

| | **Frazaro** | **Copilot in Excel** | **Office Scripts** | **Python in Excel** |
|---|---|---|---|---|
| **You write** | English sentences from a published grammar | A request in your own words | TypeScript [O4], or recorded actions [O1] | Python in a cell [P1] |
| **Same input, same result?** | Yes: same sentences and phrasebooks, same program; ambiguous grammar is refused at load (SD-16) | Not promised: "can sometimes make mistakes" [C1]; `COPILOT()` results "may change over time, even with the same arguments" [C3] | Yes, it is a program (a `fetch` answer can vary) | Yes, it is a program |
| **Where, and network** | In your desktop Excel; no network call of its own ([IT_REVIEW.md](IT_REVIEW.md) §3) | Microsoft 365 services, network required; `COPILOT()` models "are hosted on Azure" [C2][C3] | In Excel; scripts live in OneDrive or SharePoint, so network is required [O3]; may call web services [O5] | Its "calculations run in the Microsoft Cloud"; "you need internet access" [P1] |
| **A reviewer reads** | The sentences; *Show me the VBA* shows the VBA Compile runs, in steps tied to sentence lines | What it produced | The TypeScript; admins get an audit log [O3] | The Python |
| **Platforms; needs** | Windows desktop Excel; a macro policy that allows it | Windows, Mac, web, iPad, iPhone, Android; a Copilot-eligible subscription [C1] | Web, Windows (2210+), Mac; a Microsoft 365 business licence [O2] | Windows, web, Mac; Microsoft 365 [P1] |

**Use Copilot in Excel** for an answer or a draft now, from a question in your
own words, that a person will check. It wins first contact outright: there
is no syntax to learn, while Frazaro may refuse your first sentence.
Microsoft's own advice is to "avoid using Copilot for decisions in
sensitive areas such as finance, legal, or medical topics" [C1].

**Use Office Scripts** when the organization runs Microsoft 365 with OneDrive
and has someone who writes TypeScript. Scripts are as deterministic as
Frazaro and readable by anyone who reads TypeScript. They reach further
(web, Mac, Power Automate [O1][O2]), and their activity shows up in an admin
audit log, which Frazaro does not have.

**Use Python in Excel** for analysis (statistics, dataframes, plots) where
running the referenced data through Microsoft's cloud is acceptable
[P1]; the Python itself has no network access [P2]. Its analytical depth is
far beyond Frazaro's.

**Use Frazaro** for a procedure that must run the same way every time, on the
machine, with no cloud service in the path, in words the accountable person
can read without programming. Its limits: a beta from one maintainer;
Windows desktop only; a grammar not yet complete for any profession; a VBA
add-in your macro policy must allow ([IT_REVIEW.md](IT_REVIEW.md) §1); no
AI drafting or per-run log yet.

**Check it yourself:** `scripts/instructions_golden.vba` (the corpus's
compiled VBA: 613 numbered steps, 566 `src:` tags to sentence lines); the
same goldens run against both backends; `tools/check_no_network.ps1`.

*Sources (Microsoft Learn and Support), accessed 2026-09-10:*
[C1] [Copilot in Excel FAQ](https://support.microsoft.com/en-us/excel/copilot/frequently-asked-questions-about-copilot-in-excel) ·
[C2] [Copilot network requirements](https://learn.microsoft.com/en-us/microsoft-365/copilot/microsoft-365-copilot-requirements) ·
[C3] [COPILOT function](https://support.microsoft.com/en-us/excel/functions/copilot-function) ·
Office Scripts: [O1] [overview](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel),
[O2] [requirements](https://learn.microsoft.com/en-us/office/dev/scripts/testing/platform-limits),
[O3] [storage](https://learn.microsoft.com/en-us/office/dev/scripts/overview/script-storage),
[O4] [fundamentals](https://learn.microsoft.com/en-us/office/dev/scripts/develop/scripting-fundamentals),
[O5] [external calls](https://learn.microsoft.com/en-us/office/dev/scripts/develop/external-calls) ·
Python in Excel: [P1] [introduction](https://support.microsoft.com/en-us/office/introduction-to-python-in-excel-55643c2e-ff56-4168-b1ce-9428c8308545),
[P2] [data security](https://support.microsoft.com/en-us/office/data-security-and-python-in-excel-33cc88a4-4a87-485e-9ff9-f35958278327).
Product names are trademarks of Microsoft or the Python Software
Foundation, named only to identify the products.
