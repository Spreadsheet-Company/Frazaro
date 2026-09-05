# SECURITY.md — SEC.5

*The minimum credible artifact before any external pilot (`PI.*`) or
public download: a place a finder is told to report to, and a stated
response commitment. This document does not resolve the disclosure
mechanism itself — see the honest tension in §3 — it only makes sure a
report has somewhere to land.*

## 1. Reporting a vulnerability

**Contact: `english@spreadsheet.company`.**

See `THREAT_MODEL.md`
(`SEC.0`) for what's already known to be worth reporting against: the
`Application`-reachability finding (§1.2) and the `raw`-no-consent gap
(§1.4) are the two most likely categories a finder would surface first.

Please include:
- The Frazaro version (`VLA_RELEASE_VERSION`, `SD-14`'s own format) or
  commit hash you're running.
- A minimal `.vla` phrasebook rule or English sentence that reproduces
  the issue, if the report is about something reachable through the
  language rather than the build tooling.
- Whether the issue was found through the interpreter, the emitter
  (module-injection export path), or both — `IN.9`'s own split matters
  here, since the two backends don't share every code path.

## 2. Response commitment

We will acknowledge a report within **3 business days**. We do not
commit to a fixed remediation timeline — severity and complexity vary,
and a promised deadline this project can't reliably meet is worse than
an honest one. Instead: the reporter is kept informed of our own
assessment and progress until the issue is resolved, and a critical,
actively-exploitable issue is prioritized above all other work the
moment it's confirmed.

## 3. The honest tension, named rather than solved here

`SD-13` — no outbound network call, ever, without a dedicated
re-litigation of that standing decision — means a disclosed
vulnerability has **no push-update path** to people already running an
affected build. There is no auto-update mechanism, and `SD-13` means
there deliberately isn't one. This document does not resolve that; it
only makes sure a report has somewhere to land once found. Whoever picks
up distribution/patching (`DI.*`) is the real owner of *how* a fix
reaches an already-installed copy — this file's own job stops at "a
report gets read by a person," not "a fix gets pushed."

## 4. Scope

Covers `Frazaro`/`VLA` — the add-in, the phrasebook loader, the two
backends (interpreter and emitter), and the shipped base corpus. Does
**not** cover third-party phrasebooks a user or organization has loaded
themselves — see `THREAT_MODEL.md` §2 for why that boundary matters
(no trust-tier distinction exists yet between the base corpus and a
community-sourced phrasebook, which is itself a finding worth reporting
against, not an exclusion meant to dodge it).
