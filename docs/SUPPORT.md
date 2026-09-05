# SUPPORT.md — SIG.3

*The ordinary-support half of `SEC.5`'s `SECURITY.md`, which covers only
vulnerability disclosure. This document is the other half: what to
expect when something is broken, confusing, or missing — not exploitable.
Where a user reports a bug, what response they can expect, and where to
send the diagnostics the product already knows how to gather but has
never had anywhere to point.*

**Not the right document if you've found a security vulnerability** — see
`SECURITY.md` instead.

## 1. Getting help

**Contact: `english@spreadsheet.company`.**

Before emailing, use **Copy Feedback** (the ribbon/panel button) — it
already gathers what a support request actually needs: your Frazaro and
engine versions, every sentence loaded so far, and every instruction the
corpus didn't recognize (the log behind the red cells in column C). It
puts that report on the clipboard; this document is the part that was
missing — where to paste it. Include:

- What you were trying to do, in your own words, alongside the Copy
  Feedback report.
- Whether the sentence failed with a red cell (a refusal) or produced
  the wrong result silently (the more serious of the two — a refusal is
  this project's own intended failure mode; a silent wrong answer is a
  real bug regardless of how it presents).
- Which backend you ran it through, if you know — Interpret or Compile
  (`IN.9`'s own split; the two don't share every code path, so which one
  you hit matters for reproducing it).

## 2. Response commitment

We will acknowledge a support request within **5 business days**. As
with `SECURITY.md`, we don't commit to a fixed resolution deadline —
Frazaro is under active, ongoing development, and a promised fix date we
can't reliably hit is worse than an honest "we don't know yet." What we
do commit to: you'll hear back from a person, your report will be read
against the actual corpus/engine behavior it names, and you'll be told
either what's changing or why it isn't.

## 3. What this covers

Bug reports, sentences that refuse when they shouldn't (or the reverse —
accept something that should have been refused), unclear refusal
wording, and general "how do I say X" questions the in-product help
doesn't yet answer. Does **not** cover a security vulnerability — a
`raw`-consent bypass, unexpected `Application`-surface reachability, or
anything else `THREAT_MODEL.md` names as a real gap belongs in
`SECURITY.md` instead, at the contact and under the response commitment
stated there, not here.

## 4. What this does not promise

No SLA-backed uptime or availability guarantee — Frazaro is an add-in
running on the user's own machine, not a hosted service, so "uptime" is
not a meaningful commitment to make in the first place. No guarantee that
a requested feature or phrasing gets built — `LE.1`'s own sentence
palette and the corpus itself are how "what can I say" gets answered
today; a support request that's really a feature request is welcome, but
this document is about response to a report, not a roadmap commitment.
