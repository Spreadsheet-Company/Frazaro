# The docshelf

*An index, not a document. The files here are written for three readers
who rarely need each other's: someone **using** Frazaro, someone
**building** it, and someone **evaluating** whether to bet on it. Nothing
moved to make this list — the folder is still flat — and every entry below
is one sentence and a link. Where a summary and its document disagree, the
document is right.*

---

## If you are using Frazaro

- **[DEPLOY.md](DEPLOY.md)** — Which of the two downloads to take, how to
  install it, and how to remove it again.
- **[SUPPORT.md](SUPPORT.md)** — Where to report something broken,
  confusing, or missing, and what response to expect.
- **[SECURITY.md](SECURITY.md)** — Where to report a *vulnerability*
  instead, and the response commitment that comes with it.
- **[RELEASES.md](RELEASES.md)** — What changed in each release, newest
  first, written before the release rather than after it.
- **[GRAMMAR_SINCE.md](GRAMMAR_SINCE.md)** — Which release each sentence
  you can write first worked in.

## If you are building Frazaro

- **[BETA_ROADMAP2.md](BETA_ROADMAP2.md)** — The current plan: every item,
  open and closed, cut to one paragraph and filed under the department
  accountable for it.
- **[BETA_ROADMAP1.md](BETA_ROADMAP1.md)** — Its predecessor, kept for the
  full reasoning behind every standing decision and every closed item.
- **[TESTING.md](TESTING.md)** — The six verification passes a change to
  the grammar, a phrasebook, or the runtime has to survive, in order.
- **[REBUILD.md](REBUILD.md)** — What the modules in `src/` should be
  shaped like if they were written again from nothing.
- **[ID_REGISTRY.md](ID_REGISTRY.md)** — The register that makes "a retired
  item ID is never re-minted" mechanical rather than remembered.
- **[INTRINSICS.md](INTRINSICS.md)** — The VBA behaviours a non-VBA port of
  the translator would have to reproduce exactly.
- **[THREAT_MODEL.md](THREAT_MODEL.md)** — What a program may reach, what a
  phrasebook may reach, and who is trusted at each layer.
- **[CUTS.md](CUTS.md)** — Where the code splits into modules and where the
  licence splits into territories.
- **[LESSONS.md](LESSONS.md)** — Why each of the project's disciplines
  exists, told as the bug that created it.
- **[TRENCHES.md](TRENCHES.md)** — The same history one layer below the
  language: the ground-up debugging campaigns the runtime cost.

## If you are evaluating Frazaro

*Reviews the project commissioned against its own blind spots, plus the two
that argue the commercial case. An IT reviewer deciding whether to allow the
add-in on a managed machine wants [THREAT_MODEL.md](THREAT_MODEL.md) and
[SECURITY.md](SECURITY.md) above instead.*

- **[CONSULTANT.md](CONSULTANT.md)** — An outside project-management
  assessment of the beta, written nine days before the `0.5.0` target.
- **[AUDIT.md](AUDIT.md)** — The master roadmap read against the project's
  own philosophy essay, with the antithesis argued in borrowed uniforms.
- **[PREMORTEM.md](PREMORTEM.md)** — The coroner's report, filed in advance:
  what killed the project, written as though it already had.
- **[ADVOCATUS.md](ADVOCATUS.md)** — The devil's advocate, retained against
  the project's own premises.
- **[CONTINUITY.md](CONTINUITY.md)** — The successor's audit: what survives
  the owner.
- **[VIABILITY.md](VIABILITY.md)** — The accountant's report: the win
  condition, priced.
- **[SUBSTRATE.md](SUBSTRATE.md)** — The platform historian's survey of the
  ground this is built on, and how that ground has moved before.
- **[VENTURE.md](VENTURE.md)** — The commercial case, written for outside
  capital, and carrying its own edit-before-distributing warning at the top.
- **[MARKETING.md](MARKETING.md)** — The operating manual for finding the
  people who will break the beta.
- **[CONTEMPLATIONS.md](CONTEMPLATIONS.md)** — Append-only thought
  experiments: the questions that are not yet items and the arguments that
  are not yet decisions.

## Also in this folder

- **[archive/](archive/)** — The six alpha roadmaps and the original
  session-bootstrap brief, superseded and kept for the trail.
- **[metameta/](metameta/)** — The philosophy essays the roadmap's ordering
  doctrine is argued in, and a Kierkegaard pastiche written against them.
- **[sententiae.txt](sententiae.txt)** — Forty one-line formulae of the
  project's own argument, in successive drafts.

## Not in this folder

At the repository root, because a visitor arriving from GitHub sees them
before they see `docs/`:

- **[README.md](../README.md)** — What Frazaro is, what works today, and the
  open security items a downloader should hear from that page rather than
  from the roadmap.
- **[CONTRIBUTING.md](../CONTRIBUTING.md)** — The DCO sign-off, and what to
  expect from a patch.
- **[TRADEMARK.md](../TRADEMARK.md)**, **[OUTPUT-EXCEPTION.md](../OUTPUT-EXCEPTION.md)**,
  **[PHRASEBOOK-TERMS.md](../PHRASEBOOK-TERMS.md)** — The name; why your
  workbook does not inherit the licence of the code Frazaro writes into it;
  and the four sentences an organization's counsel will ask about
  phrasebooks.
- **[LICENSE](../LICENSE)**, **[LICENSES/](../LICENSES/)**,
  **[NOTICE](../NOTICE)**, **[REUSE.toml](../REUSE.toml)** — Four licence
  territories in one repository, with a per-file map and no directory moves.
