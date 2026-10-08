# Contributing to Frazaro

*Short on purpose. The contribution ladder, review standards, and the
escalation path for a disputed surface are open roadmap items (`GO.2`,
`GO.5`); this file is what has to exist before the first outside patch, not
after.*

## Sign your commits (DCO)

Contributions are accepted under the [Developer Certificate of
Origin](https://developercertificate.org/) 1.1. Sign each commit:

```text
git commit -s
```

That adds a `Signed-off-by:` line certifying you wrote the change or have
the right to submit it. There is no Contributor License Agreement, and there
will not be one: every file's licence is permissive enough that none is
needed.

## Inbound = outbound

By contributing, you agree your contribution is licensed under the licence
of the file it touches, as declared in `REUSE.toml`:

| Path | Licence |
|---|---|
| `src/**`, `scripts/prelude.vla`, `tools/**`, `installer/**`, the corpus fixtures | Apache-2.0 |
| `src/VLA_Runtime.bas` | 0BSD |
| `scripts/polyglotta/**` (the phrasebooks) | MPL-2.0 |
| `docs/**`, `README.md` | CC-BY-4.0 |

and that the `OUTPUT-EXCEPTION.md` additional permission applies to it.

## What a change needs

- **A phrasebook rule** needs its `test:` proof in the same file; the build
  refuses a phrasebook whose proofs fail.
- **An engine change** needs the goldens to regenerate with an empty diff,
  or a commit message that says what changed and why. `docs/TESTING.md`
  says how to run the self-tests; the owner runs the live Excel pass.
- **A refusal** goes through the message catalogue with a stable id, never a
  raw string (`SD-2`).
- **A roadmap ID** is never reused (`SD-9`); `tools/check_id_registry.ps1`
  checks.

## The kernel and its seams

The core (`core/`, the crate `frazaro-core`) is a kernel. It holds forms and
their expansion, the emitters, the sheet model, the relation set and the
ABI, and nothing a person reads or says: English, message text, defaults,
chrome, formats beyond a trait and doors live outside it, as data the kernel
reads or as implementations of a seam. The rule for a change is that the
kernel grows a seam, never a feature. Five seams are the only entrances,
each with the oracle a change through it must pass. `core/src/kernel.rs`
lists them as data, and `tools/check_kernel_boundary.ps1` holds the rule
over the sources.

| Seam | A contribution is | It must pass |
|---|---|---|
| sentences | a phrasebook `.vla` with its `test:` proofs | `frazaro prove`; the translate and refusal goldens |
| paragraphs | a library of sentences with named slots (`G-USE`) | the build golden; the stamp hashing every file used |
| engines | an implementation of `kernel::Engine` under a head-table symbol (`PORT.9`) | a proof file of its kind, clingo beside it where it applies |
| formats and hosts | a `reflect::Source` to read, a writer beside it, a `kernel::HostProfile` per door | the reflect, diff and audit goldens; the build golden |
| projections | an implementation of `kernel::Projection`; the grid, `view::Grid`, is the first (`KERNEL.4`) | a view golden under `scripts/view/` (`check_view_golden.ps1`); the free one: the view of a built model's sheet, whole, is `reflect` of its file |

A feature that fits no seam is a seam to design first, on the roadmap, not
a feature to merge. `MAINTAINERS.md` maps the paths to the seams and names
the contract files a change to which needs both sides present.

**The distro** is the unit the build tools take (`KERNEL.2`): a folder
under `distros/` whose `distro.vla` names, by reference, the prelude, the
base phrasebook, an overlay, the dialects, the libraries, the examples and
the edition's chrome and palette. The page builder bakes it into one page,
`frazaro prove <folder>` proves it whole, and the add-in builder reads the
edition's chain from it; `distros/english/README.md` has the shape and how
to make your own. An edition is a folder added, never a table edited.

## Security

Do not open a public issue for a vulnerability. See `docs/SECURITY.md`.

## The name

"Frazaro" is a trademark; see `TRADEMARK.md`. A fork is welcome under the
licences above and needs its own name.
