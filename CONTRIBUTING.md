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

## Security

Do not open a public issue for a vulnerability. See `docs/SECURITY.md`.

## The name

"Frazaro" is a trademark; see `TRADEMARK.md`. A fork is welcome under the
licences above and needs its own name.
