# Maintainers

*Who answers for which paths, in the shape Linux's `MAINTAINERS` file gives
it: a path, a seam, an oracle. One maintainer today. The map is written for
the day there are more, so that work in different seams never meets in the
same file, and so that a change to a contract file is a meeting rather than
a commit. The seams themselves are `CONTRIBUTING.md`'s "The kernel and its
seams" and `core/src/kernel.rs`.*

**Maintainer:** Spreadsheet Company (the owner), who merges to `main`. Every
live Excel pass runs through the owner's hands (`docs/TESTING.md`); every
other oracle runs without Excel, and a change is ready for integration when
it is green on all of its own.

## Paths, by seam

| Paths | Seam, or line of the roadmap | Oracle |
|---|---|---|
| `core/src/` except `english/`; `cli/`; `scripts/build/`; `scripts/reflect/`; `conformance/` | the kernel: forms, emitters, sheet model, reader, writer, ABI (the KERNEL line) | the compile, build, reflect, diff and audit goldens; `cargo test` |
| `web/`; `tools/build_web.ps1`; `tools/check_web_offline.ps1` | the page and the doors (the KERNEL line's view track; `PORT.11`) | `check_web_offline.ps1`; the view goldens when they exist |
| `scripts/polyglotta/`; `scripts/prelude.vla`; `src/VLA_SentenceEngine.bas`; `src/VLA_English.bas`; `core/src/english/`; `docs/GRAMMAR_SINCE.md` | the sentences seam, both sides: the reference leads (`SD-18`) and the port follows | the translate golden; `frazaro prove`; the token and refusal goldens |
| `src/VLA_Datalog.bas`, `src/VLA_Prolog.bas`, `src/VLA_Sql.bas`, `src/VLA_Optimize*.bas`, `src/VLA_Relation.bas`; `scripts/proofs/`; `tools/*_lp.ps1` | the engines seam (`PORT.9`; the Singularity line's engines) | `scripts/proofs/datalog.vla` with clingo beside it; `TestDSLs` |
| `tools/check_*.ps1`; `tools/run_checks.ps1`; `.github/`; `installer/`; `tools/release.ps1`; `docs/DEPLOY.md`, `docs/THREAT_MODEL.md`, `docs/IT_REVIEW.md` | assurance and release | each check's own `-Control`; the ratchets job |
| `src/` otherwise: `VLA_IDE`, `VLA_Runtime`, `VLA_Interpreter`, `VLA_Events`, `VLA_Loader`, `VLA_Provenance`, `VLA_Build`, the tests | the reference host, the Excel add-in (`SD-18`) | `VlaSelfTests`, `VerifyReports`, the live pass |
| `docs/BETA_ROADMAP.md`, `docs/BETA_REARVIEW.md`, `docs/ID_REGISTRY.md`, `docs/RELEASES.md` | the registers | `check_id_registry.ps1`; a change edits only its own line's section |

## Contract files

A change to one of these needs both sides of its seam present, since more
than one path reads it:

- `core/src/abi.rs`: the record every door reads; the ABI number stays 1.
- `core/src/kernel.rs`: the seams as data, the `Engine` and `Projection` traits, the host profiles.
- `core/data/headtable.vla`, exported from `src/VLA_HeadTable.bas`: the forms, which is to say the syscall table.
- `core/data/messages.vla`, exported from `src/VLA_Messages.bas`: the refusals; a team adds inside its own id family's block.
- `scripts/polyglotta/english.vla`: the base corpus; a sentence another seam needs is a request to the sentences seam (`SD-7`).
- `conformance/README.md`: the treaty; an oracle joins by a dated amendment.
- The goldens under `scripts/`: written by the reference or blessed by the owner, never edited by hand.
