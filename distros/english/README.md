# The english distro

A distro is a folder: the unit the build tools take, and the whole of what
an edition of Frazaro ships beyond the engine itself. This one is the
edition the three doors ship today, named by reference: the files stay where
they are, and `distro.vla` beside this README says which they are.

## The manifest

`distro.vla` is one `(distro "name" ...)` form, one directive a line, every
path relative to this folder and written with forward slashes:

| Directive | What it names |
|---|---|
| `(title "...")`, `(tagline "...")` | the page's title and the heading's tagline |
| `(opens-with "...")` | the sentence a new program's first row offers |
| `(palette (lavender "#rrggbb") (whisper "#rrggbb") (deep "#rrggbb"))` | the three shades of Frazaro Lavender the page paints with |
| `(prelude "...")` | the standard library, the one file every compile needs |
| `(phrasebook "name" "...")` | the base phrasebook, loaded first; one per distro |
| `(overlay "name" "...")` | what an edition loads after the base; at most one, none here |
| `(dialect "name" "...")` | an alternative the language picker offers, each loaded over the base |
| `(library "name" "...")` | a macro library a program includes; listed, never proved |
| `(examples "...")` | the folder of sample procedures and their data |
| `(addin "Frazaro_Name.xlam")` | the file name the add-in builder writes |
| `(readme "...")` | this file |

The name is the folder's own: lowercase letters, digits and hyphens.

## The three doors take it

- **The page.** `powershell -File tools\build_web.ps1` reads this folder
  unless `-Distro <folder>` names another, and fills `web/index.template.html`
  from it: the prelude and the base phrasebook as text, each dialect into its
  `{{BOOK:name}}` place (the template's places and the distro's dialects must
  agree exactly), the title, the tagline, the opening sentence and the three
  shades. `tools\check_web_offline.ps1` holds the template's places to this
  manifest's dialects.
- **The command-line door.** `frazaro prove distros\english` proves the
  distro whole: the base alone, as the treaty's oracle 3 proves a file, then
  each dialect over the base, as the page and the add-in load it, one line a
  book and the total last. The prelude and the base phrasebook the door
  carries inside itself (`cli/data/`) are copies of the two this manifest
  names, held byte for byte by `tools\check_crate_data.ps1`.
- **The add-in.** `VlaBuildAddin "English"` reads `distros\english\distro.vla`
  for the phrasebook chain it audits and embeds (the base, then the overlay
  where there is one), the add-in's file name, and the name of the file a
  person may put beside their own workbook to override the embedded
  phrasebook (the last file of the chain).

`tools\check_distro.ps1` holds every manifest under `distros/` to the shape
above, every path to a file that exists, and the three doors' readings to
one another.

## The terms, per kind of file

A distro's files keep the terms of where they live, resolved file by file by
`REUSE.toml` and `tools\check_spdx.ps1`:

- the prelude, this manifest and this README are the engine's, **Apache-2.0**;
- the phrasebooks, the base and every dialect, are **MPL-2.0**, per file, so
  a phrasebook you write is yours (`PHRASEBOOK-TERMS.md`);
- a library carries whatever its author chose; `alien.vla` here is Apache-2.0
  with the corpus fixtures;
- the examples are **0BSD**, made to be copied into your own procedures.

## Making your own

Copy this folder beside it, give the copy its own name, and point its paths
at your files: your organization's phrasebook as the base or as the overlay
over `english.vla`, your examples, your palette. The Org Phrasebook Template
(`Frazaro-Org-Phrasebook-Template`, the repository an organization clones)
is the starting point for the phrasebook itself. A distro is cloned, never
fetched: nothing in Frazaro reads the network (`SD-13`), and a workbook's
build stamp (`Frazaro.Build`) records the hash of every file the build used,
which is the lockfile.

Not in a distro yet: the lines a new program opens with beyond the one
sentence (`G-USE`, the library seam), a per-edition message catalogue and
ribbon (`EDITION-MESSAGES`, `EDITION-CHROME`), and `frazaro describe`, which
prints what a distro contains with each file's hash (`KERNEL.18`).
