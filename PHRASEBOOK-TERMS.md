# Phrasebook terms

*The four sentences an organization's counsel will ask for, and the
reasoning behind them.*

`english.vla` and the edition phrasebooks in `scripts/polyglotta/` are
licensed under the **Mozilla Public License 2.0**, which applies **per
file**. A phrasebook you write in your own file, including one that uses
`override:` to replace a base rule or that calls macros the base phrasebook
defines, is a *Larger Work* under MPL §3.3: it is yours, under any terms you
choose, including none. If you edit `english.vla` itself and **distribute**
the edited file, MPL requires you to make that file's source available under
MPL; using an edited copy inside your own organization creates no obligation
at all. Generated code is covered by `OUTPUT-EXCEPTION.md`, not by this file.

## Why file-scoped copyleft, and why it matches how Frazaro already works

Frazaro loads phrasebooks in layers. The base corpus ships embedded in the
add-in; an organization's own phrasebook sits beside its workbook and is
layered on top; a rule in the overlay that has the same shape as a base rule
must say `override:` to win. The overlay is a separate file by design, and
the base is never edited to customize the product.

MPL-2.0 draws its line at exactly that seam:

| What you do | MPL obligation |
|---|---|
| Write your own phrasebook file that layers on `english.vla` | none; it is a Larger Work (§3.3) and yours |
| Use `override:` in your file to replace a base rule | none; the replacement text is yours |
| Copy a rule out of `english.vla` into your file and adapt it | that file now contains Covered Software; if you **distribute** it, its source must be available under MPL (§3.1); internal use, none |
| Edit `english.vla` itself and use it internally | none; MPL obligations attach on distribution |
| Edit `english.vla` and ship the result to others | make the edited file's source available under MPL (§3.1, §3.2) |
| Build a product that embeds Frazaro and its phrasebooks | permitted; the MPL files stay MPL, the rest of your product is yours (§3.3) |

## What "distribute" means here

Sending a phrasebook file to another organization, publishing it, or
shipping it inside a product. Sharing a file between colleagues inside one
organization is not distribution in the sense that triggers §3.

## The base corpus is not the product's only language

Every edition phrasebook (`espanol.vla` and its siblings) is MPL-2.0 too. An
organization that authors a new language edition and wants it to stay
private may do so as an overlay; one that wants it to join the shipped
editions contributes it under MPL, see `CONTRIBUTING.md`.

## Contact

Questions about whether a specific use is covered: the address in
`SUPPORT.md`. This file explains the licence; it is not legal advice.
