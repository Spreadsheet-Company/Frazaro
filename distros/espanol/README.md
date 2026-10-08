# The espanol distro

The add-in's Spanish edition as it ships today, named by reference:
`english.vla` first, then `espanol.vla` as the overlay, which rides
`english.vla`'s macros, so the chain `VlaBuildAddin "Espanol"` audits and
embeds is the one `VLA_Build.bas` named by hand before this folder existed,
and the file a person may put beside their workbook to override the embedded
phrasebook is `espanol.vla`, the last of the chain.

The shape of `distro.vla`, the three doors that take it and the terms per
kind of file are `distros/english/README.md`'s: the phrasebooks it points
at keep their own file's MPL-2.0 where they sit, and this manifest and
README are Apache-2.0 with the tools. Its chrome is English, as the
Spanish edition's is today (`EDITION-CHROME`); `EDITION-ESPANOL` fills the
rest in. `frazaro prove distros\espanol` proves it whole: the base alone, the
overlay over the base, then each dialect over the base.
