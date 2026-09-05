# Frazaro Output Exception

*An additional permission, granted by the copyright holder, on top of every
licence in this repository. It exists because Frazaro copies text into your
workbook: a runtime-helper module (`Frazaro_EN_Runtime`, from
`VLA_Runtime.bas`) and generated code assembled from `prelude.vla` macro
bodies and phrasebook templates. Without this file, a lawyer could ask
whether your workbook inherits those files' licences. The answer is no.*

As an additional permission under the Apache License 2.0, the Mozilla Public
License 2.0, the BSD Zero Clause License, and any other licence applied to
files in this repository (`REUSE.toml` says which applies where):

1. **Your programs are yours.** English sentences, `.vla` programs, and
   worksheets you write are your own work. Frazaro's licences do not apply
   to them.

2. **Generated code is yours.** VBA modules, worksheet formulas, or other
   output that Frazaro generates from your programs, including any text
   copied into that output from `prelude.vla`, from a phrasebook template
   (whether from `english.vla`, an edition phrasebook, or your own), or from
   the injectable runtime helpers, may be used, modified, and distributed
   under terms of your choosing, without attribution and without any
   obligation under Apache-2.0, MPL-2.0, 0BSD, or any other licence in this
   repository.

3. **Workbooks are not derivative works.** Distributing a workbook that
   contains generated code or the injected runtime helpers does not make the
   workbook, or any other code in it, subject to Frazaro's licences.

This exception does not permit removing or altering licence notices in a
copy of Frazaro *itself* (the add-in, its source, or its phrasebooks), and it
grants no right in the Frazaro name (see `TRADEMARK.md`).

*Modelled on the GCC Runtime Library Exception and the Bison exception,
which exist for the same reason: a tool whose output contains fragments of
the tool.*
