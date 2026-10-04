# The web door

*Slice 2 of `docs/HORIZON.md` section 12, the room of section 11.4 seen through
its second door: a sentence in column B, what it means in column C.*

`index.template.html` is the source. `tools/build_web.ps1` builds the core for
the browser, checks that the module imports nothing, and fills the template's
three placeholders with the module as base64, `scripts/prelude.vla` and
`scripts/polyglotta/english.vla` as text; the result, `index.html`, is one
file that runs from disk, a build artifact like the add-in's `.xlam`
(`.gitignore`), kept by the core CI job on every push.

What the page does: the lines typed under Input are the program, in the
language the picker names (English, or a dialect loaded over `english.vla` as
the add-in loads an edition; each dialect is inside the page too, from
`{{BOOK:name}}` placeholders the builder fills); the core translates them as
the add-in's Check does, and the VLA column shows each row's form (and the
formula it writes, when it writes one), with a refusal in words on the row it
belongs to; the whole VLA and the whole VBA stand below, with a Copy button
each. The focused input cell is Frazaro Lavender, the colour the add-in
paints its input column a whisper of. Another phrasebook can be pasted to load
after the language's, and a phrasebook that holds a `(raw ...)` form loads
only once the person ticks the consent box, the same gate the add-in asks at
its file loader and the CLI asks with `--allow-raw`.

Under the two panes stands the workbook (`PORT.7`, slice 7e): the core builds
it from the rows on every change, as `frazaro build` builds it from a file,
and the strip shows its byte count and SHA-256, the digest the command-line
door prints for the same sentences. *Download as .xlsx* hands the browser
those bytes as a `data:` address named `program.xlsx`, to save where the
person chooses; the file opens in Excel, Sheets, Calc or Numbers with nothing
installed, and `frazaro rebuild` on it says yes. A sentence the writer cannot
hold with nothing running, a loop, a message, a value read from a cell, is
named in the strip with its line, and the button waits; the VLA and VBA
beside it are unaffected, since the translation stands.

What the page does not do: it loads nothing from the network, writes nothing
anywhere, and executes nothing; the download is the browser saving one file
the person asked for. The VBA is text to paste into a module in Excel, or to
leave to the add-in. `tools/check_web_offline.ps1` holds the template to that
doctrine, as `tools/check_core_imports.ps1` holds the core it carries.

The page speaks to the core through `core/src/abi.rs`: an allocator pair, the
two translate functions, the build, the gate and the version, every answer one
record in the module's memory (four little-endian `u32`, status, line, id
length and text length, then the id and the text; a build's status-0 record
is the one whose text is bytes, the workbook, with their SHA-256 as its id).
No binding layer, no generated glue.
