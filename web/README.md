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

What the page does not do: it loads nothing from the network, writes nothing
anywhere, and executes nothing. The VBA is text to paste into a module in Excel,
or to leave to the add-in. `tools/check_web_offline.ps1` holds the template to
that doctrine, as `tools/check_core_imports.ps1` holds the core it carries.

The page speaks to the core through `core/src/abi.rs`: an allocator pair, the
two translate functions, the gate and the version, every answer one record in
the module's memory (four little-endian `u32`, status, line, id length and text
length, then the id and the text). No binding layer, no generated glue.
