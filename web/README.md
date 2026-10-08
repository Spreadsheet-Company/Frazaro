# The web page

**Frazaro in a browser tab: sentences in, a workbook out, with nothing
installed and nothing fetched.**

The page is Frazaro's second door. The same engine that runs inside the
Excel add-in is compiled to WebAssembly and sealed inside one HTML file with
the prelude and the phrasebooks, so the file works from a double-click on
your desk as well as from its published address, and works the same in both
places: it reads nothing from the network, writes nothing anywhere, and
executes nothing.

## Try it

The repository publishes the page at each release, at
<https://spreadsheet-company.github.io/Frazaro/>, the address
`spreadsheet.company/frazaro` opens. From a clone, `web/index.html` is the
same file once built (see *Building it*); open it from disk.

## What you can do on it

- **Write a program, one sentence per row.** The Input column takes a
  sentence a row. Enter starts a new row, Backspace at the start of a row
  joins it to the one above, the arrow keys move between rows, Tab indents by
  two spaces (a block), and pasting several lines fills several rows. As you
  type, the VLA column shows what each row means, the form the add-in's Check
  would make of it, with the formula it writes when it writes one; a sentence
  the language cannot read is refused in words on its own row, and the status
  line says how many rows translated, or which row was refused.
- **Pick a language.** The picker at the head of the Input column offers
  English and seven dialects (Dansk, Deutsch, Español, Esperanto, Français,
  Latin and Pirate), each a phrasebook inside the page that loads over
  `english.vla` as the add-in loads an edition.
- **Bring your own phrasebook.** Under *Phrasebooks*, paste one to load after
  the language's. It is read, never saved. A phrasebook holding a `(raw ...)`
  form loads only once you tick the consent box, the same gate the add-in asks
  at its file loader and the command line asks with `--allow-raw`; one that
  needs a capability is refused, since none can be granted yet.
- **Read the whole program.** Two panes hold the whole VLA and the whole VBA,
  the text Compile writes into a module, with a Copy button each. The VBA is
  text: paste it into a module in Excel, or leave that to the add-in.
- **Download the workbook.** The Workbook strip builds the workbook your rows
  describe on every change, exactly as `frazaro build` builds it from a file,
  and shows its byte count and SHA-256, the digest the command line prints for
  the same sentences. *Download as .xlsx* turns Frazaro Lavender when a
  workbook is ready and hands the browser the bytes as a file named
  `program.xlsx`, to save where you choose; it opens in Excel, Sheets, Calc or
  Numbers with nothing installed, and `frazaro rebuild` on it says yes. A
  sentence the writer cannot hold with nothing running, a loop, a message, a
  value read from a cell, is named in the strip with its line, and the button
  waits; the VLA and VBA stand, since the translation does. So is a formula
  that would reach outside the workbook on its own (`WEBSERVICE`,
  `HYPERLINK`, a DDE link and their kin), refused by name as the add-in
  refuses it.
- **Read a workbook of your own.** Under *Reflect*, pick a workbook, `.xlsx`
  or `.ods`, and the pane prints its relations as `frazaro reflect` prints
  them. The *Show* picker beside it prints the audit list instead, as `frazaro
  audit` does, or what changed from an earlier copy picked in the second
  input, as `frazaro diff` does, with a Copy button and a line under the pane
  giving the count or the refusal. The browser reads the picked file into
  memory and the page copies the bytes into the engine's memory, where the
  same reader runs under the same bounds. The file goes nowhere.
- **Go back.** The bare link at the very top, `https://spreadsheet.company`,
  is the page's one link, followed only when you click it.

The order down the page: the link, the title with its tagline, a short
welcome, and one line the engine fills once it has loaded (the version and
the module's size); then Phrasebooks, folded; the two panes; the Workbook
strip; the Reflect pane; and the status line with the rows last, so that a
long program grows downward and pushes nothing below it.

## What it promises

Nothing you type leaves the page. It loads nothing from the network, writes
nothing anywhere, and executes nothing: the download is the browser saving
one file you asked for, and a picked workbook is read in the browser and goes
nowhere. The engine inside cannot phone home, because its WebAssembly module
imports nothing, which `tools/check_core_imports.ps1` reads off the built
module; and `tools/check_web_offline.ps1` holds the page around it to the
same doctrine on every push: no script, stylesheet, font, image, frame, form
or fetch, and one anchor, the link above, pinned whole, as spelled, once.

## Building it

```powershell
powershell -File tools\build_web.ps1
```

builds the engine for the browser (`cargo build --release -p frazaro-core
--target wasm32-unknown-unknown`), checks that the module imports nothing (a
page that could phone home is not written), reads the english distro
(`distros/english/distro.vla`, or the folder `-Distro` names; what a distro
is, and how to make your own, is `distros/english/README.md`), and fills
the placeholders of `index.template.html` from it: the module as base64,
the distro's prelude and base phrasebook as text, each `{{BOOK:name}}` of
the language picker with the distro's dialect of that name (the two sets
must agree exactly), its title, tagline and opening sentence, and the three
shades of its palette. The result,
`index.html`, is a build artifact like the add-in's `.xlam` (gitignored): the
template is the source. `-NoBuild` takes the module already built, as CI
does. The core CI job builds the page on every push and keeps it as an
artifact, and `.github/workflows/pages.yml` builds and publishes it at each
release; `docs/DEPLOY.md` ("Building the web page" and "Hosting it") has the
detail.

## For contributors

- **The template is the source**, `index.template.html`: the markup, the
  styles and the script in one file, plus the placeholders the builder fills
  from the distro. The three Frazaro Lavender shades are CSS variables at
  its top, filled from the distro's palette (`distros/english/distro.vla`),
  the one place to change them; the focused input cell takes the light
  shade, the ready Download button the deep one.
- **The page speaks to the engine through `core/src/abi.rs`**, with no
  binding layer and no generated glue: an allocator pair;
  `frazaro_translate_vla` and `frazaro_translate_vba`; `frazaro_build_xlsx`;
  the reader's three, `frazaro_reflect`, `frazaro_audit` and `frazaro_diff`,
  a file's bytes and its name in and the lines the command line prints out;
  `frazaro_view`, the build's inputs with a sheet and a window in and the
  view record out, the lines a viewport draws from (the page's own viewport
  is `KERNEL.5` on the roadmap; nothing on the page calls it yet);
  `frazaro_vocab_gate`; and the version. Every answer is one record in the
  module's memory: four little-endian `u32` (status, line, id length, text
  length), then the id and the text. A build's status-0 record is the one
  whose text is bytes, the workbook, with their SHA-256 as its id.
- **Two checks hold it.** `tools/check_web_offline.ps1` holds the template,
  and the built page beside it, to the promise above, with a `-Control` that
  proves the check on five mutants; `tools/check_core_imports.ps1` holds the
  module's import section at zero entries. A file input is no reference, so
  the Reflect pane needed no exception.
- **The command line is the same engine behind a different door**, held to
  the same goldens: `cli/README.md` is its reference, and `frazaro build`,
  `rebuild`, `reflect`, `audit` and `diff` print what the page shows.
- **Where the page is going** is `CALLOSUM.md`, beside this file: the sitting
  of 2026-10-05 that scoped its growth into a two-pane spreadsheet in Elm's
  shape, a grid that is only a view of the sentences, with the forecast, the
  reflect feature read as a Lisp, and seven moves beyond the roadmap. Nothing
  in it is an item; it is written in `docs/HORIZON.md`'s register and marks.
