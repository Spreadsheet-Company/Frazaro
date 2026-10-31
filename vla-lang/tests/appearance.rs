//! A format's appearance in the record (KERNEL.50), held through the public
//! API an outside user has:
//!
//! 1. A format's groups, what it says of a font, an alignment and a border,
//!    print as said, in one order whatever order they were read in, and a
//!    record holding them loads and views again the same record; a format
//!    that says none of them prints as a format printed before them.
//! 2. A group holding what a format does not take is refused by name, with
//!    its line, the group and what it takes; a row out of its shape is
//!    malformed.
//! 3. A `style` row takes a range, each of its cells written by an earlier
//!    row, and the record prints a `style` row a cell; a cell no earlier row
//!    wrote is refused naming it, for a `sentence` row as for a `style` row.
//! 4. The painter's fixture: a record whose looks fall over cells that say
//!    their own font, alignment and borders (`testdata/appearance/
//!    appearance.vla`), and the same window with every look applied by
//!    `Style::over` and the looks dropped (`appearance_drawn.vla`), which a
//!    painter draws alike. `VLA_BLESS=1` writes both again.
//! 5. A format of groups alone (KERNEL.58) says what its groups say and no
//!    more: it prints as said, loads back, draws the cell's own for what it
//!    leaves unsaid, and has a painter's fixture of its own, `partial.vla`
//!    and `partial_drawn.vla`, its `text` rows the drawn record's.

use vla_lang::calc::library;
use vla_lang::machine::{LineRefusal, Machine, Viewed};
use vla_lang::rows::quoted;
use vla_lang::sheet::{number_text, Content, NumFmt, Rgb, Style, Unsaid, Workbook};
use vla_lang::view::{view_text, window_of};

const FIXTURE: &str = include_str!("../testdata/appearance/appearance.vla");
const DRAWN: &str = include_str!("../testdata/appearance/appearance_drawn.vla");
const PARTIAL: &str = include_str!("../testdata/appearance/partial.vla");
const PARTIAL_DRAWN: &str = include_str!("../testdata/appearance/partial_drawn.vla");

fn load(rows: &str) -> Machine {
    Machine::load(rows, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}\n{rows}", e.line, e.refusal.text))
}

fn refused(rows: &str) -> LineRefusal {
    match Machine::load(rows, library::language()) {
        Ok(_) => panic!("loaded:\n{rows}"),
        Err(e) => e,
    }
}

fn grid(m: &Machine, sheet: &str) -> String {
    match m.view("grid", sheet, None) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{other:?}"),
    }
}

fn formats(record: &str) -> Vec<&str> {
    record
        .lines()
        .filter(|l| l.starts_with("(format "))
        .collect()
}

/// Every sheet's grid record of a machine, as a save prints them.
fn save(m: &Machine) -> String {
    (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|name| grid(m, name))
        .collect()
}

#[test]
fn a_format_s_groups_print_as_said_and_a_record_of_them_loads_back() {
    let rows = concat!(
        "(format 1 \"6B4FC8\" general nowrap (align center) (font \"ffffff\" bold))\n",
        "(format 2 none \"0.00\" wrap (border (bottom double \"800020\") (left hair)) (font 10.25 italic))\n",
        "(format 3 \"E8E0F7\" text nowrap (font not-bold not-italic) (align general bottom) (border (top none)))\n",
        "(format 4 \"F7F4FC\" general nowrap)\n",
        "(cell \"S\" \"A1:D1\" 5)\n",
        "(style \"S\" \"A1\" 1)\n",
        "(style \"S\" \"B1\" 2)\n",
        "(style \"S\" \"C1\" 3)\n",
        "(style \"S\" \"D1\" 4)\n",
    );
    let m = load(rows);
    let record = grid(&m, "S");
    assert_eq!(
        formats(&record),
        [
            "(format 0 none general nowrap)",
            "(format 1 \"6B4FC8\" general nowrap (font bold \"FFFFFF\") (align center))",
            "(format 2 none \"0.00\" wrap (font italic 10.25) (border (left hair) (bottom double \"800020\")))",
            "(format 3 \"E8E0F7\" text nowrap (font not-bold not-italic) (align general bottom) (border (top none)))",
            "(format 4 \"F7F4FC\" general nowrap)",
        ],
        "{record}"
    );
    // The record loads, and views again the same record, before a frame and
    // after one.
    let mut again = load(&save(&m));
    assert_eq!(grid(&again, "S"), record);
    let _ = again.step(0);
    assert_eq!(grid(&again, "S"), record);
}

#[test]
fn a_format_said_twice_is_taken_when_it_says_the_same_in_another_order() {
    let rows = concat!(
        "(format 1 none general nowrap (font bold 22) (align center middle))\n",
        "(format 1 none general nowrap (align middle center) (font 22 bold))\n",
        "(cell \"S\" \"A1\" 1)\n",
        "(style \"S\" \"A1\" 1)\n",
    );
    let m = load(rows);
    assert!(
        grid(&m, "S")
            .contains("(format 1 none general nowrap (font bold 22) (align center middle))\n"),
        "{}",
        grid(&m, "S")
    );
    let e = refused(concat!(
        "(format 1 none general nowrap (font bold 22))\n",
        "(format 1 none general nowrap (font bold 20))\n",
    ));
    assert_eq!(
        (e.line, e.refusal.id.as_str()),
        (2, "grid-cell-written-twice")
    );
}

#[test]
fn a_group_holding_what_a_format_does_not_take_is_refused_by_name() {
    let cases: &[(&str, &str, &[&str])] = &[
        (
            "(format 3 none general nowrap (font heavy))\n",
            "grid-format-property",
            &["The format 3's font holds heavy", "bold or not-bold"],
        ),
        (
            "(format 7 none general nowrap (align centre))\n",
            "grid-format-property",
            &[
                "The format 7's align holds centre",
                "top, middle and bottom",
            ],
        ),
        (
            "(format 2 none general nowrap (border (left heavy)))\n",
            "grid-format-property",
            &["The format 2's border holds (left heavy)", "slant-dash-dot"],
        ),
        (
            "(format 2 none general nowrap (font 410))\n",
            "grid-format-property",
            &["holds 410", "from 1 to 409.55"],
        ),
        (
            "(format 2 none general nowrap (font bold) (font 22))\n",
            "grid-row-malformed",
            &["[(font ...)] [(align ...)] [(border ...)]"],
        ),
        (
            "(format 2 none general nowrap (colour \"FFFFFF\"))\n",
            "grid-row-malformed",
            &["[(font ...)]"],
        ),
        (
            "(format 2 none general (font bold))\n",
            "grid-row-malformed",
            &["wrap|nowrap"],
        ),
    ];
    for (rows, id, words) in cases {
        let text = format!("(sheet \"S\" visible)\n{rows}");
        let e = refused(&text);
        assert_eq!((e.line, e.refusal.id.as_str()), (2, *id), "{rows}");
        for w in *words {
            assert!(e.refusal.text.contains(w), "{rows}: {}", e.refusal.text);
        }
    }
}

#[test]
fn a_style_row_takes_a_range_of_written_cells_and_the_record_prints_one_a_cell() {
    let rows = concat!(
        "(format 1 \"E8E0F7\" general wrap (border (left medium)))\n",
        "(cell \"S\" \"A1:C2\" 1)\n",
        "(style \"S\" \"A1:B2\" 1)\n",
        "(style \"S\" \"B1\" 1)\n",
    );
    let m = load(rows);
    let record = grid(&m, "S");
    let styled: Vec<&str> = record
        .lines()
        .filter(|l| l.starts_with("(style "))
        .collect();
    assert_eq!(
        styled,
        [
            "(style \"S\" \"A1\" 1)",
            "(style \"S\" \"B1\" 1)",
            "(style \"S\" \"A2\" 1)",
            "(style \"S\" \"B2\" 1)",
        ],
        "{record}"
    );
    assert_eq!(grid(&load(&save(&m)), "S"), record);
    // A range over a cell a row styled otherwise: the cell's style twice.
    let e = refused(concat!(
        "(cell \"S\" \"A1:B1\" 1)\n",
        "(style \"S\" \"B1\" 2)\n",
        "(style \"S\" \"A1:B1\" 1)\n",
    ));
    assert_eq!(
        (e.line, e.refusal.id.as_str()),
        (3, "grid-cell-written-twice")
    );
    assert!(
        e.refusal.text.contains("(style \"S\" \"B1\")"),
        "{}",
        e.refusal.text
    );
}

#[test]
fn a_style_or_sentence_row_naming_a_cell_no_earlier_row_wrote_is_refused_naming_it() {
    let cases: &[(&str, u32, &str, &str)] = &[
        // A range with a hole: its first unwritten cell, row by row.
        (
            "(cell \"S\" \"A1:B1\" 1)\n(cell \"S\" \"B2\" 1)\n(style \"S\" \"A1:B2\" 1)\n",
            3,
            "S!A2",
            "a style row follows",
        ),
        (
            "(cell \"S\" \"A1\" 1)\n(style \"S\" \"C3\" 1)\n",
            2,
            "S!C3",
            "a style row",
        ),
        // A style row before its cell's own row.
        (
            "(style \"S\" \"A1\" 1)\n(cell \"S\" \"A1\" 1)\n",
            1,
            "S!A1",
            "a style row",
        ),
        // A sheet no row made.
        (
            "(cell \"S\" \"A1\" 1)\n(style \"T\" \"A1:A2\" 1)\n",
            2,
            "T!A1",
            "a style row",
        ),
        (
            "(cell \"S\" \"A1\" 1)\n(sentence \"S\" \"B9\" 4)\n",
            2,
            "S!B9",
            "a sentence row follows",
        ),
    ];
    for (rows, line, cell, words) in cases {
        let e = refused(rows);
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (*line, "grid-cell-unwritten"),
            "{rows}"
        );
        assert!(
            e.refusal.text.contains(&format!("names {cell},")) && e.refusal.text.contains(words),
            "{rows}: {}",
            e.refusal.text
        );
    }
    // A sentence row names one cell; a range is its shape's refusal.
    let e = refused("(cell \"S\" \"A1:A2\" 1)\n(sentence \"S\" \"A1:A2\" 4)\n");
    assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-row-malformed"));
}

/// The painter's fixture's rows: formats that say a font, an alignment and
/// borders as a cell's own and as looks, every line of Excel's fourteen,
/// and a range's style.
fn fixture_rows() -> String {
    let mut rows = String::from(concat!(
        "(sheet \"Look\" visible)\n",
        "(gridlines \"Look\" off)\n",
        "(row \"Look\" 1 30 shown)\n",
        // A cell's own: a box's top-left corner, centered and wrapped.
        "(format 1 none general wrap (align center middle) (border (left medium) (top medium \"6B4FC8\")))\n",
        // A cell's own: a tint, a bold dark ink, aligned right.
        "(format 2 \"E8E0F7\" general nowrap (font bold \"1D1D1B\") (align right))\n",
        // A look: a given's digit, bold and twice the default size.
        "(format 3 none \"\\\"1\\\"\" nowrap (font bold 22))\n",
        // A look: a violation, its ink burgundy and not bold, its top edge taken back.
        "(format 4 \"FFD5D5\" general nowrap (font not-bold \"800020\") (border (top none)))\n",
        // A look that says nothing past its wrap: the cell's own whole, as before.
        "(format 5 \"A080E0\" general nowrap)\n",
        // A cell's own: white bold italic text on the deep lavender, top left.
        "(format 6 \"6B4FC8\" general nowrap (font bold italic 9.5 \"FFFFFF\") (align left top))\n",
        // A look that says the defaults, taking the cell's own alignment back.
        "(format 7 none general nowrap (align general bottom))\n",
    ));
    let lines = [
        "none",
        "thin",
        "medium",
        "thick",
        "double",
        "hair",
        "dotted",
        "dashed",
        "dash-dot",
        "dash-dot-dot",
        "medium-dashed",
        "medium-dash-dot",
        "medium-dash-dot-dot",
        "slant-dash-dot",
    ];
    for (i, line) in lines.iter().enumerate() {
        rows.push_str(&format!(
            "(format {} none general nowrap (border (left {line}) (right {line}) (top {line}) (bottom {line})))\n",
            8 + i
        ));
    }
    rows.push_str(concat!(
        "(look \"Look\" 1 3)\n",
        "(look \"Look\" 2 4)\n",
        "(look \"Look\" 3 5)\n",
        "(look \"Look\" \"\" 0)\n",
        "(look \"Look\" 4 7)\n",
        "(cell \"Look\" \"A1\" 1)\n(style \"Look\" \"A1\" 1)\n",
        "(cell \"Look\" \"B1\" 1)\n(style \"Look\" \"B1\" 2)\n",
        "(cell \"Look\" \"C1\" 2)\n(style \"Look\" \"C1\" 1)\n",
        "(cell \"Look\" \"D1\" 3)\n(style \"Look\" \"D1\" 2)\n",
        "(cell \"Look\" \"E1\" \"1 2 3\")\n(style \"Look\" \"E1\" 1)\n",
        "(cell \"Look\" \"F1\" \"\")\n(style \"Look\" \"F1\" 6)\n",
        "(cell \"Look\" \"G1\" 4)\n(style \"Look\" \"G1\" 6)\n",
    ));
    for (i, line) in lines.iter().enumerate() {
        let addr = format!("{}3", (b'A' + i as u8) as char);
        rows.push_str(&format!(
            "(cell \"Look\" \"{addr}\" \"{line}\")\n(style \"Look\" \"{addr}\" {})\n",
            8 + i
        ));
    }
    rows.push_str("(cell \"Look\" \"A5:C6\" \"box\")\n(style \"Look\" \"A5:C6\" 1)\n");
    rows
}

/// A workbook's sheet with every look applied by `Style::over` to the cells
/// whose values it meets, and its looks dropped: what a painter draws. The
/// fixture's values are spelled as the record spells them, so a look meets a
/// value by its spelling.
fn looks_applied(mut wb: Workbook, sheet: &str) -> Workbook {
    let si = wb.find_sheet(sheet).expect("the sheet");
    let looks = std::mem::take(&mut wb.sheets[si].looks);
    let mut cells: Vec<((u32, u32), u32)> = Vec::new();
    for (&at, cell) in &wb.sheets[si].cells {
        let spelled = match &cell.content {
            Content::Number(n) => number_text(*n),
            Content::Text(t) => Some(quoted(t)),
            _ => None,
        };
        let look = spelled.and_then(|s| looks.iter().find(|(v, _)| *v == s).map(|l| l.1));
        if let Some(look) = look {
            let own = wb.styles.get(cell.style).cloned().unwrap_or_default();
            let draws = wb.styles.get(look).cloned().unwrap_or_default().over(&own);
            cells.push((at, wb.styles.id(draws)));
        }
    }
    for (at, style) in cells {
        if let Some(cell) = wb.sheets[si].cells.get_mut(&at) {
            cell.style = style;
        }
    }
    wb
}

#[test]
fn the_painter_s_fixture_is_the_language_s() {
    let m = load(&fixture_rows());
    let record = grid(&m, "Look");
    let wb = looks_applied(m.workbook(), "Look");
    let window = window_of(&wb, "Look", None).expect("the window");
    let drawn = view_text(&wb, &window);
    if std::env::var("VLA_BLESS").as_deref() == Ok("1") {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/testdata/appearance/");
        std::fs::write(format!("{dir}appearance.vla"), record.replace('\n', "\r\n")).unwrap();
        std::fs::write(
            format!("{dir}appearance_drawn.vla"),
            drawn.replace('\n', "\r\n"),
        )
        .unwrap();
        return;
    }
    assert_eq!(
        record,
        FIXTURE.replace("\r\n", "\n"),
        "the fixture's record"
    );
    assert_eq!(drawn, DRAWN.replace("\r\n", "\n"), "its looks applied");
    // Each loads and views again the same record, the drawn one holding no look.
    assert_eq!(grid(&load(&save(&m)), "Look"), record);
    assert_eq!(grid(&load(&drawn), "Look"), drawn);
    assert!(!drawn.contains("(look "));
    // What a look said and what the cell's own did, cell by cell: the
    // given's digit keeps its box's corner and centering; the bold dark ink
    // keeps its color under a look of a size; the violation takes the top
    // edge back and keeps the left; a look of no group draws the cell's own
    // font and alignment; an emptied cell takes its look's lack of fill and
    // keeps its own white text; a look of the defaults takes alignment back.
    for (addr, want) in [
        ("A1", "none \"\\\"1\\\"\" nowrap (font bold 22) (align center middle) (border (left medium) (top medium \"6B4FC8\"))"),
        ("B1", "none \"\\\"1\\\"\" nowrap (font bold 22 \"1D1D1B\") (align right)"),
        ("C1", "\"FFD5D5\" general nowrap (font not-bold \"800020\") (align center middle) (border (left medium) (top none))"),
        ("D1", "\"A080E0\" general nowrap (font bold \"1D1D1B\") (align right)"),
        ("F1", "none general nowrap (font bold italic 9.5 \"FFFFFF\") (align left top)"),
        ("G1", "none general nowrap (font bold italic 9.5 \"FFFFFF\") (align general bottom)"),
    ] {
        let style = drawn
            .lines()
            .find_map(|l| l.strip_prefix(&format!("(style \"Look\" \"{addr}\" ")))
            .and_then(|n| n.strip_suffix(')'))
            .unwrap_or_else(|| panic!("{addr}: no style row\n{drawn}"));
        let row = format!("(format {style} {want})");
        assert!(drawn.lines().any(|l| l == row), "{addr}: {row}\n{drawn}");
    }
}

// ---- a format of groups alone (KERNEL.58) -----------------------------------

#[test]
fn a_format_of_groups_alone_prints_what_it_says_and_loads_back() {
    let rows = concat!(
        "(format 1 (font \"800020\"))\n",
        "(format 2 (font bold) (fill \"ffe699\"))\n",
        "(format 3 (wrap wrap) (number \"0.00\"))\n",
        "(format 4 (border (top thin)) (align center) (fill none))\n",
        "(format 5 (number text))\n",
        // All three said by groups: a format that spells them, as it prints.
        "(format 6 (wrap nowrap) (number general) (fill none))\n",
        "(format 7 (fill \"F7F4FC\") (number \"0\") (wrap wrap) (font italic))\n",
        "(cell \"S\" \"A1:G1\" 5)\n",
        "(style \"S\" \"A1\" 1)\n",
        "(style \"S\" \"B1\" 2)\n",
        "(style \"S\" \"C1\" 3)\n",
        "(style \"S\" \"D1\" 4)\n",
        "(style \"S\" \"E1\" 5)\n",
        "(style \"S\" \"F1\" 6)\n",
        "(style \"S\" \"G1\" 7)\n",
    );
    let m = load(rows);
    let record = grid(&m, "S");
    assert_eq!(
        formats(&record),
        [
            "(format 0 none general nowrap)",
            "(format 1 (font \"800020\"))",
            "(format 2 (fill \"FFE699\") (font bold))",
            "(format 3 (number \"0.00\") (wrap wrap))",
            "(format 4 (fill none) (align center) (border (top thin)))",
            "(format 5 (number text))",
            "(format 6 none general nowrap)",
            "(format 7 \"F7F4FC\" \"0\" wrap (font italic))",
        ],
        "{record}"
    );
    let mut again = load(&save(&m));
    assert_eq!(grid(&again, "S"), record);
    let _ = again.step(0);
    assert_eq!(grid(&again, "S"), record);
    // Said twice alike, in another order and spelled otherwise, it is taken; a
    // group more or less is another format.
    let m = load(concat!(
        "(format 2 (fill \"ffe699\") (font bold))\n",
        "(format 2 (font bold) (fill \"FFE699\"))\n",
        "(format 6 none general nowrap)\n",
        "(format 6 (fill none) (number general) (wrap nowrap))\n",
    ));
    drop(m);
    for rows in [
        "(format 1 (fill none))\n(format 1 none general nowrap)\n",
        "(format 1 (fill \"FFE699\"))\n(format 1 (fill \"FFE699\") (wrap wrap))\n",
    ] {
        let e = refused(rows);
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "grid-cell-written-twice"),
            "{rows}"
        );
    }
}

#[test]
fn a_format_of_groups_alone_holding_what_it_does_not_take_is_refused_by_name() {
    let cases: &[(&str, &str, &[&str])] = &[
        (
            "(format 3 (fill red))\n",
            "grid-format-property",
            &["The format 3's fill holds red", "six hex digits"],
        ),
        (
            "(format 3 (number wide))\n",
            "grid-format-property",
            &["The format 3's number holds wide", "general, text"],
        ),
        (
            "(format 3 (wrap maybe))\n",
            "grid-format-property",
            &["The format 3's wrap holds maybe", "wrap and nowrap"],
        ),
        (
            "(format 3 (font heavy))\n",
            "grid-format-property",
            &["The format 3's font holds heavy"],
        ),
        // A code this version does not draw, by the refusal a code has.
        (
            "(format 3 (number \"0.00E+00\"))\n",
            "grid-format-code",
            &["0.00E+00"],
        ),
        // Out of the row's shape: nothing said, a group twice, an item no
        // group, a group of no such head, and the three said beside the places
        // that spell them.
        (
            "(format 3)\n",
            "grid-row-malformed",
            &["a format of groups alone", "saying at least one"],
        ),
        ("(format 3 ())\n", "grid-row-malformed", &["(fill ...)"]),
        (
            "(format 3 (fill none) (fill none))\n",
            "grid-row-malformed",
            &["(fill ...)"],
        ),
        (
            "(format 3 (font bold) general)\n",
            "grid-row-malformed",
            &["(fill ...)"],
        ),
        (
            "(format 3 (colour \"FFFFFF\"))\n",
            "grid-row-malformed",
            &["(fill ...)"],
        ),
        (
            "(format 3 (font bold) none general nowrap)\n",
            "grid-row-malformed",
            &["(fill ...)"],
        ),
        (
            "(format 3 none general nowrap (fill \"FFFFFF\"))\n",
            "grid-row-malformed",
            &["(fill ...)"],
        ),
        (
            "(format 3 none general nowrap (number text))\n",
            "grid-row-malformed",
            &["(number ...)"],
        ),
        (
            "(format 3 none general nowrap (wrap wrap))\n",
            "grid-row-malformed",
            &["(wrap ...)"],
        ),
    ];
    for (rows, id, words) in cases {
        let text = format!("(sheet \"S\" visible)\n{rows}");
        let e = refused(&text);
        assert_eq!((e.line, e.refusal.id.as_str()), (2, *id), "{rows}");
        for w in *words {
            assert!(e.refusal.text.contains(w), "{rows}: {}", e.refusal.text);
        }
    }
}

#[test]
fn a_format_says_only_what_it_says_over_a_cell_s_own() {
    let mut own = Style::new(
        Some(Rgb(1, 2, 3)),
        NumFmt::Code(vla_lang::numfmt::Code::parse("0.00").unwrap()),
        true,
    );
    own.font.bold = Some(true);
    // A format of a row's places draws its three whatever their values, as
    // every format before this item did.
    let whole = Style::new(None, NumFmt::General, false);
    assert_eq!(whole.unsaid, Unsaid::NOTHING);
    let drawn = whole.over(&own);
    assert_eq!(
        (drawn.fill, &drawn.num_fmt, drawn.wrap, drawn.font.bold),
        (None, &NumFmt::General, false, Some(true))
    );
    assert_eq!(drawn.unsaid, Unsaid::NOTHING);
    // One of groups alone leaves all three to the cell.
    let partial = Style::partial();
    assert_eq!(partial.unsaid, Unsaid::ALL);
    assert_eq!(partial.over(&own), own);
    // Each said in turn, whatever its value: `none`, general and nowrap take
    // the cell's own away as any other would.
    let mut fill = Style::partial();
    fill.say_fill(None);
    let drawn = fill.over(&own);
    assert_eq!((drawn.fill, drawn.wrap), (None, true));
    assert_eq!(drawn.num_fmt, own.num_fmt);
    let mut number = Style::partial();
    number.say_number(NumFmt::General);
    let drawn = number.over(&own);
    assert_eq!(
        (drawn.fill, &drawn.num_fmt, drawn.wrap),
        (own.fill, &NumFmt::General, true)
    );
    let mut wrap = Style::partial();
    wrap.say_wrap(false);
    let drawn = wrap.over(&own);
    assert_eq!((drawn.fill, drawn.wrap), (own.fill, false));
    assert_eq!(drawn.num_fmt, own.num_fmt);
    // Formats folded before the cell's own: what either says, the first
    // winning, and what neither says left unsaid until the cell's own takes it.
    let mut first = Style::partial();
    first.say_fill(Some(Rgb(9, 9, 9)));
    first.font.bold = Some(false);
    let mut second = Style::partial();
    second.say_fill(Some(Rgb(8, 8, 8)));
    second.say_wrap(false);
    second.font.italic = Some(true);
    let folded = first.over(&second);
    assert_eq!(folded.fill, Some(Rgb(9, 9, 9)));
    assert!(folded.unsaid.number() && !folded.unsaid.fill() && !folded.unsaid.wrap());
    assert_eq!(
        (folded.font.bold, folded.font.italic),
        (Some(false), Some(true))
    );
    let drawn = folded.over(&own);
    assert_eq!(drawn.unsaid, Unsaid::NOTHING);
    assert_eq!((drawn.fill, drawn.wrap), (Some(Rgb(9, 9, 9)), false));
    assert_eq!(drawn.num_fmt, own.num_fmt);
    assert_eq!(
        drawn,
        first.over(&second.over(&own)),
        "the fold's two groupings"
    );
    // What it leaves unsaid a byte in its padding, as an underline and a
    // strike are, and a family a pointer (KERNEL.58's second slice).
    assert_eq!(std::mem::size_of::<Style>(), 64);
    assert_eq!(std::mem::size_of::<vla_lang::sheet::Font>(), 12);
}

#[test]
fn a_font_s_family_underline_and_strike_print_as_said_and_fall_over_a_cell_s_own() {
    let rows = concat!(
        "(format 1 none general nowrap (font (family \"Consolas\" mono) strike underline))\n",
        "(format 2 (font double-underline))\n",
        "(format 3 (font not-underline not-strike (family \"Arial\")))\n",
        "(format 4 \"E8E0F7\" general nowrap (font (family \"Georgia\" serif) accounting-underline))\n",
        "(cell \"S\" \"A1:D1\" 5)\n",
        "(style \"S\" \"A1\" 1)\n",
        "(style \"S\" \"B1\" 2)\n",
        "(style \"S\" \"C1\" 3)\n",
        "(style \"S\" \"D1\" 4)\n",
    );
    let m = load(rows);
    let record = grid(&m, "S");
    assert_eq!(
        formats(&record),
        [
            "(format 0 none general nowrap)",
            "(format 1 none general nowrap (font underline strike (family \"Consolas\" mono)))",
            "(format 2 (font double-underline))",
            "(format 3 (font not-underline not-strike (family \"Arial\")))",
            "(format 4 \"E8E0F7\" general nowrap (font accounting-underline (family \"Georgia\" serif)))",
        ],
        "{record}"
    );
    let mut again = load(&save(&m));
    assert_eq!(grid(&again, "S"), record);
    let _ = again.step(0);
    assert_eq!(grid(&again, "S"), record);
    // Over a cell's own, the first that says each.
    let wb = m.workbook();
    let style = |k: u32| wb.styles.get(k).expect("the format");
    let own = style(1);
    let family = |s: &Style| s.family.as_deref().map(|f| f.name.clone());
    assert_eq!(family(&style(2).over(own)).as_deref(), Some("Consolas"));
    assert_eq!(family(&style(3).over(own)).as_deref(), Some("Arial"));
    let drawn = style(2).over(own);
    assert_eq!(
        (drawn.font.underline, drawn.font.strike),
        (Some(vla_lang::sheet::Underline::Double), Some(true))
    );
    let drawn = style(3).over(own);
    assert_eq!(
        (drawn.font.underline, drawn.font.strike),
        (Some(vla_lang::sheet::Underline::None), Some(false))
    );
    assert_eq!(
        style(1).family.as_deref().and_then(|f| f.generic),
        Some(vla_lang::sheet::Generic::Mono)
    );
    assert!(style(2).says_appearance());
    let mut only_family = Style::partial();
    only_family.family = own.family.clone();
    assert!(only_family.says_appearance());
    assert_ne!(only_family, Style::partial(), "a family is said");
    // Said twice alike and in another order, taken; another name, refused.
    load(concat!(
        "(format 1 none general nowrap (font strike (family \"A\" sans)))\n",
        "(format 1 none general nowrap (font (family \"A\" sans) strike))\n",
    ));
    for rows in [
        "(format 1 (font (family \"A\")))\n(format 1 (font (family \"B\")))\n",
        "(format 1 (font (family \"A\")))\n(format 1 (font (family \"A\" sans)))\n",
        "(format 1 (font strike))\n(format 1 (font not-strike))\n",
    ] {
        let e = refused(rows);
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "grid-cell-written-twice"),
            "{rows}"
        );
    }
    let e = refused("(sheet \"S\" visible)\n(format 3 (font (family \"\")))\n");
    assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-format-property"));
    assert!(
        e.refusal.text.contains("(family \"<name>\""),
        "{}",
        e.refusal.text
    );
}

/// The second fixture's rows: a cell's own format that says everything, and
/// looks of groups alone that each say one thing over it and over format 0.
fn partial_rows() -> String {
    let mut rows = String::from(concat!(
        "(sheet \"Partial\" visible)\n",
        "(gridlines \"Partial\" off)\n",
        // A cell's own: a tint, a code, a wrap, a bold dark ink, aligned right,
        // a left edge.
        "(format 1 \"E8E0F7\" \"0.00\" wrap (font bold \"1D1D1B\") (align right) (border (left medium)))\n",
        // A fill alone.
        "(format 2 (fill \"FFE699\"))\n",
        // A text's color alone.
        "(format 3 (font \"800020\"))\n",
        // A number format alone.
        "(format 4 (number \"0\"))\n",
        // A wrap alone.
        "(format 5 (wrap nowrap))\n",
        // The fill taken away, and a said default of a font.
        "(format 6 (fill none) (font not-bold))\n",
        // All three by groups: a format that spells them.
        "(format 7 (fill \"6B4FC8\") (number \"0.000\") (wrap wrap))\n",
        // An edge taken back, and another said.
        "(format 8 (border (left none) (top thin \"6B4FC8\")))\n",
        // No number format in a code: general, taking the cell's code away.
        "(format 9 (number general))\n",
    ));
    for (value, format) in [
        (1, 2),
        (2, 3),
        (3, 4),
        (4, 5),
        (5, 6),
        (6, 7),
        (7, 8),
        (8, 9),
    ] {
        rows.push_str(&format!("(look \"Partial\" {value} {format})\n"));
    }
    rows.push_str("(look \"Partial\" \"a text\" 2)\n");
    for (k, col) in ('A'..='H').enumerate() {
        let value = k + 1;
        rows.push_str(&format!(
            "(cell \"Partial\" \"{col}1\" {value})\n(style \"Partial\" \"{col}1\" 1)\n"
        ));
        rows.push_str(&format!("(cell \"Partial\" \"{col}2\" {value})\n"));
    }
    rows.push_str("(cell \"Partial\" \"A3\" \"a text\")\n(style \"Partial\" \"A3\" 1)\n");
    rows.push_str("(cell \"Partial\" \"B3\" \"no look\")\n(style \"Partial\" \"B3\" 1)\n");
    rows.push_str("(cell \"Partial\" \"C3\" 9)\n(style \"Partial\" \"C3\" 1)\n");
    rows
}

#[test]
fn the_painter_s_fixture_of_groups_alone_is_the_language_s() {
    let m = load(&partial_rows());
    let record = grid(&m, "Partial");
    let wb = looks_applied(m.workbook(), "Partial");
    let window = window_of(&wb, "Partial", None).expect("the window");
    let drawn = view_text(&wb, &window);
    if std::env::var("VLA_BLESS").as_deref() == Ok("1") {
        let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/testdata/appearance/");
        std::fs::write(format!("{dir}partial.vla"), record.replace('\n', "\r\n")).unwrap();
        std::fs::write(
            format!("{dir}partial_drawn.vla"),
            drawn.replace('\n', "\r\n"),
        )
        .unwrap();
        return;
    }
    assert_eq!(
        record,
        PARTIAL.replace("\r\n", "\n"),
        "the fixture's record"
    );
    assert_eq!(
        drawn,
        PARTIAL_DRAWN.replace("\r\n", "\n"),
        "its looks applied"
    );
    assert_eq!(grid(&load(&save(&m)), "Partial"), record);
    assert_eq!(grid(&load(&drawn), "Partial"), drawn);
    assert!(!drawn.contains("(look "));
    // The drawn record says all three of every format, and each number's text
    // is the same in both: a look that leaves its number format unsaid draws
    // under the cell's own code, one that says one under its own.
    assert!(!drawn.contains("(fill "), "{drawn}");
    let texts = |r: &str| -> Vec<String> {
        r.lines()
            .filter(|l| l.starts_with("(text "))
            .map(str::to_string)
            .collect()
    };
    assert_eq!(texts(&record), texts(&drawn));
    for (addr, text) in [
        ("A1", Some("1.00")),
        ("B1", Some("2.00")),
        ("C1", Some("3")),
        ("D1", Some("4.00")),
        ("E1", Some("5.00")),
        ("F1", Some("6.000")),
        ("G1", Some("7.00")),
        // A look's said general takes the cell's code away.
        ("H1", None),
        ("A2", None),
        ("C2", Some("3")),
        ("F2", Some("6.000")),
        ("C3", Some("9.00")),
    ] {
        let row = format!("(text \"Partial\" \"{addr}\" ");
        let got = record
            .lines()
            .find_map(|l| l.strip_prefix(&row))
            .and_then(|t| t.strip_suffix(')'))
            .map(|t| t.trim_matches('"').to_string());
        assert_eq!(got.as_deref(), text, "{addr}\n{record}");
    }
    // What each look said and what each cell's own did, cell by cell.
    let own = "(font bold \"1D1D1B\") (align right) (border (left medium))";
    let cases: Vec<(&str, String)> = vec![
        ("A1", format!("\"FFE699\" \"0.00\" wrap {own}")),
        (
            "B1",
            "\"E8E0F7\" \"0.00\" wrap (font bold \"800020\") (align right) (border (left medium))"
                .to_string(),
        ),
        ("C1", format!("\"E8E0F7\" \"0\" wrap {own}")),
        ("D1", format!("\"E8E0F7\" \"0.00\" nowrap {own}")),
        (
            "E1",
            "none \"0.00\" wrap (font not-bold \"1D1D1B\") (align right) (border (left medium))"
                .to_string(),
        ),
        ("F1", format!("\"6B4FC8\" \"0.000\" wrap {own}")),
        (
            "G1",
            "\"E8E0F7\" \"0.00\" wrap (font bold \"1D1D1B\") (align right) (border (left none) (top thin \"6B4FC8\"))"
                .to_string(),
        ),
        ("H1", format!("\"E8E0F7\" general wrap {own}")),
        ("A2", "\"FFE699\" general nowrap".to_string()),
        ("B2", "none general nowrap (font \"800020\")".to_string()),
        ("C2", "none \"0\" nowrap".to_string()),
        ("F2", "\"6B4FC8\" \"0.000\" wrap".to_string()),
        ("A3", format!("\"FFE699\" \"0.00\" wrap {own}")),
    ];
    for (addr, want) in cases {
        let style = drawn
            .lines()
            .find_map(|l| l.strip_prefix(&format!("(style \"Partial\" \"{addr}\" ")))
            .and_then(|n| n.strip_suffix(')'))
            .unwrap_or_else(|| panic!("{addr}: no style row\n{drawn}"));
        let row = format!("(format {style} {want})");
        assert!(drawn.lines().any(|l| l == row), "{addr}: {row}\n{drawn}");
    }
    // Under format 0 with no look, nothing moved: D2 is a wrap of no effect.
    assert!(!drawn.contains("(style \"Partial\" \"D2\""), "{drawn}");
    // A cell its look does not meet keeps its own, in both.
    for r in [&record, &drawn] {
        assert!(r.contains("(style \"Partial\" \"B3\" 1)"), "{r}");
    }
}

#[test]
fn a_format_of_groups_alone_as_a_cell_s_own_draws_the_default_for_what_it_leaves_unsaid() {
    // `Style::id` tells it from the default, and it holds the defaults it
    // leaves unsaid, so a writer writing its three writes the default's.
    let m = load(concat!(
        "(format 1 (font bold))\n",
        "(cell \"S\" \"A1\" 5)\n",
        "(style \"S\" \"A1\" 1)\n",
    ));
    let wb = m.workbook();
    let style = wb.styles.get(1).expect("format 1");
    assert_eq!(style.unsaid, Unsaid::ALL);
    assert_eq!(
        (style.fill, &style.num_fmt, style.wrap),
        (None, &NumFmt::General, false)
    );
    assert_eq!(style.font.bold, Some(true));
    assert_ne!(*style, Style::default());
    let mut said_the_same = Style::new(None, NumFmt::General, false);
    said_the_same.font.bold = Some(true);
    assert_ne!(*style, said_the_same, "unsaid is not said");
    assert_eq!(style.over(&Style::default()), said_the_same);
}
