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

use vla_lang::calc::library;
use vla_lang::machine::{LineRefusal, Machine, Viewed};
use vla_lang::rows::quoted;
use vla_lang::sheet::{number_text, Content, Workbook};
use vla_lang::view::{view_text, window_of};

const FIXTURE: &str = include_str!("../testdata/appearance/appearance.vla");
const DRAWN: &str = include_str!("../testdata/appearance/appearance_drawn.vla");

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
