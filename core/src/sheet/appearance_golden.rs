//! The writer's half of the appearance golden (KERNEL.50), test-only.
//!
//! 1. `scripts/appearance.txt`'s cases are written as a workbook,
//!    `scripts/appearance_written.xlsx`, through the record, the language's
//!    loader and this crate's writer, a case every other row, its name in
//!    column A and a cell drawn with it in column B, in the fixture's order,
//!    so that `VlaWriteAppearanceGolden` (VLA_Tests.bas) reads in Excel what
//!    Excel reads of each written cell, and no two of the cells share an
//!    edge. The file is held byte for byte; `FRAZARO_BLESS=1` writes it
//!    again.
//! 2. `scripts/appearance_saved.xlsx` is Excel's own workbook of the same
//!    cases, each set through its object model on the same row, saved by
//!    `VlaWriteAppearanceGolden`; each cell's font, fill, border and
//!    alignment there, read from its styles part, is the one the writer
//!    writes for the case, property by property: a weight, a slant, a size
//!    in twentieths of a point and an ink of its own; a fill; each edge's
//!    line and color of its own; and the alignment and the wrap. The font's
//!    family, Excel's Aptos Narrow beside the writer's Calibri, is a
//!    workbook's and no case's.
//!
//! The files are read at run time, outside the crate, which the core's
//! package does not test.

use vla_lang::machine::load::read_rows;
use vla_lang::reader::read_forms;

use super::inflate::inflate;
use super::ooxml::workbook_bytes;
use super::zip;

const CASES: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../scripts/appearance.txt");
const WRITTEN: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/appearance_written.xlsx"
);
const SAVED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/appearance_saved.xlsx"
);

/// The fixture's cases: each header after its `=== ` and its one line.
fn cases() -> Vec<(String, String)> {
    let mut out: Vec<(String, String)> = Vec::new();
    let text = std::fs::read_to_string(CASES)
        .expect("scripts/appearance.txt")
        .replace("\r\n", "\n");
    let mut lines = text.lines();
    while let Some(line) = lines.next() {
        if let Some(name) = line.strip_prefix("=== ") {
            let body = lines.next().unwrap_or("").to_string();
            out.push((name.to_string(), body));
        }
    }
    out
}

/// The rows of the written workbook: a sheet named Appearance, its two
/// columns wide enough to read, and a case every other row under a header
/// row (case k on row 2k), a case's format at the case's own index. A case
/// the language refuses has no format, and its cell says so. The writer
/// writes no row's height (a sheet's `row` rows are the record's,
/// `KERNEL.22`), so Excel sizes each row to its tallest text.
fn rows() -> String {
    let mut s = String::from(concat!(
        "(sheet \"Appearance\" visible)\n",
        "(column \"Appearance\" \"A\" 44 shown none)\n",
        "(column \"Appearance\" \"B\" 30 shown none)\n",
        "(cell \"Appearance\" \"A1\" \"case\")\n",
        "(cell \"Appearance\" \"B1\" \"drawn\")\n",
    ));
    for (i, (name, body)) in cases().iter().enumerate() {
        let (k, row) = (i + 1, 2 * (i + 1));
        let quoted = vla_lang::rows::quoted(name);
        s.push_str(&format!("(cell \"Appearance\" \"A{row}\" {quoted})\n"));
        if name.contains(" !refused ") {
            s.push_str(&format!("(cell \"Appearance\" \"B{row}\" \"refused\")\n"));
            continue;
        }
        s.push_str(&format!("(format {k} {body})\n"));
        s.push_str(&format!("(cell \"Appearance\" \"B{row}\" \"Aa 123\")\n"));
        s.push_str(&format!("(style \"Appearance\" \"B{row}\" {k})\n"));
    }
    s
}

/// The cases' workbook as the writer writes it today.
fn written() -> (super::Workbook, Vec<u8>) {
    let forms = read_forms(&rows()).expect("the rows read");
    let draft = read_rows(&forms).unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    let book = draft.workbook();
    let bytes = workbook_bytes(&book).expect("the workbook's bytes");
    (book, bytes)
}

/// A part of a workbook as text, stored as the writer stores it or
/// deflated as Excel saves it.
fn part(bytes: &[u8], name: &str) -> String {
    let entries = zip::entries(bytes).expect("an archive");
    let e = entries
        .iter()
        .find(|e| e.name == name)
        .unwrap_or_else(|| panic!("no part {name}"));
    let raw = zip::raw_data(bytes, e).expect("the part's bytes");
    let data = match e.method {
        0 => raw.to_vec(),
        8 => inflate(raw, e.size as usize, 1 << 24).expect("the part inflates"),
        m => panic!("{name}: compression method {m}"),
    };
    String::from_utf8(data).expect("the part is UTF-8")
}

/// An attribute's value in an element's opening tag.
fn attr<'t>(element: &'t str, name: &str) -> Option<&'t str> {
    let open = &element[..element.find('>').unwrap_or(element.len())];
    let key = format!(" {name}=\"");
    let at = open.find(&key)? + key.len();
    open[at..].split('"').next()
}

/// The `<element>`s of a `<list>`, each whole, in their order: the first
/// list of that name, its children read up to its close.
fn elements<'t>(text: &'t str, list: &str, element: &str) -> Vec<&'t str> {
    let Some(start) = text.find(&format!("<{list}")) else {
        return Vec::new();
    };
    let body_at = start + text[start..].find('>').expect("the list's tag closes") + 1;
    let end = body_at + text[body_at..].find(&format!("</{list}>")).unwrap_or(0);
    let body = &text[body_at..end];
    let mut out = Vec::new();
    let mut at = 0;
    let open = format!("<{element}");
    while let Some(i) = body[at..].find(&open).map(|i| at + i) {
        let next = body[i + open.len()..].chars().next();
        if !matches!(next, Some(' ' | '>' | '/')) {
            at = i + open.len();
            continue;
        }
        let tag_end = i + body[i..].find('>').expect("the tag closes") + 1;
        let whole_end = if body[..tag_end].ends_with("/>") {
            tag_end
        } else {
            let close = format!("</{element}>");
            tag_end + body[tag_end..].find(&close).expect("the element closes") + close.len()
        };
        out.push(&body[i..whole_end]);
        at = whole_end;
    }
    out
}

/// The first child `<name ...>` of an element, whole, if any.
fn child<'t>(element: &'t str, name: &str) -> Option<&'t str> {
    let open = format!("<{name}");
    let mut at = 1;
    while let Some(i) = element[at..].find(&open).map(|i| at + i) {
        let next = element[i + open.len()..].chars().next();
        if matches!(next, Some(' ' | '>' | '/')) {
            let tag_end = i + element[i..].find('>')? + 1;
            if element[..tag_end].ends_with("/>") {
                return Some(&element[i..tag_end]);
            }
            let close = format!("</{name}>");
            let end = tag_end + element[tag_end..].find(&close)? + close.len();
            return Some(&element[i..end]);
        }
        at = i + open.len();
    }
    None
}

/// A color's own RGB, `RRGGBB`: an `rgb` color's, its alpha dropped; none
/// for Excel's automatic, a theme's or a palette's, which no case says.
fn own_rgb(color: Option<&str>) -> Option<String> {
    let rgb = attr(color?, "rgb")?;
    Some(rgb[rgb.len().saturating_sub(6)..].to_string())
}

/// What a cell draws with, as a case says it.
#[derive(Debug, PartialEq, Eq)]
struct Drawn {
    bold: bool,
    italic: bool,
    twentieths: u32,
    ink: Option<String>,
    fill: Option<String>,
    edges: Vec<(String, String, Option<String>)>,
    horizontal: String,
    vertical: String,
    wrap: bool,
}

/// A workbook's sheet and styles part, read as far as the cells' looks.
struct Book {
    sheet: String,
    xfs: Vec<String>,
    fonts: Vec<String>,
    fills: Vec<String>,
    borders: Vec<String>,
}

impl Book {
    fn read(bytes: &[u8]) -> Book {
        let entries = zip::entries(bytes).expect("an archive");
        let sheet_part = entries
            .iter()
            .map(|e| e.name.as_str())
            .find(|n| n.starts_with("xl/worksheets/") && n.ends_with(".xml"))
            .expect("a sheet part")
            .to_string();
        let styles = part(bytes, "xl/styles.xml");
        let own = |list: &str, element: &str| -> Vec<String> {
            elements(&styles, list, element)
                .into_iter()
                .map(str::to_string)
                .collect()
        };
        Book {
            sheet: part(bytes, &sheet_part),
            xfs: own("cellXfs", "xf"),
            fonts: own("fonts", "font"),
            fills: own("fills", "fill"),
            borders: own("borders", "border"),
        }
    }

    /// What the cell at `address` draws with.
    fn drawn(&self, address: &str) -> Drawn {
        let key = format!("<c r=\"{address}\"");
        let at = self
            .sheet
            .find(&key)
            .unwrap_or_else(|| panic!("no cell {address}"));
        let cell = &self.sheet[at..at + self.sheet[at..].find('>').expect("the cell's tag")];
        let index =
            |e: &str, a: &str| -> usize { attr(e, a).map_or(0, |v| v.parse().expect("an index")) };
        let xf = &self.xfs[index(cell, "s")];
        let font = &self.fonts[index(xf, "fontId")];
        let flag = |name: &str| {
            child(font, name).is_some_and(|t| !matches!(attr(t, "val"), Some("0" | "false")))
        };
        let size: f64 = attr(child(font, "sz").expect("a size"), "val")
            .expect("its value")
            .parse()
            .expect("a number");
        let fill = &self.fills[index(xf, "fillId")];
        let solid =
            child(fill, "patternFill").and_then(|p| attr(p, "patternType")) == Some("solid");
        let border = &self.borders[index(xf, "borderId")];
        let edges = ["left", "right", "top", "bottom"]
            .iter()
            .filter_map(|side| {
                let edge = child(border, side)?;
                let line = attr(edge, "style")?;
                Some((
                    side.to_string(),
                    line.to_string(),
                    own_rgb(child(edge, "color")),
                ))
            })
            .collect();
        let alignment = child(xf, "alignment").unwrap_or("<alignment/>");
        Drawn {
            bold: flag("b"),
            italic: flag("i"),
            twentieths: (size * 20.0).round() as u32,
            ink: own_rgb(child(font, "color")),
            fill: if solid {
                own_rgb(child(fill, "fgColor"))
            } else {
                None
            },
            edges,
            horizontal: attr(alignment, "horizontal")
                .unwrap_or("general")
                .to_string(),
            vertical: attr(alignment, "vertical").unwrap_or("bottom").to_string(),
            wrap: matches!(attr(alignment, "wrapText"), Some("1" | "true")),
        }
    }
}

#[test]
fn the_written_workbook_is_the_fixture_s_cases_as_the_writer_writes_them() {
    let all = cases();
    assert!(all.len() >= 46, "only {} cases", all.len());
    let (wb, bytes) = written();
    // Every case the language takes is a format of its own, and some say
    // each of a font, an alignment and a border.
    let last = all
        .iter()
        .rposition(|(n, _)| !n.contains(" !refused "))
        .expect("a case taken")
        + 1;
    assert_eq!(wb.styles.xfs().len(), last + 1);
    assert!(wb.styles.xfs().iter().any(|s| s.font.bold == Some(true)));
    assert!(wb.styles.xfs().iter().any(|s| s.align.vertical.is_some()));
    assert!(wb.styles.xfs().iter().any(|s| s.border.left.is_some()));
    if std::env::var("FRAZARO_BLESS").as_deref() == Ok("1") {
        std::fs::write(WRITTEN, &bytes).expect("the written workbook");
        return;
    }
    let held =
        std::fs::read(WRITTEN).expect("scripts/appearance_written.xlsx: FRAZARO_BLESS=1 writes it");
    assert!(
        held == bytes,
        "scripts/appearance_written.xlsx is not what the writer writes of the cases today: FRAZARO_BLESS=1 writes it again"
    );
}

#[test]
fn excel_saves_each_case_with_the_font_fill_border_and_alignment_the_writer_writes() {
    let saved = std::fs::read(SAVED).expect(
        "scripts/appearance_saved.xlsx: VlaWriteAppearanceGolden (VLA_Tests.bas) saves it in Excel",
    );
    let excel = Book::read(&saved);
    let ours = Book::read(&written().1);
    let all = cases();
    for (i, (name, _)) in all.iter().enumerate() {
        let address = format!("B{}", 2 * (i + 1));
        assert_eq!(
            ours.drawn(&address),
            excel.drawn(&address),
            "{name}: the writer's cell {address} beside Excel's"
        );
    }
    // The cases reach what they say in Excel's own save: a weight, a size
    // of hundredths, an ink, a fill, a colored edge and a vertical center.
    let drawn: Vec<Drawn> = (1..=all.len())
        .map(|k| excel.drawn(&format!("B{}", 2 * k)))
        .collect();
    assert!(drawn.iter().any(|d| d.bold));
    assert!(drawn.iter().any(|d| d.twentieths == 8_191));
    assert!(drawn.iter().any(|d| d.ink.as_deref() == Some("800020")));
    assert!(drawn.iter().any(|d| d.fill.as_deref() == Some("6B4FC8")));
    assert!(drawn
        .iter()
        .any(|d| d.edges.iter().any(|e| e.2.as_deref() == Some("00B050"))));
    assert!(drawn.iter().any(|d| d.vertical == "center"));
}

/// The reader's half (KERNEL.57): `reflect --looks` of Excel's own save of
/// the cases, and of the core's written workbook, reads each case's cell as
/// Excel's object model read it, the golden's `EXCEL` and `WRITTEN` lines,
/// for every case the language takes: the scoping's prototype, 44 of 44 in
/// each, held.
#[test]
fn the_looks_read_each_case_as_excel_read_it() {
    const GOLDEN: &str = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../scripts/appearance_golden.txt"
    );
    let golden = std::fs::read_to_string(GOLDEN)
        .expect("scripts/appearance_golden.txt")
        .replace("\r\n", "\n");
    // Each case's two readings, in the fixture's order.
    let mut readings: Vec<(String, String, String)> = Vec::new();
    for line in golden.lines() {
        if let Some(name) = line.strip_prefix("=== ") {
            readings.push((name.to_string(), String::new(), String::new()));
        } else if let Some((which, reading)) = line.split_once('\t') {
            let Some(last) = readings.last_mut() else {
                continue;
            };
            match which {
                "EXCEL" => last.1 = reading.to_string(),
                "WRITTEN" => last.2 = reading.to_string(),
                _ => {}
            }
        }
    }
    assert_eq!(readings.len(), cases().len(), "a reading a case");
    for (path, excel) in [(SAVED, true), (WRITTEN, false)] {
        let bytes = std::fs::read(path).unwrap();
        let text = crate::reflect::reflect_looks_text(&bytes, path).unwrap();
        let mut formats: std::collections::HashMap<u32, String> = Default::default();
        let mut styles: std::collections::HashMap<String, u32> = Default::default();
        for line in text.lines() {
            if let Some(rest) = line.strip_prefix("(format ") {
                let (index, rest) = rest.split_once(' ').unwrap();
                let rest = rest.strip_suffix(')').unwrap();
                formats.insert(index.parse().unwrap(), rest.to_string());
            } else if let Some(rest) = line.strip_prefix("(style \"Appearance\" \"") {
                let (address, index) = rest.split_once("\" ").unwrap();
                let index = index.strip_suffix(')').unwrap();
                styles.insert(address.to_string(), index.parse().unwrap());
            }
        }
        let mut read = 0;
        for (k, (name, by_excel, by_writer)) in readings.iter().enumerate() {
            let want = if excel { by_excel } else { by_writer };
            if want.starts_with("REFUSED") {
                continue;
            }
            let address = format!("B{}", 2 * (k + 1));
            let index = styles.get(&address).copied().unwrap_or(0);
            assert_eq!(
                formats.get(&index),
                Some(want),
                "{path}: {name}, the cell {address} at format {index}"
            );
            read += 1;
        }
        assert_eq!(read, 44, "{path}: the cases read");
    }
}
