//! `frazaro diff <old.xlsx> <new.xlsx>` (PORT.8, slice 8c): two workbooks'
//! files read through the one reader and compared, printed as the `changed`
//! relation of `REFLECT` (docs/SINGULARITY.md, Stage 0.3; the roadmap's
//! `AXM.11`, whose `changed(addr, old, new)` from a saved copy or another
//! file this is, with no host on the machine).
//!
//! Sheets are matched by name without case, as Excel names them; a sheet in
//! one file alone is reported and never matched, so a renamed sheet is a
//! removal and an addition, as `AXM.11` says, and its cells are not listed.
//! Names and Tables, which share one namespace in a workbook, are matched by
//! name without case; the cells of matched sheets by address. Each other
//! difference is one row, `(changed "<key>" <old> <new>)`, the key in the
//! spelling the `refers` relation uses for its second field (`Model!B3`;
//! `Rate` or `Model!Local` for a name as its `name` row spells it; `Sales`
//! for a Table), so that `cause` (`AXM.11`'s three rules) joins `changed` to
//! `refers` with no parser: a repointed name is a changed precedent of every
//! formula that refers to it. A side is `blank` where the file holds nothing,
//! a value as a `cell` row prints it, `(formula "=…")` for a formula with no
//! cached value, or `(formula "=…" <value>)` with one; a name's side is its
//! refers-to text, a Table's its sheet and range as `Data!A1:B5`. A cell is
//! changed when its formula's text or its value differs; numbers compare as
//! numbers (`800` and `800.0` are one value, since a writer other than Excel
//! spells a float with its point), every other value by kind and text.
//!
//! The fixed order, which makes the text a golden: the sheets in one file
//! alone, removed (the old file's tab order) then added (the new file's),
//! then the matched sheets whose state changed, in the new file's tab order;
//! then the names and Tables sorted by key without case; then sheet by sheet
//! in the new file's tab order, cell by cell in document order. The reader
//! streams one pair of sheets at a time: both sheets' cells are held while
//! they are compared and dropped before the next pair, so what is held is
//! two sheets, never two workbooks.

use std::collections::BTreeMap;

use super::ooxml::Package;
use super::print::datum;
use super::{sheet_prefix, Row, Sink, Value, Visibility};
use crate::form::Form;
use crate::intrinsics::fold;
use crate::messages::Refusal;
use crate::printer::write_datum;
use crate::sheet::{cell_ref, parse_a1_range};

fn quoted(s: &str) -> String {
    write_datum(&Form::string(s))
}

/// One row of the difference between two workbooks.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Change<'r> {
    /// A sheet of the old file with no sheet of that name in the new.
    SheetRemoved { name: &'r str, state: Visibility },
    /// A sheet of the new file with no sheet of that name in the old.
    SheetAdded { name: &'r str, state: Visibility },
    /// A matched sheet whose state differs.
    SheetChanged {
        name: &'r str,
        old: Visibility,
        new: Visibility,
    },
    /// `(changed "<key>" <old> <new>)`: a cell, a name or a Table, each side
    /// already spelled as its datum.
    Changed {
        key: &'r str,
        old: &'r str,
        new: &'r str,
    },
}

/// Where the rows go as the comparison finds them.
pub trait ChangeSink {
    fn change(&mut self, change: &Change<'_>);
}

/// A sink that keeps nothing, for `--counts`.
pub struct Discard;

impl ChangeSink for Discard {
    fn change(&mut self, _change: &Change<'_>) {}
}

/// A sink that holds the lines as one text, each ending in a line feed.
#[derive(Default)]
pub struct Lines {
    pub out: String,
}

impl ChangeSink for Lines {
    fn change(&mut self, change: &Change<'_>) {
        self.out.push_str(&line(change));
        self.out.push('\n');
    }
}

/// One row as its line, without the line break.
pub fn line(change: &Change<'_>) -> String {
    match change {
        Change::SheetRemoved { name, state } => {
            format!("(sheet-removed {} {})", quoted(name), state.word())
        }
        Change::SheetAdded { name, state } => {
            format!("(sheet-added {} {})", quoted(name), state.word())
        }
        Change::SheetChanged { name, old, new } => format!(
            "(sheet-changed {} {} {})",
            quoted(name),
            old.word(),
            new.word()
        ),
        Change::Changed { key, old, new } => format!("(changed {} {} {})", quoted(key), old, new),
    }
}

/// What the comparison counted.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct DiffStats {
    pub sheets_removed: u64,
    pub sheets_added: u64,
    pub sheets_changed: u64,
    /// Names whose refers-to differs, or in one file alone.
    pub names: u64,
    /// Tables whose sheet or range differs, or in one file alone.
    pub tables: u64,
    /// Cells of matched sheets whose formula or value differs.
    pub cells: u64,
    /// Addresses compared: the cells either file holds on the matched sheets.
    pub compared: u64,
}

impl DiffStats {
    /// The counts a door prints with `--counts`, before its times.
    pub fn line(&self) -> String {
        format!(
            "sheets-removed {} sheets-added {} sheets-changed {} names {} tables {} cells {} compared {}",
            self.sheets_removed,
            self.sheets_added,
            self.sheets_changed,
            self.names,
            self.tables,
            self.cells,
            self.compared
        )
    }

    /// Whether any row was found.
    pub fn differ(&self) -> bool {
        self.sheets_removed
            + self.sheets_added
            + self.sheets_changed
            + self.names
            + self.tables
            + self.cells
            > 0
    }
}

/// A cell as one file holds it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct CellState {
    value: Option<Value>,
    formula: Option<String>,
}

/// One sheet's cells by (row, column): the sink its walk fills.
#[derive(Default)]
struct SheetCells {
    cells: BTreeMap<(u32, u32), CellState>,
}

impl Sink for SheetCells {
    fn row(&mut self, row: &Row<'_>) {
        match row {
            Row::Cell { addr, value, .. } => {
                if let Some(a) = parse_a1_range(addr) {
                    self.cells.entry((a.top, a.left)).or_default().value = Some((*value).clone());
                }
            }
            Row::Formula { addr, text, .. } => {
                if let Some(a) = parse_a1_range(addr) {
                    self.cells.entry((a.top, a.left)).or_default().formula =
                        Some((*text).to_string());
                }
            }
            _ => {}
        }
    }
}

/// Two numbers are one value when both parse and are equal; every other
/// value compares by kind and text.
fn same_value(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => match (x.parse::<f64>(), y.parse::<f64>()) {
            (Ok(p), Ok(q)) => p == q,
            _ => x == y,
        },
        _ => a == b,
    }
}

fn same_cell(a: Option<&CellState>, b: Option<&CellState>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => {
            a.formula == b.formula
                && match (&a.value, &b.value) {
                    (None, None) => true,
                    (Some(x), Some(y)) => same_value(x, y),
                    _ => false,
                }
        }
        _ => false,
    }
}

/// A cell's side as its datum: `blank`, the value, `(formula "=…")` or
/// `(formula "=…" <value>)`.
fn side(state: Option<&CellState>) -> String {
    match state {
        Some(CellState {
            value,
            formula: Some(text),
        }) => match value {
            Some(v) => format!("(formula {} {})", quoted(text), datum(v)),
            None => format!("(formula {})", quoted(text)),
        },
        Some(CellState {
            value: Some(v),
            formula: None,
        }) => datum(v),
        _ => "blank".to_string(),
    }
}

/// A name's or a Table's side: its text quoted, or `blank`.
fn text_side(text: Option<&str>) -> String {
    match text {
        Some(t) => quoted(t),
        None => "blank".to_string(),
    }
}

/// What two files say about one name or Table.
#[derive(Default)]
struct Sides {
    /// The key as the new file spells it, or the old file when the new has none.
    key: String,
    old: Option<String>,
    new: Option<String>,
    table: bool,
}

/// The difference between two opened packages, to the sink in the fixed
/// order; what was counted comes back. A refusal from either file's walk
/// ends the comparison, the rows before it standing.
pub fn diff(
    old: &Package<'_>,
    new: &Package<'_>,
    sink: &mut dyn ChangeSink,
) -> Result<DiffStats, Refusal> {
    let mut stats = DiffStats::default();
    let old_sheets = old.sheets();
    let new_sheets = new.sheets();
    // Each new sheet's match in the old file, by name without case.
    let matched: Vec<Option<usize>> = new_sheets
        .iter()
        .map(|s| {
            let wanted = fold(&s.name);
            old_sheets.iter().position(|o| fold(&o.name) == wanted)
        })
        .collect();
    for (i, s) in old_sheets.iter().enumerate() {
        if !matched.contains(&Some(i)) {
            stats.sheets_removed += 1;
            sink.change(&Change::SheetRemoved {
                name: &s.name,
                state: s.visibility,
            });
        }
    }
    for (j, s) in new_sheets.iter().enumerate() {
        if matched[j].is_none() {
            stats.sheets_added += 1;
            sink.change(&Change::SheetAdded {
                name: &s.name,
                state: s.visibility,
            });
        }
    }
    for (j, s) in new_sheets.iter().enumerate() {
        if let Some(i) = matched[j] {
            if old_sheets[i].visibility != s.visibility {
                stats.sheets_changed += 1;
                sink.change(&Change::SheetChanged {
                    name: &s.name,
                    old: old_sheets[i].visibility,
                    new: s.visibility,
                });
            }
        }
    }
    // Names and Tables: one namespace, keyed by name without case; the old
    // file first, so that the new file's spelling of a key wins.
    let mut keyed: BTreeMap<String, Sides> = BTreeMap::new();
    for (package, is_new) in [(old, false), (new, true)] {
        for n in package.names().iter().filter(|n| !n.placeholder) {
            let key = package.printed_name(n);
            let e = keyed.entry(fold(&key)).or_default();
            e.key = key;
            if is_new {
                e.new = Some(n.refers_to.clone());
            } else {
                e.old = Some(n.refers_to.clone());
            }
        }
        for t in package.tables() {
            let place = format!(
                "{}!{}",
                sheet_prefix(&package.sheets()[t.sheet].name),
                t.range
            );
            let e = keyed.entry(fold(&t.name)).or_default();
            e.key = t.name.clone();
            e.table = true;
            if is_new {
                e.new = Some(place);
            } else {
                e.old = Some(place);
            }
        }
    }
    for sides in keyed.values() {
        if sides.old != sides.new {
            if sides.table {
                stats.tables += 1;
            } else {
                stats.names += 1;
            }
            sink.change(&Change::Changed {
                key: &sides.key,
                old: &text_side(sides.old.as_deref()),
                new: &text_side(sides.new.as_deref()),
            });
        }
    }
    // The cells, one matched pair of sheets at a time.
    for (j, s) in new_sheets.iter().enumerate() {
        let Some(i) = matched[j] else {
            continue;
        };
        let mut before = SheetCells::default();
        old.walk_sheet(i, &mut before)?;
        let mut after = SheetCells::default();
        new.walk_sheet(j, &mut after)?;
        let prefix = sheet_prefix(&s.name);
        let mut at: Vec<(u32, u32)> = before
            .cells
            .keys()
            .chain(after.cells.keys())
            .copied()
            .collect();
        at.sort_unstable();
        at.dedup();
        stats.compared += at.len() as u64;
        for (row, col) in at {
            let a = before.cells.get(&(row, col));
            let b = after.cells.get(&(row, col));
            if same_cell(a, b) {
                continue;
            }
            stats.cells += 1;
            let key = format!("{prefix}!{}", cell_ref(row, col));
            sink.change(&Change::Changed {
                key: &key,
                old: &side(a),
                new: &side(b),
            });
        }
    }
    Ok(stats)
}

/// The whole difference as text, one row a line in the fixed order: the
/// API's surface, the tests' and a door's that holds the text. Each label is
/// what a refusal calls its file.
pub fn diff_text(
    old_bytes: &[u8],
    old_label: &str,
    new_bytes: &[u8],
    new_label: &str,
) -> Result<String, Refusal> {
    let old = Package::open(old_bytes, old_label)?;
    let new = Package::open(new_bytes, new_label)?;
    let mut lines = Lines::default();
    diff(&old, &new, &mut lines)?;
    Ok(lines.out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::zip;

    const NS: &str = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// A package of three sheets, two names and one Table on the first
    /// sheet, whose cells and the Table's range the caller gives.
    fn package_bytes(sheets: &str, names: &str, model_rows: &str, table_ref: &str) -> Vec<u8> {
        let parts: Vec<(String, Vec<u8>)> = vec![
            (
                "xl/workbook.xml".to_string(),
                format!("<workbook {NS}><sheets>{sheets}</sheets><definedNames>{names}</definedNames></workbook>").into_bytes(),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_string(),
                format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet3.xml\"/></Relationships>").into_bytes(),
            ),
            (
                "xl/worksheets/sheet1.xml".to_string(),
                format!("<worksheet {NS}><sheetData><row r=\"1\">{model_rows}</row></sheetData><tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts></worksheet>").into_bytes(),
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
                format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/table\" Target=\"../tables/table1.xml\"/></Relationships>").into_bytes(),
            ),
            (
                "xl/worksheets/sheet2.xml".to_string(),
                format!("<worksheet {NS}><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>").into_bytes(),
            ),
            (
                "xl/worksheets/sheet3.xml".to_string(),
                format!("<worksheet {NS}><sheetData><row r=\"1\"><c r=\"A1\"><v>2</v></c></row></sheetData></worksheet>").into_bytes(),
            ),
            (
                "xl/tables/table1.xml".to_string(),
                format!("<table {NS} id=\"1\" name=\"T\" displayName=\"T\" ref=\"{table_ref}\"><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Amount\"/></tableColumns></table>").into_bytes(),
            ),
        ];
        zip::write_stored(&parts).unwrap()
    }

    fn old_bytes() -> Vec<u8> {
        package_bytes(
            "<sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Gone\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/><sheet name=\"Shy\" sheetId=\"3\" r:id=\"rId3\"/>",
            "<definedName name=\"Rate\">Model!$B$2</definedName><definedName name=\"Old\">1</definedName>",
            concat!(
                "<c r=\"A1\"><v>1200</v></c><c r=\"B1\"><v>800</v></c>",
                "<c r=\"C1\"><f>A1-B1</f><v>400</v></c><c r=\"D1\"><f>A1*2</f><v>2400</v></c>",
                "<c r=\"E1\" t=\"inlineStr\"><is><t>x</t></is></c><c r=\"F1\" t=\"b\"><v>1</v></c>",
                "<c r=\"G1\"><v>5</v></c>"
            ),
            "A1:B2",
        )
    }

    fn new_bytes() -> Vec<u8> {
        package_bytes(
            "<sheet name=\"model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Shy\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/><sheet name=\"Added\" sheetId=\"3\" r:id=\"rId3\"/>",
            "<definedName name=\"RATE\">Model!$B$3</definedName><definedName name=\"New\">2</definedName>",
            concat!(
                "<c r=\"A1\"><v>1300</v></c><c r=\"B1\"><v>800.0</v></c>",
                "<c r=\"C1\"><v>500</v></c><c r=\"D1\"><f>A1*2</f><v>2600</v></c>",
                "<c r=\"E1\" t=\"inlineStr\"><is><t>y</t></is></c><c r=\"F1\" t=\"b\"><v>0</v></c>",
                "<c r=\"H1\"><v>7</v></c>"
            ),
            "A1:B3",
        )
    }

    #[test]
    fn every_kind_of_change_is_one_row_in_the_fixed_order() {
        let old = old_bytes();
        let new = new_bytes();
        let got = diff_text(&old, "old.xlsx", &new, "new.xlsx").unwrap();
        let want = concat!(
            "(sheet-removed \"Gone\" hidden)\n",
            "(sheet-added \"Added\" visible)\n",
            "(sheet-changed \"Shy\" visible hidden)\n",
            "(changed \"New\" blank \"2\")\n",
            "(changed \"Old\" \"1\" blank)\n",
            "(changed \"RATE\" \"Model!$B$2\" \"Model!$B$3\")\n",
            "(changed \"T\" \"Model!A1:B2\" \"model!A1:B3\")\n",
            "(changed \"model!A1\" 1200 1300)\n",
            "(changed \"model!C1\" (formula \"=A1-B1\" 400) 500)\n",
            "(changed \"model!D1\" (formula \"=A1*2\" 2400) (formula \"=A1*2\" 2600))\n",
            "(changed \"model!E1\" \"x\" \"y\")\n",
            "(changed \"model!F1\" true false)\n",
            "(changed \"model!G1\" 5 blank)\n",
            "(changed \"model!H1\" blank 7)\n",
            // Shy is the third sheet of the old file and the second of the
            // new: matched by name, not by position, and its one cell differs.
            "(changed \"Shy!A1\" 2 1)\n",
        );
        assert_eq!(got, want);
        // B1: 800 and 800.0 are one value, so no row; the counts say so.
        let o = Package::open(&old, "old.xlsx").unwrap();
        let n = Package::open(&new, "new.xlsx").unwrap();
        let stats = diff(&o, &n, &mut Discard).unwrap();
        assert_eq!(
            stats,
            DiffStats {
                sheets_removed: 1,
                sheets_added: 1,
                sheets_changed: 1,
                names: 3,
                tables: 1,
                cells: 8,
                compared: 9,
            }
        );
        assert!(stats.differ());
        assert_eq!(
            stats.line(),
            "sheets-removed 1 sheets-added 1 sheets-changed 1 names 3 tables 1 cells 8 compared 9"
        );
    }

    #[test]
    fn a_file_against_itself_has_no_rows() {
        let old = old_bytes();
        assert_eq!(diff_text(&old, "a.xlsx", &old, "b.xlsx").unwrap(), "");
        let o = Package::open(&old, "a.xlsx").unwrap();
        let stats = diff(&o, &o, &mut Discard).unwrap();
        assert!(!stats.differ());
        assert_eq!(stats.compared, 9);
    }

    #[test]
    fn numbers_compare_as_numbers_and_the_rest_by_kind_and_text() {
        let n = |s: &str| Value::Number(s.to_string());
        assert!(same_value(&n("800"), &n("800.0")));
        assert!(same_value(&n("1E3"), &n("1000")));
        assert!(!same_value(&n("0.1"), &n("0.10000000000000002")));
        assert!(!same_value(&n("5"), &Value::Text("5".to_string())));
        assert!(!same_value(&n("1"), &Value::Bool(true)));
        assert!(same_value(
            &Value::Error("#N/A".to_string()),
            &Value::Error("#N/A".to_string())
        ));
        // A number the file spells in a way no float reads compares as text.
        assert!(!same_value(&n("1,5"), &n("1.5")));
        assert!(same_value(&n("1,5"), &n("1,5")));
    }

    #[test]
    fn each_row_prints_in_the_corpus_notation() {
        assert_eq!(
            line(&Change::SheetRemoved {
                name: "Q1 Data",
                state: Visibility::VeryHidden
            }),
            "(sheet-removed \"Q1 Data\" very-hidden)"
        );
        assert_eq!(
            line(&Change::SheetAdded {
                name: "It's",
                state: Visibility::Visible
            }),
            "(sheet-added \"It's\" visible)"
        );
        assert_eq!(
            line(&Change::SheetChanged {
                name: "Scratch",
                old: Visibility::Hidden,
                new: Visibility::Visible
            }),
            "(sheet-changed \"Scratch\" hidden visible)"
        );
        assert_eq!(
            line(&Change::Changed {
                key: "'Q1 Data'!B2",
                old: "blank",
                new: "(formula \"=1/0\" (error \"#DIV/0!\"))"
            }),
            "(changed \"'Q1 Data'!B2\" blank (formula \"=1/0\" (error \"#DIV/0!\")))"
        );
        let blank = CellState::default();
        assert_eq!(side(None), "blank");
        assert_eq!(side(Some(&blank)), "blank");
        assert_eq!(
            side(Some(&CellState {
                value: Some(Value::Text("say \"hi\"".to_string())),
                formula: None
            })),
            "\"say \\\"hi\\\"\""
        );
        assert_eq!(
            side(Some(&CellState {
                value: None,
                formula: Some("=B2+B3".to_string())
            })),
            "(formula \"=B2+B3\")"
        );
        assert_eq!(text_side(Some("Data!A1:B4")), "\"Data!A1:B4\"");
        assert_eq!(text_side(None), "blank");
    }

    /// The treaty's oracle 9 (conformance/README.md, the amendment of
    /// 2026-10-04 for PORT.8 slice 8c): each pair's difference is its golden
    /// whole, line endings aside. The goldens are the core's own output,
    /// blessed by the owner reading both files in Excel; a deliberate change
    /// regenerates them with the door and raises the check's floors.
    #[test]
    fn the_diff_goldens_are_reproduced() {
        /// The old file's label and bytes, the new file's, and the golden.
        type Pair = (
            &'static str,
            &'static [u8],
            &'static str,
            &'static [u8],
            &'static str,
        );
        let rows: [Pair; 3] = [
            // The first build golden against its Excel-saved copy: the five
            // cached values and the stamp's _xlfn._LONGTEXT rewrite.
            (
                "scripts/build/fixture_golden.xlsx",
                include_bytes!("../../../scripts/build/fixture_golden.xlsx"),
                "scripts/reflect/saved.xlsx",
                include_bytes!("../../../scripts/reflect/saved.xlsx"),
                include_str!("../../../scripts/reflect/build_fixture_saved_diff.vla"),
            ),
            // The model against the into golden built into it: what --into added.
            (
                "scripts/build/model.xlsx",
                include_bytes!("../../../scripts/build/model.xlsx"),
                "scripts/build/into_golden.xlsx",
                include_bytes!("../../../scripts/build/into_golden.xlsx"),
                include_str!("../../../scripts/reflect/model_into_diff.vla"),
            ),
            // The reader's fixture against its changed copy, every arm.
            (
                "scripts/reflect/fixture.xlsx",
                include_bytes!("../../../scripts/reflect/fixture.xlsx"),
                "scripts/reflect/changed.xlsx",
                include_bytes!("../../../scripts/reflect/changed.xlsx"),
                include_str!("../../../scripts/reflect/fixture_changed_diff.vla"),
            ),
        ];
        for (old_label, old, new_label, new, golden) in rows {
            let got = diff_text(old, old_label, new, new_label).unwrap_or_else(|r| {
                panic!("{old_label} against {new_label} did not read: {}", r.text)
            });
            let want = golden.replace("\r\n", "\n");
            if got != want {
                let line = got
                    .lines()
                    .zip(want.lines())
                    .position(|(a, b)| a != b)
                    .map(|i| i + 1)
                    .unwrap_or(got.lines().count().min(want.lines().count()) + 1);
                panic!("{old_label} against {new_label}: the difference differs from the golden at line {line}");
            }
        }
    }
}
