//! `frazaro audit <file.xlsx>` (PORT.8, slice 8d): the audit list of Stage
//! 2.3 (docs/SINGULARITY.md; the roadmap's `AXM.10`) over a workbook's file,
//! with no host and no engine: six of the seven questions internal audit
//! runs by eye, each a named walk over the reader's relations, each defined
//! in words in the treaty (conformance/README.md, oracle 10) before its code.
//!
//! The walks. A *typed-over constant* is a cell holding a value and no
//! formula whose nearest cells above and below in its column that the file
//! holds both hold formulas with one text in R1C1 relative to their own
//! cells. An *inconsistent formula* differs in R1C1 from those two
//! neighbours when they agree. An *unused name* is a defined name that no
//! formula and no other name refers to, by its bare name without case
//! (Excel's own `_xlnm.` names and `_xlfn.` placeholders, and Frazaro's own
//! marks, the build stamp and the add-in's `VLAt_` names, left out). An
//! *empty reference* is a single-cell reference, to a sheet the file has,
//! whose cell has no row. A *hidden sheet* is what its `sheet` row says. An
//! *external link* is a formula's reference into another workbook. A cell at
//! either end of a column is never judged, and ranges are never expanded:
//! the list is conservative by construction. The totals that do not foot
//! wait for a meaning of "total".
//!
//! The fixed order: the six walks in that order; inside a walk sheet by
//! sheet in tab order and cell by cell in document order, a formula's
//! targets sorted as its `refers` rows are, the names by name without case.
//! The audit holds every cell of the workbook, value and formula, in one
//! [`AuditIndex`] column-major by sheet, so that a column's neighbours are
//! two lookups: the second exception to streaming after `--cone`.

use std::collections::{BTreeMap, HashSet};

use super::print::datum;
use super::{sheet_prefix, Row, Sink, Source, Value, Visibility};
use crate::form::Form;
use crate::intrinsics::fold;
use crate::messages::Refusal;
use crate::printer::write_datum;
use crate::refers::{self, Kind, Shape};
use crate::sheet::{cell_ref, parse_a1_range};

fn quoted(s: &str) -> String {
    write_datum(&Form::string(s))
}

/// One finding, borrowed from the walk that made it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Finding<'r> {
    /// `(typed-over "Review!B3" 61 "=A3*2")`: the cell, its value as a
    /// datum, and the neighbours' formula rendered at the cell.
    TypedOver {
        cell: &'r str,
        value: &'r str,
        expected: &'r str,
    },
    /// `(inconsistent "Review!C3" "=A3+2" "=A3+1")`: the cell, its formula
    /// as written, and the neighbours' rendered at the cell.
    Inconsistent {
        cell: &'r str,
        formula: &'r str,
        expected: &'r str,
    },
    /// `(unused-name "Range1" "Data!$A$2:$B$4")`: the name as its `name`
    /// row spells it, and what it refers to.
    UnusedName { name: &'r str, refers_to: &'r str },
    /// `(empty-reference "Review!D1" "Review!Z9")`: the formula's cell and
    /// the target as its `refers` row spells it.
    EmptyReference { cell: &'r str, target: &'r str },
    /// `(hidden-sheet "Scratch" hidden)`.
    HiddenSheet { name: &'r str, state: Visibility },
    /// `(external-link "Model!E2" "[Rates.xlsx]Sheet1!A1")`.
    ExternalLink { cell: &'r str, target: &'r str },
}

/// Where the findings go as the walks make them.
pub trait FindingSink {
    fn finding(&mut self, finding: &Finding<'_>);
}

/// A sink that keeps nothing, for `--counts`.
pub struct Discard;

impl FindingSink for Discard {
    fn finding(&mut self, _finding: &Finding<'_>) {}
}

/// A sink that holds the lines as one text, each ending in a line feed.
#[derive(Default)]
pub struct Lines {
    pub out: String,
}

impl FindingSink for Lines {
    fn finding(&mut self, finding: &Finding<'_>) {
        self.out.push_str(&line(finding));
        self.out.push('\n');
    }
}

/// One finding as its line, without the line break.
pub fn line(finding: &Finding<'_>) -> String {
    match finding {
        Finding::TypedOver {
            cell,
            value,
            expected,
        } => format!(
            "(typed-over {} {} {})",
            quoted(cell),
            value,
            quoted(expected)
        ),
        Finding::Inconsistent {
            cell,
            formula,
            expected,
        } => format!(
            "(inconsistent {} {} {})",
            quoted(cell),
            quoted(formula),
            quoted(expected)
        ),
        Finding::UnusedName { name, refers_to } => {
            format!("(unused-name {} {})", quoted(name), quoted(refers_to))
        }
        Finding::EmptyReference { cell, target } => {
            format!("(empty-reference {} {})", quoted(cell), quoted(target))
        }
        Finding::HiddenSheet { name, state } => {
            format!("(hidden-sheet {} {})", quoted(name), state.word())
        }
        Finding::ExternalLink { cell, target } => {
            format!("(external-link {} {})", quoted(cell), quoted(target))
        }
    }
}

/// What the walks counted.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct AuditStats {
    pub typed_over: u64,
    pub inconsistent: u64,
    pub unused_names: u64,
    pub empty_references: u64,
    pub hidden_sheets: u64,
    pub external_links: u64,
}

impl AuditStats {
    /// The counts a door prints with `--counts`, before its times.
    pub fn line(&self) -> String {
        format!(
            "typed-over {} inconsistent {} unused-names {} empty-references {} hidden-sheets {} external-links {}",
            self.typed_over,
            self.inconsistent,
            self.unused_names,
            self.empty_references,
            self.hidden_sheets,
            self.external_links
        )
    }

    /// Whether any finding was made.
    pub fn any(&self) -> bool {
        self.typed_over
            + self.inconsistent
            + self.unused_names
            + self.empty_references
            + self.hidden_sheets
            + self.external_links
            > 0
    }
}

/// A cell as the index holds it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct Held {
    formula: Option<String>,
    value: Option<Value>,
}

/// Every cell of the workbook, sheet by sheet, column-major: keyed
/// (column, row), so that a column's neighbours are two range lookups.
#[derive(Debug, Default)]
pub struct AuditIndex {
    /// Each sheet's name, folded, in tab order.
    sheets: Vec<String>,
    cells: Vec<BTreeMap<(u32, u32), Held>>,
}

impl AuditIndex {
    /// An index for the sheets of a workbook, in tab order, filled by
    /// walking each sheet into it.
    pub fn new(package: &dyn Source) -> AuditIndex {
        let n = package.sheets().len();
        AuditIndex {
            sheets: package.sheets().iter().map(|s| fold(&s.name)).collect(),
            cells: (0..n).map(|_| BTreeMap::new()).collect(),
        }
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        let f = fold(name);
        self.sheets.iter().position(|s| *s == f)
    }

    fn holds(&self, sheet: usize, row: u32, col: u32) -> bool {
        self.cells[sheet].contains_key(&(col, row))
    }
}

impl Sink for AuditIndex {
    fn row(&mut self, row: &Row<'_>) {
        match row {
            Row::Cell {
                sheet, addr, value, ..
            } => {
                if let (Some(i), Some(a)) = (self.sheet_index(sheet), parse_a1_range(addr)) {
                    self.cells[i].entry((a.left, a.top)).or_default().value =
                        Some((*value).clone());
                }
            }
            Row::Formula {
                sheet, addr, text, ..
            } => {
                if let (Some(i), Some(a)) = (self.sheet_index(sheet), parse_a1_range(addr)) {
                    self.cells[i].entry((a.left, a.top)).or_default().formula =
                        Some((*text).to_string());
                }
            }
            _ => {}
        }
    }
}

/// The nearest cell the file holds above (`row`, `col`) and the nearest
/// below, when both hold formulas with one R1C1 text: the row of the one
/// above and its formula's text.
fn agreeing_neighbours(
    cells: &BTreeMap<(u32, u32), Held>,
    row: u32,
    col: u32,
) -> Option<(u32, &str)> {
    let (&(_, above_row), above) = cells.range((col, 0)..(col, row)).next_back()?;
    let (&(_, below_row), below) = cells.range((col, row + 1)..(col + 1, 0)).next()?;
    let (a, b) = (above.formula.as_deref()?, below.formula.as_deref()?);
    if refers::r1c1(a, above_row, col) != refers::r1c1(b, below_row, col) {
        return None;
    }
    Some((above_row, a))
}

/// Excel's own names and Frazaro's own marks, which no formula refers to by
/// design and the unused-name walk never reports: `_xlnm.` (print areas,
/// filter databases), the build stamp (`build::STAMP_NAME`, which `rebuild`
/// reads, and any other name under `Frazaro.`) and the add-in's `VLAt_`
/// names. `name` is folded.
fn is_mark(name: &str) -> bool {
    name.starts_with("_xlnm.") || name.starts_with("frazaro.") || name.starts_with("vlat_")
}

/// Every name a formula's text (or a name's refers-to text) refers to, by
/// its bare name folded, into the set.
fn note_names(used: &mut HashSet<String>, text: &str) {
    for r in refers::scan(text) {
        match r.kind {
            Kind::Name => {
                used.insert(fold(&r.part));
            }
            Kind::Spill if r.shape != Shape::Cell => {
                used.insert(fold(r.part.trim_end_matches('#')));
            }
            _ => {}
        }
    }
}

/// The six walks over an index the workbook's sheets were walked into, to
/// the sink in the fixed order; what was counted comes back.
pub fn audit(index: &AuditIndex, package: &dyn Source, sink: &mut dyn FindingSink) -> AuditStats {
    let mut stats = AuditStats::default();
    let sheets = package.sheets();
    // The column walks: typed-over constants and inconsistent formulas, one
    // pass over every sheet in document order.
    let mut typed: Vec<(String, String, String)> = Vec::new();
    let mut inconsistent: Vec<(String, String, String)> = Vec::new();
    for (s, info) in sheets.iter().enumerate() {
        let prefix = sheet_prefix(&info.name);
        let cells = &index.cells[s];
        let mut at: Vec<(u32, u32)> = cells.keys().map(|&(c, r)| (r, c)).collect();
        at.sort_unstable();
        for (row, col) in at {
            let held = &cells[&(col, row)];
            let Some((above_row, above)) = agreeing_neighbours(cells, row, col) else {
                continue;
            };
            let expected = refers::shift_a1_references(above, i64::from(row - above_row), 0);
            let cell = format!("{prefix}!{}", cell_ref(row, col));
            match (&held.formula, &held.value) {
                (Some(f), _) => {
                    if refers::r1c1(f, row, col) != refers::r1c1(above, above_row, col) {
                        inconsistent.push((cell, f.clone(), expected));
                    }
                }
                (None, Some(v)) => typed.push((cell, datum(v), expected)),
                (None, None) => {}
            }
        }
    }
    // Unused names: what every formula and every name refers to, by name.
    let mut used: HashSet<String> = HashSet::new();
    for cells in &index.cells {
        for held in cells.values() {
            if let Some(f) = &held.formula {
                note_names(&mut used, f);
            }
        }
    }
    for n in package.names().iter().filter(|n| !n.placeholder) {
        note_names(&mut used, &n.refers_to);
    }
    let mut unused: Vec<(String, String)> = package
        .names()
        .iter()
        .filter(|n| {
            let f = fold(&n.name);
            !n.placeholder && !is_mark(&f) && !used.contains(&f)
        })
        .map(|n| (package.printed_name(n), n.refers_to.clone()))
        .collect();
    unused.sort_by(|a, b| fold(&a.0).cmp(&fold(&b.0)).then_with(|| a.0.cmp(&b.0)));
    // Empty references and external links: every formula in document order.
    let mut empty: Vec<(String, String)> = Vec::new();
    let mut links: Vec<(String, String)> = Vec::new();
    for (s, info) in sheets.iter().enumerate() {
        let prefix = sheet_prefix(&info.name);
        let cells = &index.cells[s];
        let mut at: Vec<(u32, u32)> = cells
            .iter()
            .filter(|(_, h)| h.formula.is_some())
            .map(|(&(c, r), _)| (r, c))
            .collect();
        at.sort_unstable();
        for (row, col) in at {
            let Some(text) = cells[&(col, row)].formula.as_deref() else {
                continue;
            };
            let cell = format!("{prefix}!{}", cell_ref(row, col));
            let mut empties: Vec<String> = Vec::new();
            let mut externals: Vec<String> = Vec::new();
            for r in refers::scan(text) {
                match r.kind {
                    Kind::Cell if r.shape == Shape::Cell => {
                        let target = if r.sheet_name.is_empty() {
                            Some(s)
                        } else {
                            index.sheet_index(&r.sheet_name)
                        };
                        if let Some(t) = target {
                            if !index.holds(t, r.row1, r.col1) {
                                empties.push(refers::spell(&r, &info.name));
                            }
                        }
                    }
                    Kind::External => externals.push(refers::spell(&r, &info.name)),
                    _ => {}
                }
            }
            for list in [&mut empties, &mut externals] {
                list.sort_by_key(|t| quoted(t));
                list.dedup();
            }
            for t in empties {
                empty.push((cell.clone(), t));
            }
            for t in externals {
                links.push((cell.clone(), t));
            }
        }
    }
    for (cell, value, expected) in &typed {
        stats.typed_over += 1;
        sink.finding(&Finding::TypedOver {
            cell,
            value,
            expected,
        });
    }
    for (cell, formula, expected) in &inconsistent {
        stats.inconsistent += 1;
        sink.finding(&Finding::Inconsistent {
            cell,
            formula,
            expected,
        });
    }
    for (name, refers_to) in &unused {
        stats.unused_names += 1;
        sink.finding(&Finding::UnusedName { name, refers_to });
    }
    for (cell, target) in &empty {
        stats.empty_references += 1;
        sink.finding(&Finding::EmptyReference { cell, target });
    }
    for info in sheets
        .iter()
        .filter(|i| i.visibility != Visibility::Visible)
    {
        stats.hidden_sheets += 1;
        sink.finding(&Finding::HiddenSheet {
            name: &info.name,
            state: info.visibility,
        });
    }
    for (cell, target) in &links {
        stats.external_links += 1;
        sink.finding(&Finding::ExternalLink { cell, target });
    }
    stats
}

/// The whole audit as text, one finding a line in the fixed order: the
/// API's surface, the tests' and a door's that holds the text. `label` is
/// what a refusal calls the file.
pub fn audit_text(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    let package = super::open(bytes, label)?;
    let mut index = AuditIndex::new(package.as_ref());
    for i in 0..package.sheets().len() {
        package.walk_sheet(i, &mut index)?;
    }
    let mut lines = Lines::default();
    audit(&index, package.as_ref(), &mut lines);
    Ok(lines.out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reflect::ooxml::Package;
    use crate::sheet::zip;

    const NS: &str = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// Two sheets, four names and a link: one case of each walk, and the
    /// cases each walk must leave alone.
    fn package_bytes() -> Vec<u8> {
        let model = concat!(
            // A: the inputs. B: a fill with B3 typed over. C: a fill with C3
            // inconsistent. D: an empty reference, a reference into a held
            // cell of another sheet, one into a sheet the file lacks, and a
            // range, which is never judged. E: a name used by a formula.
            // F: one link referred to twice. H: a fill with blanks between,
            // H4 typed over; the ends of every column are never judged.
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1*2</f><v>2</v></c><c r=\"C1\"><f>A1+1</f><v>2</v></c>",
            "<c r=\"D1\"><f>Z9</f><v>0</v></c><c r=\"E1\"><f>Used</f><v>1</v></c><c r=\"F1\"><f>[1]Sheet1!A1+[1]Sheet1!A1</f><v>2</v></c>",
            "<c r=\"H1\"><f>A1</f><v>1</v></c></row>",
            "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f>A2*2</f><v>4</v></c><c r=\"C2\"><f>A2+1</f><v>3</v></c>",
            "<c r=\"D2\"><f>Other!A1</f><v>1</v></c></row>",
            "<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><v>61</v></c><c r=\"C3\"><f>A3+2</f><v>5</v></c>",
            "<c r=\"D3\"><f>Missing!A1</f><v>0</v></c><c r=\"H3\"><f>A3</f><v>3</v></c></row>",
            "<row r=\"4\"><c r=\"A4\"><v>4</v></c><c r=\"B4\"><f>A4*2</f><v>8</v></c><c r=\"C4\"><f>A4+1</f><v>5</v></c>",
            "<c r=\"D4\"><f>SUM(A1:A2)</f><v>3</v></c><c r=\"H4\"><v>9</v></c></row>",
            "<row r=\"5\"><c r=\"A5\"><v>5</v></c><c r=\"H5\"><f>A5</f><v>5</v></c></row>",
        );
        let parts: Vec<(String, Vec<u8>)> = vec![
            (
                "xl/workbook.xml".to_string(),
                format!(concat!(
                    "<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Other\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/></sheets>",
                    "<externalReferences><externalReference r:id=\"rId3\"/></externalReferences>",
                    "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Model!$A$1:$B$2</definedName>",
                    "<definedName name=\"Base\">Model!$A$2</definedName><definedName name=\"Chain\">Base</definedName>",
                    "<definedName name=\"Frazaro.Build\">\"Frazaro 0.7.1; a stamp\"</definedName>",
                    "<definedName name=\"Used\">Model!$A$1</definedName><definedName name=\"VLAt_TimeRun\">1</definedName></definedNames></workbook>"), NS = NS).into_bytes(),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_string(),
                format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"{REL}/externalLink\" Target=\"externalLinks/externalLink1.xml\"/></Relationships>").into_bytes(),
            ),
            (
                "xl/worksheets/sheet1.xml".to_string(),
                format!("<worksheet {NS}><sheetData>{model}</sheetData></worksheet>").into_bytes(),
            ),
            (
                "xl/worksheets/sheet2.xml".to_string(),
                format!("<worksheet {NS}><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>").into_bytes(),
            ),
            (
                "xl/externalLinks/externalLink1.xml".to_string(),
                format!("<externalLink {NS}><externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Sheet1\"/></sheetNames></externalBook></externalLink>").into_bytes(),
            ),
            (
                "xl/externalLinks/_rels/externalLink1.xml.rels".to_string(),
                format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/externalLinkPath\" Target=\"Rates.xlsx\" TargetMode=\"External\"/></Relationships>").into_bytes(),
            ),
        ];
        zip::write_stored(&parts).unwrap()
    }

    #[test]
    fn each_walk_finds_its_case_and_leaves_the_rest_alone() {
        let bytes = package_bytes();
        let got = audit_text(&bytes, "t.xlsx").unwrap();
        let want = concat!(
            "(typed-over \"Model!B3\" 61 \"=A3*2\")\n",
            "(typed-over \"Model!H4\" 9 \"=A4\")\n",
            "(inconsistent \"Model!C3\" \"=A3+2\" \"=A3+1\")\n",
            "(unused-name \"Chain\" \"Base\")\n",
            "(empty-reference \"Model!D1\" \"Model!Z9\")\n",
            "(hidden-sheet \"Other\" hidden)\n",
            "(external-link \"Model!F1\" \"[Rates.xlsx]Sheet1!A1\")\n",
        );
        assert_eq!(got, want);
        let package = Package::open(&bytes, "t.xlsx").unwrap();
        let mut index = AuditIndex::new(&package);
        for i in 0..package.sheets().len() {
            package.walk_sheet(i, &mut index).unwrap();
        }
        let stats = audit(&index, &package, &mut Discard);
        assert_eq!(
            stats,
            AuditStats {
                typed_over: 2,
                inconsistent: 1,
                unused_names: 1,
                empty_references: 1,
                hidden_sheets: 1,
                external_links: 1,
            }
        );
        assert!(stats.any());
        assert_eq!(
            stats.line(),
            "typed-over 2 inconsistent 1 unused-names 1 empty-references 1 hidden-sheets 1 external-links 1"
        );
    }

    #[test]
    fn the_neighbours_must_both_be_formulas_that_agree() {
        let mut cells: BTreeMap<(u32, u32), Held> = BTreeMap::new();
        let formula = |t: &str| Held {
            formula: Some(t.to_string()),
            value: None,
        };
        let value = |t: &str| Held {
            formula: None,
            value: Some(Value::Number(t.to_string())),
        };
        // Column 2: a formula at row 1, a value at row 3, a formula at row
        // 6 that agrees with row 1 in R1C1; rows 2, 4 and 5 blank.
        cells.insert((2, 1), formula("=A1*2"));
        cells.insert((2, 3), value("7"));
        cells.insert((2, 6), formula("=A6*2"));
        assert_eq!(agreeing_neighbours(&cells, 3, 2), Some((1, "=A1*2")));
        // The ends are never judged.
        assert_eq!(agreeing_neighbours(&cells, 1, 2), None);
        assert_eq!(agreeing_neighbours(&cells, 6, 2), None);
        // A value above, or a formula that does not agree, judges nothing.
        cells.insert((2, 1), value("1"));
        assert_eq!(agreeing_neighbours(&cells, 3, 2), None);
        cells.insert((2, 1), formula("=A1*3"));
        assert_eq!(agreeing_neighbours(&cells, 3, 2), None);
        // An absolute reference agrees with itself.
        cells.insert((2, 1), formula("=$A$1*2"));
        cells.insert((2, 6), formula("=$A$1*2"));
        assert_eq!(agreeing_neighbours(&cells, 3, 2), Some((1, "=$A$1*2")));
    }

    #[test]
    fn a_name_is_used_by_a_formula_or_by_another_name() {
        let mut used = HashSet::new();
        note_names(&mut used, "=Rate*2+Model!Local+Spilled#+A1+Sales[Amount]");
        assert_eq!(used.len(), 3);
        assert!(used.contains("rate") && used.contains("local") && used.contains("spilled"));
        note_names(&mut used, "Deep+Model!$B$2");
        assert!(used.contains("deep"));
        // The marks: Excel's own names and Frazaro's, never reported.
        assert!(is_mark(&fold(crate::build::STAMP_NAME)));
        assert!(is_mark("_xlnm.print_area") && is_mark("vlat_timerun"));
        assert!(!is_mark("rate") && !is_mark("frazaro") && !is_mark("vlat"));
    }

    #[test]
    fn each_finding_prints_in_the_corpus_notation() {
        assert_eq!(
            line(&Finding::TypedOver {
                cell: "'Q1 Data'!B3",
                value: "61",
                expected: "=A3*2"
            }),
            "(typed-over \"'Q1 Data'!B3\" 61 \"=A3*2\")"
        );
        assert_eq!(
            line(&Finding::Inconsistent {
                cell: "Model!C3",
                formula: "=A3+2",
                expected: "=A3+1"
            }),
            "(inconsistent \"Model!C3\" \"=A3+2\" \"=A3+1\")"
        );
        assert_eq!(
            line(&Finding::UnusedName {
                name: "Model!Local",
                refers_to: "Model!$B$2"
            }),
            "(unused-name \"Model!Local\" \"Model!$B$2\")"
        );
        assert_eq!(
            line(&Finding::EmptyReference {
                cell: "Model!D1",
                target: "Model!Z9"
            }),
            "(empty-reference \"Model!D1\" \"Model!Z9\")"
        );
        assert_eq!(
            line(&Finding::HiddenSheet {
                name: "Secret",
                state: Visibility::VeryHidden
            }),
            "(hidden-sheet \"Secret\" very-hidden)"
        );
        assert_eq!(
            line(&Finding::ExternalLink {
                cell: "Model!E2",
                target: "[Rates.xlsx]Sheet1!A1"
            }),
            "(external-link \"Model!E2\" \"[Rates.xlsx]Sheet1!A1\")"
        );
        assert!(!AuditStats::default().any());
        // A file with nothing to report prints nothing: the into golden,
        // whose one name, Rate, its Output sheet uses, and whose stamp is
        // Frazaro's own mark, not a finding.
        let into = include_bytes!("../../../scripts/build/into_golden.xlsx");
        assert_eq!(audit_text(into, "into_golden.xlsx").unwrap(), "");
    }

    /// The treaty's oracle 10 (conformance/README.md, the amendment of
    /// 2026-10-05 for PORT.8 slice 8d): each fixture's findings are its
    /// golden whole, line endings aside. The goldens are the core's own
    /// output, blessed by the owner reading the fixture in Excel; a
    /// deliberate change regenerates them with the door and raises the
    /// check's floors.
    #[test]
    fn the_audit_goldens_are_reproduced() {
        let rows: [(&str, &[u8], &str); 5] = [
            // The twin as Excel saved it (8e): the link's formula gone, so
            // seven findings.
            (
                "scripts/reflect/opendocument_saved.ods",
                include_bytes!("../../../scripts/reflect/opendocument_saved.ods"),
                include_str!("../../../scripts/reflect/opendocument_saved_audit.vla"),
            ),
            // The fixture's OpenDocument twin (8e): the same column cases,
            // one unused name, two hidden sheets, the link.
            (
                "scripts/reflect/opendocument.ods",
                include_bytes!("../../../scripts/reflect/opendocument.ods"),
                include_str!("../../../scripts/reflect/opendocument_audit.vla"),
            ),
            (
                "scripts/reflect/fixture.xlsx",
                include_bytes!("../../../scripts/reflect/fixture.xlsx"),
                include_str!("../../../scripts/reflect/fixture_audit.vla"),
            ),
            (
                "scripts/reflect/changed.xlsx",
                include_bytes!("../../../scripts/reflect/changed.xlsx"),
                include_str!("../../../scripts/reflect/changed_audit.vla"),
            ),
            (
                "scripts/build/fixture_golden.xlsx",
                include_bytes!("../../../scripts/build/fixture_golden.xlsx"),
                include_str!("../../../scripts/reflect/build_fixture_audit.vla"),
            ),
        ];
        for (label, bytes, golden) in rows {
            let got = audit_text(bytes, label)
                .unwrap_or_else(|r| panic!("{label} did not read: {}", r.text));
            let want = golden.replace("\r\n", "\n");
            if got != want {
                let line = got
                    .lines()
                    .zip(want.lines())
                    .position(|(a, b)| a != b)
                    .map(|i| i + 1)
                    .unwrap_or(got.lines().count().min(want.lines().count()) + 1);
                panic!("{label}: the findings differ from the golden at line {line}");
            }
        }
    }
}
