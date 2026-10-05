//! `--cone <sheet>!<cell>` (PORT.8, slice 8b): a cell's cone sized through
//! the file, which is `AXM.1`'s rule made exact: Excel's `Precedents` stops
//! at the sheet boundary and puts a cone between a floor and a ceiling; a
//! reader that follows every reference, through names and Tables and across
//! sheets, sizes it, and says where it is blind.
//!
//! The one exception to the reader's streaming: an [`Index`] holds every
//! formula's text by address and every non-blank cell's address (a quarter
//! of a million formulas at forty bytes is 10 MB), filled by one walk of the
//! workbook through the ordinary sink, so that a cone costs one scan and
//! never more. The walk is breadth-first from the root over the references
//! `refers::scan` finds in each formula: a cell is one cell; a range, a
//! column or a row is the cells inside it that the file holds (blank cells
//! inside a range are not counted, since nothing feeds from them); a name is
//! what the `name` relation says it refers to, scanned in turn; a structured
//! reference is its Table's whole range; a 3D span is the sheets between its
//! ends in tab order; a spill is its anchor; an external reference is
//! counted and not followed, since it is another file; `#REF!` is counted
//! as broken; `INDIRECT` and `OFFSET` are counted as blind. The set of cells
//! seen is capped at four million, past which the walk stops and says so.
//! The output is counts alone: no address but the one the caller typed.

use std::collections::{BTreeMap, HashSet, VecDeque};

use super::ooxml::Package;
use super::{Row, Sink};
use crate::intrinsics::fold;
use crate::refers::{self, Kind, Shape};
use crate::sheet::{parse_a1_range, MAX_COLUMN, MAX_ROW};

/// The cap on cells a cone walk keeps.
pub const SEEN_CAP: usize = 4_000_000;

/// Every formula's text and every non-blank cell's address, sheet by sheet,
/// as the walk handed them over.
#[derive(Debug, Default)]
pub struct Index {
    /// Each sheet's name, folded, in tab order.
    sheets: Vec<String>,
    /// Each sheet's formulas by (row, column).
    formulas: Vec<BTreeMap<(u32, u32), String>>,
    /// Each sheet's cells that hold a value, by (row, column).
    values: Vec<BTreeMap<(u32, u32), ()>>,
}

impl Index {
    /// An index for the sheets of a package, in tab order.
    pub fn new(package: &Package<'_>) -> Index {
        let n = package.sheets().len();
        Index {
            sheets: package.sheets().iter().map(|s| fold(&s.name)).collect(),
            formulas: (0..n).map(|_| BTreeMap::new()).collect(),
            values: (0..n).map(|_| BTreeMap::new()).collect(),
        }
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        let f = fold(name);
        self.sheets.iter().position(|s| *s == f)
    }

    fn holds(&self, sheet: usize, at: (u32, u32)) -> bool {
        self.formulas[sheet].contains_key(&at) || self.values[sheet].contains_key(&at)
    }

    /// The cells the file holds inside a rectangle of a sheet.
    fn cells_in(
        &self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
    ) -> Vec<(u32, u32)> {
        let mut out: Vec<(u32, u32)> = Vec::new();
        for map_keys in [
            self.formulas[sheet]
                .range((top, 1)..=(bottom, MAX_COLUMN))
                .map(|(k, _)| *k)
                .collect::<Vec<_>>(),
            self.values[sheet]
                .range((top, 1)..=(bottom, MAX_COLUMN))
                .map(|(k, _)| *k)
                .collect::<Vec<_>>(),
        ] {
            for (row, col) in map_keys {
                if col >= left && col <= right {
                    out.push((row, col));
                }
            }
        }
        out
    }
}

impl Sink for Index {
    fn row(&mut self, row: &Row<'_>) {
        match row {
            Row::Cell { sheet, addr, .. } => {
                if let (Some(i), Some(a)) = (self.sheet_index(sheet), parse_a1_range(addr)) {
                    self.values[i].insert((a.top, a.left), ());
                }
            }
            Row::Formula { sheet, addr, text } => {
                if let (Some(i), Some(a)) = (self.sheet_index(sheet), parse_a1_range(addr)) {
                    self.formulas[i].insert((a.top, a.left), (*text).to_string());
                }
            }
            _ => {}
        }
    }
}

/// What a cone walk found: counts alone.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ConeStats {
    /// Distinct cells reached that the file holds, the root among them.
    pub cells: u64,
    /// Those that hold a formula.
    pub formulas: u64,
    /// Those that hold a value and no formula.
    pub inputs: u64,
    /// Single-cell references to a cell the file does not hold.
    pub blanks: u64,
    /// Distinct sheets reached.
    pub sheets: u64,
    /// The longest chain of references from the root.
    pub depth: u64,
    /// `INDIRECT` and `OFFSET` calls met: the cone is blind there.
    pub blind: u64,
    /// Names resolved through the `name` relation.
    pub names: u64,
    /// Structured references resolved through the `table` relation.
    pub tables: u64,
    /// References into another workbook, counted and not followed.
    pub external: u64,
    /// `#REF!` references met.
    pub broken: u64,
    /// References to a sheet, name or Table the file does not have.
    pub unresolved: u64,
    /// The walk stopped at the cap on cells seen.
    pub truncated: bool,
}

impl ConeStats {
    /// The one line a door prints, after the root the caller typed.
    pub fn line(&self) -> String {
        let mut s = format!(
            "cells {} formulas {} inputs {} blanks {} sheets {} depth {} blind {} names {} tables {} external {} broken {} unresolved {}",
            self.cells,
            self.formulas,
            self.inputs,
            self.blanks,
            self.sheets,
            self.depth,
            self.blind,
            self.names,
            self.tables,
            self.external,
            self.broken,
            self.unresolved
        );
        if self.truncated {
            s.push_str(" truncated");
        }
        s
    }
}

/// One cell to visit, with the chain length that reached it.
type Pending = (usize, u32, u32, u64);

/// The cone of the cell at (`row`, `col`) of the sheet at `sheet`, sized
/// through the index and the package's names and Tables.
pub fn cone(index: &Index, package: &Package<'_>, sheet: usize, row: u32, col: u32) -> ConeStats {
    let mut stats = ConeStats::default();
    let mut seen: HashSet<(usize, u32, u32)> = HashSet::new();
    let mut sheets_seen: HashSet<usize> = HashSet::new();
    let mut queue: VecDeque<Pending> = VecDeque::new();
    let mut walk = Walk {
        index,
        package,
        stats: &mut stats,
        seen: &mut seen,
        sheets_seen: &mut sheets_seen,
        queue: &mut queue,
    };
    walk.reach(sheet, row, col, 0, true);
    while let Some((s, r, c, depth)) = walk.queue.pop_front() {
        if walk.stats.truncated {
            break;
        }
        let Some(text) = walk.index.formulas[s].get(&(r, c)).cloned() else {
            continue;
        };
        walk.follow(&text, s, r, c, depth + 1);
    }
    stats.sheets = sheets_seen.len() as u64;
    stats
}

struct Walk<'w> {
    index: &'w Index,
    package: &'w Package<'w>,
    stats: &'w mut ConeStats,
    seen: &'w mut HashSet<(usize, u32, u32)>,
    sheets_seen: &'w mut HashSet<usize>,
    queue: &'w mut VecDeque<Pending>,
}

impl Walk<'_> {
    /// A single cell reached: counted once, queued if it holds a formula.
    /// A cell the file does not hold is a blank, unless it is the root.
    fn reach(&mut self, sheet: usize, row: u32, col: u32, depth: u64, root: bool) {
        if self.seen.len() >= SEEN_CAP {
            self.stats.truncated = true;
            return;
        }
        if !self.seen.insert((sheet, row, col)) {
            return;
        }
        self.sheets_seen.insert(sheet);
        if depth > self.stats.depth {
            self.stats.depth = depth;
        }
        let has_formula = self.index.formulas[sheet].contains_key(&(row, col));
        if has_formula {
            self.stats.cells += 1;
            self.stats.formulas += 1;
            self.queue.push_back((sheet, row, col, depth));
        } else if self.index.holds(sheet, (row, col)) || root {
            self.stats.cells += 1;
            self.stats.inputs += 1;
        } else {
            self.stats.blanks += 1;
        }
    }

    /// Every cell the file holds inside a rectangle.
    fn reach_rect(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        depth: u64,
    ) {
        let (top, bottom) = (top.min(bottom), top.max(bottom));
        let (left, right) = (left.min(right), left.max(right));
        for (row, col) in self.index.cells_in(sheet, top, left, bottom, right) {
            self.reach(sheet, row, col, depth, false);
            if self.stats.truncated {
                return;
            }
        }
    }

    /// A part with numbers, on one sheet.
    fn reach_part(&mut self, r: &refers::FormulaRef, sheet: usize, depth: u64) {
        match r.shape {
            Shape::Cell => self.reach(sheet, r.row1, r.col1, depth, false),
            Shape::Range => self.reach_rect(sheet, r.row1, r.col1, r.row2, r.col2, depth),
            Shape::Column => self.reach_rect(sheet, 1, r.col1, MAX_ROW, r.col2, depth),
            Shape::Row => self.reach_rect(sheet, r.row1, 1, r.row2, MAX_COLUMN, depth),
            Shape::None => {}
        }
    }

    /// The sheet a reference names, or the home sheet when it names none.
    fn sheet_of(&mut self, r: &refers::FormulaRef, home: usize) -> Option<usize> {
        if r.sheet_name.is_empty() {
            return Some(home);
        }
        let found = self.index.sheet_index(&r.sheet_name);
        if found.is_none() {
            self.stats.unresolved += 1;
        }
        found
    }

    /// Every reference of a formula's text, from the cell at (`row`, `col`)
    /// of `home`.
    fn follow(&mut self, text: &str, home: usize, row: u32, col: u32, depth: u64) {
        for r in refers::scan(text) {
            if self.stats.truncated {
                return;
            }
            match r.kind {
                Kind::Cell | Kind::Range | Kind::Column | Kind::Row => {
                    if let Some(s) = self.sheet_of(&r, home) {
                        self.reach_part(&r, s, depth);
                    }
                }
                Kind::Spill => {
                    if r.shape == Shape::Cell {
                        if let Some(s) = self.sheet_of(&r, home) {
                            self.reach(s, r.row1, r.col1, depth, false);
                        }
                    } else {
                        let name = r.part.trim_end_matches('#');
                        self.follow_name(name, &r.sheet_name, home, depth);
                    }
                }
                Kind::Name => self.follow_name(&r.part, &r.sheet_name, home, depth),
                Kind::Structured => self.follow_table(&r, home, row, col, depth),
                Kind::External => self.stats.external += 1,
                Kind::ThreeD => self.follow_span(&r, depth),
                Kind::Unreadable => self.stats.blind += 1,
                Kind::Broken => self.stats.broken += 1,
            }
        }
    }

    /// A name, scoped to the sheet written in front of it, to the home
    /// sheet, or to the workbook, in that order; what it refers to is
    /// followed as a formula of its own, with its sheet as home.
    fn follow_name(&mut self, name: &str, scope: &str, home: usize, depth: u64) {
        let f = fold(name);
        let scope_index = if scope.is_empty() {
            None
        } else {
            self.index.sheet_index(scope)
        };
        let names = self.package.names();
        let pick = |want: Option<usize>| {
            names
                .iter()
                .find(|n| !n.placeholder && fold(&n.name) == f && n.sheet == want)
        };
        let found = match scope_index {
            Some(s) => pick(Some(s)),
            None if scope.is_empty() => pick(Some(home)).or_else(|| pick(None)),
            None => None,
        };
        let Some(info) = found else {
            self.stats.unresolved += 1;
            return;
        };
        self.stats.names += 1;
        let text = info.refers_to.clone();
        let name_home = info.sheet.unwrap_or(home);
        // A name's text is a formula without its `=`; its bare references
        // are the sheet's it is scoped to. A name that names another name
        // is followed through it; the seen set bounds the walk.
        self.follow(&text, name_home, 0, 0, depth);
    }

    /// A structured reference: its Table's whole range, the Table named in
    /// front of the brackets, or the one the cell sits in when none is.
    fn follow_table(
        &mut self,
        r: &refers::FormulaRef,
        home: usize,
        row: u32,
        col: u32,
        depth: u64,
    ) {
        let written = r.part.as_str();
        let table_name = written.split('[').next().unwrap_or("").trim();
        let tables = self.package.tables();
        let found = if table_name.is_empty() {
            tables.iter().find(|t| {
                t.sheet == home
                    && parse_a1_range(&t.range).is_some_and(|a| {
                        row >= a.top && row <= a.bottom && col >= a.left && col <= a.right
                    })
            })
        } else {
            let f = fold(table_name);
            tables.iter().find(|t| fold(&t.name) == f)
        };
        let Some(table) = found else {
            self.stats.unresolved += 1;
            return;
        };
        let Some(range) = parse_a1_range(&table.range) else {
            self.stats.unresolved += 1;
            return;
        };
        self.stats.tables += 1;
        self.reach_rect(
            table.sheet,
            range.top,
            range.left,
            range.bottom,
            range.right,
            depth,
        );
    }

    /// A 3D span: the part on every sheet from the first to the last in tab
    /// order.
    fn follow_span(&mut self, r: &refers::FormulaRef, depth: u64) {
        let Some((first, last)) = r.sheet_name.split_once(':') else {
            self.stats.unresolved += 1;
            return;
        };
        let (Some(a), Some(b)) = (self.index.sheet_index(first), self.index.sheet_index(last))
        else {
            self.stats.unresolved += 1;
            return;
        };
        for s in a.min(b)..=a.max(b) {
            self.reach_part(r, s, depth);
            if self.stats.truncated {
                return;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::zip;

    const NS: &str = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// Two sheets, a name, a Table and an external link: every arm of the
    /// walk reachable from one root or another.
    fn package_bytes() -> Vec<u8> {
        let parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Data\" sheetId=\"2\" r:id=\"rId2\"/></sheets><externalReferences><externalReference r:id=\"rId3\"/></externalReferences><definedNames><definedName name=\"Rate\">Model!$B$2</definedName><definedName name=\"Local\" localSheetId=\"0\">Model!$B$1</definedName><definedName name=\"Deep\">Rate</definedName></definedNames></workbook>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"{REL}/externalLink\" Target=\"externalLinks/externalLink1.xml\"/></Relationships>").into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!(concat!(
                "<worksheet {NS}><sheetData>",
                "<row r=\"1\"><c r=\"A1\"><f>B1+C1</f></c><c r=\"B1\"><v>5</v></c><c r=\"C1\"><f>D1*2</f></c><c r=\"D1\"><v>7</v></c>",
                "<c r=\"E1\"><f>INDIRECT(\"B1\")</f></c><c r=\"F1\"><f>Rate+Local</f></c><c r=\"G1\"><f>SUM(Sales[Amount])</f></c>",
                "<c r=\"H1\"><f>SUM(Model:Data!B1)</f></c><c r=\"I1\"><f>[1]Sheet1!A1+#REF!</f></c><c r=\"J1\"><f>SUM(B:B)+Z9</f></c>",
                "<c r=\"K1\"><f>Deep</f></c><c r=\"L1\"><f>Missing!A1+NoSuchName+NoTable[Col]</f></c></row>",
                "<row r=\"2\"><c r=\"B2\"><v>0.2</v></c></row>",
                "</sheetData></worksheet>"), NS = NS).into_bytes()),
            ("xl/worksheets/sheet2.xml".to_string(), format!("<worksheet {NS}><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Item</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>pens</t></is></c><c r=\"B2\"><v>10</v></c></row></sheetData><tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts></worksheet>").into_bytes()),
            ("xl/worksheets/_rels/sheet2.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/table\" Target=\"../tables/table1.xml\"/></Relationships>").into_bytes()),
            ("xl/tables/table1.xml".to_string(), format!("<table {NS} id=\"1\" name=\"Sales\" displayName=\"Sales\" ref=\"A1:B2\"><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Amount\"/></tableColumns></table>").into_bytes()),
            ("xl/externalLinks/externalLink1.xml".to_string(), format!("<externalLink {NS}><externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Sheet1\"/></sheetNames></externalBook></externalLink>").into_bytes()),
            ("xl/externalLinks/_rels/externalLink1.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL}/externalLinkPath\" Target=\"Rates.xlsx\" TargetMode=\"External\"/></Relationships>").into_bytes()),
        ];
        zip::write_stored(&parts).unwrap()
    }

    fn sized(root: &str) -> ConeStats {
        let bytes = package_bytes();
        let package = Package::open(&bytes, "t.xlsx").unwrap();
        let mut index = Index::new(&package);
        for i in 0..package.sheets().len() {
            package.walk_sheet(i, &mut index).unwrap();
        }
        let (sheet, row, col) = refers::parse_home(root).unwrap();
        let s = index.sheet_index(&sheet).unwrap();
        cone(&index, &package, s, row, col)
    }

    fn counts(c: &ConeStats) -> (u64, u64, u64, u64, u64, u64) {
        (c.cells, c.formulas, c.inputs, c.blanks, c.sheets, c.depth)
    }

    #[test]
    fn a_chain_is_sized_across_cells() {
        // A1 = B1 + C1; C1 = D1 * 2: four cells, two formulas, two inputs,
        // depth two.
        let c = sized("Model!A1");
        assert_eq!(counts(&c), (4, 2, 2, 0, 1, 2));
        assert_eq!(
            (
                c.blind,
                c.names,
                c.tables,
                c.external,
                c.broken,
                c.unresolved
            ),
            (0, 0, 0, 0, 0, 0)
        );
    }

    #[test]
    fn an_input_is_a_cone_of_one() {
        let c = sized("Model!D1");
        assert_eq!(counts(&c), (1, 0, 1, 0, 1, 0));
    }

    #[test]
    fn blind_names_tables_spans_external_and_broken_are_counted() {
        let c = sized("Model!E1");
        assert_eq!((counts(&c), c.blind), ((1, 1, 0, 0, 1, 0), 1));
        // Rate -> Model!$B$2 (0.2), Local (sheet-scoped) -> Model!$B$1 (5).
        let c = sized("Model!F1");
        assert_eq!((counts(&c), c.names), ((3, 1, 2, 0, 1, 1), 2));
        // Sales on Data over A1:B2: four cells the file holds.
        let c = sized("Model!G1");
        assert_eq!((counts(&c), c.tables), ((5, 1, 4, 0, 2, 1), 1));
        // Model:Data!B1: B1 on both sheets.
        let c = sized("Model!H1");
        assert_eq!(counts(&c), (3, 1, 2, 0, 2, 1));
        let c = sized("Model!I1");
        assert_eq!(
            (counts(&c), c.external, c.broken),
            ((1, 1, 0, 0, 1, 0), 1, 1)
        );
        // B:B holds B1 and B2 on Model; Z9 is a blank.
        let c = sized("Model!J1");
        assert_eq!(counts(&c), (3, 1, 2, 1, 1, 1));
        // Deep -> Rate -> Model!$B$2: two names, one cell.
        let c = sized("Model!K1");
        assert_eq!((counts(&c), c.names), ((2, 1, 1, 0, 1, 1), 2));
        let c = sized("Model!L1");
        assert_eq!((counts(&c), c.unresolved), ((1, 1, 0, 0, 1, 0), 3));
    }

    #[test]
    fn the_line_is_counts_alone() {
        let c = ConeStats {
            cells: 4,
            formulas: 2,
            inputs: 2,
            depth: 2,
            sheets: 1,
            ..ConeStats::default()
        };
        assert_eq!(
            c.line(),
            "cells 4 formulas 2 inputs 2 blanks 0 sheets 1 depth 2 blind 0 names 0 tables 0 external 0 broken 0 unresolved 0"
        );
        let t = ConeStats {
            truncated: true,
            ..ConeStats::default()
        };
        assert!(t.line().ends_with(" truncated"));
    }
}
