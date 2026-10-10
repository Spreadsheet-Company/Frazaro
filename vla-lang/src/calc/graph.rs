//! The dependency graph and the order of evaluation (KERNEL.7,
//! 2026-10-08): from `refers`, as the roadmap says.
//!
//! Every formula cell is a node; an edge runs from a formula to each
//! formula cell it reads, through a cell, a range, a whole column or row,
//! a 3D span, or a defined name whose text reads them. The scanner's
//! records (`refers::scan`) are the edges' source, so the graph follows the
//! refers golden and nothing else; a construct the parser does not read
//! adds no edge, and the cell that holds it is not computed. Tarjan's
//! algorithm over the graph finds the strongly connected components in
//! reverse topological order, which with edges pointing from a reader to
//! what it reads is the order a dependency is computed before the formula
//! that needs it: the evaluation order. A component of more than one cell,
//! or one cell that reads itself, is a cycle, refused by name with the
//! cells that close it (`calc-cycle`), as `Alonzo/SPEC.md` section 2 asks
//! for the engine's `load` and the treaty's recalculation refuses alike.
//! The walk is iterative: a running total down ten thousand rows is a
//! chain ten thousand deep, which a recursive walk would not survive.

use std::collections::btree_map::Range;
use std::collections::HashMap;

use crate::intrinsics::fold;
use crate::refers::{self, Kind};
use crate::sheet::{Cell, Content, Sheet, Workbook};

use super::formula::{ref_expr, RefExpr};
use super::value::ErrorKind;

/// A cell by its sheet's tab index, its row and its column, 1-based.
pub type CellId = (usize, u32, u32);

/// The used rectangle of a sheet, as `Sheet::extent` gives it.
pub type Extent = Option<((u32, u32), (u32, u32))>;

/// A rectangle of cells on one sheet, rows and columns 1-based and
/// inclusive; a whole column runs to the sheet's last row, a whole row to
/// its last column, and the walk clamps either to the sheet's extent.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Area {
    pub sheet: usize,
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

impl Area {
    /// Whether the area is one cell.
    pub fn is_cell(&self) -> bool {
        self.top == self.bottom && self.left == self.right
    }
}

/// The areas a reference names, read as a slice: one area held in place, or
/// a 3D span's, one a sheet, in a vector (KERNEL.25: a reference to one
/// sheet, which every evaluation of nearly every formula makes, allocates
/// nothing).
#[derive(Clone, Debug)]
pub struct Areas(Held);

#[derive(Clone, Debug)]
enum Held {
    One(Area),
    Span(Vec<Area>),
}

impl Areas {
    /// One area.
    pub fn one(area: Area) -> Areas {
        Areas(Held::One(area))
    }
}

impl From<Vec<Area>> for Areas {
    fn from(areas: Vec<Area>) -> Areas {
        match areas.as_slice() {
            [area] => Areas::one(*area),
            _ => Areas(Held::Span(areas)),
        }
    }
}

impl std::ops::Deref for Areas {
    type Target = [Area];

    fn deref(&self) -> &[Area] {
        match &self.0 {
            Held::One(area) => std::slice::from_ref(area),
            Held::Span(areas) => areas,
        }
    }
}

impl<'a> IntoIterator for &'a Areas {
    type Item = &'a Area;
    type IntoIter = std::slice::Iter<'a, Area>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Two lists of areas are equal when they read the same, however held.
impl PartialEq for Areas {
    fn eq(&self, other: &Areas) -> bool {
        **self == **other
    }
}

impl Eq for Areas {}

/// The defined names of a workbook, each with the sheet that scopes it and
/// its text, as the reader prints a `name` row (`Model!Local` for a
/// sheet-scoped name) and as the writer holds one.
pub struct Names {
    entries: Vec<NameEntry>,
}

struct NameEntry {
    key: String,
    scope: Option<usize>,
    text: String,
}

impl Names {
    /// The workbook's names, read off `Workbook::defined_names`.
    pub fn of(model: &Workbook) -> Names {
        let entries = model
            .defined_names
            .iter()
            .map(|(printed, text)| {
                let (scope, bare) = split_scope(model, printed);
                NameEntry {
                    key: fold(bare),
                    scope,
                    text: text.clone(),
                }
            })
            .collect();
        Names { entries }
    }

    /// The text of the name a formula on `home` reads by `name`: the name
    /// scoped to the sheet written in front of it when one is, else the
    /// one scoped to the formula's own sheet, else the workbook's.
    pub fn resolve(&self, name: &str, scope: Option<usize>, home: usize) -> Option<&str> {
        let key = fold(name);
        let scoped = |s: usize| {
            self.entries
                .iter()
                .find(|e| e.key == key && e.scope == Some(s))
        };
        let global = || {
            self.entries
                .iter()
                .find(|e| e.key == key && e.scope.is_none())
        };
        let entry = match scope {
            Some(s) => scoped(s).or_else(global),
            None => scoped(home).or_else(global),
        };
        entry.map(|e| e.text.as_str())
    }

    /// How many names there are.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether there is none.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

/// `Model!Local` or `'Q1 Data'!Local` as its scope and its bare name; a
/// name with no `!` is the workbook's.
fn split_scope<'a>(model: &Workbook, printed: &'a str) -> (Option<usize>, &'a str) {
    let Some(bang) = printed.rfind('!') else {
        return (None, printed);
    };
    let (sheet_part, bare) = (&printed[..bang], &printed[bang + 1..]);
    let sheet =
        if sheet_part.len() >= 2 && sheet_part.starts_with('\'') && sheet_part.ends_with('\'') {
            sheet_part[1..sheet_part.len() - 1].replace("''", "'")
        } else {
            sheet_part.to_string()
        };
    (model.find_sheet(&sheet), bare)
}

/// The areas a reference names against the model: one per sheet of a
/// span, the formula's own sheet when none is written. A sheet the model
/// does not hold is `#REF!`.
pub fn areas_of(model: &Workbook, home: usize, r: &RefExpr) -> Result<Areas, ErrorKind> {
    areas_at(model, home, r, 0, 0)
}

/// The areas a reference names from a cell `d_row` rows down and `d_col`
/// columns right of the one its formula was read at (KERNEL.22, a shape
/// evaluated at an offset): [`RefExpr::moved`]'s rectangle on each sheet,
/// `#REF!` when a corner leaves the sheet, as the moved text would read.
pub fn areas_at(
    model: &Workbook,
    home: usize,
    r: &RefExpr,
    d_row: i64,
    d_col: i64,
) -> Result<Areas, ErrorKind> {
    let first = match &r.sheet {
        None => home,
        Some(name) => model.find_sheet(name).ok_or(ErrorKind::Ref)?,
    };
    let last = match &r.last_sheet {
        None => first,
        Some(name) => model.find_sheet(name).ok_or(ErrorKind::Ref)?,
    };
    let (top, left, bottom, right) = r.moved(d_row, d_col).ok_or(ErrorKind::Ref)?;
    let area = |sheet| Area {
        sheet,
        top,
        left,
        bottom,
        right,
    };
    if first == last {
        return Ok(Areas::one(area(first)));
    }
    let (a, b) = (first.min(last), first.max(last));
    Ok(Areas::from((a..=b).map(area).collect::<Vec<Area>>()))
}

/// The cells a sheet holds inside an area, in row-major order, the area
/// clamped to the sheet's extent first so that a whole column costs its
/// rows and not Excel's million. Walked as it is read and never collected
/// (KERNEL.25): a range read by every evaluation of a formula cost a vector
/// each time.
pub fn cells_in<'a>(
    sheet: &'a Sheet,
    extent: Extent,
    area: &Area,
) -> impl Iterator<Item = ((u32, u32), &'a Cell)> + 'a {
    let Some(((e_top, e_left), (e_bottom, e_right))) = extent else {
        return Walk::Done;
    };
    let top = area.top.max(e_top);
    let bottom = area.bottom.min(e_bottom);
    let left = area.left.max(e_left);
    let right = area.right.min(e_right);
    if top > bottom || left > right {
        return Walk::Done;
    }
    let rows = usize::try_from(bottom - top + 1).unwrap_or(usize::MAX);
    if rows > sheet.cells.len() {
        // Fewer cells than rows: walk the cells and keep those inside.
        return Walk::Sparse {
            cells: sheet.cells.range((top, 0)..=(bottom, u32::MAX)),
            left,
            right,
        };
    }
    Walk::Rows {
        sheet,
        row: top,
        bottom,
        left,
        right,
        cells: sheet.cells.range((top, left)..=(top, right)),
    }
}

/// [`cells_in`]'s two walks: the sheet's cells between the area's first and
/// last rows, those in its columns kept; or the area row by row, a range of
/// the sheet's cells a row.
enum Walk<'a> {
    Done,
    Sparse {
        cells: Range<'a, (u32, u32), Cell>,
        left: u32,
        right: u32,
    },
    Rows {
        sheet: &'a Sheet,
        row: u32,
        bottom: u32,
        left: u32,
        right: u32,
        cells: Range<'a, (u32, u32), Cell>,
    },
}

impl<'a> Iterator for Walk<'a> {
    type Item = ((u32, u32), &'a Cell);

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Walk::Done => None,
            Walk::Sparse { cells, left, right } => {
                let (left, right) = (*left, *right);
                cells
                    .find(|((_, col), _)| *col >= left && *col <= right)
                    .map(|(k, c)| (*k, c))
            }
            Walk::Rows {
                sheet,
                row,
                bottom,
                left,
                right,
                cells,
            } => loop {
                if let Some((k, c)) = cells.next() {
                    return Some((*k, c));
                }
                if *row >= *bottom {
                    return None;
                }
                *row += 1;
                *cells = sheet.cells.range((*row, *left)..=(*row, *right));
            },
        }
    }
}

/// Whether a cell's content is a formula of any kind.
pub fn is_formula(content: &Content) -> bool {
    matches!(
        content,
        Content::Formula(_)
            | Content::SharedMaster { .. }
            | Content::SharedChild { .. }
            | Content::DynamicFormula(_)
    )
}

/// How deep a name may read a name.
const NAME_DEPTH: u32 = 8;

/// The formula cells a formula's text reads, through its references and
/// the names they resolve to, in formula order, repeats kept.
pub fn dependencies(
    model: &Workbook,
    names: &Names,
    extents: &[Extent],
    home: usize,
    text: &str,
    out: &mut Vec<CellId>,
) {
    collect(model, names, extents, home, text, 0, out);
}

fn collect(
    model: &Workbook,
    names: &Names,
    extents: &[Extent],
    home: usize,
    text: &str,
    depth: u32,
    out: &mut Vec<CellId>,
) {
    for r in refers::scan(text) {
        match r.kind {
            Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD => {
                let Some(re) = ref_expr(&r) else { continue };
                let Ok(areas) = areas_of(model, home, &re) else {
                    continue;
                };
                for a in &areas {
                    let sheet = &model.sheets[a.sheet];
                    for ((row, col), cell) in cells_in(sheet, extents[a.sheet], a) {
                        if is_formula(&cell.content) {
                            out.push((a.sheet, row, col));
                        }
                    }
                }
            }
            Kind::Name if depth < NAME_DEPTH => {
                let scope = if r.sheet_name.is_empty() {
                    None
                } else {
                    model.find_sheet(&r.sheet_name)
                };
                if let Some(t) = names.resolve(&r.part, scope, home) {
                    collect(
                        model,
                        names,
                        extents,
                        scope.unwrap_or(home),
                        t,
                        depth + 1,
                        out,
                    );
                }
            }
            _ => {}
        }
    }
}

/// The plan: every formula cell either in the evaluation order, each after
/// everything it reads, or in a cycle.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Plan {
    /// The formula cells with no cycle through them, dependencies first.
    pub order: Vec<CellId>,
    /// Each cycle's cells, sorted; the cycles sorted by their first cell.
    pub cycles: Vec<Vec<CellId>>,
}

impl Plan {
    /// How many formula cells the plan covers, in the order or in a cycle.
    pub fn formulas(&self) -> usize {
        self.order.len() + self.cycles.iter().map(Vec::len).sum::<usize>()
    }
}

/// The plan over the formula cells given, each with its text (without its
/// `=`), against the model and its names.
pub fn plan(
    model: &Workbook,
    names: &Names,
    extents: &[Extent],
    texts: &HashMap<CellId, String>,
) -> Plan {
    let mut ids: Vec<CellId> = texts.keys().copied().collect();
    ids.sort_unstable();
    let index: HashMap<CellId, usize> = ids.iter().enumerate().map(|(i, id)| (*id, i)).collect();
    let n = ids.len();
    let mut adj: Vec<Vec<usize>> = vec![Vec::new(); n];
    let mut deps = Vec::new();
    for (vi, id) in ids.iter().enumerate() {
        deps.clear();
        dependencies(model, names, extents, id.0, &texts[id], &mut deps);
        let edges = &mut adj[vi];
        for d in &deps {
            if let Some(&di) = index.get(d) {
                edges.push(di);
            }
        }
        edges.sort_unstable();
        edges.dedup();
    }
    let components = tarjan(&adj);
    let mut plan = Plan::default();
    for comp in components {
        if comp.len() == 1 && !adj[comp[0]].contains(&comp[0]) {
            plan.order.push(ids[comp[0]]);
        } else {
            let mut cells: Vec<CellId> = comp.iter().map(|&i| ids[i]).collect();
            cells.sort_unstable();
            plan.cycles.push(cells);
        }
    }
    plan.cycles.sort();
    plan
}

/// Tarjan's strongly connected components, iteratively, in the order the
/// algorithm emits them: a component after every component it has an edge
/// into, which is the evaluation order when an edge runs from a reader to
/// what it reads. The machine's plan of shapes orders its cells with it too
/// (KERNEL.22).
pub(crate) fn tarjan(adj: &[Vec<usize>]) -> Vec<Vec<usize>> {
    let n = adj.len();
    let mut index: Vec<Option<usize>> = vec![None; n];
    let mut low: Vec<usize> = vec![0; n];
    let mut on_stack: Vec<bool> = vec![false; n];
    let mut stack: Vec<usize> = Vec::new();
    let mut components: Vec<Vec<usize>> = Vec::new();
    let mut counter = 0usize;
    for root in 0..n {
        if index[root].is_some() {
            continue;
        }
        let mut calls: Vec<(usize, usize)> = vec![(root, 0)];
        index[root] = Some(counter);
        low[root] = counter;
        counter += 1;
        stack.push(root);
        on_stack[root] = true;
        while let Some(&mut (v, ref mut next)) = calls.last_mut() {
            if *next < adj[v].len() {
                let w = adj[v][*next];
                *next += 1;
                match index[w] {
                    None => {
                        index[w] = Some(counter);
                        low[w] = counter;
                        counter += 1;
                        stack.push(w);
                        on_stack[w] = true;
                        calls.push((w, 0));
                    }
                    Some(iw) if on_stack[w] => low[v] = low[v].min(iw),
                    Some(_) => {}
                }
                continue;
            }
            calls.pop();
            if let Some(&(u, _)) = calls.last() {
                low[u] = low[u].min(low[v]);
            }
            if low[v] == index[v].unwrap_or(0) {
                let mut comp = Vec::new();
                while let Some(w) = stack.pop() {
                    on_stack[w] = false;
                    comp.push(w);
                    if w == v {
                        break;
                    }
                }
                comp.sort_unstable();
                components.push(comp);
            }
        }
    }
    components
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tarjan_emits_dependencies_before_their_readers() {
        // 0 reads 1 and 2; 1 reads 2; 3 and 4 read each other; 5 reads itself.
        let adj = vec![vec![1, 2], vec![2], vec![], vec![4], vec![3], vec![5]];
        let comps = tarjan(&adj);
        let pos = |x: usize| comps.iter().position(|c| c.contains(&x)).unwrap();
        assert!(pos(2) < pos(1) && pos(1) < pos(0));
        assert!(comps.contains(&vec![3, 4]));
        assert!(comps.contains(&vec![5]));
        assert_eq!(comps.len(), 5);
    }

    #[test]
    fn a_reference_s_areas_are_one_held_in_place_or_a_span_s_in_order() {
        // KERNEL.25: one area is held in place and read as a slice of one;
        // a span's are one a sheet, in the sheets' order.
        use super::super::formula::{parse, Expr};
        let mut wb = Workbook::new();
        for name in ["Jan", "Feb", "Mar"] {
            wb.sheets.push(Sheet::new(name));
        }
        let areas = |text: &str| {
            let Ok(Expr::Ref(r)) = parse(text) else {
                panic!("{text}")
            };
            areas_of(&wb, 0, &r)
        };
        let rect = |sheet, top, left, bottom, right| Area {
            sheet,
            top,
            left,
            bottom,
            right,
        };
        let one = areas("=Feb!B2:C3").unwrap();
        assert_eq!(&one[..], &[rect(1, 2, 2, 3, 3)]);
        assert_eq!(one, Areas::one(rect(1, 2, 2, 3, 3)));
        assert_eq!(one, Areas::from(vec![rect(1, 2, 2, 3, 3)]));
        assert_eq!(
            &areas("=B2").unwrap()[..],
            &[rect(0, 2, 2, 2, 2)],
            "the formula's own sheet"
        );
        let span = areas("=Jan:Mar!B2").unwrap();
        let sheets: Vec<usize> = span.iter().map(|a| a.sheet).collect();
        assert_eq!(sheets, vec![0, 1, 2]);
        assert_eq!(areas("=Mar:Jan!B2").unwrap(), span);
        assert_ne!(span, one);
        assert_eq!(areas("=Apr!B2"), Err(ErrorKind::Ref));
    }

    #[test]
    fn a_range_s_cells_are_walked_row_by_row_as_they_were_collected() {
        // KERNEL.25: the walk is read as it goes. Its cells and their order
        // are the sheet's own cells inside the area clamped to the extent,
        // row by row, in both its walks: row by row through each row's
        // cells, and through every cell between the first and last rows for
        // a sheet with fewer cells than the area has rows.
        use crate::sheet::{MAX_COLUMN, MAX_ROW};
        let number = |n: u32| Cell {
            content: Content::Number(f64::from(n)),
            style: 0,
        };
        let mut dense = Sheet::new("Dense");
        for row in 1..=5u32 {
            for col in 1..=4u32 {
                if (row + col) % 5 != 0 {
                    dense.set(row, col, number(row * 10 + col));
                }
            }
        }
        let mut sparse = Sheet::new("Sparse");
        for (row, col) in [(1, 2), (40, 1), (40, 3), (700, 2), (1000, 4)] {
            sparse.set(row, col, number(row));
        }
        let empty = Sheet::new("Empty");
        let rect = |top, left, bottom, right| Area {
            sheet: 0,
            top,
            left,
            bottom,
            right,
        };
        let areas = [
            rect(2, 2, 4, 3),
            rect(1, 1, MAX_ROW, 1),
            rect(3, 1, 3, MAX_COLUMN),
            rect(1, 1, MAX_ROW, MAX_COLUMN),
            rect(40, 2, 700, 3),
            rect(6, 1, 9, 9),
            rect(2, 5, 4, 9),
            rect(1, 2, 1, 2),
        ];
        let mut walked_any = 0;
        for sheet in [&dense, &sparse, &empty] {
            let extent = sheet.extent();
            for a in &areas {
                let walked: Vec<(u32, u32)> = cells_in(sheet, extent, a).map(|(k, _)| k).collect();
                let want: Vec<(u32, u32)> = match extent {
                    None => Vec::new(),
                    Some(((t, l), (b, r))) => sheet
                        .cells
                        .keys()
                        .copied()
                        .filter(|&(row, col)| {
                            row >= a.top.max(t)
                                && row <= a.bottom.min(b)
                                && col >= a.left.max(l)
                                && col <= a.right.min(r)
                        })
                        .collect(),
                };
                assert_eq!(walked, want, "{} {a:?}", sheet.name);
                walked_any += walked.len();
            }
        }
        assert_eq!(
            walked_any, 37,
            "28 cells of the dense sheet's, 9 of the sparse one's"
        );
    }

    #[test]
    fn a_deep_chain_does_not_overflow() {
        let n = 200_000;
        let adj: Vec<Vec<usize>> = (0..n)
            .map(|i| if i + 1 < n { vec![i + 1] } else { vec![] })
            .collect();
        let comps = tarjan(&adj);
        assert_eq!(comps.len(), n);
        assert_eq!(comps[0], vec![n - 1]);
        assert_eq!(comps[n - 1], vec![0]);
    }
}
