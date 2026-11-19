//! A dynamic formula's spill as its parse bounds it (KERNEL.55, the item's
//! third slice; decisions 4 and 5): the most rows and columns its answer can
//! take, read off its text before anything computes it, so that a door plans
//! the cells it may fill, and an engine, at its slice, the cells a load fixes.
//! A bound is an upper one, `FILTER`'s its array's whole; where the parse gives
//! none (a count read from a cell, a name, a function of one's own, `OFFSET` or
//! `INDIRECT`), a door plans by the extent the file saved instead, and where
//! neither stands, by the anchor alone: an answer of one value computes, and
//! one of many is labelled.

use std::collections::{HashMap, HashSet};

use super::eval::{Library, Param};
use super::formula::{self, Expr, RefExpr, ANCHORARRAY};
use super::graph::{Area, CellId, Extent, Names};
use super::lambda::{self, Functions};
use crate::grid::{Formula, Masters, Model, Rects};
use crate::sheet::{A1Range, Content, Workbook, MAX_COLUMN, MAX_ROW};

/// Rows and columns.
pub(crate) type Dims = (u32, u32);

/// The label of a spill no plan bounds (decision 5): one whose parse gives
/// no bound and whose file saved no extent, one larger than a door plans, and
/// one whose answer passes the bound it was planned by.
pub(crate) const NO_PLAN: &str = "a spill no plan bounds";

/// The label of a spill whose cells may meet another's, which Excel settles
/// by the order it computes them in (the save of the item's probe, `Spills!K1`
/// spilling and `J3` blocked), labelled and never chased.
pub(crate) const MEETS: &str = "a spill meeting another";

/// The most cells a door plans one spill over: a bound past it, a whole
/// column's `FILTER` among them, is labelled.
const MOST: u64 = 1 << 18;

/// A dynamic formula's spill as a door plans it (KERNEL.55): the most rows
/// and columns its answer may take, the extent its file saved, whose cells
/// hold the values it last spilled, and the box of the cells it may fill, its
/// anchor first, holding both; or why it is labelled and not computed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Spill {
    pub(crate) dims: Dims,
    pub(crate) file: Option<A1Range>,
    pub(crate) cells: A1Range,
    pub(crate) label: Option<&'static str>,
}

/// A dynamic formula as a plan meets it (KERNEL.55): its text, and the extent
/// its file saved past its cell, whose cells hold the values it last spilled.
#[derive(Clone, Copy, Debug)]
pub(crate) struct Dynamic<'a> {
    pub(crate) text: &'a str,
    pub(crate) saved: Option<A1Range>,
}

/// The spills a door plans, and an engine's machine (KERNEL.55): the anchors
/// whose boxes hold more than their own cell, as runs, each a column's rows of
/// one sheet bounded alike, reaching alike past their anchors and labelled
/// alike, so that a grid's column of one shape is one run however many rows
/// it fills; each sheet's runs' boxes, the cells their anchors' boxes hold, as
/// rectangles tagged by the run's place, so that a read and a range meet them
/// by a search and no cell is listed; the extents the files saved, by anchor;
/// and the anchors whose parse gives no bound, which an engine refuses.
#[derive(Clone, Debug, Default)]
pub(crate) struct Spills {
    runs: Vec<Run>,
    boxes: Vec<Rects>,
    files: HashMap<CellId, A1Range>,
    unbounded: Vec<CellId>,
}

/// The anchors of one sheet's column from `top` to `bottom`, each bounded
/// `dims`, each one's box reaching `down` rows below it and `right` columns
/// right of it, as far as a sheet's last; `met` where each one's box may meet
/// another's ([`MEETS`]). Runs are kept by sheet, column and first row.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Run {
    sheet: u32,
    col: u32,
    top: u32,
    bottom: u32,
    dims: Dims,
    down: u32,
    right: u32,
    met: bool,
}

impl Run {
    /// The box of its anchor at a row.
    fn box_at(&self, row: u32) -> A1Range {
        A1Range {
            top: row,
            left: self.col,
            bottom: row.saturating_add(self.down).min(MAX_ROW),
            right: self.col.saturating_add(self.right).min(MAX_COLUMN),
        }
    }

    /// The cells its anchors' boxes hold, which are one rectangle.
    fn cells(&self) -> A1Range {
        A1Range {
            top: self.top,
            ..self.box_at(self.bottom)
        }
    }

    /// Its anchors whose boxes meet the rows from `top` to `bottom`, in order.
    fn meeting(&self, top: u32, bottom: u32) -> impl Iterator<Item = CellId> {
        let (sheet, col) = (self.sheet as usize, self.col);
        let first = self.top.max(top.saturating_sub(self.down));
        (first..=self.bottom.min(bottom)).map(move |row| (sheet, row, col))
    }
}

/// A workbook's dynamic formulas among its formulas' texts, each with the
/// extent its file saved.
pub(crate) fn dynamics_of<'a>(
    model: &Workbook,
    texts: &'a HashMap<CellId, String>,
) -> HashMap<CellId, Dynamic<'a>> {
    texts
        .iter()
        .filter_map(|(&id, text)| {
            let saved = file_of(model, id)?;
            Some((id, Dynamic { text, saved }))
        })
        .collect()
}

impl Spills {
    /// Each dynamic formula's spill: its bound by its parse, a spill read
    /// through `#` its anchor's, else the extent its file saved, else its
    /// anchor alone; none for a formula whose box holds one cell. Two whose
    /// boxes meet are labelled ([`MEETS`]), and one past a door's plan
    /// ([`NO_PLAN`]); one whose parse gives no bound is listed besides.
    pub(crate) fn of(
        dynamics: &HashMap<CellId, Dynamic<'_>>,
        find_sheet: &dyn Fn(&str) -> Option<usize>,
        names: &Names,
        functions: &Functions,
        lib: &Library,
    ) -> Spills {
        let text_of = |id: CellId| dynamics.get(&id).map(|d| (d.text, d.saved));
        Spills::of_each(
            &mut dynamics.keys().copied(),
            &text_of,
            &|_| None,
            find_sheet,
            names,
            functions,
            lib,
        )
    }

    /// Each dynamic formula's spill, as [`Spills::of`] finds it, of the cells
    /// given, each its text and saved extent lent by `text_of`; the cells of
    /// one group, which `group_of` names, bounded once, a grid's formulas of
    /// one shape on one sheet moving their ranges and keeping their sizes,
    /// so that an engine's load bounds a shape and never lists its cells.
    pub(crate) fn of_each<'t>(
        cells: &mut dyn Iterator<Item = CellId>,
        text_of: &dyn Fn(CellId) -> Option<(&'t str, Option<A1Range>)>,
        group_of: &dyn Fn(CellId) -> Option<u64>,
        find_sheet: &dyn Fn(&str) -> Option<usize>,
        names: &Names,
        functions: &Functions,
        lib: &Library,
    ) -> Spills {
        let mut bounds = Bounds {
            text_of,
            group_of,
            find_sheet,
            names,
            functions,
            memo: HashMap::new(),
            groups: HashMap::new(),
            visiting: HashSet::new(),
            unbounded: Vec::new(),
        };
        // Each anchor whose box holds more than its own cell: the parse's
        // bound, the extent its file saved where it gives none (decision 5),
        // and the anchor alone where neither stands; its box reaching as far
        // as the farther of its bound and its file's extent.
        let mut found: Vec<(CellId, Dims, u32, u32)> = Vec::new();
        let mut files: HashMap<CellId, A1Range> = HashMap::new();
        for id in cells {
            let Some((_, file)) = text_of(id) else {
                continue;
            };
            let (rows, columns) = bounds.dims(id, lib).unwrap_or((1, 1));
            let dims = (rows.max(1), columns.max(1));
            let down = (dims.0 - 1).max(file.map_or(0, |f| f.bottom.saturating_sub(id.1)));
            let right = (dims.1 - 1).max(file.map_or(0, |f| f.right.saturating_sub(id.2)));
            let one = id.1.saturating_add(down).min(MAX_ROW) == id.1
                && id.2.saturating_add(right).min(MAX_COLUMN) == id.2;
            if one {
                continue;
            }
            if let Some(f) = file {
                files.insert(id, f);
            }
            found.push((id, dims, down, right));
        }
        let mut unbounded = std::mem::take(&mut bounds.unbounded);
        unbounded.sort_unstable();
        unbounded.dedup();
        if found.is_empty() && unbounded.is_empty() {
            return Spills::default();
        }
        // The runs: the anchors by sheet, column and row, one run while each
        // stands below the last and is bounded and reaches as it does.
        found.sort_unstable_by_key(|&((s, row, col), ..)| (s, col, row));
        let mut runs: Vec<Run> = Vec::new();
        for &((s, row, col), dims, down, right) in &found {
            let sheet = s as u32;
            if let Some(last) = runs.last_mut() {
                let at = (last.sheet, last.col, last.bottom + 1);
                if at == (sheet, col, row)
                    && (last.dims, last.down, last.right) == (dims, down, right)
                {
                    last.bottom = row;
                    continue;
                }
            }
            runs.push(Run {
                sheet,
                col,
                top: row,
                bottom: row,
                dims,
                down,
                right,
                met: false,
            });
        }
        drop(found);
        let runs = labelled(runs);
        let boxes = boxes_of(&runs);
        Spills {
            runs,
            boxes,
            files,
            unbounded,
        }
    }

    /// Whether it holds none.
    pub(crate) fn is_empty(&self) -> bool {
        self.runs.is_empty()
    }

    /// Whether any spill is labelled ([`MEETS`], [`NO_PLAN`]), asked of its
    /// runs, once a frame, where a word kept would grow every table.
    pub(crate) fn labelled(&self) -> bool {
        self.runs
            .iter()
            .any(|r| r.met || u64::from(r.dims.0) * u64::from(r.dims.1) > MOST)
    }

    /// The run holding an anchor.
    fn run_at(&self, (s, row, col): CellId) -> Option<&Run> {
        let at = (s as u32, col, row);
        let k = self.runs.partition_point(|r| (r.sheet, r.col, r.top) <= at);
        let r = self.runs.get(k.checked_sub(1)?)?;
        ((r.sheet, r.col) == (at.0, at.1) && row <= r.bottom).then_some(r)
    }

    /// An anchor's spill.
    pub(crate) fn get(&self, anchor: CellId) -> Option<Spill> {
        let r = self.run_at(anchor)?;
        let past = u64::from(r.dims.0) * u64::from(r.dims.1) > MOST;
        let label = match r.met {
            true => Some(MEETS),
            false => past.then_some(NO_PLAN),
        };
        Some(Spill {
            dims: r.dims,
            file: self.files.get(&anchor).copied(),
            cells: r.box_at(anchor.1),
            label,
        })
    }

    /// An anchor's box, the cells its spill may fill.
    pub(crate) fn box_of(&self, anchor: CellId) -> Option<A1Range> {
        self.run_at(anchor).map(|r| r.box_at(anchor.1))
    }

    /// Each anchor and its box, by sheet, column and row.
    #[cfg(test)]
    pub(crate) fn boxes(&self) -> impl Iterator<Item = (CellId, A1Range)> + '_ {
        self.runs.iter().flat_map(|r| {
            (r.top..=r.bottom).map(move |row| ((r.sheet as usize, row, r.col), r.box_at(row)))
        })
    }

    /// Whether a box holds a cell: a cell a spill may fill, or its anchor.
    pub(crate) fn covers(&self, (s, row, col): CellId) -> bool {
        let cell = A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        };
        self.boxes
            .get(s)
            .is_some_and(|b| b.meeting(cell).next().is_some())
    }

    /// The dynamic formulas whose parse gives no bound, in sheet then
    /// row-major order: what an engine's load refuses (decision 4).
    pub(crate) fn unbounded(&self) -> &[CellId] {
        &self.unbounded
    }

    /// The anchors whose spills may fill a cell, one where an anchor stands in
    /// another's box: each whose box holds it but at its own anchor, the cell
    /// holding nothing or a value inside the extent the anchor's file saved, a
    /// value of its spill's last; never a cell holding a value of its own,
    /// which a spill meeting it is blocked by. A cell holding a formula is
    /// read, and planned, as its own before any spill's.
    pub(crate) fn owners<'s>(
        &'s self,
        model: &'s Workbook,
        (s, row, col): CellId,
    ) -> impl Iterator<Item = CellId> + 's {
        let held = model
            .sheets
            .get(s)
            .and_then(|sheet| sheet.cells.get(&(row, col)));
        let cell = A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        };
        self.meeting(s, cell).filter(move |&anchor| {
            anchor != (s, row, col)
                && match held {
                    None => true,
                    Some(_) => self.files.get(&anchor).is_some_and(|f| inside(f, row, col)),
                }
        })
    }

    /// The anchors whose boxes meet an area of a sheet, each run's in order.
    pub(crate) fn meeting(&self, sheet: usize, area: A1Range) -> impl Iterator<Item = CellId> + '_ {
        self.boxes
            .get(sheet)
            .into_iter()
            .flat_map(move |b| b.meeting(area))
            .flat_map(move |(_, k)| self.runs[k as usize].meeting(area.top, area.bottom))
    }
}

/// Each sheet's runs' boxes, tagged by the run's place.
fn boxes_of(runs: &[Run]) -> Vec<Rects> {
    let mut boxes: Vec<Rects> = Vec::new();
    for (k, r) in runs.iter().enumerate() {
        let s = r.sheet as usize;
        if boxes.len() <= s {
            boxes.resize_with(s + 1, Rects::new);
        }
        boxes[s].insert(r.cells(), k as u32);
    }
    boxes
}

/// The runs labelled where their anchors' boxes may meet ([`MEETS`]), each
/// split where its anchors are labelled and not.
///
/// Two boxes meeting are each labelled, the order Excel computes them in
/// deciding which spills; but where one anchor stands in the other's box,
/// which no order decides: the other's extent reaches a cell of the one's
/// only through its anchor, a formula, which blocks it. Every box grows down
/// and right of its anchor, so of two anchors of one column, the higher one's
/// box meeting the lower one's holds its anchor; and of two of different
/// columns, the left one's box meets the right one's without holding its
/// anchor only where the left one stands below the right one by at most the
/// right one's reach: of two runs, the anchors of each so labelled are rows
/// of it in a row, and only the part of a run's box right of its column is
/// searched, none for a box one column wide, so that a column of nested
/// boxes, a lookup's down a column, costs no search.
fn labelled(runs: Vec<Run>) -> Vec<Run> {
    let boxes = boxes_of(&runs);
    // Each run's rows labelled: the run's place, the first and the last.
    let mut met: Vec<(u32, u32, u32)> = Vec::new();
    for (k, r) in runs.iter().enumerate() {
        let cells = r.cells();
        if cells.right == r.col {
            continue;
        }
        let right = A1Range {
            left: r.col + 1,
            ..cells
        };
        for (_, other) in boxes[r.sheet as usize].meeting(right) {
            let o = &runs[other as usize];
            if o.col <= r.col || o.down == 0 {
                continue;
            }
            // The left run's anchor at row i and the right one's at row j
            // cross where j < i <= j + the right one's reach.
            let first = r.top.max(o.top + 1);
            let last = r.bottom.min(o.bottom.saturating_add(o.down));
            if first > last {
                continue;
            }
            met.push((k as u32, first, last));
            let theirs = o.top.max(r.top.saturating_sub(o.down));
            met.push((other, theirs, o.bottom.min(r.bottom - 1)));
        }
    }
    if met.is_empty() {
        return runs;
    }
    met.sort_unstable();
    let mut out: Vec<Run> = Vec::with_capacity(runs.len() + 2 * met.len());
    let mut m = 0;
    for (k, r) in runs.into_iter().enumerate() {
        let mut at = r.top;
        while m < met.len() && met[m].0 == k as u32 {
            let (_, first, mut last) = met[m];
            m += 1;
            while m < met.len() && met[m].0 == k as u32 && met[m].1 <= last + 1 {
                last = last.max(met[m].2);
                m += 1;
            }
            if first > at {
                out.push(Run {
                    top: at,
                    bottom: first - 1,
                    ..r
                });
            }
            out.push(Run {
                top: first,
                bottom: last,
                met: true,
                ..r
            });
            at = last + 1;
        }
        if at <= r.bottom {
            out.push(Run { top: at, ..r });
        }
    }
    out
}

/// Whether a rectangle holds a cell.
pub(crate) fn inside(r: &A1Range, row: u32, col: u32) -> bool {
    r.top <= row && row <= r.bottom && r.left <= col && col <= r.right
}

/// A grid as its plan reads a defined name and a text it cannot parse
/// (KERNEL.55): the formula cells inside an area its own and the anchor of
/// each spill whose box meets the area, so that a name standing for a cell a
/// spill may fill, empty or not, is read as its anchor's, as a door's plan
/// reads one (`graph::plan_placed`).
pub(crate) struct Spilled<'a, M: Model + ?Sized> {
    pub(crate) model: &'a M,
    pub(crate) spills: &'a Spills,
}

impl<M: Model + ?Sized> Model for Spilled<'_, M> {
    fn sheet_count(&self) -> usize {
        self.model.sheet_count()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.model.sheet_name(sheet)
    }

    fn find_sheet(&self, name: &str) -> Option<usize> {
        self.model.find_sheet(name)
    }

    fn defined_names(&self) -> &[(String, String)] {
        self.model.defined_names()
    }

    fn functions(&self) -> &[crate::calc::room::Compiled] {
        self.model.functions()
    }

    fn formula(&self, id: CellId) -> Option<Formula<'_>> {
        self.model.formula(id)
    }

    fn masters(&self, sheet: usize) -> Masters<'_> {
        self.model.masters(sheet)
    }

    fn each_formula(&self, sheet: usize, f: &mut dyn FnMut((u32, u32), Formula<'_>)) {
        self.model.each_formula(sheet, f);
    }

    fn formulas_in(&self, area: &Area, extent: Extent, f: &mut dyn FnMut(u32, u32)) {
        self.model.formulas_in(area, extent, f);
        let met = A1Range {
            top: area.top,
            left: area.left,
            bottom: area.bottom,
            right: area.right,
        };
        // An anchor inside the area is a formula cell of it already.
        for (_, row, col) in self.spills.meeting(area.sheet, met) {
            if !inside(&met, row, col) {
                f(row, col);
            }
        }
    }
}

/// A dynamic formula's anchor's saved extent: `Some(None)` for one whose file
/// saved none past its cell, `None` for a cell holding no dynamic formula.
fn file_of(model: &Workbook, (s, row, col): CellId) -> Option<Option<A1Range>> {
    match &model.sheets.get(s)?.cells.get(&(row, col))?.content {
        Content::DynamicFormula(_) => Some(None),
        Content::DynamicSpill { range, .. } if (range.top, range.left) == (row, col) => {
            Some((range.cells() > 1).then_some(*range))
        }
        Content::DynamicSpill { .. } => Some(None),
        _ => None,
    }
}

/// The anchors' bounds, each found once, a spill read through `#` its
/// anchor's, a ring of them none; and the anchors whose parse gives none.
struct Bounds<'a, 't> {
    text_of: &'a dyn Fn(CellId) -> Option<(&'t str, Option<A1Range>)>,
    group_of: &'a dyn Fn(CellId) -> Option<u64>,
    find_sheet: &'a dyn Fn(&str) -> Option<usize>,
    names: &'a Names,
    functions: &'a Functions,
    memo: HashMap<CellId, Option<Dims>>,
    groups: HashMap<u64, Option<Dims>>,
    visiting: HashSet<CellId>,
    unbounded: Vec<CellId>,
}

impl Bounds<'_, '_> {
    /// A dynamic formula's bound by its parse, a group's found once, else the
    /// extent its file saved; one value for a cell holding none, whose `#` is
    /// `#REF!`.
    fn dims(&mut self, id: CellId, lib: &Library) -> Option<Dims> {
        if let Some(d) = self.memo.get(&id) {
            return *d;
        }
        let Some((text, saved)) = (self.text_of)(id) else {
            return Some((1, 1));
        };
        let group = (self.group_of)(id);
        let by_parse = match group.and_then(|g| self.groups.get(&g).copied()) {
            Some(d) => d,
            None => {
                if !self.visiting.insert(id) {
                    return None;
                }
                let parsed = lambda::parsed(text, self.functions, lib).ok();
                let d = parsed.and_then(|e| self.of_parse(&e, id.0, lib, 0));
                self.visiting.remove(&id);
                if let Some(g) = group {
                    self.groups.insert(g, d);
                }
                d
            }
        };
        if by_parse.is_none() {
            self.unbounded.push(id);
        }
        let d = by_parse.or_else(|| saved.map(|f| (f.bottom - f.top + 1, f.right - f.left + 1)));
        if group.is_none() {
            self.memo.insert(id, d);
        }
        d
    }

    /// A parse's bound on a sheet: a spill read through `#` its anchor's,
    /// and a name its text's, a name read in a name's text no deeper than
    /// [`NAME_DEPTH`].
    fn of_parse(&mut self, e: &Expr, home: usize, lib: &Library, depth: u32) -> Option<Dims> {
        let find_sheet = self.find_sheet;
        bound(e, lib, &mut |asked| match asked {
            Asked::Anchor(r) => {
                let sheet = match &r.sheet {
                    Some(name) => find_sheet(name)?,
                    None => home,
                };
                self.dims((sheet, r.top, r.left), lib)
            }
            Asked::Name(name, sheet) => {
                if depth >= NAME_DEPTH {
                    return None;
                }
                let scope = match sheet {
                    Some(s) => Some(find_sheet(s)?),
                    None => None,
                };
                let text = self.names.resolve(name, scope, home)?;
                let named = formula::parse(&format!("={text}")).ok()?;
                self.of_parse(&named, home, lib, depth + 1)
            }
        })
    }
}

/// How deep a name read in a name's text is followed for a bound.
const NAME_DEPTH: u32 = 8;

/// What a bound asks of the grid it is planned in (KERNEL.55): the bound of
/// a spill read through `#`, its anchor's; and a defined name's, by what it
/// stands for.
pub(crate) enum Asked<'e> {
    Anchor(&'e RefExpr),
    Name(&'e str, Option<&'e str>),
}

/// The most rows and columns `e`'s answer can take, or `None` where its parse
/// gives none; `ask` gives an anchor's bound and a name's.
pub(crate) fn bound(
    e: &Expr,
    lib: &Library,
    ask: &mut dyn FnMut(Asked<'_>) -> Option<Dims>,
) -> Option<Dims> {
    match e {
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => {
            Some((1, 1))
        }
        Expr::Array(a) => Some((a.rows(), a.columns())),
        Expr::Ref(r) if r.last_sheet.is_none() => Some(dims(r)),
        // A span of sheets, which no answer holds as an array: one value,
        // `#REF!` where it reads many cells.
        Expr::Ref(_) => Some((1, 1)),
        Expr::Range(r) => r.reach.as_ref().map(dims),
        Expr::Name { name, sheet } => ask(Asked::Name(name, sheet.as_deref())),
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => bound(x, lib, ask),
        Expr::Binary(_, a, b) => Some(lift(bound(a, lib, ask)?, bound(b, lib, ask)?)),
        Expr::Call(name, args) => call(name, args, lib, ask),
        _ => None,
    }
}

/// A reference's rows and columns.
fn dims(r: &RefExpr) -> Dims {
    (r.bottom - r.top + 1, r.right - r.left + 1)
}

/// Two shapes lifted together, as a lift's places are ([`super::array::shape_of`]).
fn lift(a: Dims, b: Dims) -> Dims {
    (a.0.max(b.0), a.1.max(b.1))
}

/// A call's bound: the array functions by their shapes, the lookups and the
/// choices by what they pick from, and any other function of the library by
/// the lift of the arguments its parameters of one value take.
fn call(
    name: &str,
    args: &[Expr],
    lib: &Library,
    ask: &mut dyn FnMut(Asked<'_>) -> Option<Dims>,
) -> Option<Dims> {
    let mut of = |i: usize| -> Option<Dims> { bound(args.get(i)?, lib, ask) };
    let count = |i: usize| -> Option<u32> {
        match args.get(i)? {
            Expr::Number(n) if *n >= 1.0 && n.fract() == 0.0 && *n <= f64::from(u32::MAX) => {
                Some(*n as u32)
            }
            _ => None,
        }
    };
    match name {
        "SEQUENCE" => {
            let columns = match args.get(1) {
                None | Some(Expr::Empty) => 1,
                Some(_) => count(1)?,
            };
            Some((count(0)?, columns))
        }
        "MAKEARRAY" => Some((count(0)?, count(1)?)),
        "FILTER" | "SORT" | "SORTBY" | "UNIQUE" | "TAKE" | "DROP" | "ISFORMULA" | "FORMULATEXT" => {
            of(0)
        }
        "TRANSPOSE" => of(0).map(|(r, c)| (c, r)),
        // A reference's rows, and its columns, each a number.
        "ROW" => match args.first() {
            None => Some((1, 1)),
            Some(_) => of(0).map(|(r, _)| (r, 1)),
        },
        "COLUMN" => match args.first() {
            None => Some((1, 1)),
            Some(_) => of(0).map(|(_, c)| (1, c)),
        },
        "BYROW" => of(0).map(|(r, _)| (r, 1)),
        "BYCOL" => of(0).map(|(_, c)| (1, c)),
        "SCAN" => of(1),
        "MAP" if !args.is_empty() => {
            let mut shape = (1, 1);
            for i in 0..args.len() - 1 {
                shape = lift(shape, of(i)?);
            }
            Some(shape)
        }
        "TOCOL" => of(0).map(|(r, c)| (r.saturating_mul(c), 1)),
        "TOROW" => of(0).map(|(r, c)| (1, r.saturating_mul(c))),
        "HSTACK" | "VSTACK" => {
            let mut shape = (0u32, 0u32);
            for i in 0..args.len() {
                let (r, c) = of(i)?;
                shape = if name == "HSTACK" {
                    (shape.0.max(r), shape.1.saturating_add(c))
                } else {
                    (shape.0.saturating_add(r), shape.1.max(c))
                };
            }
            Some(shape)
        }
        "CHOOSECOLS" | "CHOOSEROWS" => {
            let base = of(0)?;
            let mut picked = 0u32;
            for i in 1..args.len() {
                let (r, c) = of(i)?;
                picked = picked.saturating_add(r.saturating_mul(c));
            }
            Some(if name == "CHOOSECOLS" {
                (base.0, picked)
            } else {
                (picked, base.1)
            })
        }
        "WRAPROWS" | "WRAPCOLS" => {
            let (r, c) = of(0)?;
            let n = count(1)?;
            let lines = r.saturating_mul(c).div_ceil(n);
            Some(if name == "WRAPROWS" {
                (lines, n)
            } else {
                (n, lines)
            })
        }
        "EXPAND" => {
            let base = of(0)?;
            let columns = match args.get(2) {
                None | Some(Expr::Empty) => base.1,
                Some(_) => count(2)?,
            };
            Some((count(1)?.max(base.0), columns.max(base.1)))
        }
        "MMULT" => Some((of(0)?.0, of(1)?.1)),
        // A pick of its reference or its array, no larger, where the row and
        // the column are one value each: one row where the row is a number
        // past 0 written in, one column where the column is.
        "INDEX" => {
            let base = of(0)?;
            for i in 1..args.len().min(3) {
                if of(i)? != (1, 1) {
                    return None;
                }
            }
            // A number past 0 picks one, as does a negative one, a text or a
            // match's place, each answering one value or an error.
            let picked = |i: usize, all: u32| match args.get(i) {
                Some(Expr::Number(n)) if *n >= 1.0 => 1,
                Some(Expr::Neg(x)) if matches!(&**x, Expr::Number(n) if *n >= 1.0) => 1,
                Some(Expr::Text(_)) => 1,
                Some(Expr::Call(f, _)) if f == "MATCH" || f == "XMATCH" => 1,
                _ => all,
            };
            Some((picked(1, base.0), picked(2, base.1)))
        }
        // A slice of the array returned, where one value is sought: a row of
        // it where the values sought down stand in a column, a column of it
        // where they stand in a row.
        "XLOOKUP" => {
            // Many values sought: one answer each, as the lift takes it.
            let sought = of(0)?;
            if sought != (1, 1) {
                return Some(sought);
            }
            let (rows, columns) = of(1)?;
            let returned = of(2)?;
            Some(match (rows, columns) {
                (r, 1) if r > 1 => (1, returned.1),
                (1, c) if c > 1 => (returned.0, 1),
                _ => returned,
            })
        }
        "IF" | "IFS" | "SWITCH" | "CHOOSE" | "IFERROR" | "IFNA" => {
            let mut shape = (1, 1);
            for i in 0..args.len() {
                shape = lift(shape, of(i)?);
            }
            Some(shape)
        }
        ANCHORARRAY => match args {
            [Expr::Ref(r)] => ask(Asked::Anchor(r)),
            _ => None,
        },
        // The forms of one value whatever they are given: `@`, and what tells
        // a reference, a value's kind or a parameter left out.
        "SINGLE" | "ISREF" | "TYPE" | "ISOMITTED" => Some((1, 1)),
        // The forms lifted over every argument, as a dynamic formula lifts
        // them (`ADDRESS` hands its call on to a strict function, KERNEL.60).
        "ISEVEN" | "ISODD" | "ADDRESS" => {
            let mut shape = (1, 1);
            for i in 0..args.len() {
                shape = lift(shape, of(i)?);
            }
            Some(shape)
        }
        _ => {
            // A function of values: lifted over what its parameters of one
            // value are given, as a dynamic formula lifts it; a range or an
            // array handed whole makes one value.
            lib.strict(name)?;
            let classes = lib.classes(name)?;
            let mut shape = (1, 1);
            for (i, a) in args.iter().enumerate() {
                if matches!(classes(i), Param::Value | Param::Scalar | Param::Forced) {
                    shape = lift(shape, bound(a, lib, ask)?);
                }
            }
            Some(shape)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::formula::parse;
    use crate::calc::library;

    fn of(text: &str) -> Option<Dims> {
        let lib = library::language();
        let e = parse(text).unwrap();
        // A spill read through `#` bounded three by two; a name the grid
        // does not hold, none.
        bound(&e, &lib, &mut |asked| match asked {
            Asked::Anchor(_) => Some((3, 2)),
            Asked::Name(..) => None,
        })
    }

    #[test]
    fn each_spill_of_the_probe_has_the_bound_its_text_gives() {
        // KERNEL.55's scoping, its third measure: the eighteen spills the
        // fixtures label each bounded by its text, seventeen exactly and
        // FILTER's at most its array.
        for (text, want) in [
            ("=SEQUENCE(3,2)", Some((3, 2))),
            ("=SEQUENCE(4)", Some((4, 1))),
            ("=SEQUENCE(1,3)", Some((1, 3))),
            ("=FILTER(Data!A1:A6,Data!A1:A6>Data!K1)", Some((6, 1))),
            ("=ANCHORARRAY(A1)*10", Some((3, 2))),
            ("=SORT(ANCHORARRAY(D1),1,-1)", Some((3, 2))),
            ("=Data!D1:D5", Some((5, 1))),
            ("=1/Data!C1:C3", Some((3, 1))),
            ("=Data!B1:B3&\"!\"", Some((3, 1))),
            ("=Data!A1:A3>2", Some((3, 1))),
            ("=SEQUENCE(1)", Some((1, 1))),
            ("=MAP(Data!A1:A3,LAMBDA(c,c*2))", Some((3, 1))),
            ("=BYROW(Data!E1:G2,LAMBDA(r,SUM(r)))", Some((2, 1))),
            ("=MAKEARRAY(2,3,LAMBDA(r,c,r*c))", Some((2, 3))),
            ("=INDEX(Data!E1:G2,0,2)", Some((2, 1))),
            ("=INDEX(Data!E1:G2,2,0)", Some((1, 3))),
            ("=INDEX(ANCHORARRAY(A1),3,2)", Some((1, 1))),
            ("=INDEX(Data!E1:G2,Data!X1)", Some((2, 3))),
            ("=XLOOKUP(\"b\",Data!H1:J1,Data!H2:J2)", Some((1, 1))),
            ("=XLOOKUP(2,Data!A1:A6,Data!E1:G6)", Some((1, 3))),
            ("=INDEX(Data!A1:A8,MATCH(40,Data!A1:A8,0))", Some((1, 1))),
            ("=INDEX(Data!A1:B8,-1,1)", Some((1, 1))),
            ("=INDEX(Data!A1:B8,\"x\")", Some((1, 2))),
            ("=ROW(Data!A2:A5)", Some((4, 1))),
            ("=COLUMN(Data!A2:L2)", Some((1, 12))),
            ("=ROW()", Some((1, 1))),
            ("=TRANSPOSE(Data!A1:A3)", Some((1, 3))),
            ("={1,2;3,4}+1", Some((2, 2))),
            ("=SUM(Data!A1:A6)", Some((1, 1))),
            ("=LEN(Data!B1:B3)", Some((3, 1))),
            ("=HSTACK(A1:A2,B1:C4)", Some((4, 3))),
            ("=TOCOL(A1:B3)", Some((6, 1))),
            // The forms of one value, those lifted over every argument, a span
            // of sheets, and XLOOKUP lifted over the values sought (an engine
            // refuses what gives no bound, its slice).
            ("=SINGLE(A1:A3)*2", Some((1, 1))),
            ("=TYPE(A1:A3)+ISREF(B1:B2)", Some((1, 1))),
            ("=ISEVEN(A1:A3)", Some((3, 1))),
            ("=ADDRESS(A1:A2,B1:C1)", Some((2, 2))),
            ("=Data:Other!A1:A3+1", Some((1, 1))),
            ("=XLOOKUP(A1:A3,B1:B9,C1:D9)", Some((3, 1))),
        ] {
            assert_eq!(of(text), want, "{text}");
        }
        // No bound: a count read from a cell, a name the grid does not hold,
        // a function the library lacks.
        for text in [
            "=SEQUENCE(A1)",
            "=Rate*2",
            "=NOSUCH(A1:A3)",
            "=INDEX(A1:B3,C1:C2)",
        ] {
            assert_eq!(of(text), None, "{text}");
        }
    }

    #[test]
    fn each_spill_s_box_holds_its_bound_and_its_saved_extent() {
        // A formula of one value has no spill; a bound and an extent its file
        // saved make one box; past a door's plan, labelled; two whose boxes
        // may meet labelled, and one whose anchor stands in another's box not.
        use crate::sheet::{parse_a1_range, Cell};
        let mut wb = Workbook::new();
        let s = wb.ensure_sheet("S");
        let mut put = |addr: &str, content: Content| {
            let r = parse_a1_range(addr).unwrap();
            wb.sheets[s].set(r.top, r.left, Cell { content, style: 0 });
        };
        let dynamic = |t: &str| Content::DynamicFormula(t.to_string());
        put("A1", dynamic("SEQUENCE(1)"));
        put("B1", dynamic("SEQUENCE(3)"));
        let range = parse_a1_range("C1:D2").unwrap();
        put(
            "C1",
            Content::DynamicSpill {
                text: "SEQUENCE(3)".to_string(),
                range,
            },
        );
        // Below it, one bounded alike whose file saved nothing: a run of its
        // own, its box its bound's.
        put("C2", dynamic("SEQUENCE(3)"));
        // A value its file saved, and one in a box its file saved none of.
        put("D2", Content::Number(4.0));
        put("B3", Content::Number(4.0));
        // A file saving more rows than its bound now answers: its box the
        // file's, so that the cells its last spill filled are its own.
        put(
            "H1",
            Content::DynamicSpill {
                text: "SEQUENCE(2)".to_string(),
                range: parse_a1_range("H1:H4").unwrap(),
            },
        );
        put("F1", dynamic("FILTER(A:A,A:A>0)"));
        put("K1", dynamic("SEQUENCE(4)"));
        put("J3", dynamic("SEQUENCE(1,3)"));
        put("P1", dynamic("SEQUENCE(6)"));
        put("P3", dynamic("SEQUENCE(2)"));
        let texts: HashMap<CellId, String> = wb.sheets[s]
            .cells
            .iter()
            .filter_map(|(&(row, col), c)| match &c.content {
                Content::DynamicFormula(t) | Content::DynamicSpill { text: t, .. } => {
                    Some(((s, row, col), t.clone()))
                }
                _ => None,
            })
            .collect();
        let lib = crate::calc::library::language();
        let spills = Spills::of(
            &dynamics_of(&wb, &texts),
            &|name| wb.find_sheet(name),
            &Names::of(&wb),
            &Functions::default(),
            &lib,
        );
        let at = |addr: &str| {
            let r = parse_a1_range(addr).unwrap();
            spills.get((s, r.top, r.left))
        };
        assert_eq!(at("A1"), None, "one value, no spill");
        let b1 = at("B1").unwrap();
        assert_eq!(
            (b1.dims, b1.cells.text(), b1.label),
            ((3, 1), "B1:B3".to_string(), None)
        );
        let c1 = at("C1").unwrap();
        assert_eq!((c1.dims, c1.cells.text()), ((3, 1), "C1:D3".to_string()));
        assert_eq!(c1.file, Some(range));
        let c2 = at("C2").unwrap();
        assert_eq!((c2.cells.text(), c2.file), ("C2:C4".to_string(), None));
        let h1 = at("H1").unwrap();
        assert_eq!((h1.dims, h1.cells.text()), ((2, 1), "H1:H4".to_string()));
        assert_eq!(at("F1").unwrap().label, Some(NO_PLAN), "a whole column");
        assert_eq!(at("K1").unwrap().label, Some(MEETS));
        assert_eq!(at("J3").unwrap().label, Some(MEETS));
        assert_eq!(at("P1").unwrap().label, None, "P3 stands in its box");
        assert_eq!(at("P3").unwrap().label, None);
        // A cell of a box owned by its anchor, and by both where an anchor
        // stands in another's box; never the anchor's own cell.
        let owners = |addr: &str| {
            let r = parse_a1_range(addr).unwrap();
            let mut o: Vec<CellId> = spills.owners(&wb, (s, r.top, r.left)).collect();
            o.sort_unstable();
            o
        };
        assert_eq!(owners("B2"), [(s, 1, 2)]);
        assert_eq!(
            owners("B3"),
            [],
            "a value of its own, which blocks the spill"
        );
        assert_eq!(owners("D2"), [(s, 1, 3)], "a value its anchor's file saved");
        assert_eq!(owners("H4"), [(s, 1, 8)], "a cell its last spill filled");
        assert_eq!(owners("P4"), [(s, 1, 16), (s, 3, 16)]);
        assert_eq!(owners("B1"), []);
        assert_eq!(owners("C1"), [], "an anchor its own saved extent holds");
        assert_eq!(owners("Z9"), []);
    }

    #[test]
    fn the_boxes_labelled_meeting_are_the_pairs_whose_anchors_neither_holds() {
        // The search for boxes meeting, the part of a run's box right of its
        // column alone, and the runs split where they are labelled, held to
        // every pair compared: random anchors of random bounds on one sheet,
        // alone and in columns of one bound, each pair whose boxes meet with
        // neither anchor in the other's box labelled, and no other; and every
        // box a cell, an area and an anchor meets, held to the boxes listed.
        use crate::sheet::Cell;
        let lib = crate::calc::library::language();
        let mut seed = 0x9E37_79B9_7F4A_7C15_u64;
        let mut next = |n: u64| {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            seed % n
        };
        for _ in 0..300 {
            let mut wb = Workbook::new();
            let s = wb.ensure_sheet("S");
            for _ in 0..2 + next(10) {
                let (row, col) = (1 + next(12) as u32, 1 + next(8) as u32);
                let text = format!("SEQUENCE({},{})", 1 + next(5), 1 + next(4));
                // A column of one bound, a shape filled down, or one alone.
                let rows = if next(2) == 0 { 1 + next(6) as u32 } else { 1 };
                for row in row..row + rows {
                    let content = Content::DynamicFormula(text.clone());
                    wb.sheets[s].set(row, col, Cell { content, style: 0 });
                }
            }
            let texts: HashMap<CellId, String> = wb.sheets[s]
                .cells
                .iter()
                .filter_map(|(&(row, col), cell)| match &cell.content {
                    Content::DynamicFormula(t) => Some(((s, row, col), t.clone())),
                    _ => None,
                })
                .collect();
            let spills = Spills::of(
                &dynamics_of(&wb, &texts),
                &|name| wb.find_sheet(name),
                &Names::of(&wb),
                &Functions::default(),
                &lib,
            );
            // Each anchor's box as its text bounds it, a SEQUENCE's rows and
            // columns, where it holds more than its own cell: every one listed,
            // and none else.
            let mut boxes: Vec<(CellId, A1Range)> = wb.sheets[s]
                .cells
                .iter()
                .filter_map(|(&(row, col), cell)| {
                    let Content::DynamicFormula(text) = &cell.content else {
                        return None;
                    };
                    let (r, c) = text
                        .trim_start_matches("SEQUENCE(")
                        .trim_end_matches(')')
                        .split_once(',')?;
                    let (r, c) = (r.parse::<u32>().ok()?, c.parse::<u32>().ok()?);
                    let held = A1Range {
                        top: row,
                        left: col,
                        bottom: row + r - 1,
                        right: col + c - 1,
                    };
                    (held.cells() > 1).then_some(((s, row, col), held))
                })
                .collect();
            boxes.sort_unstable_by_key(|b| b.0);
            let mut listed: Vec<(CellId, A1Range)> = spills.boxes().collect();
            listed.sort_unstable_by_key(|b| b.0);
            assert_eq!(listed, boxes);
            for &(a, box_a) in &boxes {
                let crossing = boxes.iter().any(|&(b, box_b)| {
                    b != a
                        && crate::grid::overlap(&box_a, &box_b).is_some()
                        && !inside(&box_a, b.1, b.2)
                        && !inside(&box_b, a.1, a.2)
                });
                let spill = spills.get(a).expect("an anchor's spill");
                assert_eq!(
                    spill.label == Some(MEETS),
                    crossing,
                    "{a:?} among {boxes:?}"
                );
                assert_eq!((spill.cells, spills.box_of(a)), (box_a, Some(box_a)));
                let dims = (box_a.bottom - box_a.top + 1, box_a.right - box_a.left + 1);
                assert_eq!(spill.dims, dims, "{a:?}");
            }
            // Every anchor in a box listed, and none else.
            let listed: Vec<CellId> = boxes.iter().map(|b| b.0).collect();
            for row in 1..=13 {
                for col in 1..=9 {
                    let in_box = boxes.iter().filter(|b| inside(&b.1, row, col));
                    assert_eq!(
                        spills.covers((s, row, col)),
                        in_box.clone().next().is_some()
                    );
                    let mut owners: Vec<CellId> = in_box
                        .map(|b| b.0)
                        .filter(|&b| b != (s, row, col))
                        .filter(|_| !wb.sheets[s].cells.contains_key(&(row, col)))
                        .collect();
                    owners.sort_unstable();
                    let mut got: Vec<CellId> = spills.owners(&wb, (s, row, col)).collect();
                    got.sort_unstable();
                    assert_eq!(got, owners, "{row} {col}");
                    let at = spills.get((s, row, col)).is_some();
                    assert_eq!(at, listed.contains(&(s, row, col)), "{row} {col}");
                }
            }
            for _ in 0..20 {
                let (top, left) = (1 + next(14) as u32, 1 + next(10) as u32);
                let area = A1Range {
                    top,
                    left,
                    bottom: top + next(4) as u32,
                    right: left + next(4) as u32,
                };
                let mut want: Vec<CellId> = boxes
                    .iter()
                    .filter(|b| crate::grid::overlap(&b.1, &area).is_some())
                    .map(|b| b.0)
                    .collect();
                want.sort_unstable();
                let mut got: Vec<CellId> = spills.meeting(s, area).collect();
                got.sort_unstable();
                assert_eq!(got, want, "{area:?}");
            }
        }
    }
}
