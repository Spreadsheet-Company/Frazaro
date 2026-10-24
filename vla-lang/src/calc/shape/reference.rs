//! The plan of cells, kept as the tests' reference (KERNEL.49's decision 7,
//! its fourth slice): every formula cell placed by the plan's placement,
//! each cell's edges found at it, Tarjan's components over the cells, and
//! the order, its cycles, runs, levels and stretches, as the plan found them
//! before it planned by nodes, together with each cell's position and the
//! first refusal. The plan by blocks is held to it.

use crate::fx::FxHashMap;
use crate::grid::{Formula, Model};
use crate::messages::{raise, Refusal};
use crate::refers;

use super::super::eval::unread_name;
use super::super::formula::{self, Read, RefExpr};
use super::super::graph::{self, areas_found, dependencies_in, Area, CellId, Extent, Names};
use super::{hub_at, spell_at, Shapes, RUN};

/// The plan of cells.
#[derive(Debug, PartialEq)]
pub(crate) struct Reference {
    pub(crate) order: Vec<CellId>,
    pub(crate) cycles: Vec<Vec<CellId>>,
    pub(crate) runs: Vec<(usize, usize, usize)>,
    pub(crate) level_ends: Vec<u32>,
    /// Each shape's stretches, by shape.
    pub(crate) stretches: Vec<Vec<(u32, u32)>>,
    /// Each placed cell and its shape, sorted.
    pub(crate) places: Vec<(CellId, usize)>,
    pub(crate) hubs: usize,
}

impl Reference {
    /// The plan of cells of a grid placed as `shapes` placed it.
    pub(crate) fn of<M: Model + ?Sized>(
        shapes: &Shapes,
        model: &M,
        names: &Names,
        extents: &[Extent],
    ) -> Reference {
        // Every placed cell, its shape and its offset.
        let mut places: Vec<(CellId, usize)> = Vec::new();
        for si in 0..model.sheet_count() {
            model.each_formula(si, &mut |(row, col), f| {
                let shape = match f {
                    Formula::Fill(index) => shapes.by_shared.get(&(si, index)).copied(),
                    Formula::Own(..) => shapes
                        .located
                        .alone
                        .get(si)
                        .and_then(|m| m.get(&(row, col)))
                        .map(|a| a.shape as usize),
                };
                if let Some(s) = shape {
                    places.push(((si, row, col), s));
                }
            });
        }
        places.sort_unstable();
        let ids: Vec<CellId> = places.iter().map(|p| p.0).collect();
        let index = Positions::of(&ids, model.sheet_count());
        let mut holds = vec![false; model.sheet_count()];
        for id in &ids {
            holds[id.0] = true;
        }
        let vertices = Vertices {
            model,
            extents,
            index: &index,
            holds: &holds,
        };
        let reads: Vec<ShapeReads> = shapes
            .shapes
            .iter()
            .map(|shape| vertices.reads_of(shape))
            .collect();
        let n = ids.len();
        let mut offsets: Vec<usize> = Vec::with_capacity(n + 1);
        let mut targets: Vec<u32> = Vec::with_capacity(n * 8);
        offsets.push(0);
        let mut shape_of: Vec<usize> = Vec::with_capacity(n);
        let mut deps: Vec<CellId> = Vec::new();
        let mut edges: Vec<u32> = Vec::new();
        let mut held: FxHashMap<(usize, usize), (Vec<u32>, Option<u32>)> = FxHashMap::default();
        let mut hubs: Vec<(usize, usize)> = Vec::new();
        for &(id, s) in &places {
            edges.clear();
            deps.clear();
            shape_of.push(s);
            let shape = &shapes.shapes[s];
            let (d_row, d_col) = (
                i64::from(id.1) - i64::from(shape.anchor.1),
                i64::from(id.2) - i64::from(shape.anchor.2),
            );
            let read = !shapes.unread.contains_key(&id);
            match &shape.expr {
                Ok(e) if read => {
                    let r = &reads[s];
                    if r.own.is_empty() && r.fixed.len() < hub_at() {
                        edges.extend_from_slice(&r.fixed);
                    } else {
                        let key = (s, if r.own.is_empty() { usize::MAX } else { id.0 });
                        let (fixed, hub) = held.entry(key).or_insert_with(|| {
                            let mut fixed = r.fixed.clone();
                            for re in &r.own {
                                let found = |n: &str| model.find_sheet(n);
                                if let Ok(areas) = areas_found(found, id.0, re, 0, 0) {
                                    for a in &areas {
                                        vertices.formulas_in(a, &mut fixed);
                                    }
                                }
                            }
                            fixed.sort_unstable();
                            fixed.dedup();
                            let hub = (fixed.len() >= hub_at()).then(|| {
                                hubs.push(key);
                                (n + hubs.len() - 1) as u32
                            });
                            (fixed, hub)
                        });
                        match hub {
                            Some(hub) => edges.push(*hub),
                            None => edges.extend_from_slice(fixed),
                        }
                    }
                    for re in &r.moving {
                        let found = |n: &str| model.find_sheet(n);
                        let Ok(areas) = areas_found(found, id.0, re, d_row, d_col) else {
                            continue;
                        };
                        for a in &areas {
                            vertices.formulas_in(a, &mut edges);
                        }
                    }
                    if r.named {
                        e.each_read(&mut |r| {
                            if let Read::Name { name, sheet } = r {
                                let scope = sheet.and_then(|s| model.find_sheet(s));
                                if let Some(t) = names.resolve(name, scope, id.0) {
                                    dependencies_in(
                                        model,
                                        names,
                                        extents,
                                        scope.unwrap_or(id.0),
                                        t,
                                        &mut deps,
                                    );
                                }
                            }
                        });
                    }
                }
                _ => {
                    let moved = refers::shift_a1_references(&shape.text, d_row, d_col);
                    dependencies_in(model, names, extents, id.0, &moved, &mut deps);
                }
            }
            for d in &deps {
                if let Some(v) = index.get(d.0, d.1, d.2) {
                    edges.push(v as u32);
                }
            }
            edges.sort_unstable();
            edges.dedup();
            targets.extend_from_slice(&edges);
            offsets.push(targets.len());
        }
        for key in &hubs {
            targets.extend_from_slice(&held[key].0);
            offsets.push(targets.len());
        }
        let (members, starts) = graph::tarjan_flat(&offsets, &targets);
        super::HUBS_MADE.with(|h| h.set(h.get() + hubs.len()));
        let mut r = Reference {
            order: Vec::new(),
            cycles: Vec::new(),
            runs: Vec::new(),
            level_ends: Vec::new(),
            stretches: vec![Vec::new(); shapes.shapes.len()],
            places,
            hubs: hubs.len(),
        };
        r.order_and_runs(
            shapes,
            &ids,
            (&offsets, &targets),
            &shape_of,
            (&members, &starts),
        );
        r
    }

    fn order_and_runs(
        &mut self,
        shapes: &Shapes,
        ids: &[CellId],
        (offsets, targets): (&[usize], &[u32]),
        shape_of: &[usize],
        (members, starts): (&[u32], &[usize]),
    ) {
        let adj = |v: usize| &targets[offsets[v]..offsets[v + 1]];
        let n = ids.len();
        let all = offsets.len() - 1;
        let mut component = vec![0; all];
        let mut level: Vec<usize> = vec![0; all];
        let mut alone = vec![false; n];
        for (c, ends) in starts.windows(2).enumerate() {
            let comp = &members[ends[0]..ends[1]];
            for &v in comp {
                component[v as usize] = c;
            }
            let deepest = comp
                .iter()
                .flat_map(|&v| adj(v as usize))
                .filter(|&&u| component[u as usize] != c)
                .map(|&u| level[u as usize] + 1)
                .max()
                .unwrap_or(0);
            if comp.iter().all(|&v| v as usize >= n) {
                for &v in comp {
                    level[v as usize] = deepest.saturating_sub(1);
                }
                continue;
            }
            for &v in comp {
                level[v as usize] = deepest;
            }
            if comp.len() == 1 && !adj(comp[0] as usize).contains(&comp[0]) {
                alone[comp[0] as usize] = true;
            } else {
                let mut cells: Vec<CellId> = comp
                    .iter()
                    .filter(|&&i| (i as usize) < n)
                    .map(|&i| ids[i as usize])
                    .collect();
                cells.sort_unstable();
                self.cycles.push(cells);
            }
        }
        self.cycles.sort();
        let cells: Vec<usize> = (0..n).filter(|&v| alone[v]).collect();
        let (cells, _) = by_key(&cells, shapes.shapes.len(), |v| shape_of[v]);
        let levels = level[..n].iter().max().map_or(0, |deepest| deepest + 1);
        let (order, ends) = by_key(&cells, levels, |v| level[v]);
        self.level_ends = ends;
        let mut joined = vec![usize::MAX; n];
        let mut run = (0usize, 0usize, usize::MAX);
        let mut stretches: Vec<(u32, u32, u32)> = Vec::new();
        for v in order {
            let (id, shape) = (ids[v], shape_of[v]);
            let pos = self.order.len();
            let joins = shape == run.2
                && id.0 == self.order[run.0].0
                && adj(v)
                    .iter()
                    .all(|&u| u as usize >= n || joined[u as usize] != run.0);
            if !joins {
                self.keep(shapes, run, &mut stretches);
                run = (pos, pos, shape);
            }
            joined[v] = run.0;
            run.1 = pos + 1;
            self.order.push(id);
        }
        self.keep(shapes, run, &mut stretches);
        for (s, start, end) in stretches {
            self.stretches[s as usize].push((start, end));
        }
    }

    fn keep(
        &mut self,
        shapes: &Shapes,
        run: (usize, usize, usize),
        stretches: &mut Vec<(u32, u32, u32)>,
    ) {
        let Some(shape) = shapes.shapes.get(run.2) else {
            return;
        };
        let (s, start, end) = (run.2 as u32, run.0 as u32, run.1 as u32);
        match stretches.last_mut() {
            Some(last) if last.0 == s && last.2 == start => last.2 = end,
            _ => stretches.push((s, start, end)),
        }
        if run.1 - run.0 >= RUN && !shape.risky && shape.expr.is_ok() {
            self.runs.push(run);
        }
    }

    /// Every cell's position in the order, by cell.
    pub(crate) fn positions(&self) -> FxHashMap<CellId, usize> {
        self.order
            .iter()
            .enumerate()
            .map(|(p, &id)| (id, p))
            .collect()
    }

    /// The first refusal as the plan of cells found it, every placed cell
    /// asked in order.
    pub(crate) fn first_refusal<M: Model + ?Sized>(
        &self,
        shapes: &Shapes,
        model: &M,
        only: Option<&[CellId]>,
    ) -> Option<(CellId, Refusal)> {
        if let Some(cycle) = self.cycles.first() {
            let cells: Vec<String> = cycle.iter().map(|id| spell_at(model, *id)).collect();
            return Some((
                cycle[0],
                raise("calc-cycle", &[("cells", &cells.join(", "))]),
            ));
        }
        let shape_of: FxHashMap<CellId, usize> = self.places.iter().copied().collect();
        let mut ids: Vec<CellId> = match only {
            Some(cells) => cells
                .iter()
                .copied()
                .filter(|id| shape_of.contains_key(id))
                .collect(),
            None => self.places.iter().map(|p| p.0).collect(),
        };
        let mut named: FxHashMap<(usize, usize), String> = FxHashMap::default();
        let names = (!model.defined_names().is_empty()).then(|| Names::of_model(model));
        if let Some(names) = names.filter(Names::any_unread) {
            let mut asked: FxHashMap<(usize, usize), ()> = FxHashMap::default();
            for id in &ids {
                let s = shape_of[id];
                if asked.insert((s, id.0), ()).is_some() {
                    continue;
                }
                if let Ok(e) = &shapes.shapes[s].expr {
                    if let Some(construct) = unread_name(model, &names, e, id.0) {
                        named.insert((s, id.0), construct);
                    }
                }
            }
        }
        ids.sort_unstable();
        for id in &ids {
            let shape = &shapes.shapes[shape_of[id]];
            if let Some(name) = shape.missing.first() {
                return Some((
                    *id,
                    raise(
                        "calc-function-not-computed",
                        &[("cell", &spell_at(model, *id)), ("function", name)],
                    ),
                ));
            }
        }
        for id in &ids {
            let s = shape_of[id];
            let shape = &shapes.shapes[s];
            let (d_row, d_col) = (
                i64::from(id.1) - i64::from(shape.anchor.1),
                i64::from(id.2) - i64::from(shape.anchor.2),
            );
            let label = match (&shape.expr, shapes.unread.get(id).cloned()) {
                (_, Some(unread)) => Some(unread),
                (Err(construct), None) if d_row == 0 && d_col == 0 => Some(construct.clone()),
                (Err(construct), None) => {
                    let moved = refers::shift_a1_references(&shape.text, d_row, d_col);
                    Some(
                        formula::parse(&moved)
                            .err()
                            .unwrap_or_else(|| construct.clone()),
                    )
                }
                (Ok(_), None) => named.get(&(s, id.0)).cloned(),
            };
            if let Some(label) = label {
                return Some((*id, formula::unread_refusal(&spell_at(model, *id), &label)));
            }
        }
        None
    }
}

thread_local! {
    /// Whether every plan made on this thread is held to the plan of cells:
    /// on, unless a test turns it off for a grid too large to plan twice.
    static HELD: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}

/// Every plan made on this thread held to the plan of cells or not, until
/// the guard drops.
pub(crate) struct Held(bool);

impl Drop for Held {
    fn drop(&mut self) {
        HELD.with(|h| h.set(self.0));
    }
}

/// Plans made on this thread not held to the plan of cells, until the guard
/// drops.
pub(crate) fn unheld() -> Held {
    Held(HELD.with(|h| h.replace(false)))
}

/// The plan by blocks held to the plan of cells, part for part (KERNEL.49's
/// decision 7): the order cell for cell, its runs, levels, cycles and
/// stretches, every placed cell's shape and every position, each shape's
/// cells found inside rectangles as the reach finds them, and the first
/// refusal. Asked at the end of every plan a test makes.
pub(crate) fn hold<M: Model + ?Sized>(
    shapes: &Shapes,
    model: &M,
    names: &Names,
    extents: &[Extent],
) {
    if !HELD.with(std::cell::Cell::get) {
        return;
    }
    let r = Reference::of(shapes, model, names, extents);
    let order: Vec<CellId> = shapes.order().collect();
    if order != r.order {
        let at = order
            .iter()
            .zip(&r.order)
            .position(|(a, b)| a != b)
            .unwrap_or(order.len().min(r.order.len()));
        // Each differing cell's level in either plan, and its text.
        let level_of = |ends: &[u32], p: usize| ends.partition_point(|&e| e as usize <= p);
        let pos_ref = r.positions();
        let pos_new: FxHashMap<CellId, usize> =
            order.iter().enumerate().map(|(p, &id)| (id, p)).collect();
        let tell = |id: Option<&CellId>| {
            id.map(|id| {
                let s = shapes.place(*id).map(|p| p.shape);
                let text = s.map(|s| shapes.shapes[s].text.clone());
                (
                    *id,
                    s,
                    text,
                    pos_new.get(id).map(|&p| level_of(&shapes.level_ends, p)),
                    pos_ref.get(id).map(|&p| level_of(&r.level_ends, p)),
                )
            })
        };
        panic!(
            "the order by blocks differs from the order of cells at {at} of {} and {}: {:?} against {:?}
(cell, shape, text, level by blocks, level of cells): {:?} and {:?}
levels {:?} against {:?}",
            order.len(),
            r.order.len(),
            order.get(at),
            r.order.get(at),
            tell(order.get(at)),
            tell(r.order.get(at)),
            shapes.level_ends,
            r.level_ends
        );
    }
    let runs: Vec<(usize, usize, usize)> = shapes.run_spans().collect();
    assert_eq!(runs, r.runs, "the runs");
    assert_eq!(shapes.level_ends, r.level_ends, "the levels");
    assert_eq!(shapes.cycles, r.cycles, "the cycles");
    for (s, want) in r.stretches.iter().enumerate() {
        assert_eq!(
            shapes.stretches(s),
            want.as_slice(),
            "shape {s}'s stretches"
        );
    }
    assert_eq!(shapes.formulas(), r.places.len(), "the formulas placed");
    let positions = r.positions();
    for (p, &id) in r.order.iter().enumerate() {
        assert_eq!(shapes.position(id), Some(p), "{id:?}'s position");
    }
    let mut on: FxHashMap<(usize, usize), (u32, u32, u32, u32)> = FxHashMap::default();
    for &(id, s) in &r.places {
        let place = shapes.place(id).unwrap_or_else(|| panic!("{id:?} placed"));
        assert_eq!(place.shape, s, "{id:?}'s shape");
        assert!(shapes.holds(id), "{id:?} held");
        assert_eq!(
            shapes.position(id),
            positions.get(&id).copied(),
            "{id:?}'s position"
        );
        if let Some(&p) = positions.get(&id) {
            assert_eq!(shapes.at(p), (id, s), "the cell at {p}");
        }
        let b = on.entry((s, id.0)).or_insert((id.1, id.2, id.1, id.2));
        *b = (b.0.min(id.1), b.1.min(id.2), b.2.max(id.1), b.3.max(id.2));
    }
    // Each shape's cells inside its box, inside each cell's neighbourhood
    // and inside a box one row and column wider, as the reach asks them.
    let shape_of: FxHashMap<CellId, usize> = r.places.iter().copied().collect();
    let mut areas: Vec<(usize, usize, crate::sheet::A1Range)> = Vec::new();
    for (&(s, sheet), &(top, left, bottom, right)) in &on {
        let a = |t: u32, l: u32, b: u32, r: u32| crate::sheet::A1Range {
            top: t.max(1),
            left: l.max(1),
            bottom: b,
            right: r,
        };
        areas.push((s, sheet, a(top, left, bottom, right)));
        areas.push((
            s,
            sheet,
            a(
                top.saturating_sub(1),
                left.saturating_sub(1),
                bottom + 1,
                right + 1,
            ),
        ));
        areas.push((
            s,
            sheet,
            a(
                top + (bottom - top) / 2,
                left,
                bottom,
                left + (right - left) / 2,
            ),
        ));
    }
    for &(id, s) in r.places.iter().step_by(1 + r.places.len() / 200) {
        let a = crate::sheet::A1Range {
            top: id.1.saturating_sub(1).max(1),
            left: id.2.saturating_sub(1).max(1),
            bottom: id.1 + 1,
            right: id.2 + 1,
        };
        areas.push((s, id.0, a));
    }
    for (s, sheet, a) in areas {
        let mut got: Vec<usize> = Vec::new();
        shapes
            .located
            .each_in(&shapes.blocks, s, sheet, &a, &mut |p| got.push(p));
        got.sort_unstable();
        let mut want: Vec<usize> = Vec::new();
        for row in a.top..=a.bottom {
            for col in a.left..=a.right {
                let id = (sheet, row, col);
                if shape_of.get(&id) == Some(&s) {
                    if let Some(&p) = positions.get(&id) {
                        want.push(p);
                    }
                }
            }
        }
        want.sort_unstable();
        assert_eq!(got, want, "shape {s}'s cells on sheet {sheet} in {a:?}");
    }
    let spell = |x: Option<(CellId, crate::messages::Refusal)>| x.map(|(id, r)| (id, r.id, r.text));
    assert_eq!(
        spell(shapes.first_refusal_in(model, None)),
        spell(r.first_refusal(shapes, model, None)),
        "the first refusal"
    );
}

/// Cells ordered by a key below `keys`, a counting sort, with where each
/// key's cells end.
fn by_key(cells: &[usize], keys: usize, key: impl Fn(usize) -> usize) -> (Vec<usize>, Vec<u32>) {
    let mut start = vec![0u32; keys + 1];
    for &v in cells {
        start[key(v) + 1] += 1;
    }
    for k in 0..keys {
        start[k + 1] += start[k];
    }
    let mut out = vec![0; cells.len()];
    for &v in cells {
        let k = key(v);
        out[start[k] as usize] = v;
        start[k] += 1;
    }
    start.truncate(keys);
    (out, start)
}

struct Vertices<'a, M: Model + ?Sized> {
    model: &'a M,
    extents: &'a [Extent],
    index: &'a Positions,
    holds: &'a [bool],
}

struct ShapeReads<'e> {
    fixed: Vec<u32>,
    own: Vec<&'e RefExpr>,
    moving: Vec<&'e RefExpr>,
    named: bool,
}

impl<M: Model + ?Sized> Vertices<'_, M> {
    fn formulas_in(&self, a: &Area, out: &mut Vec<u32>) {
        if !self.holds.get(a.sheet).copied().unwrap_or(false) {
            return;
        }
        let rect = (a.top, a.left, a.bottom, a.right);
        if self
            .index
            .each_in(a.sheet, rect, &mut |v| out.push(v as u32))
        {
            return;
        }
        let extent = self.extents.get(a.sheet).copied().flatten();
        self.model.formulas_in(a, extent, &mut |row, col| {
            if let Some(v) = self.index.get(a.sheet, row, col) {
                out.push(v as u32);
            }
        });
    }

    fn reads_of<'e>(&self, shape: &'e super::Shape) -> ShapeReads<'e> {
        let mut reads = ShapeReads {
            fixed: Vec::new(),
            own: Vec::new(),
            moving: Vec::new(),
            named: false,
        };
        let Ok(e) = &shape.expr else {
            return reads;
        };
        let mut each: Vec<&RefExpr> = Vec::new();
        e.each_read(&mut |r| match r {
            Read::Ref(re) => {
                if !each.contains(&re) {
                    each.push(re);
                }
            }
            Read::Name { .. } => reads.named = true,
        });
        for re in each {
            let found_once = match (re.found, &re.last_sheet) {
                (Some((first, last)), None) if first == last => Some(usize::from(first)),
                _ => None,
            };
            if let Some(sheet) = found_once {
                if !self.holds.get(sheet).copied().unwrap_or(false) {
                    continue;
                }
                if every_dollar(re) {
                    let found = |n: &str| self.model.find_sheet(n);
                    if let Ok(areas) = areas_found(found, sheet, re, 0, 0) {
                        for a in &areas {
                            self.formulas_in(a, &mut reads.fixed);
                        }
                    }
                    continue;
                }
            }
            if re.sheet.is_none() && re.last_sheet.is_none() && every_dollar(re) {
                reads.own.push(re);
                continue;
            }
            reads.moving.push(re);
        }
        reads.fixed.sort_unstable();
        reads.fixed.dedup();
        reads
    }
}

fn every_dollar(r: &RefExpr) -> bool {
    let c = &r.corners;
    match r.shape {
        refers::Shape::Cell => c.row_abs1 && c.col_abs1,
        refers::Shape::Range => c.row_abs1 && c.col_abs1 && c.row_abs2 && c.col_abs2,
        refers::Shape::Column => c.col_abs1 && c.col_abs2,
        refers::Shape::Row => c.row_abs1 && c.row_abs2,
        refers::Shape::None => false,
    }
}

const NONE: u32 = u32::MAX;

/// Each formula cell's position among the cells: an array over a sheet's
/// formulas where they stand close enough together, the rest by hash.
struct Positions {
    grids: Vec<Option<Grid>>,
    rest: FxHashMap<CellId, u32>,
}

struct Grid {
    top: u32,
    left: u32,
    width: u32,
    height: u32,
    at: Vec<u32>,
}

type Bounds = (u32, u32, u32, u32, u64);

impl Positions {
    fn of(order: &[CellId], sheets: usize) -> Positions {
        let mut bounds: Vec<Option<Bounds>> = vec![None; sheets];
        for &(s, row, col) in order {
            if let Some(b) = bounds.get_mut(s) {
                *b = Some(match *b {
                    None => (row, col, row, col, 1),
                    Some((t, l, bt, r, n)) => {
                        (t.min(row), l.min(col), bt.max(row), r.max(col), n + 1)
                    }
                });
            }
        }
        let grids = bounds
            .iter()
            .map(|b| {
                let (top, left, bottom, right, n) = (*b)?;
                let (width, height) = (right - left + 1, bottom - top + 1);
                let cells = u64::from(width) * u64::from(height);
                (cells <= 4 * n + 64).then(|| Grid {
                    top,
                    left,
                    width,
                    height,
                    at: vec![NONE; cells as usize],
                })
            })
            .collect();
        let mut p = Positions {
            grids,
            rest: FxHashMap::default(),
        };
        for (i, &(s, row, col)) in order.iter().enumerate() {
            match p.grids.get_mut(s).and_then(Option::as_mut) {
                Some(g) => {
                    let k = (row - g.top) as usize * g.width as usize + (col - g.left) as usize;
                    g.at[k] = i as u32;
                }
                None => {
                    p.rest.insert((s, row, col), i as u32);
                }
            }
        }
        p
    }

    fn get(&self, sheet: usize, row: u32, col: u32) -> Option<usize> {
        match self.grids.get(sheet) {
            Some(Some(g)) => {
                let (r, c) = (row.checked_sub(g.top)?, col.checked_sub(g.left)?);
                if r >= g.height || c >= g.width {
                    return None;
                }
                let at = g.at[r as usize * g.width as usize + c as usize];
                (at != NONE).then_some(at as usize)
            }
            _ => self.rest.get(&(sheet, row, col)).map(|&p| p as usize),
        }
    }

    fn each_in(
        &self,
        sheet: usize,
        (top, left, bottom, right): (u32, u32, u32, u32),
        f: &mut impl FnMut(usize),
    ) -> bool {
        let Some(Some(g)) = self.grids.get(sheet) else {
            return false;
        };
        let (t, l) = (top.max(g.top), left.max(g.left));
        let (b, r) = (
            bottom.min(g.top + g.height - 1),
            right.min(g.left + g.width - 1),
        );
        if t > b || l > r {
            return true;
        }
        for row in t..=b {
            let base = (row - g.top) as usize * g.width as usize;
            for col in l..=r {
                let at = g.at[base + (col - g.left) as usize];
                if at != NONE {
                    f(at as usize);
                }
            }
        }
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::formula::Expr;
    use crate::calc::reach::Dim;

    #[test]
    fn every_dollar_holds_a_reference_whose_every_number_is_held() {
        // KERNEL.33: a plan finds a reference once for its whole shape only
        // when every `$` holds it, so that it names the same cells at every
        // offset: a cell's two numbers, a range's four, a whole column's
        // columns, a whole row's rows. One unheld number moves it. The plan
        // of nodes reads an area every cell reads alike by its two
        // dimensions held (KERNEL.49's decision 7), the same references.
        for (text, held) in [
            ("$C$2", true),
            ("$C2", false),
            ("C$2", false),
            ("C2", false),
            ("$C$2:$D$4", true),
            ("$C$2:D4", false),
            ("$C$2:$D4", false),
            ("$C$2:D$4", false),
            ("C2:$D$4", false),
            ("$C:$D", true),
            ("$C:D", false),
            ("C:D", false),
            ("$2:$4", true),
            ("$2:4", false),
            ("2:4", false),
            ("S!$C$2:$D$4", true),
            ("S!C2:$D$4", false),
        ] {
            let Ok(Expr::Ref(r)) = formula::parse(text) else {
                panic!("{text} reads as one reference")
            };
            assert_eq!(every_dollar(&r), held, "{text}");
            let fixed =
                Dim::rows(&r).is_some_and(Dim::fixed) && Dim::cols(&r).is_some_and(Dim::fixed);
            assert_eq!(fixed, held, "{text}: its dimensions");
        }
    }
}
