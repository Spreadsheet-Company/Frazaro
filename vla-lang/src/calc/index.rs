//! A vector's index (KERNEL.8, slice a, 2026-10-18): what lets a lookup or a
//! criterion asked again and again of one range cost a probe and not a scan,
//! exact by construction.
//!
//! A *vector* is the cells of one row or one column a lookup searches
//! (`VLOOKUP`'s first column, `HLOOKUP`'s first row, the vector of `MATCH`,
//! `XMATCH`, `XLOOKUP` and `LOOKUP`), or the range a criterion of `COUNTIF`
//! or `SUMIF` tests, any rectangle, row by row. It is known by the rectangle
//! its reference names at the cell computing it and by the equality it keys,
//! never by its text, so that the 20,000 cells of one shape reading
//! `$A$2:$A$20001` share one index, and a defined name naming that rectangle
//! shares it too; a range that moves with its cell is a vector at each cell.
//!
//! An index holds each distinct key of its vector once: for a lookup its
//! first and its last position, and the vector's first and last reason, a
//! cell not computed, where it holds one; for a criterion each key's count,
//! or, for a `SUMIF` and its sum range, the numbers its cells contribute,
//! added in the scan's order from the scan's start, and each key's first
//! error there; and, made when first asked for, each kind's distinct values
//! in order, for `XMATCH`'s and `XLOOKUP`'s next smaller and next larger. A
//! key is a number's bits, 0 and -0 one, for a lookup exactly and for a
//! criterion to fifteen significant digits, as each compares two numbers
//! (`eval::exact_order`, `eval::compare_numbers`); a text as Excel's order
//! of texts holds two one (`order::keyed`: its characters folded, `ß` as
//! `ss`, `æ` as `ae`), hashed and compared as its characters are read, so
//! that a probe makes no string; a truth value; and for a criterion an
//! error and the blank, and a text spelling a number, keyed as the number
//! for a sum and apart from it for a count, since `<>` counts it and `=`
//! does not (Excel's save of the lookups' fixture, `Criteria!H4`).
//!
//! The scan is the specification, kept beside: it answers a vector's first
//! probe in an evaluation, every pattern, every criterion by size, a next
//! smaller or larger among texts beyond Excel's known order, which it
//! labels where their order decides, and every environment that keeps no
//! index, and a test switch has every probe answered both ways and
//! compared. An index is built at the second probe of its vector in one
//! evaluation, a frame or a pass of `Calc`, so that a vector probed once
//! never pays a build; kept until a cell of its vector changes, which its
//! environment tells it (`Indexes::changed`), and built again at its next
//! probe; and the indexes' slots together are held to the grid's own limit,
//! `machine::CELLS`, a vector past it scanned. Nothing here draws at random
//! or reads a map's order: an index moves no answer, only the time a probe
//! takes.

use std::cell::{Cell, OnceCell, RefCell};
use std::cmp::Ordering;
use std::hash::Hasher;

use crate::fx::{FxHashMap, FxHasher};
use crate::machine::CELLS;

use super::criteria::{cells_of, has_wildcard, Criterion, Target};
use super::eval::{exact_order, finite, Ctx, Reason};
use super::graph::{cells_in, places, Area, CellId};
use super::order;
use super::value::{round15, text_to_number, ErrorKind, Value};

/// A row or a column of cells a lookup searches; a single cell is a column
/// of one.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(crate) struct Line {
    pub(crate) area: Area,
}

impl Line {
    /// The line an area is, when it is one row or one column.
    pub(crate) fn of(area: Area) -> Option<Line> {
        (area.left == area.right || area.top == area.bottom).then_some(Line { area })
    }

    /// Whether the line runs across a row.
    fn across(&self) -> bool {
        self.area.left != self.area.right
    }

    /// How many cells the line spans, held or not.
    pub(crate) fn len(&self) -> u32 {
        if self.across() {
            self.area.right - self.area.left + 1
        } else {
            self.area.bottom - self.area.top + 1
        }
    }

    /// The cell at a position, from 0.
    pub(crate) fn cell(&self, k: u32) -> CellId {
        if self.across() {
            (self.area.sheet, self.area.top, self.area.left + k)
        } else {
            (self.area.sheet, self.area.top + k, self.area.left)
        }
    }

    /// A cell's position along the line.
    fn position(&self, row: u32, col: u32) -> u32 {
        if self.across() {
            col - self.area.left
        } else {
            row - self.area.top
        }
    }

    /// The first and last positions inside the sheet's extent, where every
    /// cell the line holds stands; `None` for a line that holds none.
    fn within(&self, ctx: &Ctx<'_>) -> Option<(u32, u32)> {
        let ((e_top, e_left), (e_bottom, e_right)) = ctx.env.extent(self.area.sheet)?;
        let a = &self.area;
        let (top, bottom) = (a.top.max(e_top), a.bottom.min(e_bottom));
        let (left, right) = (a.left.max(e_left), a.right.min(e_right));
        if top > bottom || left > right {
            return None;
        }
        Some((self.position(top, left), self.position(bottom, right)))
    }
}

/// What a walk of a vector hands on for each cell it holds: its value lent,
/// or its reason.
pub(crate) type Read<'v> = Result<&'v Value, &'v Reason>;

/// Each cell an area holds, row by row, its row, its column and its value
/// lent or its reason, to `f`, as `Ctx::each` walks an area: a sheet the
/// environment keeps as an array by index, its empty slots passed by; a
/// sheet holding reasons beside its cells at every place (KERNEL.8, slice
/// b); any other through its cells, each read through the environment, a
/// cell that reads as nothing passed by. `f` answers whether to go on.
pub(crate) fn walk<F>(ctx: &Ctx<'_>, area: &Area, mut f: F)
where
    F: FnMut(u32, u32, Read<'_>) -> bool,
{
    if let Some(kept) = ctx.env.dense(area.sheet) {
        let _ = kept.each_at(
            area,
            |row, col, v| if f(row, col, Ok(v)) { Ok(()) } else { Err(()) },
        );
        return;
    }
    let model = ctx.env.model();
    let Some(sheet) = model.sheets.get(area.sheet) else {
        return;
    };
    let extent = ctx.env.extent(area.sheet);
    let mut visit = |row: u32, col: u32| -> bool {
        match ctx.env.read((area.sheet, row, col)) {
            Ok(Value::Empty) => true,
            Ok(v) => f(row, col, Ok(&v)),
            Err(r) => f(row, col, Err(&r)),
        }
    };
    if ctx.env.reasons_beside(area.sheet) {
        for (row, col) in places(extent, area) {
            if !visit(row, col) {
                return;
            }
        }
        return;
    }
    for ((row, col), _) in cells_in(sheet, extent, area) {
        if !visit(row, col) {
            return;
        }
    }
}

/// A key as a probe holds it, borrowed: a number's bits, 0 and -0 one; a
/// text as written, folded as it is read; a truth value; an error, the
/// blank and a text spelling a number by the number's bits, which only a
/// criterion keys.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Probe<'a> {
    Number(u64),
    Text(&'a str),
    Bool(bool),
    Error(ErrorKind),
    Blank,
    Spelled(u64),
}

/// A key as an index keeps it, a text folded once.
#[derive(Clone, Debug, PartialEq)]
enum Key {
    Number(u64),
    Text(Box<str>),
    Bool(bool),
    Error(ErrorKind),
    Blank,
    Spelled(u64),
}

/// A number's bits as a key, -0 made 0, so that the two are one, as `=`
/// holds them.
#[inline]
fn bits(n: f64) -> u64 {
    if n == 0.0 {
        0
    } else {
        n.to_bits()
    }
}

/// A number as a criterion's equality keys it: its fifteen significant
/// digits' bits, as [`super::eval::compare_numbers`] holds two numbers
/// equal; a whole number under 10^15 is its own fifteen digits.
#[inline]
fn fifteen(n: f64) -> u64 {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        bits(n)
    } else {
        bits(round15(n))
    }
}

impl Probe<'_> {
    /// The hash a key is found by: the house's own (`fx`, KERNEL.25), a
    /// text's keyed bytes fed sixty-four at a time (`order::keyed`), so that
    /// a text keyed and one probed as written hash alike, a character
    /// folding to one of another length among them, and `ß` as `ss`.
    fn hash(self) -> u64 {
        let mut h = FxHasher::default();
        match self {
            Probe::Number(b) => {
                h.write_u8(0);
                h.write_u64(b);
            }
            Probe::Text(t) => {
                h.write_u8(1);
                let mut buf = [0u8; 64];
                if t.is_ascii() {
                    for chunk in t.as_bytes().chunks(64) {
                        for (slot, &b) in buf.iter_mut().zip(chunk) {
                            *slot = b.to_ascii_lowercase();
                        }
                        h.write(&buf[..chunk.len()]);
                    }
                } else {
                    let mut n = 0;
                    for c in order::keyed(t) {
                        let mut one = [0u8; 4];
                        for &b in c.encode_utf8(&mut one).as_bytes() {
                            buf[n] = b;
                            n += 1;
                            if n == buf.len() {
                                h.write(&buf);
                                n = 0;
                            }
                        }
                    }
                    if n > 0 {
                        h.write(&buf[..n]);
                    }
                }
            }
            Probe::Bool(b) => {
                h.write_u8(2);
                h.write_u8(u8::from(b));
            }
            Probe::Error(k) => {
                h.write_u8(3);
                h.write_u8(k as u8);
            }
            Probe::Blank => h.write_u8(4),
            Probe::Spelled(b) => {
                h.write_u8(5);
                h.write_u64(b);
            }
        }
        h.finish()
    }

    /// Whether the probe is a key.
    fn is(self, key: &Key) -> bool {
        match (self, key) {
            (Probe::Number(a), Key::Number(b)) => a == *b,
            (Probe::Text(t), Key::Text(k)) => order::equal(t, k),
            (Probe::Bool(a), Key::Bool(b)) => a == *b,
            (Probe::Error(a), Key::Error(b)) => a == *b,
            (Probe::Blank, Key::Blank) => true,
            (Probe::Spelled(a), Key::Spelled(b)) => a == *b,
            _ => false,
        }
    }

    /// The key an index keeps for it.
    fn key(self) -> Key {
        match self {
            Probe::Number(b) => Key::Number(b),
            Probe::Text(t) => Key::Text(order::key(t).into_boxed_str()),
            Probe::Bool(b) => Key::Bool(b),
            Probe::Error(k) => Key::Error(k),
            Probe::Blank => Key::Blank,
            Probe::Spelled(b) => Key::Spelled(b),
        }
    }
}

/// A value as a lookup's equality keys it: a number, a text and a truth
/// value; an error, the blank and an empty cell are found by no probe.
pub(crate) fn lookup_probe(v: &Value) -> Option<Probe<'_>> {
    match v {
        Value::Number(n) => Some(Probe::Number(bits(*n))),
        Value::Text(t) => Some(Probe::Text(t)),
        Value::Bool(b) => Some(Probe::Bool(*b)),
        Value::Error(_) | Value::Empty => None,
    }
}

/// Whether a cell's value is a lookup's probe, as an index finds it.
fn lookup_matches(cell: &Value, p: Probe<'_>) -> bool {
    match (cell, p) {
        (Value::Number(n), Probe::Number(b)) => bits(*n) == b,
        (Value::Text(t), Probe::Text(q)) => order::equal(t, q),
        (Value::Bool(x), Probe::Bool(y)) => *x == y,
        _ => false,
    }
}

/// A cell's value as a count keys it (`criteria::meets`): a number to
/// fifteen digits; a text spelling a number apart from the number, an empty
/// text as the blank, any other text folded; a truth value; an error;
/// nothing as the blank.
fn count_probe(v: &Value) -> Probe<'_> {
    match v {
        Value::Number(n) => Probe::Number(fifteen(*n)),
        Value::Text(t) => match text_to_number(t) {
            Some(x) => Probe::Spelled(fifteen(x)),
            None if t.is_empty() => Probe::Blank,
            None => Probe::Text(t),
        },
        Value::Bool(b) => Probe::Bool(*b),
        Value::Error(k) => Probe::Error(*k),
        Value::Empty => Probe::Blank,
    }
}

/// A cell's value as a sum keys it: as a count does, but a text spelling a
/// number as the number, since a sum answers `=` alone, which meets both,
/// and adds the two in the scan's order.
fn sum_probe(v: &Value) -> Probe<'_> {
    match count_probe(v) {
        Probe::Spelled(b) => Probe::Number(b),
        p => p,
    }
}

/// A criterion's target as its equality keys it; `None` for a pattern,
/// which only a scan matches.
fn target_probe(t: &Target) -> Option<Probe<'_>> {
    match t {
        Target::Number(n) => Some(Probe::Number(fifteen(*n))),
        Target::Text(p) if has_wildcard(p) => None,
        Target::Text(p) => Some(Probe::Text(p)),
        Target::Bool(b) => Some(Probe::Bool(*b)),
        Target::Error(k) => Some(Probe::Error(*k)),
        Target::Blank => Some(Probe::Blank),
    }
}

/// The distinct keys of a vector, each with the next of its hash's chain.
#[derive(Default)]
struct Keys {
    keys: Vec<Key>,
    next: Vec<u32>,
    heads: FxHashMap<u64, u32>,
}

/// The end of a chain.
const END: u32 = u32::MAX;

impl Keys {
    fn find(&self, p: Probe<'_>) -> Option<usize> {
        let mut i = *self.heads.get(&p.hash())?;
        while i != END {
            if p.is(&self.keys[i as usize]) {
                return Some(i as usize);
            }
            i = self.next[i as usize];
        }
        None
    }

    /// The key's place, made where it is new; and whether it was.
    fn add(&mut self, p: Probe<'_>) -> (usize, bool) {
        let h = p.hash();
        let mut i = self.heads.get(&h).copied().unwrap_or(END);
        let head = i;
        while i != END {
            if p.is(&self.keys[i as usize]) {
                return (i as usize, false);
            }
            i = self.next[i as usize];
        }
        let at = self.keys.len() as u32;
        self.heads.insert(h, at);
        self.keys.push(p.key());
        self.next.push(head);
        (at as usize, true)
    }
}

/// A lookup's index: each key's first and last position, the vector's first
/// and last reason with theirs, and its kinds' values in order, made when
/// first asked for.
struct LookupIndex {
    keys: Keys,
    first: Vec<u32>,
    last: Vec<u32>,
    reasons: Option<((u32, Reason), (u32, Reason))>,
    ordered: OnceCell<Ordered>,
}

/// Each kind's distinct values in order, by key: numbers by value, texts in
/// Excel's order of texts (`order::compare`), and the two truth values;
/// and whether a text holds a character beyond the order's table, whose
/// place among the others is not known.
struct Ordered {
    numbers: Vec<(f64, u32)>,
    texts: Vec<u32>,
    bools: [Option<u32>; 2],
    beyond: bool,
}

/// A count's index: each key's count, the cells held, and the first reason
/// that stops the scan.
struct CountIndex {
    keys: Keys,
    counts: Vec<u64>,
    present: u64,
    stop: Option<Reason>,
}

/// A sum's index: each key's sum of the sum range's numbers, added in the
/// scan's order, each key's first error, and the first reason that stops
/// the scan, before which every error kept stands.
struct SumIndex {
    keys: Keys,
    sums: Vec<f64>,
    errors: Vec<Option<ErrorKind>>,
    stop: Option<Reason>,
}

enum Built {
    Lookup(LookupIndex),
    Count(CountIndex),
    Sum(SumIndex),
}

/// What a vector is keyed for.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Kind {
    Lookup,
    Count,
    /// `SUMIF`'s, its sum range beside.
    Sum(Area),
}

/// A vector's identity: its rectangle and what it is keyed for.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
struct Vector {
    area: Area,
    kind: Kind,
}

impl Vector {
    /// Whether a cell of either of its ranges is the one given.
    fn holds(&self, id: CellId) -> bool {
        let inside = |a: &Area| {
            a.sheet == id.0
                && (a.top..=a.bottom).contains(&id.1)
                && (a.left..=a.right).contains(&id.2)
        };
        inside(&self.area)
            || match &self.kind {
                Kind::Sum(s) => inside(s),
                _ => false,
            }
    }

    /// The sheets its ranges stand on.
    fn sheets(&self) -> (usize, Option<usize>) {
        match &self.kind {
            Kind::Sum(s) if s.sheet != self.area.sheet => (self.area.sheet, Some(s.sheet)),
            _ => (self.area.sheet, None),
        }
    }
}

/// Where a vector stands with its environment.
enum Held {
    /// Probed once in the evaluation numbered, and answered by the scan.
    Probed(u32),
    /// Built, and its slots.
    Built(Box<Built>, usize),
    /// Built once and reached by a change since: built again at its next
    /// probe.
    Stale,
}

/// The indexes an environment keeps (see the module's comment): a machine's
/// and `Calc`'s, each answering [`super::Env::indexes`]. An environment of
/// its own may keep them too, telling them each changed cell.
#[derive(Default)]
pub struct Indexes {
    held: RefCell<FxHashMap<Vector, Held>>,
    /// By sheet, the vectors built over it.
    over: RefCell<FxHashMap<usize, Vec<Vector>>>,
    evaluation: Cell<u32>,
    slots: Cell<usize>,
    built: Cell<usize>,
    /// Every build, every built index a change reached and every probe an
    /// index answered, for a test to count.
    #[cfg(test)]
    counted: Cell<(usize, usize, usize)>,
}

impl std::fmt::Debug for Indexes {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Indexes")
            .field("built", &self.built.get())
            .field("slots", &self.slots.get())
            .finish()
    }
}

impl Indexes {
    /// No index kept yet.
    pub fn new() -> Indexes {
        Indexes::default()
    }

    /// An evaluation begins: a vector probed once before is probed once
    /// again before it is built.
    pub(crate) fn begin(&self) {
        self.evaluation.set(self.evaluation.get().wrapping_add(1));
    }

    /// A cell changed: every index whose vector holds it, built again at its
    /// next probe. Nothing to ask where none is built, which every changed
    /// value of a frame asks, inline.
    #[inline]
    pub fn changed(&mut self, id: CellId) {
        if self.built.get() > 0 {
            self.reached(id);
        }
    }

    /// [`Indexes::changed`] where an index is built.
    #[inline(never)]
    fn reached(&mut self, id: CellId) {
        let over = self.over.get_mut();
        let Some(vectors) = over.get_mut(&id.0) else {
            return;
        };
        let reached: Vec<Vector> = vectors.iter().copied().filter(|v| v.holds(id)).collect();
        for v in reached {
            self.stale(v);
        }
    }

    /// Every index over a sheet built again at its next probe: its cells
    /// changed too many to tell.
    pub(crate) fn sheet_changed(&mut self, sheet: usize) {
        if self.built.get() == 0 {
            return;
        }
        let vectors: Vec<Vector> = self.over.get_mut().get(&sheet).cloned().unwrap_or_default();
        for v in vectors {
            self.stale(v);
        }
    }

    /// Every index dropped: a plan made again.
    pub(crate) fn clear(&mut self) {
        self.held.get_mut().clear();
        self.over.get_mut().clear();
        self.slots.set(0);
        self.built.set(0);
    }

    /// Whether any index is built, so that a change has one to reach.
    pub(crate) fn any(&self) -> bool {
        self.built.get() > 0
    }

    /// How many indexes are built, for a test to count.
    #[cfg(test)]
    pub(crate) fn built(&self) -> usize {
        self.built.get()
    }

    /// How many builds there were, how many built indexes a change reached
    /// and how many probes an index answered, since the indexes were made,
    /// for a test to count.
    #[cfg(test)]
    pub(crate) fn counted(&self) -> (usize, usize, usize) {
        self.counted.get()
    }

    #[cfg(test)]
    fn tally(&self, build: usize, reach: usize, answer: usize) {
        let (b, r, a) = self.counted.get();
        self.counted.set((b + build, r + reach, a + answer));
    }

    fn stale(&mut self, v: Vector) {
        let held = self.held.get_mut();
        if let Some(Held::Built(_, slots)) = held.get(&v) {
            self.slots.set(self.slots.get() - slots);
            self.built.set(self.built.get() - 1);
            held.insert(v, Held::Stale);
            #[cfg(test)]
            self.tally(0, 1, 0);
        }
        let over = self.over.get_mut();
        let (a, b) = v.sheets();
        for s in [Some(a), b].into_iter().flatten() {
            if let Some(list) = over.get_mut(&s) {
                list.retain(|w| *w != v);
            }
        }
    }

    /// `f` over the vector's index, built at its second probe in this
    /// evaluation, or at the first after a change reached it; `None` where
    /// the scan answers: the first probe, or slots past the limit. No
    /// borrow is held while the build reads the vector's cells.
    fn with<R>(&self, ctx: &Ctx<'_>, v: Vector, f: impl FnOnce(&Built) -> R) -> Option<R> {
        let now = self.evaluation.get();
        {
            let mut held = self.held.borrow_mut();
            match held.get(&v) {
                Some(Held::Built(b, _)) => {
                    #[cfg(test)]
                    self.tally(0, 0, 1);
                    return Some(f(b));
                }
                Some(Held::Probed(at)) if *at == now => {}
                Some(Held::Stale) => {}
                _ => {
                    held.insert(v, Held::Probed(now));
                    return None;
                }
            }
        }
        let slots = slots_of(ctx, &v);
        if self.slots.get() + slots > CELLS {
            self.held.borrow_mut().insert(v, Held::Probed(now));
            return None;
        }
        let built = build(ctx, &v);
        let answer = f(&built);
        self.held
            .borrow_mut()
            .insert(v, Held::Built(Box::new(built), slots));
        self.slots.set(self.slots.get() + slots);
        self.built.set(self.built.get() + 1);
        #[cfg(test)]
        self.tally(1, 0, 1);
        let mut over = self.over.borrow_mut();
        let (a, b) = v.sheets();
        for s in [Some(a), b].into_iter().flatten() {
            over.entry(s).or_default().push(v);
        }
        Some(answer)
    }

    /// A lookup's exact match from its index (see [`exact`]).
    fn exact(
        &self,
        ctx: &Ctx<'_>,
        line: Line,
        p: Probe<'_>,
        from_end: bool,
    ) -> Option<Result<Option<u32>, Reason>> {
        let v = Vector {
            area: line.area,
            kind: Kind::Lookup,
        };
        self.with(ctx, v, |b| match b {
            Built::Lookup(ix) => ix.exact(p, from_end),
            _ => unreachable!("a lookup's vector keyed for a lookup"),
        })
    }

    /// A lookup's exact match or its next smaller or larger from its index
    /// (see [`nearest`]).
    fn nearest(
        &self,
        ctx: &Ctx<'_>,
        line: Line,
        probe: &Value,
        larger: bool,
        from_end: bool,
    ) -> Option<Result<Option<u32>, Reason>> {
        let v = Vector {
            area: line.area,
            kind: Kind::Lookup,
        };
        self.with(ctx, v, |b| match b {
            Built::Lookup(ix) => ix.nearest(probe, larger, from_end),
            _ => unreachable!("a lookup's vector keyed for a lookup"),
        })
        .flatten()
    }

    /// `COUNTIF` of an equality or its `<>` from its range's index; `None`
    /// where the scan answers (`criteria::countif`).
    pub(crate) fn count(
        &self,
        ctx: &Ctx<'_>,
        area: Area,
        c: &Criterion,
    ) -> Option<Result<Value, Reason>> {
        let (target, negated) = match c {
            Criterion::Eq(t) => (t, false),
            Criterion::Ne(t) => (t, true),
            _ => return None,
        };
        let p = target_probe(target)?;
        let v = Vector {
            area,
            kind: Kind::Count,
        };
        let answer = self.with(ctx, v, |b| match b {
            Built::Count(ix) => ix.count(p, negated, cells_of(&area)),
            _ => unreachable!("a criterion's vector keyed for a count"),
        })?;
        #[cfg(test)]
        if checking() {
            assert_eq!(
                answer,
                super::criteria::scan_count(ctx, area, c),
                "{c:?} over {area:?}"
            );
        }
        Some(answer)
    }

    /// `SUMIF` of an equality from its pair's index; `None` where the scan
    /// answers (`criteria::sumif`).
    pub(crate) fn sum(
        &self,
        ctx: &Ctx<'_>,
        area: Area,
        sums: Area,
        c: &Criterion,
    ) -> Option<Result<Value, Reason>> {
        let Criterion::Eq(target) = c else {
            return None;
        };
        let p = target_probe(target)?;
        let v = Vector {
            area,
            kind: Kind::Sum(sums),
        };
        let answer = self.with(ctx, v, |b| match b {
            Built::Sum(ix) => ix.sum(p),
            _ => unreachable!("a sum's vector keyed for a sum"),
        })?;
        #[cfg(test)]
        if checking() {
            assert_eq!(
                answer,
                super::criteria::scan_sum(ctx, area, sums, c),
                "{c:?} over {area:?}, summing {sums:?}"
            );
        }
        Some(answer)
    }
}

/// The slots an index of a vector takes, the cells of its ranges inside
/// their sheets' extents.
fn slots_of(ctx: &Ctx<'_>, v: &Vector) -> usize {
    let inside = |a: &Area| -> usize {
        let Some(((t, l), (b, r))) = ctx.env.extent(a.sheet) else {
            return 0;
        };
        let (top, bottom) = (a.top.max(t), a.bottom.min(b));
        let (left, right) = (a.left.max(l), a.right.min(r));
        if top > bottom || left > right {
            return 0;
        }
        (bottom - top + 1) as usize * (right - left + 1) as usize
    };
    inside(&v.area)
        + match &v.kind {
            Kind::Sum(s) => inside(s),
            _ => 0,
        }
}

/// A vector's index built from its cells as they stand.
fn build(ctx: &Ctx<'_>, v: &Vector) -> Built {
    match v.kind {
        Kind::Lookup => Built::Lookup(build_lookup(ctx, Line { area: v.area })),
        Kind::Count => Built::Count(build_count(ctx, &v.area)),
        Kind::Sum(sums) => Built::Sum(build_sum(ctx, &v.area, &sums)),
    }
}

fn build_lookup(ctx: &Ctx<'_>, line: Line) -> LookupIndex {
    let mut ix = LookupIndex {
        keys: Keys::default(),
        first: Vec::new(),
        last: Vec::new(),
        reasons: None,
        ordered: OnceCell::new(),
    };
    walk(ctx, &line.area, |row, col, read| {
        let k = line.position(row, col);
        match read {
            Ok(v) => {
                if let Some(p) = lookup_probe(v) {
                    let (i, new) = ix.keys.add(p);
                    if new {
                        ix.first.push(k);
                        ix.last.push(k);
                    } else {
                        ix.last[i] = k;
                    }
                }
            }
            Err(r) => match &mut ix.reasons {
                None => ix.reasons = Some(((k, r.clone()), (k, r.clone()))),
                Some((_, last)) => *last = (k, r.clone()),
            },
        }
        true
    });
    ix
}

impl LookupIndex {
    /// The exact match as [`scan_exact`] finds it.
    fn exact(&self, p: Probe<'_>, from_end: bool) -> Result<Option<u32>, Reason> {
        let found = self.keys.find(p);
        match (found, &self.reasons) {
            (None, None) => Ok(None),
            (None, Some(((_, first), (_, last)))) => {
                Err(if from_end { last } else { first }.clone())
            }
            (Some(i), None) => Ok(Some(if from_end {
                self.last[i]
            } else {
                self.first[i]
            })),
            (Some(i), Some(((fp, first), (lp, last)))) => {
                if from_end {
                    let m = self.last[i];
                    if *lp > m {
                        Err(last.clone())
                    } else {
                        Ok(Some(m))
                    }
                } else {
                    let m = self.first[i];
                    if *fp < m {
                        Err(first.clone())
                    } else {
                        Ok(Some(m))
                    }
                }
            }
        }
    }

    fn ordered(&self) -> &Ordered {
        self.ordered.get_or_init(|| {
            let mut o = Ordered {
                numbers: Vec::new(),
                texts: Vec::new(),
                bools: [None, None],
                beyond: false,
            };
            for (i, key) in self.keys.keys.iter().enumerate() {
                match key {
                    Key::Number(b) => o.numbers.push((f64::from_bits(*b), i as u32)),
                    Key::Text(_) => o.texts.push(i as u32),
                    Key::Bool(b) => o.bools[usize::from(*b)] = Some(i as u32),
                    Key::Error(_) | Key::Blank | Key::Spelled(_) => {}
                }
            }
            o.numbers.sort_by(|a, b| exact_order(a.0, b.0));
            let text = |i: u32| match &self.keys.keys[i as usize] {
                Key::Text(t) => &**t,
                _ => "",
            };
            o.beyond = o.texts.iter().any(|&i| !order::within(text(i)));
            o.texts
                .sort_by(|&a, &b| order::compare(text(a), text(b)).order);
            o
        })
    }

    /// The exact match, else the next smaller or larger, as [`scan_nearest`]
    /// finds it; `None` where the scan answers, a text's next smaller or
    /// larger where the probe or a text of the vector holds a character
    /// beyond Excel's known order, which the scan labels where it decides.
    fn nearest(
        &self,
        probe: &Value,
        larger: bool,
        from_end: bool,
    ) -> Option<Result<Option<u32>, Reason>> {
        if let Some(((_, first), (_, last))) = &self.reasons {
            return Some(Err(if from_end { last } else { first }.clone()));
        }
        let Some(p) = lookup_probe(probe) else {
            return Some(Ok(None));
        };
        let at = |i: usize| {
            if from_end {
                self.last[i]
            } else {
                self.first[i]
            }
        };
        if let Some(i) = self.keys.find(p) {
            return Some(Ok(Some(at(i))));
        }
        let o = self.ordered();
        let key = match probe {
            Value::Number(x) => {
                let n = o
                    .numbers
                    .partition_point(|&(v, _)| exact_order(v, *x) == Ordering::Less);
                // `n` is the first not under the probe, which is above it, the
                // probe not being a key.
                if larger {
                    o.numbers.get(n).map(|&(_, i)| i)
                } else {
                    n.checked_sub(1).map(|j| o.numbers[j].1)
                }
            }
            Value::Text(t) => {
                if o.beyond || !order::within(t) {
                    return None;
                }
                let text = |i: u32| match &self.keys.keys[i as usize] {
                    Key::Text(k) => &**k,
                    _ => "",
                };
                let n = o
                    .texts
                    .partition_point(|&i| order::compare(text(i), t).order == Ordering::Less);
                if larger {
                    o.texts.get(n).copied()
                } else {
                    n.checked_sub(1).map(|j| o.texts[j])
                }
            }
            Value::Bool(b) => match (b, larger) {
                (false, true) => o.bools[1],
                (true, false) => o.bools[0],
                _ => None,
            },
            _ => None,
        };
        Some(Ok(key.map(|i| at(i as usize))))
    }
}

fn build_count(ctx: &Ctx<'_>, area: &Area) -> CountIndex {
    let mut ix = CountIndex {
        keys: Keys::default(),
        counts: Vec::new(),
        present: 0,
        stop: None,
    };
    walk(ctx, area, |_, _, read| match read {
        Ok(v) => {
            ix.present += 1;
            let (i, new) = ix.keys.add(count_probe(v));
            if new {
                ix.counts.push(1);
            } else {
                ix.counts[i] += 1;
            }
            true
        }
        Err(r) => {
            ix.stop = Some(r.clone());
            false
        }
    });
    ix
}

impl CountIndex {
    /// The cells meeting an equality, or its `<>`, as `criteria::scan_count`
    /// counts them: the cells keyed alike, the texts spelling a number with
    /// the numbers equal to it, the empty cells of the whole range with them
    /// where the target is the blank, and for `<>` every other cell of the
    /// range, of a number every cell but the numbers equal to it.
    fn count(&self, p: Probe<'_>, negated: bool, cells: u64) -> Result<Value, Reason> {
        if let Some(r) = &self.stop {
            return Err(r.clone());
        }
        let of = |p: Probe<'_>| self.keys.find(p).map_or(0, |i| self.counts[i]);
        let n = match p {
            Probe::Number(_) if negated => return Ok(Value::Number((cells - of(p)) as f64)),
            Probe::Number(b) => of(p) + of(Probe::Spelled(b)),
            Probe::Blank => of(p) + (cells - self.present),
            _ => of(p),
        };
        let n = if negated { cells - n } else { n };
        Ok(Value::Number(n as f64))
    }
}

fn build_sum(ctx: &Ctx<'_>, area: &Area, sums: &Area) -> SumIndex {
    let mut ix = SumIndex {
        keys: Keys::default(),
        sums: Vec::new(),
        errors: Vec::new(),
        stop: None,
    };
    walk(ctx, sums, |row, col, read| {
        let held = match read {
            Ok(Value::Number(n)) => Ok(*n),
            Ok(Value::Error(k)) => Err(*k),
            Ok(_) => return true,
            Err(r) => {
                ix.stop = Some(r.clone());
                return false;
            }
        };
        let at = (
            area.sheet,
            area.top + (row - sums.top),
            area.left + (col - sums.left),
        );
        let tested = match ctx.env.read(at) {
            Ok(v) => v,
            Err(r) => {
                ix.stop = Some(r);
                return false;
            }
        };
        let (i, new) = ix.keys.add(sum_probe(&tested));
        if new {
            ix.sums.push(0.0);
            ix.errors.push(None);
        }
        match held {
            Ok(n) => ix.sums[i] += n,
            Err(k) => {
                ix.errors[i].get_or_insert(k);
            }
        }
        true
    });
    ix
}

impl SumIndex {
    /// The sum of an equality as `criteria::scan_sum` adds it: the key's
    /// first error, which stands before any reason; else the first reason;
    /// else its sum.
    fn sum(&self, p: Probe<'_>) -> Result<Value, Reason> {
        let found = self.keys.find(p);
        if let Some(k) = found.and_then(|i| self.errors[i]) {
            return Ok(Value::Error(k));
        }
        if let Some(r) = &self.stop {
            return Err(r.clone());
        }
        Ok(finite(found.map_or(0.0, |i| self.sums[i])))
    }
}

// ---- the lookups' searches, each from an index where one stands, and the
// ---- scans that are their meaning ---------------------------------------

/// A lookup's exact match along a line (`MATCH` 0, `VLOOKUP` and `HLOOKUP`
/// exactly, `XMATCH` and `XLOOKUP` mode 0): the position of the first cell
/// equal to the probe under the lookups' equality (no two kinds equal, a
/// number by its value, a text without case, a truth value), or the last
/// `from_end`; a reason met before it, scanning in that direction, answers
/// the reason. `None` where nothing matches, an error, the blank and an
/// empty probe among them.
pub(crate) fn exact(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    from_end: bool,
) -> Result<Option<u32>, Reason> {
    let Some(p) = lookup_probe(probe) else {
        return Ok(None);
    };
    if let Some(ix) = ctx.env.indexes() {
        if let Some(answer) = ix.exact(ctx, line, p, from_end) {
            #[cfg(test)]
            if checking() {
                assert_eq!(
                    answer,
                    scan_exact(ctx, line, p, from_end),
                    "{probe:?} along {line:?}"
                );
            }
            return answer;
        }
    }
    scan_exact(ctx, line, p, from_end)
}

/// [`exact`]'s meaning: the line walked in order, stopping at its first
/// match or its first reason; from the end, every cell, its last match
/// answering unless a reason stands after it.
fn scan_exact(
    ctx: &Ctx<'_>,
    line: Line,
    p: Probe<'_>,
    from_end: bool,
) -> Result<Option<u32>, Reason> {
    let mut found: Option<u32> = None;
    let mut reason: Option<(u32, Reason)> = None;
    walk(ctx, &line.area, |row, col, read| {
        let k = line.position(row, col);
        match read {
            Ok(v) if lookup_matches(v, p) => {
                found = Some(k);
                from_end
            }
            Ok(_) => true,
            Err(r) => {
                reason = Some((k, r.clone()));
                from_end
            }
        }
    });
    match (found, reason) {
        (_, Some((rk, r))) if found.map_or(true, |m| if from_end { rk > m } else { rk < m }) => {
            Err(r)
        }
        (found, _) => Ok(found),
    }
}

/// A pattern's match along a line (`MATCH` 0 and `VLOOKUP` exactly, with a
/// text holding `*`, `?` or `~`; `XMATCH` mode 2): the first text matching
/// it by Excel's wildcards (`criteria::wildcard`), or the last `from_end`,
/// a reason met before it answering. A scan, every time: a pattern is no
/// key.
pub(crate) fn pattern(
    ctx: &Ctx<'_>,
    line: Line,
    pattern: &str,
    from_end: bool,
) -> Result<Option<u32>, Reason> {
    let mut found: Option<u32> = None;
    let mut reason: Option<(u32, Reason)> = None;
    walk(ctx, &line.area, |row, col, read| {
        let k = line.position(row, col);
        match read {
            Ok(Value::Text(t)) if super::criteria::wildcard(pattern, t) => {
                found = Some(k);
                from_end
            }
            Ok(_) => true,
            Err(r) => {
                reason = Some((k, r.clone()));
                from_end
            }
        }
    });
    match (found, reason) {
        (_, Some((rk, r))) if found.map_or(true, |m| if from_end { rk > m } else { rk < m }) => {
            Err(r)
        }
        (found, _) => Ok(found),
    }
}

/// The kind a lookup compares a value within: a number, a text, a truth
/// value; `None` for an error and an empty cell, which no lookup compares.
fn kind_of(v: &Value) -> Option<u8> {
    match v {
        Value::Number(_) => Some(0),
        Value::Text(_) => Some(1),
        Value::Bool(_) => Some(2),
        Value::Error(_) | Value::Empty => None,
    }
}

/// Two values of one kind as a lookup orders them: numbers by value,
/// exactly, texts in Excel's order of texts, FALSE before TRUE; the label
/// of two texts a character beyond its known order decides
/// (`order::sized`).
fn ordered(a: &Value, b: &Value) -> Result<Ordering, Reason> {
    match (a, b) {
        (Value::Text(x), Value::Text(y)) => order::sized(x, y),
        _ => Ok(same_kind_order(a, b)),
    }
}

/// [`ordered`]'s order, a text beyond Excel's known order weighed after the
/// known ones by its code point, never labelled: whether two values are
/// one, which is always known.
fn same_kind_order(a: &Value, b: &Value) -> Ordering {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => exact_order(*x, *y),
        (Value::Text(x), Value::Text(y)) => order::compare(x, y).order,
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        _ => Ordering::Equal,
    }
}

/// `XMATCH`'s and `XLOOKUP`'s exact match or, where none, the next larger
/// (`larger`) or the next smaller, along a line sorted or not (modes 1 and
/// -1, searching first to last or last to first): the position of the cell
/// equal to the probe, else of the least value above it or the greatest
/// below it among the cells of its kind, its first position, or its last
/// `from_end`; any reason in the line answers, its first, or its last
/// `from_end`, a search for the nearest reading every cell.
pub(crate) fn nearest(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    larger: bool,
    from_end: bool,
) -> Result<Option<u32>, Reason> {
    if lookup_probe(probe).is_none() {
        return Ok(None);
    }
    if let Some(ix) = ctx.env.indexes() {
        if let Some(answer) = ix.nearest(ctx, line, probe, larger, from_end) {
            #[cfg(test)]
            if checking() {
                assert_eq!(
                    answer,
                    scan_nearest(ctx, line, probe, larger, from_end),
                    "{probe:?} along {line:?}"
                );
            }
            return answer;
        }
    }
    scan_nearest(ctx, line, probe, larger, from_end)
}

/// [`nearest`]'s meaning: the line walked whole.
fn scan_nearest(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    larger: bool,
    from_end: bool,
) -> Result<Option<u32>, Reason> {
    let kind = kind_of(probe);
    let mut exact: Option<(u32, u32)> = None;
    let mut best: Option<(Value, u32, u32)> = None;
    let mut reasons: Option<((u32, Reason), (u32, Reason))> = None;
    // The first comparison a character beyond Excel's known order decided.
    let mut unknown: Option<Reason> = None;
    walk(ctx, &line.area, |row, col, read| {
        let k = line.position(row, col);
        let v = match read {
            Ok(v) => v,
            Err(r) => {
                match &mut reasons {
                    None => reasons = Some(((k, r.clone()), (k, r.clone()))),
                    Some((_, last)) => *last = (k, r.clone()),
                }
                return from_end;
            }
        };
        if kind_of(v) != kind {
            return true;
        }
        let o = match ordered(v, probe) {
            Ok(o) => o,
            Err(r) => {
                unknown.get_or_insert(r);
                return true;
            }
        };
        match o {
            Ordering::Equal => match &mut exact {
                None => exact = Some((k, k)),
                Some((_, last)) => *last = k,
            },
            o if (o == Ordering::Greater) == larger => {
                let better = match &best {
                    None => true,
                    Some((held, _, _)) => match ordered(v, held) {
                        Ok(o) if larger => o == Ordering::Less,
                        Ok(o) => o == Ordering::Greater,
                        Err(r) => {
                            unknown.get_or_insert(r);
                            return true;
                        }
                    },
                };
                if better {
                    best = Some((v.clone(), k, k));
                } else if let Some((held, _, last)) = &mut best {
                    if same_kind_order(v, held) == Ordering::Equal {
                        *last = k;
                    }
                }
            }
            _ => {}
        }
        true
    });
    if let Some(((_, first), (_, last))) = reasons {
        return Err(if from_end { last } else { first });
    }
    let pick = |(first, last): (u32, u32)| if from_end { last } else { first };
    // A cell equal to the probe answers whatever the order elsewhere.
    if let Some(e) = exact {
        return Ok(Some(pick(e)));
    }
    if let Some(r) = unknown {
        return Err(r);
    }
    Ok(best.map(|(_, f, l)| pick((f, l))))
}

/// The cell at a line's position, as a read of it answers.
fn read_at(ctx: &Ctx<'_>, line: Line, k: u32) -> Result<Value, Reason> {
    ctx.env.read(line.cell(k))
}

/// `MATCH`'s type -1, its smallest value not under the probe on a
/// descending line: the line walked from its first cell, a cell of the
/// probe's kind not under it kept, and the walk done at the first under it;
/// other kinds, empty cells and errors passed by; a reason read answers.
/// Not the bisection, as Excel's save of the lookups' fixture answers 1
/// for 15 over 50, 10, 70, 20, 80, 30, 60, 40, where any bisection keeps a
/// later cell (`Lookups!A26`); on a descending line the two are one.
pub(crate) fn descending(ctx: &Ctx<'_>, line: Line, probe: &Value) -> Result<Option<u32>, Reason> {
    let Some(kind) = kind_of(probe) else {
        return Ok(None);
    };
    let mut kept: Option<u32> = None;
    let mut stop: Option<Reason> = None;
    walk(ctx, &line.area, |row, col, read| match read {
        Ok(v) if kind_of(v) == Some(kind) => match ordered(v, probe) {
            Ok(Ordering::Less) => false,
            Ok(_) => {
                kept = Some(line.position(row, col));
                true
            }
            Err(r) => {
                stop = Some(r);
                false
            }
        },
        Ok(_) => true,
        Err(r) => {
            stop = Some(r.clone());
            false
        }
    });
    match stop {
        Some(r) => Err(r),
        None => Ok(kept),
    }
}

/// The bisection every approximate match takes but `MATCH`'s -1 (`MATCH`
/// 1, `VLOOKUP` and `HLOOKUP` but exactly, `LOOKUP`, `XMATCH` and `XLOOKUP`
/// searching 2 and -2), over the line as it stands, sorted or not, and
/// needing no index: low at its first position inside the sheet's extent
/// and high at its last; the middle at `(low + high) / 2` rounded down, and
/// from it the nearest cell of the probe's kind toward low, else toward
/// high, other kinds, empty cells and errors passed by; a cell `keep`
/// keeps, by how it compares with the probe, the answer so far and low
/// moved past it, any other moving high before it; done when low passes
/// high, the answer the last kept. On a line sorted by kind and value, as
/// every approximate match asks, that is the last cell of the probe's kind
/// that `keep` keeps. A reason read answers.
pub(crate) fn bisect(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    keep: impl Fn(Ordering) -> bool,
) -> Result<Option<u32>, Reason> {
    let Some(kind) = kind_of(probe) else {
        return Ok(None);
    };
    let Some((lo, hi)) = line.within(ctx) else {
        return Ok(None);
    };
    let (mut low, mut high) = (i64::from(lo), i64::from(hi));
    let mut kept: Option<u32> = None;
    while low <= high {
        let mid = (low + high) / 2;
        // The nearest cell of the kind toward low, else toward high.
        let mut at: Option<(i64, Value)> = None;
        let mut j = mid;
        while j >= low {
            let v = read_at(ctx, line, j as u32)?;
            if kind_of(&v) == Some(kind) {
                at = Some((j, v));
                break;
            }
            j -= 1;
        }
        if at.is_none() {
            let mut j = mid + 1;
            while j <= high {
                let v = read_at(ctx, line, j as u32)?;
                if kind_of(&v) == Some(kind) {
                    at = Some((j, v));
                    break;
                }
                j += 1;
            }
        }
        let Some((j, v)) = at else {
            break;
        };
        if keep(ordered(&v, probe)?) {
            kept = Some(j as u32);
            low = j + 1;
        } else {
            high = j - 1;
        }
    }
    Ok(kept)
}

/// The next cell of a value's kind after a position along a line, or the
/// first where `after` is `None`, reading as the bisection reads: what a
/// next larger or smaller by the bisection steps to.
pub(crate) fn next_of_kind(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    after: Option<u32>,
) -> Result<Option<u32>, Reason> {
    let Some(kind) = kind_of(probe) else {
        return Ok(None);
    };
    let Some((lo, hi)) = line.within(ctx) else {
        return Ok(None);
    };
    let from = after.map_or(lo, |a| (a + 1).max(lo));
    for k in from..=hi {
        if kind_of(&read_at(ctx, line, k)?) == Some(kind) {
            return Ok(Some(k));
        }
    }
    Ok(None)
}

// ---- the test switch ------------------------------------------------------

#[cfg(test)]
thread_local! {
    static CHECK: Cell<bool> = const { Cell::new(false) };
}

/// Whether every probe an index answers is held to its scan, on this thread
/// (the laws' switch).
#[cfg(test)]
pub(crate) fn checking() -> bool {
    CHECK.with(Cell::get)
}

/// Every probe an index answers held to its scan, on this thread, from now
/// on or no longer.
#[cfg(test)]
pub(crate) fn check(on: bool) {
    CHECK.with(|c| c.set(on));
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::{library, Calc};
    use crate::sheet::{Cell, Content, Workbook};

    struct Dice(u64);

    impl Dice {
        fn next(&mut self) -> u64 {
            self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
            let mut z = self.0;
            z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
            z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
            z ^ (z >> 31)
        }

        fn below(&mut self, n: usize) -> usize {
            (self.next() % n as u64) as usize
        }

        fn pick<'a>(&mut self, xs: &[&'a str]) -> &'a str {
            xs[self.below(xs.len())]
        }
    }

    /// A cell of a vector, of every kind a key tells apart or keeps alike:
    /// both zeros, numbers of both signs, two numbers equal to fifteen
    /// digits alone, texts in two cases and beyond ASCII, texts spelling
    /// numbers and the empty text, truth values, errors, a formula's error
    /// and a formula not computed; or nothing.
    fn content(d: &mut Dice) -> Option<Content> {
        Some(match d.below(16) {
            13 => Content::Formula("0.1+0.2".to_string()),
            14 => Content::Number(0.3),
            0 | 1 => return None,
            2 => Content::Number(0.0),
            3 => Content::Number(-0.0),
            4 | 5 => Content::Number(d.below(4) as f64),
            6 => Content::Number(-0.5 - d.below(2) as f64),
            7 => Content::Text(d.pick(&["a", "A", "b", "é", "É", "ab", "AB"]).to_string()),
            8 => Content::Text(d.pick(&["1", "2", " 2 ", "-0", "1E0", ""]).to_string()),
            9 => Content::Bool(d.below(2) == 0),
            10 => Content::Error(d.pick(&["#N/A", "#DIV/0!", "#VALUE!"]).to_string()),
            11 => Content::Formula("1/0".to_string()),
            12 if d.below(3) == 0 => Content::Formula("FACT(1)".to_string()),
            _ => Content::Number(d.below(4) as f64),
        })
    }

    /// What a probe searches for: a value of each kind the vectors hold, a
    /// cell of a vector, or an error.
    fn probe(d: &mut Dice) -> String {
        match d.below(6) {
            0 => format!("V!$A${}", 1 + d.below(24)),
            _ => d
                .pick(&[
                    "0", "-0", "1", "2", "3", "-0.5", "\"a\"", "\"A\"", "\"é\"", "\"É\"", "\"ab\"",
                    "\"1\"", "\"\"", "TRUE", "FALSE", "#N/A", "0.3", "0.1+0.2",
                ])
                .to_string(),
        }
    }

    /// A criterion: a probe, or a text with each of the operators, the
    /// blank, an error, a truth value, a pattern or a size, which the scan
    /// answers alone.
    fn criterion(d: &mut Dice) -> String {
        match d.below(2) {
            0 => probe(d),
            _ => d
                .pick(&[
                    "\"=1\"",
                    "\"<>a\"",
                    "\"<>\"",
                    "\"=\"",
                    "\"<>1\"",
                    "\"a*\"",
                    "\">1\"",
                    "\"#N/A\"",
                    "\"=TRUE\"",
                    "\"<>#DIV/0!\"",
                    "\"=-0\"",
                    "\"É\"",
                    "\"<>2\"",
                    "\"=0.3\"",
                    "\"<=0.3\"",
                    "\"<>0.3\"",
                ])
                .to_string(),
        }
    }

    /// A formula searching the vectors by every path an index answers, and
    /// by the bisection beside them.
    fn search(d: &mut Dice) -> String {
        let p = probe(d);
        let mode = d.pick(&["0", "-1", "1", "2"]);
        let search = d.pick(&["1", "-1", "1", "-1", "2", "-2"]);
        match d.below(8) {
            0 => format!("MATCH({p},V!$A$1:$A$24,{})", d.pick(&["0", "0", "1", "-1"])),
            1 => format!("XMATCH({p},V!$A$1:$A$24,{mode},{search})"),
            2 => format!("XMATCH({p},V!$C$1:$Z$1,{mode},{search})"),
            3 => format!(
                "VLOOKUP({p},V!$A$1:$B$24,2,{})",
                d.pick(&["FALSE", "FALSE", "TRUE"])
            ),
            4 => format!("XLOOKUP({p},V!$A$1:$A$24,V!$B$1:$B$24,\"none\",{mode},{search})"),
            5 => format!(
                "COUNTIF({},{})",
                d.pick(&["V!$C$3:$F$8", "V!$A$1:$A$24", "V!$C$1:$Z$1"]),
                criterion(d)
            ),
            6 => format!("SUMIF(V!$A$1:$A$24,{},V!$B$1:$B$24)", criterion(d)),
            _ => format!("SUMIF(V!$C$3:$F$8,{})", criterion(d)),
        }
    }

    /// A workbook of the vectors on `V` and each formula on `P`, written
    /// twice, so that a vector searched by it is built at its second probe.
    fn workbook(d: &mut Dice, formulas: &[String]) -> Workbook {
        let mut wb = Workbook::new();
        let v = wb.ensure_sheet("V");
        let mut cells: Vec<(u32, u32)> = (1..=24).flat_map(|row| [(row, 1), (row, 2)]).collect();
        cells.extend((3..=26).map(|col| (1, col)));
        cells.extend((3..=8).flat_map(|row| (3..=6).map(move |col| (row, col))));
        for (row, col) in cells {
            if let Some(content) = content(d) {
                wb.sheets[v].set(row, col, Cell { content, style: 0 });
            }
        }
        let p = wb.ensure_sheet("P");
        for (i, f) in formulas.iter().enumerate() {
            for col in [1, 2] {
                let content = Content::Formula(f.clone());
                wb.sheets[p].set(i as u32 + 1, col, Cell { content, style: 0 });
            }
        }
        wb
    }

    #[test]
    fn an_index_answers_every_probe_as_its_scan() {
        // KERNEL.8: vectors of every kind a key tells apart or keeps alike,
        // a column, a row and a block, searched by every path an index
        // answers, each formula twice: every answer an index gives is held
        // to the scan as it is given (the switch), and the two cells of a
        // formula answer alike, the first by the scan and the second by the
        // index built at it.
        check(true);
        let (mut builds, mut answered) = (0, 0);
        for seed in 1..=150u64 {
            let mut d = Dice(seed);
            let formulas: Vec<String> = (0..40).map(|_| search(&mut d)).collect();
            let wb = workbook(&mut d, &formulas);
            let lib = library::language();
            let calc = Calc::run(&wb, &lib);
            let p = wb.find_sheet("P").expect("P");
            for (i, f) in formulas.iter().enumerate() {
                let row = i as u32 + 1;
                assert_eq!(
                    calc.computed((p, row, 1)),
                    calc.computed((p, row, 2)),
                    "seed {seed}: ={f}"
                );
            }
            let (b, _, a) = calc.indexes.counted();
            builds += b;
            answered += a;
        }
        check(false);
        assert!(
            builds > 900 && answered > builds * 6,
            "{builds} indexes built and {answered} probes answered: too few for the law to \
             hold over"
        );
    }

    #[test]
    fn a_vector_is_built_at_its_second_probe_and_once() {
        // KERNEL.8: in one pass of `Calc`, a vector searched once is never
        // built, and one searched by many cells is built at the second and
        // answers the rest; a criterion's range is keyed apart from a
        // lookup's vector over the same cells, and a sum's with its sum
        // range; a pattern and a size key nothing.
        let counted = |formulas: &[&str]| {
            let mut wb = Workbook::new();
            let v = wb.ensure_sheet("V");
            for row in 1..=4 {
                let content = Content::Number(f64::from(row));
                wb.sheets[v].set(row, 1, Cell { content, style: 0 });
            }
            let p = wb.ensure_sheet("P");
            for (i, f) in formulas.iter().enumerate() {
                let content = Content::Formula(f.to_string());
                wb.sheets[p].set(i as u32 + 1, 1, Cell { content, style: 0 });
            }
            let lib = library::language();
            let calc = Calc::run(&wb, &lib);
            let (builds, reached, answered) = calc.indexes.counted();
            assert_eq!(reached, 0, "nothing changes in a pass");
            (builds, answered)
        };
        let m = "MATCH(2,V!$A$1:$A$4,0)";
        assert_eq!(counted(&[m]), (0, 0));
        assert_eq!(counted(&[m, m]), (1, 1));
        assert_eq!(counted(&[m, m, m, "MATCH(3,V!$A$1:$A$4,0)"]), (1, 3));
        let c = "COUNTIF(V!$A$1:$A$4,2)";
        assert_eq!(counted(&[m, c, m, c]), (2, 2));
        let s = "SUMIF(V!$A$1:$A$4,2,V!$A$1:$A$4)";
        let t = "SUMIF(V!$A$1:$A$4,2)";
        assert_eq!(counted(&[s, t, s, t]), (1, 3), "one sum range, read alike");
        assert_eq!(counted(&["COUNTIF(V!$A$1:$A$4,\"<3\")"; 3]), (0, 0));
        assert_eq!(counted(&["COUNTIF(V!$A$1:$A$4,\"1*\")"; 3]), (0, 0));
        assert_eq!(counted(&["MATCH(2,V!$A$1:$A$4,1)"; 3]), (0, 0));
        // A vector is its rectangle, however its reference is written.
        assert_eq!(
            counted(&["MATCH(2,V!A1:A4,0)", "MATCH(2,V!$A$1:$A$4,0)"]),
            (1, 1)
        );
    }

    #[test]
    fn a_text_is_keyed_folded_and_found_as_written() {
        // A key's hash is its folded bytes, however long the text, so that a
        // text keyed folded and one probed as written are found alike, beyond
        // ASCII too, a character folding to one of another length among
        // them; and a number's two zeros are one key.
        let long = "Ab".repeat(100);
        let mut keys = Keys::default();
        let (i, new) = keys.add(Probe::Text(&long));
        assert!(new);
        assert_eq!(keys.find(Probe::Text(&long.to_ascii_lowercase())), Some(i));
        assert_eq!(keys.find(Probe::Text(&long.to_ascii_uppercase())), Some(i));
        assert_eq!(keys.find(Probe::Text("ab")), None);
        let (z, _) = keys.add(Probe::Number(bits(0.0)));
        assert_eq!(keys.find(Probe::Number(bits(-0.0))), Some(z));
        assert_eq!(keys.add(Probe::Text("É")), (2, true));
        assert_eq!(
            keys.find(Probe::Text("é")),
            Some(2),
            "a case beyond ASCII too"
        );
        assert_eq!(keys.find(Probe::Text("e")), None, "an accent is no case");
        assert_eq!(keys.find(Probe::Bool(true)), None);
        assert_eq!(keys.add(Probe::Blank), (3, true));
        assert_eq!(keys.find(Probe::Blank), Some(3));
        let kelvin = "\u{212A}".repeat(30) + &"É".repeat(40);
        let (k, _) = keys.add(Probe::Text(&kelvin));
        let folded = "k".repeat(30) + &"é".repeat(40);
        assert_eq!(keys.find(Probe::Text(&folded)), Some(k));
        assert_eq!(keys.find(Probe::Text(&folded.to_uppercase())), Some(k));
        // A text spelling a number is a key apart from the number, and a
        // criterion's number is its fifteen digits.
        assert_eq!(keys.find(Probe::Spelled(bits(0.0))), None);
        assert_eq!(fifteen(0.1 + 0.2), fifteen(0.3));
        assert_ne!(fifteen(0.3), fifteen(0.300_000_000_000_1));
        assert_eq!(fifteen(-0.0), fifteen(0.0));
        assert_eq!(fifteen(1e15 + 1.0), fifteen(1e15));
    }
}
