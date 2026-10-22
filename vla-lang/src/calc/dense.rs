//! A sheet's values kept as an array over its extent (KERNEL.20, slice 1,
//! 2026-10-10): a read of a cell is an index into it and a range a walk of
//! it, where a sheet's cells are a B-tree looked up one cell at a time.
//!
//! An engine's machine keeps its sheets' values in such arrays, by place
//! (KERNEL.49, `crate::machine`'s stores): each rows' sheet's, every value
//! of the frame a slot, and each `.last` twin's that a formula reads, filled
//! at every frame's end, since most of a frame's reads are of a twin: Life's
//! rule reads 21 cells of `Screen.last` for each cell it computes, 18 of them
//! in two ranges (Alonzo's `CART.1`, its profile). The evaluator reads them
//! through [`super::Env::dense`]. A slot holds what a read of its cell
//! answers, a formula's reason among them, or nothing for a cell that holds
//! nothing, which a walk skips as a walk of a sheet's cells skips a cell that
//! is not there; so a walk of the array, clamped to the sheet's extent, is
//! the walk of the cells, value for value and in the same order, and an
//! aggregate folding over either gives the same bits.
//!
//! The array is never larger than its owner allows (`limit`), so that a sheet
//! of two cells far apart is not a million empty slots; such a sheet's store
//! keeps arrays down its dense columns alone and every other cell by place.
//!
//! A slot is two numbers (KERNEL.49's decision 4): its number, eight bytes,
//! and its kind, one byte, a number or nothing, a text, a truth value or an
//! error by its kind; a text's string kept beside, the slot's number naming
//! its place among the texts, which only a slot of a number is ever read as.
//! Nine bytes a slot, where a [`Value`] of 24 and the numbers kept beside it
//! for slices, nine more (KERNEL.43), were 33. For each row of the extent the
//! store keeps how many of its slots hold no number (KERNEL.43), so that an
//! evaluation of many cells side by side in one row reads a reference's cells
//! as one slice of numbers, and the slots one at a time where any of them
//! holds something else.
//!
//! A range of the frame's own cells is folded and walked by index as a
//! twin's is (KERNEL.31), and a run's cells read one by slices: side by side
//! in a row, a slice of one row (KERNEL.43), and one below the other in a
//! column, the slots of one column read at the store's width.

use super::eval::{Computed, Reason};
use super::graph::{Area, Extent};
use super::value::{ErrorKind, Value};
use crate::sheet::Content;

/// The most slots a store holds for each cell of its sheet (KERNEL.20): a
/// slot is nine bytes (KERNEL.49), and a cell of a sheet's B-tree about 137
/// (measured 2026-10-10), so a store within this is smaller than the cells it
/// shadows, and a sheet sparser than that, two cells far apart, is read
/// through its cells.
pub(crate) const SLOTS_A_CELL: u64 = 4;

/// The most slots a store of a sheet of `cells` cells holds: [`SLOTS_A_CELL`]
/// a cell, and never more than `most`, a grid's own limit. A twin's rule, and
/// a frame's own sheet's (KERNEL.31), so that a machine and a door's census
/// of the same grid keep the same sheets as arrays.
pub(crate) fn limit(cells: usize, most: u64) -> u64 {
    (cells as u64).saturating_mul(SLOTS_A_CELL).min(most)
}

/// Whether a store over an extent is within a limit, as [`Dense::over`]
/// asks: a sheet that holds a cell, over no more cells than `limit`.
pub(crate) fn fits(extent: Extent, limit: u64) -> bool {
    extent.is_some_and(|((top, left), (bottom, right))| {
        u64::from(right - left + 1) * u64::from(bottom - top + 1) <= limit
    })
}

/// A slot's kind (KERNEL.49's decision 4): a number, which a read by slice
/// reads; nothing; a text, its string among the store's texts; a truth
/// value; a formula's reason, why it was not computed, among the store's
/// reasons; a formula with no value yet; or an error, by its place in
/// [`ERRORS`] past [`ERROR`].
const NUMBER: u8 = 0;
const EMPTY: u8 = 1;
const TEXT: u8 = 2;
const FALSE: u8 = 3;
const TRUE: u8 = 4;
const REASON: u8 = 5;
const PENDING: u8 = 6;
const ERROR: u8 = 8;

/// The error kinds a slot holds, each at its place past [`ERROR`].
const ERRORS: [ErrorKind; 8] = [
    ErrorKind::Null,
    ErrorKind::Div0,
    ErrorKind::Value,
    ErrorKind::Ref,
    ErrorKind::Name,
    ErrorKind::Num,
    ErrorKind::NA,
    ErrorKind::Calc,
];

/// An error's kind as a slot holds it.
#[inline]
fn error_kind(k: ErrorKind) -> u8 {
    ERROR
        + match k {
            ErrorKind::Null => 0,
            ErrorKind::Div0 => 1,
            ErrorKind::Value => 2,
            ErrorKind::Ref => 3,
            ErrorKind::Name => 4,
            ErrorKind::Num => 5,
            ErrorKind::NA => 6,
            ErrorKind::Calc => 7,
        }
}

/// The blocks a store keeps its texts in (KERNEL.49): the first of
/// [`FIRST`] places, each after it twice the one before up to [`TEXTS`], and
/// every one after those of [`TEXTS`]. The texts grow a block at a time and
/// never move, so that a sheet of many texts holds no more room than a block
/// past its texts and no copy of them at its peak (the netlist's program,
/// 40,984 texts on its first sheet, held 65,536 places when they grew by
/// doubling, and copied them as they grew), and a sheet of a few labels a
/// block of a few.
const FIRST: usize = 8;
const TEXTS: usize = 512;
/// The blocks that double, and the places they hold together.
const DOUBLING: usize = 7;
const DOUBLED: usize = FIRST * ((1 << DOUBLING) - 1);

/// A text's block and its place in it, from its place among the texts.
#[inline]
fn block_of(at: usize) -> (usize, usize) {
    if at < DOUBLED {
        let b = (usize::BITS - 1 - (at / FIRST + 1).leading_zeros()) as usize;
        (b, at - FIRST * ((1 << b) - 1))
    } else {
        (DOUBLING + (at - DOUBLED) / TEXTS, (at - DOUBLED) % TEXTS)
    }
}

/// How many texts a block holds.
#[inline]
fn block_size(b: usize) -> usize {
    if b < DOUBLING {
        FIRST << b
    } else {
        TEXTS
    }
}

/// What a slot holds, read where it stands (KERNEL.49): a text and a reason
/// lent from the store, every other kind by value, so that a read of a
/// number makes nothing.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Slot<'a> {
    Empty,
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Error(ErrorKind),
    /// A formula not computed, and why.
    Reason(&'a Reason),
    /// A formula with no value yet.
    Pending,
}

impl Slot<'_> {
    /// The value a slot holding a value holds, a text copied; nothing for a
    /// reason and for a formula with no value, which [`Slot::read`] answers.
    pub(crate) fn value(self) -> Value {
        match self {
            Slot::Empty | Slot::Reason(_) | Slot::Pending => Value::Empty,
            Slot::Number(n) => Value::Number(n),
            Slot::Text(t) => Value::Text(t.to_string()),
            Slot::Bool(b) => Value::Bool(b),
            Slot::Error(k) => Value::Error(k),
        }
    }

    /// What a read of the slot's cell answers: its value; a formula's reason,
    /// which propagates; and for a formula with no value, a cycle, the one
    /// place a read can meet one, every cell being computed before what reads
    /// it.
    pub(crate) fn read(self) -> Result<Value, Reason> {
        match self {
            Slot::Reason(r) => Err(r.clone()),
            Slot::Pending => Err(Reason::Cycle),
            other => Ok(other.value()),
        }
    }

    /// What a formula's slot holds as its cell's computed value: `None` for
    /// one with no value yet.
    pub(crate) fn computed(self) -> Option<Computed> {
        match self {
            Slot::Pending => None,
            Slot::Reason(r) => Some(Computed::NotComputed(r.clone())),
            other => Some(Computed::Value(other.value())),
        }
    }

    /// Whether the slot holds nothing.
    pub(crate) fn is_empty(self) -> bool {
        matches!(self, Slot::Empty)
    }

    /// Whether two slots hold the same, a number to the bit.
    pub(crate) fn same(self, other: Slot<'_>) -> bool {
        match (self, other) {
            (Slot::Number(x), Slot::Number(y)) => x.to_bits() == y.to_bits(),
            (Slot::Empty, Slot::Empty) | (Slot::Pending, Slot::Pending) => true,
            (Slot::Text(x), Slot::Text(y)) => x == y,
            (Slot::Bool(x), Slot::Bool(y)) => x == y,
            (Slot::Error(x), Slot::Error(y)) => x == y,
            (Slot::Reason(x), Slot::Reason(y)) => x == y,
            _ => false,
        }
    }
}

/// A sheet's values over its extent, row by row. See the module's comment.
#[derive(Clone, Debug)]
pub struct Dense {
    top: u32,
    left: u32,
    width: usize,
    height: usize,
    /// Each slot's number: the number it holds, 0 where it holds nothing or
    /// a truth value or an error, and for a text the place of its string
    /// among `texts`, which no read takes for a number, a slice of numbers
    /// being read only where every slot of it is a number.
    numbers: Vec<f64>,
    /// Each slot's kind.
    kinds: Vec<u8>,
    /// By row of the extent, how many of its slots hold no number
    /// (KERNEL.43).
    odd: Vec<u32>,
    /// The texts the slots hold, each a [`Value::Text`] so that a walk lends
    /// it where it stands, in blocks ([`block_of`]); a place no slot names
    /// holds nothing and is in `free`.
    texts: Vec<Vec<Value>>,
    free: Vec<u32>,
    /// The reasons the slots hold, a formula's not computed (KERNEL.49), and
    /// the places among them no slot names; few by construction.
    reasons: Vec<Reason>,
    free_reasons: Vec<u32>,
}

impl PartialEq for Dense {
    /// Two stores over the same extent holding the same slots, a number to
    /// the bit, wherever their texts are kept.
    fn eq(&self, other: &Dense) -> bool {
        (self.top, self.left, self.width, self.height)
            == (other.top, other.left, other.width, other.height)
            && (0..self.kinds.len()).all(|i| self.slot(i).same(other.slot(i)))
    }
}

impl Dense {
    /// A store over an extent, in an old store's memory when one is given:
    /// the old store as it stands when its extent is the same, for the walk
    /// that keeps it to overwrite slot by slot, so that a text kept from one
    /// frame to the next is not copied again; every slot empty otherwise.
    /// `None` for a sheet with no cell, or for one whose extent holds more
    /// than `limit` cells.
    pub(crate) fn over(extent: Extent, limit: u64, old: Option<Dense>) -> Option<Dense> {
        let ((top, left), (bottom, right)) = extent?;
        if !fits(extent, limit) {
            return None;
        }
        let width = (right - left + 1) as usize;
        let height = (bottom - top + 1) as usize;
        let mut d = match old {
            Some(d) if d.stands_over(extent) => return Some(d),
            Some(d) => d,
            None => Dense {
                top,
                left,
                width,
                height,
                numbers: Vec::new(),
                kinds: Vec::new(),
                odd: Vec::new(),
                texts: Vec::new(),
                free: Vec::new(),
                reasons: Vec::new(),
                free_reasons: Vec::new(),
            },
        };
        (d.top, d.left, d.width, d.height) = (top, left, width, height);
        d.numbers.clear();
        d.numbers.resize(width * height, 0.0);
        d.kinds.clear();
        d.kinds.resize(width * height, EMPTY);
        d.odd.clear();
        d.odd.resize(height, width as u32);
        d.texts.clear();
        d.free.clear();
        d.reasons.clear();
        d.free_reasons.clear();
        Some(d)
    }

    /// Whether the store stands over an extent, as [`Dense::over`] keeps it
    /// whole when it does (KERNEL.36: the walk reads a slot before it
    /// overwrites it only in a store kept so).
    pub(crate) fn stands_over(&self, extent: Extent) -> bool {
        extent.is_some_and(|((top, left), (bottom, right))| {
            let width = (right - left + 1) as usize;
            let height = (bottom - top + 1) as usize;
            (self.top, self.left, self.width, self.height) == (top, left, width, height)
        })
    }

    /// One cell's slot overwritten with a value; `false` for a cell outside
    /// the extent, which a store made over its sheet's own extent never
    /// meets. The tests'; the walk finds a slot's index first, to read the bit
    /// of a twin's cell behind it (KERNEL.36).
    #[cfg(test)]
    pub(crate) fn set(&mut self, row: u32, col: u32, v: Value) -> bool {
        let Some(i) = self.index(row, col) else {
            return false;
        };
        self.put_value(i, &v);
        true
    }

    /// Where the store's slots stand in memory, for the tests that hold a
    /// store kept over its extent to its own memory (KERNEL.31).
    #[cfg(test)]
    pub(crate) fn memory(&self) -> *const f64 {
        self.numbers.as_ptr()
    }

    /// A cell's value: inside the extent, what a read of the cell answers,
    /// `Value::Empty` for a cell that holds nothing, a text copied; `None`
    /// outside it.
    pub fn get(&self, row: u32, col: u32) -> Option<Value> {
        self.index(row, col).map(|i| self.slot(i).value())
    }

    /// A cell's slot where it stands, inside the extent; `None` outside it.
    #[inline]
    pub(crate) fn slot_of(&self, row: u32, col: u32) -> Option<Slot<'_>> {
        self.index(row, col).map(|i| self.slot(i))
    }

    /// The text at a place among the store's texts.
    #[inline]
    fn text(&self, at: u64) -> &Value {
        let (b, k) = block_of(at as usize);
        &self.texts[b][k]
    }

    /// How many places the store's texts hold, those free among them.
    #[cfg(test)]
    fn texts_held(&self) -> usize {
        self.texts.iter().map(Vec::len).sum()
    }

    /// A cell's slot's index, row by row; `None` outside the extent.
    #[inline]
    pub(crate) fn index(&self, row: u32, col: u32) -> Option<usize> {
        let r = row.checked_sub(self.top)? as usize;
        let c = col.checked_sub(self.left)? as usize;
        (r < self.height && c < self.width).then(|| r * self.width + c)
    }

    /// How many slots the store holds, one a cell of its extent.
    pub(crate) fn slots(&self) -> usize {
        self.kinds.len()
    }

    /// The slot at an index as it stands.
    #[inline]
    pub(crate) fn slot(&self, i: usize) -> Slot<'_> {
        match self.kinds[i] {
            NUMBER => Slot::Number(self.numbers[i]),
            EMPTY => Slot::Empty,
            TEXT => match self.text(self.numbers[i].to_bits()) {
                Value::Text(t) => Slot::Text(t),
                _ => Slot::Empty,
            },
            FALSE => Slot::Bool(false),
            TRUE => Slot::Bool(true),
            REASON => Slot::Reason(&self.reasons[self.numbers[i].to_bits() as usize]),
            PENDING => Slot::Pending,
            k => Slot::Error(ERRORS[usize::from(k - ERROR)]),
        }
    }

    /// The slot at an index given a kind, its row's count of slots holding
    /// no number kept, and a text it held given up.
    #[inline]
    fn kind(&mut self, i: usize, kind: u8) {
        let was = self.kinds[i];
        if was == kind {
            return;
        }
        if was == TEXT {
            let at = self.numbers[i].to_bits() as usize;
            let (b, k) = block_of(at);
            self.texts[b][k] = Value::Empty;
            self.free.push(at as u32);
        }
        if was == REASON {
            let at = self.numbers[i].to_bits() as usize;
            self.reasons[at] = Reason::Cycle;
            self.free_reasons.push(at as u32);
        }
        if (was == NUMBER) != (kind == NUMBER) {
            let row = i / self.width;
            if kind == NUMBER {
                self.odd[row] -= 1;
            } else {
                self.odd[row] += 1;
            }
        }
        self.kinds[i] = kind;
    }

    /// The slot at an index overwritten with a number.
    #[inline]
    pub(crate) fn put_number(&mut self, i: usize, n: f64) {
        self.kind(i, NUMBER);
        self.numbers[i] = n;
    }

    /// The slot at an index overwritten with nothing.
    #[inline]
    pub(crate) fn put_empty(&mut self, i: usize) {
        self.kind(i, EMPTY);
        self.numbers[i] = 0.0;
    }

    /// The slot at an index overwritten with a text, kept where it stands
    /// when the slot holds it already and never copied again, so that a frame
    /// allocates no more than it did before the array (KERNEL.20).
    fn put_text(&mut self, i: usize, t: &str) {
        if self.kinds[i] == TEXT {
            let (b, k) = block_of(self.numbers[i].to_bits() as usize);
            if let Value::Text(held) = &mut self.texts[b][k] {
                if held.as_str() != t {
                    held.clear();
                    held.push_str(t);
                }
            }
            return;
        }
        self.kind(i, TEXT);
        let text = Value::Text(t.to_string());
        let at = match self.free.pop() {
            Some(at) => {
                let (b, k) = block_of(at as usize);
                self.texts[b][k] = text;
                at as usize
            }
            None => {
                let full = |b: usize, block: &Vec<Value>| block.len() == block_size(b);
                let b = self.texts.len();
                if b == 0 || full(b - 1, &self.texts[b - 1]) {
                    self.texts.push(Vec::with_capacity(block_size(b)));
                }
                let last = self.texts.len() - 1;
                self.texts[last].push(text);
                let before: usize = (0..last).map(block_size).sum();
                before + self.texts[last].len() - 1
            }
        };
        self.numbers[i] = f64::from_bits(at as u64);
    }

    /// The slot at an index overwritten with a value.
    #[inline]
    pub(crate) fn put_value(&mut self, i: usize, v: &Value) {
        match v {
            Value::Number(n) => self.put_number(i, *n),
            Value::Empty => self.put_empty(i),
            Value::Text(t) => self.put_text(i, t),
            Value::Bool(b) => {
                self.kind(i, if *b { TRUE } else { FALSE });
                self.numbers[i] = 0.0;
            }
            Value::Error(k) => {
                self.kind(i, error_kind(*k));
                self.numbers[i] = 0.0;
            }
        }
    }

    /// The slot at an index overwritten with a formula's reason, kept where
    /// it stands when the slot holds it already.
    pub(crate) fn put_reason(&mut self, i: usize, r: &Reason) {
        if self.kinds[i] == REASON {
            let held = &mut self.reasons[self.numbers[i].to_bits() as usize];
            if held != r {
                held.clone_from(r);
            }
            return;
        }
        self.kind(i, REASON);
        let at = match self.free_reasons.pop() {
            Some(at) => {
                self.reasons[at as usize] = r.clone();
                at as usize
            }
            None => {
                self.reasons.push(r.clone());
                self.reasons.len() - 1
            }
        };
        self.numbers[i] = f64::from_bits(at as u64);
    }

    /// The slot at an index made a formula's with no value yet.
    pub(crate) fn put_pending(&mut self, i: usize) {
        self.kind(i, PENDING);
        self.numbers[i] = 0.0;
    }

    /// Every formula's slot with no value yet made empty, as a twin holds it
    /// (KERNEL.49): its twin cell holds nothing, which a read answers. Both
    /// kinds hold no number, so each row's count stands.
    pub(crate) fn pending_emptied(&mut self) {
        for k in &mut self.kinds {
            if *k == PENDING {
                *k = EMPTY;
            }
        }
    }

    /// Whether two stores stand over one extent, slot for slot.
    pub(crate) fn same_extent(&self, other: &Dense) -> bool {
        (self.top, self.left, self.width, self.height)
            == (other.top, other.left, other.width, other.height)
    }

    /// The slot at an index overwritten with a formula's computed value: its
    /// value, its reason, or none yet.
    pub(crate) fn put_computed(&mut self, i: usize, computed: Option<&Computed>) {
        match computed {
            Some(Computed::Value(v)) => self.put_value(i, v),
            Some(Computed::NotComputed(r)) => self.put_reason(i, r),
            None => self.put_pending(i),
        }
    }

    /// The slot at an index overwritten with what a read of a cell's content
    /// answers, nothing for a cell that keeps no value, as a twin's slot is
    /// written from its cell (KERNEL.20): a value as it stands, an error text
    /// as its kind (one this version does not hold as `#VALUE!`), nothing for
    /// a formula or anything else.
    pub(crate) fn put_content(&mut self, i: usize, content: Option<&Content>) {
        match content {
            Some(Content::Text(t)) => self.put_text(i, t),
            Some(Content::Number(n)) => self.put_number(i, *n),
            Some(Content::Bool(b)) => self.put_value(i, &Value::Bool(*b)),
            Some(Content::Error(e)) => self.put_value(
                i,
                &Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Value)),
            ),
            _ => self.put_empty(i),
        }
    }

    /// Of `n` cells side by side whose reference covers rows `top` to
    /// `bottom` and `width` columns from `col` plus the cell's place, the
    /// places `a..b` of those it keeps inside the extent (KERNEL.43), every
    /// place where it `moves`, or none or all where it stands still: so that
    /// a run's cells whose reference leaves the array at the row's ends are
    /// read alone and the rest by slices.
    pub(crate) fn inside(
        &self,
        (top, bottom): (u32, u32),
        col: u32,
        width: usize,
        n: usize,
        moves: bool,
    ) -> (usize, usize) {
        let (first, last) = (self.top, self.top + (self.height as u32 - 1));
        if top < first || bottom > last {
            return (0, 0);
        }
        let (left, right) = (
            i64::from(self.left),
            i64::from(self.left) + self.width as i64 - 1,
        );
        let (col, width) = (i64::from(col), width as i64);
        if !moves {
            let fits = col >= left && col + width - 1 <= right;
            return if fits { (0, n) } else { (0, 0) };
        }
        let a = (left - col).clamp(0, n as i64);
        let b = (right - (col + width - 1) + 1).clamp(0, n as i64);
        if a >= b {
            (0, 0)
        } else {
            (a as usize, b as usize)
        }
    }

    /// The numbers of `n` slots side by side in one row, from a cell's
    /// (KERNEL.43): what an evaluation of many cells reads a reference's
    /// cells of one row as, when every one is inside the extent and holds a
    /// number; `None` otherwise, its slots read one at a time.
    #[inline]
    pub(crate) fn numbers_in_row(&self, row: u32, col: u32, n: usize) -> Option<&[f64]> {
        let r = row.checked_sub(self.top)? as usize;
        let c = col.checked_sub(self.left)? as usize;
        if r >= self.height || c + n > self.width {
            return None;
        }
        let i = r * self.width + c;
        if self.odd[r] != 0 && self.kinds[i..i + n].iter().any(|&k| k != NUMBER) {
            return None;
        }
        Some(&self.numbers[i..i + n])
    }

    /// The numbers of `n` slots one below the other in one column, from a
    /// cell's, pushed to `out` (KERNEL.31): what an evaluation of many cells
    /// reads a reference's cells as where its cells stand one below the
    /// other, when every one is inside the extent and holds a number;
    /// `false`, nothing pushed, otherwise, its slots read one at a time.
    #[inline]
    pub(crate) fn numbers_in_column(
        &self,
        row: u32,
        col: u32,
        n: usize,
        out: &mut Vec<f64>,
    ) -> bool {
        let (Some(r), Some(c)) = (row.checked_sub(self.top), col.checked_sub(self.left)) else {
            return false;
        };
        let (r, c, w) = (r as usize, c as usize, self.width);
        if r + n > self.height || c >= w {
            return false;
        }
        let first = r * w + c;
        if self.odd[r..r + n].iter().any(|&o| o != 0)
            && (0..n).any(|j| self.kinds[first + j * w] != NUMBER)
        {
            return false;
        }
        out.extend((0..n).map(|j| self.numbers[first + j * w]));
        true
    }

    /// The slot at an index, and its cell (KERNEL.27).
    pub(crate) fn at(&self, i: usize) -> (Slot<'_>, (u32, u32)) {
        let (r, c) = (i / self.width, i % self.width);
        (self.slot(i), (self.top + r as u32, self.left + c as u32))
    }

    /// Every value an area holds, row by row, to `f`, the area clamped to the
    /// extent first, as a walk of the cells clamps it; an empty slot, and a
    /// formula's with no value, skipped; a reason answered where it stands,
    /// as a walk of the cells meets it (KERNEL.49). `f` is a type of its own
    /// and not a pointer, so that an aggregate's fold is compiled into the
    /// walk (KERNEL.25).
    pub fn each<F>(&self, area: &Area, mut f: F) -> Result<(), Reason>
    where
        F: FnMut(Value) -> Result<(), Reason>,
    {
        self.each_read(area, |_, _, read| match read {
            Ok(v) => f(v.clone()),
            Err(r) => Err(r.clone()),
        })
    }

    /// [`Dense::each`]'s walk, each read lent with its row and its column: a
    /// value, a text where it stands and any other kind as the slot makes it,
    /// or a formula's reason (KERNEL.8; KERNEL.49), so that an evaluation of
    /// many cells at once folds an aggregate over the store copying nothing
    /// (KERNEL.26), and a vector's index is built from it and a lookup's scan
    /// walks it, its positions read off the cells.
    pub(crate) fn each_read<E, F>(&self, area: &Area, mut f: F) -> Result<(), E>
    where
        F: FnMut(u32, u32, Result<&Value, &Reason>) -> Result<(), E>,
    {
        let bottom = self.top + (self.height as u32 - 1);
        let right = self.left + (self.width as u32 - 1);
        let (top, left) = (area.top.max(self.top), area.left.max(self.left));
        let (last_row, last_col) = (area.bottom.min(bottom), area.right.min(right));
        if top > last_row || left > last_col {
            return Ok(());
        }
        let n = (last_col - left + 1) as usize;
        for row in top..=last_row {
            let from = (row - self.top) as usize * self.width + (left - self.left) as usize;
            for j in 0..n {
                let i = from + j;
                let col = left + j as u32;
                match self.kinds[i] {
                    NUMBER => f(row, col, Ok(&Value::Number(self.numbers[i])))?,
                    EMPTY | PENDING => {}
                    TEXT => f(row, col, Ok(self.text(self.numbers[i].to_bits())))?,
                    REASON => f(
                        row,
                        col,
                        Err(&self.reasons[self.numbers[i].to_bits() as usize]),
                    )?,
                    _ => f(row, col, Ok(&self.slot(i).value()))?,
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{MAX_COLUMN, MAX_ROW};

    fn rect(top: u32, left: u32, bottom: u32, right: u32) -> Area {
        Area {
            sheet: 0,
            top,
            left,
            bottom,
            right,
        }
    }

    fn walk(d: &Dense, area: Area) -> Vec<Value> {
        let mut out = Vec::new();
        d.each(&area, |v| {
            out.push(v);
            Ok(())
        })
        .unwrap();
        out
    }

    #[test]
    fn a_store_is_its_extent_row_by_row_and_skips_what_is_empty() {
        // B2:D4, every cell but C3 holding a value, one of each kind.
        let mut d = Dense::over(Some(((2, 2), (4, 4))), 9, None).unwrap();
        let n = Value::Number;
        let kinds = [
            n(1.0),
            Value::Text("x".to_string()),
            Value::Bool(true),
            n(4.0),
            Value::Error(ErrorKind::NA),
            n(-0.5),
            n(7.0),
            n(8.0),
        ];
        let cells = [
            (2, 2),
            (2, 3),
            (2, 4),
            (3, 2),
            (3, 4),
            (4, 2),
            (4, 3),
            (4, 4),
        ];
        for ((row, col), v) in cells.iter().zip(&kinds) {
            assert!(d.set(*row, *col, v.clone()));
        }
        assert!(!d.set(1, 2, Value::Empty), "above the extent");
        assert!(!d.set(2, 5, Value::Empty), "right of it");
        assert_eq!(d.get(2, 3), Some(kinds[1].clone()));
        assert_eq!(d.get(3, 3), Some(Value::Empty), "a cell holding nothing");
        assert_eq!(d.get(5, 2), None, "outside");
        assert_eq!(d.get(2, 1), None);
        // The extent, then areas past it, across it and outside it, each
        // clamped as a walk of the cells clamps it.
        assert_eq!(walk(&d, rect(2, 2, 4, 4)), kinds.to_vec());
        assert_eq!(walk(&d, rect(1, 1, MAX_ROW, MAX_COLUMN)), kinds.to_vec());
        assert_eq!(
            walk(&d, rect(3, 3, 9, 9)),
            vec![kinds[4].clone(), n(7.0), n(8.0)]
        );
        assert_eq!(walk(&d, rect(1, 3, 3, 3)), vec![kinds[1].clone()]);
        assert_eq!(walk(&d, rect(3, 3, 3, 3)), Vec::<Value>::new());
        assert_eq!(walk(&d, rect(5, 1, 9, 9)), Vec::<Value>::new());
        assert_eq!(walk(&d, rect(1, 5, 1, 9)), Vec::<Value>::new());
        // A fold stops at its first error, as a walk of the cells does.
        let mut seen = 0;
        let stopped = d.each_read(&rect(2, 2, 4, 4), |_, _, read| {
            seen += 1;
            if matches!(read, Ok(Value::Bool(_))) {
                Err("stop")
            } else {
                Ok(())
            }
        });
        assert_eq!((stopped, seen), (Err("stop"), 3));
    }

    #[test]
    fn a_reason_answers_where_it_stands_and_a_formula_with_no_value_is_passed() {
        // KERNEL.49: a formula's reason is a slot's kind, which a walk meets
        // where it stands, after the values before it, and a read answers; a
        // formula with no value yet a walk passes as nothing and a read
        // answers as a cycle, the one place a read can meet one; and a
        // reason's place given up is taken again.
        let mut d = Dense::over(Some(((1, 1), (1, 4))), 4, None).unwrap();
        d.set(1, 1, Value::Number(1.0));
        let r = Reason::Function("FACT".to_string());
        d.put_reason(d.index(1, 2).unwrap(), &r);
        d.put_pending(d.index(1, 3).unwrap());
        d.set(1, 4, Value::Number(4.0));
        let mut seen = Vec::new();
        let walked = d.each(&rect(1, 1, 1, 4), |v| {
            seen.push(v);
            Ok(())
        });
        assert_eq!(walked, Err(r.clone()));
        assert_eq!(seen, vec![Value::Number(1.0)], "the values before it");
        assert_eq!(d.slot_of(1, 2).unwrap().read(), Err(r.clone()));
        assert_eq!(d.slot_of(1, 3).unwrap().read(), Err(Reason::Cycle));
        assert_eq!(d.slot_of(1, 3).unwrap().computed(), None);
        assert_eq!(
            d.slot_of(1, 2).unwrap().computed(),
            Some(Computed::NotComputed(r.clone()))
        );
        assert_eq!(walk(&d, rect(1, 3, 1, 4)), vec![Value::Number(4.0)]);
        // The row holds no number at the reason, and holds one once it goes.
        assert_eq!(d.numbers_in_row(1, 1, 2), None);
        d.put_number(d.index(1, 2).unwrap(), 2.0);
        assert_eq!(d.numbers_in_row(1, 1, 2), Some(&[1.0, 2.0][..]));
        d.put_reason(d.index(1, 4).unwrap(), &Reason::Cycle);
        assert_eq!(d.reasons.len(), 1, "the place given up taken again");
        assert_eq!(d.slot_of(1, 4).unwrap().read(), Err(Reason::Cycle));
    }

    #[test]
    fn a_store_keeps_its_memory_and_refuses_what_passes_its_limit() {
        let extent = Some(((1, 1), (10, 10)));
        let mut d = Dense::over(extent, 100, None).unwrap();
        assert!(d.set(10, 10, Value::Text("kept".to_string())));
        let at = d.numbers.as_ptr();
        let Some(Slot::Text(text)) = d.slot_of(10, 10) else {
            panic!("a text")
        };
        let text_at = text.as_ptr();
        // The next frame's store over the same extent is the store as it
        // stands, for the walk to overwrite: a text in it is not copied, nor
        // when the same text is written over it.
        let mut same = Dense::over(extent, 100, Some(d)).unwrap();
        assert_eq!(same.numbers.as_ptr(), at);
        let i = same.index(10, 10).unwrap();
        same.put_value(i, &Value::Text("kept".to_string()));
        let Some(Slot::Text(text)) = same.slot_of(10, 10) else {
            panic!("the text kept")
        };
        assert_eq!(text.as_ptr(), text_at, "the same text, not a copy");
        // Over another extent, the same memory with every slot empty again.
        let again = Dense::over(Some(((3, 3), (5, 5))), 100, Some(same)).unwrap();
        assert_eq!(again.numbers.as_ptr(), at, "the old store's memory");
        assert_eq!(again.slots(), 9);
        assert!((0..9).all(|i| again.slot(i).is_empty()));
        assert!(again.texts_held() == 0 && again.free.is_empty());
        assert_eq!(again.get(3, 3), Some(Value::Empty));
        assert!(Dense::over(Some(((1, 1), (10, 10))), 99, None).is_none());
        assert!(
            Dense::over(None, 100, None).is_none(),
            "a sheet with no cell"
        );
        // The largest extent a sheet can have is refused, never allocated.
        assert!(Dense::over(Some(((1, 1), (MAX_ROW, MAX_COLUMN))), 1 << 20, None).is_none());
    }

    #[test]
    fn a_slot_is_nine_bytes_and_a_text_given_up_is_reused() {
        // KERNEL.49's decision 4: a slot's number and its kind, and the texts
        // beside, a text's place given up when its slot holds another kind
        // and taken again by the next text.
        let mut d = Dense::over(Some(((1, 1), (1, 4))), 4, None).unwrap();
        assert_eq!(
            std::mem::size_of_val(&d.numbers[0]) + std::mem::size_of_val(&d.kinds[0]),
            9
        );
        d.set(1, 1, Value::Text("a".to_string()));
        d.set(1, 2, Value::Text("b".to_string()));
        assert_eq!(d.texts_held(), 2);
        d.set(1, 1, Value::Number(3.0));
        assert_eq!(d.free, vec![0], "the first text's place given up");
        d.set(1, 3, Value::Text("c".to_string()));
        assert_eq!((d.texts_held(), d.free.len()), (2, 0), "and taken again");
        // The texts grow a block at a time and never move.
        let mut wide = Dense::over(Some(((1, 1), (1, 2000))), 2000, None).unwrap();
        for c in 1..=1200u32 {
            wide.set(1, c, Value::Text(format!("t{c}")));
        }
        // 8, 16, 32, 64, 128, 256 and 512 hold 1,016; one block of 512 more.
        assert_eq!(wide.texts.len(), 8, "eight blocks for 1,200 texts");
        for (b, block) in wide.texts.iter().enumerate() {
            assert_eq!(block.capacity(), block_size(b), "block {b}");
        }
        for c in 1..=1200u32 {
            assert_eq!(wide.get(1, c), Some(Value::Text(format!("t{c}"))));
        }
        // Every place maps to one block's place, in order.
        let mut want = (0, 0);
        for at in 0..5000 {
            assert_eq!(block_of(at), want, "place {at}");
            want.1 += 1;
            if want.1 == block_size(want.0) {
                want = (want.0 + 1, 0);
            }
        }
        assert_eq!(
            walk(&d, rect(1, 1, 1, 4)),
            vec![
                Value::Number(3.0),
                Value::Text("b".to_string()),
                Value::Text("c".to_string()),
            ]
        );
        // A text's place never reads as a number: the row holds texts.
        assert_eq!(d.numbers_in_row(1, 1, 3), None);
        assert_eq!(d.numbers_in_row(1, 1, 1), Some(&[3.0][..]));
    }

    #[test]
    fn the_cells_a_reference_keeps_inside_are_those_it_leaves_nowhere() {
        // KERNEL.43: over a store of B2:F6, for references of every width
        // moving with cells side by side, and standing still, from columns
        // before, inside and past the extent and rows inside and out, the
        // places `inside` answers are the cells whose reference stays wholly
        // inside the extent, found one by one.
        let d = Dense::over(Some(((2, 2), (6, 6))), 100, None).unwrap();
        for (top, bottom) in [(2, 2), (3, 5), (1, 3), (6, 7), (2, 6)] {
            for col in 0..=8u32 {
                for width in 1..=4usize {
                    for n in 1..=7usize {
                        for moves in [true, false] {
                            let fits = |i: usize| {
                                let left = col + if moves { i as u32 } else { 0 };
                                let right = left + width as u32 - 1;
                                top >= 2 && bottom <= 6 && left >= 2 && right <= 6
                            };
                            let want: Vec<usize> = (0..n).filter(|&i| fits(i)).collect();
                            let (a, b) = d.inside((top, bottom), col, width, n, moves);
                            assert_eq!(
                                (a..b).collect::<Vec<_>>(),
                                want,
                                "rows {top}-{bottom}, column {col}, width {width}, {n} cells, moving {moves}"
                            );
                        }
                    }
                }
            }
        }
    }

    /// What a read of slices must answer, from the slots alone: the numbers
    /// of `n` slots of a row from a cell, when each is inside the extent and
    /// holds a number.
    fn slice_of(d: &Dense, row: u32, col: u32, n: usize) -> Option<Vec<f64>> {
        (0..n)
            .map(|j| match d.get(row, col + j as u32)? {
                Value::Number(x) => Some(x),
                _ => None,
            })
            .collect()
    }

    /// What a read of a column must answer, from the slots alone: the
    /// numbers of `n` slots one below the other from a cell, when each is
    /// inside the extent and holds a number (KERNEL.31).
    fn column_of(d: &Dense, row: u32, col: u32, n: usize) -> Option<Vec<f64>> {
        (0..n)
            .map(|j| match d.get(row + j as u32, col)? {
                Value::Number(x) => Some(x),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn a_store_s_numbers_are_its_slots_after_every_write() {
        // KERNEL.43: slots written at random with values of every kind, the
        // store kept and made again over the same extent and another; after
        // each write, every row's count of slots holding no number is the
        // count read off its slots, every slot reads back the value written
        // last, and every slice a read may ask for, inside the extent, across
        // its edges and outside it, is what the slots say.
        let mut seed: u64 = 0x5EED;
        let mut next = |n: u64| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            (seed >> 33) % n
        };
        let values = [
            Value::Number(0.0),
            Value::Number(-0.0),
            Value::Number(1.5),
            Value::Number(-3.0),
            Value::Number(1e300),
            Value::Text("x".to_string()),
            Value::Text("yz".to_string()),
            Value::Bool(true),
            Value::Bool(false),
            Value::Error(ErrorKind::NA),
            Value::Error(ErrorKind::Calc),
            Value::Empty,
        ];
        let extents = [Some(((2, 3), (6, 9))), Some(((1, 1), (3, 4)))];
        let mut d = Dense::over(extents[0], 100, None).unwrap();
        let mut held: std::collections::HashMap<(u32, u32), Value> = Default::default();
        let mut columns = 0;
        for round in 0..2000 {
            // Every five hundredth round the store made again, over the same
            // extent or another, every slot empty where it moved.
            if round % 500 == 499 {
                let extent = extents[next(2) as usize];
                if !d.stands_over(extent) {
                    held.clear();
                }
                d = Dense::over(extent, 100, Some(d)).unwrap();
            }
            let (top, left) = (d.top, d.left);
            let (row, col) = (
                top + next(d.height as u64) as u32,
                left + next(d.width as u64) as u32,
            );
            let v = values[next(values.len() as u64) as usize].clone();
            assert!(d.set(row, col, v.clone()));
            held.insert((row, col), v);
            for r in 0..d.height {
                let row = top + r as u32;
                let odd = (0..d.width as u32)
                    .filter(|&c| !matches!(d.get(row, left + c), Some(Value::Number(_))))
                    .count();
                assert_eq!(d.odd[r] as usize, odd, "round {round}, row {row}");
                for c in 0..d.width as u32 {
                    let want = held.get(&(row, left + c)).cloned().unwrap_or(Value::Empty);
                    let got = d.get(row, left + c).unwrap();
                    let same = match (&got, &want) {
                        (Value::Number(x), Value::Number(y)) => x.to_bits() == y.to_bits(),
                        _ => got == want,
                    };
                    assert!(
                        same,
                        "round {round}: ({row}, {}) {got:?} for {want:?}",
                        left + c
                    );
                }
            }
            // Every text the slots name is held once, and every other place
            // is free.
            let named = d.kinds.iter().filter(|&&k| k == TEXT).count();
            assert_eq!(named + d.free.len(), d.texts_held(), "round {round}");
            for _ in 0..4 {
                let row = (top + next(d.height as u64 + 2) as u32).saturating_sub(1);
                let col = (left + next(d.width as u64 + 2) as u32).saturating_sub(1);
                let n = 1 + next(d.width as u64 + 1) as usize;
                let got = d
                    .numbers_in_row(row, col, n)
                    .map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                let want = slice_of(&d, row, col, n)
                    .map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                assert_eq!(got, want, "round {round}: {n} slots from ({row}, {col})");
                // And the slots one below the other from it (KERNEL.31),
                // pushed after what the lane holds already, or nothing.
                let n = 1 + next(d.height as u64 + 1) as usize;
                let mut out = vec![7.0];
                let read = d.numbers_in_column(row, col, n, &mut out);
                let want = column_of(&d, row, col, n);
                let got = read.then(|| out[1..].iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                assert_eq!(
                    got,
                    want.map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>()),
                    "round {round}: {n} slots down from ({row}, {col})"
                );
                assert_eq!(out.len(), if read { 1 + n } else { 1 }, "round {round}");
                columns += usize::from(read);
            }
        }
        assert!(columns > 0, "a column read whole somewhere");
    }
}
