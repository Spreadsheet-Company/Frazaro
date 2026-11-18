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
use super::formula::{Expr, RefExpr, ANCHORARRAY};
use super::graph::CellId;
use super::lambda::{self, Functions};
use crate::grid::Rects;
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

/// The spills a door plans (KERNEL.55): each dynamic formula's whose box holds
/// more than its anchor, by its anchor, and each sheet's boxes as rectangles
/// tagged by their anchor's place in `anchors`, so that a read and a range meet
/// them by a search and no cell is listed.
#[derive(Debug, Default)]
pub(crate) struct Spills {
    by_anchor: HashMap<CellId, Spill>,
    anchors: Vec<CellId>,
    boxes: Vec<Rects>,
}

impl Spills {
    /// Each dynamic formula's spill in a model: its bound by its parse, a spill
    /// read through `#` its anchor's, else the extent its file saved, else its
    /// anchor alone; none for a formula whose box holds one cell. Two whose
    /// boxes meet are labelled ([`MEETS`]), and one past a door's plan
    /// ([`NO_PLAN`]).
    pub(crate) fn of(
        model: &Workbook,
        texts: &HashMap<CellId, String>,
        functions: &Functions,
        lib: &Library,
    ) -> Spills {
        let mut anchors: Vec<CellId> = texts
            .keys()
            .copied()
            .filter(|&id| file_of(model, id).is_some())
            .collect();
        if anchors.is_empty() {
            return Spills::default();
        }
        anchors.sort_unstable();
        let mut bounds = Bounds {
            model,
            texts,
            functions,
            memo: HashMap::new(),
            visiting: HashSet::new(),
        };
        let mut spills = Spills {
            by_anchor: HashMap::new(),
            anchors: Vec::new(),
            boxes: Vec::new(),
        };
        for id in anchors {
            let file = file_of(model, id).flatten();
            // The parse's bound, the extent its file saved where it gives none
            // (decision 5), and the anchor alone where neither stands.
            let (rows, columns) = bounds.dims(id, lib).unwrap_or((1, 1));
            let dims = (rows.max(1), columns.max(1));
            let mut cells = A1Range {
                top: id.1,
                left: id.2,
                bottom: id.1.saturating_add(dims.0 - 1).min(MAX_ROW),
                right: id.2.saturating_add(dims.1 - 1).min(MAX_COLUMN),
            };
            if let Some(f) = file {
                cells.bottom = cells.bottom.max(f.bottom);
                cells.right = cells.right.max(f.right);
            }
            if cells.cells() == 1 {
                continue;
            }
            let label = (u64::from(dims.0) * u64::from(dims.1) > MOST).then_some(NO_PLAN);
            spills.by_anchor.insert(
                id,
                Spill {
                    dims,
                    file,
                    cells,
                    label,
                },
            );
            spills.anchors.push(id);
        }
        // Two boxes meeting, each labelled, the order Excel computes them in
        // deciding which spills; but where one anchor stands in the other's
        // box, which no order decides: the other's extent reaches a cell of
        // the one's only through its anchor, a formula, which blocks it.
        let mut met: Vec<CellId> = Vec::new();
        for (tag, &id) in spills.anchors.iter().enumerate() {
            let cells = spills.by_anchor[&id].cells;
            if spills.boxes.len() <= id.0 {
                spills.boxes.resize_with(id.0 + 1, Rects::new);
            }
            for (held, other) in spills.boxes[id.0].meeting(cells) {
                let other = spills.anchors[other as usize];
                if !inside(&held, id.1, id.2) && !inside(&cells, other.1, other.2) {
                    met.push(other);
                    met.push(id);
                }
            }
            spills.boxes[id.0].insert(cells, tag as u32);
        }
        for id in met {
            if let Some(s) = spills.by_anchor.get_mut(&id) {
                s.label = Some(MEETS);
            }
        }
        spills
    }

    /// Whether it holds none.
    pub(crate) fn is_empty(&self) -> bool {
        self.anchors.is_empty()
    }

    /// An anchor's spill.
    pub(crate) fn get(&self, anchor: CellId) -> Option<&Spill> {
        self.by_anchor.get(&anchor)
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
        self.boxes
            .get(s)
            .into_iter()
            .flat_map(move |b| b.meeting(cell))
            .map(|(_, tag)| self.anchors[tag as usize])
            .filter(move |&anchor| {
                anchor != (s, row, col)
                    && match held {
                        None => true,
                        Some(_) => self.by_anchor[&anchor]
                            .file
                            .is_some_and(|f| inside(&f, row, col)),
                    }
            })
    }

    /// The anchors whose boxes meet an area of a sheet.
    pub(crate) fn meeting(&self, sheet: usize, area: A1Range) -> impl Iterator<Item = CellId> + '_ {
        self.boxes
            .get(sheet)
            .into_iter()
            .flat_map(move |b| b.meeting(area))
            .map(|(_, tag)| self.anchors[tag as usize])
    }
}

/// Whether a rectangle holds a cell.
pub(crate) fn inside(r: &A1Range, row: u32, col: u32) -> bool {
    r.top <= row && row <= r.bottom && r.left <= col && col <= r.right
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
/// anchor's, a ring of them none.
struct Bounds<'a> {
    model: &'a Workbook,
    texts: &'a HashMap<CellId, String>,
    functions: &'a Functions,
    memo: HashMap<CellId, Option<Dims>>,
    visiting: HashSet<CellId>,
}

impl Bounds<'_> {
    /// A dynamic formula's bound by its parse, else the extent its file saved;
    /// one value for a cell holding none, whose `#` is `#REF!`.
    fn dims(&mut self, id: CellId, lib: &Library) -> Option<Dims> {
        if let Some(d) = self.memo.get(&id) {
            return *d;
        }
        let Some(file) = file_of(self.model, id) else {
            return Some((1, 1));
        };
        if !self.visiting.insert(id) {
            return None;
        }
        let parsed = self
            .texts
            .get(&id)
            .and_then(|t| lambda::parsed(t, self.functions, lib).ok());
        let by_parse = parsed.and_then(|e| {
            let model = self.model;
            bound(&e, lib, &mut |r: &RefExpr| {
                let sheet = match &r.sheet {
                    Some(name) => model.find_sheet(name)?,
                    None => id.0,
                };
                self.dims((sheet, r.top, r.left), lib)
            })
        });
        let d = by_parse.or_else(|| file.map(|f| (f.bottom - f.top + 1, f.right - f.left + 1)));
        self.visiting.remove(&id);
        self.memo.insert(id, d);
        d
    }
}

/// The most rows and columns `e`'s answer can take, or `None` where its parse
/// gives none; `anchor` gives a spill read through `#` its anchor's bound.
pub(crate) fn bound(
    e: &Expr,
    lib: &Library,
    anchor: &mut dyn FnMut(&RefExpr) -> Option<Dims>,
) -> Option<Dims> {
    match e {
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => {
            Some((1, 1))
        }
        Expr::Array(a) => Some((a.rows(), a.columns())),
        Expr::Ref(r) if r.last_sheet.is_none() => Some(dims(r)),
        Expr::Range(r) => r.reach.as_ref().map(dims),
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => bound(x, lib, anchor),
        Expr::Binary(_, a, b) => Some(lift(bound(a, lib, anchor)?, bound(b, lib, anchor)?)),
        Expr::Call(name, args) => call(name, args, lib, anchor),
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
    anchor: &mut dyn FnMut(&RefExpr) -> Option<Dims>,
) -> Option<Dims> {
    let mut of = |i: usize| -> Option<Dims> { bound(args.get(i)?, lib, anchor) };
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
            if of(0)? != (1, 1) {
                return None;
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
            [Expr::Ref(r)] => anchor(r),
            _ => None,
        },
        _ => {
            // A function of values: lifted over what its parameters of one
            // value are given, as a dynamic formula lifts it; a range or an
            // array handed whole makes one value.
            lib.strict(name)?;
            let classes = lib.classes(name)?;
            let mut shape = (1, 1);
            for (i, a) in args.iter().enumerate() {
                if matches!(classes(i), Param::Value | Param::Scalar | Param::Forced) {
                    shape = lift(shape, bound(a, lib, anchor)?);
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
        bound(&e, &lib, &mut |_| Some((3, 2)))
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
        ] {
            assert_eq!(of(text), want, "{text}");
        }
        // No bound: a count read from a cell, a name, a function the library
        // lacks.
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
        let spills = Spills::of(&wb, &texts, &Functions::default(), &lib);
        let at = |addr: &str| {
            let r = parse_a1_range(addr).unwrap();
            spills.get((s, r.top, r.left)).copied()
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
        assert_eq!(owners("P4"), [(s, 1, 16), (s, 3, 16)]);
        assert_eq!(owners("B1"), []);
        assert_eq!(owners("C1"), [], "an anchor its own saved extent holds");
        assert_eq!(owners("Z9"), []);
    }
}
