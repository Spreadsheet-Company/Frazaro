//! The aggregates (KERNEL.8, slice e, 2026-10-20): `COUNT`, `COUNTA`,
//! `PRODUCT`, `AVERAGE`, `STDEV`, `STDEVP`, `VAR` and `VARP`, with their
//! newer names `STDEV.S`, `STDEV.P`, `VAR.S` and `VAR.P`, and `SUBTOTAL`,
//! whose eleven they and `SUM`, `MAX` and `MIN` are. One tally serves them
//! all, and the database functions (`super::database`) and the criteria's
//! `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and `MINIFS` (`super::criteria`)
//! with them (the owner's decisions, 2026-10-20: the language, the two
//! families whole). `AVERAGE` moved down from `frazaro-core`'s Excel
//! library with its meaning, `SUBTOTAL`'s first.
//!
//! A value given directly and a value read out of a range count as `SUM`
//! counts them (`library::OverNumbers`): given, a number, a truth value or
//! a text spelling a number is the number, any other text `#VALUE!` and an
//! error the answer; in a range only a number counts and the range's first
//! error is the answer once it is walked. `COUNT` counts what would be a
//! number, given or in a range, and passes every error by; `COUNTA` counts
//! every value given and every cell of a range holding one, an error and an
//! empty text among them. A sum and a mean add in the walk's order from
//! -0.0, as `SUM` does; a variance is taken in two passes, the mean first
//! and then the squares of the distances from it, each in the walk's order.
//! Every rule where Excel's documentation is silent is held to Excel's own
//! save (`scripts/recalc/criteria.txt`).
//!
//! `SUBTOTAL(number, reference, ...)` aggregates its references by the
//! number, 1 to 11 or, passing every hidden row by, 101 to 111, and passes
//! by every cell whose formula is itself a `SUBTOTAL`. A row hidden by hand
//! counts to 1 to 11 and a row a filter hides does not, and an xlsx file
//! tells the two apart only by its filter's range, which the reader does
//! not read: where a hidden row holds a cell of a reference, 1 to 11 are
//! labelled, `a row a filter may hide` (the owner's decision, 2026-10-20);
//! over a hidden row holding none they compute, nothing there counting
//! either way. A hidden row is the model's (`Sheet::rows`): a cartridge's
//! `row` rows, or the rows a reader reads hidden, an xlsx file's `hidden`
//! and an OpenDocument file's `collapse` and `filter`, told as one.

use super::eval::{finite, Arg, Ctx, Reason};
use super::graph::{Area, CellId};
use super::index::walk;
use super::texts::{answer, wrong, Fault};
use super::value::{text_to_number, ErrorKind, Value};

/// What an aggregate computes, `SUBTOTAL`'s eleven.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Stat {
    Average,
    Count,
    CountA,
    Max,
    Min,
    Product,
    StDevS,
    StDevP,
    Sum,
    VarS,
    VarP,
}

impl Stat {
    /// `SUBTOTAL`'s function numbers 1 to 11.
    pub(crate) fn numbered(n: i64) -> Option<Stat> {
        Some(match n {
            1 => Stat::Average,
            2 => Stat::Count,
            3 => Stat::CountA,
            4 => Stat::Max,
            5 => Stat::Min,
            6 => Stat::Product,
            7 => Stat::StDevS,
            8 => Stat::StDevP,
            9 => Stat::Sum,
            10 => Stat::VarS,
            11 => Stat::VarP,
            _ => return None,
        })
    }

    /// Whether the aggregate keeps its numbers for a second pass.
    fn keeps(self) -> bool {
        matches!(self, Stat::StDevS | Stat::StDevP | Stat::VarS | Stat::VarP)
    }
}

/// What an aggregate has read, in the walk's order.
pub(crate) struct Tally {
    stat: Stat,
    sum: f64,
    product: f64,
    numbers: u64,
    values: u64,
    most: f64,
    least: f64,
    kept: Vec<f64>,
    error: Option<ErrorKind>,
}

impl Tally {
    pub(crate) fn new(stat: Stat) -> Tally {
        Tally {
            stat,
            sum: -0.0,
            product: 1.0,
            numbers: 0,
            values: 0,
            most: f64::NEG_INFINITY,
            least: f64::INFINITY,
            kept: Vec::new(),
            error: None,
        }
    }

    /// A tally of a sum, a mean, a maximum or a minimum from what an index
    /// kept of the numbers it walked (KERNEL.8, slice e): their sum added in
    /// the walk's order from -0.0, their count, the greatest and the least.
    pub(crate) fn grouped(stat: Stat, sum: f64, numbers: u64, most: f64, least: f64) -> Tally {
        debug_assert!(matches!(
            stat,
            Stat::Sum | Stat::Average | Stat::Max | Stat::Min
        ));
        Tally {
            stat,
            sum,
            product: 1.0,
            numbers,
            values: numbers,
            most,
            least,
            kept: Vec::new(),
            error: None,
        }
    }

    fn number(&mut self, n: f64) {
        self.numbers += 1;
        self.values += 1;
        self.sum += n;
        self.product *= n;
        self.most = self.most.max(n);
        self.least = self.least.min(n);
        if self.stat.keeps() {
            self.kept.push(n);
        }
    }

    /// A value read out of a range: a number counts; an error is kept, the
    /// answer once the walk ends; any other value counts to `COUNTA` alone,
    /// an empty cell to none.
    pub(crate) fn ranged(&mut self, v: &Value) {
        match v {
            Value::Number(n) => self.number(*n),
            Value::Empty => {}
            Value::Error(k) => {
                self.values += 1;
                self.error.get_or_insert(*k);
            }
            Value::Text(_) | Value::Bool(_) => self.values += 1,
        }
    }

    /// A value given directly: the number it reads as counts, or its error
    /// is the answer at once; `COUNT` counts what would be a number and
    /// passes the rest by, `COUNTA` counts every value.
    pub(crate) fn given(&mut self, v: &Value) -> Result<(), ErrorKind> {
        match self.stat {
            Stat::CountA => self.values += 1,
            Stat::Count => match v {
                Value::Number(_) | Value::Bool(_) | Value::Empty => self.numbers += 1,
                Value::Text(t) if text_to_number(t).is_some() => self.numbers += 1,
                _ => {}
            },
            _ => self.number(v.number()?),
        }
        Ok(())
    }

    /// The aggregate's answer: a range's first error where the aggregate
    /// keeps one, `COUNT` and `COUNTA` keeping none; a mean or a variance
    /// of too few numbers `#DIV/0!`; a maximum, a minimum or a product of
    /// none 0.
    pub(crate) fn answer(self) -> Value {
        let n = self.numbers as f64;
        match self.stat {
            Stat::Count => return Value::Number(n),
            Stat::CountA => return Value::Number(self.values as f64),
            _ => {}
        }
        if let Some(k) = self.error {
            return Value::Error(k);
        }
        let div0 = Value::Error(ErrorKind::Div0);
        match self.stat {
            Stat::Sum => finite(self.sum),
            Stat::Average if self.numbers == 0 => div0,
            Stat::Average => finite(self.sum / n),
            Stat::Max if self.numbers == 0 => Value::Number(0.0),
            Stat::Max => Value::Number(self.most),
            Stat::Min if self.numbers == 0 => Value::Number(0.0),
            Stat::Min => Value::Number(self.least),
            Stat::Product if self.numbers == 0 => Value::Number(0.0),
            Stat::Product => finite(self.product),
            Stat::VarS | Stat::StDevS if self.numbers < 2 => div0,
            Stat::VarP | Stat::StDevP if self.numbers == 0 => div0,
            Stat::VarS | Stat::StDevS | Stat::VarP | Stat::StDevP => {
                let mean = self.sum / n;
                let squares = self
                    .kept
                    .iter()
                    .fold(0.0, |s, x| s + (x - mean) * (x - mean));
                let over = if matches!(self.stat, Stat::VarS | Stat::StDevS) {
                    n - 1.0
                } else {
                    n
                };
                let variance = squares / over;
                finite(if matches!(self.stat, Stat::StDevS | Stat::StDevP) {
                    variance.sqrt()
                } else {
                    variance
                })
            }
            Stat::Count | Stat::CountA => unreachable!("answered above"),
        }
    }
}

/// An aggregate over a cell's arguments, each given value and each range's
/// cells counted in order (see the module's comment).
fn aggregate(ctx: &Ctx<'_>, args: &[Arg], stat: Stat) -> Result<Value, Reason> {
    let mut tally = Tally::new(stat);
    for a in args {
        match a {
            Arg::Scalar(v) => {
                if let Err(k) = tally.given(v) {
                    return Ok(Value::Error(k));
                }
            }
            Arg::Area { areas, .. } => ctx.each(areas, |v| {
                tally.ranged(&v);
                Ok(())
            })?,
            // An array's values are counted as a range's cells are, an empty
            // one skipped but by `COUNTA`, which counts it where it skips an
            // empty cell, as Excel's own save of the arrays' fixture counts
            // `COUNTA(TRANSPOSE(B1:B6))` over an empty `B6` (KERNEL.8,
            // slice f).
            Arg::Array(a) => {
                for v in a.values() {
                    match v {
                        Value::Empty if stat == Stat::CountA => tally.values += 1,
                        Value::Empty => {}
                        v => tally.ranged(v),
                    }
                }
            }
        }
    }
    Ok(tally.answer())
}

/// `AVERAGE(value, ...)`.
pub(crate) fn average(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::Average)
}

/// `COUNT(value, ...)`.
pub(crate) fn count(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::Count)
}

/// `COUNTA(value, ...)`.
pub(crate) fn counta(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::CountA)
}

/// `PRODUCT(value, ...)`.
pub(crate) fn product(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::Product)
}

/// `STDEV(value, ...)` and `STDEV.S`.
pub(crate) fn stdev_s(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::StDevS)
}

/// `STDEVP(value, ...)` and `STDEV.P`.
pub(crate) fn stdev_p(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::StDevP)
}

/// `VAR(value, ...)` and `VAR.S`.
pub(crate) fn var_s(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::VarS)
}

/// `VARP(value, ...)` and `VAR.P`.
pub(crate) fn var_p(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    aggregate(ctx, args, Stat::VarP)
}

/// The label of `SUBTOTAL`'s 1 to 11 over a hidden row (see the module's
/// comment).
fn filtered() -> Reason {
    Reason::Construct("a row a filter may hide".to_string())
}

/// Whether a cell's formula calls `SUBTOTAL`, which a `SUBTOTAL` over it
/// passes by: the name read as a function's, a call's open bracket after
/// it, outside every text in quotes. A shared formula's child, which an
/// engine's formula written over a range makes, reads as its master's
/// text, whose functions are its own.
fn calls_subtotal(ctx: &Ctx<'_>, id: CellId) -> bool {
    // A legacy array formula's cell is no `SUBTOTAL`'s to pass by.
    match ctx.env.written(id) {
        Some(w) if !w.array() => w.text.is_some_and(|text| calls(text, "SUBTOTAL")),
        _ => false,
    }
}

/// Whether a formula's text calls the function `name`, by its name without
/// case, outside every text in quotes.
pub(crate) fn calls(text: &str, name: &str) -> bool {
    let bytes = text.as_bytes();
    let mut quoted = false;
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b == b'"' {
            quoted = !quoted;
            i += 1;
            continue;
        }
        let word = |c: u8| c.is_ascii_alphanumeric() || matches!(c, b'_' | b'.');
        if !quoted && word(b) && (i == 0 || !word(bytes[i - 1])) {
            let start = i;
            while i < bytes.len() && word(bytes[i]) {
                i += 1;
            }
            let mut after = i;
            while after < bytes.len() && bytes[after] == b' ' {
                after += 1;
            }
            if text[start..i].eq_ignore_ascii_case(name) && bytes.get(after) == Some(&b'(') {
                return true;
            }
            continue;
        }
        i += 1;
    }
    false
}

/// `SUBTOTAL(number, reference, ...)` (see the module's comment): the number
/// read down to a whole number, 1 to 11 or 101 to 111, `#VALUE!` otherwise;
/// each argument after it a reference, `#VALUE!` otherwise.
pub(crate) fn subtotal(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [number, references @ ..] = args else {
            return wrong();
        };
        if references.is_empty() {
            return wrong();
        }
        let n = ctx.one(number)?.number()?.floor();
        let (stat, visible) = match n as i64 {
            k @ 1..=11 => (Stat::numbered(k), false),
            k @ 101..=111 => (Stat::numbered(k - 100), true),
            _ => return wrong(),
        };
        let Some(stat) = stat else { return wrong() };
        let mut tally = Tally::new(stat);
        for r in references {
            let Arg::Area { areas, .. } = r else {
                return wrong();
            };
            for area in areas.iter() {
                if area.sheet >= ctx.env.sheets() {
                    return Err(Fault::Error(ErrorKind::Ref));
                }
                let rows = ctx.env.rows(area.sheet);
                let hidden = |row: u32| rows.is_some_and(|r| r.get(&row).is_some_and(|s| s.hidden));
                // A hidden row holding no cell of the area is passed by or
                // counted alike, and computes.
                let holds = |row: u32| {
                    let line = Area {
                        top: row,
                        bottom: row,
                        ..*area
                    };
                    ctx.env.walk(&line).next().is_some()
                };
                if !visible
                    && rows.is_some_and(|r| {
                        r.range(area.top..=area.bottom)
                            .any(|(&row, s)| s.hidden && holds(row))
                    })
                {
                    return Err(Fault::Reason(filtered()));
                }
                let mut stop: Option<Reason> = None;
                walk(ctx, area, |row, col, read| {
                    if (visible && hidden(row)) || calls_subtotal(ctx, (area.sheet, row, col)) {
                        return true;
                    }
                    match read {
                        Ok(v) => {
                            tally.ranged(v);
                            true
                        }
                        Err(r) => {
                            stop = Some(r.clone());
                            false
                        }
                    }
                });
                if let Some(r) = stop {
                    return Err(Fault::Reason(r));
                }
            }
        }
        Ok(tally.answer())
    })
}

/// The table the tests of the aggregates, the criteria's family and the
/// database functions read (KERNEL.8, slice e), on a sheet `T`: twelve
/// records under names in row 1, an account (one blank, one in lower
/// case), a month (one a text spelling 3), a region (cases mixed), an
/// amount, a quantity, a note (wildcards, an empty text, texts beginning
/// alike, texts spelling numbers), a mixed column (a truth value, a text
/// spelling 5, two errors, a blank) and a day; criteria for the database
/// functions in K to W, AF and AL; and a `SUBTOTAL` among the cells it sums
/// in X. A text is written after a `'`, a formula after a `=`.
#[cfg(test)]
pub(crate) mod fixture {
    use crate::calc::value::ErrorKind;
    use crate::calc::{library, Calc};
    use crate::sheet::{parse_a1_range, Cell, Content, RowSettings, Workbook};

    pub(crate) const TABLE: &[(&str, &str)] = &[
        ("A1", "'Account"),
        ("B1", "'Month"),
        ("C1", "'Region"),
        ("D1", "'Amount"),
        ("E1", "'Qty"),
        ("F1", "'Note"),
        ("G1", "'Mixed"),
        ("H1", "'Day"),
        ("A2", "'Cash"),
        ("B2", "1"),
        ("C2", "'East"),
        ("D2", "100"),
        ("E2", "5"),
        ("F2", "'plain"),
        ("G2", "1"),
        ("H2", "46023"),
        ("A3", "'Bank"),
        ("B3", "1"),
        ("C3", "'West"),
        ("D3", "250.5"),
        ("E3", "3"),
        ("F3", "'a*b"),
        ("G3", "TRUE"),
        ("H3", "46024"),
        ("A4", "'Cash"),
        ("B4", "2"),
        ("C4", "'east"),
        ("D4", "75"),
        ("E4", "2"),
        ("F4", "'a?b"),
        ("G4", "'5"),
        ("H4", "46054"),
        ("A5", "'Card"),
        ("B5", "2"),
        ("C5", "'North"),
        ("D5", "40"),
        ("E5", "1"),
        ("F5", "'~tilde"),
        ("G5", "#DIV/0!"),
        ("H5", "46055"),
        ("A6", "'Bank"),
        ("B6", "3"),
        ("C6", "'West"),
        ("D6", "300"),
        ("E6", "4"),
        ("H6", "46082"),
        ("A7", "'Cash"),
        ("B7", "3"),
        ("C7", "'East"),
        ("D7", "125"),
        ("E7", "6"),
        ("F7", "=\"\""),
        ("G7", "2"),
        ("H7", "46083"),
        ("A8", "'Card"),
        ("B8", "1"),
        ("C8", "'South"),
        ("D8", "-20"),
        ("E8", "7"),
        ("F8", "'Dav"),
        ("G8", "#VALUE!"),
        ("H8", "46025"),
        ("A9", "'Cash"),
        ("B9", "12"),
        ("C9", "'East"),
        ("D9", "10"),
        ("E9", "0"),
        ("F9", "'Davolio"),
        ("G9", "3"),
        ("H9", "46357"),
        ("B10", "2"),
        ("C10", "'West"),
        ("D10", "60"),
        ("E10", "2"),
        ("F10", "'David"),
        ("G10", "4"),
        ("H10", "46056"),
        ("A11", "'bank"),
        ("B11", "2"),
        ("C11", "'WEST"),
        ("D11", "0.1"),
        ("E11", "1"),
        ("F11", "'davis"),
        ("G11", "5"),
        ("H11", "46057"),
        ("A12", "'Cash"),
        ("B12", "3"),
        ("C12", "'East"),
        ("D12", "0.2"),
        ("E12", "9"),
        ("F12", "'100"),
        ("G12", "6"),
        ("H12", "46084"),
        ("A13", "'Card"),
        ("B13", "'3"),
        ("C13", "'North"),
        ("D13", "55"),
        ("E13", "3"),
        ("F13", "'1e2"),
        ("G13", "7"),
        ("H13", "46085"),
        ("K1", "'Region"),
        ("K2", "'East"),
        ("L1", "'Note"),
        ("L2", "'Dav"),
        ("M1", "'Note"),
        ("M2", "'=Dav"),
        ("N1", "'Month"),
        ("N2", "3"),
        ("O1", "'Account"),
        ("O2", "'Cash"),
        ("P1", "'Month"),
        ("P2", "3"),
        ("Q1", "'Account"),
        ("Q2", "'Card"),
        ("Q3", "'Bank"),
        ("R1", "'Account"),
        ("R2", "'Cash"),
        ("S1", "'region"),
        ("S2", "'west"),
        ("T1", "'Nope"),
        ("T2", "'x"),
        ("U1", "'Big"),
        ("U2", "=T!D2>100"),
        ("W1", "'Account"),
        ("W2", "'Nope"),
        ("AF1", "'Account"),
        ("AF2", "'="),
        ("AL1", "'Amount"),
        ("AL2", "300"),
        ("X1", "1"),
        ("X2", "2"),
        ("X3", "=SUBTOTAL(9,X1:X2)"),
        ("X4", "=SUBTOTAL(9,X1:X2)*2"),
        ("X5", "=SUM(X1:X2)"),
    ];

    /// Each formula's value over [`TABLE`], the rows given hidden on `T`,
    /// as `Calc` computes it, spelled.
    pub(crate) fn computed(hidden: &[u32], formulas: &[&str]) -> Vec<String> {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        for (addr, text) in TABLE {
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(s) = text.strip_prefix('\'') {
                Content::Text(s.to_string())
            } else if let Some(f) = text.strip_prefix('=') {
                Content::Formula(f.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else if let Some(e) = ErrorKind::parse(text) {
                Content::Error(e.text().to_string())
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        }
        for row in hidden {
            let settings = RowSettings {
                height: None,
                hidden: true,
            };
            wb.sheets[t].rows.insert(*row, settings);
        }
        let f = wb.ensure_sheet("F");
        for (i, formula) in formulas.iter().enumerate() {
            let content = Content::Formula(formula.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        (0..formulas.len())
            .map(|i| match calc.computed((f, i as u32 + 1, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            })
            .collect()
    }

    /// Each formula of the pairs against the value it is to compute, the
    /// rows given hidden.
    pub(crate) fn holds(hidden: &[u32], cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|(f, _)| *f).collect();
        for ((formula, want), got) in cases.iter().zip(computed(hidden, &formulas)) {
            assert_eq!(got, *want, "={formula}");
        }
    }

    pub(crate) const VALUE: &str = "(error \"#VALUE!\")";
    pub(crate) const DIV0: &str = "(error \"#DIV/0!\")";
    pub(crate) const NUM: &str = "(error \"#NUM!\")";
}

#[cfg(test)]
mod tests {
    use super::fixture::{holds, DIV0, VALUE};
    use super::*;

    #[test]
    fn the_aggregates_count_what_they_are_given_and_what_a_range_holds() {
        holds(
            &[],
            &[
                ("COUNT(T!D2:D13)", "12"),
                ("COUNT(T!G2:G13)", "7"),
                ("COUNT(1,TRUE,\"2\",\"x\",1/0)", "3"),
                ("COUNT(1,,2)", "3"),
                ("COUNTA(T!F2:F13)", "11"),
                ("COUNTA(1,\"\",TRUE,1/0)", "4"),
                ("COUNTA(T!J1)", "0"),
                ("COUNTA(T!G2:G13)", "11"),
                ("PRODUCT(T!E2:E5)", "30"),
                ("PRODUCT(T!G2:G13)", DIV0),
                ("PRODUCT(2,\"3\",TRUE)", "6"),
                ("PRODUCT(\"x\")", VALUE),
                ("PRODUCT(T!J1:J2)", "0"),
                ("AVERAGE(T!E2:E13)", "3.5833333333333335"),
                ("AVERAGE(1,\"2\",TRUE)", "1.3333333333333333"),
                ("AVERAGE(T!J1:J2)", DIV0),
                ("AVERAGE(1,,2)", "1"),
                ("AVERAGE(T!F2:F13)", DIV0),
                ("AVERAGE(T!B2:B13)", "2.909090909090909"),
                ("VAR(1,2,3,4)", "1.6666666666666667"),
                ("VAR.S(1,2,3,4)", "1.6666666666666667"),
                ("VARP(1,2,3,4)", "1.25"),
                ("VAR.P(1,2,3,4)", "1.25"),
                ("STDEV(1,2,3,4)", "1.2909944487358056"),
                ("STDEV.S(1,2,3,4)", "1.2909944487358056"),
                ("STDEVP(1,2,3,4)", "1.118033988749895"),
                ("STDEV.P(1,2,3,4)", "1.118033988749895"),
                ("STDEV(5)", DIV0),
                ("STDEVP(5)", "0"),
                ("VARP(T!J1)", DIV0),
                ("VAR(1000000001,1000000002,1000000003)", "1"),
                ("STDEV(T!E2:E13)", "2.7122058561364044"),
                ("STDEVP(T!E2:E13)", "2.5967394084804805"),
                ("VAR(T!D2:D13)", "10008.22696969697"),
                ("VARP(T!D2:D13)", "9174.208055555557"),
                ("STDEV(T!G2:G13)", DIV0),
            ],
        );
    }

    #[test]
    fn subtotal_aggregates_by_its_number_and_passes_hidden_rows_and_subtotals_by() {
        holds(
            &[],
            &[
                ("SUBTOTAL(1,T!D2:D13)", "82.98333333333333"),
                ("SUBTOTAL(2,T!G2:G13)", "7"),
                ("SUBTOTAL(3,T!F2:F13)", "11"),
                ("SUBTOTAL(4,T!D2:D13)", "300"),
                ("SUBTOTAL(5,T!D2:D13)", "-20"),
                ("SUBTOTAL(6,T!E2:E5)", "30"),
                ("SUBTOTAL(7,T!D2:D13)", "100.04112639158443"),
                ("SUBTOTAL(8,T!D2:D13)", "95.78208629778095"),
                ("SUBTOTAL(9,T!D2:D13)", "995.8000000000001"),
                ("SUBTOTAL(10,T!D2:D13)", "10008.22696969697"),
                ("SUBTOTAL(11,T!D2:D13)", "9174.208055555557"),
                ("SUBTOTAL(109,T!D2:D13)", "995.8000000000001"),
                ("SUBTOTAL(101,T!D2:D13)", "82.98333333333333"),
                ("SUBTOTAL(111,T!D2:D13)", "9174.208055555557"),
                ("SUBTOTAL(9,T!G2:G13)", DIV0),
                ("SUBTOTAL(9.7,T!D2:D3)", "350.5"),
                ("SUBTOTAL(9,T!X1:X5)", "6"),
                ("SUBTOTAL(0,T!D2:D3)", VALUE),
                ("SUBTOTAL(12,T!D2:D3)", VALUE),
                ("SUBTOTAL(100,T!D2:D3)", VALUE),
                ("SUBTOTAL(112,T!D2:D3)", VALUE),
                ("SUBTOTAL(9,5)", VALUE),
                ("SUBTOTAL(9)", VALUE),
            ],
        );
        // Row 6, 300 of the amounts, hidden: 101 to 111 pass it by, and 1 to
        // 11 over a cell of it are labelled, a filter's row and a hand's one
        // to a file; over its columns that hold no cell, `F6` and `G6`, they
        // compute, and so they do over row 20, hidden and empty.
        let filtered = "(not-computed \"a row a filter may hide\")";
        holds(
            &[6, 20],
            &[
                ("SUBTOTAL(109,T!D2:D13)", "695.8000000000001"),
                ("SUBTOTAL(109,T!D2:D13,T!E2:E13)", "734.8000000000001"),
                ("SUBTOTAL(102,T!G2:G13)", "7"),
                ("SUBTOTAL(103,T!A2:A13)", "10"),
                ("SUBTOTAL(104,T!D2:D13)", "250.5"),
                ("SUBTOTAL(9,T!D2:D13)", filtered),
                ("SUBTOTAL(9,T!D2:D5)", "465.5"),
                ("SUBTOTAL(9,T!G2:G13)", DIV0),
                ("SUBTOTAL(3,T!F2:F13)", "11"),
                ("SUBTOTAL(9,T!F2:H13)", filtered),
                ("SUBTOTAL(9,T!D14:D30)", "0"),
                ("SUBTOTAL(9,T!D6)", filtered),
            ],
        );
    }

    #[test]
    fn a_call_is_read_by_its_name_outside_quotes() {
        assert!(calls("SUBTOTAL(9,A1:A3)", "SUBTOTAL"));
        assert!(calls("=1+subtotal (9,A1:A3)*2", "SUBTOTAL"));
        assert!(!calls("\"SUBTOTAL(9,A1)\"&A1", "SUBTOTAL"));
        assert!(!calls("MYSUBTOTAL(1)", "SUBTOTAL"));
        assert!(!calls("SUBTOTALS(1)", "SUBTOTAL"));
        assert!(!calls("SUM(A1:A3)", "SUBTOTAL"));
    }
}
