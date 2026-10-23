//! The database functions (KERNEL.8, slice e, 2026-10-20): `DSUM`,
//! `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`,
//! `DSTDEV`, `DSTDEVP`, `DVAR` and `DVARP` (the owner's decisions,
//! 2026-10-20: the language, the family whole).
//!
//! `DSUM(database, field, criteria)`: a database is a range whose first row
//! names its columns and whose every row after it is a record; the field is
//! a column, by a name its first row holds, without case, or by its place,
//! the number read down to a whole number, `#VALUE!` for none; the criteria
//! are a range whose first row names columns of the database and whose
//! every row after it is a test of a record, each cell under a name a test
//! of that column, a record meeting a row whose every test it meets and the
//! criteria when it meets any row, so that a row of no tests, a blank row,
//! meets every record. A test reads as `SUMIF`'s criterion reads
//! (`super::criteria`), but for a text with no operator before it, which
//! is a record's text beginning with it, as Excel's filters read one:
//! `Dav` meets `Davolio`, and `=Dav` `Dav` alone. Each function aggregates
//! the field's cells of the records the criteria meet as its namesake
//! aggregates a range (`super::aggregates`); `DCOUNT` and `DCOUNTA` with the
//! field left out count the records; `DGET` answers the one record's cell,
//! `#VALUE!` for none and for an empty cell, and `#NUM!` for more than one.
//! A criteria range of one row, its names alone, is `#VALUE!`. A database's
//! rows end with its sheet's last cell, so that a whole column's records are
//! the rows its sheet holds anything in; a criteria range's rows past it
//! read blank. Excel's save settled the three (`scripts/recalc/criteria.txt`,
//! `functions!A208`, `A211` and `A213`).
//!
//! A test under a name no column of the database holds is labelled: a
//! formula there is Excel's computed criterion, a formula moved to each
//! record, and a value there names no column, both left to Excel's save to
//! say (`a computed criterion`, `a criterion of no field`).

use super::aggregates::{Stat, Tally};
use super::criteria::{criterion, meets, one_area, Criterion, Target};
use super::eval::{Arg, Ctx, Reason};
use super::graph::Area;
use super::order;
use super::texts::{answer, wrong, Fault};
use super::value::{ErrorKind, Value};

/// `DSUM(database, field, criteria)`.
pub(crate) fn dsum(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Sum))
}

/// `DAVERAGE(database, field, criteria)`.
pub(crate) fn daverage(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Average))
}

/// `DCOUNT(database, [field], criteria)`.
pub(crate) fn dcount(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Count))
}

/// `DCOUNTA(database, [field], criteria)`.
pub(crate) fn dcounta(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::CountA))
}

/// `DGET(database, field, criteria)`.
pub(crate) fn dget(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, None)
}

/// `DMAX(database, field, criteria)`.
pub(crate) fn dmax(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Max))
}

/// `DMIN(database, field, criteria)`.
pub(crate) fn dmin(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Min))
}

/// `DPRODUCT(database, field, criteria)`.
pub(crate) fn dproduct(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::Product))
}

/// `DSTDEV(database, field, criteria)`.
pub(crate) fn dstdev(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::StDevS))
}

/// `DSTDEVP(database, field, criteria)`.
pub(crate) fn dstdevp(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::StDevP))
}

/// `DVAR(database, field, criteria)`.
pub(crate) fn dvar(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::VarS))
}

/// `DVARP(database, field, criteria)`.
pub(crate) fn dvarp(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    database(ctx, args, Some(Stat::VarP))
}

/// A label of a test the database function does not read (see the module's
/// comment).
fn label<T>(what: &str) -> Result<T, Fault> {
    Err(Fault::Reason(Reason::Construct(what.to_string())))
}

/// The last row of an area a sheet holds anything in, the area's own where
/// the sheet's used rectangle reaches past it.
fn last_row(ctx: &Ctx<'_>, area: &Area) -> u32 {
    match ctx.env.extent(area.sheet) {
        Some((_, (bottom, _))) => area.bottom.min(bottom),
        None => area.top,
    }
}

/// The column of a database a name names: the first whose first row's text
/// is the name without case, as Excel's order of texts holds two one.
fn column_of(ctx: &Ctx<'_>, db: &Area, name: &str) -> Result<Option<u32>, Reason> {
    for col in db.left..=db.right {
        if let Value::Text(head) = ctx.env.read((db.sheet, db.top, col))? {
            if order::equal(name, &head) {
                return Ok(Some(col));
            }
        }
    }
    Ok(None)
}

/// A test as a database's criteria read it: `SUMIF`'s criterion, a text
/// with no operator before it a beginning (see the module's comment).
fn test_of(v: &Value) -> Criterion {
    let c = criterion(v);
    match (v, c) {
        (Value::Text(t), Criterion::Eq(Target::Text(p)))
            if !t.starts_with('=') && !p.ends_with('*') =>
        {
            Criterion::Eq(Target::Text(p + "*"))
        }
        (_, c) => c,
    }
}

/// The criteria's rows of tests, each a column of the database and its
/// test, a blank cell no test; a test under a name no column holds
/// labelled.
fn tests_of(ctx: &Ctx<'_>, db: &Area, crit: &Area) -> Result<Vec<Vec<(u32, Criterion)>>, Fault> {
    let mut columns = Vec::new();
    for col in crit.left..=crit.right {
        columns.push(match ctx.env.read((crit.sheet, crit.top, col))? {
            Value::Text(name) => column_of(ctx, db, &name)?,
            _ => None,
        });
    }
    let mut rows = Vec::new();
    let last = last_row(ctx, crit);
    for row in crit.top + 1..=last {
        let mut tests = Vec::new();
        for (k, col) in (crit.left..=crit.right).enumerate() {
            let v = ctx.env.read((crit.sheet, row, col))?;
            if matches!(&v, Value::Empty) || matches!(&v, Value::Text(t) if t.is_empty()) {
                continue;
            }
            match columns[k] {
                Some(at) => tests.push((at, test_of(&v))),
                None => {
                    // A legacy array formula's cell read as no field's,
                    // as before (KERNEL.8).
                    let formula = ctx
                        .env
                        .written((crit.sheet, row, col))
                        .is_some_and(|w| !w.array());
                    return label(if formula {
                        "a computed criterion"
                    } else {
                        "a criterion of no field"
                    });
                }
            }
        }
        rows.push(tests);
    }
    // A row of the criteria past every cell the sheet holds is blank, and
    // a blank row meets every record.
    if last < crit.bottom {
        rows.push(Vec::new());
    }
    Ok(rows)
}

/// Whether a record meets the criteria: any of its rows, each of whose
/// tests the record's cell meets.
fn met(ctx: &Ctx<'_>, db: &Area, row: u32, rows: &[Vec<(u32, Criterion)>]) -> Result<bool, Fault> {
    'rows: for tests in rows {
        for (col, c) in tests {
            let v = ctx.env.read((db.sheet, row, *col))?;
            if !meets(&v, c)? {
                continue 'rows;
            }
        }
        return Ok(true);
    }
    Ok(false)
}

/// A database function by its aggregate, `DGET` by none (see the module's
/// comment).
fn database(ctx: &Ctx<'_>, args: &[Arg], stat: Option<Stat>) -> Result<Value, Reason> {
    answer(|| {
        let [db, field, crit] = args else {
            return wrong();
        };
        let (Some(db), Some(crit)) = (one_area(Some(db)), one_area(Some(crit))) else {
            return wrong();
        };
        if db.sheet >= ctx.env.sheets() || crit.sheet >= ctx.env.sheets() {
            return Err(Fault::Error(ErrorKind::Ref));
        }
        let field = match ctx.one(field)? {
            Value::Number(n) => {
                let k = n.floor();
                if k < 1.0 || k > f64::from(db.right - db.left + 1) {
                    return wrong();
                }
                Some(db.left + k as u32 - 1)
            }
            Value::Text(name) => match column_of(ctx, &db, &name)? {
                Some(col) => Some(col),
                None => return wrong(),
            },
            Value::Empty if matches!(stat, Some(Stat::Count | Stat::CountA)) => None,
            Value::Error(k) => return Err(Fault::Error(k)),
            _ => return wrong(),
        };
        // A criteria range of one row, its names alone, tests nothing, which
        // Excel's save answers `#VALUE!` (`functions!A208`).
        if crit.bottom == crit.top {
            return wrong();
        }
        let rows = tests_of(ctx, &db, &crit)?;
        let mut tally = stat.map(Tally::new);
        let mut found: Option<Value> = None;
        let mut records = 0u64;
        let mut more = false;
        let mut take = |v: Value| match &mut tally {
            Some(t) => t.ranged(&v),
            None if found.is_some() => more = true,
            None => found = Some(v),
        };
        let last = last_row(ctx, &db);
        for row in db.top + 1..=last {
            if !met(ctx, &db, row, &rows)? {
                continue;
            }
            records += 1;
            if let Some(col) = field {
                take(ctx.env.read((db.sheet, row, col))?);
            }
        }
        Ok(match (tally, field) {
            (Some(_), None) => Value::Number(records as f64),
            (Some(t), Some(_)) => t.answer(),
            (None, _) if more => Value::Error(ErrorKind::Num),
            (None, _) => match found {
                None | Some(Value::Empty) => Value::Error(ErrorKind::Value),
                Some(v) => v,
            },
        })
    })
}

#[cfg(test)]
mod tests {
    use crate::calc::aggregates::fixture::{holds, DIV0, NUM, VALUE};

    #[test]
    fn a_database_s_records_are_aggregated_where_its_criteria_meet_them() {
        holds(
            &[],
            &[
                ("DSUM(T!A1:H13,\"Amount\",T!K1:K2)", "310.2"),
                ("DSUM(T!A1:H13,4,T!K1:K2)", "310.2"),
                ("DSUM(T!A1:H13,4.7,T!K1:K2)", "310.2"),
                ("DSUM(T!A1:H13,\"amount\",T!K1:K2)", "310.2"),
                ("DSUM(T!A1:H13,\"Amt\",T!K1:K2)", VALUE),
                ("DSUM(T!A1:H13,0,T!K1:K2)", VALUE),
                ("DSUM(T!A1:H13,9,T!K1:K2)", VALUE),
                ("DSUM(T!A1:H13,,T!K1:K2)", VALUE),
                ("DCOUNTA(T!A1:H13,\"Note\",T!L1:L2)", "4"),
                ("DCOUNTA(T!A1:H13,\"Note\",T!M1:M2)", "1"),
                ("DSUM(T!A1:H13,\"Amount\",T!N1:N2)", "480.2"),
                ("DSUM(T!A1:H13,\"Amount\",T!O1:P2)", "125.2"),
                ("DSUM(T!A1:H13,\"Amount\",T!Q1:Q3)", "625.6"),
                ("DSUM(T!A1:H13,\"Amount\",T!R1:R3)", "995.8000000000001"),
                ("DCOUNT(T!A1:H13,\"Qty\",T!S1:S2)", "4"),
                (
                    "DSUM(T!A1:H13,\"Amount\",T!T1:T2)",
                    "(not-computed \"a criterion of no field\")",
                ),
                (
                    "DSUM(T!A1:H13,\"Amount\",T!U1:U2)",
                    "(not-computed \"a computed criterion\")",
                ),
                ("DGET(T!A1:H13,\"Amount\",T!O1:P2)", NUM),
                ("DGET(T!A1:H13,\"Amount\",T!M1:M2)", "-20"),
                ("DGET(T!A1:H13,\"Amount\",T!W1:W2)", VALUE),
                ("DGET(T!A1:H13,\"Note\",T!AL1:AL2)", VALUE),
                ("DCOUNT(T!A1:H13,,T!K1:K2)", "5"),
                ("DCOUNTA(T!A1:H13,,T!K1:K2)", "5"),
                ("DCOUNT(T!A1:H13,\"Mixed\",T!K1:K2)", "4"),
                ("DCOUNTA(T!A1:H13,\"Mixed\",T!K1:K2)", "5"),
                ("DAVERAGE(T!A1:H13,\"Amount\",T!K1:K2)", "62.04"),
                ("DAVERAGE(T!A1:H13,\"Amount\",T!W1:W2)", DIV0),
                ("DMAX(T!A1:H13,\"Amount\",T!K1:K2)", "125"),
                ("DMIN(T!A1:H13,\"Amount\",T!K1:K2)", "0.2"),
                ("DMAX(T!A1:H13,\"Amount\",T!W1:W2)", "0"),
                ("DPRODUCT(T!A1:H13,\"Qty\",T!K1:K2)", "0"),
                ("DSTDEV(T!A1:H13,\"Qty\",T!K1:K2)", "3.5071355833500366"),
                ("DSTDEVP(T!A1:H13,\"Qty\",T!K1:K2)", "3.1368774282716245"),
                ("DVAR(T!A1:H13,\"Qty\",T!K1:K2)", "12.3"),
                ("DVARP(T!A1:H13,\"Qty\",T!K1:K2)", "9.84"),
                ("DSUM(T!A1:H13,\"Mixed\",T!K1:K2)", "12"),
                ("DSUM(T!A1:H13,\"Mixed\",T!Q1:Q3)", DIV0),
                ("DSUM(T!A1:H13,\"Amount\",T!K1)", VALUE),
                ("DSUM(T!A1:H13,\"Amount\",T!K2:K3)", "995.8000000000001"),
                ("DSUM(T!A1:H13,\"Amount\",T!K1:K20)", "995.8000000000001"),
                ("DSUM(T!A1:H13,\"Amount\",T!K13:K15)", "995.8000000000001"),
                ("DSUM(T!A:H,\"Amount\",T!K1:K2)", "310.2"),
                ("DCOUNTA(T!A:H,,T!AF1:AF2)", "1"),
                ("DGET(T!A:H,\"Amount\",T!AF1:AF2)", "60"),
            ],
        );
    }
}
