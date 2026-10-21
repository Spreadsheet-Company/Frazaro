//! The arrays a formula computes (KERNEL.8, slice f, 2026-10-20): an array
//! constant (`{1,2;3,4}`), a range read whole where an array is wanted, a
//! function's answer (`SEQUENCE`, `FILTER`), and an operator or a function
//! lifted over either, each value at its place.
//!
//! An array is no cell's value: it lives inside an evaluation, as an
//! operand ([`super::eval::Arg::Array`]), and a cell holds one value of it
//! ([`super::eval::Ctx::result`]). Its values are held once and shared by a
//! count, so that an array read twice, as `IF`'s branches and a lookup read
//! one, is never copied (the owner's decision 4, 2026-10-20); and one holds
//! at most [`LIMIT`] values, the grid's own limit, an array past it labelled
//! (decision 3), since Excel computes one there and the language does not
//! answer what Excel never would.
//!
//! **Lifting** is Excel's rule for an operator or a function of one value
//! given arrays: the answer has the most rows and the most columns of its
//! operands, and its value at each place is the operator's over the
//! operands' values there, an operand of one row standing at every row and
//! one of one column at every column, and `#N/A` at a place one of them
//! does not reach (`{1,2,3}+{10,20}` is `{11,22,#N/A}`).

use std::rc::Rc;

use super::eval::Reason;
use super::value::{ErrorKind, Value};

/// The most values one array holds: the grid's own limit, `CELLS`, sixteen
/// Screens, so that `SEQUENCE(1E9)` asks for nothing.
pub const LIMIT: usize = crate::machine::CELLS;

/// Rows by columns of values, in row-major order, held once and shared.
#[derive(Clone, Debug, PartialEq)]
pub struct Array {
    rows: u32,
    columns: u32,
    values: Rc<[Value]>,
}

impl Array {
    /// An array of these rows and columns from its values in row-major
    /// order. The caller sizes it by [`room`] first, so that no array past
    /// the limit is made; the count of values must be rows times columns.
    pub fn new(rows: u32, columns: u32, values: Vec<Value>) -> Array {
        assert_eq!(
            values.len() as u64,
            u64::from(rows) * u64::from(columns),
            "an array's values are its rows times its columns"
        );
        assert!(rows >= 1 && columns >= 1, "an array holds a value");
        Array {
            rows,
            columns,
            values: values.into(),
        }
    }

    /// An array of one value.
    pub fn single(v: Value) -> Array {
        Array::new(1, 1, vec![v])
    }

    pub fn rows(&self) -> u32 {
        self.rows
    }

    pub fn columns(&self) -> u32 {
        self.columns
    }

    /// Its values, in row-major order.
    pub fn values(&self) -> &[Value] {
        &self.values
    }

    /// How many values it holds.
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Never: an array holds a value. Here for the lint's pair to `len`.
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Whether it holds one value, which reads as that value wherever one
    /// value is wanted.
    pub fn is_single(&self) -> bool {
        self.rows == 1 && self.columns == 1
    }

    /// The value at a row and a column, each from 0.
    pub fn get(&self, row: u32, column: u32) -> &Value {
        &self.values[(row as usize) * (self.columns as usize) + column as usize]
    }

    /// The value at a place of a lift ([`shape_of`]): a row of its own, or
    /// its one row at every row, and likewise for its columns; `None` at a
    /// place it does not reach, which a lift answers `#N/A` at.
    pub fn at(&self, row: u32, column: u32) -> Option<&Value> {
        let r = if self.rows == 1 {
            0
        } else if row < self.rows {
            row
        } else {
            return None;
        };
        let c = if self.columns == 1 {
            0
        } else if column < self.columns {
            column
        } else {
            return None;
        };
        Some(self.get(r, c))
    }

    /// Whether it is one row or one column, a vector a lookup searches.
    pub fn is_vector(&self) -> bool {
        self.rows == 1 || self.columns == 1
    }
}

/// The label of an array past [`LIMIT`].
pub(crate) fn past_limit() -> Reason {
    Reason::Construct("an array past the grid's limit".to_string())
}

/// The rows and the columns of an array about to be made, as `u32`s, or the
/// label past the limit; `#CALC!` for an array of no row or no column,
/// Excel's answer for an empty one, given as `Ok(Err(..))` so that the
/// caller answers it as a value.
pub(crate) fn room(rows: u64, columns: u64) -> Result<Result<(u32, u32), ErrorKind>, Reason> {
    if rows == 0 || columns == 0 {
        return Ok(Err(ErrorKind::Calc));
    }
    match rows.checked_mul(columns) {
        Some(n) if n <= LIMIT as u64 => Ok(Ok((rows as u32, columns as u32))),
        _ => Err(past_limit()),
    }
}

/// The shape a lift over operands of these shapes answers: the most rows
/// and the most columns, a scalar standing as one by one.
pub(crate) fn shape_of(shapes: impl IntoIterator<Item = (u32, u32)>) -> (u32, u32) {
    shapes.into_iter().fold((1, 1), |(r, c), (rows, columns)| {
        (r.max(rows), c.max(columns))
    })
}

/// One operand of a lift: one value, the same at every place, or an array.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Lifted {
    One(Value),
    Many(Array),
}

impl Lifted {
    /// Its shape in a lift.
    pub(crate) fn shape(&self) -> (u32, u32) {
        match self {
            Lifted::One(_) => (1, 1),
            Lifted::Many(a) => (a.rows, a.columns),
        }
    }

    /// Its value at a place of a lift, `None` where it does not reach.
    pub(crate) fn at(&self, row: u32, column: u32) -> Option<&Value> {
        match self {
            Lifted::One(v) => Some(v),
            Lifted::Many(a) => a.at(row, column),
        }
    }
}

/// An operator of two values lifted over two operands, one of which is an
/// array: `f` at each place both reach, `#N/A` where one does not.
pub(crate) fn lift2<F>(a: &Lifted, b: &Lifted, mut f: F) -> Result<Array, Reason>
where
    F: FnMut(&Value, &Value) -> Result<Value, Reason>,
{
    let (rows, columns) = shape_of([a.shape(), b.shape()]);
    let (rows, columns) = match room(u64::from(rows), u64::from(columns))? {
        Ok(rc) => rc,
        // Never: every operand holds a value.
        Err(k) => return Ok(Array::single(Value::Error(k))),
    };
    let mut values = Vec::with_capacity((rows as usize) * (columns as usize));
    for r in 0..rows {
        for c in 0..columns {
            values.push(match (a.at(r, c), b.at(r, c)) {
                (Some(x), Some(y)) => f(x, y)?,
                _ => Value::Error(ErrorKind::NA),
            });
        }
    }
    Ok(Array::new(rows, columns, values))
}

/// A control function lifted over its operands (`IF`, `CHOOSE`, `IFERROR`):
/// the most rows and columns of them, its value at each place `f`'s, which
/// reads only the operand it takes there, so that an operand not taken
/// there may fall short without `#N/A`, as Excel's own save of the arrays'
/// fixture computes `IF({FALSE,FALSE,FALSE},{1,2},5)` (KERNEL.8, slice f).
pub(crate) fn lift_by<F>(operands: &[&Lifted], mut f: F) -> Result<Array, Reason>
where
    F: FnMut(u32, u32) -> Result<Value, Reason>,
{
    let (rows, columns) = shape_of(operands.iter().map(|o| o.shape()));
    let (rows, columns) = match room(u64::from(rows), u64::from(columns))? {
        Ok(rc) => rc,
        Err(k) => return Ok(Array::single(Value::Error(k))),
    };
    let mut values = Vec::with_capacity((rows as usize) * (columns as usize));
    for r in 0..rows {
        for c in 0..columns {
            values.push(f(r, c)?);
        }
    }
    Ok(Array::new(rows, columns, values))
}

/// An operator of one value lifted over an array.
pub(crate) fn lift1<F>(a: &Array, mut f: F) -> Result<Array, Reason>
where
    F: FnMut(&Value) -> Result<Value, Reason>,
{
    let mut values = Vec::with_capacity(a.len());
    for v in a.values() {
        values.push(f(v)?);
    }
    Ok(Array::new(a.rows, a.columns, values))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn n(x: f64) -> Value {
        Value::Number(x)
    }

    fn row(xs: &[f64]) -> Lifted {
        Lifted::Many(Array::new(
            1,
            xs.len() as u32,
            xs.iter().map(|&x| n(x)).collect(),
        ))
    }

    fn column(xs: &[f64]) -> Lifted {
        Lifted::Many(Array::new(
            xs.len() as u32,
            1,
            xs.iter().map(|&x| n(x)).collect(),
        ))
    }

    fn add(a: &Lifted, b: &Lifted) -> Array {
        lift2(a, b, |x, y| {
            Ok(n(x.number().unwrap() + y.number().unwrap()))
        })
        .unwrap()
    }

    #[test]
    fn a_lift_takes_the_most_rows_and_columns_and_n_a_where_one_falls_short() {
        // A row against a column is a table; a scalar stands everywhere; two
        // rows of different lengths reach as far as the longer, `#N/A` past
        // the shorter.
        let table = add(&row(&[1.0, 2.0, 3.0]), &column(&[10.0, 20.0]));
        assert_eq!((table.rows(), table.columns()), (2, 3));
        assert_eq!(
            table.values(),
            &[n(11.0), n(12.0), n(13.0), n(21.0), n(22.0), n(23.0)]
        );
        let scalar = add(&row(&[1.0, 2.0]), &Lifted::One(n(5.0)));
        assert_eq!(scalar.values(), &[n(6.0), n(7.0)]);
        let short = add(&row(&[1.0, 2.0, 3.0]), &row(&[10.0, 20.0]));
        assert_eq!(
            short.values(),
            &[n(11.0), n(22.0), Value::Error(ErrorKind::NA)]
        );
        let two = Lifted::Many(Array::new(2, 2, vec![n(1.0), n(2.0), n(3.0), n(4.0)]));
        let three = add(&two, &column(&[1.0, 1.0, 1.0]));
        assert_eq!((three.rows(), three.columns()), (3, 2));
        assert_eq!(three.get(2, 0), &Value::Error(ErrorKind::NA));
        assert_eq!(three.get(1, 1), &n(5.0));
    }

    #[test]
    fn room_refuses_past_the_limit_and_answers_calc_for_nothing() {
        assert_eq!(room(2, 3).unwrap(), Ok((2, 3)));
        assert_eq!(room(LIMIT as u64, 1).unwrap(), Ok((LIMIT as u32, 1)));
        assert!(room(LIMIT as u64 + 1, 1).is_err());
        assert!(room(u64::MAX, u64::MAX).is_err());
        assert_eq!(room(0, 5).unwrap(), Err(ErrorKind::Calc));
        assert_eq!(room(5, 0).unwrap(), Err(ErrorKind::Calc));
        assert_eq!(
            past_limit().spell(),
            "(not-computed \"an array past the grid's limit\")"
        );
    }

    #[test]
    fn an_array_is_shared_by_a_count() {
        let a = Array::new(1, 2, vec![n(1.0), n(2.0)]);
        let b = a.clone();
        assert!(Rc::ptr_eq(&a.values, &b.values));
        assert!(Array::single(n(1.0)).is_single() && !a.is_single() && a.is_vector());
    }
}
