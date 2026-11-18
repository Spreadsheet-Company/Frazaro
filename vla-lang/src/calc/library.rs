//! The language's own functions (KERNEL.7, 2026-10-08): what a cartridge
//! rule needs on day one, `Alonzo/SPEC.md` section 13's list to this item.
//!
//! `IF`, `AND`, `OR`, `NOT`, `SUM`, `MIN`, `MAX`, `ABS`, `INT`, `MOD`,
//! `ROW`, `COLUMN`, `CHOOSE` and `SIN`, with the comparisons and the
//! arithmetic the evaluator holds as operators: `ROW` and `COLUMN` because
//! a cell painting at the mouse must know where it is, `MOD` because every
//! counter wraps, `CHOOSE` for a score until `INDEX` comes, `SIN` for a
//! slow oscillator. Each is the formula language's own, with the semantics
//! ISO/IEC 29500 and OpenFormula publish and Excel computes, held here by
//! tests with exact values and no tolerance (`Alonzo/docs/ALGEBRA.md`
//! section 7, crack 1); an engine's replay golden is the oracle above
//! them. Excel's wider library registers through the same seam from
//! `frazaro-core`, by measurement, and a function the engine comes to
//! need moves down here with its tests.
//!
//! The rule every aggregate follows, Excel's: a value given directly is
//! coerced (a truth value is 1 or 0, a numeric text its number, another
//! text `#VALUE!`), a value read out of a range counts only when it is of
//! the kind the function wants (`SUM` skips a text and a truth value in a
//! range), an empty cell is skipped, and an error anywhere is the answer.
//!
//! The rule is written once, for numbers and for truth values ([`Rule`]),
//! and so is each aggregate's fold ([`Fold`]: where `SUM`, `MIN`, `MAX`,
//! `AND` and `OR` start, what each value they count does, what they answer),
//! and `NOT`'s answer ([`negated`]); a cell's evaluation drives them over its
//! arguments here, and an evaluation of many cells at once drives the same
//! rule and the same fold over its own (`super::vector`, KERNEL.26), so that
//! no meaning of the library is written twice.
//!
//! So is the meaning of each of the others (KERNEL.28): `ABS`, `INT`, `SIN`
//! and `MOD` over the numbers their arguments read as ([`abs_of`],
//! [`int_of`], [`sin_of`], [`mod_of`]), `ROW` and `COLUMN` over the cell or
//! the reference they place ([`placed`]), and `CHOOSE`'s index ([`chosen`]).
//! The functions below call them after their own coercions, and an
//! evaluation of many cells at once calls them over its lanes, a function of
//! numbers where every operand at a cell is a number and the function below
//! at any other cell, so that a coercion is written here alone.
//!
//! `BITAND`, `BITOR` and `BITXOR` (KERNEL.71, 2026-10-30) stand beside `MOD`, which
//! reads its numbers as they do: a game's marks read as bit masks by one call
//! where its source wrote `MOD(INT(x/2^k),2)` (Alonzo's `CART.14.1`). One meaning
//! ([`bitwise`]), held by Excel's save of `scripts/recalc/bits.txt` to what the
//! pages of `BITOR` and `BITXOR` do not say: a number is whole exactly, on all
//! three, with no tolerance at any magnitude ([`whole_bits`]). They need no arm of
//! their own in `super::program`, which calls them through their pointers.

use super::array::{self, Array, Lifted};
use super::eval::{finite, intersection, ranges, values, Arg, Ctx, Known, Library, Param, Reason};
use super::formula::Expr;
use super::graph::Areas;
use super::value::{round15, ErrorKind, Value};
use super::{
    aggregates, arrays, criteria, database, dates, formulas, helpers, information, lookups, texts,
};

// The classes of the language's functions' parameters (KERNEL.8, slice f),
// each by its place from 0, as Excel's documentation of each function gives
// them: a value, lifted over an array given; a reference or an array, handed
// whole; an array, its expression read as arrays (`super::eval::Param`).

/// `MATCH`'s and `XMATCH`'s: the value sought, the array searched, values
/// after.
fn searched(j: usize) -> Param {
    if j == 1 {
        Param::Array
    } else {
        Param::Value
    }
}

/// `VLOOKUP`'s and `HLOOKUP`'s: the table, and every other a value forced
/// in a legacy formula, as Excel's own saves of the arrays' fixture read
/// `VLOOKUP(3,A1:C6,{1,3},FALSE)` there as its first column and
/// `VLOOKUP({3,4},A1:C6,3,FALSE)` as its first value's row (KERNEL.8,
/// slice f), the match's kind with them. The value sought and the match's
/// kind are forced in a legacy array formula too, as Excel's save of
/// `KERNEL.55`'s probe holds `HLOOKUP({"b","c"},...)` and a kind of
/// `{FALSE,TRUE}` there; the index, which no save shows there, lifts in one
/// as in a dynamic formula.
fn looked_across(j: usize) -> Param {
    match j {
        1 => Param::Array,
        0 | 3 => Param::Forced,
        _ => Param::Scalar,
    }
}

/// `LOOKUP`'s: the value sought, then two arrays.
fn looked(j: usize) -> Param {
    if j == 0 {
        Param::Value
    } else {
        Param::Array
    }
}

/// `COUNTIF`'s, `SUMIF`'s and `AVERAGEIF`'s, and the database functions': a
/// range, a criterion (or a field), a range.
fn criterion(j: usize) -> Param {
    if j == 1 {
        Param::Value
    } else {
        Param::Range
    }
}

/// `SUMIFS`'s, `AVERAGEIFS`'s, `MAXIFS`'s and `MINIFS`'s: a range, then a
/// range and a criterion in turn.
fn criteria_after(j: usize) -> Param {
    if j >= 2 && j % 2 == 0 {
        Param::Value
    } else {
        Param::Range
    }
}

/// `COUNTIFS`'s: a range and a criterion in turn.
fn criteria_pairs(j: usize) -> Param {
    if j % 2 == 1 {
        Param::Value
    } else {
        Param::Range
    }
}

/// `SUBTOTAL`'s: the function's number, then references.
fn subtotaled(j: usize) -> Param {
    if j == 0 {
        Param::Value
    } else {
        Param::Range
    }
}

/// `NETWORKDAYS`'s and `WORKDAY`'s: two values, then the holidays.
fn holidays(j: usize) -> Param {
    if j >= 2 {
        Param::Range
    } else {
        Param::Value
    }
}

/// `TEXTJOIN`'s: the delimiter, which may be a range of them used in turn,
/// whether to skip empty texts, then the texts.
fn joined(j: usize) -> Param {
    if j == 1 {
        Param::Value
    } else {
        Param::Range
    }
}

/// The language's library, every function above registered, each as one an
/// evaluation of many cells at once computes by itself: `IF` as the form it
/// computes as `if_` computes it at one (KERNEL.20), the aggregates and `NOT`
/// by the same rule and fold (KERNEL.26), and the others, `CHOOSE` among
/// them as a form, by the same meanings (KERNEL.28). Since KERNEL.8 the
/// lookups and the error tests a lookup's miss is read by
/// (`super::lookups`), and `COUNTIF` and `SUMIF` with their criteria
/// (`super::criteria`), moved down from Excel's library: a run computes
/// `INDEX`, `IFERROR` and `IFNA` by itself, and calls the rest at each cell.
/// Since KERNEL.8's second slice `SINGLE`, Excel's `@`, a form a run
/// computes at each cell; since its third the text functions
/// (`super::texts`), since its fourth the date functions
/// (`super::dates`), and since its fifth the aggregates and `SUBTOTAL`
/// (`super::aggregates`, `AVERAGE` moved down from Excel's library),
/// `SUMIFS` and its family (`super::criteria`) and the database functions
/// (`super::database`), which a run calls at each cell. Since KERNEL.53
/// `ISFORMULA` and `FORMULATEXT` (`super::formulas`), a cell's formula read
/// by its reference, each answering an operand so that every path computes
/// it as `Ctx` does. Since KERNEL.56 the functions that read a sheet's
/// extent itself are declared so ([`READS_EXTENT`]). Since KERNEL.60
/// `ADDRESS` (`super::texts`), a cell's address as a text from its row and
/// its column, a form a run computes at each cell. Since KERNEL.68 the
/// information functions (`super::information`), and `ROUND` and `IFS` moved
/// down from Excel's library with `ISNUMBER` and `ISBLANK`, so that a door
/// computes nothing an engine refuses. Since KERNEL.71 `BITAND`, `BITOR` and
/// `BITXOR` ([`bitwise`]), a mask read by one call, which a program calls through
/// its pointer.
pub fn language() -> Library {
    let mut lib = Library::empty();
    lib.register_known("IF", if_, Known::If);
    lib.register_known_strict("AND", and, Known::And, ranges);
    lib.register_known_strict("OR", or, Known::Or, ranges);
    lib.register_known_strict("NOT", not, Known::Not, values);
    lib.register_known_strict("SUM", sum, Known::Sum, ranges);
    lib.register_known_strict("MIN", min, Known::Min, ranges);
    lib.register_known_strict("MAX", max, Known::Max, ranges);
    lib.register_known_strict("ABS", abs, Known::Abs, values);
    lib.register_known_strict("INT", int, Known::Int, values);
    lib.register_known_strict("MOD", modulo, Known::Mod, values);
    lib.register_known_strict("ROW", row, Known::Row, ranges);
    lib.register_known_strict("COLUMN", column, Known::Column, ranges);
    lib.register_known("CHOOSE", choose, Known::Choose);
    lib.register_known_strict("SIN", sin, Known::Sin, values);
    lib.register_known("INDEX", lookups::index, Known::Index);
    lib.register_known_strict("MATCH", lookups::match_, Known::Match, searched);
    lib.register_known_strict("XMATCH", lookups::xmatch, Known::XMatch, searched);
    lib.register_known_strict("VLOOKUP", lookups::vlookup, Known::VLookup, looked_across);
    lib.register_known_strict("HLOOKUP", lookups::hlookup, Known::HLookup, looked_across);
    lib.register_known_strict("LOOKUP", lookups::lookup, Known::Lookup, looked);
    lib.register_form("XLOOKUP", lookups::xlookup);
    lib.register_known("IFERROR", lookups::iferror, Known::IfError);
    lib.register_known("IFNA", lookups::ifna, Known::IfNa);
    lib.register_known_strict("ISERROR", lookups::iserror, Known::IsError, values);
    lib.register_known_strict("ISNA", lookups::isna, Known::IsNa, values);
    lib.register_known_strict("COUNTIF", criteria::countif, Known::CountIf, criterion);
    lib.register_known_strict("SUMIF", criteria::sumif, Known::SumIf, criterion);
    lib.register_form("SINGLE", single);
    lib.register("CONCATENATE", texts::concatenate);
    lib.register_classed("CONCAT", texts::concat, ranges);
    lib.register_classed("TEXTJOIN", texts::textjoin, joined);
    lib.register("LEFT", texts::left);
    lib.register("RIGHT", texts::right);
    lib.register("MID", texts::mid);
    lib.register("LEN", texts::len);
    lib.register("FIND", texts::find);
    lib.register("SEARCH", texts::search);
    lib.register("SUBSTITUTE", texts::substitute);
    lib.register("REPLACE", texts::replace);
    lib.register("UPPER", texts::upper);
    lib.register("LOWER", texts::lower);
    lib.register("PROPER", texts::proper);
    lib.register("TRIM", texts::trim);
    lib.register("REPT", texts::rept);
    lib.register("EXACT", texts::exact);
    lib.register("VALUE", texts::value);
    lib.register("T", texts::t);
    lib.register("N", texts::n);
    lib.register("UNICODE", texts::unicode);
    lib.register("UNICHAR", texts::unichar);
    lib.register("CHAR", texts::char_);
    lib.register("CODE", texts::code);
    lib.register("TEXT", texts::text_);
    lib.register_form("ADDRESS", texts::address);
    lib.register("BITAND", bitand);
    lib.register("BITOR", bitor);
    lib.register("BITXOR", bitxor);
    lib.register("DATE", dates::date);
    lib.register("YEAR", dates::year);
    lib.register("MONTH", dates::month);
    lib.register("DAY", dates::day_);
    lib.register("WEEKDAY", dates::weekday_);
    lib.register("EOMONTH", dates::eomonth);
    lib.register("EDATE", dates::edate);
    lib.register("DAYS", dates::days);
    lib.register_classed("NETWORKDAYS", dates::networkdays, holidays);
    lib.register_classed("WORKDAY", dates::workday, holidays);
    lib.register("TIME", dates::time);
    lib.register("HOUR", dates::hour);
    lib.register("MINUTE", dates::minute);
    lib.register("SECOND", dates::second);
    lib.register("DATEDIF", dates::datedif);
    lib.register("YEARFRAC", dates::yearfrac);
    lib.register("WEEKNUM", dates::weeknum);
    lib.register_classed("AVERAGE", aggregates::average, ranges);
    lib.register_classed("COUNT", aggregates::count, ranges);
    lib.register_classed("COUNTA", aggregates::counta, ranges);
    lib.register_classed("PRODUCT", aggregates::product, ranges);
    lib.register_classed("STDEV", aggregates::stdev_s, ranges);
    lib.register_classed("STDEV.S", aggregates::stdev_s, ranges);
    lib.register_classed("STDEVP", aggregates::stdev_p, ranges);
    lib.register_classed("STDEV.P", aggregates::stdev_p, ranges);
    lib.register_classed("VAR", aggregates::var_s, ranges);
    lib.register_classed("VAR.S", aggregates::var_s, ranges);
    lib.register_classed("VARP", aggregates::var_p, ranges);
    lib.register_classed("VAR.P", aggregates::var_p, ranges);
    lib.register_classed("SUBTOTAL", aggregates::subtotal, subtotaled);
    lib.register_classed("SUMIFS", criteria::sumifs, criteria_after);
    lib.register_classed("COUNTIFS", criteria::countifs, criteria_pairs);
    lib.register_classed("AVERAGEIF", criteria::averageif, criterion);
    lib.register_classed("AVERAGEIFS", criteria::averageifs, criteria_after);
    lib.register_classed("MAXIFS", criteria::maxifs, criteria_after);
    lib.register_classed("MINIFS", criteria::minifs, criteria_after);
    lib.register_classed("DSUM", database::dsum, criterion);
    lib.register_classed("DAVERAGE", database::daverage, criterion);
    lib.register_classed("DCOUNT", database::dcount, criterion);
    lib.register_classed("DCOUNTA", database::dcounta, criterion);
    lib.register_classed("DGET", database::dget, criterion);
    lib.register_classed("DMAX", database::dmax, criterion);
    lib.register_classed("DMIN", database::dmin, criterion);
    lib.register_classed("DPRODUCT", database::dproduct, criterion);
    lib.register_classed("DSTDEV", database::dstdev, criterion);
    lib.register_classed("DSTDEVP", database::dstdevp, criterion);
    lib.register_classed("DVAR", database::dvar, criterion);
    lib.register_classed("DVARP", database::dvarp, criterion);
    arrays::register(&mut lib);
    lib.register_arrayed("ISFORMULA", formulas::isformula, ranges);
    lib.register_arrayed("FORMULATEXT", formulas::formulatext, ranges);
    for (name, f) in INFORMATION {
        lib.register_known_strict(name, f, Known::Information, values);
    }
    lib.register_form("ISREF", information::isref);
    lib.register_form("TYPE", information::type_);
    lib.register_form("ISEVEN", information::iseven);
    lib.register_form("ISODD", information::isodd);
    lib.register("ROUND", round);
    lib.register_form("IFS", ifs);
    // `LET` and `LAMBDA` are read by the parse (KERNEL.29, `super::lambda`),
    // which leaves no call of either; a form of each holds the name, so that
    // a text calling them, read or not, calls no function the library lacks.
    lib.register_form("LET", bound);
    lib.register_form("LAMBDA", bound);
    helpers::register(&mut lib);
    for name in READS_EXTENT {
        lib.declare_reads_extent(name);
    }
    for name in READS_KINDS {
        lib.declare_reads_kinds(name);
    }
    lib
}

/// The information functions of one value, which read what kind it is, a
/// function's among them (KERNEL.29, [`Library::declare_reads_kinds`]).
const READS_KINDS: [&str; 9] = [
    "ISERROR",
    "ISNA",
    "ISNUMBER",
    "ISTEXT",
    "ISNONTEXT",
    "ISLOGICAL",
    "ISBLANK",
    "ISERR",
    "ERROR.TYPE",
];

/// `LET` or `LAMBDA` as a call, which a parse never leaves (KERNEL.29): a
/// malformed binding's answer, `#VALUE!`.
fn bound<'e>(_: &Ctx<'_>, _: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    Ok(Arg::Scalar(Value::Error(ErrorKind::Value)))
}

/// The information functions of one value (KERNEL.68, `super::information`).
const INFORMATION: [(&str, super::eval::Strict); 10] = [
    ("ISNUMBER", information::isnumber),
    ("ISTEXT", information::istext),
    ("ISNONTEXT", information::isnontext),
    ("ISLOGICAL", information::islogical),
    ("ISBLANK", information::isblank),
    ("ISERR", information::iserr),
    ("ERROR.TYPE", information::error_type),
    ("NA", information::na),
    ("TRUE", information::true_),
    ("FALSE", information::false_),
];

/// The language's functions whose answer the extent of a sheet they read
/// decides beside its cells (KERNEL.56), each declared so
/// ([`Library::declare_reads_extent`]): the lookups whose approximate match
/// bisects a vector bounded by its sheet's extent (`super::index`), over a
/// vector not sorted answering by where the extent ends; and the database
/// functions, whose records end at their sheet's last row
/// (`super::database`), as Excel's save settled it. No other function of
/// the language reads an extent but to clamp a walk of the cells it reads,
/// none of which stands outside it.
pub(crate) const READS_EXTENT: [&str; 18] = [
    "MATCH", "XMATCH", "VLOOKUP", "HLOOKUP", "LOOKUP", "XLOOKUP", "DSUM", "DAVERAGE", "DCOUNT",
    "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP",
];

/// The day-one list the language's library began with (KERNEL.7), the
/// functions `Alonzo/SPEC.md` section 13 named. A fixed array whose length is
/// its type, so it stays as it was when the library grew (KERNEL.8's
/// decision 9): a page lists what a library holds by `Library::names`.
pub const NAMES: [&str; 14] = [
    "IF", "AND", "OR", "NOT", "SUM", "MIN", "MAX", "ABS", "INT", "MOD", "ROW", "COLUMN", "CHOOSE",
    "SIN",
];

/// The functions KERNEL.8 brought into the language, beside the day-one
/// list: the lookups, `SINGLE`, the text functions, the date functions, the
/// aggregates and `SUBTOTAL`, the `SUMIFS` family, the database functions,
/// and the array functions.
#[cfg(test)]
const KERNEL_8: [&str; 108] = [
    "INDEX",
    "MATCH",
    "XMATCH",
    "VLOOKUP",
    "HLOOKUP",
    "LOOKUP",
    "XLOOKUP",
    "IFERROR",
    "IFNA",
    "ISERROR",
    "ISNA",
    "COUNTIF",
    "SUMIF",
    "SINGLE",
    "CONCATENATE",
    "CONCAT",
    "TEXTJOIN",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "UPPER",
    "LOWER",
    "PROPER",
    "TRIM",
    "REPT",
    "EXACT",
    "VALUE",
    "T",
    "N",
    "UNICODE",
    "UNICHAR",
    "CHAR",
    "CODE",
    "TEXT",
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "WEEKDAY",
    "EOMONTH",
    "EDATE",
    "DAYS",
    "NETWORKDAYS",
    "WORKDAY",
    "TIME",
    "HOUR",
    "MINUTE",
    "SECOND",
    "DATEDIF",
    "YEARFRAC",
    "WEEKNUM",
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "PRODUCT",
    "STDEV",
    "STDEV.S",
    "STDEVP",
    "STDEV.P",
    "VAR",
    "VAR.S",
    "VARP",
    "VAR.P",
    "SUBTOTAL",
    "SUMIFS",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "MAXIFS",
    "MINIFS",
    "DSUM",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DGET",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DVAR",
    "DVARP",
    "SEQUENCE",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "TRANSPOSE",
    "SUMPRODUCT",
    "MMULT",
    "ROWS",
    "COLUMNS",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "TAKE",
    "DROP",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "WRAPROWS",
    "WRAPCOLS",
    "EXPAND",
];

/// The functions KERNEL.53 brought: a cell's formula read by its reference.
#[cfg(test)]
const KERNEL_53: [&str; 2] = ["ISFORMULA", "FORMULATEXT"];

/// The function KERNEL.60 brought: a cell's address as a text.
#[cfg(test)]
const KERNEL_60: [&str; 1] = ["ADDRESS"];

/// The functions KERNEL.68 brought: the information functions, and `ROUND` and
/// `IFS` from Excel's library beside `ISNUMBER` and `ISBLANK`.
#[cfg(test)]
const KERNEL_68: [&str; 14] = [
    "ISNUMBER",
    "ISTEXT",
    "ISNONTEXT",
    "ISLOGICAL",
    "ISBLANK",
    "ISERR",
    "ISREF",
    "ISEVEN",
    "ISODD",
    "ERROR.TYPE",
    "TYPE",
    "NA",
    "ROUND",
    "IFS",
];

/// The functions KERNEL.71 brought: a mask read by one call.
#[cfg(test)]
const KERNEL_71: [&str; 3] = ["BITAND", "BITOR", "BITXOR"];

/// The functions KERNEL.58 brought: a boolean as OpenFormula spells one.
#[cfg(test)]
const KERNEL_58: [&str; 2] = ["TRUE", "FALSE"];

/// The function KERNEL.55 brought: a spill read through `A1#`.
#[cfg(test)]
const KERNEL_55: [&str; 1] = ["ANCHORARRAY"];

/// The functions KERNEL.29 brought: `LET` and `LAMBDA`, whose names a form
/// holds, and the functions that take a function (`super::helpers`).
#[cfg(test)]
const KERNEL_29: [&str; 9] = [
    "LET",
    "LAMBDA",
    "MAP",
    "REDUCE",
    "SCAN",
    "BYROW",
    "BYCOL",
    "MAKEARRAY",
    "ISOMITTED",
];

/// What a numeric aggregate collects: the numbers, or the first error.
pub type Numbers = Result<Vec<f64>, ErrorKind>;

/// What an aggregate makes of a value read out of a range (KERNEL.26): an
/// item it counts, an error it keeps to answer with once the range is
/// walked, or nothing.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum Ranged<T> {
    Counts(T),
    Error(ErrorKind),
    Skipped,
}

/// The aggregate rule above for one kind of item (KERNEL.26): how a value
/// given directly is coerced, and what a value read out of a range is.
pub(crate) trait Rule {
    type Item: Copy;
    /// A value given directly: its item, or the error that is the answer.
    fn given(v: &Value) -> Result<Self::Item, ErrorKind>;
    /// A value read out of a range.
    fn ranged(v: &Value) -> Ranged<Self::Item>;
    /// [`Rule::given`] of a number and of a truth value, and
    /// [`Rule::ranged`] of a number, each an item (KERNEL.43): what an
    /// evaluation of many cells folds a lane of numbers or of truth values
    /// and an array's slices of numbers by, with no value made for each.
    fn given_number(n: f64) -> Self::Item;
    fn given_truth(b: bool) -> Self::Item;
    fn ranged_number(n: f64) -> Self::Item;
}

/// The rule for numbers: a value given directly reads as a number or is
/// `#VALUE!`; in a range only a number counts.
pub(crate) struct OverNumbers;

impl Rule for OverNumbers {
    type Item = f64;

    fn given(v: &Value) -> Result<f64, ErrorKind> {
        v.number()
    }

    fn ranged(v: &Value) -> Ranged<f64> {
        match v {
            Value::Number(n) => Ranged::Counts(*n),
            Value::Error(k) => Ranged::Error(*k),
            _ => Ranged::Skipped,
        }
    }

    #[inline]
    fn given_number(n: f64) -> f64 {
        n
    }

    #[inline]
    fn given_truth(b: bool) -> f64 {
        if b {
            1.0
        } else {
            0.0
        }
    }

    #[inline]
    fn ranged_number(n: f64) -> f64 {
        n
    }
}

/// The rule for truth values: a value given directly must read as one
/// (`#VALUE!` otherwise); in a range a truth value counts, a number counts
/// as `0` or not, a text is skipped.
pub(crate) struct OverTruths;

impl Rule for OverTruths {
    type Item = bool;

    fn given(v: &Value) -> Result<bool, ErrorKind> {
        v.truth()
    }

    fn ranged(v: &Value) -> Ranged<bool> {
        match v {
            Value::Bool(b) => Ranged::Counts(*b),
            Value::Number(n) => Ranged::Counts(*n != 0.0),
            Value::Error(k) => Ranged::Error(*k),
            _ => Ranged::Skipped,
        }
    }

    #[inline]
    fn given_number(n: f64) -> bool {
        n != 0.0
    }

    #[inline]
    fn given_truth(b: bool) -> bool {
        b
    }

    #[inline]
    fn ranged_number(n: f64) -> bool {
        n != 0.0
    }
}

/// Each item of the arguments by a rule, to `f` in the arguments' order and
/// a range's walk: a value given directly that is an error, or does not read
/// as an item, answers at once; a range is walked to its end and answers
/// with its first error, so that a cell in it not computed still answers
/// with its reason. An aggregate folds over this as it walks, and nothing is
/// collected (KERNEL.25).
fn each_item<R: Rule, F: FnMut(R::Item)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    mut f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    for a in args {
        match a {
            Arg::Scalar(v) => match R::given(v) {
                Ok(x) => f(x),
                Err(k) => return Ok(Err(k)),
            },
            // A function's value, read as one value (KERNEL.29).
            Arg::Function(_) => return Ok(Err(ErrorKind::Value)),
            Arg::Area { areas, .. } => {
                let mut err: Option<ErrorKind> = None;
                ctx.each(areas, |v| {
                    match R::ranged(&v) {
                        Ranged::Counts(x) => f(x),
                        Ranged::Error(k) => {
                            err.get_or_insert(k);
                        }
                        Ranged::Skipped => {}
                    }
                    Ok(())
                })?;
                if let Some(k) = err {
                    return Ok(Err(k));
                }
            }
            // An array's values are read as a range's are, as Excel's
            // documentation of `SUM` says of both (KERNEL.8, slice f).
            Arg::Array(a) => {
                let mut err: Option<ErrorKind> = None;
                for v in a.values() {
                    match R::ranged(v) {
                        Ranged::Counts(x) => f(x),
                        Ranged::Error(k) => {
                            err.get_or_insert(k);
                        }
                        Ranged::Skipped => {}
                    }
                }
                if let Some(k) = err {
                    return Ok(Err(k));
                }
            }
        }
    }
    Ok(Ok(()))
}

/// Each number of the arguments by the aggregate rule above, to `f`, as
/// [`each_item`] walks them ([`OverNumbers`]).
pub fn each_number<F: FnMut(f64)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    each_item::<OverNumbers, F>(ctx, args, f)
}

/// Each truth value of the arguments by the same rule, to `f`, as
/// [`each_number`] walks ([`OverTruths`]).
pub fn each_truth<F: FnMut(bool)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    each_item::<OverTruths, F>(ctx, args, f)
}

/// An aggregate's meaning (KERNEL.26): its rule, where it starts, what each
/// item it counts does, and its answer once its arguments are walked. A
/// cell's evaluation drives it through [`folded`]; an evaluation of many
/// cells at once drives it through its own walk (`super::vector`).
pub(crate) trait Fold: Default {
    type Rule: Rule;
    fn add(&mut self, x: <Self::Rule as Rule>::Item);
    fn answer(self) -> Value;
}

/// `SUM` adds in the walk's order from -0.0, `Iterator::sum`'s own start,
/// so that its bits are those of the numbers collected and summed.
pub(crate) struct Sum(f64);

impl Default for Sum {
    fn default() -> Sum {
        Sum(-0.0)
    }
}

impl Fold for Sum {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.0 += n;
    }

    fn answer(self) -> Value {
        finite(self.0)
    }
}

/// `MIN` of no number is 0.
pub(crate) struct Least {
    least: f64,
    any: bool,
}

impl Default for Least {
    fn default() -> Least {
        Least {
            least: f64::INFINITY,
            any: false,
        }
    }
}

impl Fold for Least {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.least = self.least.min(n);
        self.any = true;
    }

    fn answer(self) -> Value {
        Value::Number(self.least.min(if self.any { f64::INFINITY } else { 0.0 }))
    }
}

/// `MAX` of no number is 0.
pub(crate) struct Most {
    most: f64,
    any: bool,
}

impl Default for Most {
    fn default() -> Most {
        Most {
            most: f64::NEG_INFINITY,
            any: false,
        }
    }
}

impl Fold for Most {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.most = self.most.max(n);
        self.any = true;
    }

    fn answer(self) -> Value {
        Value::Number(
            self.most
                .max(if self.any { f64::NEG_INFINITY } else { 0.0 }),
        )
    }
}

/// `AND` of no truth value is `#VALUE!`.
pub(crate) struct All {
    all: bool,
    any: bool,
}

impl Default for All {
    fn default() -> All {
        All {
            all: true,
            any: false,
        }
    }
}

impl Fold for All {
    type Rule = OverTruths;

    fn add(&mut self, b: bool) {
        self.all &= b;
        self.any = true;
    }

    fn answer(self) -> Value {
        if self.any {
            Value::Bool(self.all)
        } else {
            Value::Error(ErrorKind::Value)
        }
    }
}

/// `OR` of no truth value is `#VALUE!`.
#[derive(Default)]
pub(crate) struct Some_ {
    some: bool,
    any: bool,
}

impl Fold for Some_ {
    type Rule = OverTruths;

    fn add(&mut self, b: bool) {
        self.some |= b;
        self.any = true;
    }

    fn answer(self) -> Value {
        if self.any {
            Value::Bool(self.some)
        } else {
            Value::Error(ErrorKind::Value)
        }
    }
}

/// An aggregate over a cell's arguments: its fold driven by its rule's walk.
fn folded<F: Fold>(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    let mut fold = F::default();
    Ok(match each_item::<F::Rule, _>(ctx, args, |x| fold.add(x))? {
        Err(k) => Value::Error(k),
        Ok(()) => fold.answer(),
    })
}

/// `NOT`'s answer over the one value it reads (KERNEL.26).
pub(crate) fn negated(v: &Value) -> Value {
    match v.truth() {
        Ok(b) => Value::Bool(!b),
        Err(k) => Value::Error(k),
    }
}

/// `ABS`'s answer over the number its argument reads as (KERNEL.28).
pub(crate) fn abs_of(n: f64) -> Value {
    Value::Number(n.abs())
}

/// `INT`'s, rounding down to the nearest integer, so that `INT(-1.5)` is -2.
pub(crate) fn int_of(n: f64) -> Value {
    Value::Number(n.floor())
}

/// `SIN`'s, of an angle in radians.
pub(crate) fn sin_of(n: f64) -> Value {
    finite(n.sin())
}

/// `MOD(n, d)`'s: `n - d * INT(n / d)`, so that the result takes the
/// divisor's sign; a divisor of 0 is `#DIV/0!`.
pub(crate) fn mod_of(n: f64, d: f64) -> Value {
    if d == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    finite(n - d * (n / d).floor())
}

/// Which of a cell's coordinates `ROW` and `COLUMN` answer (KERNEL.28).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Axis {
    Row,
    Column,
}

impl Axis {
    /// A cell's row or column, as a number.
    pub(crate) fn of(self, row: u32, column: u32) -> f64 {
        f64::from(match self {
            Axis::Row => row,
            Axis::Column => column,
        })
    }
}

/// `ROW`'s and `COLUMN`'s answer (KERNEL.28): the row or the column of the
/// cell its argument places, the cell being computed or a reference's first
/// cell; `#VALUE!` for an argument that places none, a value or a reference
/// to no cell (a sheet the grid does not hold, a cell moved off the sheet),
/// which reaches it as the value `#REF!`.
pub(crate) fn placed(axis: Axis, cell: Option<(u32, u32)>) -> Value {
    match cell {
        Some((row, column)) => Value::Number(axis.of(row, column)),
        None => Value::Error(ErrorKind::Value),
    }
}

/// What `INDEX` picks of an area `rows` high and `cols` wide (KERNEL.8), as
/// offsets from its corner, its first row and column and its last: the row
/// and the column, each read down to a whole number; a `column` left out
/// takes the one number as the column of an area one row high and as the
/// row of any other; a row of 0 is every row, a column of 0 every column;
/// `#VALUE!` below 0 and `#REF!` past the area. One meaning for the walk,
/// a run and a program (`super::lookups`, `super::vector`,
/// `super::program`).
pub(crate) fn picked(
    rows: u32,
    cols: u32,
    row: f64,
    column: Option<f64>,
) -> Result<(u32, u32, u32, u32), ErrorKind> {
    let (row, column) = (row.trunc(), column.map(f64::trunc));
    if row < 0.0 || column.is_some_and(|c| c < 0.0) {
        return Err(ErrorKind::Value);
    }
    let (r, c) = match column {
        Some(c) => (row, c),
        None if rows == 1 => (0.0, row),
        None => (row, 0.0),
    };
    if r > f64::from(rows) || c > f64::from(cols) {
        return Err(ErrorKind::Ref);
    }
    let (r, c) = (r as u32, c as u32);
    let (top, bottom) = if r == 0 {
        (0, rows - 1)
    } else {
        (r - 1, r - 1)
    };
    let (left, right) = if c == 0 {
        (0, cols - 1)
    } else {
        (c - 1, c - 1)
    };
    Ok((top, left, bottom, right))
}

/// `CHOOSE`'s index (KERNEL.28): which of its `count` values it picks, the
/// index read down to a whole number; an error the index's own, and
/// `#VALUE!` outside 1 to the count.
pub(crate) fn chosen(index: &Value, count: usize) -> Result<usize, ErrorKind> {
    let k = index.number()?.floor();
    if k < 1.0 || k > count as f64 {
        return Err(ErrorKind::Value);
    }
    Ok(k as usize)
}

/// The numbers of the arguments by the aggregate rule above, collected.
pub fn numbers(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Numbers, Reason> {
    let mut out = Vec::new();
    let walked = each_number(ctx, args, |n| out.push(n))?;
    Ok(walked.map(|()| out))
}

/// The truth values of the arguments by the same rule, collected.
pub fn truths(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Result<Vec<bool>, ErrorKind>, Reason> {
    let mut out = Vec::new();
    let walked = each_truth(ctx, args, |b| out.push(b))?;
    Ok(walked.map(|()| out))
}

/// One numeric argument, coerced as an operand is.
pub fn number_arg(
    ctx: &Ctx<'_>,
    args: &[Arg],
    at: usize,
) -> Result<Result<f64, ErrorKind>, Reason> {
    let Some(arg) = args.get(at) else {
        return Ok(Err(ErrorKind::Value));
    };
    Ok(ctx.one(arg)?.number())
}

/// `IF`, which an evaluation of many cells at once computes by itself
/// (`super::vector`, its `if_`): a change here is a change there, and the
/// machine's tests hold the two to each other (KERNEL.20).
fn if_<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    let condition = match kept_class(ctx, &args[0])? {
        Lifted::One(v) => match v.truth() {
            Ok(b) => b,
            Err(k) => return Ok(Arg::Scalar(Value::Error(k))),
        },
        Lifted::Many(conditions) => return lifted_if(ctx, args, conditions),
    };
    // A branch left out reads as 0 when taken; a missing else is FALSE.
    match (condition, args.get(1), args.get(2)) {
        (true, Some(Expr::Empty), _) | (false, _, Some(Expr::Empty)) => {
            Ok(Arg::Scalar(Value::Number(0.0)))
        }
        (true, Some(then), _) => ctx.eval(then),
        (false, _, Some(otherwise)) => ctx.eval(otherwise),
        (false, _, None) => Ok(Arg::Scalar(Value::Bool(false))),
        (true, None, _) => Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
    }
}

/// A control function's condition or index as an operand (KERNEL.8, slice
/// f): in a legacy formula it keeps its own class, Excel's, so that an
/// array given lifts the function while a range is met at the formula's
/// cell even inside an array parameter (`SUMPRODUCT(IF(A1:A9>2,1,0))`
/// wants Ctrl+Shift+Enter); a dynamic formula reads it as arrays.
fn kept_class(ctx: &Ctx<'_>, e: &Expr) -> Result<Lifted, Reason> {
    if ctx.reads_arrays_by_kind() {
        ctx.operand(e)
    } else {
        ctx.as_values(|| ctx.operand(e))
    }
}

/// `IF` of an array of conditions (KERNEL.8, slice f): both branches
/// computed, and at each place of the lift the value of the branch its
/// condition takes there, `#N/A` where that branch falls short and a
/// branch not taken free to, as Excel's own save computes
/// `IF({FALSE,FALSE,FALSE},{1,2},5)` as three 5s; an error condition its
/// error; a branch left out reads as 0, a missing else as FALSE, as one
/// condition reads them.
fn lifted_if(ctx: &Ctx<'_>, args: &[Expr], conditions: Array) -> Result<Arg<'static>, Reason> {
    let branch = |e: Option<&Expr>| -> Result<Lifted, Reason> {
        match e {
            None => Ok(Lifted::One(Value::Bool(false))),
            Some(Expr::Empty) => Ok(Lifted::One(Value::Number(0.0))),
            Some(e) => ctx.operand(e),
        }
    };
    let yes = branch(args.get(1))?;
    let no = branch(args.get(2))?;
    let conditions = Lifted::Many(conditions);
    let na = || Value::Error(ErrorKind::NA);
    let lifted = array::lift_by(&[&conditions, &yes, &no], |r, c| {
        Ok(match conditions.at(r, c).map(Value::truth) {
            None => na(),
            Some(Ok(true)) => yes.at(r, c).cloned().unwrap_or_else(na),
            Some(Ok(false)) => no.at(r, c).cloned().unwrap_or_else(na),
            Some(Err(k)) => Value::Error(k),
        })
    })?;
    Ok(Arg::Array(lifted))
}

fn and(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<All>(ctx, args)
}

fn or(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Some_>(ctx, args)
}

fn not(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(negated(&ctx.one(&args[0])?))
}

fn sum(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Sum>(ctx, args)
}

fn min(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Least>(ctx, args)
}

fn max(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Most>(ctx, args)
}

/// `ABS` ([`abs_of`]).
fn abs(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => abs_of(n),
        Err(k) => Value::Error(k),
    })
}

/// `INT` ([`int_of`]).
fn int(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => int_of(n),
        Err(k) => Value::Error(k),
    })
}

/// `MOD` ([`mod_of`]), the first argument's error before the second's.
fn modulo(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let n = match number_arg(ctx, args, 0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let d = match number_arg(ctx, args, 1)? {
        Ok(d) => d,
        Err(k) => return Ok(Value::Error(k)),
    };
    Ok(mod_of(n, d))
}

/// One past the largest number `BITAND`, `BITOR` and `BITXOR` take, 2^48.
const BITS_PAST: f64 = 281_474_976_710_656.0;

/// A number as `BITAND`, `BITOR` and `BITXOR` take it (KERNEL.71): whole and
/// from 0 to 2^48-1, exactly, or `#NUM!`, so that `0.1*3*10`, which a double
/// holds a little above 3, is refused.
fn whole_bits(n: f64) -> Result<u64, ErrorKind> {
    if (0.0..BITS_PAST).contains(&n) && n.fract() == 0.0 {
        Ok(n as u64)
    } else {
        Err(ErrorKind::Num)
    }
}

/// `BITAND`, `BITOR` and `BITXOR` (KERNEL.71), one meaning under `op`, as
/// Excel's save of `scripts/recalc/bits.txt` holds it: both arguments read
/// first, left to right, as an operand reads, an error or a text spelling no
/// number answering there; only then each held to [`whole_bits`]. An answer is
/// below 2^48, which a double holds exactly.
fn bitwise(ctx: &Ctx<'_>, args: &[Arg], op: fn(u64, u64) -> u64) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let first = match number_arg(ctx, args, 0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let second = match number_arg(ctx, args, 1)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    Ok(match (whole_bits(first), whole_bits(second)) {
        (Ok(a), Ok(b)) => Value::Number(op(a, b) as f64),
        (Err(k), _) | (_, Err(k)) => Value::Error(k),
    })
}

/// `BITAND` ([`bitwise`]).
fn bitand(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    bitwise(ctx, args, |a, b| a & b)
}

/// `BITOR` ([`bitwise`]).
fn bitor(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    bitwise(ctx, args, |a, b| a | b)
}

/// `BITXOR` ([`bitwise`]).
fn bitxor(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    bitwise(ctx, args, |a, b| a ^ b)
}

/// The row of the first cell of a reference, or of the cell being computed.
fn row(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, Axis::Row)
}

/// The column of the first cell of a reference, or of the cell being computed.
fn column(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, Axis::Column)
}

/// [`placed`] over a call's arguments: none, the cell being computed; a
/// reference, its first area's first cell; anything else, no cell.
fn position(ctx: &Ctx<'_>, args: &[Arg], axis: Axis) -> Result<Value, Reason> {
    let (_, row, column) = ctx.env.here();
    Ok(match args {
        [] => placed(axis, Some((row, column))),
        [Arg::Area { areas, .. }] => match areas.first() {
            Some(a) => placed(axis, Some((a.top, a.left))),
            None => Value::Error(ErrorKind::Ref),
        },
        _ => placed(axis, None),
    })
}

/// `SINGLE(value)`, Excel's `@` as its file spells it (`_xlfn.SINGLE`, the
/// prefix dropped by the reader, KERNEL.8's second slice): a reference of
/// one area intersected with the cell being computed, as a legacy formula
/// reads a range where one value is wanted (`eval::intersection`), and
/// answered as a reference to that one cell, in a formula of either kind;
/// `#VALUE!` where the area holds no such cell, `#REF!` for a span of
/// sheets, as a legacy formula meets one; an array its first value, as
/// Microsoft's page on `@` has it ("If the value is an array, then pick the
/// top-left value", KERNEL.8, slice f); any other value itself.
fn single<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    let [e] = args else {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    };
    Ok(match ctx.eval(e)? {
        Arg::Area {
            areas,
            written,
            moved,
        } => match &areas[..] {
            [a] => match intersection(a, ctx.env.here()) {
                Some((sheet, row, col)) => Arg::Area {
                    areas: Areas::one(super::graph::Area {
                        sheet,
                        top: row,
                        left: col,
                        bottom: row,
                        right: col,
                    }),
                    written,
                    moved,
                },
                None => Arg::Scalar(Value::Error(ErrorKind::Value)),
            },
            _ => Arg::Scalar(Value::Error(ErrorKind::Ref)),
        },
        Arg::Array(a) => Arg::Scalar(a.get(0, 0).clone()),
        value => value,
    })
}

/// `CHOOSE(k, v1, v2, ...)`: the k-th value ([`chosen`]). Only the value
/// chosen is computed, and a reference chosen stays a reference. An
/// evaluation of many cells at once computes each value at the cells that
/// choose it (`super::vector`, its `choose`, KERNEL.28): a change here is a
/// change there, and the machine's tests hold the two to each other. An
/// index that is an array computes every value and picks at each place of
/// the lift the one its index chooses there, `#N/A` where that one falls
/// short, as `IF` picks its branch (KERNEL.8, slice f); the index keeps
/// its class in a legacy formula ([`kept_class`]).
fn choose<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    let index = match kept_class(ctx, &args[0])? {
        Lifted::One(v) => v,
        Lifted::Many(indices) => {
            let mut operands = vec![Lifted::Many(indices)];
            for a in &args[1..] {
                operands.push(ctx.operand(a)?);
            }
            let held: Vec<&Lifted> = operands.iter().collect();
            let count = args.len() - 1;
            let na = || Value::Error(ErrorKind::NA);
            let lifted = array::lift_by(&held, |r, c| {
                Ok(match held[0].at(r, c).map(|i| chosen(i, count)) {
                    None => na(),
                    Some(Ok(k)) => held[k].at(r, c).cloned().unwrap_or_else(na),
                    Some(Err(e)) => Value::Error(e),
                })
            })?;
            return Ok(Arg::Array(lifted));
        }
    };
    match chosen(&index, args.len() - 1) {
        Ok(k) => ctx.eval(&args[k]),
        Err(e) => Ok(Arg::Scalar(Value::Error(e))),
    }
}

/// `IFS(c1, v1, c2, v2, ...)`: the value beside the first condition that
/// holds; a condition's error is the answer; none holding is `#N/A`. Only the
/// value chosen is computed.
fn ifs<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 || args.len() % 2 == 1 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::NA)));
    }
    for pair in args.chunks(2) {
        match ctx.truth(&pair[0])? {
            Ok(true) => return ctx.eval(&pair[1]),
            Ok(false) => {}
            Err(k) => return Ok(Arg::Scalar(Value::Error(k))),
        }
    }
    Ok(Arg::Scalar(Value::Error(ErrorKind::NA)))
}

/// The powers of ten [`round`] scales by, each exact in a double.
const TENS: [f64; 23] = [
    1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16,
    1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
];

/// `10^digits` for `ROUND`'s digits, at most 300: the nearest double to it, read
/// from its decimal past the exact ones, the same on every platform, where
/// `powi`'s rounding is the platform's to choose and a machine's frames are bit
/// for bit between a native run and a browser's (KERNEL.68).
fn ten_to(digits: f64) -> f64 {
    let d = digits.min(300.0) as usize;
    match TENS.get(d) {
        Some(&power) => power,
        None => format!("1e{d}").parse().unwrap_or(f64::INFINITY),
    }
}

/// `ROUND(x, digits)`: half away from zero, the digits read down to a whole
/// number, a negative count rounding to tens and hundreds. The scaled value is
/// first rounded to Excel's fifteen significant digits, so that `ROUND(2.675, 2)`
/// is 2.68 as Excel has it and not the 2.67 the double 267.49999999999997 would
/// give.
fn round(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let x = match number_arg(ctx, args, 0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let digits = match number_arg(ctx, args, 1)? {
        Ok(n) => n.trunc(),
        Err(k) => return Ok(Value::Error(k)),
    };
    let rounded = if digits >= 0.0 {
        let m = ten_to(digits);
        round15(x * m).round() / m
    } else {
        let p = ten_to(-digits);
        round15(x / p).round() * p
    };
    Ok(finite(rounded))
}

/// `SIN` ([`sin_of`]).
fn sin(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => sin_of(n),
        Err(k) => Value::Error(k),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::Calc;
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    /// A rule's items of a number and of a truth value are its `given` and
    /// its `ranged` of them, to the bit (KERNEL.43).
    fn items_agree<R: Rule>(same: impl Fn(R::Item, R::Item) -> bool) {
        let numbers = [
            0.0,
            -0.0,
            1.0,
            -1.5,
            0.25,
            1e300,
            -1e-300,
            f64::MIN_POSITIVE,
        ];
        for n in numbers {
            let given = R::given(&Value::Number(n)).expect("a number is an item");
            assert!(same(R::given_number(n), given), "given {n}");
            let Ranged::Counts(ranged) = R::ranged(&Value::Number(n)) else {
                panic!("a number in a range counts")
            };
            assert!(same(R::ranged_number(n), ranged), "ranged {n}");
        }
        for b in [true, false] {
            let given = R::given(&Value::Bool(b)).expect("a truth value is an item");
            assert!(same(R::given_truth(b), given), "given {b}");
        }
    }

    #[test]
    fn a_rule_s_items_of_numbers_and_truth_values_are_its_given_and_ranged() {
        items_agree::<OverNumbers>(|a, b| a.to_bits() == b.to_bits());
        items_agree::<OverTruths>(|a, b| a == b);
    }

    /// Each `(row, content)` of a sheet `F`'s column A, as `Calc` computes it and
    /// spelled, beside a sheet `T` of the values the bit functions read: 13, a
    /// text of digits, a text, a truth value, an empty cell, an error and a
    /// fraction in A1:A7, 1, 10 and 100 in C1:C3, and 1 and 4 around an empty cell
    /// in D1:D3.
    fn bits_computed(rows: Vec<(u32, Content)>) -> Vec<String> {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        let cells = [
            ("A1", Content::Number(13.0)),
            ("A2", Content::Text("5".to_string())),
            ("A3", Content::Text("x".to_string())),
            ("A4", Content::Bool(true)),
            ("A6", Content::Error("#DIV/0!".to_string())),
            ("A7", Content::Number(0.5)),
            ("C1", Content::Number(1.0)),
            ("C2", Content::Number(10.0)),
            ("C3", Content::Number(100.0)),
            ("D1", Content::Number(1.0)),
            ("D3", Content::Number(4.0)),
        ];
        for (addr, content) in cells {
            let r = parse_a1_range(addr).unwrap();
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        }
        let f = wb.ensure_sheet("F");
        let at: Vec<u32> = rows.iter().map(|(row, _)| *row).collect();
        for (row, content) in rows {
            wb.sheets[f].set(row, 1, Cell { content, style: 0 });
        }
        let lib = language();
        let calc = Calc::run(&wb, &lib);
        at.into_iter()
            .map(|row| match calc.computed((f, row, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            })
            .collect()
    }

    /// Each legacy formula of the pairs against the value it is to compute, from
    /// row 11 down, a row no range of `T` is met at.
    fn bits_hold(cases: &[(&str, &str)]) {
        let rows = cases
            .iter()
            .enumerate()
            .map(|(i, (f, _))| (i as u32 + 11, Content::Formula(f.to_string())))
            .collect();
        for ((formula, want), got) in cases.iter().zip(bits_computed(rows)) {
            assert_eq!(got, *want, "={formula}");
        }
    }

    const NUM: &str = "(error \"#NUM!\")";
    const VALUE: &str = "(error \"#VALUE!\")";
    const DIV0: &str = "(error \"#DIV/0!\")";
    const NA: &str = "(error \"#N/A\")";
    const REF: &str = "(error \"#REF!\")";

    #[test]
    fn a_number_is_whole_bits_exactly_from_0_to_2_to_the_48_less_1() {
        assert_eq!(whole_bits(0.0), Ok(0));
        assert_eq!(whole_bits(-0.0), Ok(0));
        assert_eq!(whole_bits(5.0), Ok(5));
        assert_eq!(BITS_PAST, 2f64.powi(48));
        assert_eq!(whole_bits(BITS_PAST - 1.0), Ok((1 << 48) - 1));
        for n in [
            BITS_PAST,
            BITS_PAST + 2.0,
            2f64.powi(53),
            1e300,
            f64::MAX,
            -1.0,
            -0.5,
            -1e-300,
            1e-300,
            0.5,
            1.5,
            1.0 + f64::EPSILON,
            1.0 - f64::EPSILON / 2.0,
            0.1 * 3.0 * 10.0,
            2f64.powi(47) + 0.5,
            BITS_PAST - 1.0 + 2f64.powi(-5),
            BITS_PAST - 2f64.powi(-5),
        ] {
            assert_eq!(whole_bits(n), Err(ErrorKind::Num), "{n}");
        }
    }

    /// `TRUE()` and `FALSE()` are the values their literals are (KERNEL.58's *Found*
    /// 14), as LibreOffice's every boolean is written; given a value, which Excel's
    /// entry refuses and a file may hold, `#VALUE!`, as `NA` is.
    #[test]
    fn true_and_false_are_the_values_their_literals_are() {
        bits_hold(&[
            ("TRUE()", "true"),
            ("FALSE()", "false"),
            ("true()", "true"),
            ("NOT(TRUE())", "false"),
            ("TRUE()+1", "2"),
            ("IF(FALSE(),1,2)", "2"),
            ("AND(TRUE(),FALSE())", "false"),
            ("OR(TRUE(),FALSE())", "true"),
            ("ISLOGICAL(TRUE())", "true"),
            ("TYPE(FALSE())", "4"),
            ("TRUE()=TRUE", "true"),
            ("FALSE()=TRUE", "false"),
            ("TRUE(1)", VALUE),
            ("FALSE(1)", VALUE),
        ]);
    }

    #[test]
    fn the_bit_functions_compute_their_pages_examples_and_every_bit_to_the_top() {
        bits_hold(&[
            ("BITAND(1,5)", "1"),
            ("BITAND(13,25)", "9"),
            ("BITOR(23,10)", "31"),
            ("BITXOR(5,3)", "6"),
            ("bitxor(5,3)", "6"),
            ("BITAND(0,7)", "0"),
            ("BITOR(0,7)", "7"),
            ("BITXOR(7,0)", "7"),
            ("BITAND(13,13)", "13"),
            ("BITOR(13,13)", "13"),
            ("BITXOR(13,13)", "0"),
            ("BITOR(2^31,2^32)", "6442450944"),
            ("BITAND(2^32+1,2^31+1)", "1"),
            ("BITXOR(2^47,1)", "140737488355329"),
            ("BITAND(2^48-1,2^48-1)", "281474976710655"),
            ("BITXOR(2^48-1,2^47)", "140737488355327"),
            ("BITOR(2^47,2^46)", "211106232532992"),
            ("BITOR(1E+14,0)", "100000000000000"),
            ("BITOR(300%,0)", "3"),
            ("BITOR(-0,5)", "5"),
        ]);
    }

    #[test]
    fn the_bit_functions_hold_one_range_at_each_argument() {
        for name in ["BITAND", "BITOR", "BITXOR"] {
            for x in [
                "-1",
                "-0.5",
                "0.5",
                "1.9",
                "2^48",
                "2^53",
                "1E+300",
                "1+2^-52",
                "1-2^-53",
                "0.1*3*10",
                "2^47+0.5",
                "2^48-1+2^-5",
            ] {
                let first = format!("{name}({x},1)");
                let second = format!("{name}(1,{x})");
                bits_hold(&[(&first, NUM), (&second, NUM)]);
            }
            let (first, second) = (format!("{name}(2^48-1,0)"), format!("{name}(0,2^48-1)"));
            let top = if name == "BITAND" {
                "0"
            } else {
                "281474976710655"
            };
            bits_hold(&[(&first, top), (&second, top)]);
        }
    }

    #[test]
    fn the_bit_functions_read_both_arguments_before_they_hold_a_range() {
        for name in ["BITAND", "BITOR", "BITXOR"] {
            for (a, b, want) in [
                ("-1", "\"x\"", VALUE),
                ("\"x\"", "-1", VALUE),
                ("0.5", "\"x\"", VALUE),
                ("\"x\"", "2^48", VALUE),
                ("-1", "1/0", DIV0),
                ("1/0", "-1", DIV0),
                ("2^48", "T!A6", DIV0),
                ("T!A6", "0.5", DIV0),
                ("1/0", "\"x\"", DIV0),
                ("\"x\"", "1/0", VALUE),
                ("1/0", "#N/A", DIV0),
                ("#N/A", "1/0", NA),
            ] {
                let formula = format!("{name}({a},{b})");
                bits_hold(&[(&formula, want)]);
            }
        }
    }

    #[test]
    fn the_bit_functions_read_each_kind_as_an_operand_reads_it() {
        bits_hold(&[
            ("BITOR(TRUE,0)", "1"),
            ("BITOR(FALSE,0)", "0"),
            ("BITOR(\"3\",0)", "3"),
            ("BITOR(\" 3 \",0)", "3"),
            ("BITOR(\"1E2\",0)", "100"),
            ("BITOR(\"200%\",0)", "2"),
            ("BITOR(\"x\",0)", VALUE),
            ("BITOR(\"\",0)", VALUE),
            ("BITOR(\"TRUE\",0)", VALUE),
            ("BITOR(\"-1\",0)", NUM),
            ("BITOR(\"1.5\",0)", NUM),
            ("BITOR(\"281474976710655\",0)", "281474976710655"),
            ("BITOR(\"281474976710656\",0)", NUM),
            ("BITOR(#N/A,0)", NA),
            ("BITOR(T!A1,0)", "13"),
            ("BITOR(T!A2,0)", "5"),
            ("BITOR(T!A3,0)", VALUE),
            ("BITOR(T!A4,0)", "1"),
            ("BITOR(T!A5,5)", "5"),
            ("BITOR(5,T!A5)", "5"),
            ("BITOR(T!A6,0)", DIV0),
            ("BITOR(T!A7,0)", NUM),
            ("BITOR(T:F!A1,0)", REF),
            ("BITAND(1,T:F!A1)", REF),
        ]);
    }

    #[test]
    fn an_argument_left_out_of_a_bit_function_reads_0_however_it_is_left_out() {
        bits_hold(&[
            ("BITOR(,5)", "5"),
            ("BITOR(5,)", "5"),
            ("BITOR(,)", "0"),
            ("BITAND(,5)", "0"),
            ("BITAND(5,)", "0"),
            ("BITXOR(,5)", "5"),
            ("BITXOR(5,)", "5"),
            ("BITOR(CHOOSE(1,,2),5)", "5"),
            ("BITOR(5,CHOOSE(2,1,))", "5"),
            ("BITOR(IF(TRUE,,2),5)", "5"),
            ("BITOR(IF(FALSE,2,),5)", "5"),
            ("BITOR(IF(FALSE,2),5)", "5"),
        ]);
    }

    #[test]
    fn a_bit_function_of_any_count_but_two_is_value() {
        for name in ["BITAND", "BITOR", "BITXOR"] {
            for args in ["", "1", "1,2,3", "{1,2}", "{1,2},3,4"] {
                let formula = format!("{name}({args})");
                bits_hold(&[(&formula, VALUE)]);
            }
        }
    }

    #[test]
    fn the_bit_functions_lift_over_an_array_and_a_range_read_as_arrays() {
        bits_hold(&[
            ("BITAND({3,2,1},1)", "1"),
            ("SUM(BITAND({1,2,3},1))", "2"),
            ("SUM(BITOR({1,2,3},{4,8,16}))", "34"),
            ("SUM(BITAND({1,-1,3},1))", NUM),
            ("SUMPRODUCT(BITAND(T!C1:C3,1))", "1"),
            ("SUMPRODUCT(BITOR(1,T!C1:C3))", "113"),
            ("SUMPRODUCT(BITXOR(T!C1:C3,1))", "112"),
            ("SUMPRODUCT(BITAND(T!C1:C3,T!D1:D3))", "5"),
            ("SUMPRODUCT(BITOR(T!D1:D3,5))", "15"),
            ("SUMPRODUCT(BITOR(5,T!D1:D3))", "15"),
            ("SUMPRODUCT(BITAND(ROW(T!C1:C3),1))", "2"),
            ("SUMPRODUCT(--ISERROR(BITAND({1,-1,3},1)))", "1"),
            ("SUMPRODUCT(BITAND(T!C1:C3,T!C1:C2))", NA),
            ("ROWS(BITOR({1,2,3},{1;2}))", "2"),
            ("COLUMNS(BITOR({1,2,3},{1;2}))", "3"),
            ("ROWS(BITAND(T!C1:C3,1))", "3"),
        ]);
        // A dynamic formula reads a range whole where a legacy one meets it.
        let rows = vec![
            (11, Content::DynamicFormula("BITOR(T!C1:C3,1)".to_string())),
            (
                12,
                Content::DynamicFormula("INDEX(BITOR(T!C1:C3,1),3)".to_string()),
            ),
            (
                13,
                Content::DynamicFormula("INDEX(BITOR({1,2},{1,2,3}),1,3)".to_string()),
            ),
            (
                14,
                Content::DynamicFormula("INDEX(BITOR(T!D1:D3,5),2)".to_string()),
            ),
        ];
        assert_eq!(
            bits_computed(rows),
            [
                // Its spill blocked by the formula below it (KERNEL.55).
                "(error \"#SPILL!\")".to_string(),
                "101".to_string(),
                NA.to_string(),
                "5".to_string(),
            ]
        );
    }

    #[test]
    fn a_legacy_bit_function_meets_a_range_at_its_row_and_off_it_is_value() {
        let legacy = |text: &str| Content::Formula(text.to_string());
        let rows = vec![
            (1, legacy("BITAND(T!C1:C3,3)")),
            (2, legacy("BITAND(T!C1:C3,3)")),
            (3, legacy("BITAND(T!C1:C3,3)")),
            (4, legacy("BITAND(T!C1:C3,3)")),
            (5, legacy("SUM(BITAND(T!C1:C3,3))")),
        ];
        assert_eq!(bits_computed(rows), ["1", "2", "0", VALUE, VALUE]);
    }

    #[test]
    fn a_mask_read_by_the_bit_functions_is_the_mask_read_by_mod_and_int() {
        // What Seiyaku's source writes at every site (Alonzo's `CART.14.1`), for
        // every mask of nine marks at every mark.
        let mut cases = Vec::new();
        for mask in [0, 1, 2, 3, 170, 256, 341, 510, 511] {
            for k in 0..9 {
                let bit = 1 << k;
                let set = if mask & bit != 0 { "true" } else { "false" };
                cases.push((format!("BITAND({mask},{bit})>0"), set));
                cases.push((format!("MOD(INT({mask}/{bit}),2)=1"), set));
            }
        }
        let cases: Vec<(&str, &str)> = cases.iter().map(|(f, w)| (f.as_str(), *w)).collect();
        bits_hold(&cases);
    }

    #[test]
    fn round_scales_by_the_same_power_on_every_platform() {
        // KERNEL.68: exact to 10^22, the nearest double to its decimal past it,
        // and no power past the 300 digits a cell's number can hold.
        for d in 0..=22 {
            assert_eq!(ten_to(f64::from(d)), 10f64.powi(d), "10^{d}");
        }
        assert_eq!(ten_to(23.0), 1e23);
        assert_eq!(ten_to(300.0), 1e300);
        assert_eq!(ten_to(1000.0), 1e300);
    }

    #[test]
    fn the_language_declares_the_functions_that_read_an_extent_and_no_other() {
        // KERNEL.56: the lookups that bisect and the database functions are
        // declared to read a sheet's extent, each a function the language
        // holds, without case; no other is; a host's own function is
        // declared once registered, and registered again it declares what it
        // reads of its own.
        let lib = language();
        for name in lib.names() {
            assert_eq!(
                lib.reads_extent(name),
                READS_EXTENT.contains(&name),
                "{name}"
            );
        }
        assert!(READS_EXTENT.iter().all(|name| lib.has(name)));
        assert!(lib.reads_extent("dcounta") && lib.reads_extent("XLookup"));
        let mut lib = language();
        lib.register("LAST", |_, _| Ok(Value::Empty));
        assert!(!lib.reads_extent("LAST"));
        lib.declare_reads_extent("last");
        assert!(lib.reads_extent("LAST"));
        lib.register("LAST", |_, _| Ok(Value::Empty));
        assert!(!lib.reads_extent("LAST"), "registered again");
        lib.register("MATCH", |_, _| Ok(Value::Empty));
        assert!(!lib.reads_extent("MATCH"), "the language's replaced");
    }

    #[test]
    fn the_language_registers_the_day_one_list_and_kernel_8_s_53_s_and_60_s_exactly() {
        let lib = language();
        let mut want: Vec<&str> = NAMES
            .iter()
            .chain(KERNEL_8.iter())
            .chain(KERNEL_53.iter())
            .chain(KERNEL_60.iter())
            .chain(KERNEL_68.iter())
            .chain(KERNEL_71.iter())
            .chain(KERNEL_58.iter())
            .chain(KERNEL_29.iter())
            .chain(KERNEL_55.iter())
            .copied()
            .collect();
        want.sort_unstable();
        assert_eq!(lib.names(), want);
        for name in KERNEL_53 {
            assert!(
                lib.arrayed(name).is_some(),
                "{name} answers an operand, so that every path computes it as Ctx does"
            );
        }
        for name in ["INDEX", "XLOOKUP", "IFERROR", "IFNA"] {
            assert!(
                lib.form(name).is_some(),
                "{name} computes only what it answers"
            );
        }
        assert!(
            lib.form("IF").is_some(),
            "IF computes only the branch it takes"
        );
        assert!(
            lib.form("CHOOSE").is_some(),
            "CHOOSE computes only the value chosen"
        );
        assert!(lib.strict("SUM").is_some());
    }
}
