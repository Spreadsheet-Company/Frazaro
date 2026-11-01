//! Excel's library, registered through the language's functions seam
//! (KERNEL.7, 2026-10-08): the declared subset `frazaro calc` computes, each
//! function with the evidence that admitted it.
//!
//! The rule is `KERNEL.6`'s: a function enters the subset with its fixture
//! and its measured frequency, never by taste. The measure is the one the
//! roadmap cites, the two spreadsheet corpora of Jansen and Hermans, "Enron
//! versus EUSES: A Comparison of Two Spreadsheet Corpora" (2015, arXiv
//! 1503.04055), Table IV: the fifteen most used functions in each corpus,
//! which cover 69% of the Enron corpus's spreadsheets and 70% of EUSES's.
//! Of the Enron fifteen (SUM, IF, AVERAGE, VLOOKUP, ROUND, SUBTOTAL, OFFSET,
//! CONCATENATE, NOW, DAVERAGE, SUMIF, INDEX, MATCH, LOOKUP, MONTH) `KERNEL.7`
//! computed the five that are arithmetic and logic over values, and
//! `KERNEL.8`'s first slice the four lookups (2026-10-18), its third
//! `CONCATENATE` with the text functions, its fourth `MONTH` with the
//! dates and its fifth `SUBTOTAL` and `DAVERAGE` with the aggregates and
//! the database functions (2026-10-20); `OFFSET` is unreadable by the
//! reader's own rule, and `NOW` is volatile, which an engine refuses and a
//! door shows as not computed. Of the EUSES
//! fifteen (SUM, IF, ROUND, HYPERLINK, CONCATENATE, AND, COUNTIF, AVERAGE,
//! OR, INDIRECT, MIN, ISNUMBER, MAX, VLOOKUP, ISBLANK) it computes ten. `IFS` enters on its fixture alone,
//! `scripts/reflect/saved.xlsx`, the owner's first Excel-saved workbook,
//! whose cached values are this oracle's first row. The language's own
//! functions (`vla_lang::calc::library`) are here as they are there: its
//! day-one fourteen, which `Alonzo/SPEC.md` section 13 named for the
//! engine's first cartridges (`IF`, `SUM`, `AND`, `OR`, `MIN` and `MAX` in
//! the tables too), and the fourteen `KERNEL.8` brought, the lookups, the
//! error tests a lookup's miss is read by, and `SUMIF` and `COUNTIF`, moved
//! down from this file with their criteria since a vector's index keys by
//! them and a cartridge counts duplicates (Alonzo's `CART.13`),
//! `SINGLE`, Excel's `@`, with implicit intersection (its second slice),
//! the twenty-five text functions with `TEXT` (its third), the
//! seventeen date functions (its fourth), and the thirty-one aggregates,
//! `SUBTOTAL`, `SUMIFS` and its family and the database functions (its
//! fifth), `AVERAGE` moved down from this file with them, and the
//! twenty-one array functions (its sixth); `ISFORMULA` and
//! `FORMULATEXT`, a cell's formula read by its reference (`KERNEL.53`),
//! which enter on their fixture alone as `IFS` did, Excel having added them
//! in 2013, after both corpora were gathered; `ADDRESS` (`KERNEL.60`) and
//! `BITAND`, `BITOR` and `BITXOR` (`KERNEL.71`), on their fixtures alone too, in
//! neither table, for what a game asked of them; and the
//! information functions (`KERNEL.68`), with `ROUND`, `IFS`, `ISNUMBER` and
//! `ISBLANK` moved down from this file, so that a door computes nothing an engine
//! refuses and this file's library is the language's, the seam kept for what comes
//! by measurement.
//!
//! Every function here is held to the language's exact tests or to its
//! fixture: `scripts/recalc/subset.txt`, built by the door into
//! `subset.xlsx` with one formula per function and per coercion this slice
//! claims. Its cached values are Excel's the day the owner opens it in
//! Excel 365 and saves it beside as `subset_saved.xlsx`, and
//! `tools/check_recalc_golden.ps1` then holds every row of that file to
//! `agree`; until then the door prints `unchecked` beside each, which is
//! the honest word. A function outside this list shows `not computed here`
//! with its name (`HORIZON.md` section 12.6), and the one line that admits
//! it is a row in [`SUBSET`] with its evidence.

use vla_lang::calc::library::language;
use vla_lang::calc::Library;

/// One function of the subset and why it is in.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Entry {
    pub name: &'static str,
    /// The measurement or the fixture that admitted it.
    pub evidence: &'static str,
}

/// The declared subset: the language's functions, its day-one fourteen first,
/// then the hundred and eight `KERNEL.8` brought, the two `KERNEL.53` brought, the
/// one `KERNEL.60` brought, the fourteen of `KERNEL.68` and the three of
/// `KERNEL.71`, each keeping the
/// evidence that admitted it.
pub const SUBSET: &[Entry] = &[
    Entry {
        name: "IF",
        evidence: "Enron 2, EUSES 2; SPEC.md s13",
    },
    Entry {
        name: "AND",
        evidence: "EUSES 6; SPEC.md s13",
    },
    Entry {
        name: "OR",
        evidence: "EUSES 9; SPEC.md s13",
    },
    Entry {
        name: "NOT",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "SUM",
        evidence: "Enron 1, EUSES 1; SPEC.md s13",
    },
    Entry {
        name: "MIN",
        evidence: "EUSES 11; SPEC.md s13",
    },
    Entry {
        name: "MAX",
        evidence: "EUSES 13; SPEC.md s13",
    },
    Entry {
        name: "ABS",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "INT",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "MOD",
        evidence: "SPEC.md s13: every counter wraps",
    },
    Entry {
        name: "ROW",
        evidence: "SPEC.md s13: a cell painting at the mouse",
    },
    Entry {
        name: "COLUMN",
        evidence: "SPEC.md s13: a cell painting at the mouse",
    },
    Entry {
        name: "CHOOSE",
        evidence: "SPEC.md s13: a score until INDEX comes",
    },
    Entry {
        name: "SIN",
        evidence: "SPEC.md s13: a slow oscillator",
    },
    Entry {
        name: "INDEX",
        evidence: "Enron 12; Alonzo's KIT.3 and CART.5, Buffet's MISE.6; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "MATCH",
        evidence: "Enron 13; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "XMATCH",
        evidence: "Excel 365's spelling of MATCH's search; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "VLOOKUP",
        evidence: "Enron 4, EUSES 14; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "HLOOKUP",
        evidence: "VLOOKUP's search across a row; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "LOOKUP",
        evidence: "Enron 14; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "XLOOKUP",
        evidence: "Excel 365's spelling of the lookups; Alonzo's ALGEBRA.md s12, the word game; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "IFERROR",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "IFNA",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "ISERROR",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "ISNA",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "COUNTIF",
        evidence: "EUSES 7; moved into the language, KERNEL.8",
    },
    Entry {
        name: "SUMIF",
        evidence: "Enron 11; moved into the language, KERNEL.8",
    },
    Entry {
        name: "SINGLE",
        evidence: "Excel's @ as its file spells it, with implicit intersection; scripts/recalc/intersect.txt",
    },
    Entry {
        name: "CONCATENATE",
        evidence: "Enron 8, EUSES 5; the text functions, KERNEL.8; scripts/recalc/texts.txt",
    },
    Entry {
        name: "CONCAT",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "TEXTJOIN",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "LEFT",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "RIGHT",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "MID",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "LEN",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "FIND",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "SEARCH",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "SUBSTITUTE",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "REPLACE",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "UPPER",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "LOWER",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "PROPER",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "TRIM",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "REPT",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "EXACT",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "VALUE",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "T",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "N",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "UNICODE",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "UNICHAR",
        evidence: "the text functions, KERNEL.8, Alonzo's CART.13 among their readers; scripts/recalc/texts.txt",
    },
    Entry {
        name: "CHAR",
        evidence: "over ASCII, another code page's labelled (the owner's decision), KERNEL.8; scripts/recalc/texts.txt",
    },
    Entry {
        name: "CODE",
        evidence: "over ASCII, another code page's labelled (the owner's decision), KERNEL.8; scripts/recalc/texts.txt",
    },
    Entry {
        name: "TEXT",
        evidence: "a number format's code drawn by numfmt, a date's among them, KERNEL.8; scripts/recalc/texts.txt",
    },
    Entry {
        name: "DATE",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "YEAR",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "MONTH",
        evidence: "Enron 15; the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "DAY",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "WEEKDAY",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "EOMONTH",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "EDATE",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "DAYS",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "NETWORKDAYS",
        evidence: "the dates, KERNEL.8, a game's holidays among them (Buffet's Caveats); scripts/recalc/dates.txt",
    },
    Entry {
        name: "WORKDAY",
        evidence: "the dates, KERNEL.8, a game's holidays among them (Buffet's Caveats); scripts/recalc/dates.txt",
    },
    Entry {
        name: "TIME",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "HOUR",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "MINUTE",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "SECOND",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "DATEDIF",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "YEARFRAC",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "WEEKNUM",
        evidence: "the dates, KERNEL.8; scripts/recalc/dates.txt",
    },
    Entry {
        name: "AVERAGE",
        evidence: "Enron 3, EUSES 8; moved down into the language with SUBTOTAL, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "COUNT",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "COUNTA",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "PRODUCT",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "STDEV",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "STDEV.S",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "STDEVP",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "STDEV.P",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "VAR",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "VAR.S",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "VARP",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "VAR.P",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "SUBTOTAL",
        evidence: "Enron 6; the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "SUMIFS",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "COUNTIFS",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "AVERAGEIF",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "AVERAGEIFS",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "MAXIFS",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "MINIFS",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DSUM",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DAVERAGE",
        evidence: "Enron 10; the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DCOUNT",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DCOUNTA",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DGET",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DMAX",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DMIN",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DPRODUCT",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DSTDEV",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DSTDEVP",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DVAR",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "DVARP",
        evidence: "the aggregates by criteria, KERNEL.8; scripts/recalc/criteria.txt",
    },
    Entry {
        name: "SEQUENCE",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "FILTER",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "SORT",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "SORTBY",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "UNIQUE",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "TRANSPOSE",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "SUMPRODUCT",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "MMULT",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "ROWS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "COLUMNS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "CHOOSECOLS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "CHOOSEROWS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "TAKE",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "DROP",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "HSTACK",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "VSTACK",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "TOCOL",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "TOROW",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "WRAPROWS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "WRAPCOLS",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "EXPAND",
        evidence: "the arrays within a formula, KERNEL.8; scripts/recalc/arrays.txt",
    },
    Entry {
        name: "ISFORMULA",
        evidence: "a cell's formula read by its reference, KERNEL.53 for Alonzo's CART.9; scripts/recalc/formulas.txt",
    },
    Entry {
        name: "FORMULATEXT",
        evidence: "a cell's formula read by its reference, KERNEL.53 for Alonzo's CART.9; scripts/recalc/formulas.txt",
    },
    Entry {
        name: "ADDRESS",
        evidence: "a cell's address from its row and column, KERNEL.60 for Alonzo's KIT.6 and CART.14.1; scripts/recalc/address.txt",
    },
    Entry {
        name: "BITAND",
        evidence: "a mask read by one call, KERNEL.71 for Alonzo's CART.14.1; scripts/recalc/bits.txt",
    },
    Entry {
        name: "BITOR",
        evidence: "a mask read by one call, KERNEL.71 for Alonzo's CART.14.1; scripts/recalc/bits.txt",
    },
    Entry {
        name: "BITXOR",
        evidence: "a mask read by one call, KERNEL.71 for Alonzo's CART.14.1; scripts/recalc/bits.txt",
    },
    Entry {
        name: "ISTEXT",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISNONTEXT",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISLOGICAL",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISERR",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISREF",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISEVEN",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ISODD",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ERROR.TYPE",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "TYPE",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "NA",
        evidence: "the information functions, KERNEL.68 for Alonzo's CART.9; scripts/recalc/info.txt",
    },
    Entry {
        name: "ROUND",
        evidence: "Enron 5, EUSES 3",
    },
    Entry {
        name: "ISNUMBER",
        evidence: "EUSES 12",
    },
    Entry {
        name: "ISBLANK",
        evidence: "EUSES 15",
    },
    Entry {
        name: "IFS",
        evidence: "scripts/reflect/saved.xlsx, the first Excel-saved fixture",
    },
];

/// The functions named by the two corpora's top fifteen that the subset
/// does not compute, each with the slice that takes it: what the honest
/// label names most often on a real workbook.
pub const NAMED_NOT_COMPUTED: &[(&str, &str)] = &[
    (
        "DATEVALUE",
        "a date spelled in a text, which a host reads by its locale",
    ),
    (
        "TIMEVALUE",
        "a time spelled in a text, which a host reads by its locale",
    ),
    ("NOW", "volatile: an engine refuses it, a door shows it"),
    ("TODAY", "volatile: an engine refuses it, a door shows it"),
    ("OFFSET", "unreadable by the reader's rule"),
    ("INDIRECT", "unreadable by the reader's rule"),
    (
        "HYPERLINK",
        "a link, not a value; the formula sink refuses it",
    ),
];

/// Excel's library: the language's functions, none of its own since
/// `KERNEL.68` moved its last four down; the one line that admits another is a
/// registration here and a row in [`SUBSET`] with its evidence.
pub fn library() -> Library {
    language()
}

/// The names the subset holds, for a page to list: the language's first.
pub fn subset_names() -> Vec<&'static str> {
    SUBSET.iter().map(|e| e.name).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use vla_lang::calc::library::NAMES as LANGUAGE_NAMES;
    use vla_lang::calc::{Calc, Computed, ErrorKind, Value, Values};
    use vla_lang::sheet::{parse_a1_range, Cell, Content, Workbook};

    fn model(rows: &[(&str, &str, &str)]) -> Workbook {
        let mut wb = Workbook::new();
        for (sheet, addr, text) in rows {
            let i = wb.ensure_sheet(sheet);
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(f) = text.strip_prefix('=') {
                Content::Formula(f.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[i].set(r.top, r.left, Cell { content, style: 0 });
        }
        wb
    }

    fn at(calc: &Calc<'_>, addr: &str) -> Computed {
        let r = parse_a1_range(addr).unwrap();
        calc.computed((0, r.top, r.left)).cloned().unwrap()
    }

    fn number(calc: &Calc<'_>, addr: &str) -> f64 {
        match at(calc, addr) {
            Computed::Value(Value::Number(n)) => n,
            other => panic!("{addr}: {other:?}"),
        }
    }

    #[test]
    fn the_subset_table_is_the_library_and_the_language_comes_first() {
        let lib = library();
        let mut names: Vec<&str> = SUBSET.iter().map(|e| e.name).collect();
        names.sort_unstable();
        assert_eq!(
            lib.names(),
            names,
            "every function admitted has its evidence row, and no other"
        );
        assert_eq!(
            &SUBSET[..LANGUAGE_NAMES.len()]
                .iter()
                .map(|e| e.name)
                .collect::<Vec<_>>(),
            &LANGUAGE_NAMES.to_vec()
        );
        for e in SUBSET {
            assert!(!e.evidence.is_empty(), "{} has no evidence", e.name);
        }
        for (name, _) in NAMED_NOT_COMPUTED {
            assert!(!lib.has(name), "{name} is listed as not computed");
        }
    }

    /// The typed formulas' fixture as Excel 365 saved it, each written through
    /// `Range.Formula2` as a person's typing is (KERNEL.70).
    const TYPED_SAVED: &str = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../scripts/recalc/typed_saved.xlsx"
    );

    #[test]
    fn a_row_s_word_computes_what_excel_365_answers_the_same_text_typed() {
        // KERNEL.70: each of the 56 written by a row with the word into a machine
        // under Excel's library, at the cell Excel's save holds it in, and its
        // frame's value held to the value Excel saved there; the nine of many
        // values labelled, where Excel spilled them; and a typed `@`, which Excel
        // holds as a legacy formula, refused by name. A package holds no fixture,
        // and the test stands down.
        use crate::reflect::calc::collect;
        use vla_lang::machine::Machine;
        use vla_lang::rows::{datum, quoted};
        use vla_lang::sheet::cell_ref;
        let Ok(bytes) = std::fs::read(TYPED_SAVED) else {
            return;
        };
        let read = collect(&bytes, "typed_saved.xlsx").expect("the save");
        let (data, typed) = (
            read.model.find_sheet("data").expect("data"),
            read.model.find_sheet("typed").expect("typed"),
        );
        let mut rows = String::from("(sheet \"data\" visible)\n(sheet \"typed\" visible)\n");
        for (&(row, col), cell) in &read.model.sheets[data].cells {
            let addr = cell_ref(row, col);
            rows.push_str(&match &cell.content {
                Content::Number(n) => format!("(cell \"data\" \"{addr}\" {n})\n"),
                Content::Text(t) => format!("(cell \"data\" \"{addr}\" {})\n", quoted(t)),
                Content::Formula(f) => {
                    format!(
                        "(formula \"data\" \"{addr}\" {})\n",
                        quoted(&format!("={f}"))
                    )
                }
                other => panic!("{addr}: {other:?}"),
            });
        }
        let (mut written, mut refused) = (Vec::new(), 0);
        for (&(row, col), cell) in &read.model.sheets[typed].cells {
            if col != 2 {
                continue;
            }
            // A spill's other cells hold its values, and no formula.
            let (Content::Formula(f) | Content::DynamicFormula(f)) = &cell.content else {
                continue;
            };
            let as_typed = match read.model.sheets[typed].cells.get(&(row, 1)) {
                Some(Cell {
                    content: Content::Text(t),
                    ..
                }) => t.clone(),
                other => panic!("A{row}: {other:?}"),
            };
            // An `@` as typed, refused by name by the parser, and left out.
            if as_typed.contains('@') {
                let one = format!(
                    "(sheet \"data\" visible)\n(sheet \"typed\" visible)\n(formula \"typed\" \"B{row}\" {} dynamic)\n",
                    quoted(&as_typed)
                );
                let e = Machine::load(&one, library()).expect_err(&as_typed);
                assert!(e.refusal.text.contains('@'), "{}", e.refusal.text);
                refused += 1;
                continue;
            }
            let addr = cell_ref(row, col);
            let text = quoted(&format!("={f}"));
            rows.push_str(&format!("(formula \"typed\" \"{addr}\" {text} dynamic)\n"));
            written.push(((typed, row, col), addr));
        }
        assert_eq!((written.len(), refused), (54, 2));
        let mut m = Machine::load(&rows, library()).unwrap_or_else(|e| {
            panic!("line {}: {}", e.line, e.refusal.text);
        });
        let typed_in_machine = m.find_sheet("typed").expect("typed");
        assert!(m.step(0).done);
        let mut spills = Vec::new();
        for (id, addr) in &written {
            let got = m
                .computed((typed_in_machine, id.1, id.2))
                .unwrap_or_else(|| panic!("{addr}: no value"))
                .spell();
            let want = datum(
                read.cached
                    .get(id)
                    .unwrap_or_else(|| panic!("{addr}: no cache")),
            );
            if got == "(not-computed \"a spill\")" {
                spills.push(addr.as_str());
                continue;
            }
            let same = match (got.parse::<f64>(), want.parse::<f64>()) {
                (Ok(a), Ok(b)) => a == b,
                _ => got == want,
            };
            assert!(same, "{addr}: {got} where Excel saved {want}");
        }
        assert_eq!(
            spills,
            ["B451", "B461", "B471", "B481", "B491", "B501", "B511", "B521", "B531"],
            "Excel spilled these nine"
        );
    }

    #[test]
    fn excel_s_functions_compute_as_excel_does() {
        let wb = model(&[
            ("S", "A1", "10"),
            ("S", "A2", "20"),
            ("S", "A3", "x"),
            ("S", "A4", "30"),
            ("S", "B1", "1"),
            ("S", "B2", "2"),
            ("S", "B4", "4"),
            ("S", "C1", "=AVERAGE(A1:A4)"),
            ("S", "C2", "=AVERAGE(A3)"),
            ("S", "C3", "=ROUND(2.675,2)"),
            ("S", "C4", "=ROUND(-2.5,0)"),
            ("S", "C5", "=ROUND(1234.5678,-2)"),
            ("S", "C6", "=ROUND(1.234,1.9)"),
            ("S", "C7", "=SUMIF(A1:A4,\">15\")"),
            ("S", "C8", "=SUMIF(A1:A4,\">15\",B1:B4)"),
            ("S", "C9", "=COUNTIF(A1:A4,\">15\")"),
            ("S", "C10", "=COUNTIF(A1:A4,\"x\")"),
            ("S", "C11", "=COUNTIF(A1:A6,\"\")"),
            ("S", "C12", "=COUNTIF(A1:A4,\"<>10\")"),
            ("S", "C13", "=COUNTIF(A1:A4,\"*\")"),
            ("S", "C14", "=COUNTIF(A1:A4,10)"),
            ("S", "C15", "=ISNUMBER(A1)"),
            ("S", "C16", "=ISNUMBER(A3)"),
            ("S", "C17", "=ISNUMBER(1/0)"),
            ("S", "C18", "=ISBLANK(A5)"),
            ("S", "C19", "=ISBLANK(A1)"),
            ("S", "C20", "=IFS(A1>15,\"big\",TRUE,\"small\")"),
            ("S", "C21", "=IFS(A1>15,\"big\")"),
            ("S", "C22", "=IFS(A3,1,TRUE,2)"),
            ("S", "C23", "=COUNTIF(B1:B4,\"<>\")"),
            ("S", "C24", "=COUNTIF(A1:A4,\"?\")"),
            ("S", "C25", "=SUMIF(A1:A4,\"x\",B1:B4)"),
        ]);
        let lib = library();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(number(&calc, "C1"), 20.0, "the text is skipped");
        assert_eq!(
            at(&calc, "C2"),
            Computed::Value(Value::Error(ErrorKind::Div0))
        );
        assert_eq!(number(&calc, "C3"), 2.68);
        assert_eq!(number(&calc, "C4"), -3.0);
        assert_eq!(number(&calc, "C5"), 1200.0);
        assert_eq!(number(&calc, "C6"), 1.2);
        assert_eq!(number(&calc, "C7"), 50.0);
        assert_eq!(number(&calc, "C8"), 6.0);
        assert_eq!(number(&calc, "C9"), 2.0);
        assert_eq!(number(&calc, "C10"), 1.0);
        assert_eq!(number(&calc, "C11"), 2.0, "A5 and A6 are blank");
        assert_eq!(number(&calc, "C12"), 3.0, "the text counts as not 10");
        assert_eq!(number(&calc, "C13"), 1.0, "* matches texts alone");
        assert_eq!(number(&calc, "C14"), 1.0);
        assert_eq!(at(&calc, "C15"), Computed::Value(Value::Bool(true)));
        assert_eq!(at(&calc, "C16"), Computed::Value(Value::Bool(false)));
        assert_eq!(
            at(&calc, "C17"),
            Computed::Value(Value::Bool(false)),
            "an error is not a number and does not propagate"
        );
        assert_eq!(at(&calc, "C18"), Computed::Value(Value::Bool(true)));
        assert_eq!(at(&calc, "C19"), Computed::Value(Value::Bool(false)));
        assert_eq!(
            at(&calc, "C20"),
            Computed::Value(Value::Text("small".to_string()))
        );
        assert_eq!(
            at(&calc, "C21"),
            Computed::Value(Value::Error(ErrorKind::NA))
        );
        assert_eq!(
            at(&calc, "C22"),
            Computed::Value(Value::Error(ErrorKind::Value)),
            "a text condition"
        );
        assert_eq!(number(&calc, "C23"), 3.0, "the non-blank cells");
        assert_eq!(number(&calc, "C24"), 1.0);
        assert_eq!(number(&calc, "C25"), 0.0, "B3 is absent");
    }
}
