//! The file format's names for Excel's newer functions (PORT.7, slice 7b).
//!
//! New ground, said plainly: the reference writes a formula through
//! `Range.Formula2`, and Excel's object model spells the file for it. In the
//! file, a function added after Excel 2007 is stored under a prefix,
//! `_xlfn.` (and `_xlfn._xlws.` for two of the dynamic-array functions), and
//! a formula written without it shows `#NAME?` until a person re-enters the
//! cell. Excel strips the prefix again when it shows the formula, so a user
//! never sees it; Sheets, Calc and Numbers strip it too. The table is the
//! "future functions" list of the format's documentation (MS-XLSX), as far
//! as this writer knows it; a function missing from it shows up as `#NAME?`
//! in a built workbook, and the fix is one more row here.
//!
//! What is not done here: the parameter names of `LET` and `LAMBDA`, which
//! Excel stores under `_xlpm.`; they are written as given, and the owner's
//! live pass says whether Excel accepts them (step 0, slice 7b).

pub use super::strip_future_prefixes;
use super::{is_name_char, is_name_start};

/// Functions stored as `_xlfn.<NAME>`.
const XLFN: &[&str] = &[
    // Excel 2010
    "AGGREGATE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.INV",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FLOOR.PRECISE",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "HYPGEOM.DIST",
    "ISO.CEILING",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MODE.MULT",
    "MODE.SNGL",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANK.AVG",
    "RANK.EQ",
    "STDEV.P",
    "STDEV.S",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "VAR.P",
    "VAR.S",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "Z.TEST",
    // Excel 2013
    "ACOT",
    "ACOTH",
    "ARABIC",
    "BASE",
    "BINOM.DIST.RANGE",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "CEILING.MATH",
    "COMBINA",
    "COT",
    "COTH",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "ENCODEURL",
    "FILTERXML",
    "FLOOR.MATH",
    "FORMULATEXT",
    "GAMMA",
    "GAUSS",
    "IFNA",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISOWEEKNUM",
    "MUNIT",
    "NUMBERVALUE",
    "PDURATION",
    "PERMUTATIONA",
    "PHI",
    "RRI",
    "SEC",
    "SECH",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "UNICHAR",
    "UNICODE",
    "WEBSERVICE",
    "XOR",
    // Excel 2016 and 2019
    "CONCAT",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "IFS",
    "MAXIFS",
    "MINIFS",
    "SWITCH",
    "TEXTJOIN",
    // Microsoft 365
    "ARRAYTOTEXT",
    "BYCOL",
    "BYROW",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DETECTLANGUAGE",
    "DROP",
    "EXPAND",
    "GROUPBY",
    "HSTACK",
    "IMAGE",
    "ISOMITTED",
    "LAMBDA",
    "LET",
    "MAKEARRAY",
    "MAP",
    "PERCENTOF",
    "PIVOTBY",
    "RANDARRAY",
    "REDUCE",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "SCAN",
    "SEQUENCE",
    "SORTBY",
    "STOCKHISTORY",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSLATE",
    "TRIMRANGE",
    "UNIQUE",
    "VALUETOTEXT",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
];

/// Functions stored as `_xlfn._xlws.<NAME>`.
const XLWS: &[&str] = &["FILTER", "SORT"];

/// Functions that can return a reference or an array, so that Excel 365
/// marks a legacy formula calling one with the implicit-intersection `@`
/// (`=@IFS(...)`, seen on the owner's first live pass of slice 7b) and
/// stores a `Formula2` entry of one as a dynamic-array formula. The writer
/// stores such a formula the same way (`sheet::ooxml`, the `cm` metadata),
/// so the formula bar reads as the add-in's would. The table is Excel's own
/// classification as far as this writer knows it: a function missing here
/// shows `@` in the bar and, filled over a range, computes one value per
/// cell; one wrongly here is a single-cell dynamic array, harmless in Excel
/// and a braced array in an older host. `IF`, `IFERROR` and the lookup
/// functions that return one value are not here, as old workbooks full of
/// them open without `@`.
const ARRAY_CAPABLE: &[&str] = &[
    "ANCHORARRAY",
    "BYCOL",
    "BYROW",
    "CHOOSE",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "FILTER",
    "FREQUENCY",
    "GROUPBY",
    "GROWTH",
    "HSTACK",
    "IFS",
    "INDEX",
    "INDIRECT",
    "LAMBDA",
    "LET",
    "LINEST",
    "LOGEST",
    "LOOKUP",
    "MAKEARRAY",
    "MAP",
    "MINVERSE",
    "MMULT",
    "MODE.MULT",
    "MUNIT",
    "OFFSET",
    "PIVOTBY",
    "RANDARRAY",
    "REDUCE",
    "REGEXEXTRACT",
    "SCAN",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "SWITCH",
    "TAKE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSPOSE",
    "TREND",
    "TRIMRANGE",
    "UNIQUE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
];

/// Whether a formula's text (without its `=`, prefixes applied or not)
/// calls a function that can return a reference or an array, and so must be
/// stored as a dynamic-array formula to read without `@`.
pub fn can_return_array(formula: &str) -> bool {
    let chars: Vec<char> = formula.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            i += 1;
            while i < chars.len() {
                if chars[i] == '"' {
                    if i + 1 < chars.len() && chars[i + 1] == '"' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if is_name_start(c) && (i == 0 || !is_name_char(chars[i - 1])) {
            let start = i;
            while i < chars.len() && is_name_char(chars[i]) {
                i += 1;
            }
            let mut j = i;
            while j < chars.len() && chars[j] == ' ' {
                j += 1;
            }
            if j < chars.len() && chars[j] == '(' {
                let name: String = chars[start..i]
                    .iter()
                    .collect::<String>()
                    .to_ascii_uppercase();
                let bare = name
                    .strip_prefix("_XLFN._XLWS.")
                    .or_else(|| name.strip_prefix("_XLFN."))
                    .unwrap_or(&name);
                if ARRAY_CAPABLE.contains(&bare) {
                    return true;
                }
            }
            continue;
        }
        i += 1;
    }
    false
}

/// The prefix a function name takes in the file, or `None` for one Excel
/// 2007 already had (or one this table does not know).
pub fn prefix_for(upper_name: &str) -> Option<&'static str> {
    if XLWS.contains(&upper_name) {
        Some("_xlfn._xlws.")
    } else if XLFN.contains(&upper_name) {
        Some("_xlfn.")
    } else {
        None
    }
}

/// A formula's text (without its leading `=`) with every call of a newer
/// function written under its prefix, in the upper case Excel stores. Text
/// inside a string literal is left alone; a name already prefixed is one
/// the table does not hold (`_xlfn.IFS` is not `IFS`), so it is left too.
/// A name is a call only when `(` follows it, so a defined name or a cell
/// reference spelled like a function is untouched.
pub fn prefix_future_functions(formula: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 16);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // A string literal, to its closing quote; a doubled quote stays
            // inside it.
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if i + 1 < chars.len() && chars[i + 1] == '"' {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if is_name_start(c) && (i == 0 || !is_name_char(chars[i - 1])) {
            let start = i;
            while i < chars.len() && is_name_char(chars[i]) {
                i += 1;
            }
            let name: String = chars[start..i].iter().collect();
            let mut j = i;
            while j < chars.len() && chars[j] == ' ' {
                j += 1;
            }
            let is_call = j < chars.len() && chars[j] == '(';
            let upper = name.to_ascii_uppercase();
            match (is_call, prefix_for(&upper)) {
                (true, Some(prefix)) => {
                    out.push_str(prefix);
                    out.push_str(&upper);
                }
                _ => out.push_str(&name),
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn newer_functions_take_their_prefix_and_older_ones_do_not() {
        assert_eq!(
            prefix_future_functions("IFS(B2>3,\"big\",TRUE,\"small\")"),
            "_xlfn.IFS(B2>3,\"big\",TRUE,\"small\")"
        );
        assert_eq!(prefix_future_functions("SUM(B2:B9)"), "SUM(B2:B9)");
        assert_eq!(prefix_future_functions("ifs(1,2)"), "_xlfn.IFS(1,2)");
        assert_eq!(
            prefix_future_functions("FILTER(A:A,B:B>1)"),
            "_xlfn._xlws.FILTER(A:A,B:B>1)"
        );
        assert_eq!(
            prefix_future_functions("STDEV.S(B2:B50)+STDEV(B2:B50)"),
            "_xlfn.STDEV.S(B2:B50)+STDEV(B2:B50)"
        );
        assert_eq!(
            prefix_future_functions("UNICHAR(160)&CHAR(120)"),
            "_xlfn.UNICHAR(160)&CHAR(120)"
        );
    }

    #[test]
    fn what_is_not_a_call_is_left_alone() {
        // A string literal, with a doubled quote inside it.
        assert_eq!(
            prefix_future_functions("\"IFS(\"\"x\"\")\"&SWITCH(1,1,\"a\")"),
            "\"IFS(\"\"x\"\")\"&_xlfn.SWITCH(1,1,\"a\")"
        );
        // Already prefixed: the name read is `_xlfn.IFS`, not in the table.
        assert_eq!(prefix_future_functions("_xlfn.IFS(1,2)"), "_xlfn.IFS(1,2)");
        // A name or a reference spelled like a function, with no `(` after it.
        assert_eq!(prefix_future_functions("LET+1"), "LET+1");
        assert_eq!(prefix_future_functions("Data!IFS1"), "Data!IFS1");
        // A space before the parenthesis still makes a call.
        assert_eq!(prefix_future_functions("IFS (1,2)"), "_xlfn.IFS (1,2)");
    }

    #[test]
    fn a_function_that_can_return_an_array_is_known() {
        assert!(can_return_array("IFS(B2>3,\"big\",TRUE,\"small\")"));
        assert!(can_return_array("_xlfn.IFS(1,2)"));
        assert!(can_return_array("_xlfn._xlws.FILTER(A:A,B:B>1)"));
        assert!(can_return_array("1+INDEX(A1:A3,2)"));
        assert!(can_return_array("sequence(3)"));
        assert!(!can_return_array("B2*2"));
        assert!(!can_return_array("SUM(B2:B9)"));
        assert!(!can_return_array("IF(B2>3,\"big\",\"small\")"));
        assert!(!can_return_array("VLOOKUP(A1,B:C,2,FALSE)"));
        assert!(!can_return_array("\"INDEX(\"&A1"));
        assert!(!can_return_array("INDEX+1"));
    }
}
