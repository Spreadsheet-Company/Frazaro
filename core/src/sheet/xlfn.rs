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

fn is_name_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_'
}

fn is_name_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '.'
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
}
