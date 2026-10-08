//! The VBA behaviours the port must match: docs/INTRINSICS.md, one function
//! per entry that is a function at all (PORT.5, slice 5a).
//!
//! The reference implementation is VBA (SD-18), and six of its behaviours
//! leak into the language's own meaning: how a name folds, how a number is
//! read and written, how text is indexed, what a Collection is, how text
//! compares, and what a refusal is. INTRINSICS.md names each with its VBA
//! citation; this module is where the core answers each one, so that every
//! other module reaches for `intrinsics::fold` and never for `to_lowercase`.
//!
//! 1. Case folding: [`fold`], ASCII A-Z only, exactly `VLA_Identity.Fold`.
//! 2. Numbers: [`val`] is VBA's `Val`, [`str_`] is VBA's `Str$`, both
//!    locale-free by construction here, where the VBA had to choose them
//!    over `CDbl`/`CStr` to be so. [`is_numeric`] answers what `IsNumeric`
//!    answers for a token on an invariant-locale machine, and nothing a
//!    locale could change (a thousands separator, a currency sign) is a
//!    number here. [`is_numeric_literal_text`] is VLA.bas's own stricter
//!    test, the one expand-time arithmetic uses.
//! 3. Text is 1-based in VBA except `Split`: a convention for the porter,
//!    not a function. Rust's `&str` is 0-based throughout; every `Mid$(s,
//!    i, n)` becomes a slice at `i - 1`.
//! 4. `Collection`: 1-based, insertion-ordered, positional. A `Vec` with
//!    `v[i - 1]`; nothing on the translate path uses keyed access.
//! 5. `Option Compare Binary`: Rust's `==` on `str` is already exact.
//! 6. `Err.Raise` at the seam: a refusal is a value, `Result<_, Refusal>`,
//!    built from the message catalogue (the `messages` module).
//!
//! The one place this module is allowed to be *more* than the VBA: there is
//! no locale here to bend to, so the second entry's class of bug
//! (`DATALOG.17`, a number spelled by `CStr` on a comma-decimal Windows)
//! cannot occur, and a test below says so.

/// VBA `VLA_Identity.Fold`: A-Z to a-z, every other code point untouched.
/// Not `to_lowercase`, which is Unicode-aware and would fold the Turkish
/// dotted I and the diacritics `Fold` deliberately leaves alone.
pub fn fold(s: &str) -> String {
    s.chars()
        .map(|c| {
            if c.is_ascii_uppercase() {
                c.to_ascii_lowercase()
            } else {
                c
            }
        })
        .collect()
}

/// VLA.bas's `IsNumericLiteralText`: an optional leading `-`, then digits
/// with at most one `.`, at least one digit, and nothing else. The strict
/// test expand-time arithmetic (`+expand` and its comparisons) applies to
/// its operands before `Val` ever reads them.
pub fn is_numeric_literal_text(s: &str) -> bool {
    let body = s.strip_prefix('-').unwrap_or(s);
    if body.is_empty() {
        return false;
    }
    let mut saw_digit = false;
    let mut saw_dot = false;
    for c in body.chars() {
        if c.is_ascii_digit() {
            saw_digit = true;
        } else if c == '.' && !saw_dot {
            saw_dot = true;
        } else {
            return false;
        }
    }
    saw_digit
}

/// What VBA's `IsNumeric` says of a token, read without a locale: an
/// optional sign, digits with at most one `.`, at least one digit, an
/// optional exponent (`e`, `E`, `d` or `D`, an optional sign, digits), or
/// a `&H` hexadecimal or `&O` octal literal. `SymName`, `FormulaQuote` and
/// `QuoteDatum` ask this of an atom to pass a number through unmangled.
///
/// The VBA's own answer bends to Windows' regional settings for a token
/// holding `,` or a currency sign; here such a token is not a number, and
/// the tokenizer never produces one from the corpus.
pub fn is_numeric(s: &str) -> bool {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return false;
    }
    let mut i = 0;
    if bytes[i] == b'+' || bytes[i] == b'-' {
        i += 1;
    }
    if i + 1 < bytes.len() && bytes[i] == b'&' {
        let radix = match bytes[i + 1] {
            b'H' | b'h' => 16,
            b'O' | b'o' => 8,
            _ => return false,
        };
        let digits = &bytes[i + 2..];
        return !digits.is_empty() && digits.iter().all(|b| (*b as char).is_digit(radix));
    }
    let mut saw_digit = false;
    let mut saw_dot = false;
    while i < bytes.len() {
        let b = bytes[i];
        if b.is_ascii_digit() {
            saw_digit = true;
        } else if b == b'.' && !saw_dot {
            saw_dot = true;
        } else {
            break;
        }
        i += 1;
    }
    if !saw_digit {
        return false;
    }
    if i == bytes.len() {
        return true;
    }
    if !matches!(bytes[i], b'e' | b'E' | b'd' | b'D') {
        return false;
    }
    i += 1;
    if i < bytes.len() && (bytes[i] == b'+' || bytes[i] == b'-') {
        i += 1;
    }
    let exp = &bytes[i..];
    !exp.is_empty() && exp.iter().all(u8::is_ascii_digit)
}

/// VBA `Val`: blanks, tabs and line feeds are dropped from the whole text,
/// then the longest leading number is read and the rest ignored; no number
/// at all is 0. `&H` and `&O` prefixes read hexadecimal and octal. The
/// decimal separator is `.` on every machine, which is why VLA.bas chose
/// `Val` over `CDbl` (INTRINSICS.md, entry 2).
pub fn val(s: &str) -> f64 {
    let cleaned: Vec<u8> = s
        .bytes()
        .filter(|b| !matches!(b, b' ' | b'\t' | b'\n'))
        .collect();
    let mut i = 0;
    let mut negative = false;
    if i < cleaned.len() && (cleaned[i] == b'+' || cleaned[i] == b'-') {
        negative = cleaned[i] == b'-';
        i += 1;
    }
    if i + 1 < cleaned.len() && cleaned[i] == b'&' {
        let radix = match cleaned[i + 1] {
            b'H' | b'h' => 16,
            b'O' | b'o' => 8,
            _ => return 0.0,
        };
        let mut value: f64 = 0.0;
        for b in &cleaned[i + 2..] {
            match (*b as char).to_digit(radix) {
                Some(d) => value = value * f64::from(radix) + f64::from(d),
                None => break,
            }
        }
        return if negative { -value } else { value };
    }
    let start = i;
    let mut saw_digit = false;
    let mut saw_dot = false;
    while i < cleaned.len() {
        let b = cleaned[i];
        if b.is_ascii_digit() {
            saw_digit = true;
        } else if b == b'.' && !saw_dot {
            saw_dot = true;
        } else {
            break;
        }
        i += 1;
    }
    if !saw_digit {
        return 0.0;
    }
    let mut end = i;
    if i < cleaned.len() && matches!(cleaned[i], b'e' | b'E' | b'd' | b'D') {
        let mut j = i + 1;
        if j < cleaned.len() && (cleaned[j] == b'+' || cleaned[j] == b'-') {
            j += 1;
        }
        let digits_start = j;
        while j < cleaned.len() && cleaned[j].is_ascii_digit() {
            j += 1;
        }
        if j > digits_start {
            end = j;
        }
    }
    let mut text = String::with_capacity(end - start);
    for b in &cleaned[start..end] {
        text.push(match *b {
            b'd' | b'D' => 'e',
            other => other as char,
        });
    }
    let magnitude: f64 = text.parse().unwrap_or(0.0);
    if negative {
        -magnitude
    } else {
        magnitude
    }
}

/// VBA `Str$` of a Double: a leading space for a number that is not
/// negative, at most fifteen significant digits, trailing zeros dropped,
/// no zero before the decimal point of a fraction (` .5`, not ` 0.5`), and
/// exponent notation (`1E-05`, `1E+15`) once the exponent is below -4 or at
/// 15 or above, the C library's `%.15G` rule the VBA runtime follows.
/// `+expand` writes `Trim$(Str$(a + b))`, so a sum spelled here must agree
/// with VBA's to the byte.
pub fn str_(x: f64) -> String {
    if x == 0.0 || x.is_nan() {
        // Str$(0) is " 0"; -0 is 0 to VBA. NaN cannot arise from Val.
        return " 0".to_string();
    }
    let sign = if x < 0.0 { "-" } else { " " };
    let mag = x.abs();
    if mag.is_infinite() {
        return format!("{sign}1.#INF");
    }
    // Fifteen significant digits: d.dddddddddddddde<exp>.
    let sci = format!("{mag:.14e}");
    let (mantissa, exp_text) = sci
        .split_once('e')
        .expect("Rust's exponent format always writes an e");
    let exp: i32 = exp_text.parse().expect("an exponent is an integer");
    let digits: String = mantissa.chars().filter(|c| *c != '.').collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let body = if !(-4..15).contains(&exp) {
        let (first, rest) = digits.split_at(1);
        let mut m = first.to_string();
        if !rest.is_empty() {
            m.push('.');
            m.push_str(rest);
        }
        let exp_sign = if exp < 0 { '-' } else { '+' };
        format!("{m}E{exp_sign}{:02}", exp.abs())
    } else if exp < 0 {
        // A fraction: Str$ writes no zero before the point.
        format!(".{}{}", "0".repeat((-exp - 1) as usize), digits)
    } else {
        let point = exp as usize + 1;
        if digits.len() <= point {
            format!("{digits}{}", "0".repeat(point - digits.len()))
        } else {
            let (int_part, frac_part) = digits.split_at(point);
            format!("{int_part}.{frac_part}")
        }
    };
    format!("{sign}{body}")
}

#[cfg(test)]
mod tests {
    use super::*;

    // Entry 1: Fold, the INTRINSICS.md examples and the trap it names.
    #[test]
    fn fold_touches_only_ascii_capitals() {
        assert_eq!(fold("Set-Formula"), "set-formula");
        assert_eq!(fold("ABCxyz019"), "abcxyz019");
        // The Turkish dotted capital I and every accented letter pass
        // through unchanged, where to_lowercase would fold them.
        assert_eq!(fold("İSTANBUL"), "İstanbul");
        assert_eq!(fold("ÉCOLE Ñ"), "École Ñ");
        assert_eq!(fold(""), "");
    }

    // Entry 2: Val and Str$, locale-free.
    #[test]
    fn val_reads_the_leading_number_as_documented() {
        assert_eq!(val("5.5"), 5.5);
        assert_eq!(val("-3"), -3.0);
        assert_eq!(val("  12abc"), 12.0);
        // Microsoft's own example: blanks are dropped, not stopped at.
        assert_eq!(val(" 1615 198th Street N.E."), 1615198.0);
        assert_eq!(val("&HFFFF"), 65535.0);
        assert_eq!(val("&O17"), 15.0);
        assert_eq!(val("1e3"), 1000.0);
        assert_eq!(val("1d2"), 100.0);
        assert_eq!(val("1e"), 1.0);
        assert_eq!(val(".5"), 0.5);
        assert_eq!(val("5."), 5.0);
        assert_eq!(val("abc"), 0.0);
        assert_eq!(val(""), 0.0);
        assert_eq!(val("-"), 0.0);
    }

    #[test]
    fn str_spells_a_double_as_vba_does() {
        assert_eq!(str_(459.0), " 459");
        assert_eq!(str_(-459.65), "-459.65");
        assert_eq!(str_(459.001), " 459.001");
        assert_eq!(str_(0.0), " 0");
        assert_eq!(str_(3.0), " 3");
        assert_eq!(str_(0.5), " .5");
        assert_eq!(str_(-0.25), "-.25");
        assert_eq!(str_(0.0001), " .0001");
        assert_eq!(str_(0.00001), " 1E-05");
        assert_eq!(str_(1e15), " 1E+15");
        assert_eq!(str_(1e14), " 100000000000000");
        assert_eq!(str_(123456789012345680.0), " 1.23456789012346E+17");
        // Fifteen digits, not seventeen: 0.1 + 0.2 is .3 to VBA.
        assert_eq!(str_(0.1 + 0.2), " .3");
        assert_eq!(str_(1.0 / 3.0), " .333333333333333");
    }

    #[test]
    fn expand_time_sum_spells_as_trim_of_str() {
        // (+expand a b) in VLA.bas: Trim$(Str$(Val(a) + Val(b))).
        let sum = |a: &str, b: &str| str_(val(a) + val(b)).trim().to_string();
        assert_eq!(sum("1", "2"), "3");
        assert_eq!(sum("2.5", "0.25"), "2.75");
        assert_eq!(sum("-1", "0.5"), "-.5");
    }

    #[test]
    fn comma_decimal_cannot_occur_here() {
        // DATALOG.17's class: the VBA spells a cell's number with CStr,
        // which writes 0,5 on a comma-decimal Windows. This core has no
        // locale to consult, so the period is the only separator it can
        // write or read, on every machine.
        assert_eq!(str_(0.5), " .5");
        assert!(!str_(0.5).contains(','));
        // And a comma is where Val stops reading, as VBA's own does.
        assert_eq!(val("0,5"), 0.0);
        assert_eq!(val("2,5"), 2.0);
        assert!(!is_numeric("0,5"));
    }

    #[test]
    fn is_numeric_reads_a_token_without_a_locale() {
        for yes in [
            "5", "-5", "+5", "0.08", ".5", "5.", "1e5", "1E+05", "1d5", "&HFF", "&o17",
        ] {
            assert!(is_numeric(yes), "{yes} is numeric");
        }
        for no in [
            "", "-", ".", "e5", "1e", "1-2", "x", "1.2.3", "&H", "1,000", "$5", "5 ",
        ] {
            assert!(!is_numeric(no), "{no} is not numeric");
        }
    }

    #[test]
    fn numeric_literal_text_is_the_strict_test() {
        for yes in ["5", "-5", "0.08", "-.5", "5.", "007"] {
            assert!(is_numeric_literal_text(yes), "{yes}");
        }
        for no in ["", "-", ".", "+5", "1e5", "1.2.3", "5x", "&HFF"] {
            assert!(!is_numeric_literal_text(no), "{no}");
        }
    }
}
