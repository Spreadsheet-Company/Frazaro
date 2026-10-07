//! The formula sink's scan (SEC.15): `VLA_Runtime.bas`'s `VlaFormulaEgress`,
//! one arm per arm, held to the egress golden.
//!
//! Every formula a program writes goes through one sink: the add-in's
//! `VlaSetFormula` on both backends, and the writer's `write_formula` here
//! (`build`). The sink asks this scan first: a text that calls a function
//! which reaches outside the workbook on its own (`WEBSERVICE`, `HYPERLINK`,
//! `RTD`, the XLM functions that start a program or a DLL, Sheets' `IMPORT*`
//! and `GOOGLE*`, Calc's `DDE`), or holds a DDE link (a `|` outside a
//! string, a quoted name or a bracket group), is refused by name and nothing
//! is written. The list is [`EGRESS_NAMES`], the one place in the core, each
//! name's reason in SEC.15's entry (`docs/BETA_REARVIEW.md`) and in
//! `VlaFormulaEgress`'s header; `tools/check_egress_golden.ps1` holds it
//! equal to the VBA's and every name reached by a case of the golden.
//!
//! The rule, the reference's: any text written through the formula member
//! is scanned, whatever it begins with; a string literal (a doubled quote
//! inside), a quoted name and a bracket group are stepped over with
//! `refers`' own helpers, so the crate has one string-skipping rule; a run
//! of identifier characters followed by spaces, tabs or line breaks and then
//! `(` is a call; the file prefixes `_xlfn.` and `_xlws.` are stripped from
//! its front, as many as stand there; the rest is compared without case to
//! the list, whole; the first hit in text order is named in upper case, a
//! DDE link as `DDE`. `scripts/egress.txt` to `scripts/egress_golden.txt`,
//! written by the reference's `VlaWriteEgressGolden`, is the golden
//! `the_egress_golden_is_reproduced` holds this module to.

use crate::intrinsics::fold;
use crate::refers::{bracket_group_end, is_ident_char, skip_delimited};

/// The functions the sink refuses, folded: the one list in the core.
pub const EGRESS_NAMES: &[&str] = &[
    "webservice",
    "filterxml",
    "hyperlink",
    "rtd",
    "image",
    "stockhistory",
    "call",
    "register",
    "register.id",
    "exec",
    "initiate",
    "execute",
    "poke",
    "request",
    "send.mail",
    "dde",
    "importxml",
    "importdata",
    "importhtml",
    "importrange",
    "importfeed",
    "googlefinance",
    "googletranslate",
];

/// What the scan names a DDE link, `application|topic!item`.
pub const DDE: &str = "DDE";

/// The function a formula text calls that reaches outside the workbook, in
/// upper case, or `DDE` for a DDE link; `None` when it holds none.
pub fn egress_call(formula_text: &str) -> Option<String> {
    let s: Vec<char> = formula_text.chars().collect();
    let n = s.len();
    let mut i = 0;
    while i < n {
        let ch = s[i];
        if ch == '"' {
            i = skip_delimited(&s, i, '"');
        } else if ch == '\'' {
            i = skip_delimited(&s, i, '\'');
        } else if ch == '[' {
            i = bracket_group_end(&s, i);
        } else if ch == '|' {
            return Some(DDE.to_string());
        } else if is_ident_char(ch) {
            let mut k = i;
            while k < n && is_ident_char(s[k]) {
                k += 1;
            }
            let mut tok = fold(&s[i..k].iter().collect::<String>());
            i = k;
            while i < n && matches!(s[i], ' ' | '\t' | '\r' | '\n') {
                i += 1;
            }
            if i < n && s[i] == '(' {
                loop {
                    let stripped = tok
                        .strip_prefix("_xlfn.")
                        .or_else(|| tok.strip_prefix("_xlws."))
                        .map(str::to_string);
                    match stripped {
                        Some(t) => tok = t,
                        None => break,
                    }
                }
                if EGRESS_NAMES.contains(&tok.as_str()) {
                    return Some(tok.to_ascii_uppercase());
                }
            }
        } else {
            i += 1;
        }
    }
    None
}

/// One case's record, as `VlaWriteEgressGolden` writes it: the case's lines
/// joined, blank lines at the end dropped (the fixture's own last line is
/// one), then `REFUSED<TAB><name>` or `WRITTEN`.
pub fn report_for(body: &[&str]) -> String {
    let mut text = body.join("\n");
    while text.ends_with('\n') {
        text.pop();
    }
    match egress_call(&text) {
        Some(name) => format!("REFUSED\t{name}\r\n"),
        None => "WRITTEN\r\n".to_string(),
    }
}

/// `VlaWriteEgressGolden`'s text for a fixture: each `=== ` header, then its
/// case's record; the lines before the first header are ignored.
pub fn fixture_report(fixture: &str) -> String {
    let text = fixture.replace("\r\n", "\n");
    let mut out = String::new();
    let mut header: Option<&str> = None;
    let mut body: Vec<&str> = Vec::new();
    for line in text.split('\n') {
        if line.starts_with("=== ") {
            if let Some(h) = header {
                out.push_str(h);
                out.push_str("\r\n");
                out.push_str(&report_for(&body));
            }
            header = Some(line);
            body.clear();
        } else if header.is_some() {
            body.push(line);
        }
    }
    if let Some(h) = header {
        out.push_str(h);
        out.push_str("\r\n");
        out.push_str(&report_for(&body));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &str = include_str!("../../scripts/egress.txt");
    const GOLDEN: &str = include_str!("../../scripts/egress_golden.txt");

    #[test]
    fn the_egress_golden_is_reproduced() {
        // scripts/egress_golden.txt is written by VlaWriteEgressGolden
        // (VLA_Tests.bas) from the reference's VlaFormulaEgress; compared as
        // the treaty compares a golden (LF, trailing blank lines dropped).
        let norm = |s: &str| s.replace("\r\n", "\n").trim_end_matches('\n').to_string();
        let got = norm(&fixture_report(FIXTURE));
        let want = norm(GOLDEN);
        if got != want {
            let at = got
                .chars()
                .zip(want.chars())
                .position(|(a, b)| a != b)
                .unwrap_or(got.len().min(want.len()));
            let line = want[..at].matches('\n').count() + 1;
            let context = |s: &str| {
                let start = s[..at].rfind('\n').map(|i| i + 1).unwrap_or(0);
                let end = s[at..].find('\n').map(|i| at + i).unwrap_or(s.len());
                s[start..end].to_string()
            };
            panic!(
                "the egress golden differs at char {at} (line {line}):\n  want: {}\n  got:  {}",
                context(&want),
                context(&got)
            );
        }
    }

    #[test]
    fn the_pieces_the_table_reaches_only_through_the_whole() {
        assert_eq!(egress_call(""), None);
        assert_eq!(egress_call("=A1+B1"), None);
        assert_eq!(
            egress_call("=register.id(1)").as_deref(),
            Some("REGISTER.ID")
        );
        assert_eq!(egress_call("=a|b!c").as_deref(), Some("DDE"));
        // Every listed name is a call with ( after it, and a name without.
        for name in EGRESS_NAMES {
            let want = name.to_ascii_uppercase();
            assert_eq!(egress_call(&format!("={name}(1)")), Some(want), "{name}");
            assert_eq!(egress_call(&format!("={name}+1")), None, "{name}");
            assert_eq!(egress_call(&format!("=\"{name}(\"")), None, "{name}");
        }
        assert_eq!(EGRESS_NAMES.len(), 23);
    }

    /// Prints the fixture's report, for checking a prediction of the golden:
    /// `cargo test -p frazaro-core print_egress_report -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn print_egress_report() {
        print!("{}", fixture_report(FIXTURE));
    }
}
