//! The sentence tokenizer: VLA_SentenceEngine.bas's `EnTokenize` and
//! `IsWordChar` (PORT.6, slice 6b).
//!
//! The shape is the VBA's, branch for branch, one pass over the characters:
//! a line feed counts a line, and two in a row with nothing between them
//! put down one paragraph marker (`|`, `PARA_TOK`); a carriage return, a
//! space and a tab are skipped; `#` runs a comment to the end of its line;
//! `"` opens a text whose only escape is a doubled quote mark, kept as one
//! token behind a leading quote mark, the VBA's own string tag; `.`, `!`
//! and `?` are the one sentence end `.`; `,`, `%` and `:` are themselves;
//! a `(` at the start of a row opens a raw VLA form, captured whole across
//! as many rows as it takes, honouring VLA's strings and `;` comments, with
//! a trailing period or comment consumed; a run of word characters
//! (`A-Za-z0-9_-`) is a word, folded, with a range colon (`A1:B10`), a
//! sheet bang (`Data!B2`), a digit-flanked decimal point and a strict
//! thousands comma (`1,000,000`) glued in; `'Q1 Data'!A1` is Excel's own
//! spaced-sheet reference, a text token with its case kept; a dropped word
//! (`the`, `please`) leaves no token and a number word becomes its digits.
//! Anything else is refused with words, never skipped: the catalogue's
//! `english-unknown-character`, with the character, `StrayCharHint`'s
//! teaching half and the line; a raw form still open at the end of the text
//! is `english-form-never-closes`.
//!
//! Positions are 1-based wherever a parser will read them ([`Tokens::tok_at`],
//! [`Tokens::line_at`]), as `TokAt` and `TokLine` are; the VBA's `toks()`
//! array and `mTokLines` are one struct here. Curly double quotes are read
//! as straight ones before anything else, as the VBA reads them.

use super::words::{is_dropped_word, number_word, stray_char_hint};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// The paragraph marker: a blank line, which closes every open block.
pub const PARA_TOK: &str = "|";

/// The tokens of a text with the raw line each began on: VBA's `toks()`
/// and `mTokLines`, 1-based through the accessors.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Tokens {
    texts: Vec<String>,
    lines: Vec<u32>,
}

impl Tokens {
    /// VBA `UBound(toks)`: how many tokens there are.
    pub fn count(&self) -> usize {
        self.texts.len()
    }

    /// VBA `TokAt`: the token at a 1-based position, or nothing past
    /// either end (every reader is total).
    pub fn tok_at(&self, p: usize) -> &str {
        if p == 0 || p > self.texts.len() {
            return "";
        }
        &self.texts[p - 1]
    }

    /// VBA `TokLine`: the line of the token at `p`; past the end, the last
    /// token's line (an end-of-text refusal names the last line); 0 with
    /// no tokens at all.
    pub fn line_at(&self, p: usize) -> u32 {
        let hi = self.lines.len();
        if p >= 1 && p <= hi {
            self.lines[p - 1]
        } else if p > hi && hi > 0 {
            self.lines[hi - 1]
        } else {
            0
        }
    }

    /// Every token's text, in order.
    pub fn texts(&self) -> &[String] {
        &self.texts
    }

    /// Every token's line, in order.
    pub fn lines(&self) -> &[u32] {
        &self.lines
    }

    fn push(&mut self, text: impl Into<String>, line: u32) {
        self.texts.push(text.into());
        self.lines.push(line);
    }
}

/// A tokenizer refusal with the line `mErrLine` records for it, which
/// `EnglishLastErrorLine` hands the host so the row can be marked.
#[derive(Clone, Debug, PartialEq)]
pub struct TokenizeRefusal {
    pub refusal: Refusal,
    pub line: u32,
}

impl std::fmt::Display for TokenizeRefusal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.refusal.text)
    }
}

/// VBA `IsWordChar`: `[A-Za-z0-9_-]`, ASCII only.
pub fn is_word_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '-'
}

/// VBA `LineSuf`: ` (line N)`, or nothing for line 0.
pub fn line_suf(n: u32) -> String {
    if n > 0 {
        format!(" (line {n})")
    } else {
        String::new()
    }
}

fn unknown_character(c: char, line: u32) -> TokenizeRefusal {
    let ch = c.to_string();
    TokenizeRefusal {
        refusal: raise(
            "english-unknown-character",
            &[
                ("char", &ch),
                ("hint", &stray_char_hint(&ch)),
                ("loc", &line_suf(line)),
            ],
        ),
        line,
    }
}

/// VBA `EnTokenize`.
pub fn tokenize(text: &str) -> Result<Tokens, TokenizeRefusal> {
    // B5: curly double quotes (Word's autocorrect) are straight quotes.
    let text = text.replace(['\u{201C}', '\u{201D}'], "\"");
    let chars: Vec<char> = text.chars().collect();
    let n = chars.len();
    // VBA `Mid$(text, i, 1)` past the end is "", which no test matches.
    let at = |i: usize| -> Option<char> { chars.get(i).copied() };
    let digit_at = |i: usize| at(i).is_some_and(|c| c.is_ascii_digit());
    let word_at = |i: usize| at(i).is_some_and(is_word_char);
    let is_break = |c: char| c == '\r' || c == '\n';

    let mut out = Tokens::default();
    let mut line_no: u32 = 1;
    // Consecutive line breaks seen; starts at 1 so the first character
    // counts as a line start (L0's raw-form rule).
    let mut nl_run: i64 = 1;
    let mut i = 0;
    while i < n {
        let c = chars[i];
        if c == '\n' {
            nl_run += 1;
            line_no += 1;
            // A blank line becomes one paragraph marker.
            if nl_run == 2
                && out.count() > 0
                && out.texts.last().map(String::as_str) != Some(PARA_TOK)
            {
                out.push(PARA_TOK, line_no);
            }
            i += 1;
        } else if c == '\r' || c == ' ' || c == '\t' {
            i += 1;
        } else if c == '#' {
            // A comment runs to the end of its line and counts as content.
            while i < n && !is_break(chars[i]) {
                i += 1;
            }
            nl_run = 0;
        } else if c == '"' {
            nl_run = 0;
            i += 1;
            let mut w = String::new();
            while i < n {
                let c = chars[i];
                if c == '"' {
                    // B7: VBA's own doubling escape.
                    if at(i + 1) == Some('"') {
                        w.push('"');
                        i += 2;
                    } else {
                        i += 1;
                        break;
                    }
                } else {
                    w.push(c);
                    i += 1;
                }
            }
            out.push(format!("\"{w}"), line_no);
        } else if c == '.' || c == '!' || c == '?' {
            nl_run = 0;
            out.push(".", line_no);
            i += 1;
        } else if c == ',' || c == '%' || c == ':' {
            nl_run = 0;
            out.push(c.to_string(), line_no);
            i += 1;
        } else if c == '(' && nl_run >= 1 {
            // L0.2: a row that begins with '(' is a raw VLA form, one token,
            // balanced across rows, honouring VLA's strings and comments.
            let v_start = line_no;
            let mut depth: i64 = 0;
            let mut form = String::new();
            while i < n {
                let c = chars[i];
                if c == '\n' {
                    form.push('\n');
                    line_no += 1;
                    i += 1;
                } else if c == '\r' {
                    i += 1;
                } else if c == '"' {
                    form.push(c);
                    i += 1;
                    while i < n {
                        let c = chars[i];
                        if is_break(c) {
                            break;
                        }
                        if c == '\\' {
                            form.push(c);
                            if let Some(d) = at(i + 1) {
                                form.push(d);
                            }
                            i += 2;
                        } else {
                            form.push(c);
                            i += 1;
                            if c == '"' {
                                break;
                            }
                        }
                    }
                } else if c == ';' {
                    while i < n && !is_break(chars[i]) {
                        form.push(chars[i]);
                        i += 1;
                    }
                } else if c == '(' {
                    depth += 1;
                    form.push(c);
                    i += 1;
                } else if c == ')' {
                    depth -= 1;
                    form.push(c);
                    i += 1;
                    if depth == 0 {
                        break;
                    }
                } else {
                    form.push(c);
                    i += 1;
                }
            }
            if depth != 0 {
                return Err(TokenizeRefusal {
                    refusal: raise(
                        "english-form-never-closes",
                        &[("depth", &depth.to_string()), ("loc", &line_suf(v_start))],
                    ),
                    line: v_start,
                });
            }
            out.push(form, v_start);
            nl_run = 0;
            // A trailing period and/or ; comment is politely consumed.
            while i < n && (chars[i] == ' ' || chars[i] == '\t') {
                i += 1;
            }
            if at(i) == Some('.') {
                i += 1;
            }
            while i < n && (chars[i] == ' ' || chars[i] == '\t') {
                i += 1;
            }
            if at(i) == Some(';') {
                while i < n && !is_break(chars[i]) {
                    i += 1;
                }
            }
        } else if is_word_char(c) {
            nl_run = 0;
            let mut w = String::new();
            while i < n {
                let c = chars[i];
                let last_digit = w.chars().last().is_some_and(|d| d.is_ascii_digit());
                if is_word_char(c) {
                    w.push(c);
                    i += 1;
                } else if c == ':' && word_at(i + 1) {
                    // A colon glued between word characters is a range colon.
                    w.push(':');
                    i += 1;
                } else if c == '!' && word_at(i + 1) {
                    // B5.2: a bang glued between word characters is Excel's
                    // sheet qualifier.
                    w.push('!');
                    i += 1;
                } else if c == '.' && !w.is_empty() {
                    // A digit-flanked period is a decimal point.
                    if last_digit && digit_at(i + 1) {
                        w.push('.');
                        i += 1;
                    } else {
                        break;
                    }
                } else if c == ',' && !w.is_empty() {
                    // A thousands separator only in the strict shape
                    // digit,DDD(non-digit); anything looser stays a comma.
                    if last_digit
                        && digit_at(i + 1)
                        && digit_at(i + 2)
                        && digit_at(i + 3)
                        && !digit_at(i + 4)
                    {
                        i += 1;
                    } else {
                        break;
                    }
                } else {
                    break;
                }
            }
            let w = fold(&w);
            if !is_dropped_word(&w) {
                out.push(number_word(&w), line_no);
            }
        } else if c == '\'' {
            // B5.3: Excel's spaced-sheet syntax, 'Q1 Data'!A1, complete or
            // refused: the quote closes on this line and !reference follows.
            let mut q_end = i + 1;
            let mut closed = false;
            while q_end < n {
                let rc = chars[q_end];
                if rc == '\'' {
                    closed = true;
                    break;
                }
                if is_break(rc) {
                    break;
                }
                q_end += 1;
            }
            let mut taken = false;
            if closed && at(q_end + 1) == Some('!') && word_at(q_end + 2) {
                let mut ref_end = q_end + 2;
                while ref_end < n {
                    let rc = chars[ref_end];
                    // A word character, or a range colon that chains.
                    if is_word_char(rc) || (rc == ':' && word_at(ref_end + 1)) {
                        ref_end += 1;
                    } else {
                        break;
                    }
                }
                // String-tagged, case preserved.
                let tok: String = chars[i..ref_end].iter().collect();
                out.push(format!("\"{tok}"), line_no);
                nl_run = 0;
                i = ref_end;
                taken = true;
            }
            if !taken {
                return Err(unknown_character(c, line_no));
            }
        } else {
            // B5: an unrecognized character is refused with words, never
            // skipped.
            return Err(unknown_character(c, line_no));
        }
    }
    Ok(out)
}

/// `EnglishTokenReport` (VLA_SentenceEngine.bas): one line per token,
/// `<line><TAB><token>`, a backslash, a line feed, a carriage return and a
/// tab in a token escaped as `\\`, `\n`, `\r` and `\t`, so that a report
/// line is one token. The token golden is written in this shape.
pub fn token_report(text: &str) -> Result<String, TokenizeRefusal> {
    let toks = tokenize(text)?;
    let mut r = String::new();
    for (t, l) in toks.texts.iter().zip(&toks.lines) {
        let e = t
            .replace('\\', "\\\\")
            .replace('\n', "\\n")
            .replace('\r', "\\r")
            .replace('\t', "\\t");
        r.push_str(&format!("{l}\t{e}\n"));
    }
    Ok(r)
}

/// `VlaWriteTokenGolden` (VLA_Tests.bas): the fixture `scripts/tokenize.txt`
/// holds programs, each under a line `=== <name>`; the report repeats the
/// header and then `token_report`'s lines, or one line
/// `REFUSED<TAB><line><TAB><id><TAB><text>` where the tokenizer refuses.
/// Sections are split as the VBA splits them: lines before the first header
/// are ignored, a program's lines are joined with a line feed, and the
/// file's own trailing line break gives the last program a trailing one.
pub fn fixture_report(fixture: &str) -> String {
    let text = fixture.replace("\r\n", "\n");
    let mut out = String::new();
    let mut header: Option<&str> = None;
    let mut program = String::new();
    let mut first = true;
    for ln in text.split('\n').chain(std::iter::once("=== ")) {
        if ln.starts_with("=== ") {
            if let Some(h) = header.take() {
                out.push_str(h);
                out.push('\n');
                match token_report(&program) {
                    Ok(r) => out.push_str(&r),
                    Err(e) => out.push_str(&format!(
                        "REFUSED\t{}\t{}\t{}\n",
                        e.line, e.refusal.id, e.refusal.text
                    )),
                }
            }
            header = Some(ln);
            program.clear();
            first = true;
        } else if header.is_some() {
            if first {
                program.push_str(ln);
                first = false;
            } else {
                program.push('\n');
                program.push_str(ln);
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn texts(text: &str) -> Vec<String> {
        tokenize(text).unwrap().texts().to_vec()
    }

    fn lines(text: &str) -> Vec<u32> {
        tokenize(text).unwrap().lines().to_vec()
    }

    #[test]
    fn words_fold_drop_and_rewrite_numbers() {
        assert_eq!(
            texts("Set the total to 5."),
            vec!["set", "total", "to", "5", "."]
        );
        assert_eq!(texts("Show total!"), vec!["show", "total", "."]);
        assert_eq!(texts("Is it done?"), vec!["is", "it", "done", "."]);
        // "please" and "the" leave no token; "a" stays (it can be a column).
        assert_eq!(
            texts("Please add a value to the list, then add twelve to twenty."),
            vec!["add", "a", "value", "to", "list", ",", "then", "add", "12", "to", "20", "."]
        );
        let t = tokenize("Set x to 1.").unwrap();
        assert_eq!(t.count(), 5);
        assert_eq!(t.tok_at(1), "set");
        assert_eq!(t.tok_at(0), "");
        assert_eq!(t.tok_at(6), "");
        assert_eq!(t.line_at(5), 1);
        assert_eq!(t.line_at(9), 1);
        assert_eq!(Tokens::default().line_at(1), 0);
    }

    #[test]
    fn numbers_glue_a_decimal_point_and_a_strict_thousands_comma() {
        assert_eq!(
            texts("Set n to 1,000,000 and m to 3.14159 and k to 1,00 and p to 15%."),
            vec![
                "set", "n", "to", "1000000", "and", "m", "to", "3.14159", "and", "k", "to", "1",
                ",", "00", "and", "p", "to", "15", "%", "."
            ]
        );
        // A period after a word is the sentence end, even before a digit.
        assert_eq!(texts("Log x.5"), vec!["log", "x", ".", "5"]);
    }

    #[test]
    fn quotes_keep_their_text_behind_the_tag() {
        assert_eq!(
            texts("Say \"He said \"\"hi\"\" to me\"."),
            vec!["say", "\"He said \"hi\" to me", "."]
        );
        assert_eq!(
            texts("Say \u{201C}curly\u{201D}."),
            vec!["say", "\"curly", "."]
        );
        // An unterminated text runs to the end.
        assert_eq!(texts("Say \"open"), vec!["say", "\"open"]);
    }

    #[test]
    fn references_glue_colons_bangs_and_spaced_sheets() {
        assert_eq!(
            texts("Copy Data!A1:B10 to 'Q1 Data'!C1 and column C:D."),
            vec![
                "copy",
                "data!a1:b10",
                "to",
                "\"'Q1 Data'!C1",
                "and",
                "column",
                "c:d",
                "."
            ]
        );
        // A block colon is followed by whitespace, so it stays a token.
        assert_eq!(texts("Repeat 3 times:"), vec!["repeat", "3", "times", ":"]);
    }

    #[test]
    fn blank_lines_mark_paragraphs_and_comments_count_as_content() {
        let text = "Repeat 3 times:\n  Increase x by 1.\n\n# a comment\nLog x.\n\n\nLog x again.";
        assert_eq!(
            texts(text),
            vec![
                "repeat", "3", "times", ":", "increase", "x", "by", "1", ".", "|", "log", "x", ".",
                "|", "log", "x", "again", "."
            ]
        );
        assert_eq!(
            lines(text),
            vec![1, 1, 1, 1, 2, 2, 2, 2, 2, 4, 5, 5, 5, 7, 8, 8, 8, 8]
        );
        // A marker is never doubled and never the first token; blank lines
        // after the last sentence leave one, as the VBA's do.
        assert_eq!(texts("\n\nLog x.\n\n\n\n"), vec!["log", "x", ".", "|"]);
        assert_eq!(lines("\n\nLog x.\n\n\n\n"), vec![3, 3, 3, 5]);
        assert!(texts("").is_empty());
        assert!(texts("# only a comment\n").is_empty());
    }

    #[test]
    fn a_raw_row_is_one_token_across_rows() {
        let text =
            "(begin ; opening\n  (set! x \"a ) b \\\" c\")\n  (log x)) . ; trailing\nLog x.\n(log 2).";
        let t = tokenize(text).unwrap();
        assert_eq!(
            t.texts(),
            &[
                "(begin ; opening\n  (set! x \"a ) b \\\" c\")\n  (log x))",
                "log",
                "x",
                ".",
                "(log 2)"
            ]
        );
        assert_eq!(t.lines(), &[1, 4, 4, 4, 5]);
        // A '(' not at a row's start is a stray character.
        let r = tokenize("Set x to (1).").unwrap_err();
        assert_eq!(r.refusal.id, "english-unknown-character");
        assert_eq!(
            r.refusal.text,
            "I don't understand the character '(' - a VLA form must begin its row - start the cell with '(' to write VLA, or say it in words (line 1)"
        );
        // A comment line still leaves the next row a row start.
        assert_eq!(texts("# c\n(log 1)"), vec!["(log 1)"]);
    }

    #[test]
    fn the_refusals_name_the_character_the_hint_and_the_line() {
        let r = tokenize("Set x to 1.\nSet price to $5.").unwrap_err();
        assert_eq!(r.line, 2);
        assert_eq!(
            r.refusal.text,
            "I don't understand the character '$' - write the plain number (or 'Format ... as currency.' for display) (line 2)"
        );
        let r = tokenize("Don't do that.").unwrap_err();
        assert!(r
            .refusal
            .text
            .starts_with("I don't understand the character '''"));
        let r = tokenize("Set x to 5 \u{2014} done.").unwrap_err();
        assert!(r.refusal.text.contains("use a plain hyphen"));
        let r = tokenize("Set x to 1.\n(begin (set! x").unwrap_err();
        assert_eq!(r.refusal.id, "english-form-never-closes");
        assert_eq!(r.line, 2);
        assert_eq!(
            r.refusal.text,
            "this VLA form never closes - 2 '(' still open by the end of the program (line 2)"
        );
    }

    #[test]
    fn the_report_escapes_what_breaks_a_line() {
        assert_eq!(
            token_report("Set x to \"a\\b\".\n(log\n1)").unwrap(),
            "1\tset\n1\tx\n1\tto\n1\t\"a\\\\b\n1\t.\n2\t(log\\n1)\n"
        );
        let fx = "ignored\n=== one\nLog x.\n=== refused\nSet x to $1.\n=== empty\n";
        assert_eq!(
            fixture_report(fx),
            "=== one\n1\tlog\n1\tx\n1\t.\n=== refused\nREFUSED\t1\tenglish-unknown-character\tI don't understand the character '$' - write the plain number (or 'Format ... as currency.' for display) (line 1)\n=== empty\n"
        );
    }

    const FIXTURE: &str = include_str!("../../../scripts/tokenize.txt");
    const GOLDEN: &str = include_str!("../../../scripts/tokenize_golden.txt");

    #[test]
    fn the_token_golden_is_reproduced() {
        // scripts/tokenize_golden.txt is written by VlaWriteTokenGolden
        // (VLA_Tests.bas) from the reference's EnTokenize; this is the
        // tokenizer's oracle, compared as the treaty compares a golden (LF,
        // trailing blank lines dropped).
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
                "the token golden differs at char {at} (line {line}):\n  want: {}\n  got:  {}",
                context(&want),
                context(&got)
            );
        }
    }

    /// Prints the fixture's report, for writing the prediction of the golden
    /// before the reference has written it:
    /// `cargo test -p frazaro-core print_token_report -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn print_token_report() {
        print!("{}", fixture_report(FIXTURE));
    }
}
