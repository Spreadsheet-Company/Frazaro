//! The refusal golden (PORT.6, slice 6f): every refusal id the sentence
//! engine raises on the path, in the situation that raises it, with its
//! text. `scripts/refusals.txt` holds the cases, each under a line
//! `=== program <name>` or `=== phrasebook <name>`, and
//! `scripts/refusals_golden.txt` the reference's reading of each, written
//! by `VlaWriteRefusalGolden` (VLA_Tests.bas); this module reads the same
//! fixture the same way and its test holds the core to that file, as the
//! tokenizer is held to the token golden.
//!
//! A program case is translated by `EnglishToVla` after `english.vla` has
//! loaded; a line of exactly `---` inside it divides a phrasebook text
//! (above, loaded after `english.vla` under the case's name) from the
//! program (below). A phrasebook case is loaded after `english.vla` under
//! the case's name. The record is one of `REFUSED<TAB><line><TAB><id><TAB>
//! <text>` (the line `EnglishLastErrorLine()` for a program and `-` for a
//! phrasebook, whose refusals carry their own location), `TRANSLATED`, or
//! `LOADED<TAB><n> rules`.

use super::grammar::Grammar;

/// One case of the fixture: its header line, kind (`program` or
/// `phrasebook`), name, and body as the VBA assembles it (the lines joined
/// with a line feed; the file's own trailing break gives the last case a
/// trailing one).
#[derive(Clone, Debug, PartialEq)]
pub struct Case {
    pub header: String,
    pub kind: String,
    pub name: String,
    pub body: String,
}

/// The fixture split as `VlaWriteRefusalGolden` splits it: lines before the
/// first header are ignored; a header's second word is the kind and the rest
/// the name.
pub fn fixture_cases(fixture: &str) -> Vec<Case> {
    let text = fixture.replace("\r\n", "\n");
    let mut cases: Vec<Case> = Vec::new();
    let mut header: Option<&str> = None;
    let mut body = String::new();
    let mut first = true;
    for ln in text.split('\n').chain(std::iter::once("=== ")) {
        if ln.starts_with("=== ") {
            if let Some(h) = header.take() {
                let rest = &h[4..];
                let (kind, name) = match rest.find(' ') {
                    Some(sp) => (&rest[..sp], &rest[sp + 1..]),
                    None => (rest, ""),
                };
                cases.push(Case {
                    header: h.to_string(),
                    kind: kind.to_string(),
                    name: name.to_string(),
                    body: std::mem::take(&mut body),
                });
            }
            header = Some(ln);
            body.clear();
            first = true;
        } else if header.is_some() {
            if first {
                body.push_str(ln);
                first = false;
            } else {
                body.push('\n');
                body.push_str(ln);
            }
        }
    }
    cases
}

/// A program case's body split at its `---` line: the phrasebook text, if
/// any, and the program.
fn split_divider(body: &str) -> (Option<&str>, &str) {
    match body.find("\n---\n") {
        Some(p) => (Some(&body[..p]), &body[p + 5..]),
        None => (None, body),
    }
}

/// One case's record, from a grammar that holds `english.vla` and nothing
/// else (it is cloned, never changed).
fn case_record(english: &Grammar, case: &Case) -> String {
    let mut g = english.clone();
    match case.kind.as_str() {
        "phrasebook" => match g.load_vocabulary_text(&case.body, &case.name) {
            Ok(n) => format!("LOADED\t{n} rules\n"),
            Err(e) => format!("REFUSED\t-\t{}\t{}\n", e.id, e.text),
        },
        "program" => {
            let (book, program) = split_divider(&case.body);
            if let Some(book) = book {
                if let Err(e) = g.load_vocabulary_text(book, &case.name) {
                    return format!("REFUSED\t-\t{}\t{}\n", e.id, e.text);
                }
            }
            match g.translate_program_at(program) {
                Ok(_) => "TRANSLATED\n".to_string(),
                Err(e) => format!(
                    "REFUSED\t{}\t{}\t{}\n",
                    e.line, e.refusal.id, e.refusal.text
                ),
            }
        }
        other => format!("UNKNOWN KIND\t{other}\n"),
    }
}

/// `VlaWriteRefusalGolden`'s text for a fixture: each header, then its
/// record. `english` is `english.vla`'s text and `prelude` the prelude's.
pub fn refusal_report(fixture: &str, prelude: &str, english: &str) -> String {
    let mut g = Grammar::new(prelude);
    let base = match g.load_vocabulary_text(english, "english.vla") {
        Ok(_) => g,
        Err(e) => return format!("REFUSED\t-\t{}\t{}\n", e.id, e.text),
    };
    let mut out = String::new();
    for case in fixture_cases(fixture) {
        out.push_str(&case.header);
        out.push('\n');
        out.push_str(&case_record(&base, &case));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../../scripts/polyglotta/english.vla");
    const FIXTURE: &str = include_str!("../../../scripts/refusals.txt");
    const GOLDEN: &str = include_str!("../../../scripts/refusals_golden.txt");

    #[test]
    fn the_fixture_splits_as_the_writer_splits_it() {
        let fx = "ignored\n=== program one\nLog 1.\n=== phrasebook two words\n(a)\n---\nSet x to 1.\n=== program three\n";
        let cases = fixture_cases(fx);
        assert_eq!(cases.len(), 3);
        assert_eq!(
            (
                cases[0].kind.as_str(),
                cases[0].name.as_str(),
                cases[0].body.as_str()
            ),
            ("program", "one", "Log 1.")
        );
        assert_eq!(cases[1].name, "two words");
        assert_eq!(cases[1].body, "(a)\n---\nSet x to 1.");
        assert_eq!(split_divider(&cases[1].body), (Some("(a)"), "Set x to 1."));
        assert_eq!(cases[2].body, "");
        assert_eq!(split_divider("Log 1."), (None, "Log 1."));
    }

    #[test]
    fn a_case_records_its_refusal_or_its_success() {
        let fx = "=== program ok\nLog 1.\n=== program refused\nSet total to $5.\n=== phrasebook loads\n(english-vla \"wobble {x:expr}\" (log {x}))\n=== phrasebook bad\nhello\n";
        let got = refusal_report(fx, PRELUDE, ENGLISH);
        let lines: Vec<&str> = got.lines().collect();
        assert_eq!(lines[0], "=== program ok");
        assert_eq!(lines[1], "TRANSLATED");
        assert!(
            lines[3].starts_with("REFUSED\t1\tenglish-unknown-character\t"),
            "{}",
            lines[3]
        );
        assert_eq!(lines[5], "LOADED\t1 rules");
        assert!(
            lines[7].starts_with("REFUSED\t-\tenglish-vocab-expected-directive\t"),
            "{}",
            lines[7]
        );
    }

    /// scripts/refusals_golden.txt is written by VlaWriteRefusalGolden
    /// (VLA_Tests.bas) from the reference; this is the English modules'
    /// refusal oracle, compared as the treaty compares a golden (LF, trailing
    /// blank lines dropped).
    #[test]
    fn the_refusal_golden_is_reproduced() {
        let norm = |s: &str| s.replace("\r\n", "\n").trim_end_matches('\n').to_string();
        let got = norm(&refusal_report(FIXTURE, PRELUDE, ENGLISH));
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
            let case = want[..at]
                .rsplit('\n')
                .find(|l| l.starts_with("=== "))
                .unwrap_or("");
            panic!(
                "the refusal golden differs at char {at} (line {line}, under '{case}'):\n  want: {}\n  got:  {}",
                context(&want),
                context(&got)
            );
        }
    }

    /// Prints the fixture's report between two sentinel lines, for writing
    /// the prediction of the golden before the reference has written it:
    /// `cargo test -p frazaro-core print_refusal_report -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn print_refusal_report() {
        println!("<<<REFUSAL REPORT");
        print!("{}", refusal_report(FIXTURE, PRELUDE, ENGLISH));
        println!(">>>");
    }
}
