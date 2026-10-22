//! The appearance golden (KERNEL.50): what Excel reads back of every case
//! of `testdata/appearance.txt`, a copy of `scripts/appearance.txt`, held
//! byte for byte by tools/check_crate_data.ps1, each case a format as the
//! record spells it after its index. `testdata/appearance_golden.txt` is what
//! `VlaWriteAppearanceGolden` (VLA_Tests.bas) writes in Excel: under each
//! case's header, `EXCEL` and Excel's reading of a cell it gave the case
//! through its object model, and `WRITTEN` and its reading of the cell
//! frazaro-core's writer gave it in `scripts/appearance_written.xlsx`, each
//! in the record's spelling with every property at Excel's default left out.
//!
//! 1. Every case is spelled as the record prints it, or is refused as its
//!    header says.
//! 2. Excel reads back of the written cell exactly what the case says, its
//!    defaults left out, and of the cell it gave the case through its object
//!    model the same; a case the language refuses Excel's object model
//!    refuses too. Its sizes are the language's, 1 to 409.55 points by
//!    twentieths: it takes 409.55, past the 409 its font box allows, and
//!    keeps 10.25 and 9.3 whole.

use vla_lang::calc::library;
use vla_lang::form::Form;
use vla_lang::machine::{Machine, Viewed};
use vla_lang::printer::write_datum;
use vla_lang::reader::read_forms;

const CASES: &str = include_str!("../testdata/appearance.txt");
const GOLDEN: &str = include_str!("../testdata/appearance_golden.txt");

/// A header and the lines under it, blank lines skipped, as the golden's
/// writer reads them.
struct Case {
    header: String,
    lines: Vec<String>,
}

fn cases(text: &str) -> Vec<Case> {
    let mut out: Vec<Case> = Vec::new();
    for line in text.replace("\r\n", "\n").lines() {
        if let Some(header) = line.strip_prefix("=== ") {
            out.push(Case {
                header: header.to_string(),
                lines: Vec::new(),
            });
        } else if let (Some(case), false) = (out.last_mut(), line.is_empty()) {
            case.lines.push(line.to_string());
        }
    }
    out
}

/// A case's format as the record prints it after its index, or the refusal's
/// id and words.
fn printed(line: &str) -> Result<String, (String, String)> {
    let rows = format!(
        "(sheet \"S\" visible)\n(format 1 {line})\n(cell \"S\" \"A1\" \"Aa 123\")\n(style \"S\" \"A1\" 1)\n"
    );
    match Machine::load(&rows, library::language()) {
        Err(e) => Err((e.refusal.id, e.refusal.text)),
        Ok(m) => {
            let record = match m.view("grid", "S", None) {
                Ok(Viewed::Text(t)) => t,
                other => panic!("{line}: {other:?}"),
            };
            let row = record
                .lines()
                .find(|l| l.starts_with("(format 1 "))
                .unwrap_or_else(|| panic!("{line}: no format row\n{record}"));
            Ok(row["(format 1 ".len()..row.len() - 1].to_string())
        }
    }
}

/// A format as Excel reads a cell's back: every property at Excel's default
/// left out, a weight or a slant not taken, the size 11, a general or a
/// bottom alignment and an edge of none, and a group left empty with them.
fn as_excel_reads(text: &str) -> String {
    let forms = read_forms(text).expect("the case's forms");
    let mut out: Vec<String> = forms[..3].iter().map(write_datum).collect();
    for group in &forms[3..] {
        let Form::List(l) = group else {
            panic!("{text}: a group that is no list")
        };
        let head = match l.items.first() {
            Some(Form::Sym(h)) => h.as_str(),
            _ => panic!("{text}: a group with no head"),
        };
        let kept: Vec<String> = l.items[1..]
            .iter()
            .filter(|item| match (head, item) {
                ("font", Form::Sym(w)) => !matches!(w.as_str(), "not-bold" | "not-italic" | "11"),
                ("align", Form::Sym(w)) => !matches!(w.as_str(), "general" | "bottom"),
                ("border", Form::List(edge)) => {
                    !matches!(edge.items.get(1), Some(Form::Sym(w)) if w == "none")
                }
                _ => true,
            })
            .map(write_datum)
            .collect();
        if !kept.is_empty() {
            out.push(format!("({head} {})", kept.join(" ")));
        }
    }
    out.join(" ")
}

#[test]
fn every_case_is_spelled_as_the_record_prints_it_or_refused_as_its_header_says() {
    let all = cases(CASES);
    assert!(all.len() >= 46, "only {} cases", all.len());
    for case in &all {
        assert_eq!(case.lines.len(), 1, "{}: one line a case", case.header);
        let line = &case.lines[0];
        match case.header.split_once(" !refused ") {
            Some((_, words)) => {
                let (id, text) = printed(line).expect_err(&case.header);
                assert_eq!(id, "grid-format-property", "{}", case.header);
                for w in words.split_whitespace() {
                    assert!(text.contains(w), "{}: {text}", case.header);
                }
            }
            None => assert_eq!(
                printed(line).as_deref(),
                Ok(line.as_str()),
                "{}",
                case.header
            ),
        }
    }
}

#[test]
fn excel_reads_back_what_each_case_says() {
    let fixture = cases(CASES);
    let golden = cases(GOLDEN);
    assert_eq!(fixture.len(), golden.len(), "the golden's cases");
    for (case, read) in fixture.iter().zip(&golden) {
        assert_eq!(
            case.header, read.header,
            "the golden's headers are the fixture's"
        );
        let [excel, written] = read.lines.as_slice() else {
            panic!("{}: two lines under each header", read.header)
        };
        let excel = excel.strip_prefix("EXCEL\t").expect("EXCEL<TAB>");
        let written = written.strip_prefix("WRITTEN\t").expect("WRITTEN<TAB>");
        let line = &case.lines[0];
        if case.header.contains(" !refused ") {
            assert_eq!(written, "REFUSED", "{}", case.header);
            assert!(
                excel.starts_with("REFUSED "),
                "{}: Excel took it: {excel}",
                case.header
            );
            continue;
        }
        let said = as_excel_reads(line);
        assert_eq!(written, said, "{}: the written cell", case.header);
        assert_eq!(excel, said, "{}: Excel's own cell", case.header);
    }
}
