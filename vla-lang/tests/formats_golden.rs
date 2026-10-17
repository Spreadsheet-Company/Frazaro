//! The formats golden (KERNEL.47): a number format's code, and the text a
//! value draws as under it, held to Excel's own, case for case.
//!
//! `testdata/formats.txt`, a copy of `scripts/formats.txt` held byte for
//! byte by `check_crate_data`, holds cases under `=== name`, each a code and
//! the values it draws, a number in the record's spelling or a text after
//! an apostrophe. `testdata/formats_golden.txt` is what
//! `VlaWriteFormatsGolden` (`VLA_Tests.bas`) writes of them in Excel: each
//! case's header, then `CODE<TAB><the code as Range.NumberFormat reads it
//! back>` or `REFUSED<TAB><Excel's error number>`, then one line a value,
//! `<the value as written><TAB><Range.Text>`. This holds `vla_lang::numfmt`
//! to that file whole: a code Excel refuses is refused here; a case whose
//! header ends `!refused <words>` is one this version does not draw, its
//! refusal naming the words; and every other code is taken, read back by
//! Excel as written or in a spelling of its own that this version takes and
//! draws alike, and draws each value as Excel's `Range.Text` does
//! (SD-18's shape: the reference's file is the one that decides, here
//! Excel's through the reference).

use vla_lang::calc::Value;
use vla_lang::numfmt::Code;

const CASES: &str = include_str!("../testdata/formats.txt");
const GOLDEN: &str = include_str!("../testdata/formats_golden.txt");

/// One case of the fixture: its header, its code, its values as written.
struct Case<'a> {
    header: &'a str,
    code: &'a str,
    values: Vec<&'a str>,
}

/// The fixture's cases, as `VlaWriteFormatsGolden` reads them: the lines
/// before the first header ignored, and blank lines skipped.
fn cases(text: &str) -> Vec<Case<'_>> {
    let mut out: Vec<Case<'_>> = Vec::new();
    for ln in text.lines() {
        if ln.starts_with("=== ") {
            out.push(Case {
                header: ln,
                code: "",
                values: Vec::new(),
            });
        } else if let Some(case) = out.last_mut() {
            if ln.is_empty() {
                continue;
            }
            if case.code.is_empty() {
                case.code = ln;
            } else {
                case.values.push(ln);
            }
        }
    }
    out
}

/// One case of the golden: its header, Excel's word on its code, and the
/// text Excel draws each value as.
struct Drawn<'a> {
    header: &'a str,
    code: &'a str,
    texts: Vec<(&'a str, &'a str)>,
}

fn golden(text: &str) -> Vec<Drawn<'_>> {
    let mut out: Vec<Drawn<'_>> = Vec::new();
    for ln in text.lines() {
        if ln.starts_with("=== ") {
            out.push(Drawn {
                header: ln,
                code: "",
                texts: Vec::new(),
            });
        } else if let Some(case) = out.last_mut() {
            if case.code.is_empty() {
                case.code = ln;
            } else {
                let (value, drawn) = ln.split_once('\t').unwrap_or((ln, ""));
                case.texts.push((value, drawn));
            }
        }
    }
    out
}

/// A value as the fixture writes it.
fn value(written: &str) -> Value {
    match written.strip_prefix('\'') {
        Some(text) => Value::Text(text.to_string()),
        None => Value::Number(
            written
                .parse()
                .unwrap_or_else(|_| panic!("{written} is no number")),
        ),
    }
}

#[test]
fn every_code_draws_each_of_its_values_as_excel_draws_it() {
    let cases = cases(CASES);
    let golden = golden(GOLDEN);
    assert_eq!(
        cases.iter().map(|c| c.header).collect::<Vec<_>>(),
        golden.iter().map(|g| g.header).collect::<Vec<_>>(),
        "the golden's cases are the fixture's, in its order"
    );
    let mut drawn = 0;
    let mut respellings = 0;
    for (case, excel) in cases.iter().zip(&golden) {
        let parsed = Code::parse(case.code);
        if let Some((_, words)) = case.header.split_once(" !refused ") {
            let u = parsed.expect_err(case.header);
            assert!(
                u.what.contains(words),
                "{}: the refusal names {:?}, not {words}",
                case.header,
                u.what
            );
            assert!(case.values.is_empty(), "{}: a value", case.header);
            continue;
        }
        if excel.code.starts_with("REFUSED\t") {
            panic!("{}: Excel refuses a code this version draws", case.header);
        }
        let read_back = excel.code.strip_prefix("CODE\t").unwrap_or_else(|| {
            panic!(
                "{}: a record neither CODE nor REFUSED: {}",
                case.header, excel.code
            )
        });
        let code = parsed.unwrap_or_else(|u| panic!("{}: {u:?}", case.header));
        // A code Excel reads back spelled otherwise, as it reads `\€` back
        // as `€`, is taken in Excel's spelling too, and draws every value
        // alike.
        let respelled = (read_back != case.code).then(|| {
            Code::parse(read_back).unwrap_or_else(|u| {
                panic!(
                    "{}: Excel's spelling {read_back} is refused: {u:?}",
                    case.header
                )
            })
        });
        respellings += usize::from(respelled.is_some());
        assert_eq!(
            case.values.len(),
            excel.texts.len(),
            "{}: one text a value",
            case.header
        );
        for (written, (echoed, text)) in case.values.iter().zip(&excel.texts) {
            assert_eq!(written, echoed, "{}: the golden's value", case.header);
            let v = value(written);
            // A text under a code with no text's section draws as itself.
            let own = match &v {
                Value::Text(t) => t.clone(),
                _ => String::new(),
            };
            let got = code.text(&v).unwrap_or(own.clone());
            assert_eq!(&got, text, "{}: {written} under {}", case.header, case.code);
            if let Some(r) = &respelled {
                let again = r.text(&v).unwrap_or(own);
                assert_eq!(
                    &again, text,
                    "{}: {written} under Excel's spelling {read_back}",
                    case.header
                );
            }
            drawn += 1;
        }
    }
    // Counted 2026-10-17: 114 values drawn, a floor the fixture only raises;
    // Excel's first run read one drawn code back spelled otherwise, `\€`.
    assert!(drawn >= 114, "too few values drawn: {drawn}");
    assert!(respellings <= 1, "Excel respelled {respellings} codes");
}
