//! The expand golden (VLA.3): the expander held to the reference's
//! expansion of one fixture, case for case.
//!
//! `testdata/expand.txt`, a copy of `scripts/expand.txt` held byte for byte
//! by `check_crate_data`, holds cases under `=== name`, each one VLA text.
//! `testdata/expand_golden.txt` is what `VlaWriteExpandGolden`
//! (`VLA_Tests.bas`) writes of them: each case's header, then
//! `EXPANDED<TAB><applications fired>` and the text `VlaExpandText` prints,
//! or `REFUSED<TAB><id><TAB><text>`. This writes the same records from the
//! port and holds them to that file whole, so that the twenty-seven
//! expand-time primitives and the macro system around them answer one text
//! the same way in both implementations (SD-18: the reference's file is the
//! one that decides).

use vla_lang::expand::expand_text;

const PRELUDE: &str = include_str!("../testdata/prelude.vla");
const CASES: &str = include_str!("../testdata/expand.txt");
const GOLDEN: &str = include_str!("../testdata/expand_golden.txt");

/// `ExpandReportFor`: the case's text, its trailing breaks dropped,
/// expanded to its fixpoint under the prelude, or its refusal.
fn report_for(body: &str) -> String {
    let body = body.trim_end_matches('\n');
    match expand_text(body, PRELUDE, true) {
        Ok((text, fired)) => {
            let text = text.replace("\r\n", "\n");
            let text = text.strip_suffix('\n').unwrap_or(&text);
            format!("EXPANDED\t{fired}\n{text}\n")
        }
        Err(e) => format!("REFUSED\t{}\t{}\n", e.id, e.text),
    }
}

/// `VlaWriteExpandGolden`'s loop: the lines before the first header are
/// ignored, and a case's lines are joined with a line break.
fn report(cases: &str) -> (String, usize) {
    let mut out = String::new();
    let mut header: Option<&str> = None;
    let mut body = String::new();
    let mut first = true;
    let mut count = 0;
    for ln in cases.split('\n').chain(std::iter::once("=== ")) {
        if ln.starts_with("=== ") {
            if let Some(h) = header {
                out.push_str(h);
                out.push('\n');
                out.push_str(&report_for(&body));
                count += 1;
            }
            header = Some(ln);
            body.clear();
            first = true;
        } else if header.is_some() {
            if !first {
                body.push('\n');
            }
            body.push_str(ln);
            first = false;
        }
    }
    (out, count)
}

#[test]
fn the_expand_golden_is_reproduced() {
    let want = GOLDEN.replace("\r\n", "\n");
    let (got, count) = report(&CASES.replace("\r\n", "\n"));
    assert!(count >= 135, "only {count} cases in the expand fixture");
    if got != want {
        let (g, w): (Vec<&str>, Vec<&str>) = (got.lines().collect(), want.lines().collect());
        let at = (0..g.len().max(w.len()))
            .find(|&i| g.get(i) != w.get(i))
            .unwrap_or(0);
        panic!(
            "the port differs from the expand golden at line {}:\n  port:   {:?}\n  golden: {:?}",
            at + 1,
            g.get(at),
            w.get(at)
        );
    }
}
