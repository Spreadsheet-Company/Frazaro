//! The expand golden (VLA.3): the expander held to the reference's
//! expansion of one fixture, case for case.
//!
//! `testdata/expand.txt`, a copy of `scripts/expand.txt` held byte for byte
//! by `check_crate_data`, holds cases under `=== name`, each one VLA text.
//! `testdata/expand_golden.txt` is what `VlaWriteExpandGolden`
//! (`VLA_Tests.bas`) writes of them: each case's header, then
//! `EXPANDED<TAB><applications fired>` and the text `VlaExpandText` prints,
//! or `REFUSED<TAB><id><TAB><text>`. This writes the same records from the
//! port and holds them to that file whole, so that the twenty-nine
//! expand-time primitives, the walk's forms and the macro system around
//! them answer one text the same way in both implementations (SD-18: the
//! reference's file is the one that decides). And it holds the loop's
//! claim in miniature (VLA.5): the kit's `cells` walker and its `each`
//! twin, two cases of the fixture, write the same rows; the define's
//! (VLA.6): a block's labels and names written from one list and the stop
//! rule by names write the rows their twin by addresses writes, Kairos's
//! own stop among them; and VLA.15's: a block's labels and words from
//! entries `list?` tells apart, a word or a pair, a fill citing a word
//! anchored, and `count` at a word through `first-cell-expand`, each fill
//! the row its twin by addresses writes.

use vla_lang::expand::expand_text;
use vla_lang::form::Form;
use vla_lang::reader::read_forms;

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

/// A case's text, by its name in the fixture.
fn case_text(name: &str) -> String {
    let cases = CASES.replace("\r\n", "\n");
    let header = format!("=== {name}\n");
    let at = cases
        .find(&header)
        .unwrap_or_else(|| panic!("no case {name} in the fixture"))
        + header.len();
    let rest = &cases[at..];
    let end = rest.find("\n=== ").map_or(rest.len(), |e| e + 1);
    rest[..end].trim_end_matches('\n').to_string()
}

/// The rows an expansion writes, as Alonzo's pass reads them: every
/// `begin` flattened, an empty one writing nothing.
fn rows(text: &str) -> Vec<Form> {
    fn flatten(f: Form, out: &mut Vec<Form>) {
        match f {
            Form::List(l) if l.head_is("begin") => {
                for item in l.items.into_iter().skip(1) {
                    flatten(item, out);
                }
            }
            other => out.push(other),
        }
    }
    let mut out = Vec::new();
    for f in read_forms(text).expect("an expansion reads back") {
        flatten(f, &mut out);
    }
    out
}

#[test]
fn the_walker_and_its_loop_write_the_same_rows() {
    // The two records cannot be equal: the walker fires an application an
    // entry and one more, and its base case leaves a (begin); the loop
    // fires one. Their rows are what a cartridge holds.
    let (walked, walks) =
        expand_text(&case_text("cells-written-by-a-walk"), PRELUDE, true).unwrap();
    let (looped, loops) = expand_text(&case_text("cells-written-by-each"), PRELUDE, true).unwrap();
    assert_eq!((walks, loops), (4, 1));
    assert_ne!(walked, looped);
    let (walked, looped) = (rows(&walked), rows(&looped));
    assert_eq!(walked.len(), 3);
    assert_eq!(walked, looped);
}

#[test]
fn the_stop_rule_by_names_writes_the_rows_its_twin_by_addresses_writes() {
    // The records differ: the names' case defines its words, an empty begin
    // where each stood, and fires a label's macro a word. Its rows are the
    // addresses' case's, Kairos's stop row last.
    let (named, _) = expand_text(&case_text("the-stop-rule-by-names"), PRELUDE, true).unwrap();
    let (addressed, _) =
        expand_text(&case_text("the-stop-rule-by-addresses"), PRELUDE, true).unwrap();
    assert_ne!(named, addressed);
    let (named, addressed) = (rows(&named), rows(&addressed));
    assert_eq!(named.len(), 8);
    assert_eq!(named, addressed);
    let kairos =
        read_forms("(formula \"State\" \"B12\" \"=IF(AND(State.last!B1=1,State!B9=1),1,0)\")")
            .unwrap();
    assert_eq!(named.last(), kairos.first());
}

#[test]
fn a_block_s_entries_are_told_apart_by_list_and_name_their_cells() {
    // VLA.15: a word whose text is its label, and a pair whose label is no
    // word, written as one block; the define each entry leaves names the
    // value beside its label, read in the form after.
    let (named, _) = expand_text(&case_text("names-from-words-and-pairs"), PRELUDE, true).unwrap();
    let want = read_forms(
        "(cell \"State\" \"A1\" \"phase\") (cell \"State\" \"A2\" \"stopped at\") \"=IF(State!B1=1,State!B2,0)\"",
    )
    .unwrap();
    assert_eq!(rows(&named), want);
}

#[test]
fn a_fill_and_a_count_at_a_word_write_the_rows_their_twins_by_addresses_write() {
    // VLA.15: the records differ by the define's empty begin and the `at`
    // applied; the rows are the same, Kairos's anchored reading and the
    // kit's count at a word through the first cell.
    for (by_words, by_addresses, row) in [
        (
            "a-fill-citing-a-word-anchored",
            "a-fill-citing-addresses-anchored",
            "(formula \"Gate\" \"A1:O1\" \"=IF(State!$B$12=1,State.last!$B$12,0)\")",
        ),
        (
            "a-count-at-a-word",
            "a-count-at-addresses",
            "(formula \"Tally\" \"B1:E10\" \"=Tally.last!B1+Input!B16\")",
        ),
    ] {
        let (named, _) = expand_text(&case_text(by_words), PRELUDE, true).unwrap();
        let (addressed, _) = expand_text(&case_text(by_addresses), PRELUDE, true).unwrap();
        assert_ne!(named, addressed, "{by_words}");
        let (named, addressed) = (rows(&named), rows(&addressed));
        assert_eq!(named, addressed, "{by_words}");
        assert_eq!(named, read_forms(row).unwrap(), "{by_words}");
    }
}

#[test]
fn the_expand_golden_is_reproduced() {
    let want = GOLDEN.replace("\r\n", "\n");
    let (got, count) = report(&CASES.replace("\r\n", "\n"));
    assert!(count >= 362, "only {count} cases in the expand fixture");
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
