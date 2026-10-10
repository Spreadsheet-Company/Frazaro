//! The limits held on the tightest stack measured (VLA.1). Each stage that
//! recurses once a level of what it reads is driven to exactly its limit on
//! a thread of 1 MiB, the main thread's stack on Windows and WebAssembly's
//! default, in the profile `cargo test` builds (the debug one, unless asked
//! otherwise, where a frame is largest), and one level past the limit is
//! refused by name. Before VLA.1 the reader overflowed its stack at 1,000
//! nested lists in a debug build, the formula parser at 80 nested
//! parentheses, and a machine's load at one formula 72 deep; a chain of
//! 4,096 additions, within Excel's length, overflowed evaluation even in a
//! release build. An overflow here is not a failed assertion but the end of
//! the test process, which is the failure these pins exist to prevent.

use vla_lang::calc::formula::{self, LENGTH_LIMIT, NESTING_LIMIT};
use vla_lang::calc::library;
use vla_lang::expand::expand_text;
use vla_lang::machine::Machine;
use vla_lang::reader::{read_forms, FORM_DEPTH_LIMIT};

/// `f` on a thread whose stack is 1 MiB; its panics are the test's.
fn on_one_mib<F: FnOnce() + Send + 'static>(f: F) {
    std::thread::Builder::new()
        .stack_size(1 << 20)
        .spawn(f)
        .expect("a thread")
        .join()
        .expect("no panic on the small stack");
}

/// `inner` inside `n` openings and their closings.
fn nest(n: u32, open: &str, inner: &str, close: &str) -> String {
    let n = n as usize;
    format!("{}{inner}{}", open.repeat(n), close.repeat(n))
}

/// One formula cell's row, as a cartridge or a save holds it.
fn cell(formula: &str) -> String {
    format!("(formula \"S\" \"A1\" \"={formula}\")\n")
}

#[test]
fn the_reader_reads_to_its_limit_and_refuses_past_it() {
    on_one_mib(|| {
        let at = read_forms(&nest(FORM_DEPTH_LIMIT, "(", "", ")")).unwrap();
        assert_eq!(at.len(), 1);
        let past = read_forms(&nest(FORM_DEPTH_LIMIT + 1, "(", "", ")")).unwrap_err();
        assert_eq!(past.id, "vla-form-too-deep");
    });
}

#[test]
fn the_expander_walks_and_prints_to_its_limit_and_refuses_past_it() {
    on_one_mib(|| {
        // A text at the limit with nothing to expand, walked and written
        // pretty whole.
        let (text, fired) =
            expand_text(&nest(FORM_DEPTH_LIMIT, "(f ", "1", ")"), "", true).unwrap();
        assert_eq!(fired, 0);
        assert_eq!(text.matches("(f").count(), FORM_DEPTH_LIMIT as usize);
        // A ten-deep macro whose call sits at the depth that brings its
        // expansion to the limit, and one level lower, past it.
        let deep10 = "(defmacro (deep10) (a (a (a (a (a (a (a (a (a (a 1)))))))))))\r\n";
        let around = |n: u32| format!("{deep10}{}", nest(n, "(", "(deep10)", ")"));
        assert!(expand_text(&around(FORM_DEPTH_LIMIT - 10), "", true).is_ok());
        let past = expand_text(&around(FORM_DEPTH_LIMIT - 9), "", true).unwrap_err();
        assert_eq!(past.id, "vla-form-too-deep");
        // A macro that deepens its argument at every application, which
        // overflowed even a release build's stack, and one that doubles it,
        // which ran until the memory gave out.
        let deeper = expand_text("(defmacro (deeper x) (deeper (x))) (deeper 1)", "", true);
        assert_eq!(deeper.unwrap_err().id, "vla-form-too-deep");
        let dup = expand_text("(defmacro (dup x) (dup (x x))) (dup 1)", "", true);
        assert_eq!(dup.unwrap_err().id, "vla-macro-expansion-too-large");
    });
}

#[test]
fn the_formula_parser_reads_to_excel_s_limits_and_refuses_past_them() {
    on_one_mib(|| {
        // Parentheses, a call's arguments, signs and percents, each to 64.
        for (open, close) in [("(", ")"), ("ABS(", ")"), ("-", ""), ("", "%")] {
            let at = format!("={}", nest(NESTING_LIMIT, open, "1", close));
            assert!(formula::parse(&at).is_ok(), "{open}1{close} at the limit");
            let past = format!("={}", nest(NESTING_LIMIT + 1, open, "1", close));
            assert_eq!(formula::parse(&past).unwrap_err(), formula::TOO_DEEP);
        }
        let longest = format!("={}11", "1+".repeat(4095));
        assert_eq!(longest.len() - 1, LENGTH_LIMIT);
        assert!(formula::parse(&longest).is_ok());
        assert_eq!(
            formula::parse(&format!("{longest}1")).unwrap_err(),
            formula::TOO_LONG
        );
    });
}

#[test]
fn a_machine_loads_steps_and_drops_formulas_at_excel_s_limits() {
    on_one_mib(|| {
        // The deepest formulas Excel holds and its longest, each loaded,
        // stepped to the end of a frame and dropped: evaluation recurses
        // once a level of nesting and goes round a loop along a chain, and
        // the chain's tree of 4,095 links is freed recursively.
        for f in [
            nest(NESTING_LIMIT, "(", "1", ")"),
            nest(NESTING_LIMIT, "ABS(", "1", ")"),
            nest(NESTING_LIMIT, "-", "1", ""),
            format!("{}11", "1+".repeat(4095)),
        ] {
            let mut m = Machine::load(&cell(&f), library::language()).expect("loads");
            let step = m.step(usize::MAX);
            assert!(step.done && step.evaluated == 1, "{step:?}");
            drop(m);
        }
        // One past each limit is refused at its line by the limit's own id.
        let deep = cell(&nest(NESTING_LIMIT + 1, "(", "1", ")"));
        let e = Machine::load(&deep, library::language()).unwrap_err();
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (1, "calc-formula-too-deep")
        );
        let long = cell(&format!("{}111", "1+".repeat(4095)));
        let e = Machine::load(&long, library::language()).unwrap_err();
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (1, "calc-formula-too-long")
        );
        // Rows nested past the reader's limit stop the load before any row.
        let rows = nest(FORM_DEPTH_LIMIT + 1, "(", "", ")");
        let e = Machine::load(&rows, library::language()).unwrap_err();
        assert_eq!(e.refusal.id, "vla-form-too-deep");
    });
}
