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
use vla_lang::calc::{Computed, ErrorKind, Value};
use vla_lang::expand::{expand_text, FUNCTION_DEPTH_LIMIT, FUNCTION_FUEL};
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
        // stepped to the end of a frame, its census printed and dropped:
        // evaluation recurses once a level of nesting and goes round a loop
        // along a chain, the census's walk of a parse keeps its own stack
        // (KERNEL.30), and the chain's tree of 4,095 links is freed
        // recursively.
        for f in [
            nest(NESTING_LIMIT, "(", "1", ")"),
            nest(NESTING_LIMIT, "ABS(", "1", ")"),
            nest(NESTING_LIMIT, "-", "1", ""),
            nest(NESTING_LIMIT, "IF(TRUE,", "1", ",0)"),
            format!("{}11", "1+".repeat(4095)),
        ] {
            let mut m = Machine::load(&cell(&f), library::language()).expect("loads");
            let step = m.step(usize::MAX);
            assert!(step.done && step.evaluated == 1, "{step:?}");
            assert!(m
                .census()
                .starts_with("(plan (shapes 1) (texts 1) (formulas 1) "));
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

#[test]
fn a_recursion_to_its_depth_stands_on_the_same_stack_around_and_inside_the_deepest_formulas() {
    // KERNEL.29: a call too deep for the stack where it stands is computed on
    // a fresh one and remembered, so that a recursion as deep as its limit,
    // called from a cell nested as deep as a formula may be and with a body
    // as deep, stands on a thread of 1 MiB; one call deeper is `#NUM!`.
    on_one_mib(|| {
        let around = |inner: &str| nest(NESTING_LIMIT - 1, "ABS(", inner, ")");
        let body = format!(
            "=LAMBDA(n,IF(n=0,0,1+{}))",
            nest(NESTING_LIMIT - 6, "ABS(", "CNT(n-1)", ")")
        );
        for (call, want) in [
            ("CNT(4095)", Value::Number(4095.0)),
            ("CNT(4096)", Value::Error(ErrorKind::Num)),
        ] {
            let rows = format!(
                "(name \"CNT\" \"{body}\")\n(formula \"S\" \"A1\" \"={}\")\n",
                around(call)
            );
            let mut m = Machine::load(&rows, library::language()).expect("loads");
            let step = m.step(usize::MAX);
            assert!(step.done, "{step:?}");
            assert_eq!(a1(&m), Computed::Value(want), "{call}");
        }
    });
}

/// `S!A1`'s value in a machine stepped.
fn a1(m: &Machine) -> Computed {
    use vla_lang::calc::Values;
    let s = m.find_sheet("S").expect("the sheet");
    m.computed((s, 1, 1)).expect("a value")
}

#[test]
fn a_chain_of_lets_as_long_as_excel_binds_stands_on_the_same_stack() {
    // KERNEL.29: 126 names, each its predecessor and one, Excel's most, inlined
    // to one chain of additions; and 40 names each its predecessor twice over,
    // whose inlining would double 40 times, kept as written and each computed
    // once.
    on_one_mib(|| {
        // Each name spelled as no cell is, as a binding's must be.
        let mut flat = String::from("LET(v_0,1");
        for k in 1..126 {
            flat.push_str(&format!(",v_{k},v_{}+1", k - 1));
        }
        flat.push_str(",v_125)");
        let mut doubled = String::from("LET(w_0,1");
        for k in 1..40 {
            doubled.push_str(&format!(",w_{k},w_{}+w_{}", k - 1, k - 1));
        }
        doubled.push_str(",w_39)");
        for (f, want) in [(flat, 126.0), (doubled, 549_755_813_888.0)] {
            let mut m = Machine::load(&cell(&f), library::language()).expect("loads");
            let step = m.step(usize::MAX);
            assert!(step.done, "{step:?}");
            assert_eq!(a1(&m), Computed::Value(Value::Number(want)));
        }
    });
}

#[test]
fn a_function_s_calls_nest_to_their_limit_and_loop_past_it() {
    on_one_mib(|| {
        // VLA.8: a function's calls nested to the depth, the call standing at
        // the form depth's own limit, the deepest stack the expander reaches,
        // and one level more refused by name.
        let nest_fn = concat!(
            "(defexpand (nest n) ",
            "(quote-if (=expand n 0) 0 (+expand 1 (nest (-expand n 1)))))\r\n"
        );
        let at = FUNCTION_DEPTH_LIMIT - 1;
        let call = format!("(nest {at})");
        let deep = format!("{nest_fn}{}", nest(FORM_DEPTH_LIMIT - 1, "(f ", &call, ")"));
        let (text, _) = expand_text(&deep, "", true).unwrap();
        assert!(text.contains(&format!(" {at})")));
        let past = format!("{nest_fn}(nest {FUNCTION_DEPTH_LIMIT})");
        assert_eq!(
            expand_text(&past, "", true).unwrap_err().id,
            "vla-function-depth"
        );
        // A call in a tail place loops, its fuel spent to the last call and
        // refused one past it, the stack never growing.
        let down = "(defexpand (down n) (quote-if (=expand n 0) 0 (down (-expand n 1))))\r\n";
        let (text, _) =
            expand_text(&format!("{down}(down {})", FUNCTION_FUEL - 1), "", true).unwrap();
        assert_eq!(text.trim(), "0");
        let spent = format!("{down}(down {FUNCTION_FUEL})");
        assert_eq!(
            expand_text(&spent, "", true).unwrap_err().id,
            "vla-function-fuel"
        );
        // A value shared into a vast tree, refused by its weight where the walk
        // places it, and written as dots in a refusal's trail, never walked.
        let grow = concat!(
            "(defexpand (grow x n) ",
            "(quote-if (=expand n 0) x (grow (list x x) (-expand n 1))))\r\n"
        );
        let vast = expand_text(&format!("{grow}(grow 1 40)"), "", true).unwrap_err();
        assert_eq!(vast.id, "vla-macro-expansion-too-large");
        let trail = concat!(
            "(defexpand (grow x n) ",
            "(quote-if (=expand n 0) (+expand x 1) (grow (list x x) (-expand n 1))))\r\n",
            "(grow 1 40)"
        );
        let refused = expand_text(trail, "", true).unwrap_err();
        assert!(
            refused.text.ends_with("(computing (grow ... 0))"),
            "{}",
            refused.text
        );
    });
}

#[test]
fn a_value_a_working_out_holds_is_held_to_the_reader_s_depth() {
    on_one_mib(|| {
        // VLA.8: a call that wraps its argument nests it a level deeper at
        // each call, within the fuel thousands deep, which no walk reads and
        // whose freeing alone ran this stack out. An argument is held to the
        // reader's limit where it is bound, the deepest standing.
        let deep =
            "(defexpand (deep x n) (quote-if (=expand n 0) 0 (deep (list x) (-expand n 1))))\r\n";
        let at = format!("{deep}(deep 1 {FORM_DEPTH_LIMIT})");
        assert_eq!(expand_text(&at, "", true).unwrap().0.trim(), "0");
        let past = expand_text(&format!("{deep}(deep 1 4999)"), "", true).unwrap_err();
        assert_eq!(past.id, "vla-form-too-deep");
        let left = 4999 - FORM_DEPTH_LIMIT;
        assert!(
            past.text
                .ends_with(&format!("(computing (deep ... {left}))")),
            "{}",
            past.text
        );
        // The value a call answers, five levels deeper at each call nested.
        let wrap = concat!(
            "(defexpand (wrap n) (quote-if (=expand n 0) 0 ",
            "(list (list (list (list (list (wrap (-expand n 1)))))))))\r\n"
        );
        let held = expand_text(&format!("{wrap}(car (wrap 12))"), "", true);
        assert!(held.is_ok(), "{held:?}");
        let answered = expand_text(&format!("{wrap}(car (wrap 13))"), "", true).unwrap_err();
        assert_eq!(answered.id, "vla-form-too-deep");
        assert!(
            answered.text.ends_with("(computing (wrap 13))"),
            "{}",
            answered.text
        );
        // A let-expand's name, each binding the last one wrapped.
        let names = |n: u32| {
            let mut s = String::from("(let-expand ((a0 1)");
            for k in 1..=n {
                s.push_str(&format!(" (a{k} (list a{}))", k - 1));
            }
            s.push_str(") 0)");
            s
        };
        assert!(expand_text(&names(FORM_DEPTH_LIMIT), "", true).is_ok());
        let bound = expand_text(&names(FORM_DEPTH_LIMIT + 1), "", true).unwrap_err();
        assert_eq!(bound.id, "vla-form-too-deep");
    });
}
