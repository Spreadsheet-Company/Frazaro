//! A macro that walks a list costs its length (VLA.2). The reference's scale
//! pins walk a table by recursing on `(cdr spec)` (`TestTablespecDepthSafety`
//! and `TestListopsDepthSafety`, VLA_Tests.bas), and VLA.bas shares the rest
//! of the table at each step. The port copied it at every use of the
//! parameter, so a walk of N rows wrote about N squared rows over: the table
//! generator's 2,000 rows made 290 million allocations and took 10.6 s, and
//! its second thousand rows made three times the first thousand's. The
//! expander now shares what the reference shares. This file counts and does
//! not time, so that it cannot flake: the allocations the expander makes on
//! this thread and the bytes they ask for, the second thousand rows of a walk
//! held to the first thousand's, and every row to a ceiling. The bytes catch
//! what a count alone would miss, a step that copies the rest of the list's
//! references into one new array: one allocation, of a size that grows.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use vla_lang::expand::Expander;
use vla_lang::reader::{count_lf, parse_all, tokenize};

/// The system's allocator, counting what this thread asks of it while
/// `COUNTING` is on.
struct Counting;

thread_local! {
    static COUNTING: Cell<bool> = const { Cell::new(false) };
    static COUNT: Cell<u64> = const { Cell::new(0) };
    static BYTES: Cell<u64> = const { Cell::new(0) };
}

fn tally(bytes: usize) {
    let _ = COUNTING.try_with(|on| {
        if on.get() {
            let _ = COUNT.try_with(|n| n.set(n.get() + 1));
            let _ = BYTES.try_with(|b| b.set(b.get() + bytes as u64));
        }
    });
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        tally(layout.size());
        System.alloc(layout)
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        tally(layout.size());
        System.alloc_zeroed(layout)
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        tally(new_size);
        System.realloc(ptr, layout, new_size)
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        System.dealloc(ptr, layout)
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const PRELUDE: &str = include_str!("../testdata/prelude.vla");

/// `TestTablespecDepthSafety`'s generator over `n` rows: each row's name,
/// docstring, body, pattern, call, test sentence and test call, read off it
/// by `car`, `cadr`, `caddr` and `cddr`, nine uses of the table a row.
fn table(n: usize) -> String {
    let generator = "(defmacro (tablespec-row name doc body pattern call testsentence testcall) \
        (begin (defmacro (name n) doc body) (english-vla pattern call) \
        (test-success testsentence testcall))) \
        (defmacro (tablespec spec) (quote-if (null? spec) (begin) \
        (begin (tablespec-row (car (car spec)) (cadr (car spec)) (car (cddr (car spec))) \
        (cadr (cddr (car spec))) (caddr (cddr (car spec))) \
        (cadr (cddr (cddr (car spec)))) (caddr (cddr (cddr (car spec))))) \
        (tablespec (cdr spec)))))";
    let rows: Vec<String> = (1..=n)
        .map(|i| format!("(row{i} \"d\" (debug-print {i}) \"p\" (c) \"s\" (c))"))
        .collect();
    format!("{generator} (tablespec (quote ({})))", rows.join(" "))
}

/// `TestListopsDepthSafety`'s walk over `n` numbers, one statement a row.
fn walk_rows(n: usize) -> String {
    let rows: Vec<String> = (1..=n).map(|i| i.to_string()).collect();
    format!(
        "(defmacro (walk-rows lst) (quote-if (null? lst) (begin) \
         (begin (debug-print (car lst)) (walk-rows (cdr lst)))))\r\n\
         (sub t () (walk-rows (quote ({}))))",
        rows.join(" ")
    )
}

/// The allocations the expander makes on this thread, and the bytes they
/// ask for, defining the prelude's macros and the source's and expanding
/// the rest to its fixpoint, as `VlaExpandText` does; the text is read
/// beforehand and not counted.
fn spent(source: &str) -> (u64, u64) {
    let text = format!("{PRELUDE}\r\n{source}");
    let forms = parse_all(&tokenize(&text), count_lf(PRELUDE) as u32 + 1).expect("it reads");
    COUNT.with(|n| n.set(0));
    BYTES.with(|b| b.set(0));
    COUNTING.with(|on| on.set(true));
    let mut ex = Expander::new();
    let body = ex.collect(forms).expect("its macros are defined");
    for f in &body {
        ex.expand(f, 0).expect("it expands");
    }
    COUNTING.with(|on| on.set(false));
    (COUNT.with(Cell::get), BYTES.with(Cell::get))
}

#[test]
fn a_walk_over_a_table_costs_its_rows_not_their_square() {
    // Measured when VLA.2 was built, in a debug build and a release build
    // alike: the generator's first thousand rows 203,038 allocations of
    // 7,936,337 bytes and its second 203,002 of 7,936,456; walk-rows'
    // 42,020 of 1,351,193 and 42,002 of 1,351,456. Before it, the
    // generator's were 73,236,060 allocations of 2.5 GB and 217,236,002 of
    // 7.5 GB. The ceilings are half as much again as a row's allocations.
    for (what, program, ceiling) in [
        ("the table generator", table as fn(usize) -> String, 305u64),
        ("walk-rows", walk_rows, 63),
    ] {
        let none = spent(&program(0));
        let one = spent(&program(1000));
        let two = spent(&program(2000));
        let (first, second) = (
            (one.0 - none.0, one.1 - none.1),
            (two.0 - one.0, two.1 - one.1),
        );
        assert!(
            second.0 * 20 <= first.0 * 21,
            "{what}: the second thousand rows made {} allocations and the first {}: \
             a walk that costs more the longer it runs",
            second.0,
            first.0
        );
        assert!(
            second.1 * 20 <= first.1 * 21,
            "{what}: the second thousand rows asked for {} bytes and the first {}: \
             a walk that copies more the longer it runs",
            second.1,
            first.1
        );
        assert!(
            first.0 + second.0 <= ceiling * 2000,
            "{what}: 2,000 rows made {} allocations, more than {ceiling} a row",
            first.0 + second.0
        );
    }
}
