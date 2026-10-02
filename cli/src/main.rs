//! frazaro: the command-line door to frazaro-core.
//!
//! The commands this door will carry, in the order the slices of
//! docs/HORIZON.md section 12 deliver them: compile (PORT.5), check,
//! build, run, ask, diff, rebuild, prove. The door reads the files; the
//! core sees text and gives text back. The conformance contract is
//! conformance/README.md: an oracle this version does not attempt exits 3,
//! a refusal writes its words to stderr and exits 1.

use std::process::ExitCode;

const USAGE: &str = "\
frazaro - a compiler from sentences to spreadsheets

usage:
  frazaro compile <program.vla> --prelude <prelude.vla>
                         the VBA of a VLA program, to stdout (PORT.5)
  frazaro load <phrasebook.vla> --prelude <prelude.vla> [--allow-raw]
                         load a phrasebook and report what it holds, the
                         line EnglishVocabStats prints (PORT.6, slice 6c);
                         --allow-raw is the consent a (raw ...) form needs
  frazaro version        the version of the core this door is built on
  frazaro help           this text

Not in this version (docs/HORIZON.md, section 12, slice by slice):
  translate-vla, translate-vba, prove (exit 3: not attempted),
  check, build, run, ask, diff, rebuild
";

/// A file as text: UTF-8, a leading byte-order mark dropped, line endings
/// as they are on disk (the core and the reference both leave them alone).
fn read_text(path: &str) -> Result<String, String> {
    let bytes = std::fs::read(path).map_err(|e| format!("cannot read {path}: {e}"))?;
    let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(&bytes);
    String::from_utf8(bytes.to_vec()).map_err(|_| format!("{path} is not valid UTF-8"))
}

fn option_after<'a>(args: &'a [String], name: &str) -> Option<&'a str> {
    args.iter()
        .position(|a| a == name)
        .and_then(|i| args.get(i + 1))
        .map(String::as_str)
}

fn compile(args: &[String]) -> ExitCode {
    let (Some(program), Some(prelude)) = (args.first(), option_after(args, "--prelude")) else {
        eprintln!("usage: frazaro compile <program.vla> --prelude <prelude.vla>");
        return ExitCode::from(2);
    };
    let (source, prelude) = match (read_text(program), read_text(prelude)) {
        (Ok(s), Ok(p)) => (s, p),
        (Err(e), _) | (_, Err(e)) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    match frazaro_core::emit::compile(&source, &prelude) {
        Ok(vba) => {
            use std::io::Write;
            let mut out = std::io::stdout().lock();
            if out
                .write_all(vba.as_bytes())
                .and_then(|_| out.flush())
                .is_err()
            {
                return ExitCode::from(2);
            }
            ExitCode::SUCCESS
        }
        Err(refusal) => {
            eprintln!("{refusal}");
            ExitCode::from(1)
        }
    }
}

/// `frazaro load`: a phrasebook through the core's loader, with the two
/// gates the reference's file loader runs and its text loader leaves to
/// the door: a required capability (none can be granted yet), and SEC.2's
/// consent for a `(raw ...)` form, which the person gives with --allow-raw.
fn load(args: &[String]) -> ExitCode {
    let (Some(book), Some(prelude)) = (args.first(), option_after(args, "--prelude")) else {
        eprintln!("usage: frazaro load <phrasebook.vla> --prelude <prelude.vla> [--allow-raw]");
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    let (text, prelude) = match (read_text(book), read_text(prelude)) {
        (Ok(t), Ok(p)) => (t, p),
        (Err(e), _) | (_, Err(e)) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    use frazaro_core::english::vocab::{vocab_requires_check_capability, vocab_text_has_raw_form};
    if let Err(refusal) = vocab_requires_check_capability(&text, book) {
        eprintln!("{refusal}");
        return ExitCode::from(1);
    }
    if !allow_raw && vocab_text_has_raw_form(&text) {
        let refusal = frazaro_core::messages::raise(
            "english-vocab-raw-consent-declined",
            &[("source", book)],
        );
        eprintln!("{refusal}");
        return ExitCode::from(1);
    }
    let mut grammar = frazaro_core::english::Grammar::new(&prelude);
    match grammar.load_vocabulary_text(&text, book) {
        Ok(_rules) => {
            // The reference's line counts the proofs it ran; the core
            // collects them until `prove` (slice 6d) runs them, so the same
            // numbers are shown as collected.
            let proofs = grammar.pending_proofs();
            let fails = proofs
                .iter()
                .filter(|p| p.kind == frazaro_core::english::grammar::ProofKind::Fail)
                .count();
            let tests = proofs.len() - fails;
            let plural = |n: usize| if n == 1 { "" } else { "s" };
            let rules = grammar.rule_count() - grammar.prelude_count();
            let macros = grammar.vocab_macro_count();
            println!(
                "loaded: {rules} rule{}, {macros} macro{}, {tests} test{} ({fails} expected fail{}) from {book}",
                plural(rules),
                plural(macros),
                plural(tests),
                plural(fails)
            );
            println!("(proofs collected, not run: frazaro prove is slice 6d)");
            let warnings = grammar.lint_warnings();
            if !warnings.is_empty() {
                print!("{}", grammar.lint_report());
            }
            ExitCode::SUCCESS
        }
        Err(refusal) => {
            eprintln!("{refusal}");
            ExitCode::from(1)
        }
    }
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("compile") => compile(&args[1..]),
        Some("load") => load(&args[1..]),
        Some("translate-vla") | Some("translate-vba") | Some("prove") => {
            eprintln!(
                "frazaro: '{}' is not attempted in this version (PORT.6 and later)",
                args[0]
            );
            ExitCode::from(3)
        }
        Some("version") | Some("--version") | Some("-V") => {
            let core = frazaro_core::version();
            let abi = frazaro_core::frazaro_abi_version();
            println!("frazaro {core} (core abi {abi})");
            ExitCode::SUCCESS
        }
        None | Some("help") | Some("--help") | Some("-h") => {
            print!("{USAGE}");
            ExitCode::SUCCESS
        }
        Some(other) => {
            eprintln!(
                "frazaro: '{other}' is not a command in this version; \
                 'frazaro help' lists the commands."
            );
            ExitCode::from(2)
        }
    }
}
