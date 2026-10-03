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
                         load a phrasebook, its proofs run as the add-in
                         runs them, and report what it holds, the line
                         EnglishVocabStats prints (PORT.6, slice 6c);
                         --allow-raw is the consent a (raw ...) form needs
  frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
                         the treaty's oracle 3: every proof run, each
                         failure printed, then PASS n/n or FAIL k/n (k of
                         n passed), exit 0 or 1 (PORT.6, slice 6d); the
                         prelude is prelude.vla beside the file or in its
                         parent folder unless --prelude names one; an
                         engine proof file exits 3 (not attempted, PORT.9)
  frazaro version        the version of the core this door is built on
  frazaro help           this text

Not in this version (docs/HORIZON.md, section 12, slice by slice):
  translate-vla, translate-vba (exit 3: not attempted),
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

/// A phrasebook holds a `<lingua>-vla` rule or a proof form at the top
/// level; the same lexical rule tools/prove.ps1 inventories by (a file with
/// neither is a library, or an engine proof file).
fn is_phrasebook_text(text: &str) -> bool {
    text.lines().any(|l| {
        if l.starts_with("(test-success") || l.starts_with("(test-fail") {
            return true;
        }
        let Some(rest) = l.strip_prefix('(') else {
            return false;
        };
        let head: String = rest
            .chars()
            .take_while(|c| !c.is_whitespace() && *c != '(' && *c != ')')
            .collect();
        head.ends_with("-vla") || head.ends_with("-vla-override")
    })
}

/// `prelude.vla` beside a file, or in its parent folder, as a path string.
fn prelude_beside(file: &str) -> Option<String> {
    let dir = std::path::Path::new(file)
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .map(std::path::Path::to_path_buf)
        .unwrap_or_else(|| std::path::PathBuf::from("."));
    for candidate in [dir.join("prelude.vla"), dir.join("..").join("prelude.vla")] {
        if candidate.is_file() {
            return Some(candidate.to_string_lossy().into_owned());
        }
    }
    None
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
            // `EnglishVocabStats`' line: the proofs ran at load, as the
            // reference runs them, and a failing one refused the load above.
            println!("{}", grammar.vocab_stats());
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

/// `frazaro prove <phrasebook.vla> --prelude <prelude.vla> [--allow-raw]`:
/// the treaty's oracle 3. The phrasebook loads with every proof run and
/// every failure kept, each failure is printed as the reference would have
/// refused it, and the last line is `PASS n/n` or `FAIL k/n` (k passed of
/// n), exit 0 or 1. A load that refuses before its proofs (a library, a
/// malformed rule) prints the refusal and exits 1 with no verdict line.
fn prove(args: &[String]) -> ExitCode {
    let Some(book) = args.first() else {
        eprintln!("usage: frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]");
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    let text = match read_text(book) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    // The treaty's oracle 3 is a phrasebook's proofs; an engine proof file
    // (scripts/proofs/) is PORT.9's, not attempted, which the runner
    // reads from exit 3.
    if !is_phrasebook_text(&text) {
        eprintln!(
            "frazaro: 'prove' attempts a phrasebook's proofs; engine proofs are not attempted in this version (PORT.9)"
        );
        return ExitCode::from(3);
    }
    // The contract is `prove <proofs.vla>` alone, so the prelude is found
    // beside the file or in its parent folder (scripts/polyglotta/x.vla
    // reads scripts/prelude.vla); --prelude names another.
    let prelude_path = match option_after(args, "--prelude") {
        Some(p) => p.to_string(),
        None => match prelude_beside(book) {
            Some(p) => p,
            None => {
                eprintln!(
                    "frazaro: no prelude.vla beside {book} or in its parent folder; pass --prelude <prelude.vla>"
                );
                return ExitCode::from(2);
            }
        },
    };
    let prelude = match read_text(&prelude_path) {
        Ok(p) => p,
        Err(e) => {
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
    grammar.set_proof_mode(frazaro_core::english::grammar::ProofMode::Collect);
    match grammar.load_vocabulary_text(&text, book) {
        Ok(_rules) => {
            let (tests, fails) = grammar.proofs_run();
            let n = tests + fails;
            let failures = grammar.proof_failures();
            for f in failures {
                println!("{}", f.refusal.text);
                println!();
            }
            let k = n - failures.len() as u64;
            if failures.is_empty() {
                println!("PASS {n}/{n}");
                ExitCode::SUCCESS
            } else {
                println!("FAIL {k}/{n}");
                ExitCode::from(1)
            }
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
        Some("prove") => prove(&args[1..]),
        Some("translate-vla") | Some("translate-vba") => {
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
