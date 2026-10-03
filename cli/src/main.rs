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
  frazaro translate-vla <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
                         the treaty's oracle 1: the program's VLA, the text
                         EnglishToVla writes, to stdout (PORT.6, slice 6e);
                         each phrasebook loads in order, its proofs run
  frazaro translate-vba <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
                         oracle 1a: that VLA compiled, the VBA the add-in
                         would write (EnglishToVba)
  frazaro build <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] --out <file.xlsx> [--replace] [--allow-raw]
                         the writer (PORT.7): the program translated as
                         translate-vla translates it, then written as a
                         workbook: the Frazaro sheet holds the sentences in
                         column B with OK beside each in column C, as the
                         add-in's Check marks them (slice 7a), and the sheets
                         the sentences describe hold what a sheet can hold
                         with nothing running - values and formulas into
                         cells, on the sheet the program names (slice 7b); a
                         sentence that needs the add-in's Run is refused by
                         name, and a file already at --out is refused unless
                         --replace is given; the workbook carries its build
                         stamp, the defined name Frazaro.Build (slice 7c)
  frazaro rebuild <file.xlsx> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]
                         the sentences read back out of a built workbook's
                         Frazaro sheet, its stamp checked against them and
                         the files given, the build done again and compared
                         whole: one line, yes (exit 0) or no with why (exit
                         1); a file that is not a build is refused (slice 7c)
  frazaro version        the version of the core this door is built on
  frazaro help           this text

Not in this version (docs/HORIZON.md, section 12, slice by slice):
  check, run, ask, diff
";

/// A file as text: UTF-8, a leading byte-order mark dropped, line endings
/// as they are on disk (the core and the reference both leave them alone).
fn read_text(path: &str) -> Result<String, String> {
    let bytes = std::fs::read(path).map_err(|e| format!("cannot read {path}: {e}"))?;
    let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(&bytes);
    String::from_utf8(bytes.to_vec()).map_err(|_| format!("{path} is not valid UTF-8"))
}

/// A file a command needs but cannot find: the catalogue's refusal for that
/// kind of file (the reference's `english-program-file-not-found` and
/// `english-vocab-file-not-found`; a VLA source and the prelude take
/// VLA.bas's two), to stderr with exit 1, as any refusal.
fn refuse_missing(id: &str, path: &str) -> ExitCode {
    eprintln!("{}", frazaro_core::messages::raise(id, &[("path", path)]));
    ExitCode::from(1)
}

fn is_file(path: &str) -> bool {
    std::path::Path::new(path).is_file()
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
    if !is_file(program) {
        return refuse_missing("vla-source-not-found", program);
    }
    if !is_file(prelude) {
        return refuse_missing("vla-file-not-found", prelude);
    }
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
    if !is_file(book) {
        return refuse_missing("english-vocab-file-not-found", book);
    }
    if !is_file(prelude) {
        return refuse_missing("vla-file-not-found", prelude);
    }
    let (text, prelude) = match (read_text(book), read_text(prelude)) {
        (Ok(t), Ok(p)) => (t, p),
        (Err(e), _) | (_, Err(e)) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    if let Err(refusal) = frazaro_core::api::vocab_gate(&text, book, allow_raw) {
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
    if !is_file(book) {
        return refuse_missing("english-vocab-file-not-found", book);
    }
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
    if !is_file(&prelude_path) {
        return refuse_missing("vla-file-not-found", &prelude_path);
    }
    let prelude = match read_text(&prelude_path) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    if let Err(refusal) = frazaro_core::api::vocab_gate(&text, book, allow_raw) {
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

/// `frazaro translate-vla|translate-vba <program.txt> --prelude <prelude.vla>
/// --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]`: the treaty's
/// oracles 1 and 1a. Each phrasebook loads in order, its proofs run as the
/// add-in runs them (a failing proof refuses, as EnglishLoadVocabulary
/// does), then EnglishToVla (or EnglishToVba: its text compiled with the
/// prelude) writes to stdout. A refusal goes to stderr with exit 1.
fn translate(kind: &str, args: &[String]) -> ExitCode {
    let books: Vec<&str> = args
        .windows(2)
        .filter(|w| w[0] == "--phrasebook")
        .map(|w| w[1].as_str())
        .collect();
    let (Some(program), Some(prelude)) = (args.first(), option_after(args, "--prelude")) else {
        eprintln!(
            "usage: frazaro {kind} <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    if books.is_empty() {
        eprintln!("frazaro: {kind} needs at least one --phrasebook <file.vla>");
        return ExitCode::from(2);
    }
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    if !is_file(program) {
        return refuse_missing("english-program-file-not-found", program);
    }
    if !is_file(prelude) {
        return refuse_missing("vla-file-not-found", prelude);
    }
    let (text, prelude) = match (read_text(program), read_text(prelude)) {
        (Ok(t), Ok(p)) => (t, p),
        (Err(e), _) | (_, Err(e)) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    let grammar = match load_books(&prelude, &books, allow_raw) {
        Ok((g, _)) => g,
        Err(code) => return code,
    };
    let out = if kind == "translate-vla" {
        grammar.translate_program(&text).map(|t| t.vla)
    } else {
        grammar.translate_program_to_vba(&text)
    };
    match out {
        Ok(s) => {
            print!("{s}");
            ExitCode::SUCCESS
        }
        Err(refusal) => {
            eprintln!("{refusal}");
            ExitCode::from(1)
        }
    }
}

/// Each phrasebook file read and passed through the door's two gates
/// (`api::vocab_gate`): the texts in order, or the exit code to end the
/// command with.
fn read_books(books: &[&str], allow_raw: bool) -> Result<Vec<String>, ExitCode> {
    let mut texts = Vec::with_capacity(books.len());
    for book in books {
        if !is_file(book) {
            return Err(refuse_missing("english-vocab-file-not-found", book));
        }
        let vocab = read_text(book).map_err(|e| {
            eprintln!("frazaro: {e}");
            ExitCode::from(2)
        })?;
        if let Err(refusal) = frazaro_core::api::vocab_gate(&vocab, book, allow_raw) {
            eprintln!("{refusal}");
            return Err(ExitCode::from(1));
        }
        texts.push(vocab);
    }
    Ok(texts)
}

/// The grammar a command translates with: the prelude, then each phrasebook
/// text loaded in order under its file's name, its proofs run; the first
/// refusal ends the command, and the exit code to end it with comes back as
/// the error. The texts come back too, for the build stamp.
fn load_books(
    prelude: &str,
    books: &[&str],
    allow_raw: bool,
) -> Result<(frazaro_core::english::Grammar, Vec<String>), ExitCode> {
    let texts = read_books(books, allow_raw)?;
    let mut grammar = frazaro_core::english::Grammar::new(prelude);
    for (book, vocab) in books.iter().zip(texts.iter()) {
        if let Err(refusal) = grammar.load_vocabulary_text(vocab, book) {
            eprintln!("{refusal}");
            return Err(ExitCode::from(1));
        }
    }
    Ok((grammar, texts))
}

/// `frazaro rebuild` (PORT.7, slice 7c): a built workbook read back, its
/// stamp checked against its sentences and the files given, the build done
/// again and compared whole; one line, yes or no with why, exit 0 or 1. A
/// file that is not a build this core can verify is refused through the
/// catalogue.
fn rebuild(args: &[String]) -> ExitCode {
    let books: Vec<&str> = args
        .windows(2)
        .filter(|w| w[0] == "--phrasebook")
        .map(|w| w[1].as_str())
        .collect();
    let (Some(file), Some(prelude)) = (args.first(), option_after(args, "--prelude")) else {
        eprintln!(
            "usage: frazaro rebuild <file.xlsx> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    if books.is_empty() {
        eprintln!("frazaro: rebuild needs at least one --phrasebook <file.vla>");
        return ExitCode::from(2);
    }
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    if !is_file(file) {
        return refuse_missing("vla-file-not-found", file);
    }
    if !is_file(prelude) {
        return refuse_missing("vla-file-not-found", prelude);
    }
    let bytes = match std::fs::read(file) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("frazaro: cannot read {file}: {e}");
            return ExitCode::from(2);
        }
    };
    let prelude_text = match read_text(prelude) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    let texts = match read_books(&books, allow_raw) {
        Ok(t) => t,
        Err(code) => return code,
    };
    let texts: Vec<&str> = texts.iter().map(String::as_str).collect();
    match frazaro_core::api::english_rebuild_xlsx(&bytes, &prelude_text, &texts) {
        Ok(r) => {
            println!("{}", r.line());
            if r.matches {
                ExitCode::SUCCESS
            } else {
                ExitCode::from(1)
            }
        }
        Err(r) => {
            eprintln!("{}", r.refusal);
            ExitCode::from(1)
        }
    }
}

/// `frazaro build` (PORT.7): the program translated as translate-vla
/// translates it, then written as a workbook at --out. The door writes the
/// one file and nothing else; a file already there is refused through the
/// catalogue unless --replace says to replace it.
fn build(args: &[String]) -> ExitCode {
    let books: Vec<&str> = args
        .windows(2)
        .filter(|w| w[0] == "--phrasebook")
        .map(|w| w[1].as_str())
        .collect();
    let (Some(program), Some(prelude), Some(out)) = (
        args.first(),
        option_after(args, "--prelude"),
        option_after(args, "--out"),
    ) else {
        eprintln!(
            "usage: frazaro build <program.txt> --prelude <prelude.vla> --phrasebook <file.vla> [--phrasebook ...] --out <file.xlsx> [--replace] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    if books.is_empty() {
        eprintln!("frazaro: build needs at least one --phrasebook <file.vla>");
        return ExitCode::from(2);
    }
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    let replace = args.iter().any(|a| a == "--replace");
    if !is_file(program) {
        return refuse_missing("english-program-file-not-found", program);
    }
    if !is_file(prelude) {
        return refuse_missing("vla-file-not-found", prelude);
    }
    if std::path::Path::new(out).exists() && !replace {
        eprintln!(
            "{}",
            frazaro_core::messages::raise("build-output-exists", &[("path", out)])
        );
        return ExitCode::from(1);
    }
    let (text, prelude_text) = match (read_text(program), read_text(prelude)) {
        (Ok(t), Ok(p)) => (t, p),
        (Err(e), _) | (_, Err(e)) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    let (grammar, texts) = match load_books(&prelude_text, &books, allow_raw) {
        Ok(gt) => gt,
        Err(code) => return code,
    };
    let translation = match grammar.translate_program(&text) {
        Ok(t) => t,
        Err(refusal) => {
            eprintln!("{refusal}");
            return ExitCode::from(1);
        }
    };
    let texts: Vec<&str> = texts.iter().map(String::as_str).collect();
    let bytes =
        match frazaro_core::build::build_xlsx(&text, &translation.vla, &prelude_text, &texts) {
            Ok(b) => b,
            Err(refusal) => {
                eprintln!("{refusal}");
                return ExitCode::from(1);
            }
        };
    if let Err(e) = std::fs::write(out, &bytes) {
        eprintln!("frazaro: cannot write {out}: {e}");
        return ExitCode::from(2);
    }
    // The file's own digest, for a reader to compare with another door's
    // (the page shows the same digest beside its download, slice 7e).
    println!(
        "wrote {out}: {} bytes, sha256 {}",
        bytes.len(),
        frazaro_core::sha256::sha256_hex(&bytes)
    );
    ExitCode::SUCCESS
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("compile") => compile(&args[1..]),
        Some("load") => load(&args[1..]),
        Some("prove") => prove(&args[1..]),
        Some("translate-vla") => translate("translate-vla", &args[1..]),
        Some("translate-vba") => translate("translate-vba", &args[1..]),
        Some("build") => build(&args[1..]),
        Some("rebuild") => rebuild(&args[1..]),
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
