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
  frazaro compile <program.vla> [--prelude <prelude.vla>]
                         the VBA of a VLA program, to stdout (PORT.5)
  frazaro load <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
                         load a phrasebook, its proofs run as the add-in
                         runs them, and report what it holds, the line
                         EnglishVocabStats prints (PORT.6, slice 6c);
                         --allow-raw is the consent a (raw ...) form needs
  frazaro prove <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]
                         the treaty's oracle 3: every proof run, each
                         failure printed, then PASS n/n or FAIL k/n (k of
                         n passed), exit 0 or 1 (PORT.6, slice 6d); the
                         prelude is prelude.vla beside the file or in its
                         parent folder, else the one inside frazaro, unless
                         --prelude names one; an engine proof file exits 3
                         (not attempted, PORT.9)
  frazaro translate-vla <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
                         the treaty's oracle 1: the program's VLA, the text
                         EnglishToVla writes, to stdout (PORT.6, slice 6e);
                         each phrasebook loads in order, its proofs run
  frazaro translate-vba <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
                         oracle 1a: that VLA compiled, the VBA the add-in
                         would write (EnglishToVba)
  frazaro build <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] --out <file.xlsx> [--into <model.xlsx>] [--replace] [--allow-raw]
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
                         stamp, the defined name Frazaro.Build (slice 7c);
                         with --into, the sheets are added to that workbook,
                         whose own parts are copied byte for byte and never
                         written to by a sentence (slice 7d)
  frazaro rebuild <file.xlsx> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]
                         the sentences read back out of a built workbook's
                         Frazaro sheet, its stamp checked against them and
                         the files given, the build done again and compared
                         whole: one line, yes (exit 0) or no with why (exit
                         1); a file that is not a build is refused (slice 7c)
  frazaro reflect <file.xlsx|.ods> [--counts] [--cone <Sheet!A1> ...]
                         the reader (PORT.8, slices 8a and 8b; an
                         OpenDocument .ods file reads the same from slice
                         8e, its formulas spelled as the formula bar shows
                         them): the workbook's relations to stdout in a
                         fixed order,
                         one form a line - every sheet with its state, every
                         name, every Table, then sheet by sheet each cell's
                         value as the file holds it, each formula's text as
                         the formula bar shows it, and what the formula
                         refers to, one row per distinct reference; with
                         --counts, counts and times alone, one line a sheet
                         by position and one for the workbook, the distinct
                         formulas in R1C1 and the INDIRECT and OFFSET calls
                         among them, so that nothing confidential leaves the
                         machine; with --cone, a cell's cone sized through
                         names and Tables and across sheets, counts alone,
                         and where it is blind; a file that is not a
                         workbook, a part with a DOCTYPE, or a shape this
                         version does not read is refused by name
  frazaro diff <old.xlsx|.ods> <new.xlsx|.ods> [--counts]
                         the difference between two workbooks (PORT.8, slice
                         8c), read as reflect reads them: the sheets in one
                         file alone, then the names and Tables that differ,
                         then cell by cell on the matched sheets what
                         changed, old and new side by side - a value, a
                         formula with its cached value, or blank - one form
                         a line in a fixed order, nothing when the two hold
                         the same; with --counts, the counts and times alone
  frazaro audit <file.xlsx|.ods> [--counts]
                         the audit list (PORT.8, slice 8d): where a
                         workbook's risks are, read from the file - a
                         constant typed over a column of formulas, a formula
                         inconsistent with its neighbours, a name nothing
                         refers to, a reference to an empty cell, a hidden
                         sheet, a link to another workbook - one finding a
                         line in a fixed order, nothing when there is
                         nothing to report; with --counts, the six counts
                         and the times alone
  frazaro version        the version of the core this door is built on
  frazaro help           this text

Without --prelude a command uses the prelude inside frazaro, and without
--phrasebook its english.vla: the repository's scripts/prelude.vla and
scripts/polyglotta/english.vla, byte for byte, so that an installed frazaro
needs no file beside it. --phrasebook names the whole list, in order, as
the add-in loads them, english.vla first; `--phrasebook english.vla` with
no such file beside you is the one inside frazaro.

Not in this version (docs/HORIZON.md, section 12, slice by slice):
  check, run, ask
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

/// The prelude and the English phrasebook inside the door, byte for byte the
/// repository's scripts/prelude.vla and scripts/polyglotta/english.vla
/// (tools/check_crate_data.ps1 holds the copies under cli/data/ to them):
/// what a command uses when --prelude or --phrasebook is not given, so that
/// an installed frazaro needs no file beside it.
const BUILT_IN_PRELUDE: &str = include_str!("../data/prelude.vla");
const BUILT_IN_ENGLISH: &str = include_str!("../data/english.vla");
/// The name the built-in phrasebook loads under, which a refusal shows.
const BUILT_IN_ENGLISH_NAME: &str = "english.vla (inside frazaro)";

/// The prelude a command works with: the file --prelude names, read and
/// checked, or the one inside the door; the exit code to end the command
/// with comes back as the error.
fn prelude_of(args: &[String]) -> Result<String, ExitCode> {
    match option_after(args, "--prelude") {
        Some(path) => {
            if !is_file(path) {
                return Err(refuse_missing("vla-file-not-found", path));
            }
            read_text(path).map_err(|e| {
                eprintln!("frazaro: {e}");
                ExitCode::from(2)
            })
        }
        None => Ok(BUILT_IN_PRELUDE.to_string()),
    }
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
    let Some(program) = args.first() else {
        eprintln!("usage: frazaro compile <program.vla> [--prelude <prelude.vla>]");
        return ExitCode::from(2);
    };
    if !is_file(program) {
        return refuse_missing("vla-source-not-found", program);
    }
    let prelude = match prelude_of(args) {
        Ok(p) => p,
        Err(code) => return code,
    };
    let source = match read_text(program) {
        Ok(s) => s,
        Err(e) => {
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
    let Some(book) = args.first() else {
        eprintln!("usage: frazaro load <phrasebook.vla> [--prelude <prelude.vla>] [--allow-raw]");
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    if !is_file(book) {
        return refuse_missing("english-vocab-file-not-found", book);
    }
    let prelude = match prelude_of(args) {
        Ok(p) => p,
        Err(code) => return code,
    };
    let text = match read_text(book) {
        Ok(t) => t,
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
    // reads scripts/prelude.vla), else it is the one inside the door;
    // --prelude names another.
    let prelude_path = option_after(args, "--prelude")
        .map(str::to_string)
        .or_else(|| prelude_beside(book));
    let prelude = match prelude_path {
        Some(path) => {
            if !is_file(&path) {
                return refuse_missing("vla-file-not-found", &path);
            }
            match read_text(&path) {
                Ok(p) => p,
                Err(e) => {
                    eprintln!("frazaro: {e}");
                    return ExitCode::from(2);
                }
            }
        }
        None => BUILT_IN_PRELUDE.to_string(),
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

/// `frazaro translate-vla|translate-vba <program.txt> [--prelude <prelude.vla>]
/// [--phrasebook <file.vla> ...] [--allow-raw]`: the treaty's oracles 1 and
/// 1a. Each phrasebook loads in order, its proofs run as the add-in runs
/// them (a failing proof refuses, as EnglishLoadVocabulary does), then
/// EnglishToVla (or EnglishToVba: its text compiled with the prelude) writes
/// to stdout. A refusal goes to stderr with exit 1.
fn translate(kind: &str, args: &[String]) -> ExitCode {
    let Some(program) = args.first() else {
        eprintln!(
            "usage: frazaro {kind} <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    if !is_file(program) {
        return refuse_missing("english-program-file-not-found", program);
    }
    let prelude = match prelude_of(args) {
        Ok(p) => p,
        Err(code) => return code,
    };
    let text = match read_text(program) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    let grammar = match load_books(&prelude, args, allow_raw) {
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

/// The phrasebooks a command loads, in order, as (name, text): the files
/// --phrasebook names, each read and passed through the door's two gates
/// (`api::vocab_gate`); `english.vla` with no such file beside the caller is
/// the one inside the door, as is the whole list when no --phrasebook is
/// given. The exit code to end the command with comes back as the error.
fn read_books(args: &[String], allow_raw: bool) -> Result<Vec<(String, String)>, ExitCode> {
    let named: Vec<&str> = args
        .windows(2)
        .filter(|w| w[0] == "--phrasebook")
        .map(|w| w[1].as_str())
        .collect();
    let built_in = || {
        (
            BUILT_IN_ENGLISH_NAME.to_string(),
            BUILT_IN_ENGLISH.to_string(),
        )
    };
    if named.is_empty() {
        return Ok(vec![built_in()]);
    }
    let mut books = Vec::with_capacity(named.len());
    for book in named {
        if !is_file(book) {
            if book == "english.vla" {
                books.push(built_in());
                continue;
            }
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
        books.push((book.to_string(), vocab));
    }
    Ok(books)
}

/// The grammar a command translates with: the prelude, then each phrasebook
/// text loaded in order under its name, its proofs run; the first refusal
/// ends the command, and the exit code to end it with comes back as the
/// error. The texts come back too, for the build stamp.
fn load_books(
    prelude: &str,
    args: &[String],
    allow_raw: bool,
) -> Result<(frazaro_core::english::Grammar, Vec<String>), ExitCode> {
    let books = read_books(args, allow_raw)?;
    let mut grammar = frazaro_core::english::Grammar::new(prelude);
    for (name, vocab) in &books {
        if let Err(refusal) = grammar.load_vocabulary_text(vocab, name) {
            eprintln!("{refusal}");
            return Err(ExitCode::from(1));
        }
    }
    Ok((grammar, books.into_iter().map(|(_, text)| text).collect()))
}

/// `frazaro rebuild` (PORT.7, slice 7c): a built workbook read back, its
/// stamp checked against its sentences and the files given, the build done
/// again and compared whole; one line, yes or no with why, exit 0 or 1. A
/// file that is not a build this core can verify is refused through the
/// catalogue.
fn rebuild(args: &[String]) -> ExitCode {
    let Some(file) = args.first() else {
        eprintln!(
            "usage: frazaro rebuild <file.xlsx> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    if !is_file(file) {
        return refuse_missing("vla-file-not-found", file);
    }
    let prelude_text = match prelude_of(args) {
        Ok(p) => p,
        Err(code) => return code,
    };
    let bytes = match std::fs::read(file) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("frazaro: cannot read {file}: {e}");
            return ExitCode::from(2);
        }
    };
    let books = match read_books(args, allow_raw) {
        Ok(b) => b,
        Err(code) => return code,
    };
    let texts: Vec<&str> = books.iter().map(|(_, text)| text.as_str()).collect();
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
    let (Some(program), Some(out)) = (args.first(), option_after(args, "--out")) else {
        eprintln!(
            "usage: frazaro build <program.txt> [--prelude <prelude.vla>] [--phrasebook <file.vla> ...] --out <file.xlsx> [--into <model.xlsx>] [--replace] [--allow-raw]"
        );
        return ExitCode::from(2);
    };
    let allow_raw = args.iter().any(|a| a == "--allow-raw");
    let replace = args.iter().any(|a| a == "--replace");
    let into = option_after(args, "--into");
    if !is_file(program) {
        return refuse_missing("english-program-file-not-found", program);
    }
    let prelude_text = match prelude_of(args) {
        Ok(p) => p,
        Err(code) => return code,
    };
    if let Some(model) = into {
        if !is_file(model) {
            return refuse_missing("vla-file-not-found", model);
        }
    }
    if std::path::Path::new(out).exists() && !replace {
        eprintln!(
            "{}",
            frazaro_core::messages::raise("build-output-exists", &[("path", out)])
        );
        return ExitCode::from(1);
    }
    let text = match read_text(program) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("frazaro: {e}");
            return ExitCode::from(2);
        }
    };
    let (grammar, texts) = match load_books(&prelude_text, args, allow_raw) {
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
    let built = match into {
        None => frazaro_core::build::build_xlsx(&text, &translation.vla, &prelude_text, &texts),
        Some(model) => match std::fs::read(model) {
            Ok(bytes) => frazaro_core::build::build_xlsx_into(
                &text,
                &translation.vla,
                &prelude_text,
                &texts,
                bytes,
                model,
            ),
            Err(e) => {
                eprintln!("frazaro: cannot read {model}: {e}");
                return ExitCode::from(2);
            }
        },
    };
    let bytes = match built {
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

/// Milliseconds since `since`, for the counts' lines.
fn ms_since(since: std::time::Instant) -> f64 {
    since.elapsed().as_secs_f64() * 1000.0
}

/// `frazaro reflect <file.xlsx> [--counts] [--cone <Sheet!A1> ...]`
/// (PORT.8, slices 8a and 8b): the workbook's relations to stdout in the
/// fixed order, streamed sheet by sheet through the core's reader, which
/// holds no model of the workbook. With --counts, counts and times alone:
/// one line a sheet by position and one for the workbook, AXM.1's
/// discipline (no file name, sheet name, address, formula text or value
/// leaves the machine), which is how the reader gives that measurement its
/// second number; from 8b the distinct formulas in R1C1 and the unreadable
/// calls are among them. With --cone, the one exception to streaming: an
/// index of every formula is built in one walk and each root's cone is
/// sized through it, counts alone, the root being the one address printed
/// because the caller typed it. A refusal goes to stderr with exit 1; the
/// rows printed before it stand, since the read streams.
fn reflect(args: &[String]) -> ExitCode {
    const USAGE_LINE: &str =
        "usage: frazaro reflect <file.xlsx|.ods> [--counts] [--cone <Sheet!A1> ...]";
    let Some(file) = args.first() else {
        eprintln!("{USAGE_LINE}");
        return ExitCode::from(2);
    };
    let counts = args.iter().any(|a| a == "--counts");
    let mut cones: Vec<String> = Vec::new();
    let mut at = 1;
    while at < args.len() {
        if args[at] == "--cone" {
            match args.get(at + 1) {
                Some(root) => cones.push(root.clone()),
                None => {
                    eprintln!("{USAGE_LINE}");
                    return ExitCode::from(2);
                }
            }
            at += 2;
        } else {
            at += 1;
        }
    }
    if !is_file(file) {
        return refuse_missing("vla-file-not-found", file);
    }
    let bytes = match std::fs::read(file) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("frazaro: cannot read {file}: {e}");
            return ExitCode::from(2);
        }
    };
    let opened = std::time::Instant::now();
    let package = match frazaro_core::reflect::open(&bytes, file) {
        Ok(p) => p,
        Err(refusal) => {
            eprintln!("{refusal}");
            return ExitCode::from(1);
        }
    };
    let open_ms = ms_since(opened);
    if counts {
        let (mut cells, mut formulas, mut arrays, mut unreadable, mut part_bytes) =
            (0u64, 0u64, 0u64, 0u64, 0usize);
        let mut book_distinct: std::collections::HashSet<String> = std::collections::HashSet::new();
        let mut read_ms = 0.0f64;
        for i in 0..package.sheets().len() {
            let started = std::time::Instant::now();
            let s = match package.walk_sheet_with(
                i,
                &mut frazaro_core::reflect::Discard,
                &mut book_distinct,
            ) {
                Ok(s) => s,
                Err(refusal) => {
                    eprintln!("{refusal}");
                    return ExitCode::from(1);
                }
            };
            let took = ms_since(started);
            println!(
                "sheet {}: cells {} formulas {} distinct-r1c1 {} unreadable {} arrays {} rows {} columns {} bytes {} read {took:.1} ms",
                i + 1,
                s.cells,
                s.formulas,
                s.distinct_r1c1,
                s.unreadable,
                s.array_anchors,
                s.rows,
                s.columns,
                s.part_bytes
            );
            cells += s.cells;
            formulas += s.formulas;
            arrays += s.array_anchors;
            unreadable += s.unreadable;
            part_bytes += s.part_bytes;
            read_ms += took;
        }
        let summary = package.summary();
        println!(
            "workbook: sheets {} cells {cells} formulas {formulas} distinct-r1c1 {} unreadable {unreadable} arrays {arrays} names {} placeholders {} tables {} books {} strings {} bytes {part_bytes} string-bytes {} open {open_ms:.1} ms read {read_ms:.1} ms",
            package.sheets().len(),
            book_distinct.len(),
            summary.names,
            summary.placeholders,
            summary.tables,
            summary.external_books,
            summary.strings,
            summary.string_bytes
        );
        return ExitCode::SUCCESS;
    }
    if !cones.is_empty() {
        let mut index = frazaro_core::reflect::cone::Index::new(package.as_ref());
        let started = std::time::Instant::now();
        for i in 0..package.sheets().len() {
            if let Err(refusal) = package.walk_sheet(i, &mut index) {
                eprintln!("{refusal}");
                return ExitCode::from(1);
            }
        }
        let indexed_ms = ms_since(started);
        for root in &cones {
            let Some((sheet, row, col)) = frazaro_core::refers::parse_home(root) else {
                eprintln!("frazaro: --cone wants a cell as Sheet!A1 or 'Q1 Data'!A1, not {root}");
                return ExitCode::from(2);
            };
            let wanted = frazaro_core::intrinsics::fold(&sheet);
            let Some(s) = package
                .sheets()
                .iter()
                .position(|info| frazaro_core::intrinsics::fold(&info.name) == wanted)
            else {
                eprintln!("frazaro: --cone {root}: the workbook has no sheet so named");
                return ExitCode::from(2);
            };
            let started = std::time::Instant::now();
            let stats = frazaro_core::reflect::cone::cone(&index, package.as_ref(), s, row, col);
            println!(
                "cone {root}: {} indexed {indexed_ms:.1} ms walked {:.1} ms",
                stats.line(),
                ms_since(started)
            );
        }
        return ExitCode::SUCCESS;
    }
    use std::io::Write;
    struct Lines<W: Write>(W);
    impl<W: Write> frazaro_core::reflect::Sink for Lines<W> {
        fn row(&mut self, row: &frazaro_core::reflect::Row<'_>) {
            let _ = writeln!(self.0, "{}", frazaro_core::reflect::print::line(row));
        }
    }
    let mut out = Lines(std::io::BufWriter::new(std::io::stdout().lock()));
    package.header_rows(&mut out);
    for i in 0..package.sheets().len() {
        if let Err(refusal) = package.walk_sheet(i, &mut out) {
            let _ = out.0.flush();
            eprintln!("{refusal}");
            return ExitCode::from(1);
        }
    }
    if out.0.flush().is_err() {
        return ExitCode::from(2);
    }
    ExitCode::SUCCESS
}

/// `frazaro diff <old.xlsx> <new.xlsx> [--counts]` (PORT.8, slice 8c): both
/// files opened as `reflect` opens them, the difference streamed one row a
/// line in the fixed order, or with `--counts` one line of counts and times.
/// Exit 0 whenever the comparison ran, rows or none; a refusal exits 1 and
/// the rows before it stand; a usage error or an unreadable file exits 2.
fn diff(args: &[String]) -> ExitCode {
    const USAGE_LINE: &str = "usage: frazaro diff <old.xlsx|.ods> <new.xlsx|.ods> [--counts]";
    let mut files: Vec<&str> = Vec::new();
    let mut counts = false;
    for a in args {
        if a == "--counts" {
            counts = true;
        } else if a.starts_with("--") || files.len() == 2 {
            eprintln!("{USAGE_LINE}");
            return ExitCode::from(2);
        } else {
            files.push(a);
        }
    }
    if files.len() != 2 {
        eprintln!("{USAGE_LINE}");
        return ExitCode::from(2);
    }
    let mut bytes: Vec<Vec<u8>> = Vec::new();
    for file in &files {
        if !is_file(file) {
            return refuse_missing("vla-file-not-found", file);
        }
        match std::fs::read(file) {
            Ok(b) => bytes.push(b),
            Err(e) => {
                eprintln!("frazaro: cannot read {file}: {e}");
                return ExitCode::from(2);
            }
        }
    }
    let opened = std::time::Instant::now();
    let mut packages = Vec::new();
    for (i, file) in files.iter().enumerate() {
        match frazaro_core::reflect::open(&bytes[i], file) {
            Ok(p) => packages.push(p),
            Err(refusal) => {
                eprintln!("{refusal}");
                return ExitCode::from(1);
            }
        }
    }
    let open_ms = ms_since(opened);
    let (old, new) = (packages[0].as_ref(), packages[1].as_ref());
    use frazaro_core::reflect::diff::{self as d, ChangeSink};
    if counts {
        let started = std::time::Instant::now();
        return match d::diff(old, new, &mut d::Discard) {
            Ok(stats) => {
                println!(
                    "diff: {} open {open_ms:.1} ms compared {:.1} ms",
                    stats.line(),
                    ms_since(started)
                );
                ExitCode::SUCCESS
            }
            Err(refusal) => {
                eprintln!("{refusal}");
                ExitCode::from(1)
            }
        };
    }
    use std::io::Write;
    struct Changes<W: Write>(W);
    impl<W: Write> ChangeSink for Changes<W> {
        fn change(&mut self, change: &d::Change<'_>) {
            let _ = writeln!(self.0, "{}", d::line(change));
        }
    }
    let mut out = Changes(std::io::BufWriter::new(std::io::stdout().lock()));
    if let Err(refusal) = d::diff(old, new, &mut out) {
        let _ = out.0.flush();
        eprintln!("{refusal}");
        return ExitCode::from(1);
    }
    if out.0.flush().is_err() {
        return ExitCode::from(2);
    }
    ExitCode::SUCCESS
}

/// `frazaro audit <file.xlsx> [--counts]` (PORT.8, slice 8d): the file read
/// as `reflect` reads it into the audit's index, then the six walks, one
/// finding a line in the fixed order, or with `--counts` one line of counts
/// and times. Exit 0 whenever the walks ran, findings or none; a refusal
/// exits 1; a usage error or an unreadable file exits 2.
fn audit(args: &[String]) -> ExitCode {
    const USAGE_LINE: &str = "usage: frazaro audit <file.xlsx|.ods> [--counts]";
    let mut file: Option<&str> = None;
    let mut counts = false;
    for a in args {
        if a == "--counts" {
            counts = true;
        } else if a.starts_with("--") || file.is_some() {
            eprintln!("{USAGE_LINE}");
            return ExitCode::from(2);
        } else {
            file = Some(a);
        }
    }
    let Some(file) = file else {
        eprintln!("{USAGE_LINE}");
        return ExitCode::from(2);
    };
    if !is_file(file) {
        return refuse_missing("vla-file-not-found", file);
    }
    let bytes = match std::fs::read(file) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("frazaro: cannot read {file}: {e}");
            return ExitCode::from(2);
        }
    };
    let opened = std::time::Instant::now();
    let package = match frazaro_core::reflect::open(&bytes, file) {
        Ok(p) => p,
        Err(refusal) => {
            eprintln!("{refusal}");
            return ExitCode::from(1);
        }
    };
    let open_ms = ms_since(opened);
    use frazaro_core::reflect::audit::{self as au, FindingSink};
    let started = std::time::Instant::now();
    let mut index = au::AuditIndex::new(package.as_ref());
    for i in 0..package.sheets().len() {
        if let Err(refusal) = package.walk_sheet(i, &mut index) {
            eprintln!("{refusal}");
            return ExitCode::from(1);
        }
    }
    let indexed_ms = ms_since(started);
    if counts {
        let started = std::time::Instant::now();
        let stats = au::audit(&index, package.as_ref(), &mut au::Discard);
        println!(
            "audit: {} open {open_ms:.1} ms indexed {indexed_ms:.1} ms walked {:.1} ms",
            stats.line(),
            ms_since(started)
        );
        return ExitCode::SUCCESS;
    }
    use std::io::Write;
    struct Findings<W: Write>(W);
    impl<W: Write> FindingSink for Findings<W> {
        fn finding(&mut self, finding: &au::Finding<'_>) {
            let _ = writeln!(self.0, "{}", au::line(finding));
        }
    }
    let mut out = Findings(std::io::BufWriter::new(std::io::stdout().lock()));
    au::audit(&index, package.as_ref(), &mut out);
    if out.0.flush().is_err() {
        return ExitCode::from(2);
    }
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
        Some("reflect") => reflect(&args[1..]),
        Some("diff") => diff(&args[1..]),
        Some("audit") => audit(&args[1..]),
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
