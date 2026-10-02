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

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("compile") => compile(&args[1..]),
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
