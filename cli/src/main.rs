//! frazaro: the command-line door to frazaro-core.
//!
//! The commands this door will carry, in the order the slices of
//! docs/HORIZON.md section 12 deliver them: check, build, run, ask, diff,
//! rebuild, prove. Until a slice lands, the door reports its version and the
//! core's ABI number, and refuses every other command by name, in words.

use std::process::ExitCode;

const USAGE: &str = "\
frazaro - a compiler from sentences to spreadsheets

usage:
  frazaro version        the version of the core this door is built on
  frazaro help           this text

Not in this version (docs/HORIZON.md, section 12, slice by slice):
  check, build, run, ask, diff, rebuild, prove
";

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
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
