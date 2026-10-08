//! frazaro-core: the host-free home of the Frazaro language's bridges.
//!
//! The contract (docs/HORIZON.md, section 12): bytes in, bytes out, nothing
//! else. No clock, no randomness, no threads, no file system, no socket. On
//! the wasm32 target the built module's import section is empty, and
//! tools/check_core_imports.ps1 pins that count at zero.
//!
//! Since PORT.12 (2026-10-08) the language itself is a crate of its own,
//! `vla-lang`, this crate's one dependency: the reader, the forms and their
//! printing, macro expansion, the head table, the intrinsics, the message
//! catalogue's mechanism with the language's own refusals, the version
//! predicates, the grid, the references read out of a formula's text, the
//! relation rows and the view record. Those modules are re-exported here
//! whole, below, so every path a door used before the cut resolves as it
//! did. What this crate holds is the bridges: the English engine
//! (`english`), the VBA and formula emitters (`emit`), the workbook writer
//! (`build` over `sheet`'s formats) and the readers (`reflect`), the formula
//! sink's scan (`egress`), the distro manifest (`distro`), the kernel's
//! registries (`kernel`), the core's half of the catalogue (`messages`), the
//! API (`api`) and the C surface (`abi`). The three-crate order is
//! Alonzo/CHARTER.md section 4, rule 1: an engine consumes `vla-lang` and
//! never this crate.
//!
//! SD-18 governs what this crate may claim: the VBA in src/ is the reference
//! implementation, and this crate follows the goldens in scripts/ and never
//! leads them. conformance/README.md is the treaty; tools/prove.ps1 scores an
//! implementation against it.
//!
//! The API a door calls is VLA_Browser.bas's pair, text in and text out:
//! `api::english_translate_text_to_vla` and `api::english_translate_text_to_vba`
//! (PORT.6, slice 6g), and `abi` exports the pair with C linkage for a wasm
//! host or a native embedding (6h): an allocator pair, the two functions, the
//! door's gate and the version, every answer one record in the module's
//! memory. The version and the ABI number were exported from the first
//! commit, so that CI held the zero-import property before anything else
//! existed.
//!
//! The writer (PORT.7) is `build` over `sheet`: a translated program to a
//! workbook's bytes, deterministic, with no clock and no dependency outside
//! the repository; `api::english_build_xlsx` is its surface.
//!
//! The reader (PORT.8) is `reflect`: a workbook's bytes, OOXML first, read
//! as the relations of `REFLECT` with no host, streamed sheet by sheet;
//! `api::reflect_relations` is its surface. `refers` (slice 8b, the
//! language's since the cut) is the formula-reference reader,
//! `VLA_Refers.bas` ported and held to its golden: the `refers` rows, the
//! writer's mover and the R1C1 rendering all come off its one scan.
//!
//! `egress` (SEC.15) is the formula sink's scan, `VlaFormulaEgress` of
//! `VLA_Runtime.bas` held to the egress golden: the writer asks it before
//! it writes a formula, as the add-in's two backends do, and refuses one
//! that reaches outside the workbook on its own.
//!
//! `kernel` (KERNEL.1) writes the kernel's boundary and its five seams down
//! as data and names the engines trait; the projections seam's trait is the
//! language's (`vla_lang::projection`) and is re-exported there. `view`
//! (KERNEL.4, the language's) is the first implementation of that seam: one
//! window of the sheet model as the view record, a pure function of the
//! program and the books, the lines a viewport draws from;
//! `api::english_view` is its surface. `distro` (KERNEL.2) reads a distro's
//! manifest, the folder the build tools take: the prelude, the phrasebooks
//! in their load order, the dialects, the libraries and an edition's chrome,
//! by reference; the command-line door proves a distro whole with it.

pub mod abi;
pub mod api;
pub mod build;
pub mod distro;
pub mod egress;
pub mod emit;
pub mod english;
pub mod kernel;
pub mod messages;
pub mod reflect;
pub mod sha256;
pub mod sheet;

// The language, re-exported whole (PORT.12): `frazaro_core::form`,
// `crate::view::Grid` and every other path a door or a test used before
// the cut resolve as they did. The modules are vla-lang's, documented there.
pub use vla_lang::{expand, form, headtable, intrinsics, printer, reader, refers, view};

/// The version predicates (`parse`, `compare`, `at_least`), the language's,
/// re-exported as the module they were; the function `version` below is
/// the core's own.
pub mod version {
    pub use vla_lang::version::*;
}

/// The crate's version, which tools/check_version_twin.ps1 holds equal to
/// VLA_RELEASE_VERSION in src/VLA.bas and to the language crate's: one
/// corpus, one version.
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

/// The number a door checks before calling anything else. It changes only
/// when an exported signature changes its meaning.
pub const ABI_VERSION: u32 = 1;

/// The version as text.
pub fn version() -> &'static str {
    VERSION
}

/// The ABI number, exported with C linkage so that a wasm host or a native
/// embedding can read it without any binding layer.
#[no_mangle]
pub extern "C" fn frazaro_abi_version() -> u32 {
    ABI_VERSION
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn version_is_major_minor_patch() {
        let parts: Vec<&str> = version().split('.').collect();
        assert_eq!(parts.len(), 3, "SD-14 numbers releases MAJOR.MINOR.PATCH");
        for part in parts {
            assert!(
                part.parse::<u32>().is_ok(),
                "non-numeric version part: {part}"
            );
        }
    }

    #[test]
    fn abi_version_is_one() {
        assert_eq!(frazaro_abi_version(), 1);
        assert_eq!(ABI_VERSION, 1);
    }

    #[test]
    fn the_language_and_the_core_are_one_version() {
        assert_eq!(vla_lang::VERSION, VERSION, "one corpus, one version");
    }
}
