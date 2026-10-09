//! vla-lang: the VLA language with nothing around it (PORT.12, 2026-10-08).
//!
//! The language cut out of `frazaro-core` so that an engine can stand on it
//! without the bridges: the reader (`reader`), the forms (`form`) and their
//! printing (`printer`), macro expansion (`expand`), the head table
//! (`headtable`), the six intrinsics (`intrinsics`), the message catalogue's
//! mechanism with the language's own refusals (`messages`), the version
//! predicates (`version`), the grid (`sheet`: sheets of cells with A1
//! coordinates and a style table), the references read out of a formula's
//! text (`refers`), the relation rows and their spelling (`rows`), the
//! projections seam (`projection`) and its first implementation, the view
//! record of one window (`view`); and, since KERNEL.7 (2026-10-08),
//! recalculation's mechanism (`calc`): the formula parser over the
//! reference scanner, the dependency graph, evaluation in order with a
//! budget, a cycle refused by name, the error values, and the registry seam
//! for functions with the language's own day-one functions registered.
//!
//! The contract (docs/HORIZON.md section 12; Alonzo/CHARTER.md section 4,
//! rule 1): bytes in, bytes out, nothing else. No English, no file format,
//! no clock, no randomness, no threads, no file system, no socket, and no
//! dependency. Built for wasm32 the module's import section is empty, which
//! tools/check_core_imports.ps1 reads off the artifact. Everything that
//! consumes this crate is a bridge: `frazaro-core` adds the English engine,
//! the emitters and the workbook formats; an engine adds its devices and
//! its clock, and never `frazaro-core`.
//!
//! SD-18 governs what this crate may claim: the VBA in src/ is the reference
//! implementation, and this crate follows the goldens in scripts/ and never
//! leads them; its tests read them from their test modules.

pub mod calc;
pub mod expand;
pub mod form;
pub mod headtable;
pub mod intrinsics;
pub mod messages;
pub mod printer;
pub mod projection;
pub mod reader;
pub mod refers;
pub mod rows;
pub mod sheet;
pub mod version;
pub mod view;

/// The crate's version: the workspace's, which tools/check_version_twin.ps1
/// holds equal to VLA_RELEASE_VERSION in src/VLA.bas and to the exact pins
/// the core and the door hold. One corpus, one version.
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

/// The version as text.
pub fn version() -> &'static str {
    VERSION
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
}
