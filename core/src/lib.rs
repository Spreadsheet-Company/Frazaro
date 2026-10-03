//! frazaro-core: the host-free home of the Frazaro language.
//!
//! The contract (docs/HORIZON.md, section 12): bytes in, bytes out, nothing
//! else. No clock, no randomness, no threads, no file system, no socket. On
//! the wasm32 target the built module's import section is empty, and
//! tools/check_core_imports.ps1 pins that count at zero.
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
//! workbook's bytes, deterministic, with no clock and no dependency;
//! `api::english_build_xlsx` is its surface.

pub mod abi;
pub mod api;
pub mod build;
pub mod emit;
pub mod english;
pub mod expand;
pub mod form;
pub mod headtable;
pub mod intrinsics;
pub mod messages;
pub mod printer;
pub mod reader;
pub mod sheet;
pub mod version;

/// The crate's version, which tools/check_version_twin.ps1 holds equal to
/// VLA_RELEASE_VERSION in src/VLA.bas: one corpus, one version.
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
}
