//! The core's half of the message catalogue, over the language's mechanism
//! (PORT.12): every refusal is `RaiseMsg "<id>", slot, value, ...` in the
//! reference (SD-2), and since the cut the exported catalogue is two files
//! from the one VBA source, `src/VLA_Messages.bas`, written by
//! `tools/export_messages.ps1` and held to it by
//! `tools/check_data_exports.ps1`: `vla-lang/data/messages.vla` holds the
//! language's families (`vla`, `interp`, `lint`, `view`), and
//! `core/data/messages.vla`, embedded here, holds the rest: the English
//! engine's, the build's, the reader's, the distro's, the door's, the
//! engines' and the add-in's. [`raise`] asks this half first and the
//! language's after it, so an id raised from either crate is found, and the
//! two internal errors are the reference's own words as before.
//!
//! [`Refusal`] is one type, the language's, so a refusal flows from the
//! reader through the emitters to a door unchanged.

use std::sync::OnceLock;

pub use vla_lang::messages::{Catalogue, Refusal};

/// The core's half of the exported catalogue, embedded at build time: no
/// file is read when the core runs (the wasm import section stays empty).
const CATALOGUE_TEXT: &str = include_str!("../data/messages.vla");

fn own() -> &'static Catalogue {
    static CATALOGUE: OnceLock<Catalogue> = OnceLock::new();
    CATALOGUE.get_or_init(|| Catalogue::parse(CATALOGUE_TEXT))
}

/// How many entries the two halves hold together.
pub fn count() -> usize {
    vla_lang::messages::count() + own().count()
}

/// The template of an id, unfilled, from either half, or `None` for an id
/// neither knows.
pub fn template(id: &str) -> Option<&'static str> {
    own()
        .template(id)
        .or_else(|| vla_lang::messages::template(id))
}

/// Every id, the language's half first and this half after it, each in
/// the order the VBA registers them.
pub fn ids() -> Vec<&'static str> {
    let mut all = vla_lang::messages::ids();
    all.extend(own().ids());
    all
}

/// VBA `RaiseMsg`: the entry's number and source, and its template with
/// each `{slot}` replaced by the value given for that name, from this half
/// or, failing that, the language's; an id neither knows is the reference's
/// internal error.
pub fn raise(id: &str, slots: &[(&str, &str)]) -> Refusal {
    own()
        .raise(id, slots)
        .or_else(|| vla_lang::messages::catalogue().raise(id, slots))
        .unwrap_or_else(|| vla_lang::messages::unknown(id))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_catalogue_is_read_whole() {
        // 574 entries at the export of 2026-10-01 (VLA_Messages.bas, LX.14);
        // tools/check_data_exports.ps1 holds the floors on the two files.
        assert!(count() >= 574, "{} entries read", count());
        assert_eq!(ids()[0], "vla-source-not-found");
        assert_eq!(
            template("vla-unbalanced-parens").unwrap(),
            "unbalanced parentheses: missing ')'{loc}"
        );
        // A template exported with a line break inside it reads as one
        // string, with LF on every platform.
        assert!(template("english-test-failed")
            .unwrap()
            .contains("test FAILED\n  sentence: {sentence}"));
        // A named error number, resolved by the exporter from VLA.bas,
        // found through the language's half.
        let r = raise(
            "vla-interpreter-only-handler",
            &[("name", "on:sheet-change")],
        );
        assert_eq!(r.number, -2147221504 + 7001);
        assert_eq!(r.source, "VLA");
        // The two halves share no id, and this half holds none of the
        // language's families.
        let language = vla_lang::messages::ids();
        for id in own().ids() {
            assert!(!language.contains(&id), "{id} is in both halves");
            let family = id.split('-').next().unwrap();
            assert!(
                !["vla", "interp", "lint", "view", "calc", "grid"].contains(&family),
                "{id} belongs to the language's half"
            );
        }
    }

    #[test]
    fn raise_fills_slots_as_raisemsg_does() {
        let r = raise("vla-form-missing-element", &[("n", "3")]);
        assert_eq!(r.id, "vla-form-missing-element");
        assert_eq!(r.number, 5);
        assert_eq!(r.source, "VLA");
        assert_eq!(r.text, "form is missing required element 3");
        let r = raise("vla-unbalanced-parens", &[("loc", "")]);
        assert_eq!(r.text, "unbalanced parentheses: missing ')'");
        let r = raise("build-output-exists", &[("path", "x.xlsx")]);
        assert_eq!(r.source, "VLA-Build");
        assert!(r.text.contains("x.xlsx"), "{}", r.text);
    }

    #[test]
    fn the_two_internal_errors_are_refusals_too() {
        let r = raise("no-such-id", &[]);
        assert_eq!(r.number, 5);
        assert_eq!(r.source, "VLA-Messages");
        assert_eq!(r.text, "RaiseMsg: unknown message id 'no-such-id'");
        let r = raise("vla-form-missing-element", &[]);
        assert_eq!(r.text, "RaiseMsg: no value supplied for slot '{n}'");
        let r = raise("english-test-failed", &[]);
        assert!(r.text.starts_with("RaiseMsg: no value supplied for slot"));
    }
}
