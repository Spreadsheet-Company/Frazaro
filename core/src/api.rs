//! The translate API (PORT.6, slice 6g): VLA_Browser.bas's two functions,
//! the surface a door calls - the C-ABI export (6h), the web page, any
//! embedding. Text in, text out, and a refusal back as the teaching text the
//! add-in shows, as a value instead of a dialog.
//!
//! `EnglishTranslateTextToVla` is `EnglishResetGrammar`, then
//! `EnglishLoadVocabularyText` for each phrasebook text under the name
//! `vocab-N` (N from 1, in the order given), then `EnglishToVla`;
//! `EnglishTranslateTextToVba` is the same and then `VlaTranspile` with the
//! prelude set by `VlaSetPreludeOverride`. One difference the port cannot
//! avoid: the reference's VLA stage reads its prelude from the workbook (for
//! `ValidateRawVla`'s probe and a generator's expansion), and a host-free core
//! has no workbook, so both functions take the prelude's text.
//!
//! The text path is ungated, as the reference's is (`TestRawConsentTextPathUngated`):
//! SEC.2's consent for a `(raw ...)` form and F.10's capability check are a
//! door's questions, asked before this is called, with
//! `english::vocab::vocab_text_has_raw_form` and
//! `english::vocab::vocab_requires_check_capability`, as the CLI asks them.

use crate::english::program::RefusalAtLine;
use crate::english::vocab::{vocab_requires_check_capability, vocab_text_has_raw_form};
use crate::english::Grammar;
use crate::messages::{raise, Refusal};

/// VBA `EnglishTranslateTextToVla`: the program's VLA text, or the refusal
/// with its id and the program line it stands on (0 when a phrasebook was
/// refused at load, whose text carries its own `vocab-N line L` location).
pub fn english_translate_text_to_vla(
    program_text: &str,
    prelude_text: &str,
    vocab_texts: &[&str],
) -> Result<String, RefusalAtLine> {
    let g = load_grammar(prelude_text, vocab_texts)?;
    g.translate_program_at(program_text).map(|t| t.vla)
}

/// VBA `EnglishTranslateTextToVba`: the program's VBA text, or the English
/// stage's refusal as above; a refusal of the compile stage carries line 0
/// and names its VLA line in its own text.
pub fn english_translate_text_to_vba(
    program_text: &str,
    prelude_text: &str,
    vocab_texts: &[&str],
) -> Result<String, RefusalAtLine> {
    let g = load_grammar(prelude_text, vocab_texts)?;
    let t = g.translate_program_at(program_text)?;
    crate::emit::compile(&t.vla, prelude_text).map_err(|refusal| RefusalAtLine { refusal, line: 0 })
}

/// The writer's surface (PORT.7): the program translated as
/// [`english_translate_text_to_vla`] translates it, then built into a
/// workbook's bytes (`build::build_xlsx`); the English stage's refusal with
/// its line, or the build's with line 0.
pub fn english_build_xlsx(
    program_text: &str,
    prelude_text: &str,
    vocab_texts: &[&str],
) -> Result<Vec<u8>, RefusalAtLine> {
    let g = load_grammar(prelude_text, vocab_texts)?;
    let t = g.translate_program_at(program_text)?;
    crate::build::build_xlsx(program_text, &t.vla)
        .map_err(|refusal| RefusalAtLine { refusal, line: 0 })
}

/// `EnglishResetGrammar` and the loads: a fresh grammar over the prelude, each
/// phrasebook text loaded in order as `vocab-N`, its proofs run as the add-in
/// runs them, the first refusal ending the call.
fn load_grammar(prelude_text: &str, vocab_texts: &[&str]) -> Result<Grammar, RefusalAtLine> {
    let mut g = Grammar::new(prelude_text);
    for (i, text) in vocab_texts.iter().enumerate() {
        g.load_vocabulary_text(text, &format!("vocab-{}", i + 1))
            .map_err(|refusal| RefusalAtLine { refusal, line: 0 })?;
    }
    Ok(g)
}

/// The door's two gates before a phrasebook text it did not ship loads, as
/// the reference's file loader asks them and its text loader does not
/// (`EnglishLoadVocabulary`): F.10's capability check, then SEC.2's consent
/// for a `(raw ...)` form, which a person gives (`--allow-raw`, a checkbox).
/// `source_name` is the name the refusal shows.
pub fn vocab_gate(text: &str, source_name: &str, allow_raw: bool) -> Result<(), Refusal> {
    vocab_requires_check_capability(text, source_name)?;
    if !allow_raw && vocab_text_has_raw_form(text) {
        return Err(raise(
            "english-vocab-raw-consent-declined",
            &[("source", source_name)],
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../scripts/polyglotta/english.vla");
    const ESPANOL: &str = include_str!("../../scripts/polyglotta/espanol.vla");
    const DANSK: &str = include_str!("../../scripts/polyglotta/dansk.vla");
    const DEUTSCHE: &str = include_str!("../../scripts/polyglotta/deutsche.vla");
    const ESPERANTO: &str = include_str!("../../scripts/polyglotta/esperanto.vla");
    const FRANCAIS: &str = include_str!("../../scripts/polyglotta/francais.vla");
    const LATIN: &str = include_str!("../../scripts/polyglotta/latin.vla");
    const PIRATE: &str = include_str!("../../scripts/polyglotta/pirate.vla");
    const PROGRAM: &str = include_str!("../../scripts/instructions.txt");
    const GOLDEN_VLA: &str = include_str!("../../scripts/instructions_golden.vla");
    const GOLDEN_VBA: &str = include_str!("../../scripts/instructions_golden.vba");

    fn normalized(s: &str) -> String {
        s.replace("\r\n", "\n").trim_end_matches('\n').to_string()
    }

    /// Oracle 1 through the API: `instructions.txt` to the `.vla` golden less
    /// its stamp line.
    #[test]
    fn to_vla_reproduces_the_translate_golden() {
        let got = english_translate_text_to_vla(PROGRAM, PRELUDE, &[ENGLISH])
            .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        let want = GOLDEN_VLA
            .split_once('\n')
            .map(|(_, rest)| rest)
            .unwrap_or("");
        assert_eq!(normalized(&got), normalized(want));
    }

    /// Oracle 1a through the API: `instructions.txt` to the `.vba` golden.
    #[test]
    fn to_vba_reproduces_the_compile_golden() {
        let got = english_translate_text_to_vba(PROGRAM, PRELUDE, &[ENGLISH])
            .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert_eq!(normalized(&got), normalized(GOLDEN_VBA));
    }

    #[test]
    fn a_phrasebook_refusal_names_the_browser_s_source() {
        let e = english_translate_text_to_vla("Log 1.", PRELUDE, &[ENGLISH, "hello"]).unwrap_err();
        assert_eq!(e.refusal.id, "english-vocab-expected-directive");
        // The reader gives a bare atom at the top level no line, so the
        // location says line 0: the reference's own reading, which the
        // refusal golden pins (scripts/refusals_golden.txt).
        assert!(
            e.refusal.text.starts_with("vocab-2 line 0: "),
            "{}",
            e.refusal.text
        );
        assert_eq!(e.line, 0);
    }

    #[test]
    fn a_program_refusal_carries_its_line_and_to_vba_passes_it_through() {
        let program = "Log 1.\nSet total to $5.";
        let e = english_translate_text_to_vla(program, PRELUDE, &[ENGLISH]).unwrap_err();
        assert_eq!(
            (e.refusal.id.as_str(), e.line),
            ("english-unknown-character", 2)
        );
        let e2 = english_translate_text_to_vba(program, PRELUDE, &[ENGLISH]).unwrap_err();
        assert_eq!(e2, e);
        let vba = english_translate_text_to_vba("Log 1.", PRELUDE, &[ENGLISH])
            .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert!(vba.contains("Debug.Print 1"), "{vba}");
    }

    /// The web page's language picker loads a dialect after english.vla, as
    /// the add-in loads an edition; each dialect's first proof sentence is
    /// the picker's example, and must read over the base.
    #[test]
    fn every_dialect_loads_over_english_and_reads_its_example() {
        let cases = [
            ("dansk", DANSK, "Saet formlen \"=B2*2\" i cellen B3."),
            (
                "deutsche",
                DEUTSCHE,
                "Setze die Formel \"=B2*2\" in Zelle B3.",
            ),
            (
                "espanol",
                ESPANOL,
                "Pon la formula \"=B2*2\" en la celda B3.",
            ),
            (
                "esperanto",
                ESPERANTO,
                "Metu la formulon \"=B2*2\" en la chelon B3.",
            ),
            (
                "francais",
                FRANCAIS,
                "Mets la formule \"=B2*2\" dans la cellule B3.",
            ),
            ("latin", LATIN, "Pone formulam \"=B2*2\" in cellula B3."),
            (
                "pirate",
                PIRATE,
                "Chart the course \"=B2*2\" onto the cell B3, arr.",
            ),
        ];
        for (name, book, sentence) in cases {
            let vla = english_translate_text_to_vla(sentence, PRELUDE, &[ENGLISH, book])
                .unwrap_or_else(|e| panic!("{name}: {}", e.refusal.text));
            assert!(
                vla.contains("(range \"b3\")") && vla.contains("\"=B2*2\""),
                "{name}: {vla}"
            );
        }
    }

    #[test]
    fn the_gate_asks_capability_first_then_consent() {
        let raw = "(english-vla \"char cell {r:cell}\" (raw \"Debug.Print 2\"))";
        let e = vocab_gate(raw, "pasted.vla", false).unwrap_err();
        assert_eq!(e.id, "english-vocab-raw-consent-declined");
        assert!(
            e.text.starts_with("'pasted.vla' was not loaded"),
            "{}",
            e.text
        );
        assert!(vocab_gate(raw, "pasted.vla", true).is_ok());
        let both = "(requires-capability \"network\")\n(english-vla \"x\" (raw \"y\"))";
        let e = vocab_gate(both, "pasted.vla", true).unwrap_err();
        assert_eq!(e.id, "english-vocab-requires-capability-ungranted");
        assert!(vocab_gate(
            "(english-vla \"wobble {x:expr}\" (debug-print {x}))",
            "p",
            false
        )
        .is_ok());
    }

    #[test]
    fn phrasebooks_load_in_the_order_given_and_the_text_path_is_ungated() {
        // A dialect over the base, as the add-in loads them.
        let vla =
            english_translate_text_to_vla("Pon 5 en la celda B2.", PRELUDE, &[ENGLISH, ESPANOL])
                .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert!(vla.contains("(set! (range \"b2\") 5)"), "{vla}");
        // SEC.2: the text path asks no consent; a door asks before calling.
        let book = "(english-vla \"char cell {r:cell}\" (raw \"Debug.Print 2\"))";
        let vla = english_translate_text_to_vla("Char cell A1.", PRELUDE, &[book])
            .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert!(vla.contains("(raw \"Debug.Print 2\")"), "{vla}");
    }
}
