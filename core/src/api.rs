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
    crate::build::build_xlsx(program_text, &t.vla, prelude_text, vocab_texts)
        .map_err(|refusal| RefusalAtLine { refusal, line: 0 })
}

/// `frazaro rebuild` (PORT.7, slice 7c): a built workbook's sentences read
/// back out of its `Frazaro` sheet, its stamp checked against them and
/// against the prelude and phrasebooks given, the workbook built again from
/// them and compared whole. The refusal is `rebuild-not-a-build` for a file
/// that is not a build this core can verify; a build that does not match is
/// not a refusal but an answer, `Rebuilt::matches` false with its reason.
pub fn english_rebuild_xlsx(
    file: &[u8],
    prelude_text: &str,
    vocab_texts: &[&str],
) -> Result<crate::build::Rebuilt, RefusalAtLine> {
    use crate::build::{build_workbook_with, own_part_text, Rebuilt};
    use crate::sheet::merge::HostInfo;
    use crate::sheet::ooxml::{sheet_part, sheet_xml, Render};

    let read =
        crate::build::read_build(file).map_err(|refusal| RefusalAtLine { refusal, line: 0 })?;
    let partial = read.stamp.into.is_some();
    let answer = |matches: bool, why: String| Rebuilt {
        sentences: read.sentences,
        built_by: read.stamp.version.clone(),
        matches,
        why,
        partial,
    };
    if let Some(why) = read
        .stamp
        .disagreement(&read.program_text, prelude_text, vocab_texts)
    {
        return Ok(answer(false, why));
    }
    let version_note = || {
        if read.stamp.version != crate::VERSION {
            format!(
                "it was built by Frazaro {} and this is Frazaro {}, which writes the file differently",
                read.stamp.version,
                crate::VERSION
            )
        } else {
            "its parts are not what this core builds from these sentences: a host has saved it since, or it was built differently".to_string()
        }
    };
    match &read.stamp.into {
        None => {
            let again = english_build_xlsx(&read.program_text, prelude_text, vocab_texts)?;
            let matches = again == file;
            Ok(answer(
                matches,
                if matches {
                    String::new()
                } else {
                    version_note()
                },
            ))
        }
        Some(into) => {
            // The model is not at hand, so the build's own sheets are rendered
            // again with what the stamp recorded and compared part by part.
            let host = HostInfo {
                label: "the workbook".to_string(),
                sheet_names: read.host_sheets.clone(),
                model_hex: into.model_hex.clone(),
                style_base: into.style_base,
                cm: Some(into.cm),
            };
            let g = load_grammar(prelude_text, vocab_texts)?;
            let t = g.translate_program_at(&read.program_text)?;
            let wb = build_workbook_with(
                &read.program_text,
                &t.vla,
                prelude_text,
                vocab_texts,
                Some(&host),
            )
            .map_err(|refusal| RefusalAtLine { refusal, line: 0 })?;
            let render = Render {
                style_base: into.style_base,
                cm: into.cm,
            };
            let mut matches = true;
            for (i, sheet) in wb.sheets.iter().enumerate() {
                let want = sheet_xml(sheet, false, &render);
                if own_part_text(file, &sheet_part(i + 1)).as_deref() != Some(want.as_str()) {
                    matches = false;
                }
            }
            if own_part_text(file, &sheet_part(wb.sheets.len() + 1)).is_some() {
                matches = false; // a sheet the build did not make
            }
            Ok(answer(
                matches,
                if matches {
                    String::new()
                } else {
                    version_note()
                },
            ))
        }
    }
}

/// The writer's surface for adding to a workbook (`--into`): the program
/// translated and built into `model_bytes`, the model's own parts as they
/// were. `label` is what a refusal calls the model.
pub fn english_build_xlsx_into(
    program_text: &str,
    prelude_text: &str,
    vocab_texts: &[&str],
    model_bytes: Vec<u8>,
    label: &str,
) -> Result<Vec<u8>, RefusalAtLine> {
    let g = load_grammar(prelude_text, vocab_texts)?;
    let t = g.translate_program_at(program_text)?;
    crate::build::build_xlsx_into(
        program_text,
        &t.vla,
        prelude_text,
        vocab_texts,
        model_bytes,
        label,
    )
    .map_err(|refusal| RefusalAtLine { refusal, line: 0 })
}

/// The view record (KERNEL.4): the program translated and built into the
/// sheet model as [`english_build_xlsx`] builds it, nothing written, and one
/// window of one sheet projected as the lines a viewport draws from
/// (`view::Grid`, the first projection through the kernel's projections
/// seam). `sheet` is the sheet's name, compared as Excel compares names;
/// `window` an A1 rectangle, or `None` for the sheet's whole extent. The
/// English stage's refusal with its line, or the build's or the view's
/// (`view-sheet-unknown`, `view-window-not-a-range`) with line 0.
pub fn english_view(
    program_text: &str,
    prelude_text: &str,
    vocab_texts: &[&str],
    sheet: &str,
    window: Option<&str>,
) -> Result<String, RefusalAtLine> {
    let g = load_grammar(prelude_text, vocab_texts)?;
    let t = g.translate_program_at(program_text)?;
    let at_line_0 = |refusal| RefusalAtLine { refusal, line: 0 };
    let model = crate::build::build_workbook(program_text, &t.vla, prelude_text, vocab_texts)
        .map_err(at_line_0)?;
    let window = crate::view::window_of(&model, sheet, window).map_err(at_line_0)?;
    Ok(crate::view::view_text(&model, &window))
}

/// The reader's surface (PORT.8, slice 8a): a workbook's bytes, an OOXML
/// package or an OpenDocument spreadsheet (8e), as the relations of
/// `REFLECT`, one row a line in the fixed order (`reflect::reflect_text`);
/// `label` is what a refusal calls the file. Seven of the eight relations,
/// as the module says; `refers` joins in slice 8b.
pub fn reflect_relations(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    crate::reflect::reflect_text(bytes, label)
}

/// The difference between two workbooks (PORT.8, slice 8c): both files'
/// bytes read through the reader and compared, the `changed` relation and
/// the sheets in one file alone, one row a line in the fixed order
/// (`reflect::diff::diff_text`); each label is what a refusal calls its
/// file. Empty when the two hold the same sheets, names, Tables and cells.
pub fn diff_relations(
    old_bytes: &[u8],
    old_label: &str,
    new_bytes: &[u8],
    new_label: &str,
) -> Result<String, Refusal> {
    crate::reflect::diff::diff_text(old_bytes, old_label, new_bytes, new_label)
}

/// The audit list (PORT.8, slice 8d): a workbook's bytes read through the
/// reader into the audit's index and the six walks run, one finding a line
/// in the fixed order (`reflect::audit::audit_text`); `label` is what a
/// refusal calls the file. Empty when there is nothing to report.
pub fn audit_findings(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    crate::reflect::audit::audit_text(bytes, label)
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

    const FIXTURE: &str = include_str!("../../scripts/build/fixture.txt");
    const INTO: &str = include_str!("../../scripts/build/into.txt");

    /// A golden as the door prints it: the files are CRLF on disk.
    fn golden(text: &str) -> String {
        text.replace("\r\n", "\n")
    }

    /// Oracle 11 through the API (KERNEL.4). The test is here, and not with
    /// the view in `vla-lang`, since it builds through English (PORT.12).
    #[test]
    fn the_view_goldens_are_reproduced() {
        for (program, sheet, window, golden_text, path) in [
            (
                FIXTURE,
                "Frazaro",
                None,
                include_str!("../../scripts/view/fixture_frazaro.vla"),
                "scripts/view/fixture_frazaro.vla",
            ),
            (
                FIXTURE,
                "Output",
                Some("A1:F20"),
                include_str!("../../scripts/view/fixture_output.vla"),
                "scripts/view/fixture_output.vla",
            ),
            (
                FIXTURE,
                "Output",
                Some("B2:C3"),
                include_str!("../../scripts/view/fixture_output_b2_c3.vla"),
                "scripts/view/fixture_output_b2_c3.vla",
            ),
            (
                FIXTURE,
                "Data",
                None,
                include_str!("../../scripts/view/fixture_data.vla"),
                "scripts/view/fixture_data.vla",
            ),
            (
                INTO,
                "Output",
                None,
                include_str!("../../scripts/view/into_output.vla"),
                "scripts/view/into_output.vla",
            ),
            (
                INTO,
                "Checks",
                None,
                include_str!("../../scripts/view/into_checks.vla"),
                "scripts/view/into_checks.vla",
            ),
        ] {
            let got = english_view(program, PRELUDE, &[ENGLISH], sheet, window)
                .unwrap_or_else(|r| panic!("{path}: {}", r.refusal));
            let want = golden(golden_text);
            if got != want {
                let at = got
                    .lines()
                    .zip(want.lines())
                    .position(|(a, b)| a != b)
                    .map(|i| i + 1)
                    .unwrap_or(got.lines().count().min(want.lines().count()) + 1);
                panic!(
                    "{path} differs at line {at}: got {:?}, want {:?}",
                    got.lines().nth(at - 1),
                    want.lines().nth(at - 1)
                );
            }
        }
    }

    /// The free oracle (web/CALLOSUM.md section 8, slice 1): the window is
    /// drawn from the model, the file is written from the model, and the
    /// reader reads the file; the cell and formula rows agree row for row,
    /// and the sheet rows too.
    #[test]
    fn the_view_of_a_built_model_is_the_reflect_of_its_file() {
        use crate::kernel::Window;
        use crate::reflect::print::quoted;
        use crate::view::{view_text, whole};
        for (program, label) in [(FIXTURE, "fixture.txt"), (INTO, "into.txt")] {
            let vla = english_translate_text_to_vla(program, PRELUDE, &[ENGLISH])
                .unwrap_or_else(|r| panic!("{label}: {}", r.refusal));
            let model = crate::build::build_workbook(program, &vla, PRELUDE, &[ENGLISH])
                .unwrap_or_else(|r| panic!("{label}: {r}"));
            let bytes = crate::sheet::ooxml::workbook_bytes(&model).unwrap();
            let reflected = crate::reflect::reflect_text(&bytes, label).unwrap();
            let file_sheets: Vec<&str> = reflected
                .lines()
                .filter(|l| l.starts_with("(sheet "))
                .collect();
            let mut seen = 0;
            for sheet in &model.sheets {
                let view = view_text(
                    &model,
                    &Window {
                        sheet: sheet.name.clone(),
                        range: whole(sheet),
                    },
                );
                let view_sheets: Vec<&str> =
                    view.lines().filter(|l| l.starts_with("(sheet ")).collect();
                assert_eq!(view_sheets, file_sheets, "{label}: the sheet rows");
                let body: Vec<&str> = view
                    .lines()
                    .filter(|l| l.starts_with("(cell ") || l.starts_with("(formula "))
                    .collect();
                let prefix_cell = format!("(cell {} ", quoted(&sheet.name));
                let prefix_formula = format!("(formula {} ", quoted(&sheet.name));
                let read: Vec<&str> = reflected
                    .lines()
                    .filter(|l| l.starts_with(&prefix_cell) || l.starts_with(&prefix_formula))
                    .collect();
                assert_eq!(body, read, "{label}: sheet {}", sheet.name);
                seen += body.len();
            }
            assert!(seen > 10, "{label}: the comparison covered {seen} rows");
        }
    }
}
