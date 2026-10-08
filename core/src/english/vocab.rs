//! Phrasebook loading: VLA_SentenceEngine.bas's `EnglishLoadVocabularyText`,
//! `DispatchVocabForm`, `ExpandVocabMacroCall`, `RegisterVocabMacro`,
//! `RegisterFunctionDirective`, the `requires-` scan, and SEC.2's
//! `VocabTextHasRawForm` (PORT.6, slice 6c).
//!
//! A phrasebook is real VLA: every directive a parenthesized form,
//! dispatched by its own head. `<lingua>-vla` and `<lingua>-vla-override`
//! register a phrase rule; `test-success` and `test-fail` are proofs,
//! collected now and run after every rule of the file exists (slice 6d);
//! `defmacro` is carried into every translation; `<lingua>-function`
//! declares a function word or phrase; `keyword-alias` a language's own
//! control-flow word; `requires-` is read from the source text before
//! anything loads; `begin` and `at-row` are wrappers; and any other head
//! is tried once as a call to a macro the file carries, whose expansion
//! is dispatched again.
//!
//! The one thing the text path does not do is ask: the reference's file
//! loader asks consent for a `(raw ...)` form and checks a required
//! capability; its text loader, which `VLA_Browser.bas` calls, does
//! neither. The door decides both here, with [`vocab_text_has_raw_form`]
//! and [`vocab_requires_check_capability`] at hand.

use crate::expand::{expand_text, Expander};
use crate::form::{Form, List};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::printer::write_datum;
use crate::reader::{read_forms, read_forms_with_lines};

use super::grammar::{Grammar, Proof, ProofKind};

/// What one load gathers as it dispatches (`count`, the proof
/// collections and the expanded-form record of `EnglishLoadVocabularyText`).
#[derive(Default)]
struct LoadState {
    count: usize,
    proofs: Vec<Proof>,
    /// GEXPANDER.0: each real directive's text and its row tag.
    expanded: Vec<(String, String)>,
}

impl Grammar {
    /// VBA `EnglishLoadVocabularyText`: the rules a phrasebook text adds,
    /// its macros carried, its function words and aliases registered, its
    /// proofs collected. Returns the number of rules added.
    pub fn load_vocabulary_text(
        &mut self,
        text: &str,
        source_name: &str,
    ) -> Result<usize, Refusal> {
        // F.10: the host-free requirements, before a single rule registers.
        vocab_requires_check_pure(text, source_name)?;
        self.load_source = source_name.to_string();
        let forms = read_forms_with_lines(text)?;
        let mut st = LoadState::default();
        for (f, line) in &forms {
            self.dispatch_vocab_form(f, *line, "", source_name, &mut st, true, 0)?;
        }
        // The file's proofs run against the fully loaded grammar, in file
        // order, positives and negatives alike (slice 6d).
        self.run_file_proofs(std::mem::take(&mut st.proofs))?;
        self.load_source.clear();
        self.set_last_load_source(source_name);
        self.expanded_blob = build_expanded_blob(&st.expanded);
        Ok(st.count)
    }

    /// VBA `DispatchVocabForm`.
    #[allow(clippy::too_many_arguments)]
    fn dispatch_vocab_form(
        &mut self,
        fl: &Form,
        start_line: u32,
        row_tag: &str,
        source_name: &str,
        st: &mut LoadState,
        allow_expansion: bool,
        depth: u32,
    ) -> Result<(), Refusal> {
        let loc = || prov_loc(source_name, start_line, row_tag);
        if depth > 20 {
            return Err(raise("english-vocab-nesting-too-deep", &[("loc", &loc())]));
        }
        let Form::List(flc) = fl else {
            return Err(raise(
                "english-vocab-expected-directive",
                &[
                    ("loc", &loc()),
                    ("word", &fl.atom_text().unwrap_or_default()),
                ],
            ));
        };
        let head = cstr(item(flc, 1)?)?;

        // A generator's own expansion may be (begin dir1 dir2 ...).
        if head == "begin" {
            for j in 2..=flc.count() {
                self.dispatch_vocab_form(
                    item(flc, j)?,
                    start_line,
                    row_tag,
                    source_name,
                    st,
                    false,
                    depth + 1,
                )?;
            }
            return Ok(());
        }
        // LISTOPS-PROVENANCE: (at-row label form), a transparent wrapper.
        if head == "at-row" {
            if flc.count() != 3 {
                return Err(raise("english-vocab-at-row-arity", &[("loc", &loc())]));
            }
            let new_tag = strip_quote_sigil(&cstr(item(flc, 2)?)?);
            return self.dispatch_vocab_form(
                item(flc, 3)?,
                start_line,
                &new_tag,
                source_name,
                st,
                allow_expansion,
                depth + 1,
            );
        }

        if head.len() > 13 && head.ends_with("-vla-override") {
            // G3: an explicit, earned replacement.
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            let pattern = strip_quote_sigil(&cstr(item(flc, 2)?)?);
            self.add_phrase_rule(&pattern, &join_forms_text(flc, 3), true)?;
            st.count += 1;
        } else if head.len() > 4 && head.ends_with("-vla") {
            // F.13: the source language rides in the head on purpose.
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            let pattern = strip_quote_sigil(&cstr(item(flc, 2)?)?);
            self.add_phrase_rule(&pattern, &join_forms_text(flc, 3), false)?;
            st.count += 1;
        } else if head == "test-success" {
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            st.proofs.push(Proof {
                sentence: strip_quote_sigil(&cstr(item(flc, 2)?)?),
                expected: join_forms_text(flc, 3),
                line: start_line,
                kind: ProofKind::Success,
                row_tag: row_tag.to_string(),
                source: source_name.to_string(),
            });
        } else if head == "test-fail" {
            // G4: the negative proof protects the message, not just the
            // refusal, so the fragment must be given.
            let fragment = strip_quote_sigil(&cstr(item(flc, 3)?)?);
            if fragment.is_empty() {
                return Err(raise(
                    "english-vocab-test-fail-empty-fragment",
                    &[("loc", &loc())],
                ));
            }
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            st.proofs.push(Proof {
                sentence: strip_quote_sigil(&cstr(item(flc, 2)?)?),
                expected: fragment,
                line: start_line,
                kind: ProofKind::Fail,
                row_tag: row_tag.to_string(),
                source: source_name.to_string(),
            });
        } else if head == "defmacro" {
            // L4: carried into every translation, appended after the program.
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            self.register_vocab_macro(flc, source_name, start_line, row_tag)?;
        } else if head.len() > 9 && head.ends_with("-function") {
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            self.register_function_directive(flc, &loc())?;
        } else if head == "keyword-alias" {
            // LX5.1: a language file's own control-flow vocabulary.
            st.expanded.push((write_datum(fl), row_tag.to_string()));
            let surface = strip_quote_sigil(&cstr(item(flc, 2)?)?);
            // LX.14: a phrase that uses the word could never match.
            self.refuse_phrase_word_alias(&surface, &loc())?;
            let canonical = strip_quote_sigil(&cstr(item(flc, 3)?)?);
            self.register_keyword_alias(&surface, &canonical);
        } else if head.len() > 9 && head.starts_with("requires-") {
            // F.10: read, checked and refused by the pre-pass; here at top
            // level it passed. Out of an expansion it was never checked.
            if depth > 0 {
                return Err(raise(
                    "english-vocab-requires-from-expansion",
                    &[("loc", &loc()), ("head", &head)],
                ));
            }
            st.expanded.push((write_datum(fl), row_tag.to_string()));
        } else if allow_expansion {
            self.expand_vocab_macro_call(fl, &head, source_name, start_line, row_tag, st, depth)?;
        } else {
            return Err(raise(
                "english-vocab-expansion-not-directive",
                &[("loc", &loc()), ("head", &head)],
            ));
        }
        Ok(())
    }

    /// VBA `ExpandVocabMacroCall` (METAVOCAB): the one real attempt at a
    /// head nothing recognizes - a call to a macro the file carries,
    /// expanded with the prelude as `VlaExpandText` expands it, its one
    /// resulting form dispatched again with no second attempt.
    #[allow(clippy::too_many_arguments)]
    fn expand_vocab_macro_call(
        &mut self,
        fl: &Form,
        head: &str,
        source_name: &str,
        start_line: u32,
        row_tag: &str,
        st: &mut LoadState,
        depth: u32,
    ) -> Result<(), Refusal> {
        let loc = prov_loc(source_name, start_line, row_tag);
        let call_text = write_datum(fl);
        let source = format!("{}\n{call_text}", self.vocab_macros_text());
        let expanded_text = match expand_text(&source, &self.prelude, true) {
            Ok((text, _fired)) => text,
            Err(e) => {
                // A real macro matched, but its own expansion raised: this
                // file's real line, not the synthetic blob's.
                return Err(raise(
                    "english-vocab-macro-expansion-failed",
                    &[("loc", &loc), ("head", head), ("detail", &e.text)],
                ));
            }
        };
        let unrecognized = || {
            raise(
                "english-vocab-unrecognized-directive",
                &[("loc", &loc), ("head", head)],
            )
        };
        let Ok(expanded_forms) = read_forms(&expanded_text) else {
            return Err(unrecognized());
        };
        if expanded_forms.len() != 1 {
            return Err(unrecognized());
        }
        let ef = &expanded_forms[0];
        if !ef.is_list() || write_datum(ef) == call_text {
            return Err(unrecognized()); // no macro matched - unchanged
        }
        self.dispatch_vocab_form(ef, start_line, row_tag, source_name, st, false, depth + 1)
    }

    /// VBA `RegisterVocabMacro` (F.13, L4): one `(defmacro ...)` form,
    /// its name unique among the carried macros, its form proven by the
    /// macro system's own definition check, its text carried.
    fn register_vocab_macro(
        &mut self,
        mac_form: &List,
        source_name: &str,
        start_line: u32,
        row_tag: &str,
    ) -> Result<(), Refusal> {
        let loc = prov_loc(source_name, start_line, row_tag);
        // LISTOPS-PROVENANCE: an at-row-tagged macro's template half is
        // rewrapped in (gen-row "label" ...), so every call carries its row.
        let eff_form: Form = if row_tag.is_empty() {
            Form::List(mac_form.clone())
        } else {
            let body_start = macro_body_start_idx(mac_form);
            let mut wrapped = vec![Form::sym("gen-row"), Form::string(row_tag)];
            for bi in body_start..=mac_form.count() {
                wrapped.push(item(mac_form, bi)?.clone());
            }
            let mut rebuilt = vec![Form::sym("defmacro"), item(mac_form, 2)?.clone()];
            if body_start == 4 {
                rebuilt.push(item(mac_form, 3)?.clone()); // the docstring, untouched
            }
            rebuilt.push(Form::list(wrapped));
            Form::list(rebuilt)
        };
        let mac_text = write_datum(&eff_form);

        let Form::List(name_form) = item(mac_form, 2)? else {
            return Err(vba_runtime_error(424, "Object required"));
        };
        let mac_name = cstr(item(name_form, 1)?)?;
        let key = fold(&mac_name);
        if let Some(prev) = self.vocab_macro_names.get(&key) {
            // U.9: the same path re-carried is a raw re-load without a
            // reset; a different path is the cross-file collision.
            if prev == source_name {
                return Err(raise(
                    "english-vocab-file-already-carried-same",
                    &[("loc", &loc)],
                ));
            }
            return Err(raise(
                "english-vocab-macro-name-collision",
                &[("loc", &loc), ("name", &mac_name), ("prev", prev)],
            ));
        }
        // P-PROBE: VlaProbeMacroForm - the text read again and DefineMacro's
        // own check, with no prelude and no second pass.
        if let Err(e) = probe_macro_form(&mac_text) {
            return Err(raise(
                "english-vocab-macro-form-invalid",
                &[("loc", &loc), ("detail", &e.text)],
            ));
        }
        self.vocab_macros.push(mac_text);
        self.vocab_macro_names.insert(key, source_name.to_string());
        Ok(())
    }

    /// VBA `RegisterFunctionDirective` (LX.14): a `-function` directive,
    /// checked and handed on, its target a name or a form's text.
    fn register_function_directive(&mut self, flc: &List, context: &str) -> Result<(), Refusal> {
        let arity = || {
            raise(
                "english-function-arity",
                &[
                    ("loc", context),
                    ("example", "(english-function \"sum of\" sum-of)"),
                ],
            )
        };
        if flc.count() != 3 {
            return Err(arity());
        }
        if item(flc, 2)?.is_list() {
            return Err(arity());
        }
        let lhs = strip_quote_sigil(&cstr(item(flc, 2)?)?);
        let target = item(flc, 3)?;
        if target.is_list() {
            self.register_function_word(&lhs, &write_datum(target), true, context)
        } else {
            self.register_function_word(&lhs, &cstr(target)?, false, context)
        }
    }
}

/// VBA `VlaProbeMacroForm`: the macro's text read again and each defmacro
/// in it put through `DefineMacro`'s own check.
fn probe_macro_form(mac_text: &str) -> Result<(), Refusal> {
    let forms = read_forms(mac_text)?;
    let mut ex = Expander::new();
    for f in &forms {
        if let Form::List(l) = f {
            if l.head_is("defmacro") {
                ex.define_macro(l)?;
            }
        }
    }
    Ok(())
}

/// VBA `MacroBodyStartIdx`: 4 when the third element is a docstring and a
/// body follows it, else 3.
fn macro_body_start_idx(mac_form: &List) -> usize {
    if mac_form.count() >= 4 {
        if let Some(Form::Str(_)) = mac_form.items.get(2) {
            return 4;
        }
    }
    3
}

/// VBA `ProvLoc`: `<source> line <n>`, with the row label when there is one.
pub fn prov_loc(source_name: &str, line_no: u32, row_tag: &str) -> String {
    let mut r = format!("{source_name} line {line_no}");
    if !row_tag.is_empty() {
        r.push_str(&format!(" ({row_tag})"));
    }
    r
}

/// VBA `JoinFormsText` (F.13): the written forms from a 1-based index on,
/// space-separated.
pub fn join_forms_text(fl: &List, from_idx: usize) -> String {
    fl.items
        .iter()
        .skip(from_idx.saturating_sub(1))
        .map(write_datum)
        .collect::<Vec<_>>()
        .join(" ")
}

/// VBA `StripQuoteSigil`: a string token's content behind its quote mark.
pub fn strip_quote_sigil(s: &str) -> String {
    s.strip_prefix('"').unwrap_or(s).to_string()
}

/// VBA `BuildExpandedBlob` (GEXPANDER.0): one line per real directive, an
/// at-row label as a comment line above the form it labels.
fn build_expanded_blob(expanded: &[(String, String)]) -> String {
    let mut r = String::new();
    for (text, tag) in expanded {
        if !tag.is_empty() {
            r.push_str(&format!("; row: {tag}\r\n"));
        }
        r.push_str(text);
        r.push_str("\r\n");
    }
    r
}

/// VBA `CStr` on a form: a symbol's text, a string's content behind its
/// quote mark, and VBA's own runtime error 13 for a list.
fn cstr(f: &Form) -> Result<String, Refusal> {
    f.atom_text()
        .ok_or_else(|| vba_runtime_error(13, "Type mismatch"))
}

/// VBA `Collection.Item(i)`, 1-based, and its runtime error 9 past the end.
fn item(l: &List, i: usize) -> Result<&Form, Refusal> {
    l.items
        .get(i.wrapping_sub(1))
        .ok_or_else(|| vba_runtime_error(9, "Subscript out of range"))
}

/// The reference meets a malformed directive as a VBA runtime error, not a
/// catalogue refusal; the same words, as a value.
fn vba_runtime_error(number: i64, text: &str) -> Refusal {
    Refusal {
        id: format!("vba-runtime-{number}"),
        number,
        source: "VLA-English".to_string(),
        text: text.to_string(),
    }
}

// ---- F.10: requires- declarations, read from the source text ----------------

/// One `(requires-<namespace> "<value>")` found in the text: the namespace,
/// the quoted value (empty when there is none) and the 0-based index of
/// its `(`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Requirement {
    pub namespace: String,
    pub value: String,
    pub pos: usize,
}

/// VBA `VocabRequiresScan`: every `requires-` head in the text, outside
/// strings and comments, in file order.
pub fn vocab_requires_scan(text: &str) -> Vec<Requirement> {
    let chars: Vec<char> = text.chars().collect();
    let n = chars.len();
    let mut out = Vec::new();
    let mut in_lit = false;
    let mut i = 0;
    while i < n {
        let c = chars[i];
        if in_lit {
            if c == '\\' {
                i += 2;
            } else {
                if c == '"' {
                    in_lit = false;
                }
                i += 1;
            }
        } else if c == '"' {
            in_lit = true;
            i += 1;
        } else if c == ';' {
            while i < n && chars[i] != '\r' && chars[i] != '\n' {
                i += 1;
            }
        } else if c == '(' {
            let mut head_start = i + 1;
            while head_start < n && chars[head_start] == ' ' {
                head_start += 1;
            }
            let mut h_end = head_start;
            while h_end < n && !matches!(chars[h_end], ' ' | '(' | ')' | '\r' | '\n' | '\t') {
                h_end += 1;
            }
            let head = fold(&chars[head_start..h_end].iter().collect::<String>());
            if head.len() > 9 && head.starts_with("requires-") {
                out.push(Requirement {
                    namespace: head[9..].to_string(),
                    value: requires_read_value(&chars, h_end),
                    pos: i,
                });
            }
            i = h_end;
        } else {
            i += 1;
        }
    }
    out
}

/// VBA `RequiresReadValue`: the quoted value directly after a requires-
/// head, or nothing when there isn't one.
fn requires_read_value(chars: &[char], from: usize) -> String {
    let n = chars.len();
    let mut i = from;
    while i < n {
        let c = chars[i];
        if c == '"' {
            break;
        }
        if !matches!(c, ' ' | '\r' | '\n' | '\t') {
            return String::new();
        }
        i += 1;
    }
    if i >= n {
        return String::new();
    }
    i += 1;
    let mut r = String::new();
    while i < n {
        let d = chars[i];
        if d == '\\' {
            if let Some(&e) = chars.get(i + 1) {
                r.push(e);
            }
            i += 2;
        } else if d == '"' {
            return r;
        } else {
            r.push(d);
            i += 1;
        }
    }
    String::new()
}

/// VBA `RequiresLineOf`: the 1-based line of a character position.
fn requires_line_of(text: &str, pos: usize) -> u32 {
    1 + text.chars().take(pos).filter(|&c| c == '\n').count() as u32
}

/// VBA `VocabRequiresCheckPure`: the host-free namespaces, version first
/// whatever the file order, then form and the unknown ones; every exit a
/// refusal, never a dialog.
pub fn vocab_requires_check_pure(text: &str, source_name: &str) -> Result<(), Refusal> {
    let reqs = vocab_requires_scan(text);
    if reqs.is_empty() {
        return Ok(());
    }
    for r in &reqs {
        if r.namespace == "version" {
            requires_check_version(&r.value, text, r.pos, source_name)?;
        }
    }
    for r in &reqs {
        let loc = prov_loc(source_name, requires_line_of(text, r.pos), "");
        match r.namespace.as_str() {
            "version" | "capability" => {}
            "form" => {
                return Err(raise(
                    "english-vocab-requires-form-unsupported",
                    &[("loc", &loc), ("form", &r.value)],
                ))
            }
            other => {
                return Err(raise(
                    "english-vocab-requires-unknown-namespace",
                    &[
                        ("loc", &loc),
                        ("namespace", other),
                        ("have", crate::VERSION),
                    ],
                ))
            }
        }
    }
    Ok(())
}

/// VBA `RequiresCheckVersion`.
fn requires_check_version(
    wanted: &str,
    text: &str,
    pos: usize,
    source_name: &str,
) -> Result<(), Refusal> {
    let loc = prov_loc(source_name, requires_line_of(text, pos), "");
    if wanted.is_empty() {
        return Err(raise(
            "english-vocab-requires-missing-value",
            &[("loc", &loc), ("namespace", "version")],
        ));
    }
    if crate::version::parse(wanted).is_none() {
        return Err(raise(
            "english-vocab-requires-version-malformed",
            &[("loc", &loc), ("wanted", wanted)],
        ));
    }
    if !crate::version::at_least(wanted)? {
        return Err(raise(
            "english-vocab-requires-version-unmet",
            &[("loc", &loc), ("wanted", wanted), ("have", crate::VERSION)],
        ));
    }
    Ok(())
}

/// VBA `VocabRequiresCheckCapability`: the consent-shaped half, which the
/// reference's file loader runs and its text loader does not. Today every
/// capability refuses, since nothing can grant one yet (SEC.7).
pub fn vocab_requires_check_capability(text: &str, source_name: &str) -> Result<(), Refusal> {
    for r in vocab_requires_scan(text) {
        if r.namespace == "capability" {
            let loc = prov_loc(source_name, requires_line_of(text, r.pos), "");
            if r.value.is_empty() {
                return Err(raise(
                    "english-vocab-requires-missing-value",
                    &[("loc", &loc), ("namespace", "capability")],
                ));
            }
            return Err(raise(
                "english-vocab-requires-capability-ungranted",
                &[("loc", &loc), ("capability", &r.value)],
            ));
        }
    }
    Ok(())
}

/// VBA `VocabTextHasRawForm` (SEC.2): does the text carry a `(raw ...)`
/// form anywhere, read by every paren's own head outside strings and
/// comments?
pub fn vocab_text_has_raw_form(text: &str) -> bool {
    let chars: Vec<char> = text.chars().collect();
    let n = chars.len();
    let mut in_lit = false;
    let mut i = 0;
    while i < n {
        let c = chars[i];
        if in_lit {
            if c == '\\' {
                i += 2;
            } else {
                if c == '"' {
                    in_lit = false;
                }
                i += 1;
            }
        } else if c == '"' {
            in_lit = true;
            i += 1;
        } else if c == ';' {
            while i < n && chars[i] != '\r' && chars[i] != '\n' {
                i += 1;
            }
        } else if c == '(' {
            let mut j = i + 1;
            while j < n && chars[j] == ' ' {
                j += 1;
            }
            let mut h_end = j;
            while h_end < n && !matches!(chars[h_end], ' ' | '(' | ')' | '\r' | '\n' | '\t') {
                h_end += 1;
            }
            if fold(&chars[j..h_end].iter().collect::<String>()) == "raw" {
                return true;
            }
            i = h_end;
        } else {
            i += 1;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../../scripts/polyglotta/english.vla");
    const ESPANOL: &str = include_str!("../../../scripts/polyglotta/espanol.vla");
    const ALIEN: &str = include_str!("../../../scripts/polyglotta/alien.vla");
    const DANSK: &str = include_str!("../../../scripts/polyglotta/dansk.vla");
    const DEUTSCHE: &str = include_str!("../../../scripts/polyglotta/deutsche.vla");
    const ESPERANTO: &str = include_str!("../../../scripts/polyglotta/esperanto.vla");
    const FRANCAIS: &str = include_str!("../../../scripts/polyglotta/francais.vla");
    const LATIN: &str = include_str!("../../../scripts/polyglotta/latin.vla");
    const PIRATE: &str = include_str!("../../../scripts/polyglotta/pirate.vla");

    /// A grammar that keeps a failing proof instead of refusing the load:
    /// these tests are about loading, and until slice 6e a proof whose
    /// sentence begins with a built-in statement form cannot pass.
    fn fresh() -> Grammar {
        let mut g = Grammar::new(PRELUDE);
        g.set_proof_mode(super::super::grammar::ProofMode::Collect);
        g
    }

    fn load(g: &mut Grammar, text: &str, source: &str) -> Result<usize, Refusal> {
        g.load_vocabulary_text(text, source)
    }

    fn refused(text: &str, source: &str) -> Refusal {
        let mut g = fresh();
        let e = load(&mut g, text, source).unwrap_err();
        assert!(!e.text.starts_with("RaiseMsg:"), "{}", e.text);
        e
    }

    fn proofs(g: &Grammar) -> (usize, usize) {
        let s = g
            .proofs()
            .iter()
            .filter(|p| p.kind == ProofKind::Success)
            .count();
        let f = g
            .proofs()
            .iter()
            .filter(|p| p.kind == ProofKind::Fail)
            .count();
        (s, f)
    }

    #[test]
    fn english_vla_loads_as_the_reference_counts_it() {
        // The owner's reading of 2026-10-02: 240 rules, 220 macros, 460; L-SHEET-HELPERS
        // (2026-10-07) added eleven sheet macros, so 231 macros. 460
        // test-success and 22 test-fail proofs (EnglishVocabStats, read
        // after EnglishLoadVocabulary).
        let mut g = fresh();
        let n = load(&mut g, ENGLISH, "english.vla").unwrap_or_else(|e| panic!("{}", e.text));
        assert_eq!(n, 240);
        assert_eq!(g.rule_count(), 251);
        assert_eq!(g.vocab_macro_count(), 231);
        assert_eq!(proofs(&g), (460, 22));
        assert_eq!(g.fn_of_count(), 20);
        assert_eq!(g.phrases().len(), 8);
        assert_eq!(g.keyword_alias_count(), 0);
        assert_eq!(
            g.vocab_stats(),
            "loaded: 240 rules, 231 macros, 460 tests (22 expected fails) from english.vla"
        );
        assert!(g.lint_warnings().is_empty(), "{:?}", g.lint_warnings());
        // The three generators made their rules and proofs.
        assert!(g
            .rules()
            .iter()
            .any(|r| r.text == "set style of table {n:text} to {s:text}"));
        assert!(g
            .proofs()
            .iter()
            .any(|p| p.sentence == "Hide the total row of table Sales."));
        assert!(g
            .vocab_macros_text()
            .contains("(defmacro (unhide-column c)"));
        let shapes: Vec<&str> = g.phrases().iter().map(|p| p.shape.as_str()).collect();
        assert!(shapes.contains(&"standard deviation of ... as sample|population"));
        assert!(shapes.contains(&"last filled row of column <column>"));
        // The expanded record is one line per real directive.
        assert_eq!(g.expanded_blob().lines().count(), 240 + 231 + 460 + 22 + 14);
    }

    #[test]
    fn every_dialect_loads_alone_as_the_reference_loads_it() {
        // Seven dialects register their rules, and their proofs run at load
        // and pass whole (stmt.rs's test proves each in Raise mode). Until
        // LX.15 (2026-10-02) the reference refused five of them at load:
        // esperanto, latin and deutsche registered their general put rule
        // before their literal today rule (espanol fixed its own under
        // LX.10), deutsche and francais carried F.4's noise-word defect
        // ('an', 'a' before a slot), which the lint named here, and pirate
        // glued a comma to a slot. LX.15 mended the files, so each loads
        // alone, every proof registered, no warning.
        for (text, name, rules, macros, ok, fail, warnings) in [
            (DANSK, "dansk.vla", 19, 0, 20, 0, 0),
            (DEUTSCHE, "deutsche.vla", 19, 0, 20, 0, 0),
            (ESPERANTO, "esperanto.vla", 19, 0, 20, 0, 0),
            (FRANCAIS, "francais.vla", 19, 0, 20, 0, 0),
            (LATIN, "latin.vla", 19, 0, 20, 0, 0),
            (PIRATE, "pirate.vla", 19, 0, 20, 0, 0),
            (ESPANOL, "espanol.vla", 19, 134, 29, 1, 0),
        ] {
            let mut g = fresh();
            let n = load(&mut g, text, name).unwrap_or_else(|e| panic!("{name}: {}", e.text));
            assert_eq!(n, rules, "{name} rules");
            assert_eq!(g.vocab_macro_count(), macros, "{name} macros");
            assert_eq!(proofs(&g), (ok, fail), "{name} proofs");
            assert_eq!(
                g.lint_warnings().len(),
                warnings,
                "{name}: {:?}",
                g.lint_warnings()
            );
        }
        // F.4's lint still names the shape francais.vla carried until LX.15
        // (TestF4NoiseWordBeforeSlot keeps the same rule as its fixture).
        let mut g = fresh();
        load(
            &mut g,
            "(francais-vla \"envoie un courriel a {who:expr} avec objet {s:expr} et message {m:expr}\" (vlasendmail {who} {s} {m}))",
            "f4-fixture.vla",
        )
        .unwrap();
        assert!(
            g.lint_warnings()[0].contains("noise word 'a' immediately precedes slot {who:expr}")
        );
        // alien.vla is a library of macros a program includes, not a
        // phrasebook: no rule, no proof (the treaty inventories it and never
        // scores it, LX.15), and its one top-level call expands to a
        // (sub ...), which the loader refuses as no directive.
        let e = refused(ALIEN, "alien.vla");
        assert_eq!(e.id, "english-vocab-expansion-not-directive");
        assert!(
            e.text
                .starts_with("alien.vla line 117: a generator's expansion produced 'sub'"),
            "{}",
            e.text
        );
        let mut g = fresh();
        load(&mut g, ESPANOL, "espanol.vla").unwrap();
        assert_eq!(g.keyword_alias_count(), 5);
        assert_eq!(g.keyword_alias("si"), Some("if"));
        // "la suma de" is three words before its value: a phrase, as the
        // VBA registers it, not a one-word head ("la" is no English noise
        // word).
        assert_eq!(g.fn_of_count(), 14);
        assert_eq!(g.phrases().len(), 8);
        assert!(g.phrases().iter().any(|p| p.shape == "la suma de ..."));
        assert!(g
            .phrases()
            .iter()
            .any(|p| p.shape == "la suma de la region <range>"));
        // A second file loads beside the first; a reset clears both.
        load(&mut g, DANSK, "dansk.vla").unwrap();
        assert_eq!(g.rule_count(), 11 + 19 + 19);
        g.reset();
        assert_eq!(g.rule_count(), 11);
        assert_eq!(g.vocab_macro_count(), 0);
    }

    #[test]
    fn overrides_and_duplicates_as_testg3_pins_them() {
        let mut g = fresh();
        load(&mut g, "(english-vla \"warm cell {r:cell}\" (debug-print 1))\n(test-success \"Warm cell B2.\" (debug-print 1))", "base-vocab").unwrap();
        load(&mut g, "(english-vla-override \"warm cell {r:cell}\" (debug-print 2))\n(test-success \"Warm cell B2.\" (debug-print 2))", "dialect-vocab").unwrap();
        assert_eq!(g.rule_count(), 12);
        assert_eq!(g.rule_source(12), "dialect-vocab (overrides base-vocab)");
        let e = refused("(english-vla \"warm cell {r:cell}\" (debug-print 1))\n(english-vla \"warm cell {r:cell}\" (debug-print 2))", "base-vocab");
        assert!(
            e.text.contains("duplicates") && e.text.contains("-vla-override"),
            "{}",
            e.text
        );
        let e = refused(
            "(english-vla-override \"warm cell {r:cell}\" (debug-print 2))",
            "dialect-vocab",
        );
        assert!(e.text.contains("matches no earlier rule"));
        let e = refused("(english-vla \"fix cell {r:range}\" (debug-print 1))\n(english-vla \"fix range {r:range}\" (debug-print 2))\n(english-vla-override \"fix {w:cell|range} {r:range}\" (debug-print 3))", "v");
        assert!(e.text.contains("ambiguous"));
        let e = refused(
            "(english-vla-override \"set {v:var} to {e:expr}\" (debug-print 9))",
            "v",
        );
        assert!(e.text.contains("not overridable"));
        // An in-file override replaces its own base.
        let mut g = fresh();
        load(&mut g, "(english-vla \"warm cell {r:cell}\" (debug-print 1))\n(english-vla-override \"warm cell {r:cell}\" (debug-print 5))", "one-file-vocab").unwrap();
        assert_eq!(g.rule_count(), 12);
        assert_eq!(
            g.rules()[11].forms[0],
            Form::list(vec![Form::sym("debug-print"), Form::sym("5")])
        );
    }

    #[test]
    fn carried_macros_as_testl4_and_the_probe_pin_them() {
        let mut g = fresh();
        load(&mut g, "(defmacro (vla-mark r) (begin (set! (range r) 1) (set! (. (range r) font.bold) true)))\n(english-vla \"mark cell {r:cell}\" (vla-mark {r}))\n(test-success \"Mark cell B2.\" (vla-mark \"b2\"))", "selftest-vocab").unwrap();
        assert!(g
            .vocab_macros_text()
            .starts_with("(defmacro (vla-mark r) (begin"));
        load(&mut g, "(defmacro (vla-a x)\n    (set! x 1))", "base-vocab").unwrap();
        load(&mut g, "(defmacro (vla-b x) (set! x 2))", "dialect-vocab").unwrap();
        assert_eq!(g.vocab_macros_text(), "(defmacro (vla-mark r) (begin (set! (range r) 1) (set! (. (range r) font.bold) true)))\r\n(defmacro (vla-a x) (set! x 1))\r\n(defmacro (vla-b x) (set! x 2))");
        assert_eq!(g.vocab_macro_count(), 3);
        g.reset();
        assert_eq!(g.vocab_macros_text(), "");
        let mut g = fresh();
        load(&mut g, "(defmacro (vla-a x) (set! x 1))", "base-vocab").unwrap();
        let e = load(&mut g, "(defmacro (vla-a x) (set! x 2))", "dialect-vocab").unwrap_err();
        assert_eq!(e.id, "english-vocab-macro-name-collision");
        assert!(
            e.text.contains("already carried by base-vocab"),
            "{}",
            e.text
        );
        let e = load(&mut g, "(defmacro (vla-a x) (set! x 2))", "base-vocab").unwrap_err();
        assert_eq!(e.id, "english-vocab-file-already-carried-same");
        let e = refused(
            "(defmacro (vla-c & body :rescue & h) (begin body))",
            "selftest-vocab",
        );
        assert_eq!(e.id, "english-vocab-macro-form-invalid");
        assert!(e.text.contains("rest parameter must be last"), "{}", e.text);
        let e = refused("(set! x 1)", "selftest-vocab");
        assert_eq!(e.id, "english-vocab-unrecognized-directive");
        assert!(e.text.contains("unrecognized top-level directive"));
        let e = refused("(defmacro (vla-d x)\n    (set! x 1)", "selftest-vocab");
        assert_eq!(e.id, "vla-unbalanced-parens");
        let e = refused("(defmacro (probe-ok-a x) (set! x 1))\n(defmacro (probe-ok-b x) (set! x 2))\n(defmacro (quote x) x)", "probe-vocab");
        assert!(
            e.text.contains("does not stand") && e.text.contains("quote"),
            "{}",
            e.text
        );
        assert!(e.text.starts_with("probe-vocab line 3: "), "{}", e.text);
    }

    #[test]
    fn generators_expand_as_testmetavocab_pins_them() {
        let mut g = fresh();
        let n = load(&mut g, "(defmacro (vocab-antonym-pair pat1 call1 pat2 call2)\n    (begin (english-vla pat1 call1) (english-vla pat2 call2)))\n(vocab-antonym-pair\n    \"warm cell {r:cell}\" (debug-print {r})\n    \"cool cell {r:cell}\" (debug-print {r}))\n(test-success \"Warm cell B2.\" (debug-print \"b2\"))\n(test-success \"Cool cell B3.\" (debug-print \"b3\"))", "metavocab-vocab").unwrap();
        assert_eq!(n, 2);
        assert_eq!(g.rules()[11].text, "warm cell {r:cell}");
        assert_eq!(g.rules()[12].text, "cool cell {r:cell}");
        let e = refused("(zzz-not-a-thing \"x\")", "metavocab-vocab");
        assert_eq!(
            e.text,
            "metavocab-vocab line 1: unrecognized top-level directive 'zzz-not-a-thing'"
        );
        let e = refused("(defmacro (oops x) (+ x 1))\n(oops 5)", "metavocab-vocab");
        assert_eq!(e.id, "english-vocab-expansion-not-directive");
        assert!(e.text.contains("is not a vocabulary directive"));
        let e = refused(
            "(defmacro (needs-two a b) (english-vla a b))\n\n(needs-two \"only one\")",
            "metavocab-vocab",
        );
        assert_eq!(e.id, "english-vocab-macro-expansion-failed");
        assert!(
            e.text.contains("line 3")
                && e.text.contains("needs-two")
                && e.text.contains("expects 2 argument"),
            "{}",
            e.text
        );
        let mut g = fresh();
        let n = load(&mut g, "(defmacro (one-rule pat call) (english-vla pat call))\n(defmacro (two-rules pat1 call1 pat2 call2)\n    (begin (one-rule pat1 call1) (one-rule pat2 call2)))\n(two-rules\n    \"glow cell {r:cell}\" (debug-print {r})\n    \"fade cell {r:cell}\" (debug-print {r}))", "metavocab-vocab").unwrap();
        assert_eq!(n, 2);
        let mut g = fresh();
        let n = load(&mut g, "(defmacro (nested-begin-rule)\n    (begin (begin (english-vla \"spark cell {r:cell}\" (debug-print {r})))\n           (test-success \"Spark cell B2.\" (debug-print \"b2\"))))\n(nested-begin-rule)", "metavocab-vocab").unwrap();
        assert_eq!(n, 1);
        assert_eq!(proofs(&g), (1, 0));
    }

    #[test]
    fn at_row_tags_ride_as_testatrow_pins_them() {
        let mut g = fresh();
        let n = load(&mut g, "(english-vla \"warm cell {r:cell}\" (debug-print {r}))\n(at-row \"row A\" (test-success \"Warm cell B2.\" (debug-print \"b2\")))", "atrow-vocab").unwrap();
        assert_eq!(n, 1);
        assert_eq!(g.proofs()[0].row_tag, "row A");
        assert!(g.expanded_blob().contains("; row: row A\r\n(test-success"));
        let e = refused("(defmacro (needs-two a b) (english-vla a b))\n\n(at-row \"row 9\" (needs-two \"only one\"))", "atrow-vocab");
        assert!(
            e.text.contains("row 9") && e.text.contains("needs-two"),
            "{}",
            e.text
        );
        let e = refused("(at-row \"row 3\" (defmacro (quote x) x))", "atrow-vocab");
        assert!(e.text.contains("row 3"), "{}", e.text);
        let e = refused("(at-row \"onlyonearg\")", "atrow-vocab");
        assert!(e.text.contains("takes exactly two arguments"));
        let e = refused("(zzz-not-a-thing \"x\")", "atrow-vocab");
        assert!(!e.text.contains('('), "{}", e.text);
        // A tagged macro's template is rewrapped in gen-row.
        let mut g = fresh();
        load(&mut g, "(at-row \"42\" (defmacro (tagged-set-style n s) \"doc\" (set! (. (activesheet.listobjects n) tablestyle) s)))", "atrow-vocab").unwrap();
        assert_eq!(g.vocab_macros_text(), "(defmacro (tagged-set-style n s) \"doc\" (gen-row \"42\" (set! (. (activesheet.listobjects n) tablestyle) s)))");
        let mut g = fresh();
        load(&mut g, "(defmacro (plain-set-style n s) \"doc\" (set! (. (activesheet.listobjects n) tablestyle) s))", "atrow-vocab").unwrap();
        assert!(!g.vocab_macros_text().contains("gen-row"));
        // Too deep a nesting of generators' begins.
        let deep = format!(
            "{}(english-vla \"z {{r:cell}}\" (x)){}",
            "(begin ".repeat(21),
            ")".repeat(21)
        );
        let e = refused(&deep, "v");
        assert_eq!(e.id, "english-vocab-nesting-too-deep");
        let e = refused("5", "v");
        assert_eq!(e.id, "english-vocab-expected-directive");
        assert!(e.text.ends_with("got a bare word '5'"));
        let e = refused("(test-fail \"Warm cell banana.\" \"\")", "selftest-vocab");
        assert!(e.text.contains("non-empty fragment"));
    }

    #[test]
    fn requires_as_testf10_pins_them() {
        let mut g = fresh();
        assert!(load(&mut g, "(requires-version \"0.0.0\")\n(english-vla \"zzmet cell {r:text}\" (set! (range {r}) 1))", "selftest-vocab").is_ok());
        let e = refused("(requires-version \"999.0.0\")", "selftest-vocab");
        assert_eq!(e.id, "english-vocab-requires-version-unmet");
        assert!(
            e.text.contains("needs Frazaro 999.0.0") && e.text.contains(crate::VERSION),
            "{}",
            e.text
        );
        assert!(e.text.starts_with("selftest-vocab line 1: "));
        let e = refused("(requires-version \"banana\")", "v");
        assert!(e.text.contains("is not a version"));
        let e = refused("(requires-version)", "v");
        assert!(e.text.contains("names no value"));
        let e = refused("(requires-signature \"acme\")", "v");
        assert!(e.text.contains("is not a kind of requirement"));
        let e = refused(
            "(requires-signature \"acme\")\n(requires-version \"999.0.0\")",
            "v",
        );
        assert!(e.text.contains("needs Frazaro 999.0.0"));
        // Capability is not checked on the text path; a door checks it.
        let mut g = fresh();
        let cap = "(requires-capability \"sendmail\")\n(english-vla \"zzcap cell {r:text}\" (set! (range {r}) 1))";
        assert!(load(&mut g, cap, "v").is_ok());
        let e = vocab_requires_check_capability(cap, "v").unwrap_err();
        assert_eq!(e.id, "english-vocab-requires-capability-ungranted");
        assert!(e.text.contains("'sendmail'"));
        let e = vocab_requires_check_capability("(requires-capability)", "v").unwrap_err();
        assert_eq!(e.id, "english-vocab-requires-missing-value");
        let e = refused("(requires-form \"paint cell\")", "v");
        assert!(e.text.contains("not yet enforceable"));
        let mut g = fresh();
        assert!(load(&mut g, "; (requires-version \"999.0.0\")\n(english-vla \"zzcomment cell {r:text}\" (set! (range {r}) 1))", "v").is_ok());
        assert!(load(&mut g, "(english-vla \"zzsay cell {r:text}\"\n    (set! (range {r}) \"(requires-version \\\"999.0.0\\\")\"))", "v").is_ok());
        // An unmet requirement refuses before a single rule registers.
        let mut g = fresh();
        assert!(load(&mut g, "(english-vla \"zzlate cell {r:text}\" (set! (range {r}) 1))\n(requires-version \"999.0.0\")", "v").is_err());
        assert_eq!(g.rule_count(), 11);
        let e = refused(
            "(defmacro (zzgen) \"g\" (requires-version \"0.0.0\"))\n(zzgen)",
            "v",
        );
        assert!(e.text.contains("produced by a generator's expansion"));
        let reqs = vocab_requires_scan(
            "x (requires-version \"1.2.3\") ; (requires-x \"n\")\n(REQUIRES-Form   \"a\\\"b\")",
        );
        assert_eq!(reqs.len(), 2);
        assert_eq!(
            (
                reqs[0].namespace.as_str(),
                reqs[0].value.as_str(),
                reqs[0].pos
            ),
            ("version", "1.2.3", 2)
        );
        assert_eq!(
            (reqs[1].namespace.as_str(), reqs[1].value.as_str()),
            ("form", "a\"b")
        );
    }

    #[test]
    fn function_directives_keyword_aliases_and_raw_forms() {
        let mut g = fresh();
        load(&mut g, "(english-function \"zeta of range {r:range}\" (zeta-of (range {r})))\n(english-function \"zeta de\" zeta-de)\n(english-function \"zeta\" zeta-now)", "v").unwrap();
        assert_eq!(g.phrases().len(), 2);
        assert_eq!(g.fn_nullary_target("zeta"), Some("zeta-now"));
        let e = refused("(english-function \"zeta of\")", "v");
        assert_eq!(e.id, "english-function-arity");
        assert!(e.text.contains("takes a quoted pattern and one target"));
        let e = refused("(english-function (zeta) zeta)", "v");
        assert_eq!(e.id, "english-function-arity");
        let e = refused("(keyword-alias \"zeta\" \"if\")\n(english-function \"zeta of range {r:range}\" (zeta (range {r})))", "v");
        assert!(
            e.text
                .contains("'zeta' is both a keyword alias and a word of the phrase"),
            "{}",
            e.text
        );
        let e = refused("(english-function \"lorem ipsum of {x:value}\" (lorem {x}))\n(keyword-alias \"ipsum\" \"if\")", "v");
        assert!(
            e.text
                .contains("'ipsum' is both a keyword alias and a word of the phrase"),
            "{}",
            e.text
        );
        assert!(vocab_text_has_raw_form(
            "(english-vla \"x\" (raw \"' hi\"))"
        ));
        assert!(vocab_text_has_raw_form(
            "(english-vla \"x\" ( RAW \"' hi\"))"
        ));
        assert!(!vocab_text_has_raw_form(
            "; (raw \"x\")\n(english-vla \"raw deal\" (set! raw 1))"
        ));
        assert!(!vocab_text_has_raw_form(ENGLISH));
        let e = refused("(english-vla (x) (y))", "v");
        assert_eq!(
            (e.id.as_str(), e.text.as_str()),
            ("vba-runtime-13", "Type mismatch")
        );
        let e = refused("(test-success)", "v");
        assert_eq!(e.text, "Subscript out of range");
    }
}
