//! The grammar: VLA_SentenceEngine.bas's rule store and the tables a
//! phrasebook fills - `EnsureInit`'s built-in rules, `AddPhraseRule`,
//! `LintRule`, `EnglishResetGrammar`, the function words and phrases of
//! `RegisterBuiltinFuncWords`, `EnglishAddFunctionWord` and
//! `RegisterFunctionWord`, the keyword aliases, and the counters
//! `EnglishVocabStats` reports (PORT.6, slice 6c).
//!
//! The VBA keeps the rule store as five parallel Collections walked by
//! position (`mPatItems`, `mPatForms`, `mPatTexts`, `mPatSigs`,
//! `mPatSources`); here they are one `Vec<Rule>`, spoken of 1-based as the
//! signature owners and the matcher speak of a rule's index, since that
//! index is dispatch order and provenance. Every table the VBA keys by a
//! folded word (`mFnOf`, `mKeywordAlias`, `mSigOwner`, ...) is a map here;
//! nothing on the translate path walks one in order.

use std::collections::HashMap;

use crate::form::Form;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::reader::read_forms;

use super::rules::{
    bare_surfaces, expanded_signatures, is_alt_cat, opt_tok, slot_tok, surface_forms,
    validate_rule_items,
};
use super::words::{
    builtin_function_words, function_word_display, is_after_value_word, is_expr_op_word,
    is_noise_word, number_word,
};

/// One phrase rule: `mPatItems`, `mPatForms`, `mPatTexts`, `mPatSigs` and
/// `mPatSources` at one index.
#[derive(Clone, Debug)]
pub struct Rule {
    /// The pattern's items, folded, noise words gone, number words as digits.
    pub items: Vec<String>,
    /// F.2: the template, pre-parsed into forms.
    pub forms: Vec<Form>,
    /// The pattern as written, for messages.
    pub text: String,
    /// G1: every shape the rule can present (`ExpandedSignatures`).
    pub sigs: Vec<String>,
    /// G3: where the rule came from, `(built-in)`, a source name, or
    /// `<new> (overrides <old>)`.
    pub source: String,
}

/// LX.14: one function phrase, as `RegisterFunctionWord` records it.
#[derive(Clone, Debug)]
pub struct Phrase {
    /// The fixed words before the hole, folded.
    pub words: Vec<String>,
    pub hole_name: String,
    pub hole_cat: String,
    /// The closing clause's items: fixed words and at most one alternation
    /// slot, as pattern tokens.
    pub clause: Vec<String>,
    /// The template, one form.
    pub forms: Vec<Form>,
    /// The shape "What can I say?" and the refusal show.
    pub shape: String,
}

/// A proof form a phrasebook carries: `test-success` or `test-fail`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProofKind {
    Success,
    Fail,
}

/// One proof, collected at load and run after every rule of the file
/// exists (`RunVocabTest`, `RunVocabFailTest`).
#[derive(Clone, Debug)]
pub struct Proof {
    pub sentence: String,
    /// The expected VLA text, or the fragment a refusal must contain.
    pub expected: String,
    pub line: u32,
    pub kind: ProofKind,
    pub row_tag: String,
    pub source: String,
}

/// How a load treats a failing proof: the reference refuses the load at
/// the first one (`Raise`, `EnglishLoadVocabularyText`'s own behaviour);
/// `frazaro prove` keeps every failure and loads on (`Collect`), to score
/// the file whole.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProofMode {
    Raise,
    Collect,
}

/// A proof that failed under [`ProofMode::Collect`], with the refusal the
/// reference would have raised for it.
#[derive(Clone, Debug)]
pub struct ProofFailure {
    pub proof: Proof,
    pub refusal: Refusal,
}

/// VLA_SentenceEngine.bas's grammar state: what `EnsureInit` makes,
/// `EnglishLoadVocabularyText` fills and `EnglishResetGrammar` empties.
pub struct Grammar {
    /// `scripts/prelude.vla`'s text, which a generator call's expansion
    /// reads beside the carried macros, as `VlaExpandText` splices it.
    pub(super) prelude: String,
    rules: Vec<Rule>,
    /// `mPreludeCount`: rules 1..=this are built in and never removed.
    prelude_count: usize,
    /// G3 perf: one exact signature to the earliest rule (1-based) that
    /// claims it.
    sig_owner: HashMap<String, usize>,
    lint_warnings: Vec<String>,
    /// G3: the file being loaded; empty for a rule added directly.
    pub(super) load_source: String,
    /// G3: an audit records what a load would refuse and walks on.
    pub(super) audit_mode: bool,
    /// L4: each carried macro's text, in load order (`mVocabMacros`).
    pub(super) vocab_macros: Vec<String>,
    /// L4: folded macro name to the file that carries it.
    pub(super) vocab_macro_names: HashMap<String, String>,
    /// LX5.1: folded surface word to its canonical English spelling.
    keyword_aliases: HashMap<String, String>,
    /// Function words read with `of`, folded word to target (`mFnOf`).
    fn_of: HashMap<String, String>,
    /// Value words (`mFnNullary`).
    fn_nullary: HashMap<String, String>,
    /// Display strings for the phrase list (`mFnDisplay`).
    fn_display: Vec<String>,
    /// LX.14: a vocabulary's own `<word> of` and the file that declared it.
    vocab_of_words: HashMap<String, String>,
    /// LX.14: the words a program's own definitions masked, with the
    /// targets to give back at the next translation (slice 6e).
    pub(super) mask_words: Vec<String>,
    pub(super) mask_targets: Vec<String>,
    pub(super) mask_notes: Vec<String>,
    phrases: Vec<Phrase>,
    /// LX.14: a phrase's first word to the phrases beginning with it,
    /// 1-based indices, longest first.
    phrase_buckets: HashMap<String, Vec<usize>>,
    /// U.10: proofs run since the last reset.
    pub(super) tests_run: u64,
    pub(super) fails_run: u64,
    last_load_source: String,
    /// GEXPANDER.0: the last load's own expanded-form text.
    pub(super) expanded_blob: String,
    /// Every proof the loads collected and ran, in order.
    pub(super) proofs: Vec<Proof>,
    pub(super) proof_mode: ProofMode,
    /// The failures a load under [`ProofMode::Collect`] kept.
    pub(super) proof_failures: Vec<ProofFailure>,
}

impl Grammar {
    /// VBA `EnsureInit`: the built-in rules, the engine's function words,
    /// and empty tables.
    pub fn new(prelude: &str) -> Grammar {
        let mut g = Grammar {
            prelude: prelude.to_string(),
            rules: Vec::new(),
            prelude_count: 0,
            sig_owner: HashMap::new(),
            lint_warnings: Vec::new(),
            load_source: String::new(),
            audit_mode: false,
            vocab_macros: Vec::new(),
            vocab_macro_names: HashMap::new(),
            keyword_aliases: HashMap::new(),
            fn_of: HashMap::new(),
            fn_nullary: HashMap::new(),
            fn_display: Vec::new(),
            vocab_of_words: HashMap::new(),
            mask_words: Vec::new(),
            mask_targets: Vec::new(),
            mask_notes: Vec::new(),
            phrases: Vec::new(),
            phrase_buckets: HashMap::new(),
            tests_run: 0,
            fails_run: 0,
            last_load_source: String::new(),
            expanded_blob: String::new(),
            proofs: Vec::new(),
            proof_mode: ProofMode::Raise,
            proof_failures: Vec::new(),
        };
        g.register_builtin_func_words();
        // The prelude vocabulary. Each line is one DCG production.
        g.load_source = "(built-in)".to_string();
        let built_in: [(&str, &str); 11] = [
            ("set {v:var} to {e:expr}", "(set! {v} {e})"),
            ("add {e:expr} to {v:var}", "(add! {v} {e})"),
            ("increase {v:var} by {e:expr}", "(add! {v} {e})"),
            ("decrease {v:var} by {e:expr}", "(set! {v} (- {v} {e}))"),
            // B7: the multiplicative verbs.
            ("grow {v:var} by {e:expr}", "(set! {v} (* {v} (+ 1 {e})))"),
            ("shrink {v:var} by {e:expr}", "(set! {v} (* {v} (- 1 {e})))"),
            // B1: lists.
            ("append {e:expr} to {v:var}", "(. {v} add {e})"),
            ("show {e:expr}", "(msgbox {e})"),
            ("say {e:expr}", "(msgbox {e})"),
            ("log {e:expr}", "(debug-print {e})"),
            ("stop", "(exit-sub)"),
        ];
        for (pattern, template) in built_in {
            g.add_phrase_rule(pattern, template, false)
                .expect("the built-in rules register");
        }
        g.load_source.clear();
        g.prelude_count = g.rules.len();
        g
    }

    /// VBA `EnglishResetGrammar`: back to the built-in rules alone, every
    /// carried macro, alias, function word and phrase gone.
    pub fn reset(&mut self) {
        self.rules.truncate(self.prelude_count);
        self.load_source.clear();
        self.vocab_macros.clear();
        self.vocab_macro_names.clear();
        self.keyword_aliases.clear();
        self.tests_run = 0;
        self.fails_run = 0;
        self.last_load_source.clear();
        self.lint_warnings.clear();
        self.register_builtin_func_words();
        self.rebuild_sig_owner();
        self.proofs.clear();
        self.proof_failures.clear();
        self.expanded_blob.clear();
    }

    pub fn rules(&self) -> &[Rule] {
        &self.rules
    }

    /// VBA `EnglishRuleCount`: every rule, the built-in ones included.
    pub fn rule_count(&self) -> usize {
        self.rules.len()
    }

    pub fn prelude_count(&self) -> usize {
        self.prelude_count
    }

    /// VBA `RuleSourceOf`: a rule's provenance, by its 1-based index.
    pub fn rule_source(&self, idx: usize) -> &str {
        &self.rules[idx - 1].source
    }

    /// `mVocabMacros`: every carried macro's text, one per line (CRLF),
    /// as the translation appends it and a generator's expansion reads it.
    pub fn vocab_macros_text(&self) -> String {
        self.vocab_macros.join("\r\n")
    }

    pub fn vocab_macro_count(&self) -> usize {
        self.vocab_macro_names.len()
    }

    /// LX5.1: the canonical spelling a phrasebook gave a word, if any.
    pub fn keyword_alias(&self, word: &str) -> Option<&str> {
        self.keyword_aliases.get(word).map(String::as_str)
    }

    pub fn keyword_alias_count(&self) -> usize {
        self.keyword_aliases.len()
    }

    /// `IsFnWord(mFnOf, word)` and `FnTarget`.
    pub fn fn_of_target(&self, word: &str) -> Option<&str> {
        self.fn_of.get(&fold(word)).map(String::as_str)
    }

    pub fn fn_nullary_target(&self, word: &str) -> Option<&str> {
        self.fn_nullary.get(&fold(word)).map(String::as_str)
    }

    pub fn fn_of_count(&self) -> usize {
        self.fn_of.len()
    }

    pub fn fn_nullary_count(&self) -> usize {
        self.fn_nullary.len()
    }

    pub fn fn_display(&self) -> &[String] {
        &self.fn_display
    }

    pub fn phrases(&self) -> &[Phrase] {
        &self.phrases
    }

    /// LX.14: the phrases beginning with a word, longest first, 1-based.
    pub fn phrase_bucket(&self, word: &str) -> Option<&[usize]> {
        self.phrase_buckets.get(word).map(Vec::as_slice)
    }

    pub fn lint_warnings(&self) -> &[String] {
        &self.lint_warnings
    }

    /// VBA `EnglishLintReport`.
    pub fn lint_report(&self) -> String {
        if self.lint_warnings.is_empty() {
            return "No grammar warnings.".to_string();
        }
        self.lint_warnings
            .iter()
            .map(|w| format!("WARNING: {w}\r\n"))
            .collect()
    }

    /// VBA `EnglishVocabStats`: the load-report counters line.
    pub fn vocab_stats(&self) -> String {
        let nr = self.rules.len() - self.prelude_count;
        let nm = self.vocab_macro_names.len();
        let plural = |n: u64| if n == 1 { "" } else { "s" };
        let mut r = format!(
            "loaded: {nr} rule{}, {nm} macro{}, {} test{} ({} expected fail{})",
            plural(nr as u64),
            plural(nm as u64),
            self.tests_run,
            plural(self.tests_run),
            self.fails_run,
            plural(self.fails_run)
        );
        if !self.last_load_source.is_empty() {
            r.push_str(&format!(" from {}", self.last_load_source));
        }
        r
    }

    pub fn proofs(&self) -> &[Proof] {
        &self.proofs
    }

    /// What a failing proof does to a load; `Raise` unless told otherwise.
    pub fn set_proof_mode(&mut self, mode: ProofMode) {
        self.proof_mode = mode;
    }

    pub fn proof_mode(&self) -> ProofMode {
        self.proof_mode
    }

    /// The proofs that failed under [`ProofMode::Collect`], in file order.
    pub fn proof_failures(&self) -> &[ProofFailure] {
        &self.proof_failures
    }

    /// `mTestsRun` and `mFailsRun`: the positive and negative proofs run
    /// since the last reset.
    pub fn proofs_run(&self) -> (u64, u64) {
        (self.tests_run, self.fails_run)
    }

    /// The last load's expanded-form text (`mLastExpandedBlob`).
    pub fn expanded_blob(&self) -> &str {
        &self.expanded_blob
    }

    pub(super) fn set_last_load_source(&mut self, source: &str) {
        self.last_load_source = source.to_string();
    }

    /// VBA `EnglishAddPhrase`: one production added directly.
    pub fn add_phrase(&mut self, pattern: &str, template: &str) -> Result<(), Refusal> {
        self.add_phrase_rule(pattern, template, false)
    }

    /// VBA `AddPhraseRule`.
    pub(super) fn add_phrase_rule(
        &mut self,
        pattern: &str,
        template: &str,
        is_override: bool,
    ) -> Result<(), Refusal> {
        let mut items: Vec<String> = Vec::new();
        for part in pattern.split(' ') {
            let w = fold(part.trim_matches(' '));
            if !w.is_empty() && !is_noise_word(&w) {
                // G-TEXT slice 2: a number word is read as the tokenizer
                // reads it in a sentence.
                items.push(number_word(&w));
            }
        }
        if items.is_empty() {
            return Err(raise("english-empty-pattern", &[]));
        }
        validate_rule_items(pattern, &items)?;

        let src = if self.load_source.is_empty() {
            "(added directly)".to_string()
        } else {
            self.load_source.clone()
        };

        if is_override {
            // G3: an override replaces exactly one earlier same-shape rule,
            // in place; matching nothing, several, or the built-in core
            // refuses. Deduped by owner index, not by signature.
            let new_sigs = expanded_signatures(&items)?;
            let mut hits: Vec<usize> = Vec::new();
            for e in &new_sigs {
                if let Some(&k) = self.sig_owner.get(e) {
                    if !hits.contains(&k) {
                        hits.push(k);
                    }
                }
            }
            if hits.is_empty() {
                return Err(raise("english-override-no-match", &[("pattern", pattern)]));
            }
            if hits.len() > 1 {
                let list = hits
                    .iter()
                    .map(|&k| format!("'{}'", self.rules[k - 1].text))
                    .collect::<Vec<_>>()
                    .join("; ");
                return Err(raise(
                    "english-override-ambiguous",
                    &[
                        ("pattern", pattern),
                        ("count", &hits.len().to_string()),
                        ("list", &list),
                    ],
                ));
            }
            let tgt = hits[0];
            if tgt <= self.prelude_count {
                return Err(raise(
                    "english-override-matches-builtin",
                    &[("pattern", pattern), ("builtin", &self.rules[tgt - 1].text)],
                ));
            }
            self.lint_rule(pattern, &items, tgt)?;
            let forms = template_forms(template)?;
            let old_src = self.rules[tgt - 1].source.clone();
            let old_sigs = self.rules[tgt - 1].sigs.clone();
            self.rules[tgt - 1] = Rule {
                items,
                forms,
                text: pattern.to_string(),
                sigs: new_sigs.clone(),
                source: format!("{src} (overrides {old_src})"),
            };
            // G3 perf: only an entry the target itself still owns is
            // cleared; a shape it merely shared is not its to remove.
            for osl in &old_sigs {
                if self.sig_owner.get(osl) == Some(&tgt) {
                    self.sig_owner.remove(osl);
                }
            }
            for nsv in &new_sigs {
                self.add_sig_owner(nsv, tgt);
            }
            return Ok(());
        }

        let dup_msg = self.lint_rule(pattern, &items, 0)?;
        if !dup_msg.is_empty() && !self.audit_mode {
            // G3: a silent same-shape shadow refuses the load unless the
            // replacement is earned with the override marker.
            return Err(raise("english-rule-shadow", &[("dupMsg", &dup_msg)]));
        }
        let append_sigs = expanded_signatures(&items)?;
        let forms = template_forms(template)?;
        self.rules.push(Rule {
            items,
            forms,
            text: pattern.to_string(),
            sigs: append_sigs.clone(),
            source: src,
        });
        let idx = self.rules.len();
        for asv in &append_sigs {
            self.add_sig_owner(asv, idx);
        }
        Ok(())
    }

    /// VBA `AddSigOwner`: idempotent; the first rule to claim a shape
    /// keeps owning it, as dispatch's first match wins.
    fn add_sig_owner(&mut self, sig: &str, idx: usize) {
        self.sig_owner.entry(sig.to_string()).or_insert(idx);
    }

    /// VBA `RebuildSigOwner`.
    fn rebuild_sig_owner(&mut self) {
        self.sig_owner.clear();
        let all: Vec<Vec<String>> = self.rules.iter().map(|r| r.sigs.clone()).collect();
        for (i, sigs) in all.iter().enumerate() {
            for sig in sigs {
                self.add_sig_owner(sig, i + 1);
            }
        }
    }

    /// VBA `LintNoiseWordBeforeSlot` (F.4, shape 2): a noise word right
    /// before a slot in the raw pattern is stripped at registration, so the
    /// slot would swallow whatever the input leaves in its place.
    fn lint_noise_word_before_slot(&mut self, pattern: &str) {
        let parts: Vec<&str> = pattern.split(' ').collect();
        for i in 0..parts.len().saturating_sub(1) {
            let w = fold(parts[i].trim_matches(' '));
            if is_noise_word(&w) {
                let nxt = parts[i + 1].trim_matches(' ');
                if nxt.starts_with('{') {
                    self.lint_warnings.push(format!(
                        "pattern '{pattern}': noise word '{w}' immediately precedes slot {nxt} - '{w}' is stripped from the pattern at registration, so the slot will silently swallow whatever token the input leaves in its place"
                    ));
                }
            }
        }
    }

    /// VBA `LintRule`: the unreachable-literal warnings and the duplicate
    /// check over expanded signatures. Returns the first duplicate's
    /// warning, or nothing.
    fn lint_rule(
        &mut self,
        pattern: &str,
        items: &[String],
        skip_idx: usize,
    ) -> Result<String, Refusal> {
        self.lint_noise_word_before_slot(pattern);
        let mut prev_cat = String::new();
        let after_value = |prev: &str| prev == "expr" || prev == "cond";
        for t in items {
            if let Some(slot) = slot_tok(t)? {
                if is_alt_cat(&slot.cat) {
                    // G1: an alternation consumes a literal token, so a
                    // branch that is an operator word after an expression
                    // slot is unreachable.
                    if after_value(&prev_cat) {
                        for branch in slot.cat.split('|') {
                            for sfv in surface_forms(&fold(branch)) {
                                if is_expr_op_word(&sfv)
                                    || (prev_cat == "cond" && (sfv == "and" || sfv == "or"))
                                {
                                    self.lint_warnings.push(format!("pattern '{pattern}': alternative '{sfv}' directly follows a {{:{prev_cat}}} slot - the expression will consume it and that branch can never match"));
                                }
                            }
                        }
                    }
                    prev_cat.clear();
                } else {
                    prev_cat = slot.cat.clone();
                }
            } else if t.contains('|') || t.contains('/') {
                // G10: a bare surface token consumes a literal.
                if after_value(&prev_cat) {
                    for sfv in bare_surfaces(t) {
                        if is_expr_op_word(&sfv)
                            || (prev_cat == "cond" && (sfv == "and" || sfv == "or"))
                        {
                            self.lint_warnings.push(format!("pattern '{pattern}': alternative '{sfv}' directly follows a {{:{prev_cat}}} slot - the expression will consume it and that branch can never match"));
                        }
                    }
                }
                prev_cat.clear();
            } else if let Some(ow) = opt_tok(t)? {
                // G1: an optional may be absent, so the expression stays
                // adjacent to whatever follows: prevCat persists.
                if after_value(&prev_cat) && is_expr_op_word(&ow) {
                    self.lint_warnings.push(format!("pattern '{pattern}': optional literal '{ow}' directly follows a {{:{prev_cat}}} slot - the expression will consume it whenever it is present"));
                } else if prev_cat == "cond" && (ow == "and" || ow == "or") {
                    self.lint_warnings.push(format!("pattern '{pattern}': optional literal '{ow}' directly follows a {{:cond}} slot - the condition will consume it whenever it is present"));
                }
            } else {
                if after_value(&prev_cat) && is_expr_op_word(t) {
                    self.lint_warnings.push(format!("pattern '{pattern}': literal '{t}' directly follows a {{:{prev_cat}}} slot - the expression will consume it and the rule can never match"));
                } else if prev_cat == "cond" && (t == "and" || t == "or") {
                    self.lint_warnings.push(format!("pattern '{pattern}': literal '{t}' directly follows a {{:cond}} slot - the condition will consume it and the rule can never match"));
                }
                prev_cat.clear();
            }
        }

        // The duplicate check, over expanded signatures, against the
        // earliest owner of each shape; an override skips its own target.
        let mut first = String::new();
        for e in expanded_signatures(items)? {
            if let Some(&owner) = self.sig_owner.get(&e) {
                if owner != skip_idx {
                    let w = format!(
                        "pattern '{pattern}' duplicates '{}' (same shape: {e}) - the earlier rule always wins",
                        self.rules[owner - 1].text
                    );
                    self.lint_warnings.push(w.clone());
                    if first.is_empty() {
                        first = w;
                    }
                }
            }
        }
        Ok(first)
    }

    /// VBA `RegisterBuiltinFuncWords`: the engine's own function words,
    /// from `scripts/names.vla`, over fresh tables with no phrase and no mask.
    fn register_builtin_func_words(&mut self) {
        self.fn_of.clear();
        self.fn_nullary.clear();
        self.fn_display.clear();
        self.vocab_of_words.clear();
        self.mask_words.clear();
        self.mask_targets.clear();
        self.mask_notes.clear();
        self.phrases.clear();
        self.phrase_buckets.clear();
        for fw in builtin_function_words() {
            let table = if fw.takes_of {
                &mut self.fn_of
            } else {
                &mut self.fn_nullary
            };
            // AddFnEntry: remove then add, so a redefinition wins.
            table.insert(fold(&fw.word), fw.target.clone());
        }
        self.fn_display
            .extend(function_word_display().iter().cloned());
    }

    /// VBA `EnglishAddFunctionWord`: a value word, or a word read with
    /// `of`, whose provenance is kept for LX.14's masking note.
    pub fn add_function_word(&mut self, word: &str, target: &str, nullary: bool) {
        let key = fold(word);
        if nullary {
            self.fn_nullary.insert(key.clone(), target.to_string());
            self.fn_display.push(key);
        } else {
            self.fn_of.insert(key.clone(), target.to_string());
            self.fn_display.push(format!("{key} of ..."));
            let src = if self.load_source.is_empty() {
                "(added directly)".to_string()
            } else {
                self.load_source.clone()
            };
            self.vocab_of_words.insert(key, src);
        }
    }

    /// LX.14: may a program's own definition take this word? Only a
    /// vocabulary's `<word> of`, never one the engine seeds.
    pub fn is_maskable_fn_word(&self, word: &str) -> bool {
        self.vocab_of_words.contains_key(&fold(word))
    }

    /// LX.14: the file that declared a vocabulary's own `<word> of`, for the
    /// note a program's masking definition leaves.
    pub fn vocab_of_source(&self, word: &str) -> Option<&str> {
        self.vocab_of_words.get(&fold(word)).map(String::as_str)
    }

    /// VBA `RegisterKeywordAlias` (LX5.1): last write wins.
    pub(super) fn register_keyword_alias(&mut self, surface: &str, canonical: &str) {
        self.keyword_aliases.insert(fold(surface), fold(canonical));
    }

    /// VBA `RegisterFunctionWord` (LX.14): a `-function` directive's
    /// pattern and target, as a value word, a one-word head, or a phrase,
    /// each audited as it registers.
    pub(super) fn register_function_word(
        &mut self,
        lhs: &str,
        target: &str,
        target_is_form: bool,
        context: &str,
    ) -> Result<(), Refusal> {
        let mut items: Vec<String> = Vec::new();
        for part in lhs.trim_matches(' ').split(' ') {
            let w = fold(part.trim_matches(' '));
            if !w.is_empty() && !is_noise_word(&w) {
                items.push(number_word(&w));
            }
        }
        if items.is_empty() {
            return Err(phrase_shape(context, lhs, "has no words"));
        }

        // The hole, and the one alternation after it.
        let mut hole_at = 0usize;
        let mut alt_at = 0usize;
        for (i, w) in items.iter().enumerate() {
            let i = i + 1;
            if let Some(slot) = slot_tok(w)? {
                if slot.has_default {
                    return Err(phrase_shape(
                        context,
                        lhs,
                        "gives a slot a default, and a phrase's value is always said",
                    ));
                }
                if is_alt_cat(&slot.cat) {
                    if hole_at == 0 {
                        return Err(phrase_shape(
                            context,
                            lhs,
                            "has a choice of words before its value",
                        ));
                    }
                    if alt_at > 0 {
                        return Err(phrase_shape(context, lhs, "has a second choice of words"));
                    }
                    alt_at = i;
                } else {
                    if hole_at > 0 {
                        return Err(phrase_shape(
                            context,
                            lhs,
                            "takes a second value, which makes it a sentence rule, not a phrase",
                        ));
                    }
                    match slot.cat.as_str() {
                        "value" | "range" | "column" | "cell" => {}
                        other => {
                            return Err(phrase_shape(
                                context,
                                lhs,
                                &format!("takes a '{other}', where a phrase takes a value, a range, a column or a cell"),
                            ))
                        }
                    }
                    hole_at = i;
                }
            } else if w.starts_with('[') || w.contains('|') || w.contains('/') {
                return Err(phrase_shape(
                    context,
                    lhs,
                    "has an optional word or a choice of words outside {braces}, where a phrase's words are fixed",
                ));
            }
        }

        // No hole: one word is a value word; two or more take a value after them.
        if hole_at == 0 {
            if items.len() == 1 {
                if target_is_form {
                    return Err(phrase_shape(
                        context,
                        lhs,
                        "is a value word, whose target is one name",
                    ));
                }
                self.add_function_word(&items[0], target, true);
                return Ok(());
            }
            items.push("{x:value}".to_string());
            hole_at = items.len();
        }
        if hole_at == 1 {
            return Err(phrase_shape(
                context,
                lhs,
                "begins with its value, where a phrase begins with a word",
            ));
        }
        let hole = slot_tok(&items[hole_at - 1])?.expect("the hole is a slot");
        let (hole_name, hole_cat) = (hole.name, hole.cat);

        // A word and "of", then a value, called by name: a one-word head.
        if hole_at == 3
            && items.len() == 3
            && items[1] == "of"
            && hole_cat == "value"
            && !target_is_form
        {
            self.refuse_same_phrase_words(&format!("{} of", items[0]), context, lhs)?;
            let word = items[0].clone();
            self.add_function_word(&word, target, false);
            return Ok(());
        }

        // A phrase: its words before the hole, then its closing clause.
        let mut words: Vec<String> = Vec::new();
        let mut joined = String::new();
        for (i, w) in items[..hole_at - 1].iter().enumerate() {
            self.audit_phrase_word(w, i > 0, context, lhs)?;
            words.push(w.clone());
            if !joined.is_empty() {
                joined.push(' ');
            }
            joined.push_str(w);
        }
        let mut clause: Vec<String> = Vec::new();
        let mut alt_name = String::new();
        for (i, w) in items.iter().enumerate().skip(hole_at) {
            let i = i + 1;
            if i == alt_at {
                let alt = slot_tok(w)?.expect("the alternation is a slot");
                alt_name = alt.name;
                for branch in alt.cat.split('|') {
                    self.audit_phrase_word(&stem_of(&fold(branch)), true, context, lhs)?;
                }
            } else {
                self.audit_phrase_word(w, true, context, lhs)?;
            }
            clause.push(w.clone());
        }
        self.refuse_same_phrase_words(&joined, context, lhs)?;
        if words.len() == 2 && words[1] == "of" && self.fn_of.contains_key(&words[0]) {
            return Err(raise(
                "english-function-phrase-duplicate",
                &[
                    ("loc", context),
                    ("pattern", lhs),
                    ("other", &format!("{} of ...", words[0])),
                ],
            ));
        }

        // Its template: one form, using the value and the word chosen.
        let template = if target_is_form {
            target.to_string()
        } else {
            format!("({target} {{{hole_name}}})")
        };
        if !template.contains(&format!("{{{hole_name}}}")) {
            return Err(phrase_template(
                context,
                lhs,
                &format!("never uses {{{hole_name}}}, so the value would be lost"),
            ));
        }
        if alt_at > 0 && !template.contains(&format!("{{{alt_name}}}")) {
            return Err(phrase_template(
                context,
                lhs,
                &format!("never uses {{{alt_name}}}, so the word chosen would be lost"),
            ));
        }
        let forms = template_forms(&template)?;
        if forms.len() != 1 {
            return Err(phrase_template(
                context,
                lhs,
                &format!("is {} forms, where a value is one", forms.len()),
            ));
        }

        // Its shape, as "What can I say?" and its refusal show it.
        let mut shape = format!("{joined} {}", phrase_hole_shape(&hole_cat));
        for cv in &clause {
            match slot_tok(cv)? {
                Some(slot) => shape.push_str(&format!(" {}", slot.cat)),
                None => shape.push_str(&format!(" {cv}")),
            }
        }

        self.phrases.push(Phrase {
            words,
            hole_name,
            hole_cat,
            clause,
            forms,
            shape: shape.clone(),
        });
        self.add_phrase_to_bucket(self.phrases.len());
        self.fn_display.push(shape);
        Ok(())
    }

    /// VBA `AuditPhraseWord`: no word after the first that the value or
    /// condition grammar reads right after a value; no keyword alias's word.
    fn audit_phrase_word(
        &self,
        w: &str,
        after_first: bool,
        context: &str,
        lhs: &str,
    ) -> Result<(), Refusal> {
        if after_first && is_after_value_word(w) {
            return Err(raise(
                "english-function-phrase-after-value-word",
                &[("loc", context), ("pattern", lhs), ("word", w)],
            ));
        }
        if self.keyword_aliases.contains_key(w) {
            return Err(raise(
                "english-function-phrase-alias-word",
                &[("loc", context), ("pattern", lhs), ("word", w)],
            ));
        }
        Ok(())
    }

    /// VBA `RefuseSamePhraseWords`: no two phrases with the same words
    /// before the hole.
    fn refuse_same_phrase_words(
        &self,
        joined: &str,
        context: &str,
        lhs: &str,
    ) -> Result<(), Refusal> {
        for p in &self.phrases {
            if p.words.join(" ") == joined {
                return Err(raise(
                    "english-function-phrase-duplicate",
                    &[("loc", context), ("pattern", lhs), ("other", &p.shape)],
                ));
            }
        }
        Ok(())
    }

    /// VBA `RefusePhraseWordAlias`: a keyword alias rewrites its word in
    /// every sentence before any phrase is read, so a phrase using the
    /// word could never be said; refused whichever is declared first.
    pub(super) fn refuse_phrase_word_alias(
        &self,
        surface: &str,
        context: &str,
    ) -> Result<(), Refusal> {
        let k = fold(surface);
        for p in &self.phrases {
            let mut hit = p.words.contains(&k);
            if !hit {
                for wv in &p.clause {
                    match slot_tok(wv)? {
                        Some(slot) => {
                            if slot.cat.split('|').any(|b| stem_of(&fold(b)) == k) {
                                hit = true;
                            }
                        }
                        None => {
                            if *wv == k {
                                hit = true;
                            }
                        }
                    }
                }
            }
            if hit {
                return Err(raise(
                    "english-function-phrase-alias-word",
                    &[("loc", context), ("pattern", &p.shape), ("word", &k)],
                ));
            }
        }
        Ok(())
    }

    /// VBA `AddPhraseToBucket`: a phrase joins its first word's bucket,
    /// longest first; a tie keeps registration order.
    fn add_phrase_to_bucket(&mut self, idx: usize) {
        let first = self.phrases[idx - 1].words[0].clone();
        let n = self.phrases[idx - 1].words.len();
        let lengths: Vec<usize> = self.phrases.iter().map(|p| p.words.len()).collect();
        let bucket = self.phrase_buckets.entry(first).or_default();
        let at = bucket
            .iter()
            .position(|&other| lengths[other - 1] < n)
            .unwrap_or(bucket.len());
        bucket.insert(at, idx);
    }
}

/// VBA `TemplateForms` (F.2): a template pre-parsed into forms, refusing
/// one that does not read as at least one form.
pub(super) fn template_forms(template: &str) -> Result<Vec<Form>, Refusal> {
    let detail = match read_forms(template) {
        Ok(forms) if !forms.is_empty() => return Ok(forms),
        Ok(_) => String::new(),
        Err(e) => format!(" ({})", e.text),
    };
    Err(raise(
        "english-template-not-well-formed",
        &[("template", template), ("detail", &detail)],
    ))
}

/// VBA `StemOf`: a branch's stem, `center/ed` is `center`.
pub(super) fn stem_of(spec: &str) -> String {
    match spec.find('/') {
        Some(sp) => spec[..sp].to_string(),
        None => spec.to_string(),
    }
}

/// VBA `PhraseHoleShape`: what a phrase's hole shows in its shape.
fn phrase_hole_shape(cat: &str) -> String {
    if cat == "value" {
        "...".to_string()
    } else {
        format!("<{cat}>")
    }
}

/// VBA `RefusePhraseShape`.
fn phrase_shape(context: &str, lhs: &str, why: &str) -> Refusal {
    raise(
        "english-function-phrase-shape",
        &[
            ("loc", context),
            ("pattern", lhs),
            ("why", why),
            (
                "holes",
                "{x:value}, or {r:range}, {c:column} or {c:cell} for a reference",
            ),
            ("choice", "{k:sample|population}"),
        ],
    )
}

/// VBA `RefusePhraseTemplate`.
fn phrase_template(context: &str, lhs: &str, why: &str) -> Refusal {
    raise(
        "english-function-phrase-template",
        &[("loc", context), ("pattern", lhs), ("why", why)],
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");

    fn fresh() -> Grammar {
        Grammar::new(PRELUDE)
    }

    fn not_internal(r: &Refusal) {
        assert!(!r.text.starts_with("RaiseMsg:"), "{}", r.text);
    }

    #[test]
    fn the_built_in_rules_are_ensureinit_s_eleven() {
        let g = fresh();
        assert_eq!(g.rule_count(), 11);
        assert_eq!(g.prelude_count(), 11);
        assert_eq!(g.rules()[0].text, "set {v:var} to {e:expr}");
        assert_eq!(g.rules()[0].items, vec!["set", "{v:var}", "to", "{e:expr}"]);
        assert_eq!(g.rules()[0].sigs, vec!["set|{var}|to|{expr}"]);
        assert_eq!(g.rules()[0].source, "(built-in)");
        assert_eq!(g.rules()[10].text, "stop");
        assert_eq!(g.fn_of_count(), 14);
        assert_eq!(g.fn_nullary_count(), 2);
        assert_eq!(g.fn_of_target("length"), Some("len"));
        assert_eq!(g.fn_nullary_target("today"), Some("date"));
        assert_eq!(g.fn_display().len(), 17);
        assert_eq!(
            g.vocab_stats(),
            "loaded: 0 rules, 0 macros, 0 tests (0 expected fails)"
        );
        assert_eq!(g.lint_report(), "No grammar warnings.");
    }

    #[test]
    fn a_rule_registers_with_its_items_forms_and_shapes() {
        let mut g = fresh();
        g.add_phrase("warm the cell {r:cell} [up]", "(debug-print {r})")
            .unwrap();
        let r = &g.rules()[11];
        assert_eq!(r.items, vec!["warm", "cell", "{r:cell}", "[up]"]);
        assert_eq!(r.sigs, vec!["warm|cell|{cell}", "warm|cell|{cell}|up"]);
        assert_eq!(r.forms.len(), 1);
        assert_eq!(r.source, "(added directly)");
        assert_eq!(g.rule_count(), 12);
        // A number word in a pattern reads as the tokenizer writes it.
        g.add_phrase("count as one list", "(x)").unwrap();
        assert_eq!(g.rules()[12].items, vec!["count", "as", "1", "list"]);
        // An empty pattern, and a template that is not a form.
        assert_eq!(
            g.add_phrase("the a", "(x)").unwrap_err().id,
            "english-empty-pattern"
        );
        let e = g.add_phrase("zap", "").unwrap_err();
        assert_eq!(e.id, "english-template-not-well-formed");
        assert_eq!(e.text, "template does not parse as well-formed VLA: ");
        let e = g.add_phrase("zap", "(x").unwrap_err();
        assert!(e.text.contains("(unbalanced parentheses"), "{}", e.text);
    }

    #[test]
    fn a_same_shape_rule_is_refused_and_an_override_replaces_in_place() {
        let mut g = fresh();
        g.load_source = "base-vocab".to_string();
        g.add_phrase_rule("warm cell {r:cell}", "(debug-print 1)", false)
            .unwrap();
        g.load_source = "dialect-vocab".to_string();
        let e = g
            .add_phrase_rule("warm cell {r:cell}", "(debug-print 2)", false)
            .unwrap_err();
        assert_eq!(e.id, "english-rule-shadow");
        not_internal(&e);
        assert!(
            e.text.contains("duplicates") && e.text.contains("-vla-override"),
            "{}",
            e.text
        );
        assert!(e.text.starts_with("pattern 'warm cell {r:cell}' duplicates 'warm cell {r:cell}' (same shape: warm|cell|{cell}) - the earlier rule always wins"));
        // The override wins the index and the provenance names both files.
        g.add_phrase_rule("warm cell {r:cell}", "(debug-print 2)", true)
            .unwrap();
        assert_eq!(g.rule_count(), 12);
        assert_eq!(g.rules()[11].source, "dialect-vocab (overrides base-vocab)");
        assert_eq!(
            g.rules()[11].forms[0],
            Form::list(vec![Form::sym("debug-print"), Form::sym("2")])
        );
        // Matching nothing, several, or the built-in core.
        let e = g
            .add_phrase_rule("cool cell {r:cell}", "(debug-print 2)", true)
            .unwrap_err();
        assert_eq!(e.id, "english-override-no-match");
        g.add_phrase_rule("fix cell {r:range}", "(debug-print 1)", false)
            .unwrap();
        g.add_phrase_rule("fix range {r:range}", "(debug-print 2)", false)
            .unwrap();
        let e = g
            .add_phrase_rule("fix {w:cell|range} {r:range}", "(debug-print 3)", true)
            .unwrap_err();
        assert_eq!(e.id, "english-override-ambiguous");
        assert!(
            e.text
                .contains("matches 2 earlier rules ('fix cell {r:range}'; 'fix range {r:range}')"),
            "{}",
            e.text
        );
        let e = g
            .add_phrase_rule("set {v:var} to {e:expr}", "(debug-print 9)", true)
            .unwrap_err();
        assert_eq!(e.id, "english-override-matches-builtin");
        assert!(e.text.contains("not overridable"));
        // A shape shared across spellings is a duplicate too; a different
        // category is a different shape.
        let e = g
            .add_phrase_rule("fix cell {r:range} [up]", "(x)", false)
            .unwrap_err();
        assert_eq!(e.id, "english-rule-shadow");
        assert!(
            e.text
                .contains("'fix cell {r:range} [up]' duplicates 'fix cell {r:range}'"),
            "{}",
            e.text
        );
        assert!(g
            .add_phrase_rule("fix cell {r:text} [up]", "(x)", false)
            .is_ok());
        // A reset leaves the eleven and their owners.
        g.reset();
        assert_eq!(g.rule_count(), 11);
        assert!(g
            .add_phrase("warm cell {r:cell}", "(debug-print 1)")
            .is_ok());
        assert_eq!(g.rules()[11].source, "(added directly)");
    }

    #[test]
    fn the_lint_names_unreachable_literals_after_a_value() {
        let mut g = fresh();
        g.add_phrase("zap {e:expr} plus one", "(x)").unwrap();
        g.add_phrase("zip {c:cond} and more", "(x)").unwrap();
        g.add_phrase("zop {e:expr} [minus] it", "(x)").unwrap();
        g.add_phrase("zup {e:expr} {d:times|by}", "(x)").unwrap();
        g.add_phrase("send an email a {who:expr}", "(x)").unwrap();
        let w = g.lint_warnings();
        assert!(w.iter().any(|x| x == "pattern 'zap {e:expr} plus one': literal 'plus' directly follows a {:expr} slot - the expression will consume it and the rule can never match"), "{w:?}");
        assert!(w.iter().any(|x| x == "pattern 'zip {c:cond} and more': literal 'and' directly follows a {:cond} slot - the condition will consume it and the rule can never match"));
        assert!(w.iter().any(|x| x == "pattern 'zop {e:expr} [minus] it': optional literal 'minus' directly follows a {:expr} slot - the expression will consume it whenever it is present"));
        assert!(w.iter().any(|x| x == "pattern 'zup {e:expr} {d:times|by}': alternative 'times' directly follows a {:expr} slot - the expression will consume it and that branch can never match"));
        assert!(w.iter().any(|x| x.starts_with("pattern 'send an email a {who:expr}': noise word 'a' immediately precedes slot {who:expr}")));
        assert!(g.lint_report().starts_with("WARNING: pattern 'zap"));
    }

    #[test]
    fn function_words_and_phrases_register_as_lx14_does() {
        let mut g = fresh();
        g.load_source = "english.vla".to_string();
        g.register_function_word("sum of", "sum-of", false, "english.vla line 1")
            .unwrap();
        assert_eq!(g.fn_of_target("sum"), Some("sum-of"));
        assert!(g.is_maskable_fn_word("sum") && !g.is_maskable_fn_word("length"));
        g.register_function_word("today-ish", "date", false, "ctx")
            .unwrap();
        assert_eq!(g.fn_nullary_target("today-ish"), Some("date"));
        g.register_function_word(
            "median of range {r:range}",
            "(median-of (range {r}))",
            true,
            "ctx",
        )
        .unwrap();
        g.register_function_word("median of", "median-of", false, "ctx")
            .unwrap();
        g.register_function_word(
            "standard deviation of {x:value} as {k:sample|population}",
            "({k}-standard-deviation-of {x})",
            true,
            "ctx",
        )
        .unwrap();
        g.register_function_word("zeta de", "zeta-de", false, "ctx")
            .unwrap();
        assert_eq!(g.phrases().len(), 3);
        assert_eq!(g.phrases()[0].shape, "median of range <range>");
        assert_eq!(
            g.phrases()[1].shape,
            "standard deviation of ... as sample|population"
        );
        assert_eq!(g.phrases()[1].clause, vec!["as", "{k:sample|population}"]);
        assert_eq!(g.phrases()[2].shape, "zeta de ...");
        assert_eq!(g.phrases()[2].hole_name, "x");
        assert_eq!(g.phrase_bucket("median"), Some(&[1usize][..]));
        assert_eq!(g.fn_display().last().unwrap(), "zeta de ...");
        // Longest first in a bucket, ties in registration order; a phrase
        // whose two words are a one-word head's is refused below.
        g.register_function_word("median of column {c:column}", "(mc {c})", true, "ctx")
            .unwrap();
        g.register_function_word(
            "median of column total {c:column}",
            "(mct {c})",
            true,
            "ctx",
        )
        .unwrap();
        assert_eq!(g.phrase_bucket("median"), Some(&[5usize, 1, 4][..]));
        let e = g
            .register_function_word("median of {x:value} as {k:a|b}", "({k} {x})", true, "ctx")
            .unwrap_err();
        assert_eq!(e.id, "english-function-phrase-duplicate");
        assert!(e.text.contains("as 'median of ...'"), "{}", e.text);

        let refused = |g: &mut Grammar, lhs: &str, target: &str, form: bool| -> Refusal {
            let e = g
                .register_function_word(lhs, target, form, "ctx")
                .unwrap_err();
            not_internal(&e);
            e
        };
        let e = refused(&mut g, "total plus {x:value}", "total-plus", false);
        assert_eq!(e.id, "english-function-phrase-after-value-word");
        assert!(e.text.contains("uses 'plus' after its first word"));
        let e = refused(&mut g, "zeta of {x:value} is here", "(zeta {x})", true);
        assert!(e.text.contains("uses 'is' after its first word"));
        let e = refused(&mut g, "median of range {c:column}", "(zeta {c})", true);
        assert_eq!(e.id, "english-function-phrase-duplicate");
        assert!(e
            .text
            .contains("has the same words before its value as 'median of range <range>'"));
        let e = refused(&mut g, "length of {r:range}", "(len (range {r}))", true);
        assert!(e.text.contains("'length of ...'"));
        let e = refused(
            &mut g,
            "ratio of {a:value} to {b:value}",
            "(ratio {a} {b})",
            true,
        );
        assert_eq!(e.id, "english-function-phrase-shape");
        assert!(e.text.contains("takes a second value"));
        let e = refused(&mut g, "zeta of {x:text}", "(zeta {x})", true);
        assert!(e.text.contains("takes a 'text'"));
        let e = refused(&mut g, "zeta of {x:value} as {k:one|two}", "zeta-of", false);
        assert_eq!(e.id, "english-function-phrase-template");
        assert!(e.text.contains("never uses {k}"));
        let e = refused(&mut g, "zeta", "(zeta 1)", true);
        assert!(e.text.contains("is a value word"));
        let e = refused(&mut g, "{x:value} squared", "(sq {x})", true);
        assert!(e.text.contains("begins with its value"));
        let e = refused(&mut g, "zeta [of] {x:value}", "(z {x})", true);
        assert!(e.text.contains("outside {braces}"));
        let e = refused(&mut g, "zeta {k:a|b} of {x:value}", "({k} {x})", true);
        assert!(e.text.contains("has a choice of words before its value"));
        let e = refused(&mut g, "zeta of {x:value=5}", "(z {x})", true);
        assert!(e.text.contains("gives a slot a default"));
        let e = refused(&mut g, "zeta of {x:value}", "(z {x}) (z 2)", true);
        assert!(e.text.contains("is 2 forms, where a value is one"));
        // An alias's word, whichever is declared first.
        g.register_keyword_alias("ipsum", "if");
        let e = refused(&mut g, "lorem ipsum of {x:value}", "(lorem {x})", true);
        assert_eq!(e.id, "english-function-phrase-alias-word");
        assert!(e
            .text
            .contains("'ipsum' is both a keyword alias and a word of the phrase"));
        let e = g.refuse_phrase_word_alias("median", "ctx2").unwrap_err();
        assert_eq!(e.id, "english-function-phrase-alias-word");
        assert!(e.text.contains("'median' is both a keyword alias and a word of the phrase 'median of range <range>'"), "{}", e.text);
        let e = g
            .refuse_phrase_word_alias("population", "ctx2")
            .unwrap_err();
        assert!(e
            .text
            .contains("'standard deviation of ... as sample|population'"));
        assert!(g.refuse_phrase_word_alias("nothing", "ctx2").is_ok());
        assert_eq!(g.keyword_alias("ipsum"), Some("if"));
        g.reset();
        assert!(g.phrases().is_empty() && g.keyword_alias_count() == 0 && g.fn_of_count() == 14);
    }
}
