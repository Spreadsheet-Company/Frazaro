//! The DCG matcher (PORT.6, slice 6d): VLA_SentenceEngine.bas's `TryPhrase`
//! and what it leans on, one match arm per `Case` arm: `MatchRefToken` and
//! `MatchPathToken`, G2's reference shapes (`RefShapeOk`, `IsCellPart`,
//! `IsRangePart`), `NoteFail` and `BuildParseError` with `DidYouMean`, V7's
//! first-token dispatch index (`BuildDispatchIndex`), and F.2's form path
//! (`TryFormPath`, `FormSubstitute`, `SpliceEmbeddedSlots`). The slot
//! categories G-PROLOG added (`role`, `relation`, `conditions`, `clause`,
//! `question`) are slice 6e's and refuse by name until then.
//!
//! The reference keeps a translation's state in module-level variables
//! (`mBestProgress`, `mBestExpect`, `mBestTieIdx`, `mLastRuleIdx`,
//! `mSentenceStart`, `mClaim`, `mAssigned`, `mErrLine`, `mCurLine`); here
//! they are a [`Parser`] over one grammar, which also holds the token array
//! `EnTokenize` produced and `CanonicalizeStructuralWords` rewrote.

use std::collections::HashMap;

use super::grammar::{Grammar, Rule};
use super::rules::{
    bare_alt_match, bare_surfaces, is_alt_cat, opt_tok, slot_tok, surface_forms, surface_match,
};
use super::tokenize::{line_suf, Tokens, PARA_TOK};
use super::words::{is_color_word, slot_desc};
use crate::form::{Form, List};
use crate::intrinsics::{fold, is_numeric};
use crate::messages::{raise, Refusal};
use crate::printer::write_datum;
use crate::reader::read_forms;

/// A refusal for a shape the core does not read yet: a statement head or a
/// slot category that a later slice of PORT.6 ports. Not a catalogue
/// message, on purpose: it names the slice, never a user's mistake, and it
/// leaves with the slice that ports the shape.
pub fn port_pending(what: &str) -> Refusal {
    Refusal {
        id: "english-port-pending".to_string(),
        number: 5,
        source: "VLA-English".to_string(),
        text: format!("{what} is not yet in the core (PORT.6, slice 6e)"),
    }
}

/// VBA `IsWordTok`: a bare word begins with a letter (every bare word is
/// folded, so a lowercase one).
pub fn is_word_tok(t: &str) -> bool {
    matches!(t.as_bytes().first(), Some(b) if b.is_ascii_lowercase())
}

/// VBA `IsStrTok`: a quoted text's token carries the quote as its sigil.
pub fn is_str_tok(t: &str) -> bool {
    t.starts_with('"')
}

/// VBA `IsNumTok`: a token `IsNumeric` accepts that is not a quoted text.
pub fn is_num_tok(t: &str) -> bool {
    !t.is_empty() && !is_str_tok(t) && is_numeric(t)
}

/// VBA `RenderTok`: a quoted text's token shown with its quotes.
pub fn render_tok(t: &str) -> String {
    if is_str_tok(t) {
        format!("\"{}\"", &t[1..])
    } else {
        t.to_string()
    }
}

/// VBA `VlaStringLit`: a VLA string literal for a text, backslashes and
/// quotes escaped.
pub fn vla_string_lit(content: &str) -> String {
    let c = content.replace('\\', "\\\\").replace('"', "\\\"");
    format!("\"{c}\"")
}

// ---- G2: the reference shapes -------------------------------------------

/// VBA `IsColLetters`: one to three letters.
pub fn is_col_letters(s: &str) -> bool {
    (1..=3).contains(&s.len()) && s.bytes().all(|b| b.is_ascii_lowercase())
}

/// VBA `IsRowDigits`: one to seven digits.
pub fn is_row_digits(s: &str) -> bool {
    (1..=7).contains(&s.len()) && s.bytes().all(|b| b.is_ascii_digit())
}

/// VBA `IsCellPart`: letters then digits, both parts present and sized.
pub fn is_cell_part(s: &str) -> bool {
    let i = s.bytes().take_while(|b| b.is_ascii_lowercase()).count();
    if i == 0 || i >= s.len() {
        return false;
    }
    is_col_letters(&s[..i]) && is_row_digits(&s[i..])
}

/// VBA `IsRangePart`: a cell is a range; a span pairs two parts of one
/// kind, cells, column letters or row numbers.
pub fn is_range_part(s: &str) -> bool {
    let Some(cp) = s.find(':') else {
        return is_cell_part(s);
    };
    let a = &s[..cp];
    let b = &s[cp + 1..];
    if b.contains(':') {
        return false;
    }
    (is_cell_part(a) && is_cell_part(b))
        || (is_col_letters(a) && is_col_letters(b))
        || (is_row_digits(a) && is_row_digits(b))
}

/// VBA `RefShapeOk`: the shape a bare token must have for a typed
/// reference slot; a sheet qualifier (`data!b2`) is split off for a cell
/// or a range, and a bang at either end is no reference at all.
pub fn ref_shape_ok(cat: &str, tok: &str) -> bool {
    let mut tail = tok;
    if cat == "cell" || cat == "range" {
        if let Some(bp) = tail.rfind('!') {
            if bp == 0 || bp == tail.len() - 1 {
                return false;
            }
            tail = &tail[bp + 1..];
        }
    }
    match cat {
        "cell" => is_cell_part(tail),
        "range" => is_range_part(tail),
        "column" => is_col_letters(tail),
        "sheet" => true,
        "color" => is_color_word(tail),
        _ => false,
    }
}

/// VBA `AltDesc`: "one of 'into'/'in'", every surface of a stem/suffix
/// branch listed, for the near-miss reporting.
pub fn alt_desc(cat: &str) -> String {
    let mut r = String::new();
    for alt in cat.split('|') {
        for f in surface_forms(&fold(alt)) {
            if !r.is_empty() {
                r.push('/');
            }
            r.push('\'');
            r.push_str(&f);
            r.push('\'');
        }
    }
    format!("one of {r}")
}

// ---- F.2: the form path ----------------------------------------------------

/// VBA `FormSubstitute`: a template form with every `{slot}` replaced by
/// its bound form, a slot glued inside a larger atom spliced as text.
pub fn form_substitute(f: &Form, bn: &[String], bv: &[String]) -> Result<Form, Refusal> {
    match f {
        Form::List(l) => {
            let mut items = Vec::with_capacity(l.items.len());
            for e in &l.items {
                items.push(form_substitute(e, bn, bv)?);
            }
            Ok(Form::List(List {
                items,
                line: l.line,
            }))
        }
        Form::Sym(s) => {
            if s.starts_with('{') && s.ends_with('}') && s.len() > 2 && !s[1..].contains('{') {
                bound_lookup(&s[1..s.len() - 1], bn, bv)
            } else if s.contains('{') {
                Ok(Form::Sym(splice_embedded_slots(s, bn, bv)?))
            } else {
                Ok(f.clone())
            }
        }
        // The reference sees a string as its quote sigil and its content,
        // which never begins with a brace, so only the splice applies.
        Form::Str(s) => {
            if s.contains('{') {
                Ok(Form::Str(splice_embedded_slots(s, bn, bv)?))
            } else {
                Ok(f.clone())
            }
        }
    }
}

/// VBA `BoundLookup`: the one form a slot's bound text reads back as, or
/// the literal `{slotname}` when nothing binds that name.
fn bound_lookup(slot_name: &str, bn: &[String], bv: &[String]) -> Result<Form, Refusal> {
    for (k, name) in bn.iter().enumerate() {
        if name == slot_name {
            let bound = read_forms(&bv[k])?;
            if bound.len() != 1 {
                return Err(raise(
                    "english-slot-value-not-one-form",
                    &[
                        ("slot", slot_name),
                        ("count", &bound.len().to_string()),
                        ("value", &bv[k]),
                    ],
                ));
            }
            return Ok(bound.into_iter().next().unwrap_or(Form::Sym(String::new())));
        }
    }
    Ok(Form::Sym(format!("{{{slot_name}}}")))
}

/// VBA `SpliceEmbeddedSlots`: every `{slotname}` run inside an atom's text
/// replaced by what that slot contributes.
fn splice_embedded_slots(atom: &str, bn: &[String], bv: &[String]) -> Result<String, Refusal> {
    let mut r = String::new();
    let mut i = 0;
    while i < atom.len() {
        let Some(open_rel) = atom[i..].find('{') else {
            r.push_str(&atom[i..]);
            break;
        };
        let open = i + open_rel;
        let Some(close_rel) = atom[open..].find('}') else {
            r.push_str(&atom[i..]);
            break;
        };
        let close = open + close_rel;
        r.push_str(&atom[i..open]);
        r.push_str(&embedded_slot_text(&atom[open + 1..close], bn, bv)?);
        i = close + 1;
    }
    Ok(r)
}

/// VBA `EmbeddedSlotText`: only a bare symbol can be glued into an
/// identifier; a list or a quoted string there is refused by name.
fn embedded_slot_text(slot_name: &str, bn: &[String], bv: &[String]) -> Result<String, Refusal> {
    for (k, name) in bn.iter().enumerate() {
        if name == slot_name {
            let bound = read_forms(&bv[k])?;
            if bound.len() != 1 {
                return Err(raise(
                    "english-slot-value-not-one-form",
                    &[
                        ("slot", slot_name),
                        ("count", &bound.len().to_string()),
                        ("value", &bv[k]),
                    ],
                ));
            }
            return match &bound[0] {
                Form::Sym(text) => Ok(text.clone()),
                _ => Err(raise(
                    "english-slot-glued-to-identifier",
                    &[("quoted", &format!("{{{slot_name}}}")), ("value", &bv[k])],
                )),
            };
        }
    }
    Ok(format!("{{{slot_name}}}"))
}

/// VBA `TryFormPath`: each template form substituted and written, the
/// forms on their own lines.
pub fn try_form_path(rule: &Rule, bn: &[String], bv: &[String]) -> Result<String, Refusal> {
    let mut r = String::new();
    for f in &rule.forms {
        if !r.is_empty() {
            r.push_str("\r\n");
        }
        r.push_str(&write_datum(&form_substitute(f, bn, bv)?));
    }
    Ok(r)
}

// ---- V7: the first-token dispatch index ---------------------------------------

/// VBA `BuildDispatchIndex`'s product: per first token the rules that can
/// begin with it, the universal rules a slot begins, and for each bucketed
/// rule the description its first item fails with.
pub struct DispatchIndex {
    buckets: HashMap<String, Vec<usize>>,
    universal: Vec<usize>,
    /// 1-based, as the rules are; entry 0 is unused.
    fail_desc: Vec<String>,
    n: usize,
}

impl DispatchIndex {
    /// Walk each rule's items as `TryPhrase` would at position zero.
    pub fn build(g: &Grammar) -> DispatchIndex {
        let rules = g.rules();
        let n = rules.len();
        let mut buckets: HashMap<String, Vec<usize>> = HashMap::new();
        let mut universal = Vec::new();
        let mut fail_desc = vec![String::new(); n + 1];
        for (i0, rule) in rules.iter().enumerate() {
            let i = i0 + 1;
            let mut keys: Vec<String> = Vec::new();
            let mut is_universal = false;
            let mut done = false;
            let mut desc = String::new();
            for t in &rule.items {
                if let Ok(Some(slot)) = slot_tok(t) {
                    if is_alt_cat(&slot.cat) {
                        for f in bare_surfaces(&slot.cat) {
                            dsp_add_key(&mut keys, f);
                        }
                        desc = alt_desc(&slot.cat);
                    } else {
                        is_universal = true;
                    }
                    done = true;
                } else if let Ok(Some(ow)) = opt_tok(t) {
                    // Free when absent - keep walking: the next item can
                    // also begin the rule.
                    for f in bare_surfaces(&ow) {
                        dsp_add_key(&mut keys, f);
                    }
                } else if t.contains('|') || t.contains('/') {
                    for f in bare_surfaces(t) {
                        dsp_add_key(&mut keys, f);
                    }
                    desc = alt_desc(t);
                    done = true;
                } else {
                    dsp_add_key(&mut keys, t.clone());
                    desc = format!("'{t}'");
                    done = true;
                }
                if done {
                    break;
                }
            }
            if !done {
                is_universal = true;
            }
            if is_universal {
                universal.push(i);
            } else {
                fail_desc[i] = desc;
                for k in keys {
                    buckets.entry(k).or_default().push(i);
                }
            }
        }
        DispatchIndex {
            buckets,
            universal,
            fail_desc,
            n,
        }
    }
}

/// VBA `DspAddKey`: a distinct key (left|left contributes one entry).
fn dsp_add_key(keys: &mut Vec<String>, k: String) {
    if !keys.contains(&k) {
        keys.push(k);
    }
}

// ---- the parser ---------------------------------------------------------------

/// One translation's state over one grammar: the token array and the
/// module-level variables the reference keeps beside it.
pub struct Parser<'g> {
    pub(super) g: &'g Grammar,
    pub(super) toks: Tokens,
    dsp: DispatchIndex,
    /// `mBestProgress`: the furthest token position any rule reached
    /// (-1 before any rule tried).
    best_progress: i64,
    /// `mBestExpect`: what it wanted there.
    best_expect: String,
    /// `mBestTieIdx`: every rule tied at the furthest progress, up to three.
    best_ties: Vec<usize>,
    /// `mLastRuleIdx`: the rule the last `TryPhrase` matched.
    pub(super) last_rule_idx: usize,
    /// `mSentenceStart`: where the sentence began, for a phrase refusal.
    pub(super) sentence_start: usize,
    /// `mClaim`: the construct that claimed the sentence; first claim wins.
    pub(super) claim: String,
    /// `mAssigned`: the names `:var` slots and loops bound.
    pub(super) assigned: Vec<String>,
    /// `mErrLine`: the line of the most recent refusal.
    pub(super) err_line: u32,
    /// `mCurLine`: the sentence's program line, 0 for a proof.
    pub(super) cur_line: u32,
    /// `mInRecovery`: inside an "If that fails:" paragraph (slice 6e).
    pub(super) in_recovery: bool,
    /// `mDictNames`: the lookups "Create a lookup called ..." declared
    /// (slice 6e).
    pub(super) dict_names: Vec<String>,
    /// `mLoopStack`: the enclosing loop kinds ("for"/"do"), which the
    /// block parsers push (slice 6e).
    pub(super) loop_stack: Vec<String>,
    /// `mInFuncDef`: inside a value-returning action's body (slice 6e).
    pub(super) in_func_def: bool,
}

impl<'g> Parser<'g> {
    /// A parser over a grammar, its dispatch index built once.
    pub fn new(g: &'g Grammar) -> Parser<'g> {
        Parser {
            g,
            toks: Tokens::default(),
            dsp: DispatchIndex::build(g),
            best_progress: -1,
            best_expect: String::new(),
            best_ties: Vec::new(),
            last_rule_idx: 0,
            sentence_start: 0,
            claim: String::new(),
            assigned: Vec::new(),
            err_line: 0,
            cur_line: 0,
            in_recovery: false,
            dict_names: Vec::new(),
            loop_stack: Vec::new(),
            in_func_def: false,
        }
    }

    /// VBA `CurrentLoop`: the innermost enclosing loop's kind, or "".
    pub(super) fn current_loop(&self) -> &str {
        self.loop_stack.last().map(String::as_str).unwrap_or("")
    }

    /// The tokens of the next text, rewritten as `CanonicalizeStructuralWords`
    /// rewrites them (LX5.1: a bare word a phrasebook aliased becomes its
    /// canonical English spelling), and the per-sentence state fresh.
    pub fn set_tokens(&mut self, mut toks: Tokens) {
        if self.g.keyword_alias_count() > 0 {
            for i in 1..=toks.count() {
                let canon = {
                    let t = toks.tok_at(i);
                    if is_word_tok(t) {
                        self.g.keyword_alias(t).map(str::to_string)
                    } else {
                        None
                    }
                };
                if let Some(c) = canon {
                    toks.replace_at(i, &c);
                }
            }
        }
        self.toks = toks;
        self.sentence_start = 0;
        self.claim.clear();
        self.last_rule_idx = 0;
        self.assigned.clear();
        self.err_line = 0;
    }

    /// VBA `TokAt`: the token at a 1-based position, or nothing.
    pub(super) fn tok(&self, p: usize) -> &str {
        self.toks.tok_at(p)
    }

    /// VBA `WordAt`: the word at p, or "" if the token there isn't a usable
    /// word (`done` and `otherwise` close blocks and are never words).
    pub(super) fn word_at(&self, p: usize) -> String {
        let t = self.tok(p);
        if !is_word_tok(t) || t == "done" || t == "otherwise" {
            String::new()
        } else {
            t.to_string()
        }
    }

    /// VBA `PeekWord`.
    pub(super) fn peek_word(&self, p: usize) -> String {
        let t = self.tok(p);
        if is_word_tok(t) {
            t.to_string()
        } else {
            String::new()
        }
    }

    /// VBA `MatchWords`: a space-separated word sequence; advances pos only
    /// on success.
    pub(super) fn match_words(&self, pos: &mut usize, phrase: &str) -> bool {
        let parts: Vec<&str> = phrase.split(' ').collect();
        for (i, part) in parts.iter().enumerate() {
            if self.tok(*pos + i) != *part {
                return false;
            }
        }
        *pos += parts.len();
        true
    }

    /// VBA `SkipArticles`: past `a`/`an` where a grammar word is expected.
    pub(super) fn skip_articles(&self, p: &mut usize) {
        while self.tok(*p) == "a" || self.tok(*p) == "an" {
            *p += 1;
        }
    }

    /// VBA `SentenceContext`: up to twelve tokens from pos, as written.
    pub(super) fn sentence_context(&self, pos: usize) -> String {
        let mut r = String::new();
        for i in pos..pos + 12 {
            let t = self.tok(i);
            if t.is_empty() || t == PARA_TOK {
                break;
            }
            let t = render_tok(t);
            // Punctuation attaches to the preceding word, as written.
            if !r.is_empty() && t != "." && t != "," && t != ":" {
                r.push(' ');
            }
            r.push_str(&t);
            if t == "." {
                break;
            }
        }
        r
    }

    /// VBA `RenderTokens`: tokens first..=last as readable text.
    pub(super) fn render_tokens(&self, first: usize, last: usize) -> String {
        let mut r = String::new();
        let mut i = first;
        while i <= last {
            if !r.is_empty() {
                r.push(' ');
            }
            r.push_str(&render_tok(self.tok(i)));
            i += 1;
        }
        r
    }

    /// VBA `RenderSentenceAt`: the sentence beginning at start, a raw VLA
    /// token alone and trimmed.
    pub(super) fn render_sentence_at(&self, start: usize) -> String {
        let t = self.tok(start);
        if t.starts_with('(') {
            let chars: Vec<char> = t.chars().collect();
            if chars.len() > 60 {
                let head: String = chars[..57].iter().collect();
                return format!("{head}...");
            }
            return t.to_string();
        }
        let mut r = String::new();
        for i in start..start + 30 {
            let t = self.tok(i);
            if t.is_empty() || t == PARA_TOK {
                break;
            }
            let t = render_tok(t);
            if !r.is_empty() && t != "." && t != "," && t != ":" {
                r.push(' ');
            }
            r.push_str(&t);
            if t == "." || t == ":" {
                break;
            }
        }
        r
    }

    /// VBA `LineTag`: " (line N)" for a refusal, recording N as the last
    /// error line.
    pub(super) fn line_tag(&mut self, pos: usize) -> String {
        self.err_line = self.toks.line_at(pos);
        line_suf(self.err_line)
    }

    /// VBA `Claim`: first claim wins within a sentence.
    pub(super) fn claim(&mut self, what: &str) {
        if self.claim.is_empty() {
            self.claim = what.to_string();
        }
    }

    /// VBA `ExpectWord`.
    pub(super) fn expect_word(&mut self, pos: &mut usize, what: &str) -> Result<String, Refusal> {
        let w = self.word_at(*pos);
        if w.is_empty() {
            let context = self.sentence_context(*pos);
            let loc = self.line_tag(*pos);
            return Err(raise(
                "english-expected-near",
                &[("what", what), ("context", &context), ("loc", &loc)],
            ));
        }
        if w.contains(':') {
            let context = self.sentence_context(*pos);
            let loc = self.line_tag(*pos);
            return Err(raise(
                "english-name-has-colon",
                &[("word", &w), ("context", &context), ("loc", &loc)],
            ));
        }
        *pos += 1;
        Ok(w)
    }

    /// VBA `ExpectWordIs`.
    pub(super) fn expect_word_is(&mut self, pos: &mut usize, word: &str) -> Result<(), Refusal> {
        if self.tok(*pos) != word {
            let context = self.sentence_context(*pos);
            let loc = self.line_tag(*pos);
            return Err(raise(
                "english-expected-word-near",
                &[("word", word), ("context", &context), ("loc", &loc)],
            ));
        }
        *pos += 1;
        Ok(())
    }

    /// VBA `ExpectTok`.
    pub(super) fn expect_tok(
        &mut self,
        pos: &mut usize,
        t: &str,
        what: &str,
    ) -> Result<(), Refusal> {
        if self.tok(*pos) != t {
            let context = self.sentence_context(*pos);
            let loc = self.line_tag(*pos);
            return Err(raise(
                "english-expected-near",
                &[("what", what), ("context", &context), ("loc", &loc)],
            ));
        }
        *pos += 1;
        Ok(())
    }

    /// VBA `NoteFail`: keep the furthest failure, and up to three rules
    /// tied at it.
    fn note_fail(&mut self, p: usize, expected: String, rule_idx: usize) {
        let p = p as i64;
        if p > self.best_progress {
            self.best_progress = p;
            self.best_expect = expected;
            self.best_ties = vec![rule_idx];
        } else if p == self.best_progress
            && self.best_ties.len() < 3
            && !self.best_ties.contains(&rule_idx)
        {
            self.best_ties.push(rule_idx);
        }
    }

    /// VBA `MatchRefToken`: a quoted string, bare word or number becomes a
    /// VBA string literal; a bare token is shape-checked for its category.
    pub(super) fn match_ref_token(&self, cat: &str, p: &mut usize, val: &mut String) -> bool {
        let t = self.tok(*p);
        if is_str_tok(t) {
            *val = vla_string_lit(&t[1..]);
            *p += 1;
            return true;
        }
        if !self.word_at(*p).is_empty() || is_num_tok(t) {
            if cat != "text" && !ref_shape_ok(cat, t) {
                return false;
            }
            *val = vla_string_lit(t);
            *p += 1;
            return true;
        }
        false
    }

    /// VBA `MatchPathToken` (G-PATH): a quoted path, or a bare word naming
    /// a variable that holds one; a bare word followed by a lone colon is
    /// an unquoted path, refused with directions.
    fn match_path_token(&mut self, p: &mut usize, val: &mut String) -> Result<bool, Refusal> {
        let t = self.tok(*p).to_string();
        if is_str_tok(&t) {
            *val = vla_string_lit(&t[1..]);
            *p += 1;
            return Ok(true);
        }
        if self.word_at(*p).is_empty() {
            return Ok(false);
        }
        if self.tok(*p + 1) == ":" {
            let context = self.sentence_context(*p);
            let loc = self.line_tag(*p);
            return Err(raise(
                "english-path-not-quoted",
                &[("context", &context), ("loc", &loc)],
            ));
        }
        *val = t;
        *p += 1;
        Ok(true)
    }

    /// VBA `TryPhrase`: one rule against the tokens at pos. `Ok(Some)` with
    /// the rule's VLA and pos past the sentence's period; `Ok(None)` when
    /// the rule does not match, its nearest miss noted.
    pub(super) fn try_phrase(
        &mut self,
        idx: usize,
        pos: &mut usize,
    ) -> Result<Option<String>, Refusal> {
        let g: &'g Grammar = self.g;
        let rule = &g.rules()[idx - 1];
        let mut p = *pos;
        let mut bn: Vec<String> = Vec::new();
        let mut bv: Vec<String> = Vec::new();
        let mut var_names: Vec<String> = Vec::new();
        for t in &rule.items {
            if let Some(slot) = slot_tok(t)? {
                let cat = slot.cat.as_str();
                let mut match_ok = true;
                let mut val = String::new();
                match cat {
                    "name" | "var" => {
                        let w = self.word_at(p);
                        if w.is_empty() || w.contains(':') {
                            match_ok = false;
                        } else {
                            p += 1;
                            if cat == "var" {
                                var_names.push(w.clone());
                            }
                            val = w;
                        }
                    }
                    "text" | "range" | "cell" | "column" | "sheet" | "color" => {
                        // A reference: quoted string, bare word, or number -
                        // always a VBA string literal; the typed categories
                        // shape-check the bare token (G2).
                        match_ok = self.match_ref_token(cat, &mut p, &mut val);
                    }
                    "path" => {
                        match_ok = self.match_path_token(&mut p, &mut val)?;
                    }
                    "text-list" | "range-list" | "cell-list" | "column-list" | "sheet-list"
                    | "color-list" => {
                        // G6: one or more items, comma-separated, the Oxford
                        // comma required: "and" is consumed only right after
                        // a comma. Lowers to `(array item1 item2 ...)`.
                        let item_cat = &cat[..cat.len() - "-list".len()];
                        match_ok = self.match_ref_token(item_cat, &mut p, &mut val);
                        if match_ok {
                            val = format!("(array {val}");
                            while self.tok(p) == "," {
                                let mut try_p = p + 1;
                                if self.tok(try_p) == "and" {
                                    try_p += 1;
                                }
                                let mut item_val = String::new();
                                if !self.match_ref_token(item_cat, &mut try_p, &mut item_val) {
                                    break;
                                }
                                val.push(' ');
                                val.push_str(&item_val);
                                p = try_p;
                            }
                            val.push(')');
                        }
                    }
                    "expr" => match self.parse_expr(&mut p)? {
                        Some(v) => val = v,
                        None => match_ok = false,
                    },
                    "cond" => match self.parse_cond(&mut p)? {
                        Some(v) => val = v,
                        None => match_ok = false,
                    },
                    "role" | "relation" | "conditions" | "clause" | "question" => {
                        return Err(port_pending(&format!(
                            "the '{cat}' slot (G-PROLOG's sub-grammars)"
                        )));
                    }
                    _ => {
                        if !is_alt_cat(cat) {
                            return Err(raise(
                                "english-unknown-slot-category-runtime",
                                &[("cat", cat)],
                            ));
                        }
                        // G1: alternation - the sentence token must be one
                        // of the branch literals, and the MATCHED literal
                        // binds into the template. Branches try in written
                        // order; articles skip first, as before any literal.
                        self.skip_articles(&mut p);
                        let mut hit = false;
                        for alt in cat.split('|') {
                            let (ok, stem) = surface_match(&fold(alt), self.tok(p));
                            if ok {
                                val = stem;
                                p += 1;
                                hit = true;
                                break;
                            }
                        }
                        if !hit {
                            match_ok = false;
                        }
                    }
                }
                if !match_ok {
                    // G7: a default makes the slot's absence a successful
                    // match; no token is consumed.
                    if slot.has_default {
                        val = slot.default.clone();
                    } else if is_alt_cat(cat) {
                        self.note_fail(p, alt_desc(cat), idx);
                        return Ok(None);
                    } else {
                        self.note_fail(p, slot_desc(cat), idx);
                        return Ok(None);
                    }
                }
                bn.push(slot.name.clone());
                bv.push(val);
            } else if let Some(ow) = opt_tok(t)? {
                // G1: optional literal - consumed when present, free when
                // absent, never failing the match. G10: it may be a bare
                // alternation.
                self.skip_articles(&mut p);
                if bare_alt_match(&ow, self.tok(p)) {
                    p += 1;
                }
            } else if t.contains('|') || t.contains('/') {
                // G10: a bare surface token matches one branch surface and
                // binds nothing.
                self.skip_articles(&mut p);
                if !bare_alt_match(t, self.tok(p)) {
                    self.note_fail(p, alt_desc(t), idx);
                    return Ok(None);
                }
                p += 1;
            } else {
                self.skip_articles(&mut p);
                if self.tok(p) != t {
                    self.note_fail(p, format!("'{t}'"), idx);
                    return Ok(None);
                }
                p += 1;
            }
        }

        self.skip_articles(&mut p);
        if self.tok(p) != "." {
            self.note_fail(p, "'.' to end the sentence".to_string(), idx);
            return Ok(None);
        }
        p += 1;

        // Success: the var slots' names are made (CheckName runs only now,
        // after the whole rule matched), then the bindings go into the
        // template via the cached forms.
        for v in &var_names {
            self.mark_assigned(v)?;
        }
        *pos = p;
        let out = try_form_path(rule, &bn, &bv)?;
        self.last_rule_idx = idx;
        Ok(Some(out))
    }

    /// `ParseStmt`'s rule walk (V7): the bucket keyed by the sentence's
    /// first non-article token, merged with the universal list in ascending
    /// rule index, each skipped rule's near miss simulated once per gap.
    pub(super) fn try_rules(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        self.best_progress = -1;
        self.best_expect.clear();
        self.best_ties.clear();
        let mut dsp_star = *pos;
        self.skip_articles(&mut dsp_star);
        let key = fold(self.tok(dsp_star));
        let bucket: Option<Vec<usize>> = self.dsp.buckets.get(&key).cloned();
        let n = self.dsp.n;
        let mut bi = 0usize;
        let mut ui = 0usize;
        let mut last = 0usize;
        loop {
            let nb = bucket
                .as_ref()
                .and_then(|b| b.get(bi))
                .copied()
                .unwrap_or(n + 1);
            let nu = self.dsp.universal.get(ui).copied().unwrap_or(n + 1);
            let next = nb.min(nu);
            // Simulate the skipped rules between the last candidate and
            // this one (or the tail, when no candidate remains).
            let sim = (next - 1).min(n);
            if last < sim {
                let d = self.dsp.fail_desc[last + 1].clone();
                if !d.is_empty() {
                    self.note_fail(dsp_star, d, last + 1);
                }
            }
            if next > n {
                break;
            }
            if let Some(txt) = self.try_phrase(next, pos)? {
                return Ok(Some(txt));
            }
            last = next;
            if next == nb {
                bi += 1;
            } else {
                ui += 1;
            }
        }
        Ok(None)
    }

    /// VBA `BuildParseError`: the failure message for an ununderstood
    /// sentence, from the best partial match when a rule made progress,
    /// with up to three suggestions.
    pub(super) fn build_parse_error(&mut self, pos: usize) -> String {
        let progressed = self.best_progress > pos as i64;
        let mut msg = if progressed {
            let bp = self.best_progress as usize;
            let found_t = self.tok(bp);
            let found = if found_t.is_empty() || found_t == "." {
                "the end of the sentence".to_string()
            } else {
                format!("'{}'", render_tok(found_t))
            };
            format!(
                "I understood '{}' - then I expected {} but found {}.",
                self.render_tokens(pos, bp - 1),
                self.best_expect,
                found
            )
        } else {
            format!("Don't understand: '{}'", self.sentence_context(pos))
        };
        let ties: Option<Vec<usize>> = if progressed {
            Some(self.best_ties.clone())
        } else {
            None
        };
        let first_word = self.peek_word(pos);
        let suggest = self.did_you_mean(&first_word, ties.as_deref());
        if !suggest.is_empty() {
            msg.push_str(&format!(" Did you mean: {suggest}"));
        } else if self.g.rule_count() <= self.g.prelude_count() {
            // Nothing beyond the prelude is registered: the phrasebook did
            // not load. The message says what happened, never how to fix
            // it in VBA (the owner's rule, 2026-09-26).
            msg.push_str(&format!(
                " Only the {} built-in phrases are loaded: Frazaro's phrasebook did not load, so it can read almost no sentence. If this happens again after Validate Instructions, reinstall Frazaro.",
                self.g.rule_count()
            ));
        } else {
            msg.push_str(&format!(
                " No loaded sentence starts with '{}' - press What can I say? on the Frazaro tab to see all {} sentences this program understands.",
                first_word,
                self.g.rule_count()
            ));
        }
        let ep = if progressed {
            self.best_progress as usize
        } else {
            pos
        };
        let tag = self.line_tag(ep);
        msg.push_str(&tag);
        msg
    }

    /// VBA `DidYouMean` (LE.2): the rules tied for the furthest progress
    /// lead; the first-word scan fills what they leave, never repeating one.
    fn did_you_mean(&self, first_word: &str, ties: Option<&[usize]>) -> String {
        let rules = self.g.rules();
        let mut r = String::new();
        let mut hits = 0;
        if let Some(ties) = ties {
            for &tv in ties {
                if hits > 0 {
                    r.push_str("  |  ");
                }
                r.push_str(&format!("'{}'", rules[tv - 1].text));
                hits += 1;
            }
        }
        if !first_word.is_empty() && hits < 3 {
            for (i0, rule) in rules.iter().enumerate() {
                let i = i0 + 1;
                if ties.is_some_and(|t| t.contains(&i)) {
                    continue;
                }
                if rule.items.first().is_some_and(|f| f == first_word) {
                    if hits > 0 {
                        r.push_str("  |  ");
                    }
                    r.push_str(&format!("'{}'", rule.text));
                    hits += 1;
                    if hits == 3 {
                        break;
                    }
                }
            }
        }
        r
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_reference_shapes_are_the_vba_s() {
        assert!(is_cell_part("a1"));
        assert!(is_cell_part("aa100"));
        assert!(is_cell_part("fy24"));
        assert!(!is_cell_part("abcd1"));
        assert!(!is_cell_part("a"));
        assert!(!is_cell_part("12"));
        assert!(is_range_part("a1:c50"));
        assert!(is_range_part("a:c"));
        assert!(is_range_part("1:5"));
        assert!(!is_range_part("a1:c"));
        assert!(!is_range_part("a1:b2:c3"));
        assert!(ref_shape_ok("cell", "data!b2"));
        assert!(!ref_shape_ok("cell", "!b2"));
        assert!(!ref_shape_ok("cell", "b2!"));
        assert!(ref_shape_ok("column", "cat"));
        assert!(ref_shape_ok("sheet", "anything"));
        assert!(ref_shape_ok("color", "red"));
        assert!(!ref_shape_ok("color", "plaid"));
    }

    #[test]
    fn a_string_literal_escapes_as_the_vba_does() {
        assert_eq!(vla_string_lit("a\\b\"c"), "\"a\\\\b\\\"c\"");
        assert_eq!(render_tok("\"hi"), "\"hi\"");
        assert_eq!(render_tok("hi"), "hi");
        assert!(is_num_tok("1000"));
        assert!(is_num_tok("-2.5"));
        assert!(!is_num_tok("\"1"));
        assert!(!is_num_tok("b2"));
        assert!(is_word_tok("abc"));
        assert!(!is_word_tok("1abc"));
    }

    #[test]
    fn form_substitution_binds_splices_and_refuses_glue() {
        let bn = vec!["r".to_string(), "d".to_string(), "e".to_string()];
        let bv = vec![
            "\"b2\"".to_string(),
            "bold".to_string(),
            "(+ 1 2)".to_string(),
        ];
        let t = read_forms("(make-{d} (range {r}) {e} {zz})").unwrap();
        let out = write_datum(&form_substitute(&t[0], &bn, &bv).unwrap());
        assert_eq!(out, "(make-bold (range \"b2\") (+ 1 2) {zz})");
        let t = read_forms("(x vb{r})").unwrap();
        let e = form_substitute(&t[0], &bn, &bv).unwrap_err();
        assert_eq!(e.id, "english-slot-glued-to-identifier");
        let t = read_forms("(x vb{e})").unwrap();
        let e = form_substitute(&t[0], &bn, &bv).unwrap_err();
        assert_eq!(e.id, "english-slot-glued-to-identifier");
        assert_eq!(
            alt_desc("left|right|center/ed"),
            "one of 'left'/'right'/'center'/'centered'"
        );
    }
}
