//! The expression and condition grammars (PORT.6, slice 6d), ported from
//! VLA_SentenceEngine.bas arm for arm: `ParseExpr`, `ParseSum`, `ParseProd`,
//! `ParsePrim` and `ParsePrimCore` (with LX.14's `TryFunctionPhrase` and
//! `CompleteFunctionPhrase`, the `<word> of <value>` function words, G8's
//! ordinals, B7's `item n of`, the cell readings and V3's keyed read),
//! `ReadTextRef`, `ParseExprReq`; `ParseCond`, `ParseCondSimple`,
//! `TryRangeContains`, `ParseCondReq`; and the name checks a matched
//! sentence meets, `CheckName`, `RefuseCellShaped` and
//! `RefuseCellShapedRead` (LX.13).
//!
//! Precedence, loosest to tightest: `joined with` / `followed by` (&), then
//! `plus` / `minus`, then `times` / `divided by` / `multiplied by`, then a
//! primary with its `%` postfix. Each level backtracks when the right
//! operand fails, so a word like `times` in "Repeat 10 times:" is left for
//! the sentence.

use super::matcher::{
    alt_desc, is_cell_part, is_num_tok, is_str_tok, port_pending, render_tok, vla_string_lit,
    Parser,
};
use super::rules::{slot_tok, surface_match};
use super::tokenize::{line_suf, PARA_TOK};
use super::words::{is_engine_call_name, is_reserved_name, ordinal_word, slot_desc};
use crate::form::Form;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::printer::write_datum;

/// VBA `FirstCellShapedWord` (LX.13): the first bare word in VLA text that
/// is shaped like a cell, or ""; quoted text is skipped.
pub fn first_cell_shaped_word(vla: &str) -> String {
    let chars: Vec<char> = vla.chars().collect();
    let n = chars.len();
    let mut i = 0;
    while i < n {
        let ch = chars[i];
        if ch == '"' {
            i += 1;
            while i < n {
                let ch = chars[i];
                if ch == '\\' {
                    i += 2;
                } else if ch == '"' {
                    break;
                } else {
                    i += 1;
                }
            }
            i += 1;
        } else if matches!(ch, '(' | ')' | ' ' | '\t' | '\r' | '\n') {
            i += 1;
        } else {
            let mut word = String::new();
            while i < n {
                let ch = chars[i];
                if matches!(ch, '(' | ')' | ' ' | '\t' | '\r' | '\n' | '"') {
                    break;
                }
                word.push(ch);
                i += 1;
            }
            let folded = fold(&word);
            if is_cell_part(&folded) {
                return folded;
            }
        }
    }
    String::new()
}

/// VBA `CellWordShown` (LX.13): a cell-shaped word as a cell is written,
/// its letters in capitals; ASCII only, `Fold`'s mirror.
pub fn cell_word_shown(word: &str) -> String {
    word.chars()
        .map(|c| {
            if c.is_ascii_lowercase() {
                c.to_ascii_uppercase()
            } else {
                c
            }
        })
        .collect()
}

impl<'g> Parser<'g> {
    /// VBA `ParseExpr`: `joined with` / `followed by` over sums.
    pub(super) fn parse_expr(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let Some(mut a) = self.parse_sum(pos)? else {
            return Ok(None);
        };
        loop {
            let sp = *pos;
            if self.match_words(pos, "joined with") || self.match_words(pos, "followed by") {
                match self.parse_sum(pos)? {
                    Some(b) => a = format!("(& {a} {b})"),
                    None => {
                        *pos = sp;
                        break;
                    }
                }
            } else {
                break;
            }
        }
        Ok(Some(a))
    }

    /// VBA `ParseSum`: `plus` / `minus` over products.
    fn parse_sum(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let Some(mut a) = self.parse_prod(pos)? else {
            return Ok(None);
        };
        loop {
            let op = match self.tok(*pos) {
                "plus" => "+",
                "minus" => "-",
                _ => break,
            };
            let sp = *pos;
            *pos += 1;
            match self.parse_prod(pos)? {
                Some(b) => a = format!("({op} {a} {b})"),
                None => {
                    *pos = sp;
                    break;
                }
            }
        }
        Ok(Some(a))
    }

    /// VBA `ParseProd`: `times` / `divided by` / `multiplied by` over
    /// primaries.
    fn parse_prod(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let Some(mut a) = self.parse_prim(pos)? else {
            return Ok(None);
        };
        loop {
            let sp = *pos;
            let op = if self.tok(*pos) == "times" {
                *pos += 1;
                "*"
            } else if self.match_words(pos, "divided by") {
                "/"
            } else if self.match_words(pos, "multiplied by") {
                "*"
            } else {
                break;
            };
            match self.parse_prim(pos)? {
                Some(b) => a = format!("({op} {a} {b})"),
                None => {
                    *pos = sp;
                    break;
                }
            }
        }
        Ok(Some(a))
    }

    /// VBA `ParsePrim`: a primary and its `%` / `percent` postfixes (B6.3,
    /// B7.2), each dividing by 100 again, as in Excel.
    pub(super) fn parse_prim(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let Some(mut r) = self.parse_prim_core(pos)? else {
            return Ok(None);
        };
        while self.tok(*pos) == "%" || self.tok(*pos) == "percent" {
            *pos += 1;
            r = format!("(/ {r} 100)");
        }
        Ok(Some(r))
    }

    /// B6: a `To ... using` value action of the program's own (slice 6e: a
    /// proof's sentence has no program around it, so none is ever declared).
    fn is_using_fn(&self, _word: &str) -> bool {
        false
    }

    /// V3: a lookup declared by "Create a lookup called ..." in the current
    /// translation (slice 6e declares them).
    fn is_dict_name(&self, w: &str) -> bool {
        self.dict_names.iter().any(|d| d == w)
    }

    /// VBA `ParsePrimCore`.
    fn parse_prim_core(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let t = self.tok(*pos).to_string();
        // LX.14: a function phrase, longest first, ahead of any one-word
        // reading. Once its words match it completes, or it is refused there.
        if let Some(n2) = self.try_function_phrase(pos)? {
            return Ok(Some(n2));
        }
        let fn_of = self.g.fn_of_target(&t).map(str::to_string);
        if is_num_tok(&t) {
            *pos += 1;
            Ok(Some(t))
        } else if is_str_tok(&t) {
            *pos += 1;
            Ok(Some(vla_string_lit(&t[1..])))
        } else if let (Some(target), true) = (fn_of, self.tok(*pos + 1) == "of") {
            // "<word> of <value>": binds tightly and composes - "length of
            // cell B2", "month of today". Read before a value word (LX.14).
            let sp = *pos;
            *pos += 2;
            match self.parse_prim(pos)? {
                Some(n2) => Ok(Some(format!("({target} {n2})"))),
                None => {
                    *pos = sp + 1; // no operand: the word is a name
                    Ok(Some(t))
                }
            }
        } else if let Some(target) = self.g.fn_nullary_target(&t) {
            let target = target.to_string();
            *pos += 1;
            Ok(Some(format!("({target})")))
        } else if !ordinal_word(&t).is_empty() {
            // G8: ordinal words, read only here and only once "<word> of"
            // has failed, so "first of X" never reaches this branch.
            *pos += 1;
            Ok(Some(ordinal_word(&t)))
        } else if (self.is_using_fn(&t) && self.tok(*pos + 1) == "using")
            || (t == "get" && self.is_using_fn(self.tok(*pos + 1)) && self.tok(*pos + 2) == "using")
        {
            Err(port_pending("a value call with 'using'"))
        } else if t == "problem" && self.in_recovery {
            // B7: inside an "If that fails:" paragraph, "the problem" is
            // what went wrong.
            *pos += 1;
            Ok(Some("vla-problem".to_string()))
        } else if t == "item" {
            // B7: "item <n> of <list>" fetches an element; falls back to the
            // plain name when the shape does not follow.
            let sp = *pos;
            *pos += 1;
            if let Some(n2) = self.parse_expr(pos)? {
                if self.tok(*pos) == "of" {
                    *pos += 1;
                    if let Some(cref) = self.parse_prim(pos)? {
                        return Ok(Some(format!("(vlaitem {cref} {n2})")));
                    }
                }
            }
            *pos = sp + 1;
            Ok(Some(t))
        } else if t == "cell" || (t == "value" && self.tok(*pos + 1) == "in") {
            // Cell references are values: "cell B2", "cell in column C row
            // k", and B5.2's "value in column ..." / "value in cell B2".
            let sp = *pos;
            *pos += 1;
            if self.tok(*pos) == "in" {
                *pos += 1;
                if t == "value" && self.tok(*pos) == "cell" {
                    *pos += 1;
                    if let Some(cref) = self.read_text_ref(pos) {
                        return Ok(Some(format!("(range {cref})")));
                    }
                } else if self.tok(*pos) == "column" {
                    *pos += 1;
                    let cref = if self.tok(*pos) == "number" {
                        // B5: "cell in column number <expr> row <expr>".
                        *pos += 1;
                        self.parse_expr(pos)?
                    } else {
                        self.read_text_ref(pos)
                    };
                    if let Some(cref) = cref {
                        if self.tok(*pos) == "row" {
                            *pos += 1;
                            if let Some(n2) = self.parse_expr(pos)? {
                                return Ok(Some(format!("(cells {n2} {cref})")));
                            }
                        }
                    }
                }
            } else if t == "cell" {
                if let Some(cref) = self.read_text_ref(pos) {
                    return Ok(Some(format!("(range {cref})")));
                }
            }
            *pos = sp + 1; // no reference followed: the word is a name
            Ok(Some(t))
        } else if !self.word_at(*pos).is_empty() {
            if t.contains(':') {
                Ok(None) // bare ranges belong in reference slots
            } else if self.is_dict_name(&t) && self.tok(*pos + 1) == "for" {
                // V3: the keyed read - "<lookup> for <key>".
                let sp = *pos;
                *pos += 2;
                match self.parse_prim(pos)? {
                    Some(n2) => Ok(Some(format!("(vladictget {t} {n2})"))),
                    None => {
                        *pos = sp + 1; // no key followed: the word is a name
                        Ok(Some(t))
                    }
                }
            } else {
                *pos += 1;
                Ok(Some(t))
            }
        } else {
            Ok(None)
        }
    }

    /// VBA `ReadTextRef`: one reference token (quoted, bare word, or
    /// number) as a VBA string literal.
    fn read_text_ref(&self, pos: &mut usize) -> Option<String> {
        let t = self.tok(*pos);
        if is_str_tok(t) {
            let r = vla_string_lit(&t[1..]);
            *pos += 1;
            Some(r)
        } else if is_num_tok(t) || super::matcher::is_word_tok(t) {
            let r = vla_string_lit(t);
            *pos += 1;
            Some(r)
        } else {
            None
        }
    }

    /// VBA `ParseExprReq`.
    pub(super) fn parse_expr_req(&mut self, pos: &mut usize) -> Result<String, Refusal> {
        match self.parse_expr(pos)? {
            Some(e) => Ok(e),
            None => Err(raise(
                "english-expected-value",
                &[("context", &self.sentence_context(*pos))],
            )),
        }
    }

    /// VBA `TryFunctionPhrase` (LX.14): the phrases beginning with this
    /// word, longest first; the first whose words all match completes.
    fn try_function_phrase(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let word = self.tok(*pos).to_string();
        let Some(bucket) = self.g.phrase_bucket(&word) else {
            return Ok(None);
        };
        let bucket: Vec<usize> = bucket.to_vec();
        for idx in bucket {
            let words = self.g.phrases()[idx - 1].words.clone();
            let mut p = *pos + 1;
            let mut all_words = true;
            for w in words.iter().skip(1) {
                self.skip_articles(&mut p);
                if self.tok(p) != w {
                    all_words = false;
                    break;
                }
                p += 1;
            }
            if all_words {
                let out = self.complete_function_phrase(idx, *pos, &mut p)?;
                *pos = p;
                return Ok(Some(out));
            }
        }
        Ok(None)
    }

    /// VBA `CompleteFunctionPhrase`: from the hole on - the value or the
    /// reference, then the closing clause - substituted into the template.
    fn complete_function_phrase(
        &mut self,
        idx: usize,
        start_pos: usize,
        p: &mut usize,
    ) -> Result<String, Refusal> {
        let ph = &self.g.phrases()[idx - 1];
        let cat = ph.hole_cat.clone();
        let hole_name = ph.hole_name.clone();
        let clause = ph.clause.clone();
        let template: Form = ph.forms[0].clone();
        let mut bn: Vec<String> = Vec::new();
        let mut bv: Vec<String> = Vec::new();
        let v = if cat == "value" {
            match self.parse_prim(p)? {
                Some(v) => v,
                None => return Err(self.refuse_phrase(idx, start_pos, *p, &slot_desc("value"))),
            }
        } else {
            let mut v = String::new();
            if !self.match_ref_token(&cat, p, &mut v) {
                return Err(self.refuse_phrase(idx, start_pos, *p, &slot_desc(&cat)));
            }
            v
        };
        bn.push(hole_name);
        bv.push(v);
        for it in &clause {
            self.skip_articles(p);
            if let Some(slot) = slot_tok(it)? {
                let mut hit = false;
                let mut stem = String::new();
                for alt in slot.cat.split('|') {
                    let (ok, s) = surface_match(&fold(alt), self.tok(*p));
                    if ok {
                        hit = true;
                        stem = s;
                        break;
                    }
                }
                if !hit {
                    return Err(self.refuse_phrase(idx, start_pos, *p, &alt_desc(&slot.cat)));
                }
                bn.push(slot.name.clone());
                bv.push(stem);
            } else if self.tok(*p) != it {
                return Err(self.refuse_phrase(idx, start_pos, *p, &format!("'{it}'")));
            }
            *p += 1;
        }
        Ok(write_datum(&super::matcher::form_substitute(
            &template, &bn, &bv,
        )?))
    }

    /// VBA `RefusePhrase` (LX.14): a phrase that stopped partway, refused in
    /// the teaching frame from where its sentence began.
    fn refuse_phrase(&mut self, idx: usize, start_pos: usize, p: usize, expected: &str) -> Refusal {
        let found_t = self.tok(p).to_string();
        let found = if found_t.is_empty() || found_t == "." || found_t == PARA_TOK {
            "the end of the sentence".to_string()
        } else {
            format!("'{}'", render_tok(&found_t))
        };
        let mut st = start_pos;
        if self.sentence_start >= 1 && self.sentence_start <= start_pos {
            st = self.sentence_start;
        }
        let understood = self.render_tokens(st, p.saturating_sub(1));
        let shape = self.g.phrases()[idx - 1].shape.clone();
        let loc = self.line_tag(p);
        raise(
            "english-phrase-incomplete",
            &[
                ("understood", &understood),
                ("expected", expected),
                ("found", &found),
                ("phrase", &shape),
                ("loc", &loc),
            ],
        )
    }

    // ---- conditions ---------------------------------------------------------

    /// VBA `ParseCond`: simple conditions joined by `and` / `or`.
    pub(super) fn parse_cond(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let Some(mut c) = self.parse_cond_simple(pos)? else {
            return Ok(None);
        };
        loop {
            let w = self.tok(*pos).to_string();
            if w != "and" && w != "or" {
                break;
            }
            let sp = *pos;
            *pos += 1;
            match self.parse_cond_simple(pos)? {
                Some(c2) => c = format!("({w} {c} {c2})"),
                None => {
                    *pos = sp;
                    break;
                }
            }
        }
        Ok(Some(c))
    }

    /// VBA `ParseCondSimple`: a value, a comparator phrase (longest first),
    /// and its value; the idiomatic conditions read through the P-layer's
    /// predicates (DF1).
    fn parse_cond_simple(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        if let Some(range_cond) = self.try_range_contains(pos)? {
            return Ok(Some(range_cond));
        }
        let Some(l) = self.parse_expr(pos)? else {
            return Ok(None);
        };
        // Nullary comparators first: no right-hand value follows.
        if self.match_words(pos, "is not empty") {
            return Ok(Some(format!("(not (blank? {l}))")));
        }
        if self.match_words(pos, "is empty") {
            return Ok(Some(format!("(blank? {l})")));
        }
        // Longest comparator phrases first, in the reference's order: each
        // row is the phrase, its operator, its idiom and whether it is the
        // divisibility test.
        const COMPARATORS: &[(&str, &str, &str, bool)] = &[
            ("is greater than or equal to", ">=", "", false),
            ("is less than or equal to", "<=", "", false),
            ("is at least", ">=", "", false),
            ("is at most", "<=", "", false),
            ("is greater than", ">", "", false),
            ("is more than", ">", "", false),
            ("is less than", "<", "", false),
            ("does not contain", "", "ncontains", false),
            ("contains", "", "contains", false),
            ("starts with", "", "starts", false),
            ("ends with", "", "ends", false),
            ("is divisible by", "", "", true),
            ("is not equal to", "<>", "", false),
            ("does not equal", "<>", "", false),
            ("is not", "<>", "", false),
            ("is equal to", "=", "", false),
            ("equals", "=", "", false),
            ("is", "=", "", false),
        ];
        let Some(&(_, op, kind, divisible)) = COMPARATORS
            .iter()
            .find(|(phrase, ..)| self.match_words(pos, phrase))
        else {
            return Ok(None);
        };
        let Some(r) = self.parse_expr(pos)? else {
            return Ok(None);
        };
        Ok(Some(if divisible {
            format!("(zero? (mod {l} {r}))")
        } else {
            match kind {
                "contains" => format!("(positive? (instr 1 {l} {r} vbtextcompare))"),
                "ncontains" => format!("(zero? (instr 1 {l} {r} vbtextcompare))"),
                "starts" => format!("(= (instr 1 {l} {r} vbtextcompare) 1)"),
                "ends" => format!("(= (lcase (right {l} (len {r}))) (lcase {r}))"),
                _ => format!("({op} {l} {r})"),
            }
        }))
    }

    /// VBA `TryRangeContains` (G-TEXT slice 3): "range A1:D50 contains x",
    /// "column C does not contain x", through `VlaFindText`.
    fn try_range_contains(&mut self, pos: &mut usize) -> Result<Option<String>, Refusal> {
        let kind = self.tok(*pos).to_string();
        if kind != "range" && kind != "column" {
            return Ok(None);
        }
        let mut p = *pos + 1;
        let mut reference = String::new();
        if !self.match_ref_token(&kind, &mut p, &mut reference) {
            return Ok(None);
        }
        let negated = if self.match_words(&mut p, "does not contain") {
            true
        } else if self.match_words(&mut p, "contains") {
            false
        } else {
            return Ok(None);
        };
        let Some(r) = self.parse_expr(&mut p)? else {
            return Ok(None);
        };
        let place = if kind == "range" {
            format!("(range {reference})")
        } else {
            format!("(columns {reference})")
        };
        let cond = if negated {
            format!("(zero? (vlafindtext {place} {r} \"row\"))")
        } else {
            format!("(positive? (vlafindtext {place} {r} \"row\"))")
        };
        *pos = p;
        Ok(Some(cond))
    }

    /// VBA `ParseCondReq`.
    #[allow(dead_code)] // the If, While and Try forms read it (slice 6e)
    pub(super) fn parse_cond_req(&mut self, pos: &mut usize) -> Result<String, Refusal> {
        match self.parse_cond(pos)? {
            Some(c) => Ok(c),
            None => Err(raise(
                "english-expected-condition",
                &[("context", &self.sentence_context(*pos))],
            )),
        }
    }

    // ---- names ------------------------------------------------------------------

    /// VBA `CheckName`: a reserved word, a cell-shaped word (LX.13), a value
    /// word, or a name the generated code calls (U.30) cannot be a name.
    pub(super) fn check_name(&mut self, n: &str) -> Result<(), Refusal> {
        if is_reserved_name(n) {
            return Err(raise("english-reserved-word-name", &[("name", n)]));
        }
        let folded = fold(n);
        if is_cell_part(&folded) {
            return Err(self.refuse_cell_shaped(&folded, self.cur_line));
        }
        if self.g.fn_nullary_target(n).is_some() {
            return Err(raise("english-value-word-name", &[("name", n)]));
        }
        // U.30: last, so date and time keep the reserved-word refusal and
        // the value-word one.
        if is_engine_call_name(n) {
            let loc = if self.cur_line > 0 {
                self.err_line = self.cur_line;
                line_suf(self.cur_line)
            } else {
                String::new()
            };
            return Err(raise(
                "english-engine-call-name",
                &[("name", &folded), ("loc", &loc)],
            ));
        }
        Ok(())
    }

    /// VBA `RefuseCellShaped` (LX.13): the one refusal for a word shaped
    /// like a cell where a name is made or a value is read.
    fn refuse_cell_shaped(&mut self, word: &str, ln: u32) -> Refusal {
        let loc = if ln > 0 {
            self.err_line = ln;
            line_suf(ln)
        } else {
            String::new()
        };
        raise(
            "english-cell-shaped-name",
            &[
                ("name", word),
                ("cell", &cell_word_shown(word)),
                ("loc", &loc),
            ],
        )
    }

    /// VBA `RefuseCellShapedRead` (LX.13): a bare cell-shaped word in a
    /// finished translation is a cell written without the word "cell".
    pub(super) fn refuse_cell_shaped_read(&mut self, vla: &str, ln: u32) -> Result<(), Refusal> {
        let word = first_cell_shaped_word(vla);
        if word.is_empty() {
            Ok(())
        } else {
            Err(self.refuse_cell_shaped(&word, ln))
        }
    }

    /// VBA `MarkAssigned`: the name is checked and remembered.
    pub(super) fn mark_assigned(&mut self, name: &str) -> Result<(), Refusal> {
        self.check_name(name)?;
        let folded = fold(name);
        if !self.assigned.contains(&folded) {
            self.assigned.push(folded);
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_first_cell_shaped_word_skips_quoted_text() {
        assert_eq!(first_cell_shaped_word("(set! (range \"b2\") 5)"), "");
        assert_eq!(first_cell_shaped_word("(set! total (+ a2 1))"), "a2");
        assert_eq!(first_cell_shaped_word("(set! x \"a\\\"b2\")"), "");
        assert_eq!(first_cell_shaped_word("(x data!a1 :key abcd1)"), "");
        assert_eq!(cell_word_shown("fy24"), "FY24");
    }
}
