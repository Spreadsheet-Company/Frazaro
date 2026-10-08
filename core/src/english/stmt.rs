//! The statement grammar (PORT.6, slices 6d and 6e): VLA_SentenceEngine.bas's
//! `ParseStmt` arm for arm - the blank-line and raw-form gates, If with its
//! Otherwise chain, Repeat, Count, Stop, While, For each, the standalone-Get
//! guard, Give back, Try with "If that fails:", When ... is, Create, the
//! To and Define guards, B7.4's percent form of Increase, Decrease and Add,
//! the rule walk, the action-call fallback and the parse error - with the
//! block parsers (`ParseBlock`, `ParseBranchBlock`, `ParseTryBody`,
//! `ContinuationAhead`), `ParseCaseValues`, `ValidateRawVla` and
//! `ParseTracked`, which numbers each statement for the step table; and the
//! proof runners `RunVocabTest`, `RunVocabFailTest` and `NormalizeWs`.

use super::grammar::{Grammar, Proof, ProofFailure, ProofKind, ProofMode};
use super::matcher::{add_keyed, is_num_tok, is_str_tok, is_word_tok, render_tok, Parser};
use super::tokenize::{tokenize, PARA_TOK};
use super::vocab::prov_loc;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// VBA `NormalizeWs`: all whitespace runs collapsed to single spaces, for
/// the proof comparison.
pub fn normalize_ws(s: &str) -> String {
    let mut s = s.replace("\r\n", " ").replace(['\n', '\t'], " ");
    while s.contains("  ") {
        s = s.replace("  ", " ");
    }
    s.trim_matches(' ').to_string()
}

/// VBA `StrConv(w, vbProperCase)` on one word.
fn proper_case(w: &str) -> String {
    let mut chars = w.chars();
    match chars.next() {
        Some(c) => c.to_ascii_uppercase().to_string() + chars.as_str(),
        None => String::new(),
    }
}

/// VBA `JoinStmts`: statements on their own lines.
pub fn join_stmts(stmts: &[String]) -> String {
    stmts.join("\r\n")
}

/// VBA `VlaBalanceHint` (VLA.bas, L12): the parenthesis count of a row in
/// words, or "" when it balances; quoted text and comments are skipped.
pub fn vla_balance_hint(t: &str) -> String {
    let chars: Vec<char> = t.chars().collect();
    let mut i = 0;
    let mut depth: i64 = 0;
    let mut in_quote = false;
    while i < chars.len() {
        let c = chars[i];
        if in_quote {
            if c == '\\' {
                i += 2;
            } else {
                if c == '"' {
                    in_quote = false;
                }
                i += 1;
            }
        } else if c == '"' {
            in_quote = true;
            i += 1;
        } else if c == ';' {
            while i < chars.len() && chars[i] != '\r' && chars[i] != '\n' {
                i += 1;
            }
        } else {
            if c == '(' {
                depth += 1;
            }
            if c == ')' {
                depth -= 1;
            }
            i += 1;
        }
    }
    if depth > 0 {
        format!(
            "this row opens {depth} form{} it never closes",
            if depth == 1 { "" } else { "s" }
        )
    } else if depth < 0 {
        format!(
            "this row closes {} form{} it never opened",
            -depth,
            if depth == -1 { "" } else { "s" }
        )
    } else {
        String::new()
    }
}

impl<'g> Parser<'g> {
    /// VBA `ParseStmt`: one statement from pos, as VLA text, indented `ind`
    /// levels.
    pub fn parse_stmt(&mut self, pos: &mut usize, ind: usize) -> Result<String, Refusal> {
        let pad = " ".repeat(ind * 2);
        // LX.14: where this sentence began, for a function phrase that stops
        // partway. A comma body keeps its sentence's start.
        if *pos >= 1 && matches!(self.tok(*pos - 1), "" | "." | ":" | PARA_TOK) {
            self.sentence_start = *pos;
        }
        if self.tok(*pos) == PARA_TOK {
            return Err(raise("english-expected-sentence-blank-line", &[]));
        }
        // L0: a token beginning with "(" is a raw VLA form the tokenizer
        // captured whole. It is a statement wherever a sentence can stand;
        // Check validates it by probe-transpiling.
        if self.tok(*pos).starts_with('(') {
            self.claim("a raw VLA form");
            let raw_f = self.tok(*pos).to_string();
            self.validate_raw_vla(&raw_f, *pos)?;
            *pos += 1;
            return Ok(format!("{pad}{raw_f}"));
        }

        let w = self.peek_word(*pos);
        match w.as_str() {
            "if" => {
                if self.at_fail_intro(*pos) {
                    self.claim("the Try recovery intro ('If that fails:')");
                    let context = self.sentence_context(*pos);
                    let loc = self.line_tag(*pos);
                    return Err(raise(
                        "english-if-fails-misplaced",
                        &[("context", &context), ("loc", &loc)],
                    ));
                }
                self.claim("the If form");
                *pos += 1;
                let mut c = self.parse_cond_req(pos)?;
                if self.tok(*pos) == "," {
                    *pos += 1;
                    let inner = self.parse_stmt(pos, 0)?;
                    return Ok(format!("{pad}(if {c} (then {inner}))"));
                }
                self.expect_tok(pos, ":", "',' or ':' after the If condition")?;
                let (then_c, mut cont_w) = self.parse_branch_block(pos, " otherwise ")?;
                let mut r = format!("{pad}(if {c}\r\n{pad}  (then\r\n{})", join_stmts(&then_c));
                while cont_w == "otherwise" {
                    *pos += 1; // the word itself
                    if self.tok(*pos) == "," {
                        *pos += 1;
                    }
                    if self.tok(*pos) == "if" {
                        *pos += 1;
                        c = self.parse_cond_req(pos)?;
                        let inner;
                        if self.tok(*pos) == "," {
                            *pos += 1;
                            inner = self.parse_tracked(pos, 2)?;
                            cont_w = self.continuation_ahead(pos, " otherwise ");
                        } else {
                            self.expect_tok(
                                pos,
                                ":",
                                "',' or ':' after the Otherwise-if condition",
                            )?;
                            let (tc, cw) = self.parse_branch_block(pos, " otherwise ")?;
                            inner = join_stmts(&tc);
                            cont_w = cw;
                        }
                        r.push_str(&format!("\r\n{pad}  (elseif {c}\r\n{inner})"));
                    } else {
                        self.expect_tok(
                            pos,
                            ":",
                            "':' (or ', if <condition>:') after 'Otherwise'",
                        )?;
                        let else_c = self.parse_block(pos)?;
                        r.push_str(&format!("\r\n{pad}  (else\r\n{})", join_stmts(&else_c)));
                        cont_w.clear();
                    }
                }
                r.push(')');
                return Ok(r);
            }
            "repeat" => {
                self.claim("the Repeat loop");
                *pos += 1;
                if self.tok(*pos) == "until" {
                    *pos += 1;
                    let c = self.parse_cond_req(pos)?;
                    if self.tok(*pos) == "," {
                        *pos += 1;
                        let inner = self.in_loop("do", |s| s.parse_stmt(pos, 0))?;
                        return Ok(format!("{pad}(do-until {c} {inner})"));
                    }
                    self.expect_tok(pos, ":", "',' or ':' after 'Repeat until ...'")?;
                    let body = self.in_loop("do", |s| s.parse_block(pos))?;
                    return Ok(format!("{pad}(do-until {c}\r\n{})", join_stmts(&body)));
                }
                let e = self.parse_expr_req(pos)?;
                self.expect_word_is(pos, "times")?;
                self.mark_assigned("counter")?;
                if self.tok(*pos) == "," {
                    *pos += 1;
                    let inner = self.in_loop("for", |s| s.parse_stmt(pos, 0))?;
                    return Ok(format!("{pad}(dotimes counter {e} {inner})"));
                }
                self.expect_tok(pos, ":", "',' or ':' after 'Repeat ... times'")?;
                let body = self.in_loop("for", |s| s.parse_block(pos))?;
                return Ok(format!(
                    "{pad}(dotimes counter {e}\r\n{})",
                    join_stmts(&body)
                ));
            }
            "count" => {
                self.claim("the Count loop");
                *pos += 1;
                let v = self.expect_word(pos, "a name after 'Count'")?;
                self.mark_assigned(&v)?;
                let mut cnt_down = false;
                if self.tok(*pos) == "down" {
                    cnt_down = true;
                    *pos += 1;
                }
                self.expect_word_is(pos, "from")?;
                let e = self.parse_expr_req(pos)?;
                self.expect_word_is(pos, "to")?;
                let c = self.parse_expr_req(pos)?;
                let kind = if self.tok(*pos) == "step" {
                    *pos += 1;
                    let cnt_step = self.parse_expr_req(pos)?;
                    if cnt_down {
                        format!(" (- 0 {cnt_step})")
                    } else {
                        format!(" {cnt_step}")
                    }
                } else if cnt_down {
                    " -1".to_string()
                } else {
                    String::new()
                };
                if self.tok(*pos) == "," {
                    *pos += 1;
                    let inner = self.in_loop("for", |s| s.parse_stmt(pos, 0))?;
                    return Ok(format!("{pad}(for ({v} {e} {c}{kind}) {inner})"));
                }
                self.expect_tok(pos, ":", "',' or ':' after 'Count ... from ... to ...'")?;
                let body = self.in_loop("for", |s| s.parse_block(pos))?;
                return Ok(format!(
                    "{pad}(for ({v} {e} {c}{kind})\r\n{})",
                    join_stmts(&body)
                ));
            }
            "stop" => {
                if self.tok(*pos + 1) == "loop" {
                    self.claim("the Stop form");
                    if self.current_loop().is_empty() {
                        let context = self.sentence_context(*pos);
                        return Err(raise(
                            "english-stop-loop-outside-loop",
                            &[("context", &context)],
                        ));
                    }
                    *pos += 2;
                    self.expect_tok(pos, ".", "'.' after 'Stop the loop'")?;
                    let exit = if self.current_loop() == "do" {
                        "(exit-do)"
                    } else {
                        "(exit-for)"
                    };
                    return Ok(format!("{pad}{exit}"));
                } else if self.tok(*pos + 1) == "." {
                    // B4: inside a value-returning action this must be
                    // Exit Function - Exit Sub there is a compile error.
                    self.claim("the Stop form");
                    *pos += 2;
                    let exit = if self.in_func_def {
                        "(exit-function)"
                    } else {
                        "(exit-sub)"
                    };
                    return Ok(format!("{pad}{exit}"));
                }
                // anything else starting with "stop" tries the phrase rules
            }
            "while" => {
                self.claim("the While loop");
                *pos += 1;
                let c = self.parse_cond_req(pos)?;
                if self.tok(*pos) == "," {
                    *pos += 1;
                    let inner = self.in_loop("do", |s| s.parse_stmt(pos, 0))?;
                    return Ok(format!("{pad}(while {c} {inner})"));
                }
                self.expect_tok(pos, ":", "',' or ':' after the While condition")?;
                let body = self.in_loop("do", |s| s.parse_block(pos))?;
                return Ok(format!("{pad}(while {c}\r\n{})", join_stmts(&body)));
            }
            "for" => {
                self.claim("the For each loop");
                *pos += 1;
                self.expect_word_is(pos, "each")?;
                let v = self.expect_word(pos, "a name after 'For each'")?;
                self.mark_assigned(&v)?;
                self.expect_word_is(pos, "in")?;
                let e = if self.is_dict_name(self.tok(*pos))
                    && (self.tok(*pos + 1) == "," || self.tok(*pos + 1) == ":")
                {
                    let d = format!("(vladictpairs {})", self.tok(*pos));
                    *pos += 1;
                    d
                } else {
                    self.parse_expr_req(pos)?
                };
                if self.tok(*pos) == "," {
                    *pos += 1;
                    let inner = self.in_loop("for", |s| s.parse_stmt(pos, 0))?;
                    return Ok(format!("{pad}(for-each ({v} {e}) {inner})"));
                }
                self.expect_tok(pos, ":", "',' or ':' after 'For each ... in ...'")?;
                let body = self.in_loop("for", |s| s.parse_block(pos))?;
                return Ok(format!(
                    "{pad}(for-each ({v} {e})\r\n{})",
                    join_stmts(&body)
                ));
            }
            "get" => {
                if self.is_using_fn(self.tok(*pos + 1)) && self.tok(*pos + 2) == "using" {
                    self.claim("the standalone-Get guard");
                    let f = self.tok(*pos + 1).to_string();
                    let loc = self.line_tag(*pos);
                    return Err(raise(
                        "english-standalone-get",
                        &[("fn", &f), ("loc", &loc)],
                    ));
                }
            }
            "give" => {
                self.claim("the Give back form");
                *pos += 1;
                self.expect_word_is(pos, "back")?;
                if !self.in_func_def {
                    let context = self.sentence_context(*pos - 2);
                    let loc = self.line_tag(*pos - 2);
                    return Err(raise(
                        "english-give-back-outside-action",
                        &[("context", &context), ("loc", &loc)],
                    ));
                }
                let e = self.parse_expr_req(pos)?;
                self.expect_tok(pos, ".", "'.' at the end of the sentence")?;
                return Ok(format!("{pad}(return {e})"));
            }
            "try" => {
                self.claim("the Try block");
                *pos += 1;
                self.expect_tok(pos, ":", "':' after 'Try'")?;
                self.try_count += 1;
                let tn = self.try_count;
                let (body, has_rec) = self.parse_try_body(pos)?;
                let else_c = if has_rec {
                    self.in_recovery = true;
                    let r = self.parse_block(pos);
                    self.in_recovery = false;
                    r?
                } else {
                    Vec::new()
                };
                let mut r = format!("{pad}(on-error goto vla-tryf-{tn})\r\n");
                if !body.is_empty() {
                    r.push_str(&join_stmts(&body));
                    r.push_str("\r\n");
                }
                r.push_str(&format!("{pad}(goto vla-tryd-{tn})\r\n"));
                r.push_str(&format!("{pad}(label vla-tryf-{tn})\r\n"));
                r.push_str(&format!("{pad}(set! vla-problem err.description)\r\n"));
                r.push_str(&format!("{pad}(resume vla-tryr-{tn})\r\n"));
                r.push_str(&format!("{pad}(label vla-tryr-{tn})\r\n"));
                r.push_str(&format!("{pad}{}\r\n", self.restore_handler_vla()));
                if !else_c.is_empty() {
                    r.push_str(&join_stmts(&else_c));
                    r.push_str("\r\n");
                }
                r.push_str(&format!("{pad}(label vla-tryd-{tn})\r\n"));
                r.push_str(&format!("{pad}{}", self.restore_handler_vla()));
                return Ok(r);
            }
            "when" => {
                self.claim("the When choices form");
                *pos += 1;
                if self.tok(*pos) == "it" {
                    let context = self.sentence_context(*pos - 1);
                    let loc = self.line_tag(*pos - 1);
                    return Err(raise(
                        "english-when-it-is-first",
                        &[("context", &context), ("loc", &loc)],
                    ));
                }
                if !is_str_tok(self.tok(*pos))
                    && self.tok(*pos + 1) == "is"
                    && self.tok(*pos + 2) == "clicked"
                    && self.tok(*pos + 3) == ":"
                {
                    let t = self.tok(*pos).to_string();
                    let loc = self.line_tag(*pos);
                    return Err(raise(
                        "english-click-handler-needs-quotes",
                        &[("tok", &t), ("loc", &loc)],
                    ));
                }
                if is_str_tok(self.tok(*pos))
                    && self.tok(*pos + 1) == "is"
                    && is_word_tok(self.tok(*pos + 2))
                    && self.tok(*pos + 2) != "clicked"
                    && self.tok(*pos + 3) == ":"
                {
                    let t = render_tok(self.tok(*pos));
                    let misspelled = self.tok(*pos + 2).to_string();
                    let loc = self.line_tag(*pos);
                    return Err(raise(
                        "english-clicked-misspelled",
                        &[("tok", &t), ("misspelled", &misspelled), ("loc", &loc)],
                    ));
                }
                let e = self.parse_expr_req(pos)?;
                self.expect_word_is(pos, "is")?;
                let mut c = self.parse_case_values(pos)?;
                self.expect_tok(pos, ":", "':' after 'When ... is ...'")?;
                let (body, mut cont_w) = self.parse_branch_block(pos, " when otherwise ")?;
                let mut r = format!(
                    "{pad}(select {e}\r\n{pad}  (case {c}\r\n{})",
                    join_stmts(&body)
                );
                while !cont_w.is_empty() {
                    *pos += 1; // "when" or "otherwise"
                    if cont_w == "when" {
                        self.expect_word_is(pos, "it")?;
                        self.expect_word_is(pos, "is")?;
                        c = self.parse_case_values(pos)?;
                        self.expect_tok(pos, ":", "':' after 'When it is ...'")?;
                        let (b, cw) = self.parse_branch_block(pos, " when otherwise ")?;
                        cont_w = cw;
                        r.push_str(&format!("\r\n{pad}  (case {c}\r\n{})", join_stmts(&b)));
                    } else {
                        self.expect_tok(pos, ":", "':' after 'Otherwise'")?;
                        let b = self.parse_block(pos)?;
                        r.push_str(&format!("\r\n{pad}  (case-else\r\n{})", join_stmts(&b)));
                        cont_w.clear();
                    }
                }
                r.push(')');
                return Ok(r);
            }
            "create" => {
                self.claim("the Create declaration");
                *pos += 1;
                self.skip_articles(pos);
                let kind = self.expect_word(
                    pos,
                    "'number', 'text', 'value', 'list', or 'lookup' after 'Create'",
                )?;
                if kind == "button" {
                    let loc = self.line_tag(*pos - 1);
                    return Err(raise(
                        "english-button-not-created-with-create",
                        &[("loc", &loc)],
                    ));
                }
                if kind == "pivot" {
                    let loc = self.line_tag(*pos - 1);
                    return Err(raise(
                        "english-pivot-not-created-with-create",
                        &[("loc", &loc)],
                    ));
                }
                self.expect_word_is(pos, "called")?;
                let v = self.expect_word(pos, "a name after 'called'")?;
                self.expect_tok(pos, ".", "'.' at the end of the sentence")?;
                self.mark_declared(&v)?;
                return Ok(match kind.as_str() {
                    "number" => format!("{pad}(dim {v} Double)"),
                    "text" => format!("{pad}(dim {v} String)"),
                    "value" => format!("{pad}(dim {v})"),
                    "list" => {
                        format!("{pad}(begin (dim {v} Collection) (obj-set! {v} (new Collection)))")
                    }
                    "lookup" => {
                        add_keyed(&mut self.dict_names, &v);
                        format!("{pad}(begin (dim {v} Object) (obj-set! {v} (vladictnew)))")
                    }
                    _ => {
                        return Err(raise("english-create-unknown-kind", &[("kind", &kind)]));
                    }
                });
            }
            "to" => {
                self.claim("the To definition");
                return Err(raise("english-to-not-top-level", &[]));
            }
            "define" => {
                self.claim("the Define declaration");
                return Err(raise("english-define-not-top-level", &[]));
            }
            "increase" | "decrease" | "add" => {
                if let Some(r) = self.percent_share_form(&w, pos, &pad)? {
                    return Ok(r);
                }
            }
            "done" | "otherwise" => {
                self.claim("the block structure (Done/Otherwise placement)");
                return Err(raise("english-unexpected-token-block", &[("tok", &w)]));
            }
            _ => {}
        }

        // Phrase rules (DCG productions), in order; first match wins.
        if let Some(txt) = self.try_rules(pos)? {
            return Ok(format!("{pad}{txt}"));
        }

        // Fallback: a name is a call to a defined action - bare ("Greet.")
        // or with named arguments ("Stamp with row of 2 and value of "x".").
        // Every call is recorded and checked against the definitions once
        // the whole file has parsed.
        if !w.is_empty() {
            if self.tok(*pos + 1) == "." {
                self.claim(&format!("a call to the action '{w}'"));
                let text = self.render_sentence_at(*pos);
                let line = self.toks.line_at(*pos);
                self.record_call(&w, Vec::new(), text, line);
                *pos += 2;
                return Ok(format!("{pad}({w})"));
            } else if self.tok(*pos + 1) == "with" {
                self.claim(&format!("a call to the action '{w}' with arguments"));
                let call_txt = self.render_sentence_at(*pos);
                let call_line = self.toks.line_at(*pos);
                *pos += 2;
                let mut inner = String::new();
                let mut call_args: Vec<String> = Vec::new();
                loop {
                    let v = self.expect_word(pos, "a parameter name after 'with'")?;
                    self.expect_word_is(pos, "of")?;
                    let e = self.parse_expr_req(pos)?;
                    inner.push_str(&format!(" :{v} {e}"));
                    call_args.push(fold(&v));
                    if self.tok(*pos) == "and" {
                        *pos += 1;
                    } else {
                        break;
                    }
                }
                self.expect_tok(pos, ".", "'.' at the end of the sentence")?;
                self.record_call(&w, call_args, call_txt, call_line);
                return Ok(format!("{pad}({w}{inner})"));
            }
        }

        let msg = self.build_parse_error(*pos);
        Err(raise("english-parse-error", &[("msg", &msg)]))
    }

    /// `PushLoop` / `PopLoop` around a loop's body.
    fn in_loop<T>(
        &mut self,
        kind: &str,
        body: impl FnOnce(&mut Self) -> Result<T, Refusal>,
    ) -> Result<T, Refusal> {
        self.loop_stack.push(kind.to_string());
        let r = body(self);
        self.loop_stack.pop();
        r
    }

    /// VBA `MarkDeclared`: a declared name, checked.
    pub(super) fn mark_declared(&mut self, name: &str) -> Result<(), Refusal> {
        self.check_name(name)?;
        add_keyed(&mut self.declared, name);
        Ok(())
    }

    /// B7.4: "Increase total by 10%." and "Add 10% to total." mean Grow and
    /// Shrink; any other amount carrying % or percent refuses with
    /// directions; an amount without % falls through to the rules.
    fn percent_share_form(
        &mut self,
        w: &str,
        pos: &mut usize,
        pad: &str,
    ) -> Result<Option<String>, Refusal> {
        let p = *pos;
        let is_pct = |t: &str| t == "%" || t == "percent";
        let sh_stop = if w == "add" {
            // Add <N><%|percent> to <name>.
            let sh_n = self.tok(p + 1).to_string();
            if is_pct(self.tok(p + 2))
                && self.tok(p + 3) == "to"
                && !self.word_at(p + 4).is_empty()
                && self.tok(p + 5) == "."
                && (is_num_tok(&sh_n) || !self.word_at(p + 1).is_empty())
            {
                let tgt = self.word_at(p + 4);
                self.claim("the percent form of Increase/Decrease/Add");
                self.mark_assigned(&tgt)?;
                *pos = p + 6;
                return Ok(Some(format!(
                    "{pad}(set! {tgt} (* {tgt} (+ 1 (/ {sh_n} 100))))"
                )));
            }
            "to"
        } else {
            // Increase/Decrease <name> by <N><%|percent>.
            let tgt = self.word_at(p + 1);
            let sh_n = self.tok(p + 3).to_string();
            if !tgt.is_empty()
                && self.tok(p + 2) == "by"
                && is_pct(self.tok(p + 4))
                && self.tok(p + 5) == "."
                && (is_num_tok(&sh_n) || !self.word_at(p + 3).is_empty())
            {
                let op = if w == "increase" { "+" } else { "-" };
                self.claim("the percent form of Increase/Decrease/Add");
                self.mark_assigned(&tgt)?;
                *pos = p + 6;
                return Ok(Some(format!(
                    "{pad}(set! {tgt} (* {tgt} ({op} 1 (/ {sh_n} 100))))"
                )));
            }
            "."
        };
        // Not the simple share: % or percent anywhere in the AMOUNT refuses
        // rather than let either reading win silently.
        let mut scan = if w == "add" {
            p + 1
        } else if self.tok(p + 2) == "by" {
            p + 3
        } else {
            p + 1
        };
        let mut hit = false;
        while scan <= self.toks.count() {
            let t = self.tok(scan);
            if t == "." || t == sh_stop || t == PARA_TOK {
                break;
            }
            if is_pct(t) {
                hit = true;
                break;
            }
            scan += 1;
        }
        if hit {
            self.claim("the percent guard on Increase/Decrease/Add");
            let suggestion = if w == "add" {
                "'Add step to <name>.'".to_string()
            } else {
                format!("'{} <name> by step.'", proper_case(w))
            };
            let sentence = self.render_sentence_at(p);
            let loc = self.line_tag(p);
            return Err(raise(
                "english-percent-mixed-amount",
                &[
                    ("sentence", &sentence),
                    ("suggestion", &suggestion),
                    ("loc", &loc),
                ],
            ));
        }
        Ok(None)
    }

    // ---- blocks -------------------------------------------------------------------

    /// VBA `ParseBlock`: statements until the block closes: "Done." (consumed),
    /// a paragraph marker (left for the outer blocks), or the end.
    pub(super) fn parse_block(&mut self, pos: &mut usize) -> Result<Vec<String>, Refusal> {
        let mut out = Vec::new();
        loop {
            if *pos > self.toks.count() || self.tok(*pos) == PARA_TOK {
                break;
            }
            if self.tok(*pos) == "done" {
                *pos += 1;
                self.expect_tok(pos, ".", "'.' after 'Done'")?;
                break;
            }
            let s = self.parse_tracked(pos, 2)?;
            out.push(s);
        }
        Ok(out)
    }

    /// VBA `ParseBranchBlock`: a block that a continuation word may also
    /// close, returned with that word ("" when none).
    pub(super) fn parse_branch_block(
        &mut self,
        pos: &mut usize,
        cont_words: &str,
    ) -> Result<(Vec<String>, String), Refusal> {
        let mut out = Vec::new();
        let mut cont_word = String::new();
        loop {
            if *pos > self.toks.count() {
                break;
            }
            if self.tok(*pos) == PARA_TOK {
                cont_word = self.continuation_ahead(pos, cont_words);
                break;
            }
            if cont_words.contains(&format!(" {} ", self.tok(*pos))) {
                cont_word = self.tok(*pos).to_string();
                break;
            }
            if self.tok(*pos) == "done" {
                *pos += 1;
                self.expect_tok(pos, ".", "'.' after 'Done'")?;
                break;
            }
            let s = self.parse_tracked(pos, 2)?;
            out.push(s);
        }
        Ok((out, cont_word))
    }

    /// VBA `ContinuationAhead`: a continuation word past blank lines, pos
    /// moved to it when found.
    pub(super) fn continuation_ahead(&self, pos: &mut usize, cont_words: &str) -> String {
        let mut j = *pos;
        while self.tok(j) == PARA_TOK {
            j += 1;
        }
        let t = self.tok(j);
        if !t.is_empty() && cont_words.contains(&format!(" {t} ")) {
            *pos = j;
            return t.to_string();
        }
        String::new()
    }

    /// VBA `ParseTryBody`: the body until "If that fails:" (past a blank
    /// line too), "Done." or the end; the flag says a recovery follows.
    fn parse_try_body(&mut self, pos: &mut usize) -> Result<(Vec<String>, bool), Refusal> {
        let mut out = Vec::new();
        loop {
            if *pos > self.toks.count() {
                break;
            }
            if self.tok(*pos) == PARA_TOK {
                let mut j = *pos;
                while self.tok(j) == PARA_TOK {
                    j += 1;
                }
                if self.at_fail_intro(j) {
                    *pos = j; // blank line separates body and recovery
                } else {
                    break; // leave marker for outer blocks
                }
            }
            if self.at_fail_intro(*pos) {
                *pos += 3; // "if" "that" "fails"
                self.expect_tok(pos, ":", "':' after 'If that fails'")?;
                return Ok((out, true));
            }
            if self.tok(*pos) == "done" {
                *pos += 1;
                self.expect_tok(pos, ".", "'.' after 'Done'")?;
                break;
            }
            let s = self.parse_tracked(pos, 2)?;
            out.push(s);
        }
        Ok((out, false))
    }

    pub(super) fn at_fail_intro(&self, p: usize) -> bool {
        self.tok(p) == "if" && self.tok(p + 1) == "that" && self.tok(p + 2) == "fails"
    }

    pub(super) fn at_sheet_change_event(&self, p: usize) -> bool {
        self.tok(p) == "when"
            && self.tok(p + 1) == "sheet"
            && self.tok(p + 2) == "changes"
            && self.tok(p + 3) == ":"
    }

    pub(super) fn at_button_click_event(&self, p: usize) -> bool {
        self.tok(p) == "when"
            && is_str_tok(self.tok(p + 1))
            && self.tok(p + 2) == "is"
            && self.tok(p + 3) == "clicked"
            && self.tok(p + 4) == ":"
    }

    fn restore_handler_vla(&self) -> &'static str {
        if self.step_tracking {
            "(on-error goto vla-fail)"
        } else {
            "(on-error goto 0)"
        }
    }

    /// VBA `ParseCaseValues`: `(v1 v2 ...)`, "or"-separated.
    fn parse_case_values(&mut self, pos: &mut usize) -> Result<String, Refusal> {
        let mut r = format!("({}", self.parse_expr_req(pos)?);
        while self.tok(*pos) == "or" {
            *pos += 1;
            r.push(' ');
            r.push_str(&self.parse_expr_req(pos)?);
        }
        r.push(')');
        Ok(r)
    }

    /// VBA `ValidateRawVla` (L0, L12): a raw form's head may not define
    /// anything at the top level, its parentheses must balance, and it must
    /// transpile inside a probe sub with the prelude and the carried macros.
    fn validate_raw_vla(&mut self, form_text: &str, pos: usize) -> Result<(), Refusal> {
        let head: String = form_text
            .chars()
            .skip(1)
            .take_while(|c| !matches!(c, ' ' | '\t' | '(' | ')'))
            .collect();
        if matches!(
            fold(&head).as_str(),
            "sub" | "function" | "defmacro" | "type" | "enum" | "public" | "private"
        ) {
            let loc = self.line_tag(pos);
            return Err(raise(
                "english-top-level-definition-in-row",
                &[("head", &head), ("loc", &loc)],
            ));
        }
        let bal_hint = vla_balance_hint(form_text);
        if !bal_hint.is_empty() {
            let loc = self.line_tag(pos);
            return Err(raise(
                "english-form-balance-hint",
                &[("hint", &bal_hint), ("loc", &loc)],
            ));
        }
        let prelude = format!("{}\n{}", self.g.prelude, self.g.vocab_macros_text());
        if let Err(e) =
            crate::emit::compile(&format!("(sub vla-check-probe () {form_text})"), &prelude)
        {
            let loc = self.line_tag(pos);
            return Err(raise(
                "english-form-doesnt-transpile",
                &[("detail", &e.text), ("loc", &loc)],
            ));
        }
        Ok(())
    }

    // ---- the step table -----------------------------------------------------------

    /// VBA `SrcLineText` (S3.2): the original line, trimmed, or the sentence
    /// rendered from its tokens when the line is empty.
    fn src_line_text(&self, ln: u32, pos: usize) -> String {
        if ln >= 1 {
            if let Some(s) = self.src_lines.get((ln - 1) as usize) {
                let s = s.trim_matches(' ');
                if !s.is_empty() {
                    return s.to_string();
                }
            }
        }
        self.render_sentence_at(pos)
    }

    /// VBA `ParseTracked`: one statement with its step number set before it
    /// and its line marked, when step tracking is on; the line mark alone
    /// under TER-10's `mLineMarks`.
    pub(super) fn parse_tracked(&mut self, pos: &mut usize, ind: usize) -> Result<String, Refusal> {
        self.cur_line = self.toks.line_at(*pos);
        let start_ln = self.cur_line;
        if !self.step_tracking {
            let lm_ln = self.toks.line_at(*pos);
            let lm_stmt = self.parse_stmt(pos, ind)?;
            self.refuse_cell_shaped_read(&lm_stmt, start_ln)?;
            if self.line_marks && lm_ln > 0 {
                return Ok(format!(
                    "{}(at-line {lm_ln}\r\n{lm_stmt})",
                    " ".repeat(ind * 2)
                ));
            }
            return Ok(lm_stmt);
        }
        self.step_count += 1;
        let my_n = self.step_count;
        let stp_ln = self.toks.line_at(*pos);
        if stp_ln > 0 {
            let text = self.src_line_text(stp_ln, *pos);
            self.step_texts.push(format!("{text} [line {stp_ln}]"));
        } else {
            let text = self.render_sentence_at(*pos);
            self.step_texts.push(text);
        }
        let stmt = self.parse_stmt(pos, ind)?;
        self.refuse_cell_shaped_read(&stmt, start_ln)?;
        let pad = " ".repeat(ind * 2);
        if stp_ln > 0 {
            Ok(format!(
                "{pad}(set! vla-step {my_n})\r\n{pad}(if (vlatraceon) (then (vlatracestep {my_n} (vla-step-text {my_n}))))\r\n{pad}(at-line {stp_ln}\r\n{stmt})"
            ))
        } else {
            Ok(format!(
                "{pad}(set! vla-step {my_n})\r\n{pad}(if (vlatraceon) (then (vlatracestep {my_n} (vla-step-text {my_n}))))\r\n{stmt}"
            ))
        }
    }

    // ---- the proof runners --------------------------------------------------------

    /// The translation both proof runners make: `EnTokenize`, the alias
    /// rewrite, one `ParseStmt` at position 1, LX.13's read check, and
    /// nothing but paragraph marks allowed after the statement.
    pub fn translate_proof_sentence(&mut self, sentence: &str) -> Result<String, Refusal> {
        let toks = tokenize(sentence).map_err(|e| e.refusal)?;
        self.set_tokens(toks);
        self.cur_line = 0; // LX.13: a proof has no program line
        self.assigned.clear();
        self.declared.clear();
        self.loop_stack.clear();
        let mut pos = 1;
        let got = self.parse_stmt(&mut pos, 0)?;
        self.refuse_cell_shaped_read(&got, 0)?;
        while self.tok(pos) == PARA_TOK {
            pos += 1;
        }
        if pos <= self.toks.count() {
            return Err(raise("english-extra-words-after-statement", &[]));
        }
        Ok(got)
    }

    /// VBA `RunVocabTest`: the sentence translates, and reads as the
    /// expected VLA with whitespace normalized.
    pub fn run_vocab_test(&mut self, proof: &Proof) -> Result<(), Refusal> {
        let loc = prov_loc(&proof.source, proof.line, &proof.row_tag);
        let got = match self.translate_proof_sentence(&proof.sentence) {
            Ok(got) => got,
            Err(d) => {
                return Err(raise(
                    "english-test-failed-to-translate",
                    &[
                        ("loc", &loc),
                        ("sentence", &proof.sentence),
                        ("error", &d.text),
                    ],
                ))
            }
        };
        if normalize_ws(&got) != normalize_ws(&proof.expected) {
            return Err(raise(
                "english-test-failed",
                &[
                    ("loc", &loc),
                    ("sentence", &proof.sentence),
                    ("expected", &normalize_ws(&proof.expected)),
                    ("got", &normalize_ws(&got)),
                ],
            ));
        }
        Ok(())
    }

    /// VBA `RunVocabFailTest` (G4): the sentence must refuse, and the
    /// refusal must contain the fragment, whitespace-normalized and
    /// case-blind.
    pub fn run_vocab_fail_test(&mut self, proof: &Proof) -> Result<(), Refusal> {
        let loc = prov_loc(&proof.source, proof.line, &proof.row_tag);
        match self.translate_proof_sentence(&proof.sentence) {
            Ok(got) => Err(raise(
                "english-failtest-translated",
                &[
                    ("loc", &loc),
                    ("sentence", &proof.sentence),
                    ("became", &normalize_ws(&got)),
                ],
            )),
            Err(d) => {
                let message = fold(&normalize_ws(&d.text));
                let wanted = fold(&normalize_ws(&proof.expected));
                if !message.contains(&wanted) {
                    return Err(raise(
                        "english-failtest-message-drifted",
                        &[
                            ("loc", &loc),
                            ("sentence", &proof.sentence),
                            ("wanted", &proof.expected),
                            ("message", &d.text),
                        ],
                    ));
                }
                Ok(())
            }
        }
    }
}

impl Grammar {
    /// `EnglishLoadVocabularyText`'s second loop: the file's proofs run
    /// against the fully loaded grammar, in file order, positives and
    /// negatives alike; each is counted before it runs. Under
    /// [`ProofMode::Raise`] the first failure refuses the load, as the
    /// reference does; under [`ProofMode::Collect`] every failure is kept
    /// and the load goes on, which is what `frazaro prove` scores.
    pub(super) fn run_file_proofs(&mut self, proofs: Vec<Proof>) -> Result<(), Refusal> {
        let mut tests = 0u64;
        let mut fails = 0u64;
        let mut failures: Vec<ProofFailure> = Vec::new();
        let mut first_err: Option<Refusal> = None;
        {
            let mut parser = Parser::new(self);
            for proof in &proofs {
                let r = match proof.kind {
                    ProofKind::Fail => {
                        fails += 1;
                        parser.run_vocab_fail_test(proof)
                    }
                    ProofKind::Success => {
                        tests += 1;
                        parser.run_vocab_test(proof)
                    }
                };
                if let Err(e) = r {
                    match self.proof_mode {
                        ProofMode::Raise => {
                            first_err = Some(e);
                            break;
                        }
                        ProofMode::Collect => failures.push(ProofFailure {
                            proof: proof.clone(),
                            refusal: e,
                        }),
                    }
                }
            }
        }
        self.tests_run += tests;
        self.fails_run += fails;
        self.proofs.extend(proofs);
        self.proof_failures.extend(failures);
        match first_err {
            Some(e) => Err(e),
            None => Ok(()),
        }
    }

    /// One sentence through the grammar as a proof's sentence goes: the
    /// VLA it translates to, or the refusal.
    pub fn translate_sentence(&self, sentence: &str) -> Result<String, Refusal> {
        let mut parser = Parser::new(self);
        parser.translate_proof_sentence(sentence)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../../scripts/polyglotta/english.vla");
    const DANSK: &str = include_str!("../../../scripts/polyglotta/dansk.vla");
    const DEUTSCHE: &str = include_str!("../../../scripts/polyglotta/deutsche.vla");
    const ESPANOL: &str = include_str!("../../../scripts/polyglotta/espanol.vla");
    const ESPERANTO: &str = include_str!("../../../scripts/polyglotta/esperanto.vla");
    const FRANCAIS: &str = include_str!("../../../scripts/polyglotta/francais.vla");
    const LATIN: &str = include_str!("../../../scripts/polyglotta/latin.vla");
    const PIRATE: &str = include_str!("../../../scripts/polyglotta/pirate.vla");

    fn translate(rules: &str, sentence: &str) -> Result<String, Refusal> {
        let mut g = Grammar::new(PRELUDE);
        g.load_vocabulary_text(rules, "t.vla").unwrap();
        g.translate_sentence(sentence)
    }

    const SHOW: &str = "(x-vla \"peek {e:expr}\" (debug-print {e}))";

    #[test]
    fn expressions_read_as_the_reference_reads_them() {
        let t = |s: &str| translate(SHOW, s).unwrap();
        assert_eq!(t("Peek 2 plus 3 times 4."), "(debug-print (+ 2 (* 3 4)))");
        assert_eq!(
            t("Peek total divided by 2 minus 1."),
            "(debug-print (- (/ total 2) 1))"
        );
        assert_eq!(t("Peek 15%."), "(debug-print (/ 15 100))");
        assert_eq!(t("Peek 10 percent."), "(debug-print (/ 10 100))");
        assert_eq!(
            t("Peek \"a\" joined with \"b\" followed by c."),
            "(debug-print (& (& \"a\" \"b\") c))"
        );
        assert_eq!(t("Peek cell B2."), "(debug-print (range \"b2\"))");
        assert_eq!(
            t("Peek cell in column C row 5."),
            "(debug-print (cells 5 \"c\"))"
        );
        assert_eq!(
            t("Peek value in column number 2 row k."),
            "(debug-print (cells k 2))"
        );
        assert_eq!(t("Peek value in cell D4."), "(debug-print (range \"d4\"))");
        assert_eq!(
            t("Peek item 2 of found-items."),
            "(debug-print (vlaitem found-items 2))"
        );
        assert_eq!(t("Peek the third."), "(debug-print 3)");
        assert_eq!(
            t("Peek length of cell B2."),
            "(debug-print (len (range \"b2\")))"
        );
        assert_eq!(t("Peek today."), "(debug-print (date))");
        assert_eq!(t("Peek length."), "(debug-print length)");
        assert_eq!(t("Peek cell."), "(debug-print cell)");
        assert_eq!(t("Peek item."), "(debug-print item)");
    }

    #[test]
    fn a_bare_range_is_no_value_and_a_cell_shaped_word_is_refused() {
        let e = translate(SHOW, "Peek a1:b2.").unwrap_err();
        assert_eq!(e.id, "english-parse-error");
        assert!(
            e.text
                .contains("I understood 'peek' - then I expected a value (like 5, "),
            "{}",
            e.text
        );
        let e = translate(SHOW, "Peek a2.").unwrap_err();
        assert_eq!(e.id, "english-cell-shaped-name");
        assert!(e.text.starts_with("'a2' is shaped like a cell"));
    }

    #[test]
    fn conditions_read_as_the_reference_reads_them() {
        const IF: &str = "(x-vla \"check {c:cond}\" (if {c} 1))";
        let t = |s: &str| translate(IF, s).unwrap();
        assert_eq!(t("Check x is 5."), "(if (= x 5) 1)");
        assert_eq!(
            t("Check x is at least 5 and y is not empty."),
            "(if (and (>= x 5) (not (blank? y))) 1)"
        );
        assert_eq!(
            t("Check name contains \"z\" or name starts with \"a\"."),
            "(if (or (positive? (instr 1 name \"z\" vbtextcompare)) (= (instr 1 name \"a\" vbtextcompare) 1)) 1)"
        );
        assert_eq!(t("Check n is divisible by 3."), "(if (zero? (mod n 3)) 1)");
        assert_eq!(
            t("Check range A1:D50 contains \"x\"."),
            "(if (positive? (vlafindtext (range \"a1:d50\") \"x\" \"row\")) 1)"
        );
        assert_eq!(
            t("Check column C does not contain \"x\"."),
            "(if (zero? (vlafindtext (columns \"c\") \"x\" \"row\")) 1)"
        );
        assert_eq!(
            t("Check name ends with \"s\"."),
            "(if (= (lcase (right name (len \"s\"))) (lcase \"s\")) 1)"
        );
    }

    #[test]
    fn the_percent_form_and_its_guard() {
        let g = Grammar::new(PRELUDE);
        assert_eq!(
            g.translate_sentence("Increase total by 10%.").unwrap(),
            "(set! total (* total (+ 1 (/ 10 100))))"
        );
        assert_eq!(
            g.translate_sentence("Decrease total by rate percent.")
                .unwrap(),
            "(set! total (* total (- 1 (/ rate 100))))"
        );
        assert_eq!(
            g.translate_sentence("Add 5% to total.").unwrap(),
            "(set! total (* total (+ 1 (/ 5 100))))"
        );
        let e = g.translate_sentence("Add 5 plus 5% to total.").unwrap_err();
        assert_eq!(e.id, "english-percent-mixed-amount");
        assert!(e.text.contains("'Add step to <name>.'"), "{}", e.text);
        let e = g
            .translate_sentence("Increase total by 10% times 2.")
            .unwrap_err();
        assert!(e.text.contains("'Increase <name> by step.'"), "{}", e.text);
    }

    #[test]
    fn the_statement_forms_read_as_the_reference_reads_them() {
        let g = Grammar::new(PRELUDE);
        let t = |s: &str| g.translate_sentence(s).unwrap();
        assert_eq!(t("If x is 5, show x."), "(if (= x 5) (then (msgbox x)))");
        assert_eq!(
            t("Repeat 3 times, log counter."),
            "(dotimes counter 3 (debug-print counter))"
        );
        assert_eq!(
            t("Repeat until n is 0, set n to n minus 1."),
            "(do-until (= n 0) (set! n (- n 1)))"
        );
        assert_eq!(
            t("While n is less than 3, add 1 to n."),
            "(while (< n 3) (add! n 1))"
        );
        assert_eq!(
            t("Count i from 1 to 10, log i."),
            "(for (i 1 10) (debug-print i))"
        );
        assert_eq!(
            t("Count i down from 10 to 1 step 2, log i."),
            "(for (i 10 1 (- 0 2)) (debug-print i))"
        );
        assert_eq!(
            t("Count i down from 3 to 1, log i."),
            "(for (i 3 1 -1) (debug-print i))"
        );
        assert_eq!(
            t("For each f in found-items, log f."),
            "(for-each (f found-items) (debug-print f))"
        );
        assert_eq!(t("Create a number called total."), "(dim total Double)");
        assert_eq!(t("Create a text called label."), "(dim label String)");
        assert_eq!(t("Create a value called thing."), "(dim thing)");
        assert_eq!(
            t("Create a list called found-items."),
            "(begin (dim found-items Collection) (obj-set! found-items (new Collection)))"
        );
        assert_eq!(
            t("Create a lookup called prices."),
            "(begin (dim prices Object) (obj-set! prices (vladictnew)))"
        );
        assert_eq!(t("Stop."), "(exit-sub)");
        let e = g.translate_sentence("Stop the loop.").unwrap_err();
        assert_eq!(e.id, "english-stop-loop-outside-loop");
        let e = g.translate_sentence("Give back 5.").unwrap_err();
        assert_eq!(e.id, "english-give-back-outside-action");
        let e = g.translate_sentence("To greet:").unwrap_err();
        assert_eq!(e.id, "english-to-not-top-level");
        let e = g.translate_sentence("Define x as 5.").unwrap_err();
        assert_eq!(e.id, "english-define-not-top-level");
        let e = g.translate_sentence("If that fails: log 1.").unwrap_err();
        assert_eq!(e.id, "english-if-fails-misplaced");
        let e = g.translate_sentence("When it is 5: log 1.").unwrap_err();
        assert_eq!(e.id, "english-when-it-is-first");
        let e = g
            .translate_sentence("Create a button called go.")
            .unwrap_err();
        assert_eq!(e.id, "english-button-not-created-with-create");
        let e = g
            .translate_sentence("Create a thing called x.")
            .unwrap_err();
        assert_eq!(e.id, "english-create-unknown-kind");
        assert_eq!(
            t("(set! (range \"h25\") \"vla-row\")"),
            "(set! (range \"h25\") \"vla-row\")"
        );
        let e = g.translate_sentence("(sub x () 1)").unwrap_err();
        assert_eq!(e.id, "english-top-level-definition-in-row");
        assert_eq!(
            vla_balance_hint("(a (b)"),
            "this row opens 1 form it never closes"
        );
        assert_eq!(
            vla_balance_hint("(a))"),
            "this row closes 1 form it never opened"
        );
        assert_eq!(vla_balance_hint("(a \"(\" ; (\n)"), "");
    }

    #[test]
    fn the_parse_error_is_the_reference_s() {
        let mut g = Grammar::new(PRELUDE);
        g.load_vocabulary_text(SHOW, "t.vla").unwrap();
        let e = g.translate_sentence("Frobnicate the widget.").unwrap_err();
        assert_eq!(e.id, "english-parse-error");
        assert!(
            e.text.starts_with("Don't understand: 'frobnicate widget.' No loaded sentence starts with 'frobnicate' - press What can I say? on the Frazaro tab to see all 12 sentences this program understands. (line 1)"),
            "{}",
            e.text
        );
        let e = g.translate_sentence("Peek 5 6.").unwrap_err();
        assert!(
            e.text.starts_with("I understood 'peek 5' - then I expected '.' to end the sentence but found '6'. Did you mean: 'peek {e:expr}' (line 1)"),
            "{}",
            e.text
        );
        // A bare name is an action call, with or without arguments.
        assert_eq!(g.translate_sentence("Greet.").unwrap(), "(greet)");
        assert_eq!(
            g.translate_sentence("Stamp with row of 2 and value of \"x\".")
                .unwrap(),
            "(stamp :row 2 :value \"x\")"
        );
        let e = g.translate_sentence("Done.").unwrap_err();
        assert_eq!(e.id, "english-unexpected-token-block");
    }

    #[test]
    fn a_failing_proof_refuses_the_load_as_the_reference_does() {
        let mut g = Grammar::new(PRELUDE);
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-success \"Warm cell B2.\" (debug-print \"b3\"))",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-test-failed");
        assert!(
            e.text.starts_with("t.vla line 2: test FAILED"),
            "{}",
            e.text
        );
        assert!(
            e.text.contains("expected: (debug-print \"b3\")")
                && e.text.contains("(debug-print \"b2\")"),
            "{}",
            e.text
        );
        // A refused load leaves its rules registered, as in Excel, so the
        // next load starts from a reset grammar (EnglishResetGrammar).
        g.reset();
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-fail \"Warm cell B2.\" \"anything\")",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-failtest-translated");
        g.reset();
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-fail \"Warm cell banana.\" \"a unicorn\")",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-failtest-message-drifted");
        g.reset();
        g.load_vocabulary_text(
            "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-success \"Warm cell B2.\" (debug-print \"b2\"))\n(test-fail \"Warm cell banana.\" \"I expected a cell\")",
            "t.vla",
        )
        .unwrap();
        assert_eq!(
            g.vocab_stats(),
            "loaded: 1 rule, 0 macros, 1 test (1 expected fail) from t.vla"
        );
    }

    #[test]
    fn every_phrasebook_s_proofs_pass_whole() {
        // Raise mode: a failing proof would refuse the load, as in Excel.
        for (text, name, stats) in [
            (
                DANSK,
                "dansk.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                DEUTSCHE,
                "deutsche.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                ESPERANTO,
                "esperanto.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                FRANCAIS,
                "francais.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                LATIN,
                "latin.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                PIRATE,
                "pirate.vla",
                "19 rules, 0 macros, 20 tests (0 expected fails)",
            ),
            (
                ESPANOL,
                "espanol.vla",
                "19 rules, 134 macros, 29 tests (1 expected fail)",
            ),
            (
                ENGLISH,
                "english.vla",
                "240 rules, 231 macros, 460 tests (22 expected fails)",
            ),
        ] {
            let mut g = Grammar::new(PRELUDE);
            g.load_vocabulary_text(text, name)
                .unwrap_or_else(|e| panic!("{name}: {}", e.text));
            assert_eq!(g.vocab_stats(), format!("loaded: {stats} from {name}"));
            assert!(g.proof_failures().is_empty());
        }
    }
}
