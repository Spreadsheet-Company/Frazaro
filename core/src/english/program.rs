//! The program frame (PORT.6, slice 6e): VLA_SentenceEngine.bas's
//! `EnglishToVla`, which turns a whole program into VLA text - Define's
//! constants, `Use library`, the sheet-change and button-click handlers,
//! the To definitions in their four shapes (a sub with its `with` parameters,
//! `To get <name>:`, `To get <name> using ...:`, `To <name> of <p>:`), the
//! main sub, and the frame around them (`vla-step`, `vla-problem`, the step
//! table `BuildStepInfra` writes, the libraries and the carried macros) -
//! with `BuildSub`, `ParseDefine`, `ParseUseLibrary`, `CheckDupAction`,
//! `ValidateActionCalls` and G-PROLOG's three range lints.

use super::grammar::Grammar;
use super::matcher::{add_keyed, has_keyed, is_num_tok, is_str_tok, vla_string_lit, Parser};
use super::prolog::ref_in_range;
use super::tokenize::{line_suf, tokenize, PARA_TOK};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// What `EnglishToVla` leaves beside its text: the lint warnings it added,
/// LX.14's masking notes (`EnglishLastNotes`), IN.7's click handlers
/// (`EnglishClickHandlerNames`/`Procs`) and the step texts.
#[derive(Clone, Debug, Default)]
pub struct Translation {
    pub vla: String,
    pub lint_warnings: Vec<String>,
    pub mask_notes: Vec<String>,
    pub click_handlers: Vec<(String, String)>,
    pub step_texts: Vec<String>,
}

/// VBA `ClickHandlerSlug`: a caption as an identifier, every other
/// character an underscore.
pub fn click_handler_slug(caption: &str) -> String {
    caption
        .chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect()
}

/// VBA `JoinNames`: ", "-joined, or "(none)".
fn join_names(names: &[String]) -> String {
    if names.is_empty() {
        "(none)".to_string()
    } else {
        names.join(", ")
    }
}

impl Grammar {
    /// VBA `EnglishToVla`: a program's VLA text, with step tracking on.
    pub fn translate_program(&self, text: &str) -> Result<Translation, Refusal> {
        let mut parser = Parser::new(self);
        parser.translate_program(text)
    }

    /// VBA `EnglishToVba`: `VlaTranspile(EnglishToVla(text))`, the carried
    /// macros riding inside the text.
    pub fn translate_program_to_vba(&self, text: &str) -> Result<String, Refusal> {
        let t = self.translate_program(text)?;
        crate::emit::compile(&t.vla, &self.prelude)
    }
}

impl<'g> Parser<'g> {
    /// `EnglishToVla`'s own state, fresh for this translation.
    fn reset_program_state(&mut self) {
        self.step_count = 0;
        self.step_texts.clear();
        self.err_line = 0;
        self.act_names.clear();
        self.act_params.clear();
        self.act_req.clear();
        self.calls.clear();
        self.rule_cells.clear();
        self.asks.clear();
        self.alias_names.clear();
        self.alias_defs.clear();
        self.loop_stack.clear();
        self.try_count = 0;
        self.in_func_def = false;
        self.in_recovery = false;
        self.saw_sheet_change_event = false;
        self.click_captions.clear();
        self.click_slugs.clear();
        self.click_names.clear();
        self.click_procs.clear();
        self.user_fn_of.clear();
        self.user_fn_nullary.clear();
        self.masked_of.clear();
        self.mask_notes.clear();
        self.fn_act_names.clear();
        self.fn_using_names.clear();
        self.fn_using_req.clear();
        self.assigned_lines.clear();
        self.cur_line = 0;
        self.dict_names.clear();
        self.include_lines.clear();
        self.include_seen.clear();
        self.declared.clear();
        self.assigned.clear();
        self.lint_warnings.clear();
        self.in_program = true;
    }

    /// VBA `EnglishToVla`.
    pub fn translate_program(&mut self, text: &str) -> Result<Translation, Refusal> {
        self.src_lines = text
            .replace("\r\n", "\n")
            .split('\n')
            .map(str::to_string)
            .collect(); // S3.2
        let toks = match tokenize(text) {
            Ok(t) => t,
            Err(e) => {
                self.err_line = e.line;
                return Err(e.refusal);
            }
        };
        self.set_tokens(toks);
        self.reset_program_state();
        let mut pos = 1usize;
        let mut out_parts = String::new();
        let mut main_stmts: Vec<String> = Vec::new();
        while pos <= self.toks.count() {
            self.claim.clear(); // claim tracking is per top-level sentence
            self.cur_line = self.toks.line_at(pos);
            if self.tok(pos) == PARA_TOK {
                pos += 1; // blank lines at top level are decoration
                continue;
            }
            let w = self.peek_word(pos);
            if w == "define" {
                self.parse_define(&mut pos)?;
            } else if (w == "use" || w == "import")
                && (self.word_at(pos + 1) == "library" || self.word_at(pos + 1) == "code")
            {
                self.parse_use_library(&mut pos)?;
            } else if self.at_sheet_change_event(pos) {
                pos += 4; // when sheet changes :
                if self.saw_sheet_change_event {
                    let loc = self.line_tag(pos - 1);
                    return Err(raise("english-sheet-change-dup", &[("loc", &loc)]));
                }
                self.saw_sheet_change_event = true;
                let sub =
                    self.define_sub(&mut pos, "on:sheet-change", "", false, false, Vec::new())?;
                out_parts.push_str(&sub);
            } else if self.at_button_click_event(pos) {
                let btn_caption = self.tok(pos + 1)[1..].to_string(); // strip the token's quote
                pos += 5; // when "caption" is clicked :
                if has_keyed(&self.click_captions, &btn_caption) {
                    let loc = self.line_tag(pos - 1);
                    return Err(raise(
                        "english-click-handler-dup",
                        &[("caption", &btn_caption), ("loc", &loc)],
                    ));
                }
                add_keyed(&mut self.click_captions, &btn_caption);
                let click_slug = click_handler_slug(&btn_caption);
                if let Some(prior) = self.click_slugs.get(&fold(&click_slug)) {
                    let prior = prior.clone();
                    let loc = self.line_tag(pos - 1);
                    return Err(raise(
                        "english-click-handler-slug-collision",
                        &[("a", &btn_caption), ("b", &prior), ("loc", &loc)],
                    ));
                }
                self.click_slugs
                    .insert(fold(&click_slug), btn_caption.clone());
                let proc_name = format!("on:click:{click_slug}");
                let sub = self.define_sub(&mut pos, &proc_name, "", false, false, Vec::new())?;
                out_parts.push_str(&sub);
                self.click_names.push(btn_caption);
                self.click_procs.push(proc_name);
            } else if w == "to" {
                pos += 1;
                self.parse_to_definition(&mut pos, &mut out_parts)?;
            } else {
                let s = self.parse_tracked(&mut pos, 1)?;
                main_stmts.push(s);
            }
        }
        self.validate_action_calls()?;
        self.validate_relation_table_names()?;
        self.validate_first_match_cells()?;
        self.validate_question_ranges()?;
        if !main_stmts.is_empty() {
            let s = self.build_sub("main", &main_stmts, "", false)?;
            out_parts.push_str(&s);
            out_parts.push_str("\r\n");
        }
        if self.step_tracking && self.step_count > 0 {
            out_parts = format!(
                "(dim vla-step Long)\r\n\r\n{out_parts}{}",
                self.build_step_infra()
            );
        }
        if self.try_count > 0 {
            out_parts = format!("(dim vla-problem String)\r\n\r\n{out_parts}");
        }
        if !self.alias_defs.is_empty() {
            out_parts = format!("{}\r\n{out_parts}", self.alias_defs);
        }
        if !self.include_lines.is_empty() {
            out_parts.push_str("\r\n; ---- libraries (G12) ----\r\n");
            out_parts.push_str(&self.include_lines);
        }
        if self.g.vocab_macro_count() > 0 {
            out_parts.push_str("\r\n; ---- carried by phrasebooks (L4) ----\r\n");
            out_parts.push_str(&self.g.vocab_macros_text());
            out_parts.push_str("\r\n\r\n");
        }
        self.in_program = false;
        Ok(Translation {
            vla: out_parts,
            lint_warnings: self.lint_warnings.clone(),
            mask_notes: self.mask_notes.clone(),
            click_handlers: self
                .click_names
                .iter()
                .cloned()
                .zip(self.click_procs.iter().cloned())
                .collect(),
            step_texts: self.step_texts.clone(),
        })
    }

    /// A sub or function body in its own scope of declared and assigned
    /// names (the reference's `Set mDeclared = New Collection` and the
    /// restore to the main scope), built with `BuildSub` and followed by a
    /// blank line.
    fn define_sub(
        &mut self,
        pos: &mut usize,
        name: &str,
        param_spec: &str,
        as_func: bool,
        func_body: bool,
        declared: Vec<String>,
    ) -> Result<String, Refusal> {
        let saved_decl = std::mem::replace(&mut self.declared, declared);
        let saved_asgn = std::mem::take(&mut self.assigned);
        let saved_func = self.in_func_def;
        self.in_func_def = func_body;
        let body = self.parse_block(pos);
        self.in_func_def = saved_func;
        let built = match body {
            Ok(b) => self.build_sub(name, &b, param_spec, as_func),
            Err(e) => Err(e),
        };
        self.declared = saved_decl;
        self.assigned = saved_asgn;
        Ok(format!("{}\r\n", built?))
    }

    /// `EnglishToVla`'s To arm: an action, in one of its four shapes.
    fn parse_to_definition(
        &mut self,
        pos: &mut usize,
        out_parts: &mut String,
    ) -> Result<(), Refusal> {
        let mut get_nullary = false;
        if self.peek_word(*pos) == "get" && !self.word_at(*pos + 1).is_empty() {
            if self.tok(*pos + 2) == "using" || self.tok(*pos + 2) == "of" {
                *pos += 1;
            } else if self.tok(*pos + 2) == ":" {
                *pos += 1;
                get_nullary = true;
            }
        }
        let sub_name = self.expect_word(pos, "an action name after 'To'")?;
        self.check_name(&sub_name)?;
        self.check_dup_action(&sub_name, *pos - 1)?;
        if get_nullary {
            self.expect_tok(pos, ":", &format!("':' after 'To get {sub_name}'"))?;
            self.mask_vocab_of_word(&sub_name, false);
            self.user_fn_nullary
                .insert(fold(&sub_name), sub_name.clone());
            self.fn_act_names.push(fold(&sub_name));
            let s = self.define_sub(pos, &sub_name, "", true, true, Vec::new())?;
            out_parts.push_str(&s);
        } else if self.tok(*pos) == "using" {
            *pos += 1;
            let mut param_spec = String::new();
            let mut sig_params: Vec<String> = Vec::new();
            let mut sig_req: Vec<bool> = Vec::new();
            let mut declared: Vec<String> = Vec::new();
            let mut saw_default = false;
            loop {
                let p_name = self.expect_word(pos, "a parameter name after 'using'")?;
                if (self.fn_of_target(&p_name).is_some()
                    || self.fn_nullary_target(&p_name).is_some())
                    && !self.g.is_maskable_fn_word(&p_name)
                {
                    let loc = self.line_tag(*pos - 1);
                    return Err(raise(
                        "english-param-name-taken",
                        &[("name", &p_name), ("loc", &loc)],
                    ));
                }
                let mut p_def = String::new();
                if self.tok(*pos) == "of" {
                    *pos += 1;
                    p_def = self.parse_expr_req(pos)?;
                    self.refuse_cell_shaped_read(&p_def, self.cur_line)?; // LX.13: a default is a value
                }
                if !p_def.is_empty() {
                    saw_default = true;
                } else if saw_default {
                    let loc = self.line_tag(*pos - 1);
                    return Err(raise(
                        "english-param-missing-default",
                        &[("name", &p_name), ("loc", &loc)],
                    ));
                }
                if !param_spec.is_empty() {
                    param_spec.push(' ');
                }
                if !p_def.is_empty() {
                    param_spec.push_str(&format!("(optional {p_name} Variant {p_def})"));
                } else {
                    param_spec.push_str(&format!("({p_name} Variant)"));
                }
                self.check_name(&p_name)?; // MarkDeclared
                add_keyed(&mut declared, &p_name);
                sig_params.push(fold(&p_name));
                sig_req.push(p_def.is_empty());
                if self.tok(*pos) == "and" {
                    *pos += 1;
                } else {
                    break;
                }
            }
            self.expect_tok(pos, ":", "':' after the 'using' parameters")?;
            self.fn_using_names.insert(fold(&sub_name), sig_params);
            self.fn_using_req.insert(fold(&sub_name), sig_req);
            self.fn_act_names.push(fold(&sub_name));
            let s = self.define_sub(pos, &sub_name, &param_spec, true, true, declared)?;
            out_parts.push_str(&s);
        } else if self.tok(*pos) == "of" {
            *pos += 1;
            let p_name = self.expect_word(pos, "a parameter name after 'of'")?;
            self.expect_tok(pos, ":", &format!("':' after 'To {sub_name} of {p_name}'"))?;
            self.check_name(&p_name)?; // the parameter is never auto-dimmed
            let declared = vec![p_name.clone()];
            self.mask_vocab_of_word(&sub_name, true); // LX.14: a vocabulary's own word, kept to give back
            self.user_fn_of.insert(fold(&sub_name), sub_name.clone());
            self.fn_act_names.push(fold(&sub_name));
            let s = self.define_sub(
                pos,
                &sub_name,
                &format!("({p_name} Variant)"),
                true,
                true,
                declared,
            )?;
            out_parts.push_str(&s);
        } else {
            let mut param_spec = String::new();
            let mut sig_params: Vec<String> = Vec::new();
            let mut sig_req: Vec<bool> = Vec::new();
            let mut declared: Vec<String> = Vec::new();
            let mut saw_default = false;
            if self.tok(*pos) == "," {
                *pos += 1;
                self.expect_word_is(pos, "with")?;
                loop {
                    let mut p_opt = false;
                    if self.tok(*pos) == "optional" {
                        p_opt = true;
                        *pos += 1;
                    }
                    let p_name = self.expect_word(pos, "a parameter name after 'with'")?;
                    let mut p_def = String::new();
                    if self.tok(*pos) == "of" {
                        *pos += 1;
                        p_def = self.parse_expr_req(pos)?;
                        self.refuse_cell_shaped_read(&p_def, self.cur_line)?; // LX.13
                    }
                    if !param_spec.is_empty() {
                        param_spec.push(' ');
                    }
                    if p_opt || !p_def.is_empty() {
                        saw_default = true;
                    } else if saw_default {
                        let loc = self.line_tag(*pos - 1);
                        return Err(raise(
                            "english-param-required-after-optional",
                            &[("name", &p_name), ("loc", &loc)],
                        ));
                    }
                    if p_opt || !p_def.is_empty() {
                        let def = if p_def.is_empty() {
                            String::new()
                        } else {
                            format!(" {p_def}")
                        };
                        param_spec.push_str(&format!("(optional {p_name} Variant{def})"));
                    } else {
                        param_spec.push_str(&format!("({p_name} Variant)"));
                    }
                    self.check_name(&p_name)?; // params are never auto-dimmed
                    add_keyed(&mut declared, &p_name);
                    sig_params.push(fold(&p_name));
                    sig_req.push(!(p_opt || !p_def.is_empty()));
                    if self.tok(*pos) == "and" {
                        *pos += 1;
                    } else {
                        break;
                    }
                }
            }
            self.expect_tok(pos, ":", "':' after the action name (or its 'with' clause)")?;
            let s = self.define_sub(pos, &sub_name, &param_spec, false, false, declared)?;
            out_parts.push_str(&s);
            self.act_names.push(fold(&sub_name));
            self.act_params.push(sig_params);
            self.act_req.push(sig_req);
        }
        Ok(())
    }

    /// VBA `BuildSub`: the sub or function, its auto-dimmed names, and the
    /// error handler when steps are tracked.
    fn build_sub(
        &mut self,
        name: &str,
        stmts: &[String],
        param_spec: &str,
        as_func: bool,
    ) -> Result<String, Refusal> {
        let instrument = self.step_tracking && !stmts.is_empty();
        let mut sb = format!(
            "({} {name} ({param_spec})\r\n",
            if as_func { "function" } else { "sub" }
        );
        for e in self.assigned.clone() {
            if has_keyed(&self.declared, &e) {
                continue;
            }
            if has_keyed(&self.alias_names, &e) {
                let line = self.assigned_lines.get(&fold(&e)).copied();
                let loc = match line {
                    Some(l) => {
                        self.err_line = l;
                        line_suf(l)
                    }
                    None => String::new(),
                };
                return Err(raise(
                    "english-define-value-immutable",
                    &[("name", &e), ("loc", &loc)],
                ));
            }
            sb.push_str(&format!("  (dim {e})\r\n"));
        }
        if instrument {
            sb.push_str("  (on-error goto vla-fail)\r\n");
        }
        for e in stmts {
            sb.push_str(e);
            sb.push_str("\r\n");
        }
        if instrument {
            sb.push_str(if as_func {
                "  (exit-function)\r\n"
            } else {
                "  (exit-sub)\r\n"
            });
            sb.push_str("  (label vla-fail)\r\n");
            sb.push_str("  (vla-report-error)\r\n");
        }
        // StackCloser: the closing paren on the last line, then CRLF.
        Ok(format!("{})\r\n", sb.trim_end_matches(['\r', '\n'])))
    }

    /// VBA `BuildStepInfra`: the error reporter and the step-text lookup.
    fn build_step_infra(&self) -> String {
        let mut sb = String::from("(sub vla-report-error ()\r\n");
        sb.push_str("  (vlareportstop vla-step (vla-step-text vla-step) err.description))\r\n\r\n");
        sb.push_str("(function vla-step-text ((byval n Long)) String\r\n");
        sb.push_str("  (select n\r\n");
        for (i0, t) in self.step_texts.iter().enumerate() {
            sb.push_str(&format!(
                "    (case ({}) (return {}))\r\n",
                i0 + 1,
                vla_string_lit(t)
            ));
        }
        sb.push_str("    (case-else (return \"an unknown step\"))))\r\n");
        sb
    }

    /// VBA `ParseDefine`: `Define <name> as <fixed value>.`, a constant.
    fn parse_define(&mut self, pos: &mut usize) -> Result<(), Refusal> {
        *pos += 1; // past "define"
        let name = self.expect_word(pos, "a name after 'Define'")?;
        self.check_name(&name)?;
        self.expect_word_is(pos, "as")?;
        let t = self.tok(*pos).to_string();
        let val = if is_str_tok(&t) {
            *pos += 1;
            vla_string_lit(&t[1..])
        } else if is_num_tok(&t)
            || (!self.word_at(*pos).is_empty() && has_keyed(&self.alias_names, &t))
        {
            // a number, or an earlier alias
            *pos += 1;
            t
        } else {
            let context = self.sentence_context(*pos);
            return Err(raise(
                "english-define-needs-fixed-value",
                &[("context", &context)],
            ));
        };
        self.expect_tok(pos, ".", "'.' at the end of the sentence")?;
        add_keyed(&mut self.alias_names, &name);
        self.alias_defs
            .push_str(&format!("(const {name} {val})\r\n"));
        Ok(())
    }

    /// VBA `ParseUseLibrary` (G12): `Use library "name".`, spliced once.
    fn parse_use_library(&mut self, pos: &mut usize) -> Result<(), Refusal> {
        *pos += 2; // use|import library|code
        if self.word_at(*pos) == "at" || self.word_at(*pos) == "from" {
            *pos += 1;
        }
        let ftok = self.tok(*pos).to_string();
        if !is_str_tok(&ftok) {
            let context = self.sentence_context(*pos);
            let loc = self.line_tag(*pos);
            return Err(raise(
                "english-library-name-not-quoted",
                &[("context", &context), ("loc", &loc)],
            ));
        }
        *pos += 1;
        let fname = &ftok[1..];
        self.expect_tok(pos, ".", "'.' after the library name")?;
        let k = fold(fname);
        if !self.include_seen.insert(k) {
            return Ok(()); // seen before - splice once
        }
        self.include_lines
            .push_str(&format!("(include \"{fname}\")\r\n"));
        Ok(())
    }

    /// VBA `CheckDupAction` (B4): one name, one definition.
    fn check_dup_action(&mut self, name: &str, name_pos: usize) -> Result<(), Refusal> {
        let n = fold(name);
        // A sub's name, a value action's name (every `using` action is one
        // of those too), or a using action's: taken.
        if self.act_names.iter().any(|e| e == &n)
            || self.fn_act_names.iter().any(|e| e == &n)
            || self.is_using_fn(&n)
        {
            let loc = self.line_tag(name_pos);
            return Err(raise(
                "english-action-name-taken",
                &[("name", name), ("loc", &loc)],
            ));
        }
        if (self.fn_of_target(&n).is_some() || self.fn_nullary_target(&n).is_some())
            && !self.g.is_maskable_fn_word(&n)
        {
            let loc = self.line_tag(name_pos);
            return Err(raise(
                "english-action-name-means-something",
                &[("name", name), ("loc", &loc)],
            ));
        }
        Ok(())
    }

    /// VBA `ValidateActionCalls`: every recorded call against the actions
    /// defined; an undefined action is a warning, a wrong argument a refusal.
    fn validate_action_calls(&mut self) -> Result<(), Refusal> {
        let calls = self.calls.clone();
        for call in &calls {
            let Some(idx) = self.act_names.iter().position(|a| a == &call.name) else {
                self.lint_warnings.push(format!(
                    "'{}' calls '{}', which is not defined in this file - if it is not defined elsewhere, running will fail",
                    call.text, call.name
                ));
                continue;
            };
            let params = self.act_params[idx].clone();
            let reqs = self.act_req[idx].clone();
            for a in &call.args {
                if !params.contains(a) {
                    self.err_line = call.line;
                    return Err(raise(
                        "english-call-unknown-param",
                        &[
                            ("call", &call.text),
                            ("action", &call.name),
                            ("param", a),
                            ("list", &join_names(&params)),
                            ("loc", &line_suf(call.line)),
                        ],
                    ));
                }
            }
            for (j, p) in params.iter().enumerate() {
                if reqs[j] && !call.args.contains(p) {
                    self.err_line = call.line;
                    return Err(raise(
                        "english-call-missing-param",
                        &[
                            ("call", &call.text),
                            ("action", &call.name),
                            ("param", p),
                            ("loc", &line_suf(call.line)),
                        ],
                    ));
                }
            }
        }
        Ok(())
    }

    /// VBA `ValidateRelationTableNames` (G-PROLOG slice 3): a relation one
    /// cell defines may not be a table another cell reads.
    fn validate_relation_table_names(&mut self) -> Result<(), Refusal> {
        let cells = self.rule_cells.clone();
        for (ci, c) in cells.iter().enumerate() {
            for (cj, other) in cells.iter().enumerate() {
                if cj == ci {
                    continue;
                }
                if other.tables.iter().any(|tb| tb == &c.rel) {
                    self.err_line = c.line;
                    return Err(raise(
                        "english-program-relation-names-table",
                        &[
                            ("relation", &c.rel),
                            ("cell", &c.addr.to_uppercase()),
                            ("tablecell", &other.addr.to_uppercase()),
                            ("loc", &line_suf(c.line)),
                        ],
                    ));
                }
            }
        }
        Ok(())
    }

    /// VBA `ValidateFirstMatchCells`: a first-match relation lives in one cell.
    fn validate_first_match_cells(&mut self) -> Result<(), Refusal> {
        let cells = self.rule_cells.clone();
        for (ci, c) in cells.iter().enumerate() {
            if !c.first {
                continue;
            }
            for (cj, other) in cells.iter().enumerate() {
                if cj != ci && other.rel == c.rel {
                    self.err_line = other.line;
                    return Err(raise(
                        "english-program-first-match-split",
                        &[
                            ("relation", &c.rel),
                            ("cell", &c.addr.to_uppercase()),
                            ("othercell", &other.addr.to_uppercase()),
                            ("loc", &line_suf(other.line)),
                        ],
                    ));
                }
            }
        }
        Ok(())
    }

    /// VBA `ValidateQuestionRanges`: every cell a question's relations
    /// reach must sit in the rules range it names, when the cells can be
    /// placed at all.
    fn validate_question_ranges(&mut self) -> Result<(), Refusal> {
        let asks = self.asks.clone();
        let cells = self.rule_cells.clone();
        for ask in &asks {
            let mut reach: Vec<String> = vec![ask.rel.clone()];
            let mut k = 0;
            while k < reach.len() {
                let nm = reach[k].clone();
                let mut cells_of: Vec<String> = Vec::new();
                let mut total = 0;
                let mut inside = 0;
                let mut missing = String::new();
                let mut unplaced = false;
                for c in cells.iter().filter(|c| c.rel == nm) {
                    for bn in &c.bodies {
                        if !reach.contains(bn) {
                            reach.push(bn.clone());
                        }
                    }
                    if !cells_of.contains(&c.addr) {
                        cells_of.push(c.addr.clone());
                        total += 1;
                        match ref_in_range(&c.addr, &ask.range) {
                            1 => inside += 1,
                            0 => {
                                if missing.is_empty() {
                                    missing = c.addr.clone();
                                }
                            }
                            _ => unplaced = true,
                        }
                    }
                }
                if !unplaced && inside > 0 && inside < total {
                    self.err_line = ask.line;
                    return Err(raise(
                        "english-question-range-splits-relation",
                        &[
                            ("relation", &nm),
                            ("cell", &missing.to_uppercase()),
                            ("range", &ask.range.to_uppercase()),
                            ("loc", &line_suf(ask.line)),
                        ],
                    ));
                }
                k += 1;
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../../scripts/polyglotta/english.vla");
    const PROGRAM: &str = include_str!("../../../scripts/instructions.txt");
    const GOLDEN: &str = include_str!("../../../scripts/instructions_golden.vla");

    fn english() -> Grammar {
        let mut g = Grammar::new(PRELUDE);
        g.load_vocabulary_text(ENGLISH, "english.vla").unwrap();
        g
    }

    fn normalized(s: &str) -> String {
        s.replace("\r\n", "\n").trim_end_matches('\n').to_string()
    }

    /// The treaty's oracle 1: `instructions.txt` to `instructions_golden.vla`
    /// less its stamp line, after the treaty's normalization.
    #[test]
    fn the_translate_golden_is_reproduced() {
        let g = english();
        let t = g
            .translate_program(PROGRAM)
            .unwrap_or_else(|e| panic!("{}", e.text));
        let want = normalized(GOLDEN.split_once('\n').map(|(_, rest)| rest).unwrap_or(""));
        let got = normalized(&t.vla);
        if got != want {
            let at = got
                .chars()
                .zip(want.chars())
                .position(|(a, b)| a != b)
                .unwrap_or(got.chars().count().min(want.chars().count()));
            let want_line = want.chars().take(at).filter(|c| *c == '\n').count() + 1;
            let ctx = |s: &str| {
                s.chars()
                    .skip(at.saturating_sub(80))
                    .take(240)
                    .collect::<String>()
            };
            panic!(
                "differs at char {at} (golden line {want_line}) of {}:\n  want: {}\n  got:  {}",
                want.chars().count(),
                ctx(&want),
                ctx(&got)
            );
        }
        assert!(t.lint_warnings.is_empty(), "{:?}", t.lint_warnings);
    }

    #[test]
    fn the_frame_and_the_definitions_are_the_reference_s() {
        let g = Grammar::new(PRELUDE);
        let t = g
            .translate_program("Define hot-pink as \"#FF69B4\".\n\nTo stamp, with row-number of 1 and value of \"ok\":\nSet x to value.\n\nTo tax of amount:\nGive back amount times 0.08.\n\nTo get vat-rate:\nGive back 0.2.\n\nTo get commission using sale of 1000 and rate of 5%:\nGive back sale times rate.\n\nSet total to tax of 100 plus vat-rate plus commission using sale of 10.\nStamp with row-number of 2.\n")
            .unwrap_or_else(|e| panic!("{}", e.text));
        let want = "(const hot-pink \"#FF69B4\")\r\n\r\n(dim vla-step Long)\r\n\r\n\
(sub stamp ((optional row-number Variant 1) (optional value Variant \"ok\"))\r\n  (dim x)\r\n  (on-error goto vla-fail)\r\n    (set! vla-step 1)\r\n    (if (vlatraceon) (then (vlatracestep 1 (vla-step-text 1))))\r\n    (at-line 4\r\n    (set! x value))\r\n  (exit-sub)\r\n  (label vla-fail)\r\n  (vla-report-error))\r\n\r\n\
(function tax ((amount Variant))\r\n  (on-error goto vla-fail)\r\n    (set! vla-step 2)\r\n    (if (vlatraceon) (then (vlatracestep 2 (vla-step-text 2))))\r\n    (at-line 7\r\n    (return (* amount 0.08)))\r\n  (exit-function)\r\n  (label vla-fail)\r\n  (vla-report-error))\r\n\r\n\
(function vat-rate ()\r\n  (on-error goto vla-fail)\r\n    (set! vla-step 3)\r\n    (if (vlatraceon) (then (vlatracestep 3 (vla-step-text 3))))\r\n    (at-line 10\r\n    (return 0.2))\r\n  (exit-function)\r\n  (label vla-fail)\r\n  (vla-report-error))\r\n\r\n\
(function commission ((optional sale Variant 1000) (optional rate Variant (/ 5 100)))\r\n  (on-error goto vla-fail)\r\n    (set! vla-step 4)\r\n    (if (vlatraceon) (then (vlatracestep 4 (vla-step-text 4))))\r\n    (at-line 13\r\n    (return (* sale rate)))\r\n  (exit-function)\r\n  (label vla-fail)\r\n  (vla-report-error))\r\n\r\n\
(sub main ()\r\n  (dim total)\r\n  (on-error goto vla-fail)\r\n  (set! vla-step 5)\r\n  (if (vlatraceon) (then (vlatracestep 5 (vla-step-text 5))))\r\n  (at-line 15\r\n  (set! total (+ (+ (tax 100) (vat-rate)) (commission :sale 10))))\r\n  (set! vla-step 6)\r\n  (if (vlatraceon) (then (vlatracestep 6 (vla-step-text 6))))\r\n  (at-line 16\r\n  (stamp :row-number 2))\r\n  (exit-sub)\r\n  (label vla-fail)\r\n  (vla-report-error))\r\n\r\n\
(sub vla-report-error ()\r\n  (vlareportstop vla-step (vla-step-text vla-step) err.description))\r\n\r\n(function vla-step-text ((byval n Long)) String\r\n  (select n\r\n    (case (1) (return \"Set x to value. [line 4]\"))\r\n    (case (2) (return \"Give back amount times 0.08. [line 7]\"))\r\n    (case (3) (return \"Give back 0.2. [line 10]\"))\r\n    (case (4) (return \"Give back sale times rate. [line 13]\"))\r\n    (case (5) (return \"Set total to tax of 100 plus vat-rate plus commission using sale of 10. [line 15]\"))\r\n    (case (6) (return \"Stamp with row-number of 2. [line 16]\"))\r\n    (case-else (return \"an unknown step\"))))\r\n";
        assert_eq!(t.vla, want);
    }

    #[test]
    fn the_program_level_refusals_are_the_reference_s() {
        let g = Grammar::new(PRELUDE);
        let id = |text: &str| g.translate_program(text).unwrap_err().id;
        assert_eq!(
            id("To greet:\nLog 1.\n\nTo greet:\nLog 2.\n"),
            "english-action-name-taken"
        );
        assert_eq!(
            id("To length of x:\nGive back 1.\n"),
            "english-action-name-means-something"
        );
        assert_eq!(
            id("To greet, with a of 1 and b:\nLog 1.\n"),
            "english-param-required-after-optional"
        );
        assert_eq!(
            id("To get f using a of 1 and b:\nGive back 1.\n"),
            "english-param-missing-default"
        );
        assert_eq!(
            id("Greet with x of 1.\n\nTo greet, with y of 1:\nLog y.\n"),
            "english-call-unknown-param"
        );
        assert_eq!(
            id("Greet.\n\nTo greet, with y:\nLog y.\n"),
            "english-call-missing-param"
        );
        assert_eq!(
            id("Define x as 5.\nSet x to 6.\n"),
            "english-define-value-immutable"
        );
        assert_eq!(
            id("Define x as total.\n"),
            "english-define-needs-fixed-value"
        );
        assert_eq!(id("Use library x.\n"), "english-library-name-not-quoted");
        assert_eq!(
            id("When the sheet changes:\nLog 1.\n\nWhen the sheet changes:\nLog 2.\n"),
            "english-sheet-change-dup"
        );
        assert_eq!(
            id("When \"Go\" is clicked:\nLog 1.\n\nWhen \"Go\" is clicked:\nLog 2.\n"),
            "english-click-handler-dup"
        );
        assert_eq!(
            id("When \"Go!\" is clicked:\nLog 1.\n\nWhen \"Go?\" is clicked:\nLog 2.\n"),
            "english-click-handler-slug-collision"
        );
        assert_eq!(
            id("When Go is clicked:\nLog 1.\n"),
            "english-click-handler-needs-quotes"
        );
        assert_eq!(
            id("When \"Go\" is clikced:\nLog 1.\n"),
            "english-clicked-misspelled"
        );
        let t = g.translate_program("Greet.\n").unwrap();
        // The call's sentence is rendered from its folded tokens, as the
        // reference renders it.
        assert_eq!(t.lint_warnings, vec!["'greet.' calls 'greet', which is not defined in this file - if it is not defined elsewhere, running will fail".to_string()]);
        let t = g.translate_program("Use library \"alien.vla\".\nUse library \"alien.vla\".\nWhen \"Go\" is clicked:\nLog 1.\n").unwrap();
        assert!(
            t.vla
                .contains("; ---- libraries (G12) ----\r\n(include \"alien.vla\")\r\n"),
            "{}",
            t.vla
        );
        assert_eq!(
            t.click_handlers,
            vec![("Go".to_string(), "on:click:Go".to_string())]
        );
        assert!(
            t.vla
                .starts_with("(dim vla-step Long)\r\n\r\n(sub on:click:Go ()\r\n"),
            "{}",
            t.vla
        );
    }
}
