//! Statements: VLA.bas's `EmitStmt` and the emitters it dispatches to,
//! `EmitCallStmt`, `EmitReturn`, `EmitIf`, `EmitFor`, `EmitForEach`,
//! `EmitForEachRow`, `EmitSelect`, `EmitDimCore`, `DimSpecOf`, `BoundList`,
//! `EmitRedim`, `EmitTypeDef`, `EmitEnumDef`, `EmitVisibility`,
//! `EmitConstCore`, `EmitDoc`, and `MapTag`, the source-map comment.

use super::expr::{kw_misuse_msg, refuse_interpreter_only_call};
use super::names::{click_handler_slug, sym_name};
use super::{Compiler, MAKEBUTTON_SUB};
use crate::form::{str_lit_content, sym_text, Form, List};
use crate::headtable::resolve_head_alias;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

fn pad_of(ind: usize) -> String {
    " ".repeat(ind * 4)
}

impl Compiler {
    /// VBA `EmitStmt`.
    pub(super) fn emit_stmt(&mut self, s: &Form, ind: usize) -> Result<String, Refusal> {
        let pad = pad_of(ind);
        let Form::List(lst) = s else {
            return Err(raise(
                "vla-bare-atom-statement",
                &[("value", &s.atom_text().unwrap_or_default())],
            ));
        };
        // C2: the source map. Remember the mapped line and tag the
        // emitted statement's first line below.
        let src_ln = lst.line;
        if src_ln > 0 {
            self.emit_line = src_ln;
        }
        let h = resolve_head_alias(&fold(&lst.head_sym()?));
        let mut r = String::new();
        match h.as_str() {
            "dim" => r = format!("{pad}{}\r\n", self.emit_dim_core(lst)?),
            "const" => r = format!("{pad}{}\r\n", self.emit_const_core(lst)?),
            "set!" => {
                // SEC.15: exactly (set! (. obj formula) v) is a call to the
                // runtime's VlaSetFormula, the one formula sink of both
                // backends (it refuses an egress call by name, then writes
                // through Formula2, as this arm wrote directly until SEC.15).
                let mut wrote_formula2 = false;
                if lst.nth(2)?.head_is(".") {
                    let set_dot = lst.nth_list(2)?;
                    if set_dot.count() == 3 && fold(&sym_text(set_dot.nth(3)?)?) == "formula" {
                        r = format!(
                            "{pad}Call VlaSetFormula({}, {})\r\n",
                            self.emit_expr(set_dot.nth(2)?)?,
                            self.emit_expr(lst.nth(3)?)?
                        );
                        wrote_formula2 = true;
                    }
                }
                if !wrote_formula2 {
                    r = format!(
                        "{pad}{} = {}\r\n",
                        self.emit_expr(lst.nth(2)?)?,
                        self.emit_expr(lst.nth(3)?)?
                    );
                }
            }
            "obj-set!" => {
                r = format!(
                    "{pad}Set {} = {}\r\n",
                    self.emit_expr(lst.nth(2)?)?,
                    self.emit_expr(lst.nth(3)?)?
                )
            }
            "if" => r = self.emit_if(lst, ind)?,
            "for" => r = self.emit_for(lst, ind)?,
            "for-each" => r = self.emit_for_each(lst, ind)?,
            "for-each-row" => r = self.emit_for_each_row(lst, ind)?,
            "while" => {
                r = format!(
                    "{pad}Do While {}\r\n{}{pad}Loop\r\n",
                    self.emit_expr(lst.nth(2)?)?,
                    self.emit_body(lst, 3, ind + 1)?
                )
            }
            "do-until" => {
                r = format!(
                    "{pad}Do Until {}\r\n{}{pad}Loop\r\n",
                    self.emit_expr(lst.nth(2)?)?,
                    self.emit_body(lst, 3, ind + 1)?
                )
            }
            "select" => r = self.emit_select(lst, ind)?,
            "with" => {
                r = format!(
                    "{pad}With {}\r\n{}{pad}End With\r\n",
                    self.emit_expr(lst.nth(2)?)?,
                    self.emit_body(lst, 3, ind + 1)?
                )
            }
            "return" => r = self.emit_return(lst, &pad)?,
            "exit-sub" => r = format!("{pad}Exit Sub\r\n"),
            "exit-function" => r = format!("{pad}Exit Function\r\n"),
            "exit-for" => r = format!("{pad}Exit For\r\n"),
            "exit-do" => r = format!("{pad}Exit Do\r\n"),
            "redim" => r = format!("{pad}{}\r\n", self.emit_redim(lst)?),
            "type" | "enum" => {
                return Err(raise("vla-type-enum-module-level-only", &[("head", &h)]))
            }
            "on-error" => {
                let k = fold(&sym_text(lst.nth(2)?)?);
                if k == "resume-next" {
                    r = format!("{pad}On Error Resume Next\r\n");
                } else if k == "goto" {
                    r = format!(
                        "{pad}On Error GoTo {}\r\n",
                        sym_name(&sym_text(lst.nth(3)?)?)?
                    );
                } else {
                    return Err(raise("vla-on-error-bad-shape", &[]));
                }
            }
            "goto" => r = format!("{pad}GoTo {}\r\n", sym_name(&sym_text(lst.nth(2)?)?)?),
            "label" => r = format!("{}:\r\n", sym_name(&sym_text(lst.nth(2)?)?)?),
            "quote" => return Err(raise("vla-quote-in-statement-position", &[])),
            "deflambda" => r = format!("{pad}{}\r\n", self.emit_deflambda(lst)?),
            "include" => return Err(raise("vla-include-must-stand-alone", &[])),
            "doc" => r = self.emit_doc(lst, &pad)?,
            "resume" => {
                if lst.count() == 1 {
                    r = format!("{pad}Resume\r\n");
                } else if fold(&sym_text(lst.nth(2)?)?) == "next" {
                    r = format!("{pad}Resume Next\r\n");
                } else {
                    r = format!("{pad}Resume {}\r\n", sym_name(&sym_text(lst.nth(2)?)?)?);
                }
            }
            "debug-print" => {
                let mut parts: Vec<String> = Vec::with_capacity(lst.count());
                for i in 2..=lst.count() {
                    parts.push(self.emit_expr(lst.nth(i)?)?);
                }
                let parts = parts.join("; ");
                r = format!(
                    "{pad}Debug.Print{}\r\n",
                    if parts.is_empty() {
                        String::new()
                    } else {
                        format!(" {parts}")
                    }
                );
            }
            "raw" => r = format!("{pad}{}\r\n", str_lit_content(lst.nth(2)?)?),
            "begin" => r = self.emit_body(lst, 2, ind)?,
            "call" => r = format!("{pad}{}\r\n", self.emit_call_stmt(lst, 2)?),
            "." => r = format!("{pad}Call {}\r\n", self.emit_dot_text(lst)?),
            "at-line" => {
                // V4: a zero-runtime annotation; N rides every map tag
                // inside as src:N, by dynamic extent. On a refusal inside,
                // the restore is skipped so the message can name the line.
                if lst.count() < 3 {
                    return Err(raise("vla-at-line-arity", &[]));
                }
                let at_txt = sym_text(lst.nth(2)?)?;
                if at_txt.is_empty() {
                    return Err(raise("vla-at-line-missing-number", &[]));
                }
                if !at_txt.chars().all(|c| c.is_ascii_digit()) {
                    return Err(raise("vla-at-line-not-a-number", &[("value", &at_txt)]));
                }
                let at_prev = self.at_line;
                self.at_line = at_txt.parse().unwrap_or(0);
                for i in 3..=lst.count() {
                    r.push_str(&self.emit_stmt(lst.nth(i)?, ind)?);
                }
                self.at_line = at_prev;
            }
            "gen-row" => {
                // LISTOPS-PROVENANCE: a generator's own row label, by
                // dynamic extent, as at-line.
                if lst.count() < 3 {
                    return Err(raise("vla-gen-row-arity", &[]));
                }
                let row_label = str_lit_content(lst.nth(2)?)?;
                let row_prev = std::mem::replace(&mut self.gen_row, row_label);
                for i in 3..=lst.count() {
                    r.push_str(&self.emit_stmt(lst.nth(i)?, ind)?);
                }
                self.gen_row = row_prev;
            }
            "then" | "else" | "elseif" | "case" | "case-else" => {
                return Err(raise("vla-clause-outside-parent", &[("head", &h)]))
            }
            "make-button" => {
                // IN.7: the handler is on:click: plus the caption's slug;
                // no such sub in this compile means a decorative button.
                let caption = str_lit_content(lst.nth(2)?)?;
                let raw = format!("on:click:{}", click_handler_slug(&caption));
                let proc_lit = if self.sub_exists(&raw) {
                    format!("\"{}\"", sym_name(&raw)?)
                } else {
                    "\"\"".to_string()
                };
                r = format!(
                    "{pad}Call {MAKEBUTTON_SUB}({}, {}, {proc_lit})\r\n",
                    self.emit_expr(lst.nth(2)?)?,
                    self.emit_expr(lst.nth(3)?)?
                );
            }
            _ => r = format!("{pad}{}\r\n", self.emit_call_stmt(lst, 1)?),
        }
        // The map tag on the first emitted line: not on begin (it only
        // delegates), raw (the user's verbatim VBA), at-line or gen-row
        // (their text is their inner statements, already tagged).
        if (src_ln > 0 || !self.gen_row.is_empty())
            && h != "begin"
            && h != "raw"
            && h != "at-line"
            && h != "gen-row"
            && !r.is_empty()
        {
            r = self.map_tag(&r, src_ln);
        }
        Ok(r)
    }

    /// VBA `MapTag`: ` ' vla:N[ src:M][ vla-row:L]` on the first emitted
    /// line; without a line (a macro body under gen-row) just the comment
    /// mark and the row.
    fn map_tag(&self, emitted: &str, src_line: u32) -> String {
        let mut tag = if src_line > 0 {
            let mut t = format!(" ' vla:{src_line}");
            if self.at_line > 0 {
                t.push_str(&format!(" src:{}", self.at_line));
            }
            t
        } else {
            " '".to_string()
        };
        if !self.gen_row.is_empty() {
            tag.push_str(&format!(" vla-row:{}", self.gen_row));
        }
        match emitted.find("\r\n") {
            None => format!("{emitted}{tag}"),
            Some(i) => format!("{}{tag}{}", &emitted[..i], &emitted[i..]),
        }
    }

    /// VBA `EmitCallStmt`: a procedure call in statement position.
    fn emit_call_stmt(&self, lst: &List, name_idx: usize) -> Result<String, Refusal> {
        let op_hd = fold(&sym_text(lst.nth(name_idx)?)?);
        // L8 rider: an operator can never head a statement.
        if matches!(
            op_hd.as_str(),
            "+" | "-"
                | "*"
                | "&"
                | "/"
                | "\\"
                | "="
                | "<>"
                | "<"
                | ">"
                | "<="
                | ">="
                | "and"
                | "or"
                | "not"
                | "xor"
                | "mod"
                | "is"
                | "like"
                | "imp"
                | "eqv"
        ) {
            return Err(raise(
                "vla-operator-in-statement-position",
                &[("op", &op_hd)],
            ));
        }
        // P.L4 guard: a keyword can never head a form.
        if Form::sym(&op_hd).is_keyword_arg() {
            return Err(raise(
                "vla-keyword-misuse",
                &[("msg", &kw_misuse_msg(&op_hd, "the head of a statement")?)],
            ));
        }
        refuse_interpreter_only_call(&op_hd)?;
        let name = sym_name(&sym_text(lst.nth(name_idx)?)?)?;
        let args = self.emit_args(lst, name_idx + 1)?;
        Ok(if args.is_empty() {
            format!("Call {name}")
        } else {
            format!("Call {name}({args})")
        })
    }

    /// VBA `EmitReturn`: `(return value)` sets the function's name and
    /// exits; under an armed TCO a self-return rebinds and jumps.
    fn emit_return(&self, lst: &List, pad: &str) -> Result<String, Refusal> {
        if lst.count() < 2 {
            return Ok(format!(
                "{pad}{}\r\n",
                if self.in_function {
                    "Exit Function"
                } else {
                    "Exit Sub"
                }
            ));
        }
        if !self.in_function {
            return Err(raise("vla-return-outside-function", &[]));
        }
        if !self.tco_name.is_empty() {
            if let Form::List(inner) = lst.nth(2)? {
                if let Some(Form::Sym(callee)) = inner.items.first() {
                    if fold(&sym_name(callee)?) == self.tco_name {
                        if inner.count() - 1 != self.tco_params.len() {
                            return Err(raise(
                                "vla-tco-arity",
                                &[
                                    ("name", callee),
                                    ("given", &(inner.count() - 1).to_string()),
                                    ("declared", &self.tco_params.len().to_string()),
                                ],
                            ));
                        }
                        let mut r = String::new();
                        if self.tco_params.len() == 1 {
                            r.push_str(&format!(
                                "{pad}{} = {}\r\n",
                                self.tco_params[0],
                                self.emit_expr(inner.nth(2)?)?
                            ));
                        } else {
                            // Every argument before any parameter changes.
                            for i in 1..=self.tco_params.len() {
                                r.push_str(&format!(
                                    "{pad}vla_tco_{i} = {}\r\n",
                                    self.emit_expr(inner.nth(i + 1)?)?
                                ));
                            }
                            for (i, p) in self.tco_params.iter().enumerate() {
                                r.push_str(&format!("{pad}{p} = vla_tco_{}\r\n", i + 1));
                            }
                        }
                        r.push_str(&format!("{pad}GoTo vla_tco\r\n"));
                        return Ok(r);
                    }
                }
            }
        }
        Ok(format!(
            "{pad}{} = {}\r\n{pad}Exit Function\r\n",
            self.func_name,
            self.emit_expr(lst.nth(2)?)?
        ))
    }

    /// `(if test (then s...) [(elseif test s...)]* [(else s...)])`.
    fn emit_if(&mut self, lst: &List, ind: usize) -> Result<String, Refusal> {
        let pad = pad_of(ind);
        let mut r = format!("{pad}If {} Then\r\n", self.emit_expr(lst.nth(2)?)?);
        for i in 3..=lst.count() {
            let cl = lst.nth_list(i)?;
            let ch = fold(&cl.head_sym()?);
            match ch.as_str() {
                "then" => r.push_str(&self.emit_body(cl, 2, ind + 1)?),
                "elseif" => {
                    r.push_str(&format!(
                        "{pad}ElseIf {} Then\r\n{}",
                        self.emit_expr(cl.nth(2)?)?,
                        self.emit_body(cl, 3, ind + 1)?
                    ));
                }
                "else" => {
                    r.push_str(&format!("{pad}Else\r\n{}", self.emit_body(cl, 2, ind + 1)?));
                }
                _ => return Err(raise("vla-if-bad-clause", &[("head", &ch)])),
            }
        }
        r.push_str(&format!("{pad}End If\r\n"));
        Ok(r)
    }

    /// `(for (i start end [step]) body...)`.
    fn emit_for(&mut self, lst: &List, ind: usize) -> Result<String, Refusal> {
        let pad = pad_of(ind);
        let hdr = lst.nth_list(2)?;
        let v = sym_name(&sym_text(hdr.nth(1)?)?)?;
        let mut r = format!(
            "{pad}For {v} = {} To {}",
            self.emit_expr(hdr.nth(2)?)?,
            self.emit_expr(hdr.nth(3)?)?
        );
        if hdr.count() >= 4 {
            r.push_str(&format!(" Step {}", self.emit_expr(hdr.nth(4)?)?));
        }
        r.push_str(&format!(
            "\r\n{}{pad}Next {v}\r\n",
            self.emit_body(lst, 3, ind + 1)?
        ));
        Ok(r)
    }

    /// `(for-each (x collection) body...)`.
    fn emit_for_each(&mut self, lst: &List, ind: usize) -> Result<String, Refusal> {
        let pad = pad_of(ind);
        let hdr = lst.nth_list(2)?;
        let v = sym_name(&sym_text(hdr.nth(1)?)?)?;
        Ok(format!(
            "{pad}For Each {v} In {}\r\n{}{pad}Next {v}\r\n",
            self.emit_expr(hdr.nth(2)?)?,
            self.emit_body(lst, 3, ind + 1)?
        ))
    }

    /// `(for-each-row (row rng) body...)`: PF.4c, one bulk read, the rows
    /// walked by index, one bulk write-back; `(row i)` in the body compiles
    /// to the array element. The hidden names carry a per-occurrence
    /// number, never a random one.
    fn emit_for_each_row(&mut self, lst: &List, ind: usize) -> Result<String, Refusal> {
        self.slab_counter += 1;
        let n = self.slab_counter;
        let pad = pad_of(ind);
        let hdr = lst.nth_list(2)?;
        let v = sym_name(&sym_text(hdr.nth(1)?)?)?;
        let rng_var = format!("vlaSlabRange{n}");
        let arr_var = format!("vlaSlabArr{n}");
        let i_var = format!("vlaSlabI{n}");
        let mut r = format!("{pad}Dim {rng_var} As Range\r\n");
        r.push_str(&format!(
            "{pad}Set {rng_var} = {}\r\n",
            self.emit_expr(hdr.nth(2)?)?
        ));
        r.push_str(&format!("{pad}Dim {arr_var} As Variant\r\n"));
        r.push_str(&format!("{pad}{arr_var} = VlaSlabRead({rng_var})\r\n"));
        r.push_str(&format!("{pad}Dim {i_var} As Long\r\n"));
        r.push_str(&format!(
            "{pad}For {i_var} = LBound({arr_var}, 1) To UBound({arr_var}, 1)\r\n"
        ));
        // Dynamic extent: the outer binding comes back after the body.
        let saved_row = std::mem::replace(&mut self.slab_row_var, fold(&v));
        let saved_arr = std::mem::replace(&mut self.slab_arr_var, arr_var.clone());
        let saved_i = std::mem::replace(&mut self.slab_i_var, i_var.clone());
        let body = self.emit_body(lst, 3, ind + 1);
        self.slab_row_var = saved_row;
        self.slab_arr_var = saved_arr;
        self.slab_i_var = saved_i;
        r.push_str(&body?);
        r.push_str(&format!("{pad}Next {i_var}\r\n"));
        r.push_str(&format!("{pad}VlaSlabWrite {arr_var}, {rng_var}\r\n"));
        Ok(r)
    }

    /// `(select expr (case (v1 v2 ...) body...) ... (case-else body...))`.
    fn emit_select(&mut self, lst: &List, ind: usize) -> Result<String, Refusal> {
        let pad = pad_of(ind);
        let pad2 = pad_of(ind + 1);
        let mut r = format!("{pad}Select Case {}\r\n", self.emit_expr(lst.nth(2)?)?);
        for i in 3..=lst.count() {
            let cl = lst.nth_list(i)?;
            let ch = fold(&cl.head_sym()?);
            if ch == "case" {
                let vals = cl.nth_list(2)?;
                let mut vparts: Vec<String> = Vec::with_capacity(vals.count());
                for j in 1..=vals.count() {
                    vparts.push(self.emit_expr(vals.nth(j)?)?);
                }
                r.push_str(&format!(
                    "{pad2}Case {}\r\n{}",
                    vparts.join(", "),
                    self.emit_body(cl, 3, ind + 2)?
                ));
            } else if ch == "case-else" {
                r.push_str(&format!(
                    "{pad2}Case Else\r\n{}",
                    self.emit_body(cl, 2, ind + 2)?
                ));
            } else {
                return Err(raise("vla-select-bad-clause", &[("head", &ch)]));
            }
        }
        r.push_str(&format!("{pad}End Select\r\n"));
        Ok(r)
    }

    /// `(dim name [type])` to `Dim name As type`.
    pub(super) fn emit_dim_core(&self, lst: &List) -> Result<String, Refusal> {
        Ok(format!("Dim {}", self.dim_spec_of(lst, 2)?))
    }

    /// C2: the `name As Type` half of a declaration; the type slot is a
    /// type name or an `(array [type] [bounds...])` spec.
    fn dim_spec_of(&self, lst: &List, name_idx: usize) -> Result<String, Refusal> {
        let nm = sym_name(&sym_text(lst.nth(name_idx)?)?)?;
        if lst.count() < name_idx + 1 {
            return Ok(format!("{nm} As Variant"));
        }
        let t = lst.nth(name_idx + 1)?;
        let al = match t {
            Form::List(al) => al,
            _ => {
                return Ok(format!(
                    "{nm} As {}",
                    sym_name(&t.atom_text().unwrap_or_default())?
                ))
            }
        };
        let head = al.head_sym()?;
        if fold(&head) != "array" {
            return Err(raise(
                "vla-decl-bad-type",
                &[("name", &nm), ("head", &head)],
            ));
        }
        let mut atype = "Variant".to_string();
        if al.count() >= 2 {
            atype = sym_name(&sym_text(al.nth(2)?)?)?;
        }
        Ok(format!("{nm}({}) As {atype}", self.bound_list(al, 3)?))
    }

    /// Bounds: an expression (the upper bound) or `(to lower upper)`,
    /// comma-joined.
    fn bound_list(&self, lst: &List, from_idx: usize) -> Result<String, Refusal> {
        let mut parts: Vec<String> = Vec::new();
        for i in from_idx..=lst.count() {
            let b = lst.nth(i)?;
            let piece = match b {
                Form::List(bl) if fold(&bl.head_sym()?) == "to" => {
                    if bl.count() != 3 {
                        return Err(raise("vla-array-bound-arity", &[]));
                    }
                    format!(
                        "{} To {}",
                        self.emit_expr(bl.nth(2)?)?,
                        self.emit_expr(bl.nth(3)?)?
                    )
                }
                _ => self.emit_expr(b)?,
            };
            parts.push(piece);
        }
        Ok(parts.join(", "))
    }

    /// C2: `(redim [preserve] name bound...)`.
    fn emit_redim(&self, lst: &List) -> Result<String, Refusal> {
        let mut i = 2;
        let mut pres = "";
        if let Form::Sym(s) = lst.nth(2)? {
            if fold(s) == "preserve" {
                pres = "Preserve ";
                i = 3;
            }
        }
        let nm = sym_name(&sym_text(lst.nth(i)?)?)?;
        if lst.count() < i + 1 {
            return Err(raise("vla-redim-missing-bound", &[("name", &nm)]));
        }
        Ok(format!(
            "ReDim {pres}{nm}({})",
            self.bound_list(lst, i + 1)?
        ))
    }

    /// C2: `(type Name member...)`, module-level only.
    pub(super) fn emit_type_def(&self, lst: &List, vis: &str) -> Result<String, Refusal> {
        let name = sym_name(&sym_text(lst.nth(2)?)?)?;
        let mut r = format!("{vis}Type {name}\r\n");
        for i in 3..=lst.count() {
            let m = lst.nth(i)?;
            match m {
                Form::List(tm) => r.push_str(&format!("    {}\r\n", self.dim_spec_of(tm, 1)?)),
                _ => r.push_str(&format!(
                    "    {} As Variant\r\n",
                    sym_name(&m.atom_text().unwrap_or_default())?
                )),
            }
        }
        if lst.count() < 3 {
            return Err(raise("vla-type-no-members", &[("name", &name)]));
        }
        r.push_str("End Type\r\n");
        Ok(r)
    }

    /// C2: `(enum Name member...)` where a member is `name | (name value)`.
    pub(super) fn emit_enum_def(&self, lst: &List, vis: &str) -> Result<String, Refusal> {
        let name = sym_name(&sym_text(lst.nth(2)?)?)?;
        let mut r = format!("{vis}Enum {name}\r\n");
        for i in 3..=lst.count() {
            let m = lst.nth(i)?;
            match m {
                Form::List(ml) => r.push_str(&format!(
                    "    {} = {}\r\n",
                    sym_name(&sym_text(ml.nth(1)?)?)?,
                    self.emit_expr(ml.nth(2)?)?
                )),
                _ => r.push_str(&format!(
                    "    {}\r\n",
                    sym_name(&m.atom_text().unwrap_or_default())?
                )),
            }
        }
        if lst.count() < 3 {
            return Err(raise("vla-enum-no-members", &[("name", &name)]));
        }
        r.push_str("End Enum\r\n");
        Ok(r)
    }

    /// C2: the visibility wrapper's dispatch.
    pub(super) fn emit_visibility(&mut self, lst: &List, vis: &str) -> Result<String, Refusal> {
        let inner = lst.nth_list(2)?;
        let ih = fold(&inner.head_sym()?);
        match ih.as_str() {
            "sub" => self.emit_proc(inner, false, vis),
            "function" => self.emit_proc(inner, true, vis),
            "dim" => Ok(format!("{vis} {}\r\n", self.dim_spec_of(inner, 2)?)),
            "const" => Ok(format!("{vis} {}\r\n", self.emit_const_core(inner)?)),
            "type" => self.emit_type_def(inner, &format!("{vis} ")),
            "enum" => self.emit_enum_def(inner, &format!("{vis} ")),
            _ => Err(raise(
                "vla-visibility-wraps-unknown",
                &[("vis", &fold(vis)), ("head", &ih)],
            )),
        }
    }

    /// `(const name value)` or `(const name type value)`.
    pub(super) fn emit_const_core(&self, lst: &List) -> Result<String, Refusal> {
        if lst.count() == 3 {
            Ok(format!(
                "Const {} = {}",
                sym_name(&sym_text(lst.nth(2)?)?)?,
                self.emit_expr(lst.nth(3)?)?
            ))
        } else {
            Ok(format!(
                "Const {} As {} = {}",
                sym_name(&sym_text(lst.nth(2)?)?)?,
                sym_name(&sym_text(lst.nth(3)?)?)?,
                self.emit_expr(lst.nth(4)?)?
            ))
        }
    }

    /// L11: `(doc macro-name)`, the docstring baked into a `Debug.Print`.
    fn emit_doc(&self, lst: &List, pad: &str) -> Result<String, Refusal> {
        if lst.count() < 2 {
            return Err(raise("vla-doc-arity", &[]));
        }
        let nm = sym_text(lst.nth(2)?)?;
        let msg = match self.expander.get_macro(&nm) {
            None => format!("{nm}: (not a known macro in this parse - vocabulary macros ride loads and translations, not raw scratches)"),
            Some(m) if m.doc.is_empty() => format!("{nm}: (no documentation)"),
            Some(m) => format!("{nm}: {}", m.doc),
        };
        Ok(format!(
            "{pad}Debug.Print \"{}\"\r\n",
            msg.replace('"', "\"\"")
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::super::compile;
    use crate::messages::Refusal;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");

    /// The body of `(sub t () ...)`, between its header and `End Sub`.
    fn body(stmts: &str) -> String {
        let text = compile(&format!("(sub t () {stmts})"), PRELUDE)
            .unwrap_or_else(|e| panic!("{stmts}: {e}"));
        let start = text.find("Public Sub t()\r\n").unwrap() + "Public Sub t()\r\n".len();
        let end = text.rfind("End Sub\r\n").unwrap();
        text[start..end].to_string()
    }

    fn err(stmts: &str) -> Refusal {
        compile(&format!("(sub t () {stmts})"), PRELUDE).unwrap_err()
    }

    #[test]
    fn assignments_and_simple_statements() {
        assert_eq!(body("(set! x 1)"), "    x = 1 ' vla:1\r\n");
        // SEC.15: the formula member is a call to the runtime's sink.
        assert_eq!(
            body("(set! (. (range \"A1\") formula) \"=1\")"),
            "    Call VlaSetFormula(range(\"A1\"), \"=1\") ' vla:1\r\n"
        );
        assert_eq!(
            body("(set! (. (range \"A1\") value) 2)"),
            "    range(\"A1\").value = 2 ' vla:1\r\n"
        );
        assert_eq!(
            body("(obj-set! x (new Collection))"),
            "    Set x = New Collection ' vla:1\r\n"
        );
        assert_eq!(body("(dim x) (dim y Long) (dim z (array Long 10))"),
            "    Dim x As Variant ' vla:1\r\n    Dim y As Long ' vla:1\r\n    Dim z(10) As Long ' vla:1\r\n");
        assert_eq!(
            body("(const k 5) (const s String \"x\")"),
            "    Const k = 5 ' vla:1\r\n    Const s As String = \"x\" ' vla:1\r\n"
        );
        assert_eq!(body("(exit-sub) (exit-for) (exit-do) (return)"),
            "    Exit Sub ' vla:1\r\n    Exit For ' vla:1\r\n    Exit Do ' vla:1\r\n    Exit Sub ' vla:1\r\n");
        assert_eq!(
            body("(redim preserve a 5) (redim b (to 1 5) 3)"),
            "    ReDim Preserve a(5) ' vla:1\r\n    ReDim b(1 To 5, 3) ' vla:1\r\n"
        );
        assert_eq!(body("(on-error resume-next) (on-error goto fail) (label fail) (goto fail) (resume) (resume next) (resume fail)"),
            "    On Error Resume Next ' vla:1\r\n    On Error GoTo fail ' vla:1\r\nfail: ' vla:1\r\n    GoTo fail ' vla:1\r\n    Resume ' vla:1\r\n    Resume Next ' vla:1\r\n    Resume fail ' vla:1\r\n");
        assert_eq!(
            body("(debug-print) (debug-print 1 \"a\")"),
            "    Debug.Print ' vla:1\r\n    Debug.Print 1; \"a\" ' vla:1\r\n"
        );
        assert_eq!(body("(raw \"' as is\")"), "    ' as is\r\n");
        assert_eq!(body("(call f 1) (f) (. (range \"A1\") clearcontents)"),
            "    Call f(1) ' vla:1\r\n    Call f ' vla:1\r\n    Call range(\"A1\").clearcontents ' vla:1\r\n");
        assert!(body("(doc nosuch)").starts_with("    Debug.Print \"nosuch: (not a known macro"));
        let when = body("(doc when)");
        assert!(when.starts_with("    Debug.Print \"when: "), "{when}");
        assert!(!when.contains("not a known macro"), "{when}");
    }

    #[test]
    fn blocks_and_their_indentation() {
        assert_eq!(
            body("(if (> x 1) (then (set! y 2)) (elseif (= x 1) (set! y 1)) (else (set! y 0)))"),
            "    If (x > 1) Then ' vla:1\r\n        y = 2 ' vla:1\r\n    ElseIf (x = 1) Then\r\n        y = 1 ' vla:1\r\n    Else\r\n        y = 0 ' vla:1\r\n    End If\r\n"
        );
        assert_eq!(
            body("(for (i 1 10 2) (debug-print i))"),
            "    For i = 1 To 10 Step 2 ' vla:1\r\n        Debug.Print i ' vla:1\r\n    Next i\r\n"
        );
        assert_eq!(
            body("(for-each (c (range \"A1:A3\")) (debug-print c))"),
            "    For Each c In range(\"A1:A3\") ' vla:1\r\n        Debug.Print c ' vla:1\r\n    Next c\r\n"
        );
        assert_eq!(
            body("(while (< i 3) (set! i (+ i 1)))"),
            "    Do While (i < 3) ' vla:1\r\n        i = (i + 1) ' vla:1\r\n    Loop\r\n"
        );
        assert_eq!(
            body("(do-until done (set! done true))"),
            "    Do Until done ' vla:1\r\n        done = True ' vla:1\r\n    Loop\r\n"
        );
        assert_eq!(
            body("(with (range \"A1\") (set! .value 1))"),
            "    With range(\"A1\") ' vla:1\r\n        .value = 1 ' vla:1\r\n    End With\r\n"
        );
        assert_eq!(
            body("(select x (case (1 2) (debug-print \"a\")) (case-else (debug-print \"b\")))"),
            "    Select Case x ' vla:1\r\n        Case 1, 2\r\n            Debug.Print \"a\" ' vla:1\r\n        Case Else\r\n            Debug.Print \"b\" ' vla:1\r\n    End Select\r\n"
        );
        assert_eq!(
            body("(for-each-row (row (range \"A1:B3\")) (set! (row 2) (row 1)))"),
            "    Dim vlaSlabRange1 As Range ' vla:1\r\n    Set vlaSlabRange1 = range(\"A1:B3\")\r\n    Dim vlaSlabArr1 As Variant\r\n    vlaSlabArr1 = VlaSlabRead(vlaSlabRange1)\r\n    Dim vlaSlabI1 As Long\r\n    For vlaSlabI1 = LBound(vlaSlabArr1, 1) To UBound(vlaSlabArr1, 1)\r\n        vlaSlabArr1(vlaSlabI1, 2) = vlaSlabArr1(vlaSlabI1, 1) ' vla:1\r\n    Next vlaSlabI1\r\n    VlaSlabWrite vlaSlabArr1, vlaSlabRange1\r\n"
        );
    }

    #[test]
    fn provenance_tags() {
        assert_eq!(
            body("(at-line 90 (set! x 1) (set! y 2))"),
            "    x = 1 ' vla:1 src:90\r\n    y = 2 ' vla:1 src:90\r\n"
        );
        assert_eq!(
            body("(gen-row \"r7\" (set! x 1))"),
            "    x = 1 ' vla:1 vla-row:r7\r\n"
        );
        assert_eq!(err("(at-line x (set! y 1))").id, "vla-at-line-not-a-number");
        assert_eq!(err("(at-line 3)").id, "vla-at-line-arity");
        assert_eq!(err("(then 1)").id, "vla-clause-outside-parent");
        assert_eq!(err("(+ 1 2)").id, "vla-operator-in-statement-position");
        assert_eq!(err("5").id, "vla-bare-atom-statement");
        assert_eq!(err("(quote (1 2))").id, "vla-quote-in-statement-position");
        assert_eq!(
            err("(type T (x Long))").id,
            "vla-type-enum-module-level-only"
        );
        assert_eq!(err("(return 1)").id, "vla-return-outside-function");
        assert_eq!(err("(if x (what 1))").id, "vla-if-bad-clause");
        assert_eq!(err("(select x (other 1))").id, "vla-select-bad-clause");
        assert_eq!(err("(dim x (list 1))").id, "vla-decl-bad-type");
        assert_eq!(err("(redim a)").id, "vla-redim-missing-bound");
        assert_eq!(err("(on-error sideways)").id, "vla-on-error-bad-shape");
    }
}
