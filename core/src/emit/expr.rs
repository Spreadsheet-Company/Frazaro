//! Expressions: VLA.bas's `EmitExpr`, `EmitChain`, `OpDisplay`,
//! `EmitDotText`, `EmitQuote`, `QuoteDatum`, `EmitArgs`, `KwMisuseMsg`,
//! `EmitInterpolateCall`, `EmitParseInterpolateTemplate` and
//! `RefuseInterpreterOnlyCall`.

use super::names::{sym_name, to_vba_string};
use super::Compiler;
use crate::form::{sym_text, Form, List};
use crate::headtable::resolve_head_alias;
use crate::intrinsics::{fold, is_numeric};
use crate::messages::{raise, Refusal};

impl Compiler {
    /// VBA `EmitExpr`.
    pub(super) fn emit_expr(&self, x: &Form) -> Result<String, Refusal> {
        let lst = match x {
            Form::Str(s) => return Ok(to_vba_string(s)),
            Form::Sym(s) => {
                // P.L4 guard: a keyword token standing where a value belongs.
                if x.is_keyword_arg() {
                    return Err(raise(
                        "vla-keyword-misuse",
                        &[("msg", &kw_misuse_msg(s, "a value on its own")?)],
                    ));
                }
                return Ok(match fold(s).as_str() {
                    "true" => "True".to_string(),
                    "false" => "False".to_string(),
                    "nothing" => "Nothing".to_string(),
                    "null" => "Null".to_string(),
                    "empty" => "Empty".to_string(),
                    _ => sym_name(s)?,
                });
            }
            Form::List(l) => l,
        };
        let h = lst.head_sym()?;
        let hl = resolve_head_alias(&fold(&h));

        // PF.4c: (row i) inside an active for-each-row's own body.
        if !self.slab_row_var.is_empty() && fold(&h) == self.slab_row_var && lst.count() == 2 {
            return Ok(format!(
                "{}({}, {})",
                self.slab_arr_var,
                self.slab_i_var,
                self.emit_expr(lst.nth(2)?)?
            ));
        }

        match hl.as_str() {
            "+" | "*" | "&" | "/" | "\\" | "=" | "<>" | "<" | ">" | "<=" | ">=" | "and" | "or"
            | "xor" | "mod" | "is" | "like" | "imp" | "eqv" => {
                self.emit_chain(lst, op_display(&hl))
            }
            "-" => {
                if lst.count() == 2 {
                    Ok(format!("(-{})", self.emit_expr(lst.nth(2)?)?))
                } else {
                    self.emit_chain(lst, "-")
                }
            }
            "not" => Ok(format!("(Not {})", self.emit_expr(lst.nth(2)?)?)),
            "quote" => emit_quote(lst),
            "include" => Err(raise("vla-include-must-stand-alone", &[])),
            "new" => Ok(format!("New {}", sym_name(&sym_text(lst.nth(2)?)?)?)),
            "array" => Ok(format!("Array({})", self.emit_args(lst, 2)?)),
            "interpolate" => self.emit_interpolate_call(lst),
            "." => self.emit_dot_text(lst),
            _ => {
                // A function call, an array index or a default property:
                // all one syntax in VBA.
                if Form::sym(&h).is_keyword_arg() {
                    return Err(raise(
                        "vla-keyword-misuse",
                        &[("msg", &kw_misuse_msg(&h, "the head of an expression")?)],
                    ));
                }
                refuse_interpreter_only_call(&fold(&h))?;
                Ok(format!("{}({})", sym_name(&h)?, self.emit_args(lst, 2)?))
            }
        }
    }

    /// VBA `EmitChain`: `(op a b c)` to `(a op b op c)`.
    pub(super) fn emit_chain(&self, lst: &List, op_text: &str) -> Result<String, Refusal> {
        if lst.count() < 3 {
            return Err(raise("vla-operator-needs-two-operands", &[("op", op_text)]));
        }
        let mut r = format!("({}", self.emit_expr(lst.nth(2)?)?);
        for i in 3..=lst.count() {
            r.push_str(&format!(" {op_text} {}", self.emit_expr(lst.nth(i)?)?));
        }
        r.push(')');
        Ok(r)
    }

    /// VBA `EmitDotText`: `(. obj member args...)`, shared by the expression
    /// and statement forms.
    pub(super) fn emit_dot_text(&self, lst: &List) -> Result<String, Refusal> {
        if lst.count() < 3 {
            return Err(raise("vla-dot-needs-object-member", &[]));
        }
        let obj = self.emit_expr(lst.nth(2)?)?;
        let mem_tok = sym_text(lst.nth(3)?)?;
        if Form::sym(&mem_tok).is_keyword_arg() {
            return Err(raise(
                "vla-keyword-misuse",
                &[("msg", &kw_misuse_msg(&mem_tok, "a member name after '.'")?)],
            ));
        }
        let mem = sym_name(&mem_tok)?;
        let args = self.emit_args(lst, 4)?;
        Ok(if args.is_empty() {
            format!("{obj}.{mem}")
        } else {
            format!("{obj}.{mem}({args})")
        })
    }

    /// VBA `EmitArgs`: arguments from `from_idx`, `:keyword value` pairs
    /// becoming `Named:=arguments`.
    pub(super) fn emit_args(&self, lst: &List, from_idx: usize) -> Result<String, Refusal> {
        let mut parts: Vec<String> = Vec::new();
        let mut i = from_idx;
        while i <= lst.count() {
            let a = lst.nth(i)?;
            if a.is_keyword_arg() {
                let tok = a.atom_text().unwrap_or_default();
                if i + 1 > lst.count() {
                    return Err(raise("vla-keyword-arg-missing-value", &[("tok", &tok)]));
                }
                parts.push(format!(
                    "{}:={}",
                    sym_name(&tok[1..])?,
                    self.emit_expr(lst.nth(i + 1)?)?
                ));
                i += 2;
            } else {
                parts.push(self.emit_expr(a)?);
                i += 1;
            }
        }
        Ok(parts.join(", "))
    }

    /// VBA `EmitInterpolateCall`: `(interpolate "text {hole}" :hole value
    /// ...)` to a `&`-chain of literals and the values, each value's text
    /// computed once.
    fn emit_interpolate_call(&self, lst: &List) -> Result<String, Refusal> {
        let tpl = match lst.nth(2)? {
            Form::List(_) => {
                return Err(raise(
                    "vla-interpolate-template-must-be-literal",
                    &[("got", "a nested form")],
                ))
            }
            Form::Sym(s) => {
                return Err(raise(
                    "vla-interpolate-template-must-be-literal",
                    &[("got", s)],
                ))
            }
            Form::Str(s) => s,
        };
        let pieces = parse_interpolate_template(tpl)?;

        struct Key {
            raw: String,
            folded: String,
            text: String,
            used: bool,
        }
        let mut keys: Vec<Key> = Vec::new();
        let mut i = 3;
        while i <= lst.count() {
            let tok = lst.nth(i)?;
            if !tok.is_keyword_arg() {
                return Err(raise(
                    "vla-interpolate-expected-keyword-arg",
                    &[("pos", &(i - 2).to_string())],
                ));
            }
            let tok_text = tok.atom_text().unwrap_or_default();
            if i + 1 > lst.count() {
                return Err(raise(
                    "vla-keyword-arg-missing-value",
                    &[("tok", &tok_text)],
                ));
            }
            let raw = tok_text[1..].to_string();
            let folded = fold(&raw);
            let text = self.emit_expr(lst.nth(i + 1)?)?;
            match keys.iter_mut().find(|k| k.folded == folded) {
                Some(k) => {
                    k.raw = raw;
                    k.text = text;
                }
                None => keys.push(Key {
                    raw,
                    folded,
                    text,
                    used: false,
                }),
            }
            i += 2;
        }

        let mut parts: Vec<String> = Vec::with_capacity(pieces.len());
        for p in &pieces {
            match p {
                Piece::Hole(name) => {
                    let folded = fold(name);
                    let Some(k) = keys.iter_mut().find(|k| k.folded == folded) else {
                        return Err(raise("vla-interpolate-unknown-key", &[("key", name)]));
                    };
                    k.used = true;
                    parts.push(k.text.clone());
                }
                Piece::Literal(text) => parts.push(to_vba_string(text)),
            }
        }
        let r = if parts.is_empty() {
            "\"\"".to_string()
        } else {
            parts.join(" & ")
        };
        if let Some(k) = keys.iter().find(|k| !k.used) {
            return Err(raise("vla-interpolate-unused-argument", &[("key", &k.raw)]));
        }
        Ok(r)
    }
}

/// VBA `OpDisplay`: the word operators in VBA's capitalisation.
pub fn op_display(hl: &str) -> &str {
    match hl {
        "and" => "And",
        "or" => "Or",
        "xor" => "Xor",
        "mod" => "Mod",
        "is" => "Is",
        "like" => "Like",
        "imp" => "Imp",
        "eqv" => "Eqv",
        _ => hl,
    }
}

/// VBA `RefuseInterpreterOnlyCall` (OPTIMIZE.3 slice 5).
pub fn refuse_interpreter_only_call(folded_head: &str) -> Result<(), Refusal> {
    if folded_head == "vlaoptimizecell" {
        return Err(raise("vla-interpreter-only-command", &[]));
    }
    Ok(())
}

/// VBA `EmitQuote`: `(quote datum)` to one VBA literal expression.
pub fn emit_quote(lst: &List) -> Result<String, Refusal> {
    if lst.count() != 2 {
        return Err(raise("vla-quote-arity", &[]));
    }
    Ok(quote_datum(lst.nth(2)?))
}

/// VBA `QuoteDatum`: a list becomes `Array(...)`; a string literal stays a
/// string, a numeric token a number, every other atom its text as a string.
pub fn quote_datum(v: &Form) -> String {
    match v {
        Form::List(lst) => {
            let parts: Vec<String> = lst.items.iter().map(quote_datum).collect();
            format!("Array({})", parts.join(", "))
        }
        Form::Str(s) => to_vba_string(s),
        Form::Sym(s) => {
            if is_numeric(s) {
                s.clone()
            } else {
                format!("\"{s}\"")
            }
        }
    }
}

/// VBA `KwMisuseMsg`.
pub fn kw_misuse_msg(tok: &str, where_: &str) -> Result<String, Refusal> {
    Ok(format!(
        "'{tok}' is a keyword argument token - inside a call's argument list it pairs with the value after it ((f {tok} x) -> f {}:=x); it cannot be {where_}",
        sym_name(&tok[1..])?
    ))
}

/// One piece of an interpolate template.
#[derive(Debug, PartialEq)]
enum Piece {
    Literal(String),
    Hole(String),
}

/// VBA `EmitParseInterpolateTemplate`: `{{` is a literal `{`, `{name}` a
/// hole, a `}` outside a hole literal (doubled or not).
fn parse_interpolate_template(tpl: &str) -> Result<Vec<Piece>, Refusal> {
    let chars: Vec<char> = tpl.chars().collect();
    let n = chars.len();
    let mut out = Vec::new();
    let mut lit = String::new();
    let mut i = 0;
    while i < n {
        let c = chars[i];
        if c == '{' {
            if chars.get(i + 1) == Some(&'{') {
                lit.push('{');
                i += 2;
            } else {
                let Some(close_rel) = chars[i + 1..].iter().position(|c| *c == '}') else {
                    return Err(raise("vla-interpolate-unclosed-hole", &[]));
                };
                let close_at = i + 1 + close_rel;
                let hole_name: String = chars[i + 1..close_at].iter().collect();
                if hole_name.is_empty() {
                    return Err(raise("vla-interpolate-empty-hole-name", &[]));
                }
                if !lit.is_empty() {
                    out.push(Piece::Literal(std::mem::take(&mut lit)));
                }
                out.push(Piece::Hole(hole_name));
                i = close_at + 1;
            }
        } else if c == '}' {
            lit.push('}');
            i += if chars.get(i + 1) == Some(&'}') { 2 } else { 1 };
        } else {
            lit.push(c);
            i += 1;
        }
    }
    if !lit.is_empty() {
        out.push(Piece::Literal(lit));
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_forms;

    fn expr(src: &str) -> String {
        Compiler::new()
            .emit_expr(&read_forms(src).unwrap()[0])
            .unwrap_or_else(|e| panic!("{src}: {e}"))
    }

    fn err(src: &str) -> Refusal {
        Compiler::new()
            .emit_expr(&read_forms(src).unwrap()[0])
            .unwrap_err()
    }

    #[test]
    fn atoms_operators_and_calls() {
        assert_eq!(expr("\"a \\\"b\\\"\""), "\"a \"\"b\"\"\"");
        assert_eq!(expr("True"), "True");
        assert_eq!(expr("nothing"), "Nothing");
        assert_eq!(expr("hot-pink"), "hot_pink");
        assert_eq!(expr("(+ 1 2 3)"), "(1 + 2 + 3)");
        assert_eq!(expr("(- x)"), "(-x)");
        assert_eq!(expr("(- x 1)"), "(x - 1)");
        assert_eq!(expr("(and a (not b))"), "(a And (Not b))");
        assert_eq!(expr("(mod a 2)"), "(a Mod 2)");
        assert_eq!(expr("(f 1 :key 2)"), "f(1, key:=2)");
        assert_eq!(
            expr("(. (range \"A1\") font.bold)"),
            "range(\"A1\").font.bold"
        );
        assert_eq!(
            expr("(. (range \"A1\") copy :destination x)"),
            "range(\"A1\").copy(destination:=x)"
        );
        assert_eq!(expr("(new Collection)"), "New Collection");
        assert_eq!(expr("(array 1 \"b\")"), "Array(1, \"b\")");
        assert_eq!(
            expr("(quote (hot-pink 1 \"s\" :k))"),
            "Array(\"hot-pink\", 1, \"s\", \":k\")"
        );
        assert_eq!(expr("(quote x)"), "\"x\"");
        assert_eq!(err("(+ 1)").id, "vla-operator-needs-two-operands");
        assert_eq!(err(":key").id, "vla-keyword-misuse");
        assert!(err(":key")
            .text
            .contains("(f :key x) -> f key:=x); it cannot be a value on its own"));
        assert_eq!(err("(f :key)").id, "vla-keyword-arg-missing-value");
        assert_eq!(err("(. x :m)").id, "vla-keyword-misuse");
        assert_eq!(err("(. x)").id, "vla-dot-needs-object-member");
        assert_eq!(
            err("(vlaoptimizecell 1)").id,
            "vla-interpreter-only-command"
        );
        assert_eq!(err("(include \"x\")").id, "vla-include-must-stand-alone");
    }

    #[test]
    fn interpolation() {
        assert_eq!(
            expr("(interpolate \"Hello {name}, {{x}} and }\" :name who)"),
            "\"Hello \" & who & \", {x} and }\""
        );
        assert_eq!(expr("(interpolate \"{a}{a}\" :a 1 :A 2)"), "2 & 2");
        assert_eq!(expr("(interpolate \"\")"), "\"\"");
        assert_eq!(
            err("(interpolate x)").id,
            "vla-interpolate-template-must-be-literal"
        );
        assert_eq!(
            err("(interpolate \"{a}\" b 1)").id,
            "vla-interpolate-expected-keyword-arg"
        );
        assert_eq!(
            err("(interpolate \"{a}\" :b 1)").id,
            "vla-interpolate-unknown-key"
        );
        assert_eq!(
            err("(interpolate \"x\" :b 1)").id,
            "vla-interpolate-unused-argument"
        );
        assert_eq!(
            err("(interpolate \"{a\")").id,
            "vla-interpolate-unclosed-hole"
        );
        assert_eq!(
            err("(interpolate \"{}\")").id,
            "vla-interpolate-empty-hole-name"
        );
        assert_eq!(
            parse_interpolate_template("a{b}c").unwrap(),
            vec![
                Piece::Literal("a".into()),
                Piece::Hole("b".into()),
                Piece::Literal("c".into())
            ]
        );
    }
}
