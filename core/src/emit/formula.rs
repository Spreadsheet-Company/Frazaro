//! The formula dialect: VLA.bas's `EmitDeflambda`, `EmitFormula`,
//! `FormulaQuote` and `FormulaText`, with F.7's two fixes (a lone operand
//! under `-` keeps its sign; a quote mark inside a text is doubled).

use super::names::{sym_name, to_vba_string};
use super::Compiler;
use crate::form::{sym_text, Form, List};
use crate::headtable::{self, resolve_head_alias, FormulaSupport};
use crate::intrinsics::{fold, is_numeric};
use crate::messages::{raise, Refusal};

impl Compiler {
    /// VBA `EmitDeflambda`: `(deflambda name (params) ["doc"] body-expr)`
    /// registers an Excel LAMBDA under a workbook Name, in one statement.
    pub(super) fn emit_deflambda(&mut self, lst: &List) -> Result<String, Refusal> {
        if lst.count() < 4 {
            return Err(raise("vla-deflambda-arity", &[]));
        }
        let name_tok = sym_text(lst.nth(2)?)?;
        let params = lst.nth_list(3)?;
        let mut body_idx = 4;
        let mut ldoc = String::new();
        if lst.count() >= 5 {
            if let Form::Str(s) = lst.nth(4)? {
                ldoc = s.clone();
                body_idx = 5;
            }
        }
        if lst.count() != body_idx {
            return Err(raise(
                "vla-deflambda-body-not-one-formula",
                &[
                    ("name", &name_tok),
                    ("n", &(lst.count() - body_idx + 1).to_string()),
                ],
            ));
        }
        let mut plist: Vec<String> = Vec::with_capacity(params.count());
        for e in &params.items {
            if e.is_list() {
                return Err(raise(
                    "vla-deflambda-params-not-plain",
                    &[("name", &name_tok)],
                ));
            }
            plist.push(sym_name(&e.atom_text().unwrap_or_default())?);
        }
        let plist = plist.join(", ");
        let formula = format!(
            "=LAMBDA({plist}{}{})",
            if plist.is_empty() { "" } else { ", " },
            emit_formula(lst.nth(body_idx)?)?
        );
        self.record_sub_doc(&name_tok, &ldoc, "lambda");
        let add_part = "ThisWorkbook.Names.Add";
        let arg_part = format!(
            "Name:=\"{}\", RefersTo:=\"{}\"",
            sym_name(&name_tok)?,
            formula.replace('"', "\"\"")
        );
        Ok(if ldoc.is_empty() {
            format!("{add_part} {arg_part}")
        } else {
            format!("{add_part}({arg_part}).Comment = {}", to_vba_string(&ldoc))
        })
    }
}

/// VBA `EmitFormula`: a VLA expression to Excel formula text, with real
/// quote characters; the caller VBA-escapes once at the Add line.
pub fn emit_formula(v: &Form) -> Result<String, Refusal> {
    let lst = match v {
        Form::Str(s) => return Ok(formula_text(s)),
        Form::Sym(s) => {
            if v.is_keyword_arg() {
                return Err(raise("vla-formula-no-named-args", &[("tok", s)]));
            }
            return Ok(match fold(s).as_str() {
                "true" => "TRUE".to_string(),
                "false" => "FALSE".to_string(),
                _ => sym_name(s)?,
            });
        }
        Form::List(l) => l,
    };
    let h = resolve_head_alias(&fold(&lst.head_sym()?));
    match h.as_str() {
        "+" | "-" | "*" | "/" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=" => {
            // F.7: a lone operand under - is a negation and keeps its sign.
            if h == "-" && lst.count() == 2 {
                return Ok(format!("(-{})", emit_formula(lst.nth(2)?)?));
            }
            let mut r = String::new();
            for i in 2..=lst.count() {
                if !r.is_empty() {
                    r.push_str(&h);
                }
                r.push_str(&emit_formula(lst.nth(i)?)?);
            }
            Ok(format!("({r})"))
        }
        "mod" => {
            if lst.count() != 3 {
                return Err(raise("vla-formula-mod-arity", &[]));
            }
            Ok(format!(
                "MOD({}, {})",
                emit_formula(lst.nth(2)?)?,
                emit_formula(lst.nth(3)?)?
            ))
        }
        "not" => Ok(format!("NOT({})", emit_formula(lst.nth(2)?)?)),
        "and" | "or" => {
            let mut parts = Vec::with_capacity(lst.count());
            for i in 2..=lst.count() {
                parts.push(emit_formula(lst.nth(i)?)?);
            }
            Ok(format!("{}({})", h.to_uppercase(), parts.join(", ")))
        }
        "if" => {
            if lst.count() >= 3 {
                if let Form::List(then_form) = lst.nth(3)? {
                    if let Some(Form::Sym(th)) = then_form.items.first() {
                        if fold(th) == "then" {
                            return Err(raise("vla-formula-if-then-block", &[]));
                        }
                    }
                }
            }
            if lst.count() < 3 || lst.count() > 4 {
                return Err(raise("vla-formula-if-arity", &[]));
            }
            let mut r = format!(
                "IF({}, {}",
                emit_formula(lst.nth(2)?)?,
                emit_formula(lst.nth(3)?)?
            );
            if lst.count() == 4 {
                r.push_str(&format!(", {}", emit_formula(lst.nth(4)?)?));
            }
            r.push(')');
            Ok(r)
        }
        "quote" => {
            if lst.count() != 2 {
                return Err(raise("vla-quote-arity", &[]));
            }
            formula_quote(lst.nth(2)?)
        }
        "lambda" => {
            // L14.1: an inner lambda for REDUCE/SCAN/MAP.
            if lst.count() != 3 {
                return Err(raise("vla-formula-lambda-arity", &[]));
            }
            let lp = lst.nth_list(2)?;
            let mut parts: Vec<String> = Vec::with_capacity(lp.count());
            for pe in &lp.items {
                if pe.is_list() {
                    return Err(raise("vla-formula-lambda-params-not-plain", &[]));
                }
                parts.push(sym_name(&pe.atom_text().unwrap_or_default())?);
            }
            let r = parts.join(", ");
            Ok(format!(
                "LAMBDA({r}{}{})",
                if r.is_empty() { "" } else { ", " },
                emit_formula(lst.nth(3)?)?
            ))
        }
        "set!" | "obj-set!" | "dim" | "begin" | "while" | "for" | "for-each" | "for-each-row"
        | "debug-print" | "on-error" | "label" | "goto" | "deflambda" | "include" => {
            Err(raise("vla-formula-is-statement", &[("head", &h)]))
        }
        _ => {
            // VLA.4 (F.7): a head the head table declares no formula's is
            // refused by name, where this arm used to write it as a call,
            // `(new x)` as `new(x)`.
            if headtable::row(&h).is_some_and(|r| r.formula != FormulaSupport::Yes) {
                return Err(raise(
                    "vla-formula-head-refused",
                    &[("head", &lst.head_sym()?)],
                ));
            }
            let mut parts = Vec::with_capacity(lst.count());
            for i in 2..=lst.count() {
                parts.push(emit_formula(lst.nth(i)?)?);
            }
            Ok(format!(
                "{}({})",
                sym_name(&lst.head_sym()?)?,
                parts.join(", ")
            ))
        }
    }
}

/// VBA `FormulaQuote`: data as an Excel array constant; strings stay
/// strings, numbers stay numbers, every other atom is its text as written,
/// as a string.
pub fn formula_quote(v: &Form) -> Result<String, Refusal> {
    match v {
        Form::List(lst) => {
            let mut parts = Vec::with_capacity(lst.count());
            for e in &lst.items {
                if e.is_list() {
                    return Err(raise("vla-formula-array-no-nest", &[]));
                }
                parts.push(formula_quote(e)?);
            }
            Ok(format!("{{{}}}", parts.join(",")))
        }
        Form::Str(s) => Ok(formula_text(s)),
        Form::Sym(s) => Ok(if is_numeric(s) {
            s.clone()
        } else {
            formula_text(s)
        }),
    }
}

/// F.7's `FormulaText`: a text as Excel reads it, each inner quote mark
/// doubled.
pub fn formula_text(content: &str) -> String {
    format!("\"{}\"", content.replace('"', "\"\""))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_forms;

    fn formula(src: &str) -> String {
        emit_formula(&read_forms(src).unwrap()[0]).unwrap()
    }

    #[test]
    fn the_dialect_renders_as_l14_pinned() {
        assert_eq!(formula("(* x 2)"), "(x*2)");
        assert_eq!(formula("(> a b)"), "(a>b)");
        assert_eq!(
            formula("(if (= (mod n 2) 0) \"even\" \"odd\")"),
            "IF((MOD(n, 2)=0), \"even\", \"odd\")"
        );
        assert_eq!(formula("(quote (1 2 3))"), "{1,2,3}");
        assert_eq!(
            formula("(and a (not b) (or c d))"),
            "AND(a, NOT(b), OR(c, d))"
        );
        assert_eq!(
            formula("(sum (lambda (x) (+ x 1)) rng)"),
            "sum(LAMBDA(x, (x+1)), rng)"
        );
        assert_eq!(formula("true"), "TRUE");
        assert_eq!(formula("tax-rate"), "tax_rate");
    }

    #[test]
    fn f7_sign_and_quote() {
        assert_eq!(formula("(- x)"), "(-x)");
        assert_eq!(formula("(- (- x) 1)"), "((-x)-1)");
        assert_eq!(
            formula("(if x \"say \\\"hi\\\"\" \"nothing\")"),
            "IF(x, \"say \"\"hi\"\"\", \"nothing\")"
        );
        assert_eq!(
            formula("(quote (\"a \\\"b\\\"\" 2 sym))"),
            "{\"a \"\"b\"\"\",2,\"sym\"}"
        );
    }

    #[test]
    fn the_refusals() {
        let err = |src: &str| emit_formula(&read_forms(src).unwrap()[0]).unwrap_err().id;
        assert_eq!(err("(set! x 1)"), "vla-formula-is-statement");
        assert_eq!(err("(if (> x 1) (then 2))"), "vla-formula-if-then-block");
        assert_eq!(err("(if x)"), "vla-formula-if-arity");
        assert_eq!(err("(mod a b c)"), "vla-formula-mod-arity");
        assert_eq!(err("(quote ((1 2) 3))"), "vla-formula-array-no-nest");
        assert_eq!(err("(f :key 1)"), "vla-formula-no-named-args");
        assert_eq!(
            err("(lambda ((x)) x)"),
            "vla-formula-lambda-params-not-plain"
        );
        // VLA.4 (F.7): a head the table refuses, where this wrote new(x).
        assert_eq!(err("(new x)"), "vla-formula-head-refused");
        assert_eq!(err("(. a b)"), "vla-formula-head-refused");
    }

    /// F.7's test (VLA.4) on deflambda's dialect: every head the head table
    /// declares no formula's is refused by name, by the fourteen this
    /// emitter always named or by the table's refusal; and the heads it
    /// declares a formula's are written, `type` and `xor` as the calls
    /// Excel's TYPE and XOR, as they always were.
    #[test]
    fn every_refused_head_of_the_table_is_refused() {
        for row in headtable::rows() {
            assert_ne!(
                row.formula,
                FormulaSupport::Undeclared,
                "{} is undeclared",
                row.symbol
            );
            let form = Form::list(vec![Form::sym(&row.symbol), Form::sym("a"), Form::sym("b")]);
            if row.formula == FormulaSupport::Refused {
                let id = emit_formula(&form).unwrap_err().id;
                assert!(
                    id == "vla-formula-is-statement" || id == "vla-formula-head-refused",
                    "{}: {id}",
                    row.symbol
                );
            }
        }
        assert_eq!(formula("(xor a (type b))"), "xor(a, type(b))");
    }

    /// The parity of the two writers (VLA.4): where deflambda's dialect and
    /// the formula syntax both write an expression, the language's parser
    /// reads their texts as one tree. They differ in spelling only: this
    /// one parenthesizes every operation and keeps a call's case.
    #[test]
    fn the_formula_syntax_reads_as_this_dialect_does() {
        use crate::calc::formula::parse;
        for src in [
            "(* x 2)",
            "(> a b)",
            "(if (= (mod n 2) 0) \"even\" \"odd\")",
            "(and a (not b) (or c d))",
            "(- (- x) 1)",
            "(- a b c)",
            "(+ 1 (* 2 3))",
            "(* (+ 1 2) 3)",
            "(& \"say \\\"hi\\\"\" 1)",
            "(< (+ a 1) (* b 2))",
            "(sum A1 B2 (max c d))",
        ] {
            let v = &read_forms(src).unwrap()[0];
            let old = parse(&emit_formula(v).unwrap()).unwrap();
            let new = parse(&crate::formula_syntax::write(v).unwrap()).unwrap();
            assert_eq!(old, new, "{src}");
        }
    }
}
