//! The VBA emitter: VLA.bas's `VlaTranspile`, `EmitTop`, `EmitProc` and
//! the compile state they carry (PORT.5, slice 5e). The statement,
//! expression and formula emitters are the submodules, one `match` arm per
//! `Select Case` arm; the names module holds `SymName` and its kin.
//!
//! The state is `VlaFrame.cls`'s snapshot of VLA.bas's compile-time
//! fields, carried here as one [`Compiler`]: the macro table lives in its
//! [`Expander`], the head aliases are the head table's, and the rest are
//! the per-compile fields `VlaTranspile` resets at its top. Generated
//! names (`vla_tco_N`, `vlaSlabRangeN`, the button helper) come out of the
//! same counters the VBA keeps, so a compile here and a compile there write
//! the same text.
//!
//! What is not here: `SpliceIncludes`, which reads files. The core is a
//! pure function over text, so a program whose line is a whole-line
//! `(include "...")` is refused with the catalogue's own `vla-include-
//! cannot-read`, naming the file; a door that can read files splices
//! before it calls [`compile`].

mod expr;
mod formula;
mod names;
mod stmt;

use std::collections::HashMap;

use crate::expand::Expander;
use crate::form::{str_lit_content, sym_text, Form, List};
use crate::headtable::resolve_head_alias;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::reader::{count_lf, parse_all, tokenize};

pub use names::sym_name;

/// `VLA_MAKEBUTTON_SUB`: the compiled module's button helper, appended
/// once when a program used `make-button`.
pub const MAKEBUTTON_SUB: &str = "VlaCompiledMakeButton";

/// P.L6: one procedure's docstring record.
#[derive(Clone, Debug)]
struct SubDoc {
    as_written: String,
    doc: String,
    kind: String,
}

/// The compile state: VLA.bas's module-level fields, one compile at a time.
#[derive(Default)]
pub struct Compiler {
    pub expander: Expander,
    sub_docs: HashMap<String, SubDoc>,
    func_name: String,
    in_function: bool,
    tco_name: String,
    tco_params: Vec<String>,
    proc_emitted: bool,
    emit_line: u32,
    at_line: u32,
    gen_row: String,
    slab_counter: u32,
    slab_row_var: String,
    slab_arr_var: String,
    slab_i_var: String,
}

/// VBA `VlaTranspile`: the program with the prelude, read, macros defined,
/// every top-level form expanded and emitted, `Option Explicit` first and
/// the button helper last when it is used.
pub fn compile(source: &str, prelude: &str) -> Result<String, Refusal> {
    Compiler::new().transpile(source, prelude)
}

impl Compiler {
    pub fn new() -> Compiler {
        Compiler::default()
    }

    /// VBA `VlaTranspile`.
    pub fn transpile(&mut self, source: &str, prelude: &str) -> Result<String, Refusal> {
        refuse_includes(source)?;
        let line_offset = count_lf(prelude) as u32 + 1;
        self.emit_line = 0;
        self.at_line = 0;
        self.gen_row.clear();
        self.slab_counter = 0;
        self.slab_row_var.clear();
        let text = format!("{prelude}\r\n{source}");
        let forms = parse_all(&tokenize(&text), line_offset)?;
        self.expander = Expander::new();
        self.sub_docs.clear();
        self.proc_emitted = false;
        self.func_name.clear();
        self.in_function = false;

        // Pass 1: collect macro definitions (order-independent).
        let body = self.expander.collect(forms)?;

        // Pass 2: expand, then emit. A refusal names the last mapped
        // statement line the emitter reached, "near" on purpose.
        let mut out = String::from("Option Explicit\r\n\r\n");
        for f in &body {
            let expanded = match self.expander.expand(f, 0) {
                Ok(e) => e,
                Err(e) => return Err(self.emit_fail(e)),
            };
            match self.emit_top(&expanded) {
                Ok(text) => out.push_str(&text),
                Err(e) => return Err(self.emit_fail(e)),
            }
            out.push_str("\r\n");
        }
        if out.contains(&format!("Call {MAKEBUTTON_SUB}(")) {
            out.push_str("\r\n");
            out.push_str(&make_button_helper_text());
        }
        Ok(out)
    }

    /// `VlaTranspile`'s `emitfail`: ` (near vla line N)` on a refusal that
    /// does not already name a line, with the source line when an
    /// `(at-line ...)` encloses the point of failure.
    fn emit_fail(&self, e: Refusal) -> Refusal {
        if self.emit_line > 0 && !e.text.contains("vla line") {
            let suffix = format!(
                " (near vla line {}{})",
                self.emit_line,
                self.at_line_suffix()
            );
            return e.with_suffix(&suffix);
        }
        e
    }

    /// VBA `AtLineSuffix`.
    fn at_line_suffix(&self) -> String {
        if self.at_line > 0 {
            format!(", source line {}", self.at_line)
        } else {
            String::new()
        }
    }

    /// VBA `TopoGuard`: a module-level declaration after a procedure.
    fn topo_guard(&self, what: &str) -> Result<(), Refusal> {
        if self.proc_emitted {
            return Err(raise("vla-module-level-decl-after-proc", &[("what", what)]));
        }
        Ok(())
    }

    /// VBA `RecordSubDoc`: last write wins on a name.
    fn record_sub_doc(&mut self, as_written: &str, doc: &str, kind: &str) {
        self.sub_docs.insert(
            fold(as_written),
            SubDoc {
                as_written: as_written.to_string(),
                doc: doc.to_string(),
                kind: kind.to_string(),
            },
        );
    }

    /// VBA `VlaSubDoc`: a procedure's docstring, or empty.
    pub fn sub_doc(&self, name: &str) -> String {
        self.sub_docs
            .get(&fold(name))
            .map(|d| d.doc.clone())
            .unwrap_or_default()
    }

    /// VBA `SubExists`: has this compile emitted a procedure by this name?
    fn sub_exists(&self, as_written: &str) -> bool {
        self.sub_docs.contains_key(&fold(as_written))
    }

    /// The recorded procedures, as written, with their kinds: what
    /// `VlaAproposText` would list.
    pub fn procedures(&self) -> Vec<(String, String)> {
        self.sub_docs
            .values()
            .map(|d| (d.as_written.clone(), d.kind.clone()))
            .collect()
    }

    /// VBA `EmitTop`.
    fn emit_top(&mut self, f: &Form) -> Result<String, Refusal> {
        let Form::List(lst) = f else {
            return Err(raise(
                "vla-top-level-not-list",
                &[("value", &f.atom_text().unwrap_or_default())],
            ));
        };
        let h = resolve_head_alias(&fold(&lst.head_sym()?));
        match h.as_str() {
            "sub" => {
                self.proc_emitted = true; // G12.1
                self.emit_proc(lst, false, "Public")
            }
            "function" => {
                self.proc_emitted = true;
                self.emit_proc(lst, true, "Public")
            }
            "dim" => {
                self.topo_guard("dim")?;
                Ok(format!("{}\r\n", self.emit_dim_core(lst)?))
            }
            "const" => {
                self.topo_guard("const")?;
                Ok(format!("{}\r\n", self.emit_const_core(lst)?))
            }
            "raw" => Ok(format!("{}\r\n", str_lit_content(lst.nth(2)?)?)),
            "type" => self.emit_type_def(lst, ""),
            "enum" => self.emit_enum_def(lst, ""),
            "public" | "private" => {
                // C2: visibility as a wrapper form.
                let vis = if h == "public" { "Public" } else { "Private" };
                self.emit_visibility(lst, vis)
            }
            "begin" => {
                let mut out = String::new();
                for i in 2..=lst.count() {
                    out.push_str(&self.emit_top(lst.nth(i)?)?);
                }
                Ok(out)
            }
            "deflambda" => Err(raise("vla-deflambda-runs", &[])),
            "include" => Err(raise("vla-include-must-stand-alone", &[])),
            _ => Err(raise("vla-unknown-top-level-form", &[("head", &h)])),
        }
    }

    /// VBA `EmitProc`: `(sub name (params) ["doc"] body...)` and
    /// `(function name (params) [type] ["doc"] body...)`.
    fn emit_proc(&mut self, lst: &List, is_func: bool, vis: &str) -> Result<String, Refusal> {
        let raw_name = sym_text(lst.nth(2)?)?;
        // IN.7: only an on:click: handler may carry a colon into a name.
        if raw_name.contains(':') && !fold(&raw_name).starts_with("on:click:") {
            return Err(raise(
                "vla-interpreter-only-handler",
                &[("name", &raw_name)],
            ));
        }
        let name = sym_name(&raw_name)?;
        let params = lst.nth_list(3)?;

        let mut body_idx = 4;
        let mut ret_type = String::from("Variant");
        if is_func && lst.count() >= 4 {
            // P.L6: an atom after the params that is not a string literal
            // is the return type.
            if let Form::Sym(s) = lst.nth(4)? {
                ret_type = sym_name(s)?;
                body_idx = 5;
            }
        }

        // P.L6: an optional docstring, documentation only when at least
        // one more form follows.
        let mut pdoc = String::new();
        if lst.count() > body_idx {
            if let Form::Str(s) = lst.nth(body_idx)? {
                pdoc = s.clone();
                body_idx += 1;
            }
        }
        self.record_sub_doc(&raw_name, &pdoc, if is_func { "function" } else { "sub" });

        // L18: TCO, armed for a function whose every parameter is
        // (byval name type) and whose body holds a self-return.
        self.tco_name.clear();
        self.tco_params.clear();
        if is_func {
            let mut all_by_val = true;
            for pv in &params.items {
                match pv {
                    Form::List(pvl) => {
                        if fold(&pvl.head_sym()?) != "byval" {
                            all_by_val = false;
                        }
                    }
                    _ => all_by_val = false,
                }
            }
            if all_by_val && tco_scan(lst, body_idx, &name)? {
                self.tco_name = fold(&name);
                for pv in &params.items {
                    let Form::List(pvl) = pv else { unreachable!() };
                    self.tco_params.push(sym_name(&sym_text(pvl.nth(2)?)?)?);
                }
            }
        }

        self.func_name = name.clone();
        self.in_function = is_func;

        let kind = if is_func { "Function" } else { "Sub" };
        let mut r = format!("{vis} {kind} {name}({})", self.emit_params(params)?);
        if is_func {
            r.push_str(&format!(" As {ret_type}"));
        }
        r.push_str("\r\n");
        if !self.tco_name.is_empty() {
            // L18: the loop head; temps only with more than one parameter.
            if self.tco_params.len() > 1 {
                for ti in 1..=self.tco_params.len() {
                    r.push_str(&format!("    Dim vla_tco_{ti} As Variant\r\n"));
                }
            }
            r.push_str("vla_tco:\r\n");
        }
        r.push_str(&self.emit_body(lst, body_idx, 1)?);
        r.push_str(&format!("End {kind}\r\n"));
        self.tco_name.clear();
        self.tco_params.clear();
        Ok(r)
    }

    /// VBA `EmitParams`: `name | (name type) | (byval name type) |
    /// (byref name type) | (optional name type [default]) | (paramarray name)`.
    fn emit_params(&self, params: &List) -> Result<String, Refusal> {
        let mut parts: Vec<String> = Vec::with_capacity(params.count());
        for p in &params.items {
            let one = match p {
                Form::List(pl) => {
                    let ph = fold(&pl.head_sym()?);
                    match ph.as_str() {
                        "byval" => format!(
                            "ByVal {} As {}",
                            sym_name(&sym_text(pl.nth(2)?)?)?,
                            sym_name(&sym_text(pl.nth(3)?)?)?
                        ),
                        "byref" => format!(
                            "ByRef {} As {}",
                            sym_name(&sym_text(pl.nth(2)?)?)?,
                            sym_name(&sym_text(pl.nth(3)?)?)?
                        ),
                        "optional" => {
                            let mut one = format!(
                                "Optional {} As {}",
                                sym_name(&sym_text(pl.nth(2)?)?)?,
                                sym_name(&sym_text(pl.nth(3)?)?)?
                            );
                            if pl.count() >= 4 {
                                one.push_str(&format!(" = {}", self.emit_expr(pl.nth(4)?)?));
                            }
                            one
                        }
                        "paramarray" => format!(
                            "ParamArray {}() As Variant",
                            sym_name(&sym_text(pl.nth(2)?)?)?
                        ),
                        _ => format!(
                            "{} As {}",
                            sym_name(&sym_text(pl.nth(1)?)?)?,
                            sym_name(&sym_text(pl.nth(2)?)?)?
                        ),
                    }
                }
                // A bare name is ByRef Variant.
                _ => sym_name(&p.atom_text().unwrap_or_default())?,
            };
            parts.push(one);
        }
        Ok(parts.join(", "))
    }

    /// VBA `EmitBody`.
    fn emit_body(&mut self, lst: &List, from_idx: usize, ind: usize) -> Result<String, Refusal> {
        let mut out = String::new();
        for i in from_idx..=lst.count() {
            out.push_str(&self.emit_stmt(lst.nth(i)?, ind)?);
        }
        Ok(out)
    }
}

/// VBA `TcoScan`: does the body hold a `(return (self ...))`?
fn tco_scan(lst: &List, from_idx: usize, self_name: &str) -> Result<bool, Refusal> {
    for i in from_idx..=lst.count() {
        if tco_scan_form(lst.nth(i)?, self_name)? {
            return Ok(true);
        }
    }
    Ok(false)
}

/// VBA `TcoScanForm`: only a return's direct operand counts.
fn tco_scan_form(v: &Form, self_name: &str) -> Result<bool, Refusal> {
    let Form::List(lst) = v else {
        return Ok(false);
    };
    let Some(Form::Sym(head)) = lst.items.first() else {
        return Ok(false);
    };
    let h = fold(head);
    if h == "quote" {
        return Ok(false); // data, never code
    }
    if h == "return" && lst.count() >= 2 {
        if let Form::List(inner) = lst.nth(2)? {
            if let Some(Form::Sym(callee)) = inner.items.first() {
                if fold(&sym_name(callee)?) == fold(self_name) {
                    return Ok(true);
                }
            }
        }
    }
    for e in &lst.items {
        if tco_scan_form(e, self_name)? {
            return Ok(true);
        }
    }
    Ok(false)
}

/// VBA `MakeButtonHelperText`: the compiled module's own button-creation
/// helper, appended once by `VlaTranspile` when `make-button` was used.
pub fn make_button_helper_text() -> String {
    [
        format!("Private Sub {MAKEBUTTON_SUB}(ByVal vlaCaption As String, ByVal vlaPlace As Range, ByVal vlaProc As String)"),
        "    Dim vlaWs As Worksheet".to_string(),
        "    Set vlaWs = vlaPlace.Worksheet".to_string(),
        "    On Error Resume Next".to_string(),
        "    vlaWs.Buttons(vlaCaption).Delete".to_string(),
        "    On Error GoTo 0".to_string(),
        "    Dim vlaBtn As Button".to_string(),
        "    Set vlaBtn = vlaWs.Buttons.Add(vlaPlace.Left, vlaPlace.Top, 120, 24)".to_string(),
        "    vlaBtn.Caption = vlaCaption".to_string(),
        "    vlaBtn.Name = vlaCaption".to_string(),
        "    If Len(vlaProc) > 0 Then".to_string(),
        "        vlaBtn.OnAction = \"'\" & ThisWorkbook.Name & \"'!\" & vlaProc".to_string(),
        "    End If".to_string(),
        "End Sub".to_string(),
        String::new(),
    ]
    .join("\r\n")
}

/// `SpliceIncludes`' recognition of a whole-line include, `IsIncludeLine`:
/// the core cannot read the file, so it refuses by name where the VBA
/// would splice; a door that reads files splices first.
fn refuse_includes(source: &str) -> Result<(), Refusal> {
    if !fold(source).contains("(include") {
        return Ok(());
    }
    for line in source.split('\n') {
        if let Some(name) = include_line_file(line) {
            return Err(raise(
                "vla-include-cannot-read",
                &[("name", &name), ("path", &name)],
            ));
        }
    }
    Ok(())
}

/// VBA `IsIncludeLine`: `(include "name")` alone on its line.
fn include_line_file(line: &str) -> Option<String> {
    let t = line.trim();
    let head = t.get(..9)?;
    if fold(head) != "(include " || !t.ends_with(')') {
        return None;
    }
    let inner = t[9..t.len() - 1].trim();
    if inner.len() < 2 || !inner.starts_with('"') || !inner.ends_with('"') {
        return None;
    }
    let name = &inner[1..inner.len() - 1];
    if name.is_empty() || name.contains('"') {
        return None;
    }
    Some(name.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const GOLDEN_VLA: &str = include_str!("../../../scripts/instructions_golden.vla");
    const GOLDEN_VBA: &str = include_str!("../../../scripts/instructions_golden.vba");

    fn vba(src: &str) -> String {
        compile(src, PRELUDE).unwrap_or_else(|e| panic!("{src}: {e}"))
    }

    fn refusal(src: &str) -> Refusal {
        match compile(src, PRELUDE) {
            Ok(text) => panic!("{src} compiled instead of refusing: {text}"),
            Err(e) => e,
        }
    }

    #[test]
    fn a_sub_and_a_function_with_their_shapes() {
        assert_eq!(
            vba("(sub t () (set! x 1))"),
            "Option Explicit\r\n\r\nPublic Sub t()\r\n    x = 1 ' vla:1\r\nEnd Sub\r\n\r\n"
        );
        let f = vba("(function tax ((amount Variant)) (return (* amount 0.08)))");
        assert!(f.contains("Public Function tax(amount As Variant) As Variant\r\n"));
        assert!(f.contains("    tax = (amount * 0.08) ' vla:1\r\n    Exit Function\r\n"));
        let g = vba("(function f ((byval n Long)) Long \"the doc\" (return (+ n 1)))");
        assert!(g.contains("Public Function f(ByVal n As Long) As Long\r\n"));
        assert!(g.contains("    f = (n + 1) ' vla:1\r\n"));
        let p = vba(
            "(sub s ((a Long) (byref b String) (optional c Variant 1) (paramarray d) e) (exit-sub))",
        );
        assert!(p.contains(
            "Public Sub s(a As Long, ByRef b As String, Optional c As Variant = 1, ParamArray d() As Variant, e)\r\n"
        ));
    }

    #[test]
    fn module_level_forms_and_the_topology_guard() {
        let m = vba("(const k 5)\n(dim x Long)\n(private (dim m Long))\n(public (const q \"s\"))\n(type Pt (x Double) y)\n(enum E a (b 2))\n(raw \"' hello\")");
        assert!(m.contains("Const k = 5\r\n\r\nDim x As Long\r\n\r\nPrivate m As Long\r\n\r\nPublic Const q = \"s\"\r\n\r\n"));
        assert!(m.contains("Type Pt\r\n    x As Double\r\n    y As Variant\r\nEnd Type\r\n\r\n"));
        assert!(m.contains("Enum E\r\n    a\r\n    b = 2\r\nEnd Enum\r\n\r\n' hello\r\n"));
        let r = refusal("(sub t () (debug-print 1))\n(dim late Long)");
        assert_eq!(r.id, "vla-module-level-decl-after-proc");
        assert!(r.text.contains("must come before the first"));
        assert_eq!(refusal("5").id, "vla-top-level-not-list");
        assert_eq!(refusal("(nonsense 1)").id, "vla-unknown-top-level-form");
        assert_eq!(refusal("(deflambda f (x) x)").id, "vla-deflambda-runs");
    }

    #[test]
    fn a_formula_write_calls_the_runtime_sink() {
        // SEC.15: exactly (set! (. obj formula) v) is a call to the runtime's
        // VlaSetFormula; a read of the member, and any other member, are
        // written as they were.
        let f = vba("(sub t () (set! (. (range \"b3\") formula) \"=A1\"))");
        assert!(
            f.contains("    Call VlaSetFormula(range(\"b3\"), \"=A1\") ' vla:1\r\n"),
            "{f}"
        );
        assert!(!f.contains("Formula2"), "{f}");
        let g = vba("(sub t () (set! x (. (range \"b3\") formula)))");
        assert!(
            g.contains("    x = range(\"b3\").formula ' vla:1\r\n"),
            "{g}"
        );
        let h = vba("(sub t () (set! (. (range \"b3\") value) \"=A1\"))");
        assert!(
            h.contains("    range(\"b3\").value = \"=A1\" ' vla:1\r\n"),
            "{h}"
        );
    }

    #[test]
    fn the_refusal_names_the_nearest_mapped_line() {
        let r = refusal("(sub t ()\n  (set! x 1)\n  (if))");
        assert!(r.text.ends_with(" (near vla line 3)"), "{}", r.text);
        let r = refusal("(sub t () (at-line 40 (set! x 1) (if)))");
        assert!(
            r.text.ends_with(" (near vla line 1, source line 40)"),
            "{}",
            r.text
        );
    }

    #[test]
    fn tail_calls_rebind_and_jump() {
        let f = vba("(function count-down ((byval n Long) (byval acc Long)) Long (if (= n 0) (then (return acc)) (else (return (count-down (- n 1) (+ acc n))))))");
        assert!(f.contains(
            "    Dim vla_tco_1 As Variant\r\n    Dim vla_tco_2 As Variant\r\nvla_tco:\r\n"
        ));
        assert!(f.contains("        vla_tco_1 = (n - 1) ' vla:1\r\n        vla_tco_2 = (acc + n)\r\n        n = vla_tco_1\r\n        acc = vla_tco_2\r\n        GoTo vla_tco\r\n"));
        let g = vba("(function loop-it ((byval n Long)) Long (if (= n 0) (then (return 0)) (else (return (loop-it (- n 1))))))");
        assert!(g.contains("vla_tco:\r\n"));
        assert!(g.contains("        n = (n - 1) ' vla:1\r\n        GoTo vla_tco\r\n"));
        // A ByRef parameter keeps true recursion: not a byte of TCO.
        let h = vba("(function r ((n Long)) Long (return (r (- n 1))))");
        assert!(!h.contains("vla_tco"));
        assert!(h.contains("    r = r((n - 1)) ' vla:1\r\n"));
    }

    #[test]
    fn make_button_appends_its_helper_once() {
        let t = vba("(sub on:click:go () (debug-print 1))\n(sub main () (make-button \"go\" (range \"A1\")))");
        assert!(t.contains("Call VlaCompiledMakeButton(\"go\", range(\"A1\"), \"on_click_go\")"));
        assert!(t.ends_with(&format!("\r\n{}", make_button_helper_text())));
        assert_eq!(t.matches("Private Sub VlaCompiledMakeButton").count(), 1);
        let u = vba("(sub main () (make-button \"go\" (range \"A1\")))");
        assert!(u.contains(", \"\")"));
        assert_eq!(
            refusal("(sub on:sheet-change () (debug-print 1))").id,
            "vla-interpreter-only-handler"
        );
    }

    #[test]
    fn an_include_line_is_refused_by_name_here() {
        let r = refusal("(include \"lib.vla\")\n(sub t () (debug-print 1))");
        assert_eq!(r.id, "vla-include-cannot-read");
        assert!(r.text.contains("lib.vla"));
        assert_eq!(
            include_line_file("  (include \"a b.vla\")  "),
            Some("a b.vla".to_string())
        );
        assert_eq!(include_line_file("(include \"a\") x"), None);
        assert_eq!(include_line_file("(include-all \"a\")"), None);
        assert_eq!(
            refusal("(sub t () (include \"x\"))").id,
            "vla-include-must-stand-alone"
        );
    }

    #[test]
    fn the_golden_compiles_to_the_golden_from_its_second_line_on() {
        // The treaty's oracle 1b: the stamp line is the writer's, not the
        // program's; what follows it compiled to instructions_golden.vba.
        let program = GOLDEN_VLA.split_once('\n').expect("a stamp line").1;
        let got = compile(program, PRELUDE).unwrap_or_else(|e| panic!("{e}"));
        let want = GOLDEN_VBA;
        let norm = |s: &str| s.replace("\r\n", "\n").trim_end_matches('\n').to_string();
        let (got, want) = (norm(&got), norm(want));
        if got != want {
            let at = got
                .chars()
                .zip(want.chars())
                .position(|(a, b)| a != b)
                .unwrap_or(got.len().min(want.len()));
            let line = want[..at].matches('\n').count() + 1;
            let context = |s: &str| {
                let start = s[..at].rfind('\n').map(|i| i + 1).unwrap_or(0);
                let end = s[at..].find('\n').map(|i| at + i).unwrap_or(s.len());
                s[start..end].to_string()
            };
            panic!(
                "differs at char {at} of {} (line {line}):\n  want: {}\n  got:  {}",
                want.len(),
                context(&want),
                context(&got)
            );
        }
    }
}
