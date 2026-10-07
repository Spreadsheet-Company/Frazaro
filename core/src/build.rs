//! The build (PORT.7): a program's sentences to a workbook's bytes, through
//! the sheet model and the OOXML writer.
//!
//! Slice 7a writes the room's first sheet, `Frazaro`: the program's lines in
//! column B from row 1 exactly as written, blank lines included, and OK in
//! green in column C beside every line that holds text, which is what the
//! add-in's Check marks (`DoCheck`, VLA_IDE.bas: once the program
//! translates, a row whose column B is not blank is marked OK). The widths,
//! the fills, the hidden column A and the gridlines are `BuildWorkspace`'s
//! (VLA_IDE.bas), so the room looks the same through this door.
//!
//! Slice 7b walks the program's VLA, read and macro-expanded exactly as
//! `VlaTranspile` reads it (the prelude first, every `defmacro` defined, every
//! other form expanded to its fixpoint), and writes what a sheet can hold
//! with nothing running: a value or a formula string into a cell, a range,
//! rows of a column or a cell of a named sheet; `Work on sheet` and `Go to
//! sheet`; a `Define`d constant or a variable set to a value that folds; a
//! parameterless step, inlined where it is called; an `If` whose condition
//! folds over literals, so that `set-formula-rows`' guard works. Everything
//! else is refused by name with the sentence quoted (`build-not-representable`),
//! and a sheet named before anything made it is refused too
//! (`build-sheet-unknown`). The scaffold `EnglishToVla` writes around a
//! program (`on-error goto vla-fail`, the `vla-step` counter, the trace
//! guards, `exit-sub`, the fail label) is recognised and stepped over.
//!
//! A formula into more than one cell is written as Excel writes a filled
//! formula, a shared formula, so the host adjusts the references for each
//! cell as `Range.Formula2` would have; a newer function takes its file
//! prefix (`sheet::xlfn`). New ground, said plainly: the reference runs the
//! program in Excel, and Excel does all of this.
//!
//! The build is a pure function of its texts: the same sentences and the
//! same core give the same bytes on every machine, which the build golden
//! (`scripts/build/fixture_golden.xlsx`, the treaty's amendment of
//! 2026-10-03) holds it to.

use std::collections::HashMap;

use crate::expand::Expander;
use crate::form::{Form, List};
use crate::intrinsics::{fold, is_numeric, val};
use crate::messages::{raise, Refusal};
use crate::reader::{count_lf, parse_all, tokenize};
use crate::sha256::sha256_hex_skipping_whitespace;
use crate::sheet::merge::{self, HostInfo, Model};
use crate::sheet::xlfn::{can_return_array, prefix_future_functions};
use crate::sheet::{
    ooxml, parse_a1_range, xml, zip, A1Range, Cell, Column, Content, NumFmt, Rgb, Sheet, Style,
    Styles, Workbook,
};

/// The defined name that carries the build stamp.
pub const STAMP_NAME: &str = "Frazaro.Build";

/// The sheet that holds the program, named as the add-in names its
/// workspace (`IDE_SHEET`).
pub const FRAZARO_SHEET: &str = "Frazaro";
/// The sheet a program's results go to unless it says `Work on sheet`
/// (`OUT_SHEET`).
pub const OUTPUT_SHEET: &str = "Output";
/// `BuildWorkspace`'s whisper of the logo's lavender on the sentence column.
pub const LAVENDER: Rgb = Rgb(247, 244, 252);
/// The result column's gray, under the marks.
pub const RESULT_GRAY: Rgb = Rgb(242, 242, 242);
/// `MarkOK`'s green.
pub const OK_GREEN: Rgb = Rgb(221, 235, 221);
/// `BuildWorkspace`'s column widths, in characters.
pub const SENTENCE_COLUMN_WIDTH: f64 = 72.0;
pub const RESULT_COLUMN_WIDTH: f64 = 60.0;
/// The most cells one sentence may fill: one whole column's worth.
pub const MAX_CELLS_PER_WRITE: u64 = 1_048_576;
/// How deep steps may call steps before the build calls it a loop.
const MAX_INLINE_DEPTH: u32 = 32;

/// The `Frazaro` sheet of a program whose text translated: each line of the
/// text in column B of its own row, and OK in column C beside a line that
/// is not blank. Blank is `Len(Trim$(...)) = 0`, and VBA's `Trim$` trims
/// spaces only, so a line of tabs is marked and a line of spaces is not.
pub fn frazaro_sheet(program_text: &str, styles: &mut Styles) -> Sheet {
    let sentence = styles.id(Style {
        fill: Some(LAVENDER),
        num_fmt: NumFmt::Text,
        wrap: false,
    });
    let result = styles.id(Style {
        fill: Some(RESULT_GRAY),
        num_fmt: NumFmt::General,
        wrap: true,
    });
    let ok = styles.id(Style {
        fill: Some(OK_GREEN),
        num_fmt: NumFmt::General,
        wrap: true,
    });
    let mut sheet = Sheet::new(FRAZARO_SHEET);
    sheet.gridlines = false;
    sheet.active_cell = (1, 2);
    sheet.columns = vec![
        Column {
            index: 1,
            width: None,
            hidden: true,
            style: None,
        },
        Column {
            index: 2,
            width: Some(SENTENCE_COLUMN_WIDTH),
            hidden: false,
            style: Some(sentence),
        },
        Column {
            index: 3,
            width: Some(RESULT_COLUMN_WIDTH),
            hidden: false,
            style: Some(result),
        },
    ];
    for (i, line) in program_text.lines().enumerate() {
        let row = i as u32 + 1;
        if line.is_empty() {
            continue;
        }
        sheet.set(
            row,
            2,
            Cell {
                content: Content::Text(line.to_string()),
                style: sentence,
            },
        );
        if !line.trim_matches(' ').is_empty() {
            sheet.set(
                row,
                3,
                Cell {
                    content: Content::Text("OK".to_string()),
                    style: ok,
                },
            );
        }
    }
    sheet
}

/// A value the build can put in a cell, or fold a condition from.
#[derive(Clone, Debug, PartialEq)]
enum Value {
    Num(f64),
    Text(String),
    Bool(bool),
}

/// VBA's `CStr` of a value, as `&` spells its operands: an integral number
/// without a decimal point, anything else as Rust spells it (close enough
/// for the cell addresses a program builds).
fn cstr(v: &Value) -> String {
    match v {
        Value::Num(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            }
        }
        Value::Text(t) => t.clone(),
        Value::Bool(b) => (if *b { "True" } else { "False" }).to_string(),
    }
}

/// Why Excel would refuse a sheet name, or `None` for a name it takes: the
/// rules `VlaCheckSheetName` (VLA_Runtime.bas) holds a program to at run
/// time, with the runtime's own words; here the build's.
fn sheet_name_problem(name: &str) -> Option<String> {
    if name.is_empty() {
        return Some("it is empty".to_string());
    }
    let n = name.chars().count();
    if n > 31 {
        return Some(format!("it is {n} characters long"));
    }
    // The first offending character in the runtime's own order, as its
    // message names it.
    if let Some(c) = ":\\/?*[]".chars().find(|c| name.contains(*c)) {
        return Some(format!("it contains {c}"));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Some("it begins or ends with an apostrophe".to_string());
    }
    if fold(name) == "history" {
        return Some("History is a name Excel keeps for itself".to_string());
    }
    None
}

/// What the build keeps of a program's procedure: how many parameters it
/// takes, and its body.
#[derive(Clone)]
struct Step {
    params: usize,
    body: Vec<Form>,
}

/// Whether a statement let the walk go on, or ended the procedure.
enum Flow {
    Next,
    Exit,
}

/// Where `(range ...)` writes: a sheet of this build's, or one of the host
/// workbook's, which a program may go to but never write into.
#[derive(Clone, Debug, PartialEq, Eq)]
enum Target {
    Ours(usize),
    Host(String),
}

struct Walker<'a> {
    lines: Vec<&'a str>,
    wb: Workbook,
    /// The host workbook's own sheets, when the build adds to one.
    host_sheets: Vec<String>,
    /// The sheet a `(range ...)` means: `None` until something writes, then
    /// `Output` unless `Work on sheet` chose another.
    current: Option<Target>,
    /// Constants and variables whose values fold, by folded name.
    bindings: HashMap<String, Value>,
    /// The program's procedures, by folded name.
    steps: HashMap<String, Step>,
    /// The program line the walk is on, from the nearest `(at-line N ...)`.
    line: u32,
    depth: u32,
}

impl<'a> Walker<'a> {
    fn new(program_text: &'a str, host_sheets: Vec<String>) -> Walker<'a> {
        Walker {
            lines: program_text.lines().collect(),
            wb: Workbook::new(),
            host_sheets,
            current: None,
            bindings: HashMap::new(),
            steps: HashMap::new(),
            line: 0,
            depth: 0,
        }
    }

    /// The sentence the walk is on, quoted in a refusal.
    fn sentence(&self) -> String {
        if self.line == 0 {
            return "the program's own setup".to_string();
        }
        match self.lines.get(self.line as usize - 1).map(|s| s.trim()) {
            Some(s) if !s.is_empty() => s.to_string(),
            _ => format!("the sentence on line {}", self.line),
        }
    }

    fn refuse_sentence(&self) -> Refusal {
        raise(
            "build-not-representable",
            &[
                ("line", &self.line.to_string()),
                ("sentence", &self.sentence()),
            ],
        )
    }

    fn refuse_sheet(&self, name: &str) -> Refusal {
        raise(
            "build-sheet-unknown",
            &[("line", &self.line.to_string()), ("name", name)],
        )
    }

    /// A top-level form: a constant is bound, a procedure kept, an include
    /// refused as the compiler refuses it (no file system here either), and
    /// the declarations the compiled module needs are nothing to a sheet.
    fn collect_top(&mut self, f: &Form) -> Result<(), Refusal> {
        let Form::List(l) = f else { return Ok(()) };
        let Ok(head) = l.head_sym() else {
            return Ok(());
        };
        match fold(&head).as_str() {
            "const" => {
                if l.count() >= 3 {
                    if let Some(name) = l.items[1].atom_text() {
                        if let Some(v) = self.fold_expr(&l.items[2]) {
                            self.bindings.insert(fold(&name), v);
                        }
                    }
                }
            }
            "sub" => {
                if l.count() >= 3 {
                    if let (Some(name), Form::List(params)) = (l.items[1].atom_text(), &l.items[2])
                    {
                        let mut body: &[Form] = &l.items[3..];
                        if let Some(Form::Str(_)) = body.first() {
                            body = &body[1..]; // a docstring
                        }
                        self.steps.insert(
                            fold(&name),
                            Step {
                                params: params.count(),
                                body: body.to_vec(),
                            },
                        );
                    }
                }
            }
            "include" => {
                let name = l
                    .items
                    .get(1)
                    .and_then(|f| f.atom_text())
                    .unwrap_or_default();
                return Err(raise(
                    "vla-include-cannot-read",
                    &[("name", &name), ("path", &name)],
                ));
            }
            _ => {}
        }
        Ok(())
    }

    fn walk_main(&mut self) -> Result<(), Refusal> {
        let Some(main) = self.steps.get("main").cloned() else {
            return Ok(()); // a program with no statements builds its sheet alone
        };
        self.walk_body(&main.body)?;
        Ok(())
    }

    fn walk_body(&mut self, body: &[Form]) -> Result<Flow, Refusal> {
        for f in body {
            if let Flow::Exit = self.walk_stmt(f)? {
                return Ok(Flow::Exit);
            }
        }
        Ok(Flow::Next)
    }

    fn walk_stmt(&mut self, f: &Form) -> Result<Flow, Refusal> {
        let Form::List(l) = f else {
            return Err(self.refuse_sentence());
        };
        let head = match l.head_sym() {
            Ok(h) => fold(&h),
            Err(_) => return Err(self.refuse_sentence()),
        };
        match head.as_str() {
            "at-line" => {
                if let Some(n) = l.items.get(1).and_then(|f| f.atom_text()) {
                    self.line = val(&n) as u32;
                }
                self.walk_body(&l.items[2.min(l.items.len())..])
            }
            "begin" => self.walk_body(&l.items[1..]),
            // Locals, and the scaffold's own report call.
            "dim" | "vla-report-error" => Ok(Flow::Next),
            // The scaffold's error handler; a `Try:` block's handlers are not it.
            "on-error" => {
                let target = l
                    .items
                    .last()
                    .and_then(|f| f.atom_text())
                    .unwrap_or_default();
                if fold(&target) == "vla-fail" || target == "0" {
                    Ok(Flow::Next)
                } else {
                    Err(self.refuse_sentence())
                }
            }
            "label" => {
                let name = l
                    .items
                    .get(1)
                    .and_then(|f| f.atom_text())
                    .unwrap_or_default();
                if fold(&name) == "vla-fail" {
                    Ok(Flow::Exit)
                } else {
                    Err(self.refuse_sentence())
                }
            }
            "exit-sub" | "exit-function" => Ok(Flow::Exit),
            "set!" => self.walk_set(l),
            "if" => self.walk_if(l),
            "vlaensuresheet" => {
                let name = self.fold_text(l.items.get(1))?;
                if self.host_sheet(&name).is_some() {
                    return Ok(Flow::Next); // the host has it; nothing to make
                }
                if let Some(reason) = sheet_name_problem(&name) {
                    return Err(raise(
                        "build-sheet-name-invalid",
                        &[
                            ("line", &self.line.to_string()),
                            ("name", &name),
                            ("reason", &reason),
                        ],
                    ));
                }
                self.wb.ensure_sheet(&name);
                Ok(Flow::Next)
            }
            "." => self.walk_dot(l),
            _ => self.walk_call(l, &head),
        }
    }

    /// `(set! place value)`: a variable bound, a cell or range written, a
    /// formula written, a cell of a named sheet written.
    fn walk_set(&mut self, l: &List) -> Result<Flow, Refusal> {
        if l.count() != 3 {
            return Err(self.refuse_sentence());
        }
        let place = &l.items[1];
        let value = &l.items[2];
        match place {
            Form::Sym(name) => {
                let v = self
                    .fold_expr(value)
                    .ok_or_else(|| self.refuse_sentence())?;
                self.bindings.insert(fold(name), v);
                Ok(Flow::Next)
            }
            Form::Str(_) => Err(self.refuse_sentence()),
            Form::List(pl) => {
                let phead = pl.head_sym().map(|h| fold(&h)).unwrap_or_default();
                if phead == "range" && pl.count() == 2 {
                    let range = self.range_of(&pl.items[1])?;
                    let v = self
                        .fold_expr(value)
                        .ok_or_else(|| self.refuse_sentence())?;
                    let sheet = self.writable_current()?;
                    self.write_value(sheet, range, &v)?;
                    return Ok(Flow::Next);
                }
                if phead != "." || pl.count() < 3 {
                    return Err(self.refuse_sentence());
                }
                let obj = &pl.items[1];
                let member = fold(&pl.items[2].atom_text().unwrap_or_default());
                // (. (range R) formula) and (. (range R) value)
                if let (Some(rl), true) = (obj.as_list(), obj.head_is("range")) {
                    if pl.count() == 3
                        && rl.count() == 2
                        && matches!(member.as_str(), "formula" | "formula2" | "value")
                    {
                        let range = self.range_of(&rl.items[1])?;
                        let v = self
                            .fold_expr(value)
                            .ok_or_else(|| self.refuse_sentence())?;
                        let sheet = self.writable_current()?;
                        if member == "value" {
                            self.write_value(sheet, range, &v)?;
                        } else {
                            self.write_formula(sheet, range, &v)?;
                        }
                        return Ok(Flow::Next);
                    }
                }
                // (. (worksheets S) range R): put-into-cell-of-sheet
                if let (Some(wl), true) = (obj.as_list(), obj.head_is("worksheets")) {
                    if pl.count() == 4 && wl.count() == 2 && member == "range" {
                        let name = self.fold_text(wl.items.get(1))?;
                        let sheet = self.writable_named(&name)?;
                        let range = self.range_of(&pl.items[3])?;
                        let v = self
                            .fold_expr(value)
                            .ok_or_else(|| self.refuse_sentence())?;
                        self.write_value(sheet, range, &v)?;
                        return Ok(Flow::Next);
                    }
                }
                Err(self.refuse_sentence())
            }
        }
    }

    /// `(if cond (then ...) [(else ...)])`: the trace guard is skipped (no
    /// trace runs in a build); another condition must fold.
    fn walk_if(&mut self, l: &List) -> Result<Flow, Refusal> {
        let Some(cond) = l.items.get(1) else {
            return Err(self.refuse_sentence());
        };
        if cond.head_is("vlatraceon") {
            return Ok(Flow::Next);
        }
        let truth = match self.fold_expr(cond) {
            Some(Value::Bool(b)) => b,
            Some(Value::Num(n)) => n != 0.0,
            _ => return Err(self.refuse_sentence()),
        };
        for branch in &l.items[2..] {
            if let Form::List(bl) = branch {
                let taken = (truth && bl.head_is("then")) || (!truth && bl.head_is("else"));
                if taken {
                    return self.walk_body(&bl.items[1..]);
                }
            }
        }
        Ok(Flow::Next)
    }

    /// `(. (worksheets S) activate)`: `Go to sheet`, and the second half of
    /// `Work on sheet`.
    fn walk_dot(&mut self, l: &List) -> Result<Flow, Refusal> {
        let obj = &l.items[1];
        let member = fold(
            &l.items
                .get(2)
                .and_then(|f| f.atom_text())
                .unwrap_or_default(),
        );
        if let (Some(wl), true) = (obj.as_list(), obj.head_is("worksheets")) {
            if l.count() == 3 && wl.count() == 2 && member == "activate" {
                let name = self.fold_text(wl.items.get(1))?;
                let target = self.target_named(&name)?;
                self.current = Some(target);
                return Ok(Flow::Next);
            }
        }
        Err(self.refuse_sentence())
    }

    /// A call of one of the program's own steps, with no values passed:
    /// its body inlined here. Anything else a sentence calls needs a host.
    fn walk_call(&mut self, l: &List, head: &str) -> Result<Flow, Refusal> {
        let Some(step) = self.steps.get(head).cloned() else {
            return Err(self.refuse_sentence());
        };
        if step.params != 0 || l.count() != 1 || self.depth >= MAX_INLINE_DEPTH {
            return Err(self.refuse_sentence());
        }
        self.depth += 1;
        let saved = self.line;
        let flow = self.walk_body(&step.body);
        self.depth -= 1;
        self.line = saved;
        flow?; // a step's own Exit Sub ends the step, not its caller
        Ok(Flow::Next)
    }

    fn fold_text(&self, f: Option<&Form>) -> Result<String, Refusal> {
        match f.and_then(|f| self.fold_expr(f)) {
            Some(Value::Text(t)) => Ok(t),
            _ => Err(self.refuse_sentence()),
        }
    }

    fn range_of(&self, f: &Form) -> Result<A1Range, Refusal> {
        let text = self.fold_text(Some(f))?;
        parse_a1_range(&text).ok_or_else(|| self.refuse_sentence())
    }

    /// The host workbook's sheet of that name, as the host spells it.
    fn host_sheet(&self, name: &str) -> Option<String> {
        let want = fold(name);
        self.host_sheets.iter().find(|s| fold(s) == want).cloned()
    }

    /// Where `(range ...)` writes: `Output` made on first use, unless the
    /// host workbook has an `Output` of its own.
    fn current_target(&mut self) -> Target {
        if let Some(t) = &self.current {
            return t.clone();
        }
        let t = match self.host_sheet(OUTPUT_SHEET) {
            Some(h) => Target::Host(h),
            None => Target::Ours(self.wb.ensure_sheet(OUTPUT_SHEET)),
        };
        self.current = Some(t.clone());
        t
    }

    /// A sheet named by the program: one it made, one of the host's, or
    /// `Output`, which the add-in has at the start of every run.
    fn target_named(&mut self, name: &str) -> Result<Target, Refusal> {
        if let Some(i) = self.wb.find_sheet(name) {
            return Ok(Target::Ours(i));
        }
        if let Some(h) = self.host_sheet(name) {
            return Ok(Target::Host(h));
        }
        if fold(name) == fold(OUTPUT_SHEET) {
            return Ok(Target::Ours(self.wb.ensure_sheet(OUTPUT_SHEET)));
        }
        Err(self.refuse_sheet(name))
    }

    /// A target a sentence may write to: one of this build's sheets. One
    /// of the host's is refused by name, since the build leaves the host's
    /// own sheets as they are.
    fn writable(&self, t: Target) -> Result<usize, Refusal> {
        match t {
            Target::Ours(i) => Ok(i),
            Target::Host(name) => Err(raise(
                "build-into-model-sheet",
                &[("line", &self.line.to_string()), ("name", &name)],
            )),
        }
    }

    fn writable_current(&mut self) -> Result<usize, Refusal> {
        let t = self.current_target();
        self.writable(t)
    }

    fn writable_named(&mut self, name: &str) -> Result<usize, Refusal> {
        let t = self.target_named(name)?;
        self.writable(t)
    }

    fn write_value(&mut self, sheet: usize, range: A1Range, v: &Value) -> Result<(), Refusal> {
        if range.cells() > MAX_CELLS_PER_WRITE {
            return Err(self.refuse_sentence());
        }
        let content = match v {
            Value::Num(n) => Content::Number(*n),
            Value::Text(t) => Content::Text(t.clone()),
            Value::Bool(b) => Content::Bool(*b),
        };
        let sh = &mut self.wb.sheets[sheet];
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                sh.set(
                    row,
                    col,
                    Cell {
                        content: content.clone(),
                        style: 0,
                    },
                );
            }
        }
        Ok(())
    }

    /// `Range.Formula2 = text`: a text beginning with `=` is a formula, with
    /// newer functions under their file prefix; anything else is the value
    /// itself, as Excel stores it. SEC.15: the text is scanned first, as the
    /// add-in's `VlaSetFormula` scans it, and a formula that reaches outside
    /// the workbook on its own is refused by name with the sentence quoted
    /// (`build-formula-egress`), nothing written.
    fn write_formula(&mut self, sheet: usize, range: A1Range, v: &Value) -> Result<(), Refusal> {
        if range.cells() > MAX_CELLS_PER_WRITE {
            return Err(self.refuse_sentence());
        }
        if let Value::Text(t) = v {
            if let Some(name) = crate::egress::egress_call(t) {
                return Err(raise(
                    "build-formula-egress",
                    &[
                        ("line", &self.line.to_string()),
                        ("name", &name),
                        ("sentence", &self.sentence()),
                    ],
                ));
            }
            if let Some(body) = t.strip_prefix('=') {
                let text = prefix_future_functions(body);
                if can_return_array(&text) {
                    // What Excel stores for a Formula2 entry of such a formula,
                    // and reads back without its @.
                    self.wb.sheets[sheet].set_formula_dynamic(range, &text, 0);
                } else {
                    self.wb.sheets[sheet].set_formula(range, &text, 0);
                }
                return Ok(());
            }
        }
        self.write_value(sheet, range, v)
    }

    /// A value from literals, constants, variables set to literals, and the
    /// operators over them; `None` for anything that needs a host.
    fn fold_expr(&self, f: &Form) -> Option<Value> {
        match f {
            Form::Str(s) => Some(Value::Text(s.clone())),
            Form::Sym(s) => {
                if is_numeric(s) {
                    return Some(Value::Num(val(s)));
                }
                let key = fold(s);
                match key.as_str() {
                    "true" => Some(Value::Bool(true)),
                    "false" => Some(Value::Bool(false)),
                    _ => self.bindings.get(&key).cloned(),
                }
            }
            Form::List(l) => {
                let head = fold(&l.head_sym().ok()?);
                let args: Vec<Value> = l.items[1..]
                    .iter()
                    .map(|a| self.fold_expr(a))
                    .collect::<Option<Vec<Value>>>()?;
                fold_op(&head, &args)
            }
        }
    }
}

fn nums(args: &[Value]) -> Option<Vec<f64>> {
    args.iter()
        .map(|v| match v {
            Value::Num(n) => Some(*n),
            _ => None,
        })
        .collect()
}

fn bools(args: &[Value]) -> Option<Vec<bool>> {
    args.iter()
        .map(|v| match v {
            Value::Bool(b) => Some(*b),
            _ => None,
        })
        .collect()
}

/// The operators the build folds: arithmetic and comparison over numbers,
/// `&` over anything, comparison over texts, `not`/`and`/`or` over truths.
fn fold_op(head: &str, args: &[Value]) -> Option<Value> {
    match head {
        "&" => Some(Value::Text(
            args.iter().map(cstr).collect::<Vec<_>>().join(""),
        )),
        "+" | "*" | "/" | "-" => {
            let n = nums(args)?;
            if n.is_empty() {
                return None;
            }
            if head == "-" && n.len() == 1 {
                return Some(Value::Num(-n[0]));
            }
            let mut acc = n[0];
            for x in &n[1..] {
                acc = match head {
                    "+" => acc + x,
                    "*" => acc * x,
                    "-" => acc - x,
                    _ => {
                        if *x == 0.0 {
                            return None;
                        }
                        acc / x
                    }
                };
            }
            Some(Value::Num(acc))
        }
        "=" | "<>" | "<" | ">" | "<=" | ">=" => {
            if args.len() != 2 {
                return None;
            }
            let ord = match (&args[0], &args[1]) {
                (Value::Num(a), Value::Num(b)) => a.partial_cmp(b)?,
                (Value::Text(a), Value::Text(b)) => a.cmp(b),
                (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
                _ => return None,
            };
            Some(Value::Bool(match head {
                "=" => ord.is_eq(),
                "<>" => ord.is_ne(),
                "<" => ord.is_lt(),
                ">" => ord.is_gt(),
                "<=" => ord.is_le(),
                _ => ord.is_ge(),
            }))
        }
        "not" => {
            let b = bools(args)?;
            (b.len() == 1).then(|| Value::Bool(!b[0]))
        }
        "and" => Some(Value::Bool(bools(args)?.iter().all(|b| *b))),
        "or" => Some(Value::Bool(bools(args)?.iter().any(|b| *b))),
        _ => None,
    }
}

/// One source's identity in the stamp: `EnglishSourceHash`'s digest of its
/// non-whitespace bytes, and how many there were.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SourceHash {
    pub hex: String,
    pub count: usize,
}

impl SourceHash {
    pub fn of(text: &str) -> SourceHash {
        let (hex, count) = sha256_hex_skipping_whitespace(text.as_bytes());
        SourceHash { hex, count }
    }

    fn text(&self) -> String {
        format!(
            "sha256:{} over {} non-whitespace bytes",
            self.hex, self.count
        )
    }

    fn parse(s: &str) -> Option<SourceHash> {
        let rest = s.strip_prefix("sha256:")?;
        let (hex, tail) = rest.split_once(" over ")?;
        let count = tail.strip_suffix(" non-whitespace bytes")?.parse().ok()?;
        if hex.len() != 64 || !hex.chars().all(|c| c.is_ascii_hexdigit()) {
            return None;
        }
        Some(SourceHash {
            hex: hex.to_string(),
            count,
        })
    }
}

/// What the stamp records of the workbook a build was added to
/// (`--into`): enough for `rebuild` to render the build's sheets again
/// exactly, without the model at hand.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IntoStamp {
    /// The SHA-256 of the model's bytes.
    pub model_hex: String,
    /// How many cell formats the model had: where this build's begin.
    pub style_base: u32,
    /// The cell-metadata index the build's dynamic-array formulas point
    /// at; 0 when none was needed.
    pub cm: u32,
}

/// The build stamp: the core's version, the workbook added to when there
/// was one, and the identity of the sentences, the prelude and each
/// phrasebook in order, as the defined name `Frazaro.Build` carries it in
/// one string.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Stamp {
    pub version: String,
    pub into: Option<IntoStamp>,
    pub sentences: SourceHash,
    pub prelude: SourceHash,
    pub phrasebooks: Vec<SourceHash>,
}

impl Stamp {
    pub fn of(program_text: &str, prelude: &str, books: &[&str], into: Option<IntoStamp>) -> Stamp {
        Stamp {
            version: crate::VERSION.to_string(),
            into,
            sentences: SourceHash::of(program_text),
            prelude: SourceHash::of(prelude),
            phrasebooks: books.iter().map(|b| SourceHash::of(b)).collect(),
        }
    }

    /// `Frazaro 0.7.1; [into sha256:... with N cell formats and metadata
    /// M; ]sentences sha256:... over N non-whitespace bytes; prelude ...;
    /// phrasebook 1 ...`.
    pub fn text(&self) -> String {
        let mut s = format!("Frazaro {}", self.version);
        if let Some(into) = &self.into {
            s.push_str(&format!(
                "; into sha256:{} with {} cell formats and metadata {}",
                into.model_hex, into.style_base, into.cm
            ));
        }
        s.push_str(&format!(
            "; sentences {}; prelude {}",
            self.sentences.text(),
            self.prelude.text()
        ));
        for (i, b) in self.phrasebooks.iter().enumerate() {
            s.push_str(&format!("; phrasebook {} {}", i + 1, b.text()));
        }
        s
    }

    pub fn parse(text: &str) -> Option<Stamp> {
        let mut fields = text.split("; ").peekable();
        let version = fields.next()?.strip_prefix("Frazaro ")?.to_string();
        let into = if fields.peek().is_some_and(|f| f.starts_with("into ")) {
            let f = fields.next()?.strip_prefix("into sha256:")?;
            let (hex, rest) = f.split_once(" with ")?;
            let (base, cm) = rest.split_once(" cell formats and metadata ")?;
            if hex.len() != 64 || !hex.chars().all(|c| c.is_ascii_hexdigit()) {
                return None;
            }
            Some(IntoStamp {
                model_hex: hex.to_string(),
                style_base: base.parse().ok()?,
                cm: cm.parse().ok()?,
            })
        } else {
            None
        };
        let sentences = SourceHash::parse(fields.next()?.strip_prefix("sentences ")?)?;
        let prelude = SourceHash::parse(fields.next()?.strip_prefix("prelude ")?)?;
        let mut phrasebooks = Vec::new();
        for (i, f) in fields.enumerate() {
            let rest = f.strip_prefix(&format!("phrasebook {} ", i + 1))?;
            phrasebooks.push(SourceHash::parse(rest)?);
        }
        Some(Stamp {
            version,
            into,
            sentences,
            prelude,
            phrasebooks,
        })
    }

    /// Why the stamp does not name these sources, or `None` when it does.
    pub fn disagreement(
        &self,
        program_text: &str,
        prelude: &str,
        books: &[&str],
    ) -> Option<String> {
        if self.sentences != SourceHash::of(program_text) {
            return Some(
                "the sentences in its Frazaro sheet are not the ones the stamp names".to_string(),
            );
        }
        if self.prelude != SourceHash::of(prelude) {
            return Some("the prelude given is not the one the stamp names".to_string());
        }
        if self.phrasebooks.len() != books.len() {
            return Some(format!(
                "the stamp names {} phrasebook(s) and {} were given",
                self.phrasebooks.len(),
                books.len()
            ));
        }
        for (i, (want, text)) in self.phrasebooks.iter().zip(books.iter()).enumerate() {
            if *want != SourceHash::of(text) {
                return Some(format!(
                    "phrasebook {} given is not the one the stamp names",
                    i + 1
                ));
            }
        }
        None
    }
}

/// The stamp as a defined name's formula: a string constant.
fn stamp_formula(stamp: &Stamp) -> String {
    format!("\"{}\"", stamp.text().replace('"', "\"\""))
}

/// What `frazaro rebuild` reads out of a built workbook before building it
/// again.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BuildFile {
    pub stamp: Stamp,
    /// The program, line by line, as the `Frazaro` sheet holds it.
    pub program_text: String,
    /// The lines with text in them, which Check marks and the stamp counts.
    pub sentences: usize,
    /// The sheets that are not the build's: the model's, when the build
    /// was added to one.
    pub host_sheets: Vec<String>,
}

/// What `frazaro rebuild` answers.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Rebuilt {
    pub sentences: usize,
    pub built_by: String,
    pub matches: bool,
    /// Empty when it matches.
    pub why: String,
    /// Whether only the build's own sheets were checked, as for a workbook
    /// the build was added to, whose own parts the rebuild cannot remake.
    pub partial: bool,
}

impl Rebuilt {
    /// The one line `frazaro rebuild` prints.
    pub fn line(&self) -> String {
        let head = if self.partial {
            format!(
                "This workbook's Frazaro sheets were built from these {} sentences by Frazaro {}, into a workbook whose own sheets are not checked:",
                self.sentences, self.built_by
            )
        } else {
            format!(
                "This workbook was built from these {} sentences by Frazaro {}:",
                self.sentences, self.built_by
            )
        };
        if self.matches {
            format!("{head} yes.")
        } else {
            format!("{head} no. Why: {}.", self.why)
        }
    }
}

/// A stored part of a build, as text, for `rebuild` to compare with what
/// it renders again; `None` when the part is missing or not stored.
pub fn own_part_text(file: &[u8], name: &str) -> Option<String> {
    let entries = zip::entries(file)?;
    let e = entries.iter().find(|e| e.name == name)?;
    String::from_utf8(zip::stored_data(file, e)?.to_vec()).ok()
}

fn refuse_rebuild(why: &str) -> Refusal {
    raise("rebuild-not-a-build", &[("why", why)])
}

/// The text between `<open ...>` and `</close>` after `from`, with the
/// element's start tag given up to its name.
fn element_text<'a>(xml_text: &'a str, start_tag: &str, end_tag: &str) -> Option<&'a str> {
    let open = xml_text.find(start_tag)?;
    let after_open = open + xml_text[open..].find('>')? + 1;
    let close = after_open + xml_text[after_open..].find(end_tag)?;
    Some(&xml_text[after_open..close])
}

/// A stored part of our own archive, as text.
fn stored_part(file: &[u8], entries: &[zip::Entry], name: &str) -> Result<String, Refusal> {
    let e = entries
        .iter()
        .find(|e| e.name == name)
        .ok_or_else(|| refuse_rebuild(&format!("it has no part {name}")))?;
    let data = zip::stored_data(file, e).ok_or_else(|| {
        refuse_rebuild(
            "its parts are not stored as frazaro build stores them, so a host has saved it since it was built, if it was built at all",
        )
    })?;
    String::from_utf8(data.to_vec())
        .map_err(|_| refuse_rebuild(&format!("its part {name} is not UTF-8")))
}

/// A built workbook read back: its stamp and its sentences.
pub fn read_build(file: &[u8]) -> Result<BuildFile, Refusal> {
    let entries = zip::entries(file)
        .ok_or_else(|| refuse_rebuild("it is not a zip archive this reader knows"))?;
    let workbook = stored_part(file, &entries, "xl/workbook.xml")?;
    let raw = element_text(
        &workbook,
        &format!("<definedName name=\"{STAMP_NAME}\""),
        "</definedName>",
    )
    .ok_or_else(|| refuse_rebuild("it carries no Frazaro.Build stamp"))?;
    let formula = xml::unescape(raw);
    let quoted = formula
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .ok_or_else(|| refuse_rebuild("its Frazaro.Build stamp is not a text"))?;
    let stamp = Stamp::parse(&quoted.replace("\"\"", "\""))
        .ok_or_else(|| refuse_rebuild("its Frazaro.Build stamp does not read"))?;
    // The Frazaro sheet is the first sheet part; its column B is the program.
    let sheet = stored_part(file, &entries, &ooxml::sheet_part(1))?;
    let mut lines: Vec<String> = Vec::new();
    let mut rest = sheet.as_str();
    while let Some(i) = rest.find("<c r=\"B") {
        let cell = &rest[i + 7..];
        let Some(q) = cell.find('"') else { break };
        let row: usize = match cell[..q].parse() {
            Ok(n) => n,
            Err(_) => {
                rest = cell;
                continue;
            }
        };
        let Some(end) = cell.find("</c>") else { break };
        let body = &cell[..end];
        let text = element_text(body, "<t", "</t>")
            .map(xml::unescape)
            .unwrap_or_default();
        if row >= 1 {
            if lines.len() < row {
                lines.resize(row, String::new());
            }
            lines[row - 1] = text;
        }
        rest = &cell[end..];
    }
    let sentences = lines
        .iter()
        .filter(|l| !l.trim_matches(' ').is_empty())
        .count();
    let rels = stored_part(file, &entries, "xl/_rels/workbook.xml.rels")?;
    let host_sheets = merge::host_sheets_of(&workbook, &rels);
    Ok(BuildFile {
        stamp,
        program_text: lines.join("\n"),
        sentences,
        host_sheets,
    })
}

/// The workbook a translated program builds: the `Frazaro` sheet, then
/// the sheets its static subset writes, and the stamp. `vla` is the
/// program's VLA as `EnglishToVla` wrote it, read with the prelude and
/// expanded as the compiler reads it; `books` are the phrasebook texts the
/// translation used, in order, for the stamp alone. With a `host`, the
/// build is being added to that workbook: the host's sheets may be named
/// but not written to, and the stamp records the host.
pub fn build_workbook_with(
    program_text: &str,
    vla: &str,
    prelude: &str,
    books: &[&str],
    host: Option<&HostInfo>,
) -> Result<Workbook, Refusal> {
    let line_offset = count_lf(prelude) as u32 + 1;
    let text = format!("{prelude}\r\n{vla}");
    let forms = parse_all(&tokenize(&text), line_offset)?;
    let mut expander = Expander::new();
    let body = expander.collect(forms)?;
    let mut expanded = Vec::with_capacity(body.len());
    for f in &body {
        expanded.push(expander.expand(f, 0)?);
    }
    let mut w = Walker::new(
        program_text,
        host.map(|h| h.sheet_names.clone()).unwrap_or_default(),
    );
    if let Some(h) = host {
        if let Some(taken) = w.host_sheet(FRAZARO_SHEET) {
            return Err(raise(
                "build-into-sheet-name-taken",
                &[("path", &h.label), ("name", &taken)],
            ));
        }
    }
    let sheet = frazaro_sheet(program_text, &mut w.wb.styles);
    w.wb.sheets.push(sheet);
    w.wb.active_sheet = 0;
    for f in &expanded {
        w.collect_top(f)?;
    }
    w.walk_main()?;
    let into = match host {
        None => None,
        Some(h) => {
            if h.cm.is_none() && w.wb.has_dynamic_formulas() {
                return Err(raise(
                    "build-into-unsupported",
                    &[
                        ("path", &h.label),
                        ("part", "xl/metadata.xml"),
                        ("why", "its cell metadata has no dynamic-array block for this build's formulas to point at"),
                    ],
                ));
            }
            Some(IntoStamp {
                model_hex: h.model_hex.clone(),
                style_base: h.style_base,
                cm: h.cm.unwrap_or(0),
            })
        }
    };
    let stamp = Stamp::of(program_text, prelude, books, into);
    w.wb.defined_names
        .push((STAMP_NAME.to_string(), stamp_formula(&stamp)));
    Ok(w.wb)
}

/// A fresh workbook from a translated program.
pub fn build_workbook(
    program_text: &str,
    vla: &str,
    prelude: &str,
    books: &[&str],
) -> Result<Workbook, Refusal> {
    build_workbook_with(program_text, vla, prelude, books, None)
}

/// The workbook's bytes, or the catalogue's refusal when they would not fit
/// the file format.
pub fn build_xlsx(
    program_text: &str,
    vla: &str,
    prelude: &str,
    books: &[&str],
) -> Result<Vec<u8>, Refusal> {
    let wb = build_workbook(program_text, vla, prelude, books)?;
    ooxml::workbook_bytes(&wb).ok_or_else(|| raise("build-workbook-too-large", &[]))
}

/// The model with the program's sheets added, as bytes (`--into`): the
/// model's parts as they were, this build's appended (`sheet::merge`).
/// `label` is what a refusal calls the model.
pub fn build_xlsx_into(
    program_text: &str,
    vla: &str,
    prelude: &str,
    books: &[&str],
    model_bytes: Vec<u8>,
    label: &str,
) -> Result<Vec<u8>, Refusal> {
    let model = Model::read(model_bytes, label)?;
    let host = model.host_info();
    let wb = build_workbook_with(program_text, vla, prelude, books, Some(&host))?;
    merge::write(&model, &wb, host.cm.unwrap_or(0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::zip;

    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../scripts/polyglotta/english.vla");
    const FIXTURE: &str = include_str!("../../scripts/build/fixture.txt");
    const GOLDEN: &[u8] = include_bytes!("../../scripts/build/fixture_golden.xlsx");

    fn text_of(cell: &Cell) -> &str {
        match &cell.content {
            Content::Text(t) => t,
            other => panic!("not a text: {other:?}"),
        }
    }

    #[test]
    fn the_frazaro_sheet_is_the_room() {
        let mut styles = Styles::new();
        let sh = frazaro_sheet(
            "Put 5 into cell B2.\n\n# a comment\n   \n\tindented\nlast",
            &mut styles,
        );
        assert_eq!(sh.name, "Frazaro");
        assert!(!sh.gridlines);
        assert_eq!(sh.active_cell, (1, 2));
        assert!(sh.columns[0].hidden);
        assert_eq!(sh.columns[1].width, Some(72.0));
        assert_eq!(sh.columns[2].width, Some(60.0));
        // Row 1: the sentence and its OK. Row 2, blank: nothing. Row 3, a
        // comment: marked, as Check marks any row with text. Row 4, spaces
        // only: the text kept, no mark. Row 5, a tab: marked, as Trim$ would
        // leave the tab. Row 6: the last line, without a line break after it.
        assert_eq!(text_of(&sh.cells[&(1, 2)]), "Put 5 into cell B2.");
        assert_eq!(text_of(&sh.cells[&(1, 3)]), "OK");
        assert!(!sh.cells.contains_key(&(2, 2)));
        assert_eq!(text_of(&sh.cells[&(3, 2)]), "# a comment");
        assert_eq!(text_of(&sh.cells[&(3, 3)]), "OK");
        assert_eq!(text_of(&sh.cells[&(4, 2)]), "   ");
        assert!(!sh.cells.contains_key(&(4, 3)));
        assert_eq!(text_of(&sh.cells[&(5, 3)]), "OK");
        assert_eq!(text_of(&sh.cells[&(6, 2)]), "last");
        assert_eq!(sh.cells.len(), 9);
        // The three formats, in order of first use, after the default.
        let xfs = styles.xfs();
        assert_eq!(xfs.len(), 4);
        assert_eq!(xfs[1].fill, Some(LAVENDER));
        assert_eq!(xfs[1].num_fmt, NumFmt::Text);
        assert_eq!(xfs[2].fill, Some(RESULT_GRAY));
        assert_eq!(xfs[3].fill, Some(OK_GREEN));
        assert!(xfs[3].wrap);
    }

    /// A program's VLA as EnglishToVla scaffolds it, written by hand.
    const STATIC_VLA: &str = r#"(const rate 0.2)

(dim vla-step Long)

(sub main ()
  (dim total)
  (on-error goto vla-fail)
  (set! vla-step 1)
  (if (vlatraceon) (then (vlatracestep 1 (vla-step-text 1))))
  (at-line 1
  (set! (range "b2") 5))
  (at-line 2
  (set! (. (range "b3") formula) "=B2*2"))
  (at-line 3
  (if (>= 4 2) (then (set! (. (range (& "c" 2 ":" "c" 4)) formula) "=IFS(B2>3,\"big\",TRUE,\"small\")"))))
  (at-line 4
  (begin (vlaensuresheet "data") (. (worksheets "data") activate)))
  (at-line 5
  (set! (range "a1") "west"))
  (at-line 6
  (set! (. (worksheets "output") range "b1") 7))
  (at-line 7
  (set! total 5))
  (at-line 8
  (set! (range "a2") (+ total rate)))
  (at-line 9
  (if (> total 3) (then (at-line 10 (set! (range "a3") true))) (else (at-line 11 (set! (range "a3") false)))))
  (at-line 12
  (set! (. (range "a4:b4") value) (& "x" 1)))
  (at-line 13
  (tally))
  (at-line 14
  (set! (. (range "a5") formula) "plain text"))
  (at-line 15
  (exit-sub))
  (at-line 16
  (set! (range "a6") 1))
  (exit-sub)
  (label vla-fail)
  (vla-report-error))

(sub tally ()
  (on-error goto vla-fail)
  (at-line 20
  (set! (range "a7") 70))
  (at-line 21
  (exit-sub))
  (at-line 22
  (set! (range "a8") 80))
  (exit-sub)
  (label vla-fail)
  (vla-report-error))

(sub vla-report-error ()
  (vlareportstop vla-step (vla-step-text vla-step) err.description))
"#;

    #[test]
    fn the_static_subset_lands_in_cells() {
        let wb = build_workbook("", STATIC_VLA, "", &[]).expect("the static subset builds");
        let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, vec!["Frazaro", "Output", "data"]);
        let out = &wb.sheets[1];
        assert_eq!(out.cells[&(2, 2)].content, Content::Number(5.0));
        assert_eq!(
            out.cells[&(3, 2)].content,
            Content::Formula("B2*2".to_string())
        );
        // IFS can return an array, so the fill is a dynamic-array formula per
        // cell with its references moved (7c), not a shared formula.
        for (row, moved) in [(2, "B2"), (3, "B3"), (4, "B4")] {
            assert_eq!(
                out.cells[&(row, 3)].content,
                Content::DynamicFormula(format!("_xlfn.IFS({moved}>3,\"big\",TRUE,\"small\")"))
            );
        }
        assert_eq!(out.cells[&(1, 2)].content, Content::Number(7.0));
        assert!(wb.has_dynamic_formulas());
        assert_eq!(wb.defined_names.len(), 1);
        assert_eq!(wb.defined_names[0].0, "Frazaro.Build");
        assert!(wb.defined_names[0]
            .1
            .starts_with(&format!("\"Frazaro {}; sentences sha256:", crate::VERSION)));
        let data = &wb.sheets[2];
        assert_eq!(
            data.cells[&(1, 1)].content,
            Content::Text("west".to_string())
        );
        assert_eq!(data.cells[&(2, 1)].content, Content::Number(5.2));
        assert_eq!(data.cells[&(3, 1)].content, Content::Bool(true));
        assert_eq!(data.cells[&(4, 1)].content, Content::Text("x1".to_string()));
        assert_eq!(data.cells[&(4, 2)].content, Content::Text("x1".to_string()));
        // The step inlined up to its own Stop; the main's Stop ends the walk.
        assert_eq!(data.cells[&(7, 1)].content, Content::Number(70.0));
        assert!(!data.cells.contains_key(&(8, 1)));
        assert_eq!(
            data.cells[&(5, 1)].content,
            Content::Text("plain text".to_string())
        );
        assert!(!data.cells.contains_key(&(6, 1)));
    }

    fn vla_with(stmt: &str) -> String {
        format!(
            "(sub main ()\n  (on-error goto vla-fail)\n  (at-line 1\n  (set! (range \"b2\") 5))\n  (at-line 2\n  {stmt})\n  (exit-sub)\n  (label vla-fail)\n  (vla-report-error))\n"
        )
    }

    #[test]
    fn a_formula_that_reaches_outside_is_refused_by_name() {
        // SEC.15: the writer asks the egress scan before it writes a formula,
        // through either spelling of the member and whatever the text begins
        // with; the refusal names the line, the function and the sentence.
        let program = "Put 5 into cell B2.\nPut formula \"=HYPERLINK(\"\"https://example.com/\"\",\"\"go\"\")\" into cell B3.\n";
        for (stmt, name) in [
            (
                "(set! (. (range \"b3\") formula) \"=HYPERLINK(\\\"https://example.com/\\\",\\\"go\\\")\")",
                "HYPERLINK",
            ),
            (
                "(set! (. (range \"b3\") formula) \"=cmd|'/c calc'!A0\")",
                "DDE",
            ),
            (
                "(set! (. (range \"b3\") formula2) \"=_xlfn.WEBSERVICE(\\\"https://example.com/\\\")\")",
                "WEBSERVICE",
            ),
            (
                "(set! (. (range \"b3:b5\") formula) \"WEBSERVICE(1)\")",
                "WEBSERVICE",
            ),
        ] {
            let r = build_workbook(program, &vla_with(stmt), "", &[]).expect_err(stmt);
            assert_eq!(r.id, "build-formula-egress", "{stmt}");
            assert!(
                r.text
                    .contains(&format!("Line 2 would write a formula that uses {name}")),
                "{}",
                r.text
            );
            assert!(r.text.contains("Put formula \"=HYPERLINK"), "{}", r.text);
            assert!(r.text.ends_with("Nothing was written."), "{}", r.text);
        }
        // The same text through the value member is a text in the cell, as
        // Excel stores a value, so nothing is refused and nothing runs.
        let wb = build_workbook(
            program,
            &vla_with(
                "(set! (. (range \"b3\") value) \"=HYPERLINK(\\\"https://example.com/\\\",\\\"go\\\")\")",
            ),
            "",
            &[],
        )
        .expect("a value");
        assert_eq!(
            wb.sheets[1].cells[&(3, 2)].content,
            Content::Text("=HYPERLINK(\"https://example.com/\",\"go\")".to_string())
        );
        // A listed name inside a string literal is text, and the formula is
        // written as before.
        let wb = build_workbook(
            program,
            &vla_with("(set! (. (range \"b3\") formula) \"=\\\"WEBSERVICE(\\\"\")"),
            "",
            &[],
        )
        .expect("a name inside a string");
        assert_eq!(
            wb.sheets[1].cells[&(3, 2)].content,
            Content::Formula("\"WEBSERVICE(\"".to_string())
        );
    }

    #[test]
    fn what_needs_a_host_is_refused_with_its_sentence() {
        let program = "Put 5 into cell B2.\nSay \"hello\".\n";
        for stmt in [
            "(msgbox \"hello\")",
            "(set! (range \"c:c\") 0)",
            "(set! (range \"a1\") (range \"b1\"))",
            "(set! (. (range \"a1\") interior.color) 255)",
            "(if (> (range \"a1\") 1) (then (set! (range \"a2\") 1)))",
            "(dotimes i 3 (set! (range \"a1\") 1))",
            "(on-error goto vla-tryf-1)",
            "(set! x (range \"a1\"))",
        ] {
            let r = build_workbook(program, &vla_with(stmt), "", &[]).expect_err(stmt);
            assert_eq!(r.id, "build-not-representable", "{stmt}");
            assert!(
                r.text.contains("Line 2 asks for more: Say \"hello\"."),
                "{}",
                r.text
            );
        }
        let r = build_workbook(
            program,
            &vla_with("(. (worksheets \"nowhere\") activate)"),
            "",
            &[],
        )
        .expect_err("an unmade sheet");
        assert_eq!(r.id, "build-sheet-unknown");
        assert!(r.text.contains("Line 2 names sheet nowhere"), "{}", r.text);
        // A step that takes values, and a step that calls itself.
        let r = build_workbook(
            program,
            &format!(
                "{}(sub taxed ((byval a Variant))\n  (at-line 5 (set! (range \"a1\") a)))\n",
                vla_with("(taxed 5)")
            ),
            "",
            &[],
        )
        .expect_err("a step with values");
        assert_eq!(r.id, "build-not-representable");
        let r = build_workbook(
            program,
            &format!(
                "{}(sub again ()\n  (at-line 5 (again)))\n",
                vla_with("(again)")
            ),
            "",
            &[],
        )
        .expect_err("a step calling itself");
        assert_eq!(r.id, "build-not-representable");
        assert!(r.text.contains("Line 5"), "{}", r.text);
        // A library the build cannot read, as the compiler says it.
        let r =
            build_workbook(program, "(include \"alien.vla\")\n", "", &[]).expect_err("an include");
        assert_eq!(r.id, "vla-include-cannot-read");
        // A sheet name Excel would refuse.
        for (name, why) in [
            (
                "abcdefghijklmnopqrstuvwxyz123456",
                "it is 32 characters long",
            ),
            ("a:b", "it contains :"),
            ("'q1", "it begins or ends with an apostrophe"),
            ("History", "History is a name Excel keeps for itself"),
        ] {
            let r = build_workbook(
                program,
                &vla_with(&format!("(vlaensuresheet \"{name}\")")),
                "",
                &[],
            )
            .expect_err(name);
            assert_eq!(r.id, "build-sheet-name-invalid", "{name}");
            assert!(r.text.contains(why), "{}", r.text);
        }
        assert_eq!(sheet_name_problem("Q1 Data"), None);
        assert_eq!(sheet_name_problem(""), Some("it is empty".to_string()));
    }

    #[test]
    fn folding_spells_values_as_vba_does() {
        assert_eq!(
            fold_op(
                "&",
                &[
                    Value::Text("c".into()),
                    Value::Num(2.0),
                    Value::Text(":".into())
                ]
            ),
            Some(Value::Text("c2:".into()))
        );
        assert_eq!(cstr(&Value::Num(2.5)), "2.5");
        assert_eq!(cstr(&Value::Num(-3.0)), "-3");
        assert_eq!(cstr(&Value::Bool(true)), "True");
        assert_eq!(fold_op("-", &[Value::Num(5.0)]), Some(Value::Num(-5.0)));
        assert_eq!(fold_op("/", &[Value::Num(5.0), Value::Num(0.0)]), None);
        assert_eq!(
            fold_op("<", &[Value::Text("a".into()), Value::Text("b".into())]),
            Some(Value::Bool(true))
        );
        assert_eq!(fold_op("and", &[Value::Bool(true), Value::Num(1.0)]), None);
        assert_eq!(fold_op("sum", &[Value::Num(1.0)]), None);
    }

    #[test]
    fn the_same_text_gives_the_same_bytes() {
        let vla = "(sub main ()\n  (at-line 1\n  (set! (range \"b2\") 5)))";
        let a = build_xlsx("Put 5 into cell B2.\n", vla, "", &[]).unwrap();
        let b = build_xlsx("Put 5 into cell B2.\n", vla, "", &[]).unwrap();
        assert_eq!(a, b);
        assert!(build_xlsx("Log 1.\n", "(sub main", "", &[]).is_err());
    }

    #[test]
    fn the_refusals_come_from_the_catalogue() {
        let r = raise("build-output-exists", &[("path", "out.xlsx")]);
        assert_eq!(r.source, "VLA-Build");
        assert!(r.text.contains("out.xlsx"));
        assert!(r.text.contains("--replace"));
        let r = raise("build-workbook-too-large", &[]);
        assert_eq!(r.source, "VLA-Build");
        let r = raise(
            "build-not-representable",
            &[("line", "4"), ("sentence", "Say \"hi\".")],
        );
        assert_eq!(r.source, "VLA-Build");
        assert!(r.text.contains("Line 4 asks for more: Say \"hi\"."));
        let r = raise("build-sheet-unknown", &[("line", "4"), ("name", "data")]);
        assert_eq!(r.source, "VLA-Build");
        assert!(r.text.contains("Work on sheet data."));
    }

    #[test]
    fn the_stamp_round_trips() {
        let stamp = Stamp::of(
            "Put 5 into cell B2.\r\n",
            "(defmacro (a) 1)",
            &["book one", "book\ttwo"],
            None,
        );
        assert_eq!(stamp.version, crate::VERSION);
        assert_eq!(stamp.phrasebooks.len(), 2);
        let text = stamp.text();
        assert!(text.starts_with(&format!("Frazaro {}; sentences sha256:", crate::VERSION)));
        assert!(text.contains("; prelude sha256:"));
        assert!(text.contains("; phrasebook 1 sha256:"));
        assert!(text.contains("; phrasebook 2 sha256:"));
        assert!(text.ends_with(" non-whitespace bytes"));
        assert_eq!(Stamp::parse(&text), Some(stamp.clone()));
        assert_eq!(Stamp::parse("Frazaro 0.1.0; sentences nonsense"), None);
        assert_eq!(Stamp::parse(""), None);
        // The same sentences in another spelling of their whitespace agree.
        assert_eq!(
            stamp.disagreement(
                "Put 5 into cell B2.",
                "(defmacro (a) 1)",
                &["book one", "book two"]
            ),
            None
        );
        assert_eq!(
            stamp.disagreement(
                "Put 6 into cell B2.",
                "(defmacro (a) 1)",
                &["book one", "book two"]
            ),
            Some("the sentences in its Frazaro sheet are not the ones the stamp names".to_string())
        );
        assert_eq!(
            stamp.disagreement(
                "Put 5 into cell B2.",
                "(defmacro (a) 2)",
                &["book one", "book two"]
            ),
            Some("the prelude given is not the one the stamp names".to_string())
        );
        assert_eq!(
            stamp.disagreement("Put 5 into cell B2.", "(defmacro (a) 1)", &["book one"]),
            Some("the stamp names 2 phrasebook(s) and 1 were given".to_string())
        );
        assert_eq!(
            stamp.disagreement(
                "Put 5 into cell B2.",
                "(defmacro (a) 1)",
                &["book one", "book three"]
            ),
            Some("phrasebook 2 given is not the one the stamp names".to_string())
        );
    }

    #[test]
    fn a_built_workbook_reads_back() {
        let program = " Put 5 into cell B2.\n\n# a & b <c>\n   \nlast";
        let vla = "(sub main ()\n  (at-line 1\n  (set! (range \"b2\") 5)))";
        let bytes = build_xlsx(program, vla, "pre", &["book"]).unwrap();
        let read = read_build(&bytes).expect("a build reads back");
        assert_eq!(
            read.program_text,
            " Put 5 into cell B2.\n\n# a & b <c>\n   \nlast"
        );
        assert_eq!(read.sentences, 3);
        assert_eq!(read.stamp, Stamp::of(program, "pre", &["book"], None));
        assert!(read.host_sheets.is_empty());
        assert_eq!(
            read.stamp
                .disagreement(&read.program_text, "pre", &["book"]),
            None
        );
        // Not a build: not a zip; a zip with no stamp.
        let r = read_build(b"not a workbook at all").expect_err("not a zip");
        assert_eq!(r.id, "rebuild-not-a-build");
        assert!(r.text.contains("not a zip archive"));
        let plain =
            zip::write_stored(&[("xl/workbook.xml".to_string(), b"<workbook/>".to_vec())]).unwrap();
        let r = read_build(&plain).expect_err("no stamp");
        assert_eq!(r.id, "rebuild-not-a-build");
        assert!(r.text.contains("no Frazaro.Build stamp"), "{}", r.text);
        // The answer's one line.
        let yes = Rebuilt {
            sentences: 9,
            built_by: "0.7.1".to_string(),
            matches: true,
            why: String::new(),
            partial: false,
        };
        assert_eq!(
            yes.line(),
            "This workbook was built from these 9 sentences by Frazaro 0.7.1: yes."
        );
        let no = Rebuilt {
            why: "the prelude given is not the one the stamp names".to_string(),
            matches: false,
            ..yes
        };
        assert_eq!(
            no.line(),
            "This workbook was built from these 9 sentences by Frazaro 0.7.1: no. Why: the prelude given is not the one the stamp names."
        );
    }

    #[test]
    fn rebuild_says_yes_to_its_own_build_and_no_to_a_changed_one() {
        let r = crate::api::english_rebuild_xlsx(GOLDEN, PRELUDE, &[ENGLISH])
            .unwrap_or_else(|r| panic!("the golden did not read back: {}", r.refusal));
        assert!(r.matches, "{}", r.why);
        assert_eq!(r.sentences, 9);
        assert_eq!(r.built_by, crate::VERSION);
        let mut changed = GOLDEN.to_vec();
        changed[100] ^= 1; // inside [Content_Types].xml, the first part
        let r = crate::api::english_rebuild_xlsx(&changed, PRELUDE, &[ENGLISH]).unwrap();
        assert!(!r.matches);
        assert!(r.why.contains("not what this core builds"), "{}", r.why);
        let r = crate::api::english_rebuild_xlsx(GOLDEN, "", &[ENGLISH]).unwrap();
        assert!(!r.matches);
        assert_eq!(r.why, "the prelude given is not the one the stamp names");
    }

    #[test]
    fn a_build_into_a_model_leaves_the_model_alone() {
        use crate::sheet::merge::test_model;
        // Writes to Output (new) and Checks (new) land; the model's Model
        // and Notes may be gone to but never written into.
        let vla = "(sub main ()\n  (at-line 1\n  (set! (range \"b1\") 5))\n  (at-line 2\n  (begin (vlaensuresheet \"checks\") (. (worksheets \"checks\") activate)))\n  (at-line 3\n  (set! (. (range \"a1\") formula) \"=IFS(Model!B3>0,\\\"profit\\\",TRUE,\\\"loss\\\")\"))\n  (at-line 4\n  (begin (vlaensuresheet \"notes\") (. (worksheets \"notes\") activate))))";
        let program = "Put 5 into cell B1.\nWork on sheet Checks.\nPut formula \"=IFS(...)\" into cell A1.\nWork on sheet Notes.\n";
        let out =
            build_xlsx_into(program, vla, "", &["book"], test_model(None), "model.xlsx").unwrap();
        let read = read_build(&out).expect("the output reads back as a build");
        assert_eq!(read.host_sheets, ["Model", "Notes"]);
        assert_eq!(read.sentences, 4);
        let into = read.stamp.into.clone().expect("an into stamp");
        assert_eq!(into.style_base, 3);
        assert_eq!(into.cm, 1);
        assert_eq!(into.model_hex, crate::sha256::sha256_hex(&test_model(None)));
        let names: Vec<String> = zip::entries(&out)
            .unwrap()
            .into_iter()
            .map(|e| e.name)
            .collect();
        assert!(names.contains(&"xl/worksheets/frazaro_1.xml".to_string()));
        assert!(names.contains(&"xl/worksheets/frazaro_3.xml".to_string()));
        assert!(names.contains(&"xl/metadata.xml".to_string()));
        // A write into the model's own sheet is refused with its line.
        let bad = "(sub main ()\n  (at-line 1\n  (. (worksheets \"model\") activate))\n  (at-line 2\n  (set! (range \"a5\") 7)))";
        let r = build_xlsx_into(
            "Go to sheet Model.\nPut 7 into cell A5.\n",
            bad,
            "",
            &[],
            test_model(None),
            "model.xlsx",
        )
        .expect_err("a write into the model");
        assert_eq!(r.id, "build-into-model-sheet");
        assert!(
            r.text.contains("Line 2 writes into sheet Model"),
            "{}",
            r.text
        );
        // So is a cross-sheet put into it.
        let bad2 =
            "(sub main ()\n  (at-line 1\n  (set! (. (worksheets \"notes\") range \"a1\") 7)))";
        let r = build_xlsx_into(
            "Put 7 into cell A1 of sheet Notes.\n",
            bad2,
            "",
            &[],
            test_model(None),
            "model.xlsx",
        )
        .expect_err("a put into the model");
        assert_eq!(r.id, "build-into-model-sheet");
        assert!(r.text.contains("sheet Notes"), "{}", r.text);
        // A model that already has a Frazaro sheet cannot take another.
        let taken = String::from_utf8(test_model(None)).unwrap_or_default();
        let _ = taken;
        let with_frazaro = {
            let bytes = test_model(None);
            let m = Model::read(bytes, "m.xlsx").unwrap();
            let host = HostInfo {
                sheet_names: vec!["Model".to_string(), "FRAZARO".to_string()],
                ..m.host_info()
            };
            build_workbook_with(
                "Put 1 into cell A1.\n",
                "(sub main ()\n  (at-line 1\n  (set! (range \"a1\") 1)))",
                "",
                &[],
                Some(&host),
            )
        };
        let r = with_frazaro.expect_err("the name is taken");
        assert_eq!(r.id, "build-into-sheet-name-taken");
        assert!(
            r.text.contains("already has a sheet named FRAZARO"),
            "{}",
            r.text
        );
        // Not a model at all.
        let r = build_xlsx_into("", "(sub main ())", "", &[], b"nope".to_vec(), "x.xlsx")
            .expect_err("not a zip");
        assert_eq!(r.id, "build-into-not-a-workbook");
    }

    #[test]
    fn the_into_stamp_round_trips() {
        let into = IntoStamp {
            model_hex: "AB".repeat(32),
            style_base: 7,
            cm: 2,
        };
        let stamp = Stamp::of("x", "p", &["b"], Some(into.clone()));
        let text = stamp.text();
        assert!(text.contains(&format!(
            "; into sha256:{} with 7 cell formats and metadata 2; sentences ",
            "AB".repeat(32)
        )));
        assert_eq!(Stamp::parse(&text), Some(stamp));
        assert_eq!(Stamp::parse(&text.replace(" with 7 ", " with x ")), None);
        let partial = Rebuilt {
            sentences: 6,
            built_by: "0.7.1".to_string(),
            matches: true,
            why: String::new(),
            partial: true,
        };
        assert_eq!(
            partial.line(),
            "This workbook's Frazaro sheets were built from these 6 sentences by Frazaro 0.7.1, into a workbook whose own sheets are not checked: yes."
        );
    }

    /// The build golden: the fixture, built with the prelude and english.vla,
    /// is `scripts/build/fixture_golden.xlsx` byte for byte. When this fails
    /// the writer changed: regenerate with `frazaro build
    /// scripts/build/fixture.txt --prelude scripts/prelude.vla --phrasebook
    /// scripts/polyglotta/english.vla --out scripts/build/fixture_golden.xlsx
    /// --replace`, open it in Excel (the owner's live pass), and raise
    /// tools/check_build_golden.ps1's floor if it grew.
    #[test]
    fn the_build_golden_is_reproduced() {
        let got = crate::api::english_build_xlsx(FIXTURE, PRELUDE, &[ENGLISH])
            .unwrap_or_else(|r| panic!("the fixture did not build: {}", r.refusal));
        if got != GOLDEN {
            let at = got
                .iter()
                .zip(GOLDEN.iter())
                .position(|(a, b)| a != b)
                .unwrap_or(got.len().min(GOLDEN.len()));
            let part = zip::entries(GOLDEN)
                .unwrap_or_default()
                .into_iter()
                .find(|e| at >= e.data_offset && at < e.data_offset + e.size as usize)
                .map(|e| format!("{} at offset {}", e.name, at - e.data_offset))
                .unwrap_or_else(|| "a header".to_string());
            panic!(
                "the build differs from scripts/build/fixture_golden.xlsx at byte {at} of {} ({part}); got {} bytes",
                GOLDEN.len(),
                got.len()
            );
        }
    }
}
