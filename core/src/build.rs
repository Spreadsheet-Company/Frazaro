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
use crate::sheet::xlfn::prefix_future_functions;
use crate::sheet::{
    ooxml, parse_a1_range, A1Range, Cell, Column, Content, NumFmt, Rgb, Sheet, Style, Styles,
    Workbook,
};

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

struct Walker<'a> {
    lines: Vec<&'a str>,
    wb: Workbook,
    /// The sheet a `(range ...)` means: `None` until something writes, then
    /// `Output` unless `Work on sheet` chose another.
    current: Option<usize>,
    /// Constants and variables whose values fold, by folded name.
    bindings: HashMap<String, Value>,
    /// The program's procedures, by folded name.
    steps: HashMap<String, Step>,
    /// The program line the walk is on, from the nearest `(at-line N ...)`.
    line: u32,
    depth: u32,
}

impl<'a> Walker<'a> {
    fn new(program_text: &'a str) -> Walker<'a> {
        Walker {
            lines: program_text.lines().collect(),
            wb: Workbook::new(),
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
                    let sheet = self.current_sheet();
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
                        let sheet = self.current_sheet();
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
                        let sheet = self.sheet_named(&name)?;
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
                let sheet = self.sheet_named(&name)?;
                self.current = Some(sheet);
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

    /// The sheet `(range ...)` writes to, `Output` made on first use.
    fn current_sheet(&mut self) -> usize {
        match self.current {
            Some(i) => i,
            None => {
                let i = self.wb.ensure_sheet(OUTPUT_SHEET);
                self.current = Some(i);
                i
            }
        }
    }

    /// A sheet named by the program: one it made, or `Output`, which the
    /// add-in has at the start of every run.
    fn sheet_named(&mut self, name: &str) -> Result<usize, Refusal> {
        if let Some(i) = self.wb.find_sheet(name) {
            return Ok(i);
        }
        if fold(name) == fold(OUTPUT_SHEET) {
            return Ok(self.wb.ensure_sheet(OUTPUT_SHEET));
        }
        Err(self.refuse_sheet(name))
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
    /// itself, as Excel stores it.
    fn write_formula(&mut self, sheet: usize, range: A1Range, v: &Value) -> Result<(), Refusal> {
        if range.cells() > MAX_CELLS_PER_WRITE {
            return Err(self.refuse_sentence());
        }
        if let Value::Text(t) = v {
            if let Some(body) = t.strip_prefix('=') {
                let text = prefix_future_functions(body);
                self.wb.sheets[sheet].set_formula(range, &text, 0);
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

/// The workbook a translated program builds: the `Frazaro` sheet, then
/// the sheets its static subset writes. `vla` is the program's VLA as
/// `EnglishToVla` wrote it, read with the prelude and expanded as the
/// compiler reads it.
pub fn build_workbook(program_text: &str, vla: &str, prelude: &str) -> Result<Workbook, Refusal> {
    let line_offset = count_lf(prelude) as u32 + 1;
    let text = format!("{prelude}\r\n{vla}");
    let forms = parse_all(&tokenize(&text), line_offset)?;
    let mut expander = Expander::new();
    let body = expander.collect(forms)?;
    let mut expanded = Vec::with_capacity(body.len());
    for f in &body {
        expanded.push(expander.expand(f, 0)?);
    }
    let mut w = Walker::new(program_text);
    let sheet = frazaro_sheet(program_text, &mut w.wb.styles);
    w.wb.sheets.push(sheet);
    w.wb.active_sheet = 0;
    for f in &expanded {
        w.collect_top(f)?;
    }
    w.walk_main()?;
    Ok(w.wb)
}

/// The workbook's bytes, or the catalogue's refusal when they would not fit
/// the file format.
pub fn build_xlsx(program_text: &str, vla: &str, prelude: &str) -> Result<Vec<u8>, Refusal> {
    let wb = build_workbook(program_text, vla, prelude)?;
    ooxml::workbook_bytes(&wb).ok_or_else(|| raise("build-workbook-too-large", &[]))
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
        let wb = build_workbook("", STATIC_VLA, "").expect("the static subset builds");
        let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, vec!["Frazaro", "Output", "data"]);
        let out = &wb.sheets[1];
        assert_eq!(out.cells[&(2, 2)].content, Content::Number(5.0));
        assert_eq!(
            out.cells[&(3, 2)].content,
            Content::Formula("B2*2".to_string())
        );
        assert_eq!(
            out.cells[&(2, 3)].content,
            Content::SharedMaster {
                text: "_xlfn.IFS(B2>3,\"big\",TRUE,\"small\")".to_string(),
                range: "C2:C4".to_string(),
                si: 0
            }
        );
        assert_eq!(out.cells[&(3, 3)].content, Content::SharedChild { si: 0 });
        assert_eq!(out.cells[&(4, 3)].content, Content::SharedChild { si: 0 });
        assert_eq!(out.cells[&(1, 2)].content, Content::Number(7.0));
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
            let r = build_workbook(program, &vla_with(stmt), "").expect_err(stmt);
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
        )
        .expect_err("a step calling itself");
        assert_eq!(r.id, "build-not-representable");
        assert!(r.text.contains("Line 5"), "{}", r.text);
        // A library the build cannot read, as the compiler says it.
        let r = build_workbook(program, "(include \"alien.vla\")\n", "").expect_err("an include");
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
        let a = build_xlsx("Put 5 into cell B2.\n", vla, "").unwrap();
        let b = build_xlsx("Put 5 into cell B2.\n", vla, "").unwrap();
        assert_eq!(a, b);
        assert!(build_xlsx("Log 1.\n", "(sub main", "").is_err());
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
