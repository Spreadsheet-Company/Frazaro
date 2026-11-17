//! A function of cells (KERNEL.29's third slice): a room of cells on one
//! sheet, compiled to the `LAMBDA` a person would write from it, so that a
//! call computes what the room computes at the call's arguments.
//!
//! A `function` row names the function, the sheet its body stands on, its
//! inputs, read row by row as its parameters, and its output, a cell or a
//! range answered as an array: `(function "SUMMA" "VIII SUMMA" "A3:A5"
//! "J3")`. The compile walks the room once from the output. An input is its
//! parameter; a formula cell the walk reaches is a gate, bound by a `LET` to
//! its own formula with every reference to the room rewritten; any other
//! cell of the room is a reference kept, its sheet named and its place
//! absolute, as a name's references are; and a reference to another sheet
//! is kept as written. VIII SUMMA's sum is
//! `=LAMBDA(_A3,_A4,_A5,LET(_C3,MOD(_A3+_A4,2),MOD(_C3+_A5,2)))`.
//!
//! The text is the grid's named function, inlined at every call as a `name`
//! row's `LAMBDA` is (`super::lambda`). A body that does not compile (a gate
//! reading itself, a range over an input or a gate, a reference no `LAMBDA`
//! writes from the room, more inputs than a `LAMBDA` takes, a text past a
//! formula's limits) is held with why: its calls answer `#CALC!`, and the
//! census says why.

use std::collections::{HashMap, HashSet};

use crate::grid::{text_of, Formula, Masters, Model};
use crate::intrinsics::fold;
use crate::refers::{self, FormulaRef, Kind};
use crate::rows::quoted;
use crate::sheet::{cell_ref, column_letters, A1Range, MAX_COLUMN, MAX_ROW};

use super::formula::{self, Expr};
use super::graph::Area;
use super::lambda::{LAMBDA_PARAMS, LET_PAIRS};

/// A function of cells as a grid holds it: its name as written, the sheet its
/// body stands on, its inputs and its output.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Function {
    pub(crate) name: String,
    pub(crate) sheet: usize,
    pub(crate) inputs: A1Range,
    pub(crate) output: A1Range,
}

/// A function compiled: its name, how many parameters it takes, and its
/// `LAMBDA`'s text, `=` first, or why its body does not compile.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Compiled {
    pub(crate) name: String,
    pub(crate) params: usize,
    pub(crate) lambda: Result<String, Uncompiled>,
}

impl Compiled {
    /// A function a write has just defined, compiled at the write's settle.
    pub(crate) fn pending(name: &str) -> Compiled {
        Compiled {
            name: name.to_string(),
            params: 0,
            lambda: Err(Uncompiled::Text(String::new())),
        }
    }
}

/// Why a body does not compile, each as the census says it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum Uncompiled {
    /// A gate reading itself through the gates, by its cell.
    Cycle(String),
    /// A range, a whole column or a whole row of the room holding an input
    /// or a gate, as written: a `LAMBDA` reads its parameters one at a time.
    Range(String),
    /// A reference no `LAMBDA` writes from the room, its kind's word: a
    /// spill of the room, a table's column, a span of sheets, another
    /// workbook, `INDIRECT` or `OFFSET`.
    Construct(&'static str),
    /// More inputs than a `LAMBDA` takes parameters.
    Inputs(usize),
    /// A text past what a formula holds, in the parse's words.
    Text(String),
}

impl Uncompiled {
    /// Why, as the census's `uncompiled` row spells it after the name.
    pub(crate) fn words(&self) -> String {
        match self {
            Uncompiled::Cycle(cell) => format!("cycle {}", quoted(cell)),
            Uncompiled::Range(written) => format!("range {}", quoted(written)),
            Uncompiled::Construct(word) => format!("construct {word}"),
            Uncompiled::Inputs(n) => format!("inputs {n}"),
            Uncompiled::Text(why) => format!("text {}", quoted(why)),
        }
    }
}

/// A function's room compiled against a grid.
pub(crate) fn compile<M: Model + ?Sized>(model: &M, f: &Function) -> Compiled {
    let params = f.inputs.cells() as usize;
    let lambda = if params > LAMBDA_PARAMS {
        Err(Uncompiled::Inputs(params))
    } else {
        Room { model, f }.lambda()
    };
    Compiled {
        name: f.name.clone(),
        params,
        lambda,
    }
}

/// What a reference in a gate reads, as the compile writes it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Read {
    /// A cell of the room: a parameter, a gate's binding or a reference kept.
    Cell(u32, u32),
    /// A range, a whole column or a whole row of the room that holds no
    /// input and no gate, kept with its sheet named and absolute.
    Area,
    /// Anything else, kept as written: a reference to another sheet, a
    /// name, a broken reference.
    Other,
}

/// A gate the walk reached: its cell, its text, and what each reference in
/// it reads, in the order the scan found them.
struct Gate {
    at: (u32, u32),
    text: String,
    refs: Vec<(FormulaRef, Read)>,
}

struct Room<'a, M: Model + ?Sized> {
    model: &'a M,
    f: &'a Function,
}

impl<M: Model + ?Sized> Room<'_, M> {
    fn lambda(&self) -> Result<String, Uncompiled> {
        let gates = self.reach()?;
        let prefix = prefix(gates.iter().map(|g| g.text.as_str()));
        let bound: HashSet<(u32, u32)> = gates.iter().map(|g| g.at).collect();
        let named = |(row, col): (u32, u32)| format!("{prefix}{}", cell_ref(row, col));
        let params: Vec<String> = cells_of(self.f.inputs).into_iter().map(named).collect();
        let mut binds: Vec<(String, String)> = gates
            .iter()
            .map(|g| (named(g.at), self.render(g, &prefix, &bound)))
            .collect();
        let outputs = cells_of(self.f.output);
        let body = match outputs.as_slice() {
            // One output that is a gate: its own formula the body, the last
            // of the walk's order.
            [one] if bound.contains(one) && !self.is_input(*one) => {
                binds.pop().map(|(_, text)| text).unwrap_or_default()
            }
            _ => {
                let width = (self.f.output.right - self.f.output.left + 1) as usize;
                let spelled: Vec<String> = outputs
                    .iter()
                    .map(|&at| self.cell_text(at, &prefix, &bound))
                    .collect();
                stacked(&spelled, width)
            }
        };
        let text = assembled(&params, &binds, &body);
        match formula::parse(&text) {
            Ok(Expr::Lambda(_)) => Ok(text),
            Ok(_) => Err(Uncompiled::Text(String::new())),
            Err(why) => Err(Uncompiled::Text(why)),
        }
    }

    fn sheet_name(&self) -> &str {
        self.model.sheet_name(self.f.sheet).unwrap_or("")
    }

    fn is_input(&self, (row, col): (u32, u32)) -> bool {
        let r = &self.f.inputs;
        row >= r.top && row <= r.bottom && col >= r.left && col <= r.right
    }

    /// The gates the outputs reach, each after every gate it reads: a walk
    /// of the room held on a stack of its own, so that a long chain of gates
    /// takes no frame of the machine's.
    fn reach(&self) -> Result<Vec<Gate>, Uncompiled> {
        // 1 while a gate's reads are walked, 2 once it is in the order.
        let mut state: HashMap<(u32, u32), u8> = HashMap::new();
        let mut order: Vec<Gate> = Vec::new();
        let mut stack: Vec<(Gate, usize)> = Vec::new();
        for at in cells_of(self.f.output) {
            self.enter(at, &mut state, &mut stack)?;
            while let Some((gate, k)) = stack.last_mut() {
                let next = gate.refs.get(*k).map(|(_, read)| *read);
                *k += 1;
                match next {
                    Some(Read::Cell(row, col)) => self.enter((row, col), &mut state, &mut stack)?,
                    Some(_) => {}
                    None => {
                        let (gate, _) = stack.pop().expect("a gate on the stack");
                        state.insert(gate.at, 2);
                        order.push(gate);
                    }
                }
            }
        }
        Ok(order)
    }

    /// A cell the walk meets: a gate not yet walked is put on the stack, its
    /// references read; a gate on the stack already reads itself.
    fn enter(
        &self,
        at: (u32, u32),
        state: &mut HashMap<(u32, u32), u8>,
        stack: &mut Vec<(Gate, usize)>,
    ) -> Result<(), Uncompiled> {
        if self.is_input(at) {
            return Ok(());
        }
        match state.get(&at) {
            Some(2) => return Ok(()),
            Some(_) => return Err(Uncompiled::Cycle(cell_ref(at.0, at.1))),
            None => {}
        }
        let Some(text) = self.text_at(at) else {
            return Ok(());
        };
        let mut refs = Vec::new();
        for r in refers::scan(&text) {
            let read = self.read(&r)?;
            refs.push((r, read));
        }
        state.insert(at, 1);
        stack.push((Gate { at, text, refs }, 0));
        Ok(())
    }

    /// A cell's formula as the formula bar shows it, `None` where it holds
    /// none.
    fn text_at(&self, (row, col): (u32, u32)) -> Option<String> {
        let formula = self.model.formula((self.f.sheet, row, col))?;
        let mut masters = Masters::new();
        if let Formula::Fill(si) = formula {
            if let Some(master) = self.model.master(self.f.sheet, si) {
                masters.insert(si, master);
            }
        }
        text_of(formula, (row, col), &masters)
    }

    /// What a reference reads, refused where no `LAMBDA` writes it.
    fn read(&self, r: &FormulaRef) -> Result<Read, Uncompiled> {
        match r.kind {
            Kind::Structured => return Err(Uncompiled::Construct("structured")),
            Kind::External => return Err(Uncompiled::Construct("external")),
            Kind::ThreeD => return Err(Uncompiled::Construct("3d")),
            Kind::Unreadable => return Err(Uncompiled::Construct(r.unreadable_name())),
            Kind::Name | Kind::Broken => return Ok(Read::Other),
            Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::Spill => {}
        }
        if !self.on_room(r) {
            return Ok(Read::Other);
        }
        match r.kind {
            Kind::Cell => Ok(Read::Cell(r.row1, r.col1)),
            Kind::Spill => Err(Uncompiled::Construct("spill")),
            _ if self.meets(r) => Err(Uncompiled::Range(r.written.clone())),
            _ => Ok(Read::Area),
        }
    }

    /// Whether a reference names the room: a bare one, or one naming its
    /// sheet, of this workbook.
    fn on_room(&self, r: &FormulaRef) -> bool {
        r.book.is_empty()
            && (r.sheet_name.is_empty()
                || self.model.find_sheet(&r.sheet_name) == Some(self.f.sheet))
    }

    /// Whether a range, a whole column or a whole row of the room holds an
    /// input or a gate.
    fn meets(&self, r: &FormulaRef) -> bool {
        let area = rect_of(r);
        let i = &self.f.inputs;
        let crosses = area.top <= i.bottom
            && i.top <= area.bottom
            && area.left <= i.right
            && i.left <= area.right;
        if crosses {
            return true;
        }
        let mut found = false;
        let area = Area {
            sheet: self.f.sheet,
            top: area.top,
            left: area.left,
            bottom: area.bottom,
            right: area.right,
        };
        let whole = Some(((1, 1), (MAX_ROW, MAX_COLUMN)));
        self.model
            .formulas_in(&area, whole, &mut |_, _| found = true);
        found
    }

    /// A gate's formula with every reference to the room rewritten.
    fn render(&self, g: &Gate, prefix: &str, bound: &HashSet<(u32, u32)>) -> String {
        let chars: Vec<char> = g.text.chars().collect();
        let mut out = String::with_capacity(g.text.len() + 16);
        let mut p = 0;
        for (r, read) in &g.refs {
            let with = match read {
                Read::Cell(row, col) => self.cell_text((*row, *col), prefix, bound),
                Read::Area => self.area_text(r),
                Read::Other => continue,
            };
            out.extend(&chars[p..r.pos]);
            out.push_str(&with);
            p = r.pos + r.span;
        }
        out.extend(&chars[p..]);
        out
    }

    /// A cell of the room as a body reads it: an input's parameter or a
    /// gate's binding by its name, any other cell kept, its sheet named and
    /// its place absolute.
    fn cell_text(&self, at: (u32, u32), prefix: &str, bound: &HashSet<(u32, u32)>) -> String {
        if self.is_input(at) || bound.contains(&at) {
            return format!("{prefix}{}", cell_ref(at.0, at.1));
        }
        format!(
            "{}!${}${}",
            refers::quote_sheet(self.sheet_name()),
            column_letters(at.1),
            at.0
        )
    }

    /// A range, a whole column or a whole row of the room kept, its sheet
    /// named and its places absolute.
    fn area_text(&self, r: &FormulaRef) -> String {
        let a = rect_of(r);
        let place = match r.kind {
            Kind::Column => format!("${}:${}", column_letters(a.left), column_letters(a.right)),
            Kind::Row => format!("${}:${}", a.top, a.bottom),
            _ => format!(
                "${}${}:${}${}",
                column_letters(a.left),
                a.top,
                column_letters(a.right),
                a.bottom
            ),
        };
        format!("{}!{place}", refers::quote_sheet(self.sheet_name()))
    }
}

/// The rectangle a reference with numbers names, a whole column running to
/// the last row and a whole row to the last column.
fn rect_of(r: &FormulaRef) -> A1Range {
    match r.kind {
        Kind::Column => A1Range {
            top: 1,
            left: r.col1.min(r.col2),
            bottom: MAX_ROW,
            right: r.col1.max(r.col2),
        },
        Kind::Row => A1Range {
            top: r.row1.min(r.row2),
            left: 1,
            bottom: r.row1.max(r.row2),
            right: MAX_COLUMN,
        },
        _ => A1Range {
            top: r.row1.min(r.row2),
            left: r.col1.min(r.col2),
            bottom: r.row1.max(r.row2),
            right: r.col1.max(r.col2),
        },
    }
}

/// A range's cells row by row.
fn cells_of(r: A1Range) -> Vec<(u32, u32)> {
    let mut out = Vec::with_capacity(r.cells() as usize);
    for row in r.top..=r.bottom {
        for col in r.left..=r.right {
            out.push((row, col));
        }
    }
    out
}

/// The shortest run of underscores no name in the gates' texts begins with,
/// so that a parameter or a binding named by it and its cell meets none:
/// `_` but where a text names something `_` already.
fn prefix<'t>(texts: impl Iterator<Item = &'t str>) -> String {
    let mut words: Vec<String> = Vec::new();
    for text in texts {
        words.extend(identifiers(text).into_iter().map(|w| fold(&w)));
    }
    let mut prefix = String::from("_");
    while words.iter().any(|w| w.starts_with(&prefix)) {
        prefix.push('_');
    }
    prefix
}

/// The words a formula's text holds outside its texts and its quoted
/// sheets, each a run of a name's characters that begins with a letter, an
/// underscore or a backslash; a number's run is passed over whole.
fn identifiers(text: &str) -> Vec<String> {
    let s: Vec<char> = text.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < s.len() {
        let c = s[i];
        if c == '"' || c == '\'' {
            i = refers::skip_delimited(&s, i, c);
        } else if c.is_alphabetic() || c == '_' || c == '\\' {
            let start = i;
            i += 1;
            while i < s.len() && refers::is_ident_char(s[i]) {
                i += 1;
            }
            out.push(s[start..i].iter().collect());
        } else if c.is_ascii_digit() {
            while i < s.len() && refers::is_ident_char(s[i]) {
                i += 1;
            }
        } else {
            i += 1;
        }
    }
    out
}

/// The outputs' values as one array, row by row: `HSTACK` a row's, `VSTACK`
/// the rows'; one output alone as itself.
fn stacked(spelled: &[String], width: usize) -> String {
    let rows: Vec<String> = spelled
        .chunks(width.max(1))
        .map(|row| match row {
            [one] => one.clone(),
            many => format!("HSTACK({})", many.join(",")),
        })
        .collect();
    match rows.as_slice() {
        [one] => one.clone(),
        many => format!("VSTACK({})", many.join(",")),
    }
}

/// `=LAMBDA(<params>,LET(<bindings>,<body>))`, a `LET` of at most Excel's
/// pairs, the rest in a `LET` inside it; no `LET` where nothing is bound.
fn assembled(params: &[String], binds: &[(String, String)], body: &str) -> String {
    let mut text = String::from("=LAMBDA(");
    for p in params {
        text.push_str(p);
        text.push(',');
    }
    let mut closing = 1;
    for chunk in binds.chunks(LET_PAIRS) {
        text.push_str("LET(");
        for (name, value) in chunk {
            text.push_str(name);
            text.push(',');
            text.push_str(value);
            text.push(',');
        }
        closing += 1;
    }
    text.push_str(body);
    for _ in 0..closing {
        text.push(')');
    }
    text
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::machine::read_rows;
    use crate::reader::read_forms;
    use crate::sheet::{parse_a1_range, Workbook};

    /// A grid's workbook from its rows.
    fn grid(rows: &str) -> Workbook {
        read_rows(&read_forms(rows).unwrap()).unwrap().workbook()
    }

    fn function(wb: &Workbook, sheet: &str, inputs: &str, output: &str) -> Function {
        Function {
            name: "F".to_string(),
            sheet: wb.find_sheet(sheet).unwrap(),
            inputs: parse_a1_range(inputs).unwrap(),
            output: parse_a1_range(output).unwrap(),
        }
    }

    const SUMMA: &str = r#"(cell "VIII SUMMA" "A3:A5" 0)
(formula "VIII SUMMA" "C3" "=MOD(A3+A4,2)")
(formula "VIII SUMMA" "C4" "=A3*A4")
(formula "VIII SUMMA" "C5" "=C3*A5")
(formula "VIII SUMMA" "J3" "=MOD(C3+A5,2)")
(formula "VIII SUMMA" "J4" "=MAX(C4,C5)")
"#;

    #[test]
    fn the_full_adder_compiles_to_the_lambda_a_person_writes() {
        let wb = grid(SUMMA);
        let sum = compile(&wb, &function(&wb, "VIII SUMMA", "A3:A5", "J3"));
        assert_eq!(
            sum.lambda.as_deref(),
            Ok("=LAMBDA(_A3,_A4,_A5,LET(_C3,MOD(_A3+_A4,2),MOD(_C3+_A5,2)))")
        );
        assert_eq!(sum.params, 3);
        let carry = compile(&wb, &function(&wb, "VIII SUMMA", "A3:A5", "J4"));
        assert_eq!(
            carry.lambda.as_deref(),
            Ok("=LAMBDA(_A3,_A4,_A5,LET(_C4,_A3*_A4,_C3,MOD(_A3+_A4,2),_C5,_C3*_A5,MAX(_C4,_C5)))")
        );
        // Both outputs as one array, a column of two.
        let both = compile(&wb, &function(&wb, "VIII SUMMA", "A3:A5", "J3:J4"));
        assert_eq!(
            both.lambda.as_deref(),
            Ok("=LAMBDA(_A3,_A4,_A5,LET(_C3,MOD(_A3+_A4,2),_J3,MOD(_C3+_A5,2),_C4,_A3*_A4,_C5,_C3*_A5,_J4,MAX(_C4,_C5),VSTACK(_J3,_J4)))")
        );
    }

    #[test]
    fn a_cell_neither_input_nor_gate_is_kept_and_another_sheets_as_written() {
        let wb = grid(
            r#"(cell "Room" "A1" 1)
(cell "Room" "B1" 7)
(cell "Other" "A1" 2)
(formula "Room" "C1" "=A1*B1+Other!A1+SUM(B1:B2)+'Room'!A1")
"#,
        );
        let f = compile(&wb, &function(&wb, "Room", "A1", "C1"));
        assert_eq!(
            f.lambda.as_deref(),
            Ok("=LAMBDA(_A1,_A1*Room!$B$1+Other!A1+SUM(Room!$B$1:$B$2)+_A1)")
        );
        // The output an input, a value, or a cell holding nothing.
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "A1"))
                .lambda
                .as_deref(),
            Ok("=LAMBDA(_A1,_A1)")
        );
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "D9"))
                .lambda
                .as_deref(),
            Ok("=LAMBDA(_A1,Room!$D$9)")
        );
        // A row of outputs, and a block of them, row by row.
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "A1:B1"))
                .lambda
                .as_deref(),
            Ok("=LAMBDA(_A1,HSTACK(_A1,Room!$B$1))")
        );
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "A1:B2"))
                .lambda
                .as_deref(),
            Ok("=LAMBDA(_A1,VSTACK(HSTACK(_A1,Room!$B$1),HSTACK(Room!$A$2,Room!$B$2)))")
        );
    }

    #[test]
    fn a_name_beginning_with_an_underscore_lengthens_the_prefix() {
        let wb = grid(
            r#"(cell "Room" "A1" 1)
(formula "Room" "B1" "=LET(_A1,2,_A1+A1)")
"#,
        );
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "B1"))
                .lambda
                .as_deref(),
            Ok("=LAMBDA(__A1,LET(_A1,2,_A1+__A1))")
        );
    }

    #[test]
    fn a_body_no_lambda_writes_says_why() {
        let wb = grid(
            r#"(cell "Room" "A1:A3" 1)
(cell "Room" "D1" 5)
(formula "Room" "B1" "=SUM(A1:A3)")
(formula "Room" "B2" "=SUM(B1:B1)")
(formula "Room" "B3" "=SUM(D1:D4)")
(formula "Room" "B4" "=INDIRECT(\"A1\")")
(formula "Room" "B5" "=SUM(A:A)")
"#,
        );
        let why = |out: &str| compile(&wb, &function(&wb, "Room", "A1:A3", out)).lambda;
        assert_eq!(why("B1"), Err(Uncompiled::Range("A1:A3".to_string())));
        assert_eq!(why("B2"), Err(Uncompiled::Range("B1:B1".to_string())));
        assert_eq!(
            why("B3"),
            Ok("=LAMBDA(_A1,_A2,_A3,SUM(Room!$D$1:$D$4))".to_string())
        );
        assert_eq!(why("B4"), Err(Uncompiled::Construct("INDIRECT")));
        assert_eq!(why("B5"), Err(Uncompiled::Range("A:A".to_string())));
        let wide = Function {
            inputs: parse_a1_range("A1:A254").unwrap(),
            ..function(&wb, "Room", "A1", "B3")
        };
        assert_eq!(compile(&wb, &wide).lambda, Err(Uncompiled::Inputs(254)));
        assert_eq!(Uncompiled::Range("A1:A3".into()).words(), "range \"A1:A3\"");
        assert_eq!(
            Uncompiled::Construct("INDIRECT").words(),
            "construct INDIRECT"
        );
    }

    #[test]
    fn a_gate_reading_itself_is_a_cycle() {
        // A draft's workbook holds what a machine's load refuses.
        let wb = grid(
            r#"(cell "Room" "A1" 1)
(formula "Room" "B1" "=C1+A1")
(formula "Room" "C1" "=B1")
"#,
        );
        assert_eq!(
            compile(&wb, &function(&wb, "Room", "A1", "B1")).lambda,
            Err(Uncompiled::Cycle("B1".to_string()))
        );
    }

    #[test]
    fn more_bindings_than_a_let_takes_nest_a_let_inside_it() {
        let mut rows =
            String::from("(cell \"Room\" \"A1\" 1)\n(formula \"Room\" \"B1\" \"=A1+1\")\n");
        for k in 2..=130 {
            rows.push_str(&format!("(formula \"Room\" \"B{k}\" \"=B{}+1\")\n", k - 1));
        }
        let wb = grid(&rows);
        let text = compile(&wb, &function(&wb, "Room", "A1", "B130"))
            .lambda
            .unwrap();
        assert_eq!(text.matches("LET(").count(), 2, "{text}");
        assert!(
            text.starts_with("=LAMBDA(_A1,LET(_B1,_A1+1,_B2,_B1+1,"),
            "{text}"
        );
        assert!(text.ends_with(",_B129+1)))"), "{text}");
    }
}
