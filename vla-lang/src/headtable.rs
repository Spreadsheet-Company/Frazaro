//! The head table: `VLA_HeadTable.bas`'s catalogue of the core forms, read
//! as data (PORT.5, slice 5c).
//!
//! One row per head symbol: its aliases, arity, interpreter routine, VBA
//! routine, its standing in the formula dialect, and whether it is
//! export-only. The VBA is the source; `tools/export_headtable.ps1` writes
//! `core/data/headtable.vla` from it and `tools/check_data_exports.ps1` fails
//! when the two drift. As in the VBA, the one live column is the aliases:
//! [`alias_map`] feeds [`resolve_head_alias`], the lookup `EmitTop`,
//! `EmitStmt`, `EmitExpr` and `EmitFormula` run before their dispatch, so an
//! aliased head (`fijar!` for `set!`) reaches the arm its canonical spelling
//! would. The other columns are the catalogue they are in the VBA, not the
//! dispatch.

use std::collections::HashMap;
use std::sync::OnceLock;

use crate::form::Form;
use crate::intrinsics::fold;
use crate::reader::read_forms;

/// The formula column's three states, as the VBA's honesty note names them.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormulaSupport {
    Yes,
    Refused,
    Undeclared,
}

/// One row of the table.
#[derive(Clone, Debug, PartialEq)]
pub struct HeadRow {
    pub symbol: String,
    pub aliases: Vec<String>,
    pub arity: String,
    pub interp_routine: String,
    pub vba_routine: String,
    pub formula: FormulaSupport,
    pub export_only: bool,
}

const TABLE_TEXT: &str = include_str!("../data/headtable.vla");

fn parse_rows() -> Vec<HeadRow> {
    let text = TABLE_TEXT.replace("\r\n", "\n");
    let forms = match read_forms(&text) {
        Ok(forms) => forms,
        Err(_) => return Vec::new(),
    };
    let mut rows = Vec::with_capacity(forms.len());
    for form in &forms {
        let Form::List(l) = form else { continue };
        if !l.head_is("head") || l.items.len() != 8 {
            continue;
        }
        let (
            Form::Sym(symbol),
            Form::List(aliases),
            Form::Str(arity),
            Form::Str(interp),
            Form::Str(vba),
            Form::Sym(formula),
            Form::Sym(export_only),
        ) = (
            &l.items[1],
            &l.items[2],
            &l.items[3],
            &l.items[4],
            &l.items[5],
            &l.items[6],
            &l.items[7],
        )
        else {
            continue;
        };
        let formula = match formula.as_str() {
            "yes" => FormulaSupport::Yes,
            "refused" => FormulaSupport::Refused,
            "undeclared" => FormulaSupport::Undeclared,
            _ => continue,
        };
        let aliases = aliases
            .items
            .iter()
            .filter_map(|a| match a {
                Form::Sym(s) => Some(s.clone()),
                _ => None,
            })
            .collect();
        rows.push(HeadRow {
            symbol: symbol.clone(),
            aliases,
            arity: arity.clone(),
            interp_routine: interp.clone(),
            vba_routine: vba.clone(),
            formula,
            export_only: export_only == "true",
        });
    }
    rows
}

/// VBA `VlaHeadTableRows`: every row, in the table's order.
pub fn rows() -> &'static [HeadRow] {
    static ROWS: OnceLock<Vec<HeadRow>> = OnceLock::new();
    ROWS.get_or_init(parse_rows)
}

/// VBA `VlaHeadTableCount`.
pub fn count() -> usize {
    rows().len()
}

/// VBA `VlaHeadTableRow`: a row by its symbol or by any of its aliases,
/// compared folded.
pub fn row(symbol: &str) -> Option<&'static HeadRow> {
    let key = fold(symbol);
    rows().iter().find(|r| fold(&r.symbol) == key).or_else(|| {
        rows()
            .iter()
            .find(|r| r.aliases.iter().any(|a| fold(a) == key))
    })
}

/// VBA `VlaHeadTableAliasMap`: each alias, folded, to its canonical symbol
/// as written. Built once; the VBA rebuilds it per compile, to the same
/// map.
pub fn alias_map() -> &'static HashMap<String, String> {
    static MAP: OnceLock<HashMap<String, String>> = OnceLock::new();
    MAP.get_or_init(|| {
        let mut m = HashMap::new();
        for r in rows() {
            for a in &r.aliases {
                m.insert(fold(a), r.symbol.clone());
            }
        }
        m
    })
}

/// VBA `ResolveHeadAlias`: a folded head to its canonical spelling, or
/// itself when it is not an alias.
pub fn resolve_head_alias(folded_head: &str) -> String {
    match alias_map().get(folded_head) {
        Some(canon) => canon.clone(),
        None => folded_head.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_table_is_read_whole() {
        // 65 rows at IN5.0; the same floor is held on the file by
        // tools/check_data_exports.ps1.
        assert_eq!(count(), 65);
        assert_eq!(rows()[0].symbol, "sub");
        assert_eq!(rows()[0].vba_routine, "EmitProc");
        assert_eq!(rows()[64].symbol, "not");
        let yes = rows()
            .iter()
            .filter(|r| r.formula == FormulaSupport::Yes)
            .count();
        let refused = rows()
            .iter()
            .filter(|r| r.formula == FormulaSupport::Refused)
            .count();
        let undeclared = rows()
            .iter()
            .filter(|r| r.formula == FormulaSupport::Undeclared)
            .count();
        assert_eq!((yes, refused, undeclared), (18, 9, 38));
    }

    #[test]
    fn export_only_is_exactly_raw_deflambda_lambda() {
        // TestHeadTableExportOnly's pin (VLA_Tests.bas), held here too.
        let set: Vec<&str> = rows()
            .iter()
            .filter(|r| r.export_only)
            .map(|r| r.symbol.as_str())
            .collect();
        assert_eq!(set, vec!["raw", "deflambda", "lambda"]);
    }

    #[test]
    fn the_three_aliases_resolve_and_nothing_else_does() {
        assert_eq!(resolve_head_alias("fijar!"), "set!");
        assert_eq!(resolve_head_alias("si"), "if");
        assert_eq!(resolve_head_alias("depurar"), "debug-print");
        assert_eq!(resolve_head_alias("set!"), "set!");
        assert_eq!(resolve_head_alias("unknown"), "unknown");
        assert_eq!(alias_map().len(), 3);
        assert_eq!(row("FIJAR!").unwrap().symbol, "set!");
        assert_eq!(row("Mod").unwrap().arity, "2 (a, b)");
        assert!(row("nope").is_none());
    }
}
