//! `frazaro reflect --functions` (KERNEL.6's instrument, built with
//! KERNEL.7, 2026-10-08): a histogram of the functions a workbook's formulas
//! call, counts alone.
//!
//! The roadmap's `KERNEL.6` asks that the declared subset be chosen by
//! measurement over real workbooks, the Enron corpus and the house's own
//! models, and never by taste; this sink is the instrument. One `function`
//! row per name, the number of formula cells calling it (a cell calling
//! `SUM` twice counts once), sorted by count then by name; nothing of the
//! workbook's contents leaves the machine, AXM.1's discipline, so it can be
//! run over anyone's models. A script sums the rows over a folder.
//!
//! ```text
//! (function "SUM" 1365)
//! (function "IF" 912)
//! ```

use std::collections::BTreeMap;

use vla_lang::calc::calls;
use vla_lang::rows::{quoted, Row, Sink};

use super::open;
use crate::messages::Refusal;

/// The functions called, by name, each the number of formula cells.
#[derive(Debug, Default)]
pub struct Histogram {
    pub counts: BTreeMap<String, u64>,
    pub formulas: u64,
}

impl Sink for Histogram {
    fn row(&mut self, row: &Row<'_>) {
        if let Row::Formula { text, .. } = row {
            self.formulas += 1;
            let mut names = calls(text);
            names.sort_unstable();
            names.dedup();
            for name in names {
                *self.counts.entry(name).or_insert(0) += 1;
            }
        }
    }
}

impl Histogram {
    /// The rows, most called first, then by name.
    pub fn text(&self) -> String {
        let mut rows: Vec<(&String, &u64)> = self.counts.iter().collect();
        rows.sort_by(|a, b| b.1.cmp(a.1).then_with(|| a.0.cmp(b.0)));
        let mut out = String::new();
        for (name, count) in rows {
            out.push_str(&format!("(function {} {count})\n", quoted(name)));
        }
        out
    }
}

/// The histogram of a workbook's file, every sheet walked.
pub fn histogram(bytes: &[u8], label: &str) -> Result<Histogram, Refusal> {
    let package = open(bytes, label)?;
    let mut h = Histogram::default();
    for i in 0..package.sheets().len() {
        package.walk_sheet(i, &mut h)?;
    }
    Ok(h)
}

/// The histogram as text.
pub fn histogram_text(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    Ok(histogram(bytes, label)?.text())
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../scripts/reflect/fixture.xlsx");

    #[test]
    fn the_fixture_s_functions_are_counted_once_per_cell() {
        let h = histogram(FIXTURE, "fixture.xlsx").unwrap();
        assert_eq!(h.counts.get("SUM"), Some(&3), "{:?}", h.counts);
        assert_eq!(h.counts.get("INDIRECT"), Some(&1));
        assert_eq!(h.counts.get("OFFSET"), Some(&1));
        assert_eq!(h.counts.get("SEQUENCE"), Some(&1));
        let text = h.text();
        assert!(text.starts_with("(function \"SUM\" 3)\n"), "{text}");
        assert!(h.formulas >= 25);
    }
}
