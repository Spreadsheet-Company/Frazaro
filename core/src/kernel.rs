//! The kernel's boundary, and the five seams through which everything else
//! enters (`KERNEL.1`, 2026-10-06; the design is `web/CALLOSUM.md` §14).
//!
//! The kernel holds mechanism only: forms and their expansion, the emitters,
//! the sheet model, recalculation over a declared subset once it lands, the
//! relation set, and the ABI. It never holds English, message text, a
//! default, chrome, a format beyond a trait, or a door. Each of those is
//! data the kernel reads (the four tables under `core/data/`, a phrasebook, a
//! library) or an implementation of one of the seams below. The rule that
//! follows: the kernel grows a seam, never a feature.
//!
//! `tools/check_kernel_boundary.ps1` pins the rule over `core/src/`. This
//! module writes the seams down as data, so that a test can hold their
//! counts and `frazaro describe` (`KERNEL.18`) can print them, and names the
//! two traits whose first implementations later items land: [`Engine`]
//! (`PORT.9`) and [`Projection`] (`KERNEL.4`).

use crate::messages::Refusal;
use crate::sheet::{A1Range, Workbook};

/// One of the five entrances into the kernel.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Seam {
    /// The seam's name, as `CONTRIBUTING.md` heads its row.
    pub name: &'static str,
    /// What a contribution through it is: a data file, or an implementation
    /// of a trait in this crate.
    pub entrance: &'static str,
    /// The oracle a change through it must pass before it lands.
    pub oracle: &'static str,
    /// The implementations the kernel ships today, by name; empty where the
    /// first is deferred to the roadmap item in `deferred_to`.
    pub implementations: &'static [&'static str],
    /// The roadmap item that lands the first implementation, or empty.
    pub deferred_to: &'static str,
}

/// The five seams, in the order `CONTRIBUTING.md` lists them.
pub const SEAMS: [Seam; 5] = [
    Seam {
        name: "sentences",
        entrance: "a phrasebook file (.vla) with its test: proofs, loaded through api::vocab_gate",
        oracle: "frazaro prove passes every proof; the translate and refusal goldens hold",
        implementations: &["the phrasebooks a door loads, english.vla first"],
        deferred_to: "",
    },
    Seam {
        name: "paragraphs",
        entrance: "a library of sentences with named slots, imported by a Use sentence",
        oracle: "the build golden, and the stamp hashing every file used",
        implementations: &[],
        deferred_to: "G-USE",
    },
    Seam {
        name: "engines",
        entrance: "an implementation of kernel::Engine under a head-table symbol",
        oracle: "a proof file of its kind, with clingo beside it where it applies",
        implementations: &[],
        deferred_to: "PORT.9",
    },
    Seam {
        name: "formats and hosts",
        entrance: "an implementation of reflect::Source to read, a writer beside it, and a kernel::HostProfile per door",
        oracle: "the reflect, diff and audit goldens for a reader; the build golden for a writer",
        implementations: &["xlsx (OOXML), read and written", "ods (OpenDocument), read"],
        deferred_to: "",
    },
    Seam {
        name: "projections",
        entrance: "an implementation of kernel::Projection",
        oracle: "a view golden per projection",
        implementations: &[],
        deferred_to: "KERNEL.4",
    },
];

/// The seams, as data.
pub fn seams() -> &'static [Seam] {
    &SEAMS
}

/// A file format the kernel reads, writes, or both.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Format {
    pub name: &'static str,
    pub extension: &'static str,
    pub reads: bool,
    pub writes: bool,
}

/// The formats seam's registry. `reflect::open` picks the reader by the
/// bytes it is handed; the one writer is `sheet::ooxml::workbook_bytes`.
pub const FORMATS: [Format; 2] = [
    Format {
        name: "OOXML",
        extension: "xlsx",
        reads: true,
        writes: true,
    },
    Format {
        name: "OpenDocument",
        extension: "ods",
        reads: true,
        writes: false,
    },
];

/// An engine: tables in, a table out, under a symbol of the head table,
/// held by proof files of one kind. The relation type is the implementor's,
/// since the port of `VLA_Relation` (`PORT.9`) defines it; the first
/// implementations are SQL, DATALOG, PROLOG and OPTIMIZE, and a fifth proves
/// the seam takes additions.
pub trait Engine {
    /// The relation an engine takes and gives.
    type Table;
    /// The head-table symbol the engine answers to, such as `datalog`.
    fn head(&self) -> &str;
    /// The proof-file kind that holds it, such as `test-datalog`.
    fn proof_kind(&self) -> &str;
    /// The tables in, the table out, or a refusal from the catalogue.
    fn run(&self, tables: &[Self::Table]) -> Result<Self::Table, Refusal>;
}

/// A window onto the model: one sheet by name, one rectangle of it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Window {
    pub sheet: String,
    pub range: A1Range,
}

/// A projection: a pure function from the model and a window to lines, the
/// shape every door prints (`SD-23`). The grid is the first (`KERNEL.4`),
/// the sentence pane the second, the dependency cone drawn as a diagram the
/// third (`KERNEL.12`). A projection never edits the model.
pub trait Projection {
    /// The projection's name, as `frazaro view` selects it.
    fn name(&self) -> &str;
    /// The window's lines, in a fixed order, to the sink.
    fn project(&self, model: &Workbook, window: &Window, out: &mut dyn FnMut(&str));
}

/// A door's host profile: the statement heads it carries out with nothing
/// running, and the catalogue id that refuses the rest by name.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct HostProfile {
    pub door: &'static str,
    /// The core forms, by head-table symbol, the door holds.
    pub forms: &'static [&'static str],
    /// The scaffold's own calls the door steps over without effect.
    pub scaffold: &'static [&'static str],
    /// The catalogue id of the refusal for a sentence outside the profile.
    pub refuses_with: &'static str,
}

/// The file door: `frazaro build`, and the page's Workbook strip. What the
/// walker in `build.rs` holds, read off its arms: a value or a formula into
/// a range, `Work on sheet` and `Go to sheet`, a condition that folds, and
/// a call of one of the program's own steps with no values, inlined. The
/// rest is refused with the sentence quoted.
pub const FILE_DOOR: HostProfile = HostProfile {
    door: "file",
    forms: &[
        "at-line",
        "begin",
        "dim",
        "on-error",
        "label",
        "exit-sub",
        "exit-function",
        "set!",
        "if",
        ".",
    ],
    scaffold: &["vla-report-error", "vlaensuresheet"],
    refuses_with: "build-not-representable",
};

/// The host profiles the kernel knows. The Excel add-in needs none, since
/// it holds every form; the page's viewport and the task pane add theirs
/// with `KERNEL.19`.
pub const HOST_PROFILES: [HostProfile; 1] = [FILE_DOOR];

#[cfg(test)]
mod tests {
    use super::*;
    use crate::headtable;
    use crate::reflect;

    const FIXTURE_XLSX: &[u8] = include_bytes!("../../scripts/reflect/fixture.xlsx");
    const FIXTURE_ODS: &[u8] = include_bytes!("../../scripts/reflect/opendocument.ods");

    #[test]
    fn five_seams_each_with_an_entrance_and_an_oracle() {
        assert_eq!(seams().len(), 5);
        let names: Vec<&str> = seams().iter().map(|s| s.name).collect();
        assert_eq!(
            names,
            [
                "sentences",
                "paragraphs",
                "engines",
                "formats and hosts",
                "projections"
            ]
        );
        for s in seams() {
            assert!(!s.entrance.is_empty(), "{} names no entrance", s.name);
            assert!(!s.oracle.is_empty(), "{} names no oracle", s.name);
            assert!(
                !s.implementations.is_empty() || !s.deferred_to.is_empty(),
                "{} has neither an implementation nor the item that lands one",
                s.name
            );
        }
    }

    #[test]
    fn the_engines_and_projections_seams_are_named_and_empty_today() {
        let engines = seams().iter().find(|s| s.name == "engines").unwrap();
        assert!(engines.implementations.is_empty());
        assert_eq!(engines.deferred_to, "PORT.9");
        let projections = seams().iter().find(|s| s.name == "projections").unwrap();
        assert!(projections.implementations.is_empty());
        assert_eq!(projections.deferred_to, "KERNEL.4");
        let paragraphs = seams().iter().find(|s| s.name == "paragraphs").unwrap();
        assert_eq!(paragraphs.deferred_to, "G-USE");
    }

    #[test]
    fn the_formats_registry_matches_the_readers_and_the_writer() {
        assert_eq!(FORMATS.len(), 2);
        assert!(reflect::open(FIXTURE_XLSX, "fixture.xlsx").is_ok());
        assert!(reflect::open(FIXTURE_ODS, "opendocument.ods").is_ok());
        let readers: Vec<&str> = FORMATS
            .iter()
            .filter(|f| f.reads)
            .map(|f| f.extension)
            .collect();
        assert_eq!(readers, ["xlsx", "ods"]);
        let writers: Vec<&str> = FORMATS
            .iter()
            .filter(|f| f.writes)
            .map(|f| f.extension)
            .collect();
        assert_eq!(writers, ["xlsx"], "the one writer is the OOXML writer");
        let formats = seams()
            .iter()
            .find(|s| s.name == "formats and hosts")
            .unwrap();
        assert_eq!(formats.implementations.len(), FORMATS.len());
    }

    #[test]
    fn the_file_door_holds_only_core_forms_and_refuses_from_the_catalogue() {
        assert_eq!(HOST_PROFILES.len(), 1);
        assert_eq!(FILE_DOOR.door, "file");
        for sym in FILE_DOOR.forms {
            let canonical = headtable::resolve_head_alias(sym);
            assert!(
                headtable::row(&canonical).is_some(),
                "{sym} is not a form of the head table"
            );
        }
        for call in FILE_DOOR.scaffold {
            assert!(
                headtable::row(call).is_none(),
                "{call} is a scaffold call, not a core form"
            );
        }
        assert_eq!(FILE_DOOR.refuses_with, "build-not-representable");
    }

    #[test]
    fn a_window_is_one_sheet_and_one_rectangle() {
        let w = Window {
            sheet: "Model".to_string(),
            range: A1Range {
                top: 1,
                left: 1,
                bottom: 20,
                right: 6,
            },
        };
        assert_eq!(w.range.cells(), 120);
        assert_eq!(w.sheet, "Model");
    }
}
