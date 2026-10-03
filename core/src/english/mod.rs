//! English in the core: VLA_SentenceEngine.bas and VLA_English.bas, ported
//! slice by slice (PORT.6; docs/HORIZON.md section 12, slice 2).
//!
//! The shape is the VBA's: each `Select Case` arm one `match` arm, each
//! rule table one `Vec`, every word list and name list the VBA holds in
//! code read here as data exported from it (`scripts/words.vla`,
//! `scripts/names.vla`), and every refusal the catalogue's, by id, in the
//! same situation. The treaty's oracles for this slice are
//! `instructions.txt` to `instructions_golden.vla` byte for byte, every
//! `(test-success ...)` and `(test-fail ...)` form of every source
//! phrasebook, and every refusal id on the path (conformance/README.md).
//!
//! Slices, each a module as it lands: `words` (6a, the tables), `tokenize`
//! (6b, `EnTokenize`), `rules`, `grammar` and `vocab` (6c, the rule store
//! and phrasebook loading), `matcher`, `expr` and `stmt` (6d, the DCG
//! matcher, the expression and condition grammars, the proof runners and
//! the part of `ParseStmt` a proof reaches), then the statement grammar
//! with `EnglishToVla`, the refusals, the translate API `VLA_Browser.bas`
//! exports, and the door.

pub mod expr;
pub mod grammar;
pub mod matcher;
pub mod program;
pub mod prolog;
pub mod refusals;
pub mod rules;
pub mod stmt;
pub mod tokenize;
pub mod vocab;
pub mod words;

pub use grammar::Grammar;
pub use program::{RefusalAtLine, Translation};
