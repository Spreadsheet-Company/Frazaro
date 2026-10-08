//! The relations as printed (PORT.8, slice 8a): the spelling is the
//! language's since PORT.12 (`vla_lang::rows`), re-exported here so that
//! `reflect::print::line`, `datum`, `target`, `quoted` and the `Printer`
//! sink are what they were for the reader, the door and the view.

pub use vla_lang::rows::{datum, line, quoted, target, Printer};
