//! The sheet model's file formats (PORT.7): the OOXML writer and the zip,
//! XML and inflate it stands on, the merge of a build into a host's
//! workbook, and the file format's names for Excel's newer functions. The
//! model itself, sheets of cells with their A1 coordinates and a style
//! table, is the language's since PORT.12 (`vla_lang::sheet`) and is
//! re-exported here whole, so that `crate::sheet::Workbook` and every
//! other name the writer, the readers and the tests use are what they were.

pub mod inflate;
pub mod merge;
pub mod ooxml;
pub mod xlfn;
pub mod xml;
pub mod zip;

pub use vla_lang::sheet::*;
