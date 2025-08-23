//! # rustclr 🦀
//!
//! A Rust library for hosting the **Common Language Runtime (CLR)** and executing .NET assemblies directly.
//!
//! ## Features
//! - Run .NET binaries fully in memory, with control over runtime configuration.
//! - Redirect and capture output from .NET programs.
//! - Patch `System.Environment.Exit` to prevent termination of the host process.
//! - Execute PowerShell commands through the .NET automation namespace.
//! - Fine-grained control via `RustClrEnv` and `ClrOutput`.
//! - `#[no_std]` support (with `alloc`).
//!
//! ## Examples
//!
//! ### Running a .NET Assembly
//! ```no_run
//! use std::fs;
//! use rustclr::{RustClr, RuntimeVersion};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let buffer = fs::read("examples/sample.exe")?;
//!
//!     let output = RustClr::new(&buffer)?
//!         .runtime_version(RuntimeVersion::V4)
//!         .output()
//!         .domain("CustomDomain")
//!         .exit()
//!         .args(vec!["arg1", "arg2"])
//!         .run()?;
//!
//!     println!("Captured output: {}", output);
//!     Ok(())
//! }
//! ```
//!
//! ### Running PowerShell Commands
//! ```no_run
//! use rustclr::PowerShell;
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let pwsh = PowerShell::new()?;
//!     println!("{}", pwsh.execute("Get-Process | Select-Object -First 1")?);
//!     Ok(())
//! }
//! ```
//!
//! ### Control with `RustClrEnv`
//! ```no_run
//! use rustclr::{RustClrEnv, RuntimeVersion};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let clr_env = RustClrEnv::new(Some(RuntimeVersion::V4))?;
//!     println!("CLR environment initialized with version {:?}", clr_env.runtime_version);
//!     Ok(())
//! }
//! ```
//!
//! # More Information
//!
//! For additional examples and CLI usage, visit the [repository].
//!
//! [repository]: https://github.com/joaoviictorti/rustclr

#![no_std]
#![allow(non_snake_case, non_camel_case_types)]
#![allow(
    clippy::not_unsafe_ptr_arg_deref,
    clippy::missing_transmute_annotations,
    clippy::mixed_case_hex_literals,
    clippy::unusual_byte_groupings,
    clippy::useless_transmute,
)]

extern crate alloc;

pub mod data;
pub mod error;

mod com;
mod host_control;
mod pwsh;
mod utils;
mod clr;

pub use clr::*;
pub use pwsh::*;
pub use utils::*;

/// Type alias for `Result` with `ClrError` as the error type.
pub(crate) type Result<T> = core::result::Result<T, crate::error::ClrError>;
