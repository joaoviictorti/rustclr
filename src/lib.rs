#![no_std]
#![doc = include_str!("../README.md")]
#![allow(non_snake_case, non_camel_case_types)]
#![allow(
    clippy::not_unsafe_ptr_arg_deref,
    clippy::missing_transmute_annotations,
    clippy::mixed_case_hex_literals,
    clippy::unusual_byte_groupings,
    clippy::useless_transmute,
)]

extern crate alloc;

mod com;
mod host_control;

/// Defines data structures and descriptions for manipulating and interacting with the CLR.
pub mod data;

/// Manages specific error types used when interacting with the CLR and COM APIs.
pub mod error;

/// Implementing the core CLR loading and interaction logic.
mod clr;
pub use clr::*;

/// Responsible for executing powershell code using CLR
mod pwsh;
pub use pwsh::*;

/// Utilities
mod utils;
pub use utils::*;

/// Type alias for `Result` with `ClrError` as the error type.
pub(crate) type Result<T> = core::result::Result<T, crate::error::ClrError>;
