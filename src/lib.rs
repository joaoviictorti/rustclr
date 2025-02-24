#![doc = include_str!("../README.md")]
#![allow(non_snake_case, non_camel_case_types)]
#![allow(clippy::not_unsafe_ptr_arg_deref)]

/// Defines data structures and descriptions for manipulating and interacting with the CLR.
pub mod data;

/// Contains definitions for COM interoperability, making it easier to call methods and manipulate COM interfaces.
pub mod com;

/// Manages specific error types used when interacting with the CLR and COM APIs.
pub mod error;

/// Main CLR module, providing functions and structures for working with the Common Language Runtime.
mod clr;

/// Auxiliary functions for common manipulations and conversions needed when interacting with the CLR and COM.
mod utils;

pub use clr::*;
pub use utils::*;