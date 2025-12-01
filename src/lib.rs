#![no_std]
#![doc = include_str!("../README.md")]
#![allow(non_snake_case, non_camel_case_types)]
#![allow(
    clippy::not_unsafe_ptr_arg_deref,
    clippy::missing_transmute_annotations,
    clippy::useless_transmute,
)]

extern crate alloc;

#[cfg(test)]
extern crate std;

pub mod com;
pub mod error;
pub mod variant;
pub mod string;

mod clr;
mod pwsh;

pub use clr::*;
pub use pwsh::PowerShell;
