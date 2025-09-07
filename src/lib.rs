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

type Result<T> = core::result::Result<T, error::ClrError>;