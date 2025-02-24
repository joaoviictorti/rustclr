//! # CLR (Common Language Runtime) Interface Bindings
//!
//! This library provides bindings for interacting with the .NET CLR, including the ability to
//! enumerate runtimes, manage AppDomains, manipulate assemblies and access type information.

mod assembly;
mod appdomain;
mod iclrmetahost;
mod iclrruntimeinfo;
mod icorruntimehost;
mod ienumunknown;
mod methodinfo;
mod itype;

pub use itype::*;
pub use assembly::*;
pub use appdomain::*;
pub use ienumunknown::*;
pub use iclrmetahost::*;
pub use iclrruntimeinfo::*;
pub use icorruntimehost::*;
pub use methodinfo::*;