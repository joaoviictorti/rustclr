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
mod assembly_identity;
mod iclrruntimehost;
mod ihostcontrol;
mod assembly_manager;
mod assembly_store;
mod ipropertyinfo;

pub use itype::*;
pub use assembly::*;
pub use appdomain::*;
pub use ienumunknown::*;
pub use iclrmetahost::*;
pub use iclrruntimeinfo::*;
pub use icorruntimehost::*;
pub use methodinfo::*;
pub use assembly_identity::*;
pub use iclrruntimehost::*;
pub use ihostcontrol::*;
pub use assembly_manager::*;
pub use assembly_store::*;
pub use ipropertyinfo::*;