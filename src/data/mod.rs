//! # CLR (Common Language Runtime) Interface Bindings
//!
//! This library provides bindings for interacting with the .NET CLR, including the ability to
//! enumerate runtimes, manage AppDomains, manipulate assemblies and access type information.

mod appdomain;
mod assembly;
mod assembly_identity;
mod assembly_manager;
mod assembly_store;
mod iclrmetahost;
mod iclrruntimehost;
mod iclrruntimeinfo;
mod icorruntimehost;
mod ienumunknown;
mod ihostcontrol;
mod ipropertyinfo;
mod itype;
mod methodinfo;

pub use appdomain::*;
pub use assembly::*;
pub use assembly_identity::*;
pub use assembly_manager::*;
pub use assembly_store::*;
pub use iclrmetahost::*;
pub use iclrruntimehost::*;
pub use iclrruntimeinfo::*;
pub use icorruntimehost::*;
pub use ienumunknown::*;
pub use ihostcontrol::*;
pub use ipropertyinfo::*;
pub use itype::*;
pub use methodinfo::*;
