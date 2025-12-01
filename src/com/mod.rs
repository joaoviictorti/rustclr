//! Raw COM interface bindings for interacting with the .NET CLR runtime.

use core::ffi::c_void;

use windows_core::{GUID, Interface};
use windows_sys::core::HRESULT;
use dinvk::module::get_proc_address; 
use dinvk::winapis::LoadLibraryA;

use crate::error::{ClrError, Result};

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
mod method_info;

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
pub use method_info::*;

/// Caches the address of the `CLRCreateInstance` function on first use.
static CLR_CREATE_INSTANCE: spin::Once<Option<CLRCreateInstanceType>> = spin::Once::new();

/// CLSID for the CLR MetaHost (`ICLRMetaHost`).
pub const CLSID_CLRMETAHOST: GUID = GUID::from_u128(0x9280188d_0e8e_4867_b30c_7fa83884e8de);

/// CLSID for the COR Runtime Host (`ICorRuntimeHost`).
pub const CLSID_COR_RUNTIME_HOST: GUID = GUID::from_u128(0xCB2F6723_AB3A_11D2_9C40_00C04FA30A3E);

/// CLSID for the ICLR Runtime Host (`ICLRRuntimeHost`).
pub const CLSID_ICLR_RUNTIME_HOST: GUID = GUID::from_u128(0x90F1_A06E_7712_4762_86B5_7A5E_BA6B_DB02);

pub(crate) type CLRIdentityManagerType = fn(
    riid: *const GUID, 
    ppv: *mut *mut c_void
) -> HRESULT;

type CLRCreateInstanceType = fn(
    clsid: *const GUID,
    riid: *const GUID,
    ppinterface: *mut *mut c_void,
) -> HRESULT;

/// Dynamically loads and invokes the `CLRCreateInstance`.
pub fn CLRCreateInstance<T>(clsid: *const GUID) -> Result<T>
where
    T: Interface,
{
    // Resolve the CLRCreateInstance function pointer.
    let CLRCreateInstance = CLR_CREATE_INSTANCE.call_once(|| {
        let module = LoadLibraryA(obfstr::obfstr!("mscoree.dll"));
        if !module.is_null() {
            let addr = get_proc_address(module, 2672818687u32, Some(dinvk::hash::murmur3));
            return Some(unsafe {
                core::mem::transmute::<*mut c_void, CLRCreateInstanceType>(addr)
            });
        }
        None
    });

    // Invoke CLRCreateInstance to create the requested COM interface
    if let Some(CLRCreateInstance) = CLRCreateInstance {
        let mut result = core::ptr::null_mut();
        let hr = CLRCreateInstance(clsid, &T::IID, &mut result);
        if hr == 0 {
            Ok(unsafe { core::mem::transmute_copy(&result) })
        } else {
            Err(ClrError::ApiError("CLRCreateInstance", hr))
        }
    } else {
        Err(ClrError::Msg(
            "CLRCreateInstance function not found",
        ))
    }
}
