use core::ffi::c_void;
use dinvk::{GetProcAddress, LoadLibraryA};
use windows_core::{GUID, Interface};
use windows_sys::core::HRESULT;
use crate::{Result, error::ClrError};

/// Caches the address of the `CLRCreateInstance` function on first use.
///
/// This avoids repeated calls to `LoadLibraryA` and `GetProcAddress` by memoizing
/// the resolved function pointer for future calls.
static CLR_CREATE_INSTANCE: spin::Once<Option<CLRCreateInstanceType>> = spin::Once::new();

/// CLSID for the CLR MetaHost (`ICLRMetaHost`).
pub const CLSID_CLRMETAHOST: GUID = GUID::from_u128(0x9280188d_0e8e_4867_b30c_7fa83884e8de);

/// CLSID for the COR Runtime Host (`ICorRuntimeHost`).
pub const CLSID_COR_RUNTIME_HOST: GUID = GUID::from_u128(0xCB2F6723_AB3A_11D2_9C40_00C04FA30A3E);

/// CLSID for the ICLR Runtime Host (`ICLRRuntimeHost`).
pub const CLSID_ICLR_RUNTIME_HOST: GUID = GUID::from_u128(0x90F1_A06E_7712_4762_86B5_7A5E_BA6B_DB02);

/// Signature of the `CLRCreateInstance` function exported by `mscoree.dll`.
type CLRCreateInstanceType = fn(
    clsid: *const GUID,
    riid: *const GUID,
    ppinterface: *mut *mut c_void,
) -> HRESULT;

/// Function pointer type for retrieving the CLR identity of the current thread.
pub type CLRIdentityManagerType = fn(
    riid: *const GUID, 
    ppv: *mut *mut c_void
) -> HRESULT;

/// Dynamically loads and invokes the `CLRCreateInstance` function from `mscoree.dll`.
pub fn CLRCreateInstance<T>(clsid: *const GUID) -> Result<T>
where
    T: Interface,
{
    // Resolve the CLRCreateInstance function pointer.
    let CLRCreateInstance = CLR_CREATE_INSTANCE.call_once(|| {
        let module = LoadLibraryA(obfstr::obfstr!("mscoree.dll"));
        if !module.is_null() {
            let addr = GetProcAddress(module, 2672818687u32, Some(dinvk::hash::murmur3));
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
        Err(ClrError::GenericError(
            "CLRCreateInstance function not found",
        ))
    }
}
