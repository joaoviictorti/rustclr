use core::ffi::c_void;
use dinvk::{GetProcAddress, LoadLibraryA};
use windows_core::{GUID, Interface};
use windows_sys::core::HRESULT;
use super::{Result, error::ClrError};

/// Static cache for the `CLRCreateInstance` function.
static CLR_CREATE_INSTANCE: spin::Once<Option<CLRCreateInstanceType>> = spin::Once::new();

/// CLR MetaHost (manages CLR versions)
pub const CLSID_CLRMETAHOST: GUID = GUID::from_u128(0x9280188d_0e8e_4867_b30c_7fa83884e8de);

/// COR Runtime Host (loads/manages CLR)
pub const CLSID_COR_RUNTIME_HOST: GUID = GUID::from_u128(0xCB2F6723_AB3A_11d2_9C40_00C04FA30A3E);

/// ICLR Runtime Host (runtime hosting interface)
pub const CLSID_ICLR_RUNTIME_HOST: GUID = GUID::from_u128(0x90F1A06E_7712_4762_86B5_7A5E_BA6B_DB02);

/// Function type for creating instances of the CLR.
type CLRCreateInstanceType = fn(
    clsid: *const GUID,
    riid: *const GUID,
    ppinterface: *mut *mut c_void,
) -> HRESULT;

/// Function type for retrieving the current thread's CLR identity.
pub(crate) type CLRIdentityManagerType = fn(
    riid: *const GUID, 
    ppv: *mut *mut c_void
) -> HRESULT;

/// Helper function to create a CLR instance based on the provided CLSID.
pub fn CLRCreateInstance<T>(clsid: *const GUID) -> Result<T>
where
    T: Interface,
{
    // Load the 'mscoree.dll' library and get the address of the 'CLRCreateInstance' function.
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
