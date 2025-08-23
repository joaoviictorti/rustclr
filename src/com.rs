use core::ffi::c_void;

use dinvk::{GetProcAddress, LoadLibraryA};
use windows_core::{GUID, Interface};
use windows_sys::core::HRESULT;

use crate::Result;
use crate::error::ClrError;

/// CLSID (Class ID) constants for various CLR components.
///
/// These constants are used to identify specific COM classes within the Common Language Runtime (CLR).
pub const CLSID_CLRMETAHOST: GUID = GUID::from_u128(0x9280188d_0e8e_4867_b30c_7fa83884e8de);
pub const CLSID_COR_RUNTIME_HOST: GUID = GUID::from_u128(0xCB2F6723_AB3A_11d2_9C40_00C04FA30A3E);
pub const CLSID_ICLR_RUNTIME_HOST: GUID = GUID::from_u128(0x90F1A06E_7712_4762_86B5_7A5E_BA6B_DB02);

/// Static cache for the `CLRCreateInstance` function.
static CLR_CREATE_INSTANCE: spin::Once<Option<CLRCreateInstanceType>> = spin::Once::new();

/// Function type for creating instances of the CLR (Common Language Runtime).
///
/// # Arguments
///
/// * `clsid` - The GUID of the class to instantiate.
/// * `riid` - The GUID of the interface to be obtained from the instance.
/// * `ppinterface` - A pointer to store the resulting interface.
///
/// # Returns
///
/// * An `HRESULT` indicating success (`S_OK`) or the failure reason.
type CLRCreateInstanceType = fn(
    clsid: *const GUID,
    riid: *const GUID,
    ppinterface: *mut *mut c_void,
) -> HRESULT;

/// Function type for retrieving the current thread's CLR identity.
///
/// # Arguments
///
/// * `riid` - A pointer to the interface ID (`GUID`) being queried.
/// * `ppv` - A pointer that receives the interface pointer if successful.
///
/// # Returns
///
/// * An `HRESULT` indicating success (`S_OK`) or the failure reason.
pub(crate) type CLRIdentityManagerType = fn(riid: *const GUID, ppv: *mut *mut c_void) -> HRESULT;

/// Helper function to create a CLR instance based on the provided CLSID.
///
/// # Arguments
///
/// * `clsid` - A pointer to the GUID of the CLR class to instantiate.
///
/// # Returns
///
/// * `Ok(T)` - if the instance is created successfully, with `T` representing the interface requested.
/// * `Err(ClrError)` - if the function fails to load `CLRCreateInstance` or if the instance creation fails.
pub fn CLRCreateInstance<T>(clsid: *const GUID) -> Result<T>
where
    T: Interface,
{
    // Load the 'mscoree.dll' library and get the address of the 'CLRCreateInstance' function.
    let CLRCreateInstance = CLR_CREATE_INSTANCE.call_once(|| {
        // Load 'mscoree.dll' and get the address of 'CLRCreateInstance'
        let module = LoadLibraryA(obfstr::obfstr!("mscoree.dll"));
        if !module.is_null() {
            // Get the address of 'CLRCreateInstance'
            let addr = GetProcAddress(module, 2672818687u32, Some(dinvk::hash::murmur3));

            // Transmute the address to the function type
            return Some(unsafe {
                core::mem::transmute::<*mut c_void, CLRCreateInstanceType>(addr)
            });
        }

        None
    });

    if let Some(CLRCreateInstance) = CLRCreateInstance {
        let mut result = core::ptr::null_mut();

        // Call 'CLRCreateInstance' to create the CLR instance.
        let hr = CLRCreateInstance(clsid, &T::IID, &mut result);
        if hr == 0 {
            // Transmute the raw pointer to the expected interface type
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
