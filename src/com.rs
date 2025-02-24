use crate::error::ClrError;
use windows_core::{Interface, GUID};    
use std::{ffi::c_void , sync::OnceLock};
use dinvk::{LoadLibraryA, GetProcAddress};
use windows_sys::core::HRESULT;

/// CLSID (Class ID) constants for various CLR components.
/// 
/// These constants are used to identify specific COM classes within the Common Language Runtime (CLR).
pub const CLSID_CLRMETAHOST: GUID = GUID::from_u128(0x9280188d_0e8e_4867_b30c_7fa83884e8de);
pub const CLSID_CLRRUNTIMEHOST: GUID = GUID::from_u128(0x90f1a06e_7712_4762_86b5_7a5eba6bdb02);
pub const CLSID_COR_RUNTIME_HOST: GUID = GUID::from_u128(0xCB2F6723_AB3A_11d2_9C40_00C04FA30A3E);

/// Static cache for the `CLRCreateInstance` function.
/// 
/// The `OnceLock` ensures that the function is loaded from `mscoree.dll` only once
/// and is reused for subsequent calls to create CLR instances.
static CLR_CREATE_INSTANCE: OnceLock<Option<CLRCreateInstanceFn>> = OnceLock::new();

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
/// * Returns an `HRESULT` indicating success or failure in creating the instance.
type CLRCreateInstanceFn = fn(
    clsid: *const windows_core::GUID,
    riid: *const windows_core::GUID,
    ppinterface: *mut *mut c_void,
) -> HRESULT;

/// Attempts to load the `CLRCreateInstance` function from `mscoree.dll`.
/// 
/// # Returns
/// 
/// * `Some(CLRCreateInstanceFn)` - if the function is found and loaded successfully.
/// * `None` - if `mscoree.dll` cannot be loaded or if `CLRCreateInstance` is not found.
fn init_clr_create_instance() -> Option<CLRCreateInstanceFn> {
    unsafe {
        // Load 'mscoree.dll' and get the address of 'CLRCreateInstance'
        let lib = LoadLibraryA("mscoree.dll");
        if !lib.is_null() {
            // Get the address of 'CLRCreateInstance'
            let addr = GetProcAddress(lib, "CLRCreateInstance", None);
            
            // Transmute the address to the function type
            return Some(core::mem::transmute::<*mut c_void, CLRCreateInstanceFn>(addr));
        }

        None
    }
}

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
pub fn CLRCreateInstance<T>(clsid: *const GUID) -> Result<T, ClrError>
where
    T: Interface
{
    // Load the 'mscoree.dll' library and get the address of the 'CLRCreateInstance' function.
    let CLRCreateInstance = CLR_CREATE_INSTANCE.get_or_init(init_clr_create_instance);

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
        Err(ClrError::ErrorClr("CLRCreateInstance function not found"))
    }
}