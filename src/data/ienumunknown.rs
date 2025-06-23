use crate::error::ClrError;
use crate::Result;
use windows_sys::core::HRESULT;
use windows_core::{IUnknown, GUID, Interface};
use core::{
    ops::Deref, ptr::null_mut, 
    ffi::c_void, mem::transmute
};

/// This struct represents the COM `IEnumUnknown` interface, 
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct IEnumUnknown(windows_core::IUnknown);

/// Implementation of the original `IEnumUnknown` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl IEnumUnknown {
    /// Retrieves the next set of interfaces from the enumerator.
    ///
    /// # Arguments
    ///
    /// * `rgelt` - A mutable slice of `Option<IUnknown>` where the results will be stored.
    /// * `pceltfetched` - An optional pointer that receives the number of elements fetched.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    #[inline]
    pub fn Next(&self, rgelt: &mut [Option<windows_core::IUnknown>], pceltfetched: Option<*mut u32>) -> HRESULT {
        unsafe {
            (Interface::vtable(self).Next)(
                Interface::as_raw(self), 
                rgelt.len() as u32, 
                transmute(rgelt.as_ptr()), 
                transmute(pceltfetched.unwrap_or(core::ptr::null_mut()))
            )
        }
    }

    /// Skips a specified number of elements in the enumeration sequence.
    ///
    /// # Arguments
    ///
    /// * `celt` - The number of elements to skip.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If successful, returns an empty `Ok`.
    /// * `Err(ClrError)` - If skipping fails, returns a `ClrError`.
    pub fn Skip(&self, celt: u32) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Skip)(Interface::as_raw(self), celt) };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("Skip", hr))
        }
    }
    
    /// Resets the enumeration sequence to the beginning.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If successful, returns an empty `Ok`.
    /// * `Err(ClrError)` - If resetting fails, returns a `ClrError`.
    pub fn Reset(&self) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Reset)(Interface::as_raw(self)) };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("Reset", hr))
        }
    }

    /// Creates a new enumerator with the same state as the current one.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut IEnumUnknown)` - If successful, returns a pointer to the new `IEnumUnknown`.
    /// * `Err(ClrError)` - If cloning fails, returns a `ClrError`.
    pub fn Clone(&self) -> Result<*mut IEnumUnknown> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).Clone)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("Clone", hr))
        }
    }
}

unsafe impl Interface for IEnumUnknown {
    type Vtable = IEnumUnknown_Vtbl;

    /// The interface identifier (IID) for the `IEnumUnknown` COM interface.
    ///
    /// This GUID is used to identify the `IEnumUnknown` interface when calling 
    /// COM methods like `QueryInterface`. It is defined based on the standard 
    /// .NET CLR IID for the `IEnumUnknown` interface.
    const IID: GUID = GUID::from_u128(0x00000100_0000_0000_c000_000000000046);
}

impl Deref for IEnumUnknown {
    type Target = windows_core::IUnknown;
    
    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `IEnumUnknown` to be used as an `IUnknown` 
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`, 
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct IEnumUnknown_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    /// 
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Retrieves the next set of interfaces from the enumerator.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `celt` - The number of elements to retrieve.
    /// * `rgelt` - Pointer to an array receiving the retrieved interfaces.
    /// * `pceltFetched` - Pointer to the actual number of elements retrieved.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    pub Next: unsafe extern "system" fn(
        this: *mut c_void, 
        celt: u32, 
        rgelt: *mut *mut IUnknown, 
        pceltFetched: *mut u32
    ) -> HRESULT,
    
    /// Skips the specified number of elements in the enumeration sequence.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `celt` - The number of elements to skip.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    pub Skip: unsafe extern "system" fn(
        this: *mut c_void, 
        celt: u32
    ) -> HRESULT,
    
    /// Resets the enumeration sequence to the beginning.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// 
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    pub Reset: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,

    /// Creates a new enumerator that contains the same state as the current one.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `ppenum` - Pointer to the new `IEnumUnknown`.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    pub Clone: unsafe extern "system" fn(
        this: *mut c_void, 
        ppenum: *mut *mut IEnumUnknown
    ) -> HRESULT,
}