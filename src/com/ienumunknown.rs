use core::{
    ffi::c_void, 
    mem::transmute, 
    ops::Deref, 
    ptr::null_mut
};

use windows_core::{GUID, IUnknown, Interface};
use windows_sys::core::HRESULT;

use crate::error::{ClrError, Result};

/// This struct represents the COM `IEnumUnknown` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct IEnumUnknown(windows_core::IUnknown);

impl IEnumUnknown {
    /// Retrieves the next set of interfaces from the enumerator.
    #[inline]
    pub fn Next(
        &self,
        rgelt: &mut [Option<windows_core::IUnknown>],
        pceltfetched: Option<*mut u32>,
    ) -> HRESULT {
        unsafe {
            (Interface::vtable(self).Next)(
                Interface::as_raw(self),
                rgelt.len() as u32,
                transmute(rgelt.as_ptr()),
                transmute(pceltfetched.unwrap_or(core::ptr::null_mut())),
            )
        }
    }

    /// Skips a specified number of elements in the enumeration sequence.
    #[inline]
    pub fn Skip(&self, celt: u32) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Skip)(Interface::as_raw(self), celt) };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("Skip", hr))
        }
    }

    /// Resets the enumeration sequence to the beginning.
    #[inline]
    pub fn Reset(&self) -> Result<()> {
        let hr = unsafe { (Interface::vtable(self).Reset)(Interface::as_raw(self)) };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("Reset", hr))
        }
    }

    /// Creates a new enumerator with the same state as the current one.
    #[inline]
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

/// Raw COM vtable for the `IEnumUnknown` interface.
#[repr(C)]
pub struct IEnumUnknown_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub Next: unsafe extern "system" fn(
        this: *mut c_void,
        celt: u32,
        rgelt: *mut *mut IUnknown,
        pceltFetched: *mut u32,
    ) -> HRESULT,
    pub Skip: unsafe extern "system" fn(this: *mut c_void, celt: u32) -> HRESULT,
    pub Reset: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub Clone: unsafe extern "system" fn(
        this: *mut c_void, 
        ppenum: *mut *mut IEnumUnknown
    ) -> HRESULT,
}
