use core::{ffi::c_void, ops::Deref};
use windows_core::{GUID, Interface};
use windows_sys::core::HRESULT;
use crate::{Result, data::IHostControl, error::ClrError};

/// This struct represents the COM `ICLRuntimeHost` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICLRuntimeHost(windows_core::IUnknown);

/// Implementation of the original `ICLRuntimeHost` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl ICLRuntimeHost {
    /// Starts the .NET runtime host.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    #[inline]
    pub fn Start(&self) -> HRESULT {
        unsafe { (Interface::vtable(self).Start)(Interface::as_raw(self)) }
    }

    /// Stops the .NET runtime host.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    #[inline]
    pub fn Stop(&self) -> HRESULT {
        unsafe { (Interface::vtable(self).Stop)(Interface::as_raw(self)) }
    }

    /// Assigns a host control implementation to the CLR runtime.
    ///
    /// # Arguments
    ///
    /// * `phostcontrol` - An object implementing the `IHostControl` interface.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the call succeeded.
    /// * `Err(ClrError)` - If the underlying COM call failed.
    pub fn SetHostControl<T>(&self, phostcontrol: T) -> Result<()>
    where
        T: windows_core::Param<IHostControl>,
    {
        let hr = unsafe {
            (Interface::vtable(self).SetHostControl)(
                Interface::as_raw(self),
                phostcontrol.param().abi(),
            )
        };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("SetHostControl", hr))
        }
    }
}

unsafe impl Interface for ICLRuntimeHost {
    type Vtable = ICLRuntimeHost_Vtbl;

    /// The interface identifier (IID) for the `ICLRuntimeHost` COM interface.
    ///
    /// This GUID is used to identify the `ICLRuntimeHost` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRuntimeHost` interface.
    const IID: GUID = GUID::from_u128(0x90f1a06c_7712_4762_86b5_7a5eba6bdb02);
}

impl Deref for ICLRuntimeHost {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `ICLRuntimeHost` COM interface.
    ///
    /// This GUID is used to identify the `ICLRuntimeHost` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRuntimeHost` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `ICLRuntimeHost` interface.
#[repr(C)]
pub struct ICLRuntimeHost_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub Start: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub Stop: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub SetHostControl: unsafe extern "system" fn(
        this: *mut c_void, 
        phostcontrol: *mut c_void
    ) -> HRESULT,
    pub GetCLRControl: *const c_void,
    pub UnloadAppDomain: *const c_void,
    pub ExecuteInAppDomain: *const c_void,
    pub GetCurrentAppDomainId: *const c_void,
    pub ExecuteApplication: *const c_void,
    pub ExecuteInDefaultAppDomain: *const c_void,
}
