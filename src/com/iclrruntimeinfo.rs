use alloc::ffi::CString;
use core::{ffi::c_void, ops::Deref};

use windows_core::{GUID, Interface, PCSTR, PCWSTR, PWSTR};
use windows_sys::{
    Win32::Foundation::{BOOL, HANDLE, HMODULE},
    core::HRESULT,
};

use crate::error::{ClrError, Result};

/// This struct represents the COM `ICLRRuntimeInfo` interface;
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICLRRuntimeInfo(windows_core::IUnknown);

impl ICLRRuntimeInfo {
    /// Checks if the CLR runtime has been started.
    #[inline]
    pub fn is_started(&self) -> bool {
        let mut started = 0;
        let mut startup_flags = 0;
        self.IsStarted(&mut started, &mut startup_flags).is_ok() && started != 0
    }

    /// Checks if the .NET runtime is loadable in the current process.
    #[inline]
    pub fn IsLoadable(&self) -> Result<BOOL> {
        unsafe {
            let mut result = 0;
            let hr = (Interface::vtable(self).IsLoadable)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("IsLoadable", hr))
            }
        }
    }

    /// Retrieves a COM interface by its class identifier.
    #[inline]
    pub fn GetInterface<T>(&self, rclsid: *const GUID) -> Result<T>
    where
        T: Interface,
    {
        unsafe {
            let mut result = core::ptr::null_mut();
            let hr = (Interface::vtable(self).GetInterface)(
                Interface::as_raw(self),
                rclsid,
                &T::IID,
                &mut result,
            );
            if hr == 0 {
                Ok(core::mem::transmute_copy(&result))
            } else {
                Err(ClrError::ApiError("GetInterface", hr))
            }
        }
    }

    /// Retrieves the version string of the CLR runtime.
    #[inline]
    pub fn GetVersionString(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetVersionString)(
                Interface::as_raw(self),
                pwzbuffer,
                pcchbuffer,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetVersionString", hr))
            }
        }
    }

    /// Retrieves the directory where the CLR runtime is installed.
    #[inline]
    pub fn GetRuntimeDirectory(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetRuntimeDirectory)(
                Interface::as_raw(self),
                pwzbuffer,
                pcchbuffer,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetRuntimeDirectory", hr))
            }
        }
    }

    /// Checks if the runtime is loaded in a specified process.
    #[inline]
    pub fn IsLoaded(&self, hndProcess: HANDLE) -> Result<BOOL> {
        unsafe {
            let mut pbLoaded = 0;
            let hr = (Interface::vtable(self).IsLoaded)(
                Interface::as_raw(self),
                hndProcess,
                &mut pbLoaded,
            );
            if hr == 0 {
                Ok(pbLoaded)
            } else {
                Err(ClrError::ApiError("IsLoaded", hr))
            }
        }
    }

    /// Loads an error string by its resource ID.
    #[inline]
    pub fn LoadErrorString(
        &self,
        iResourceID: u32,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
        iLocaleID: i32,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).LoadErrorString)(
                Interface::as_raw(self),
                iResourceID,
                pwzBuffer,
                pcchBuffer,
                iLocaleID,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("LoadErrorString", hr))
            }
        }
    }

    /// Loads a DLL by name.
    #[inline]
    pub fn LoadLibraryA(&self, pwzDllName: PCWSTR) -> Result<HMODULE> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).LoadLibraryA)(
                Interface::as_raw(self),
                pwzDllName,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("LoadLibraryA", hr))
            }
        }
    }

    /// Retrieves the address of a procedure in a loaded DLL.
    #[inline]
    pub fn GetProcAddress(&self, pszProcName: &str) -> Result<*mut c_void> {
        unsafe {
            let mut result = core::mem::zeroed();
            let cstr =
                CString::new(pszProcName).map_err(|_| ClrError::Msg("invalid String"))?;
            let hr = (Interface::vtable(self).GetProcAddress)(
                Interface::as_raw(self),
                PCSTR(cstr.as_ptr().cast()),
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetProcAddress", hr))
            }
        }
    }

    /// Sets the default startup flags for the runtime.
    #[inline]
    pub fn SetDefaultStartupFlags(
        &self,
        dwstartupflags: u32,
        pwzhostconfigfile: PCWSTR,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).SetDefaultStartupFlags)(
                Interface::as_raw(self),
                dwstartupflags,
                pwzhostconfigfile,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("SetDefaultStartupFlags", hr))
            }
        }
    }

    /// Retrieves the default startup flags for the runtime.
    #[inline]
    pub fn GetDefaultStartupFlags(
        &self,
        pdwstartupflags: *mut u32,
        pwzhostconfigfile: PWSTR,
        pcchhostconfigfile: *mut u32,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetDefaultStartupFlags)(
                Interface::as_raw(self),
                pdwstartupflags,
                core::mem::transmute(pwzhostconfigfile),
                pcchhostconfigfile,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetDefaultStartupFlags", hr))
            }
        }
    }

    /// Configures the runtime to behave as a legacy v2 runtime.
    #[inline]
    pub fn BindAsLegacyV2Runtime(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).BindAsLegacyV2Runtime)(Interface::as_raw(self));
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("BindAsLegacyV2Runtime", hr))
            }
        }
    }

    /// Checks if the runtime has started and retrieves startup flags.
    #[inline]
    pub fn IsStarted(&self, pbstarted: *mut BOOL, pdwstartupflags: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).IsStarted)(
                Interface::as_raw(self),
                pbstarted,
                pdwstartupflags,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("IsStarted", hr))
            }
        }
    }
}

unsafe impl Interface for ICLRRuntimeInfo {
    type Vtable = ICLRRuntimeInfo_Vtbl;

    /// The interface identifier (IID) for the `ICLRRuntimeInfo` COM interface.
    ///
    /// This GUID is used to identify the `ICLRRuntimeInfo` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRRuntimeInfo` interface.
    const IID: GUID = GUID::from_u128(0xbd39d1d2_ba2f_486a_89b0_b4b0cb466891);
}

impl Deref for ICLRRuntimeInfo {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `ICLRRuntimeInfo` COM interface.
    ///
    /// This GUID is used to identify the `ICLRRuntimeInfo` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRRuntimeInfo` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `ICLRRuntimeInfo` interface.
#[repr(C)]
pub struct ICLRRuntimeInfo_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub GetVersionString: unsafe extern "system" fn(
        this: *mut c_void,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,
    pub GetRuntimeDirectory: unsafe extern "system" fn(
        this: *mut c_void,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,
    pub IsLoaded: unsafe extern "system" fn(
        this: *mut c_void,
        hndProcess: HANDLE,
        pbLoaded: *mut BOOL,
    ) -> HRESULT,
    pub LoadErrorString: unsafe extern "system" fn(
        this: *mut c_void,
        iResourceID: u32,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
        iLocaleID: i32,
    ) -> HRESULT,
    pub LoadLibraryA: unsafe extern "system" fn(
        this: *mut c_void,
        pwzDllName: PCWSTR,
        phndModule: *mut HMODULE,
    ) -> HRESULT,
    pub GetProcAddress: unsafe extern "system" fn(
        this: *mut c_void,
        pszProcName: PCSTR,
        ppProc: *mut *mut c_void,
    ) -> HRESULT,
    pub GetInterface: unsafe extern "system" fn(
        this: *mut c_void,
        rclsid: *const windows_core::GUID,
        riid: *const windows_core::GUID,
        ppUnk: *mut *mut c_void,
    ) -> HRESULT,
    pub IsLoadable: unsafe extern "system" fn(this: *mut c_void, pbLoadable: *mut BOOL) -> HRESULT,
    pub SetDefaultStartupFlags: unsafe extern "system" fn(
        this: *mut c_void,
        dwStartupFlags: u32,
        pwzHostConfigFile: PCWSTR,
    ) -> HRESULT,
    pub GetDefaultStartupFlags: unsafe extern "system" fn(
        this: *mut c_void,
        dwStartupFlags: *mut u32,
        pwzHostConfigFile: windows_core::PCWSTR,
        pcchHostConfigFile: *mut u32,
    ) -> HRESULT,
    pub BindAsLegacyV2Runtime: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub IsStarted: unsafe extern "system" fn(
        this: *mut c_void,
        pbStarted: *mut BOOL,
        pdwStartupFlags: *mut u32,
    ) -> HRESULT,
}
