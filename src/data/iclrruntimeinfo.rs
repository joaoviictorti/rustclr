use core::{ffi::c_void, ops::Deref};
use alloc::ffi::CString;

use windows_core::{GUID, Interface, PCSTR, PCWSTR, PWSTR};
use windows_sys::{
    Win32::Foundation::{BOOL, HANDLE, HMODULE},
    core::HRESULT,
};

use crate::error::ClrError;
use crate::Result;

/// This struct represents the COM `ICLRRuntimeInfo` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICLRRuntimeInfo(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `ICLRRuntimeInfo` methods.
impl ICLRRuntimeInfo {
    /// Checks if the CLR runtime has been started.
    ///
    /// # Returns
    ///
    /// * `true` - If the runtime has been started (`started != 0`).
    /// * `false` - If the runtime has not been started or if `IsStarted` fails.
    pub fn is_started(&self) -> bool {
        let mut started = 0;
        let mut startup_flags = 0;

        self.IsStarted(&mut started, &mut startup_flags).is_ok() && started != 0
    }
}

/// Implementation of the original `ICLRRuntimeInfo` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl ICLRRuntimeInfo {
    /// Checks if the .NET runtime is loadable in the current process.
    ///
    /// # Returns
    ///
    /// * `Ok(BOOL)` - A `BOOL` indicating if the runtime is loadable.
    /// * `Err(ClrError)` - If the call fails, returns a `ClrError`.
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
    ///
    /// # Arguments
    ///
    /// * `rclsid` - The class identifier (`GUID`) for the COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(T)` - On success, returns an instance of the requested interface type `T`.
    /// * `Err(ClrError)` - If the call fails, returns a `ClrError`.
    pub fn GetInterface<T>(&self, rclsid: *const GUID) -> Result<T>
    where
        T: Interface,
    {
        unsafe {
            let mut result = core::ptr::null_mut();
            let hr = (Interface::vtable(self).GetInterface)(Interface::as_raw(self), rclsid, &T::IID, &mut result);
            if hr == 0 {
                Ok(core::mem::transmute_copy(&result))
            } else {
                Err(ClrError::ApiError("GetInterface", hr))
            }
        }
    }

    /// Retrieves the version string of the CLR runtime.
    ///
    /// # Arguments
    ///
    /// * `pwzbuffer` - A mutable `PWSTR` buffer for the version string.
    /// * `pcchbuffer` - A pointer to an unsigned integer that specifies
    ///   the buffer size and receives the actual length of the version string.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the version string is written to `pwzbuffer`.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
    pub fn GetVersionString(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetVersionString)(Interface::as_raw(self), pwzbuffer, pcchbuffer);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetVersionString", hr))
            }
        }
    }

    /// Retrieves the directory where the CLR runtime is installed.
    ///
    /// # Arguments
    ///
    /// * `pwzbuffer` - A mutable `PWSTR` buffer to store the runtime directory path.
    /// * `pcchbuffer` - A pointer to an unsigned integer specifying the buffer size.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the directory path is written to `pwzbuffer`.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
    pub fn GetRuntimeDirectory(&self, pwzbuffer: PWSTR, pcchbuffer: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetRuntimeDirectory)(Interface::as_raw(self), pwzbuffer, pcchbuffer);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetRuntimeDirectory", hr))
            }
        }
    }

    /// Checks if the runtime is loaded in a specified process.
    ///
    /// # Arguments
    ///
    /// * `hndProcess` - Handle to the process to check.
    ///
    /// # Returns
    ///
    /// * `Ok(BOOL)` - On success, returns a `BOOL` indicating whether the runtime is loaded.
    /// * `Err(ClrError)` - If the call fails, returns a `ClrError`.
    pub fn IsLoaded(&self, hndProcess: HANDLE) -> Result<BOOL> {
        unsafe {
            let mut pbLoaded = 0;
            let hr = (Interface::vtable(self).IsLoaded)(Interface::as_raw(self), hndProcess, &mut pbLoaded);
            if hr == 0 {
                Ok(pbLoaded)
            } else {
                Err(ClrError::ApiError("IsLoaded", hr))
            }
        }
    }

    /// Loads an error string by its resource ID.
    ///
    /// # Arguments
    ///
    /// * `iResourceID` - Resource ID of the error message to load.
    /// * `pwzBuffer` - A buffer to store the error message.
    /// * `pcchBuffer` - Pointer to the buffer size.
    /// * `iLocaleID` - Locale ID for the message.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the error string is written to `pwzBuffer`.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
    pub fn LoadErrorString(&self, iResourceID: u32, pwzBuffer: PWSTR, pcchBuffer: *mut u32, iLocaleID: i32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).LoadErrorString)(Interface::as_raw(self), iResourceID, pwzBuffer, pcchBuffer, iLocaleID);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("LoadErrorString", hr))
            }
        }
    }

    /// Loads a DLL by name.
    ///
    /// # Arguments
    ///
    /// * `pwzDllName` - The name of the DLL to load.
    ///
    /// # Returns
    ///
    /// * `Ok(HMODULE)` - On success, returns a handle to the loaded module.
    /// * `Err(ClrError)` - If loading fails, returns a `ClrError`.
    pub fn LoadLibraryA(&self, pwzDllName: PCWSTR) -> Result<HMODULE> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).LoadLibraryA)(Interface::as_raw(self), pwzDllName, &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("LoadLibraryA", hr))
            }
        }
    }

    /// Retrieves the address of a procedure in a loaded DLL.
    ///
    /// # Arguments
    ///
    /// * `pszProcName` - Name of the procedure to retrieve.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut c_void)` - On success, returns a pointer to the procedure.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
    pub fn GetProcAddress(&self, pszProcName: &str) -> Result<*mut c_void> {
        unsafe {
            let mut result = core::mem::zeroed();
            let cstr = CString::new(pszProcName).map_err(|_| ClrError::GenericError("Invalid String"))?;
            let hr = (Interface::vtable(self).GetProcAddress)(Interface::as_raw(self), PCSTR(cstr.as_ptr().cast()), &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetProcAddress", hr))
            }
        }
    }

    /// Sets the default startup flags for the runtime.
    ///
    /// # Arguments
    ///
    /// * `dwstartupflags` - Startup flags for the runtime.
    /// * `pwzhostconfigfile` - Path to a configuration file for the runtime.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
    pub fn SetDefaultStartupFlags(&self, dwstartupflags: u32, pwzhostconfigfile: PCWSTR) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).SetDefaultStartupFlags)(Interface::as_raw(self), dwstartupflags, pwzhostconfigfile);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("SetDefaultStartupFlags", hr))
            }
        }
    }

    /// Retrieves the default startup flags for the runtime.
    ///
    /// # Arguments
    ///
    /// * `pdwstartupflags` - Pointer to store the startup flags.
    /// * `pwzhostconfigfile` - Buffer to receive the configuration file path.
    /// * `pcchhostconfigfile` - Pointer to the size of the configuration file path buffer.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the startup flags and file path are written to the respective parameters.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
    pub fn GetDefaultStartupFlags(&self, pdwstartupflags: *mut u32, pwzhostconfigfile: PWSTR, pcchhostconfigfile: *mut u32) -> Result<()> {
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
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
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
    ///
    /// # Arguments
    ///
    /// * `pbstarted` - Pointer to a `BOOL` that receives the runtime's started status.
    /// * `pdwstartupflags` - Pointer to an unsigned integer to receive the startup flags.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
    pub fn IsStarted(&self, pbstarted: *mut BOOL, pdwstartupflags: *mut u32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).IsStarted)(Interface::as_raw(self), pbstarted, pdwstartupflags);
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

#[repr(C)]
pub struct ICLRRuntimeInfo_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    ///
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Retrieves the version string of the CLR runtime.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzBuffer` - A pointer to a buffer that receives the version string.
    /// * `pcchBuffer` - A pointer to an unsigned integer that holds the buffer size and
    ///   receives the length of the version string on return.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetVersionString: unsafe extern "system" fn(this: *mut c_void, pwzBuffer: PWSTR, pcchBuffer: *mut u32) -> HRESULT,

    /// Retrieves the runtime directory path.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzBuffer` - A pointer to a buffer for the directory path.
    /// * `pcchBuffer` - A pointer to an unsigned integer specifying the buffer size.
    ///
    /// # Returns
    ///
    /// * `HRESULT` indicating success or failure.
    pub GetRuntimeDirectory: unsafe extern "system" fn(this: *mut c_void, pwzBuffer: PWSTR, pcchBuffer: *mut u32) -> HRESULT,

    /// Checks if the runtime is loaded in a specified process.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `hndProcess` - A handle to the process.
    /// * `pbLoaded` - A pointer to a `BOOL` that receives the loaded status.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub IsLoaded: unsafe extern "system" fn(this: *mut c_void, hndProcess: HANDLE, pbLoaded: *mut BOOL) -> HRESULT,

    /// Loads a specified error message string.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `iResourceID` - The resource ID of the error message.
    /// * `pwzBuffer` - A buffer that receives the error message.
    /// * `pcchBuffer` - Pointer specifying the buffer size.
    /// * `iLocaleID` - The locale ID.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub LoadErrorString:
        unsafe extern "system" fn(this: *mut c_void, iResourceID: u32, pwzBuffer: PWSTR, pcchBuffer: *mut u32, iLocaleID: i32) -> HRESULT,

    /// Loads a specified DLL library.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzDllName` - The name of the DLL to load.
    /// * `phndModule` - Pointer to a handle receiving the loaded module.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure..
    pub LoadLibraryA: unsafe extern "system" fn(this: *mut c_void, pwzDllName: PCWSTR, phndModule: *mut HMODULE) -> HRESULT,

    /// Retrieves a pointer to a specified procedure within a loaded module.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pszProcName` - The name of the procedure.
    /// * `ppProc` - A pointer to a pointer receiving the procedure address.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetProcAddress: unsafe extern "system" fn(this: *mut c_void, pszProcName: PCSTR, ppProc: *mut *mut c_void) -> HRESULT,

    /// Retrieves a specified interface from the runtime.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `rclsid` - The class identifier for the interface.
    /// * `riid` - The identifier of the interface being requested.
    /// * `ppUnk` - A pointer to a pointer receiving the requested interface.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetInterface: unsafe extern "system" fn(
        this: *mut c_void,
        rclsid: *const windows_core::GUID,
        riid: *const windows_core::GUID,
        ppUnk: *mut *mut c_void,
    ) -> HRESULT,

    /// Checks if the runtime is loadable in the current process.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pbLoadable` - A pointer to a `BOOL` receiving the loadable status.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub IsLoadable: unsafe extern "system" fn(this: *mut c_void, pbLoadable: *mut BOOL) -> HRESULT,

    /// Sets the default startup flags for the runtime.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `dwStartupFlags` - Flags defining runtime startup behavior.
    /// * `pwzHostConfigFile` - A configuration file path for the runtime.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub SetDefaultStartupFlags: unsafe extern "system" fn(this: *mut c_void, dwStartupFlags: u32, pwzHostConfigFile: PCWSTR) -> HRESULT,

    /// Retrieves the default startup flags for the runtime.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `dwStartupFlags` - Pointer to receive the startup flags.
    /// * `pwzHostConfigFile` - Pointer to receive the configuration file path.
    /// * `pcchHostConfigFile` - Pointer to the size of the buffer for `pwzHostConfigFile`.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetDefaultStartupFlags: unsafe extern "system" fn(
        this: *mut c_void,
        dwStartupFlags: *mut u32,
        pwzHostConfigFile: windows_core::PCWSTR,
        pcchHostConfigFile: *mut u32,
    ) -> HRESULT,

    /// Binds the runtime as a legacy v2 runtime.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub BindAsLegacyV2Runtime: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,

    /// Checks if the runtime has started and retrieves startup flags.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pbStarted` - A pointer receiving the started status.
    /// * `pdwStartupFlags` - A pointer receiving the startup flags.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub IsStarted: unsafe extern "system" fn(this: *mut c_void, pbStarted: *mut BOOL, pdwStartupFlags: *mut u32) -> HRESULT,
}
