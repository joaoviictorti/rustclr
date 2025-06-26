use alloc::{collections::BTreeMap, string::String, vec};
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows_core::{GUID, Interface, PCWSTR, PWSTR};
use windows_sys::{Win32::Foundation::HANDLE, core::HRESULT};

use super::{ICLRRuntimeInfo, IEnumUnknown};
use crate::Result;
use crate::error::ClrError;

/// Function pointer type for the callback invoked when a runtime is loaded.
///
/// # Arguments
///
/// * `pruntimeinfo` - An optional pointer to `ICLRRuntimeInfo`, containing information about the loaded runtime.
/// * `pfncallbackthreadset` - A pointer to the callback function for setting threads.
/// * `pfncallbackthreadunset` - A pointer to the callback function for unsetting threads.
pub type RuntimeLoadedCallbackFnPtr = Option<
    unsafe extern "system" fn(
        pruntimeinfo: *mut ICLRRuntimeInfo,
        pfncallbackthreadset: CallbackThreadSetFnPtr,
        pfncallbackthreadunset: CallbackThreadUnsetFnPtr,
    ),
>;

/// Function pointer for setting the callback thread in the CLR.
pub type CallbackThreadSetFnPtr = Option<unsafe extern "system" fn() -> HRESULT>;

/// Function pointer for unsetting the callback thread in the CLR.
pub type CallbackThreadUnsetFnPtr = Option<unsafe extern "system" fn() -> HRESULT>;

/// This struct represents the COM `ICLRMetaHoste` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICLRMetaHost(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `ICLRMetaHost` methods.
impl ICLRMetaHost {
    /// Retrieves a map of available runtime versions and corresponding runtime information.
    ///
    /// # Returns
    ///
    /// * `Ok(BTreeMap<String, ICLRRuntimeInfo>)` - A map where keys are runtime versions (as strings) and values
    ///   are `ICLRRuntimeInfo` instances with details about each runtime.
    /// * `Err(ClrError)` - Returns a `ClrError::CastingError` if casting to `ICLRRuntimeInfo` fails.
    pub fn runtimes(&self) -> Result<BTreeMap<String, ICLRRuntimeInfo>> {
        let enum_unknown = self.EnumerateInstalledRuntimes()?;
        let mut fetched = 0;
        let mut rgelt = [None];
        let mut runtimes = BTreeMap::new();

        while enum_unknown.Next(&mut rgelt, Some(&mut fetched)) == 0 && fetched > 0 {
            let runtime_info = match &rgelt[0] {
                Some(unknown) => unknown
                    .cast::<ICLRRuntimeInfo>()
                    .map_err(|_| ClrError::CastingError("ICLRRuntimeInfo"))?,
                None => continue,
            };

            let mut version_string = vec![0u16; 256];
            let mut len = version_string.len() as u32;
            runtime_info.GetVersionString(PWSTR(version_string.as_mut_ptr()), &mut len)?;
            version_string.retain(|&c| c != 0);

            let version = String::from_utf16_lossy(&version_string);
            runtimes.insert(version, runtime_info);
        }

        Ok(runtimes)
    }
}

/// Implementation of the original `_Assembly` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl ICLRMetaHost {
    /// Retrieves a runtime based on the specified version.
    ///
    /// # Arguments
    ///
    /// * `pwzversion` - A `PCWSTR` reference to the .NET runtime version to retrieve (e.g., `"v4.0"`).
    ///
    /// # Returns
    ///
    /// * `Ok(T)` - Returns the requested runtime as the generic type `T` if successful.
    /// * `Err(ClrError)` - Returns a `ClrError::ApiError` if the runtime could not be retrieved.
    #[inline]
    pub fn GetRuntime<T>(&self, pwzversion: PCWSTR) -> Result<T>
    where
        T: Interface,
    {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetRuntime)(
                Interface::as_raw(self),
                pwzversion,
                &T::IID,
                &mut result,
            );
            if hr == 0 {
                Ok(core::mem::transmute_copy(&result))
            } else {
                Err(ClrError::ApiError("GetRuntime", hr))
            }
        }
    }

    /// Enumerates all installed runtimes on the system.
    ///
    /// # Returns
    ///
    /// * `Ok(IEnumUnknown)` - An enumerator containing all installed CLR runtimes.
    /// * `Err(ClrError)` - Returns a `ClrError::ApiError` if enumeration fails.
    pub fn EnumerateInstalledRuntimes(&self) -> Result<IEnumUnknown> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).EnumerateInstalledRuntimes)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr == 0 {
                Ok(IEnumUnknown::from_raw(result))
            } else {
                Err(ClrError::ApiError("EnumerateInstalledRuntimes", hr))
            }
        }
    }

    /// Retrieves the CLR version from a specified file.
    ///
    /// # Arguments
    ///
    /// * `pwzfilepath` - A `PCWSTR` pointing to the file path.
    /// * `pwzbuffer` - A mutable `PWSTR` buffer to store the version string.
    /// * `pcchbuffer` - A pointer to an unsigned integer representing the buffer size.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the version string is written to `pwzbuffer`.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
    pub fn GetVersionFromFile(
        &self,
        pwzfilepath: PCWSTR,
        pwzbuffer: PWSTR,
        pcchbuffer: *mut u32,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).GetVersionFromFile)(
                Interface::as_raw(self),
                pwzfilepath,
                pwzbuffer,
                pcchbuffer,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("GetVersionFromFile", hr))
            }
        }
    }

    /// Enumerates all loaded CLR runtimes in the specified process.
    ///
    /// # Arguments
    ///
    /// * `hndprocess` - A handle to the process to inspect.
    ///
    /// # Returns
    ///
    /// * `Ok(IEnumUnknown)` - On success, returns an enumerator for loaded runtimes.
    /// * `Err(ClrError)` - If enumeration fails, returns a `ClrError`.
    pub fn EnumerateLoadedRuntimes(&self, hndprocess: HANDLE) -> Result<IEnumUnknown> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).EnumerateLoadedRuntimes)(
                Interface::as_raw(self),
                hndprocess,
                &mut result,
            );
            if hr == 0 {
                Ok(IEnumUnknown::from_raw(result))
            } else {
                Err(ClrError::ApiError("EnumerateLoadedRuntimes", hr))
            }
        }
    }

    /// Registers a callback notification for when a runtime is loaded.
    ///
    /// # Arguments
    ///
    /// * `pcallbackfunction` - A pointer to the callback function.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the callback is registered.
    /// * `Err(ClrError)` - If registration fails, returns a `ClrError`.
    pub fn RequestRuntimeLoadedNotification(
        &self,
        pcallbackfunction: RuntimeLoadedCallbackFnPtr,
    ) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).RequestRuntimeLoadedNotification)(
                Interface::as_raw(self),
                pcallbackfunction,
            );
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("RequestRuntimeLoadedNotification", hr))
            }
        }
    }

    /// Queries for a legacy .NET v2 runtime binding.
    ///
    /// # Returns
    ///
    /// * `Ok(T)` - On success, returns an instance of the requested legacy binding as type `T`.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
    pub fn QueryLegacyV2RuntimeBinding<T>(&self) -> Result<T>
    where
        T: Interface,
    {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).QueryLegacyV2RuntimeBinding)(
                Interface::as_raw(self),
                &T::IID,
                &mut result,
            );
            if hr == 0 {
                Ok(core::mem::transmute_copy(&result))
            } else {
                Err(ClrError::ApiError("QueryLegacyV2RuntimeBinding", hr))
            }
        }
    }

    /// Terminates the process with the specified exit code.
    ///
    /// # Arguments
    ///
    /// * `iexitcode` - An integer specifying the process exit code.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success, the process is terminated.
    /// * `Err(ClrError)` - If the operation fails, returns a `ClrError`.
    pub fn ExitProcess(&self, iexitcode: i32) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).ExitProcess)(Interface::as_raw(self), iexitcode);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("ExitProcess", hr))
            }
        }
    }
}

unsafe impl Interface for ICLRMetaHost {
    type Vtable = ICLRMetaHost_Vtbl;

    /// The interface identifier (IID) for the `ICLRMetaHost` COM interface.
    ///
    /// This GUID is used to identify the `ICLRMetaHost` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRMetaHost` interface.
    const IID: GUID = GUID::from_u128(0xd332db9e_b9b3_4125_8207_a14884f53216);
}

impl Deref for ICLRMetaHost {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `ICLRMetaHost` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Vtable structure for the `ICLRMetaHost` interface, defining the available methods.
///
/// These methods provide functionality such as retrieving runtime information, enumerating installed
/// and loaded runtimes, and requesting runtime notifications.
#[repr(C)]
pub struct ICLRMetaHost_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    ///
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Retrieves a runtime based on the specified version.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzVersion` - Version of the runtime.
    /// * `riid` - GUID of the requested interface.
    /// * `ppRuntime` - Pointer to the interface.
    pub GetRuntime: unsafe extern "system" fn(
        this: *mut c_void,
        pwzVersion: PCWSTR,
        riid: *const GUID,
        ppRuntime: *mut *mut c_void,
    ) -> HRESULT,

    /// Retrieves the version of the CLR from a file.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzFilePath` - Path to the file.
    /// * `pwzBuffer` - Buffer for the version string.
    /// * `pcchBuffer` - Length of the buffer.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetVersionFromFile: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFilePath: PCWSTR,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,

    /// Enumerates all installed runtimes on the system.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `ppEnumerator` - Pointer to the enumerator.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub EnumerateInstalledRuntimes: unsafe extern "system" fn(
        this: *mut c_void, 
        ppEnumerator: *mut *mut c_void
    ) -> HRESULT,

    /// Enumerates all loaded runtimes in the specified process.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `hndProcess` - Handle to the process.
    /// * `ppEnumerator` - Pointer to the enumerator.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub EnumerateLoadedRuntimes: unsafe extern "system" fn(
        this: *mut c_void,
        hndProcess: HANDLE,
        ppEnumerator: *mut *mut c_void,
    ) -> HRESULT,

    /// Registers a notification callback for when a runtime is loaded.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pCallbackFunction` - Callback function to be invoked.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub RequestRuntimeLoadedNotification: unsafe extern "system" fn(
        this: *mut c_void,
        pCallbackFunction: RuntimeLoadedCallbackFnPtr,
    ) -> HRESULT,

    /// Queries for a legacy runtime binding.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `riid` - GUID of the legacy runtime binding.
    /// * `ppUnk` - Pointer to the resulting binding.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub QueryLegacyV2RuntimeBinding: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const GUID,
        ppUnk: *mut *mut c_void,
    ) -> HRESULT,

    /// Terminates the process by calling the CLR's `ExitProcess` method.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `iExitCode` - Exit code.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub ExitProcess: unsafe extern "system" fn(this: *mut c_void, iExitCode: i32) -> HRESULT,
}
