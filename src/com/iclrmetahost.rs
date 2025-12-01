use alloc::{collections::BTreeMap, string::String, vec};
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows_core::{GUID, Interface, PCWSTR, PWSTR};
use windows_sys::{Win32::Foundation::HANDLE, core::HRESULT};

use crate::error::{ClrError, Result};
use super::{ICLRRuntimeInfo, IEnumUnknown};

/// Function pointer for setting the callback thread in the CLR.
pub type CallbackThreadSetFnPtr = Option<unsafe extern "system" fn() -> HRESULT>;

/// Function pointer for unsetting the callback thread in the CLR.
pub type CallbackThreadUnsetFnPtr = Option<unsafe extern "system" fn() -> HRESULT>;

/// Function pointer type for the callback invoked when a runtime is loaded.
pub type RuntimeLoadedCallbackFnPtr = Option<
    unsafe extern "system" fn(
        pruntimeinfo: *mut ICLRRuntimeInfo,
        pfncallbackthreadset: CallbackThreadSetFnPtr,
        pfncallbackthreadunset: CallbackThreadUnsetFnPtr,
    ),
>;

/// This struct represents the COM `ICLRMetaHost` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICLRMetaHost(windows_core::IUnknown);

impl ICLRMetaHost {
    /// Retrieves a map of available runtime versions and corresponding runtime information.
    #[inline]
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

    /// Retrieves a runtime based on the specified version.
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
    #[inline]
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
    #[inline]
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
    #[inline]
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
    #[inline]
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
    #[inline]
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
    #[inline]
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

/// Raw COM vtable for the `ICLRMetaHost` interface.
#[repr(C)]
pub struct ICLRMetaHost_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub GetRuntime: unsafe extern "system" fn(
        this: *mut c_void,
        pwzVersion: PCWSTR,
        riid: *const GUID,
        ppRuntime: *mut *mut c_void,
    ) -> HRESULT,
    pub GetVersionFromFile: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFilePath: PCWSTR,
        pwzBuffer: PWSTR,
        pcchBuffer: *mut u32,
    ) -> HRESULT,
    pub EnumerateInstalledRuntimes: unsafe extern "system" fn(
        this: *mut c_void, 
        ppEnumerator: *mut *mut c_void
    ) -> HRESULT,
    pub EnumerateLoadedRuntimes: unsafe extern "system" fn(
        this: *mut c_void,
        hndProcess: HANDLE,
        ppEnumerator: *mut *mut c_void,
    ) -> HRESULT,
    pub RequestRuntimeLoadedNotification: unsafe extern "system" fn(
        this: *mut c_void,
        pCallbackFunction: RuntimeLoadedCallbackFnPtr,
    ) -> HRESULT,
    pub QueryLegacyV2RuntimeBinding: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const GUID,
        ppUnk: *mut *mut c_void,
    ) -> HRESULT,
    pub ExitProcess: unsafe extern "system" fn(this: *mut c_void, iExitCode: i32) -> HRESULT,
}
