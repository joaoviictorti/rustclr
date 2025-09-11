use alloc::vec::Vec;
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows_core::{GUID, IUnknown, Interface, PCWSTR};
use windows_sys::{
    Win32::Foundation::{HANDLE, HMODULE},
    core::HRESULT,
};

use super::_AppDomain;
use crate::Result;
use crate::error::ClrError;

/// This struct represents the COM `ICorRuntimeHost` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct ICorRuntimeHost(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `ICorRuntimeHost` methods.
impl ICorRuntimeHost {
    /// Creates a new .NET AppDomain with the specified name.
    ///
    /// # Arguments
    ///
    /// * `name` - A string slice (`&str`) representing the name of the AppDomain to be created.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - On success, returns an instance of `_AppDomain`, representing the created .NET AppDomain.
    /// * `Err(ClrError)` - If the domain creation fails, returns an error variant from `ClrError` describing the issue.
    pub fn create_domain(&self, name: &str) -> Result<_AppDomain> {
        let name = name.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        let domain_name = PCWSTR(name.as_ptr());
        self.CreateDomain(domain_name, null_mut())
    }
}

/// Implementation of the original `ICorRuntimeHost` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl ICorRuntimeHost {
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

    /// Retrieves the default application domain for the runtime host.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - The default application domain.
    /// * `Err(ClrError)` - An error if the domain retrieval fails.
    pub fn GetDefaultDomain(&self) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr =
                (Interface::vtable(self).GetDefaultDomain)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("GetDefaultDomain", hr))
            }
        }
    }

    /// Creates a new application domain with the specified name and identity.
    ///
    /// # Arguments
    ///
    /// * `pwzFriendlyName` - The name for the new application domain.
    /// * `pIdentityArray` - Pointer to an `IUnknown` array representing the domain identity.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - The created application domain.
    /// * `Err(ClrError)` - An error if domain creation fails.
    pub fn CreateDomain(
        &self,
        pwzFriendlyName: PCWSTR,
        pIdentityArray: *mut IUnknown,
    ) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateDomain)(
                Interface::as_raw(self),
                pwzFriendlyName,
                pIdentityArray,
                &mut result,
            );
            if hr == 0 {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CreateDomain", hr))
            }
        }
    }

    /// Creates a logical thread state within the runtime host.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn CreateLogicalThreadState(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).CreateLogicalThreadState)(Interface::as_raw(self));
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("CreateLogicalThreadState", hr))
            }
        }
    }

    /// Deletes the current logical thread state.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn DeleteLogicalThreadState(&self) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).DeleteLogicalThreadState)(Interface::as_raw(self));
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("DeleteLogicalThreadState", hr))
            }
        }
    }

    /// Switches into the logical thread state.
    ///
    /// # Returns
    ///
    /// * `Ok(u32)` - On success, returns the fiber cookie.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn SwitchInLogicalThreadState(&self) -> Result<u32> {
        unsafe {
            let mut result = 0;
            let hr = (Interface::vtable(self).SwitchInLogicalThreadState)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("SwitchInLogicalThreadState", hr))
            }
        }
    }

    /// Switches out of the logical thread state.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut u32)` - On success, returns a pointer to the fiber cookie.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn SwitchOutLogicalThreadState(&self) -> Result<*mut u32> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).SwitchOutLogicalThreadState)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("SwitchOutLogicalThreadState", hr))
            }
        }
    }

    /// Retrieves the number of locks held by the current logical thread.
    ///
    /// # Returns
    ///
    /// * `Ok(u32)` - On success, returns the count of locks.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn LocksHeldByLogicalThread(&self) -> Result<u32> {
        unsafe {
            let mut result = 0;
            let hr = (Interface::vtable(self).LocksHeldByLogicalThread)(
                Interface::as_raw(self),
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("LocksHeldByLogicalThread", hr))
            }
        }
    }

    /// Maps a file handle to an `HMODULE`.
    ///
    /// # Arguments
    ///
    /// * `h_file` - A handle to the file to be mapped.
    ///
    /// # Returns
    ///
    /// * `Ok(HMODULE)` - On success, returns an `HMODULE` for the mapped file.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn MapFile(&self, h_file: HANDLE) -> Result<HMODULE> {
        unsafe {
            let mut result = null_mut();
            let hr =
                (Interface::vtable(self).MapFile)(Interface::as_raw(self), h_file, &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("MapFile", hr))
            }
        }
    }

    /// Retrieves the configuration for the runtime host.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut c_void)` - On success, returns a pointer to the configuration object.
    /// * `Err(ClrError)` - If retrieval fails, returns an error variant from `ClrError`.
    pub fn GetConfiguration(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr =
                (Interface::vtable(self).GetConfiguration)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetConfiguration", hr))
            }
        }
    }

    /// Enumerates application domains managed by the runtime host.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut c_void)` - On success, returns an enumeration handle.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn EnumDomains(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).EnumDomains)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("EnumDomains", hr))
            }
        }
    }

    /// Retrieves the next application domain in the enumeration.
    ///
    /// # Arguments
    ///
    /// * `hEnum` - Handle to the ongoing enumeration.
    ///
    /// # Returns
    ///
    /// * `Ok(IUnknown)` - On success, returns the next app domain.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn NextDomain(&self, hEnum: *mut c_void) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).NextDomain)(Interface::as_raw(self), hEnum, &mut result);
            if hr == 0 {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("NextDomain", hr))
            }
        }
    }

    /// Closes an application domain enumeration.
    ///
    /// # Arguments
    ///
    /// * `hEnum` - Handle to the enumeration to be closed.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn CloseEnum(&self, hEnum: *mut c_void) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).CloseEnum)(Interface::as_raw(self), hEnum);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("CloseEnum", hr))
            }
        }
    }

    /// Creates a new application domain with specified configuration.
    ///
    /// # Arguments
    ///
    /// * `pwzFriendlyName` - The friendly name for the new app domain.
    /// * `psSetup` - Pointer to setup configuration.
    /// * `pEvidence` - Pointer to evidence object.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - On success, returns the new app domain.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn CreateDomainEx(
        &self,
        pwzFriendlyName: PCWSTR,
        psSetup: *mut IUnknown,
        pEvidence: *mut IUnknown,
    ) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateDomainEx)(
                Interface::as_raw(self),
                pwzFriendlyName,
                psSetup,
                pEvidence,
                &mut result,
            );
            if hr == 0 {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CreateDomainEx", hr))
            }
        }
    }

    /// Creates a setup configuration object for application domains.
    ///
    /// # Returns
    ///
    /// * `Ok(IUnknown)` - On success, returns the setup configuration object.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn CreateDomainSetup(&self) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr =
                (Interface::vtable(self).CreateDomainSetup)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("CreateDomainSetup", hr))
            }
        }
    }

    /// Creates an evidence object for application domains.
    ///
    /// # Returns
    ///
    /// * `Ok(IUnknown)` - On success, returns the evidence object.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn CreateEvidence(&self) -> Result<IUnknown> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CreateEvidence)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(IUnknown::from_raw(result as *mut c_void))
            } else {
                Err(ClrError::ApiError("CreateEvidence", hr))
            }
        }
    }

    /// Unloads the specified application domain.
    ///
    /// # Arguments
    ///
    /// * `pAppDomain` - Pointer to the app domain to unload.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - On success.
    /// * `Err(ClrError)` - If the operation fails, returns an error variant from `ClrError`.
    pub fn UnloadDomain(&self, pAppDomain: *mut IUnknown) -> Result<()> {
        unsafe {
            let hr = (Interface::vtable(self).UnloadDomain)(Interface::as_raw(self), pAppDomain);
            if hr == 0 {
                Ok(())
            } else {
                Err(ClrError::ApiError("UnloadDomain", hr))
            }
        }
    }

    /// Retrieves the current application domain.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - On success, returns the current app domain.
    /// * `Err(ClrError)` - If retrieval fails, returns an error variant from `ClrError`.
    pub fn CurrentDomain(&self) -> Result<_AppDomain> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).CurrentDomain)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                _AppDomain::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("CurrentDomain", hr))
            }
        }
    }
}

unsafe impl Interface for ICorRuntimeHost {
    type Vtable = ICorRuntimeHost_Vtbl;

    /// The interface identifier (IID) for the `ICorRuntimeHost` COM interface.
    ///
    /// This GUID is used to identify the `ICorRuntimeHost` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICorRuntimeHost` interface.
    const IID: GUID = GUID::from_u128(0xCB2F6722_AB3A_11D2_9C40_00C04FA30A3E);
}

impl Deref for ICorRuntimeHost {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `ICorRuntimeHost` COM interface.
    ///
    /// This GUID is used to identify the `ICorRuntimeHost` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICorRuntimeHost` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct ICorRuntimeHost_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    ///
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Initializes a logical thread state.
    pub CreateLogicalThreadState: unsafe extern "system" fn(*mut c_void) -> HRESULT,

    /// Deletes a logical thread state.
    pub DeleteLogicalThreadState: unsafe extern "system" fn(*mut c_void) -> HRESULT,

    /// Switches into a logical thread state.
    ///
    /// # Arguments
    ///
    /// * `pFiberCookie` - Pointer to a `u32` used to track the fiber state.s
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub SwitchInLogicalThreadState: unsafe extern "system" fn(
        this: *mut c_void, 
        pFiberCookie: *mut u32
    ) -> HRESULT,

    /// Switches out of a logical thread state.
    ///
    /// # Arguments
    ///
    /// * `pFiberCookie` - Pointer to a `u32` that holds the fiber cookie to switch out.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub SwitchOutLogicalThreadState: unsafe extern "system" fn(
        this: *mut c_void, 
        pFiberCookie: *mut *mut u32
    ) -> HRESULT,

    /// Retrieves the number of locks held by the logical thread.
    ///
    /// # Arguments
    ///
    /// * `pCount` - Pointer to a `u32` where the count is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub LocksHeldByLogicalThread: unsafe extern "system" fn(
        this: *mut c_void, 
        pCount: *mut u32
    ) -> HRESULT,

    /// Maps a file into memory.
    ///
    /// # Arguments
    ///
    /// * `hFile` - The handle to the file.
    /// * `hMapAddress` - Pointer to an `HMODULE` where the mapped address is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub MapFile: unsafe extern "system" fn(
        this: *mut c_void,
        hFile: HANDLE,
        hMapAddress: *mut HMODULE,
    ) -> HRESULT,

    /// Retrieves configuration information for the runtime host.
    ///
    /// # Arguments
    ///
    /// * `pConfiguration` - Pointer to a configuration object.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetConfiguration: unsafe extern "system" fn(
        this: *mut c_void, 
        pConfiguration: *mut *mut c_void
    ) -> HRESULT,

    /// Starts the CLR runtime host.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub Start: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,

    /// Stops the CLR runtime host.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub Stop: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,

    /// Creates a new application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzFriendlyName` - The friendly name for the new domain.
    /// * `pIdentityArray` - Pointer to an array of identities.
    /// * `pAppDomain` - Pointer to where the created `AppDomain` is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CreateDomain: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFriendlyName: PCWSTR,
        pIdentityArray: *mut IUnknown,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,

    /// Retrieves the default application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pAppDomain` - Pointer to where the default application domain is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub GetDefaultDomain: unsafe extern "system" fn(
        this: *mut c_void, 
        pAppDomain: *mut *mut IUnknown
    ) -> HRESULT,

    /// Enumerates the application domains.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `hEnum` - Pointer to the enumeration handle, where the results will be stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub EnumDomains: unsafe extern "system" fn(
        this: *mut c_void, 
        hEnum: *mut *mut c_void
    ) -> HRESULT,

    /// Retrieves the next application domain in the enumeration.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `hEnum` - Handle to the enumeration.
    /// * `pAppDomain` - Pointer to the next application domain.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub NextDomain: unsafe extern "system" fn(
        this: *mut c_void,
        hEnum: *mut c_void,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,

    /// Closes the domain enumeration.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `hEnum` - Handle to the enumeration.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CloseEnum: unsafe extern "system" fn(this: *mut c_void, hEnum: *mut c_void) -> HRESULT,

    /// Creates a new application domain with additional configuration.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pwzFriendlyName` - The friendly name for the new domain.
    /// * `pSetup` - Pointer to the setup configuration.
    /// * `pEvidence` - Pointer to the evidence object.
    /// * `pAppDomain` - Pointer to where the created `AppDomain` is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CreateDomainEx: unsafe extern "system" fn(
        this: *mut c_void,
        pwzFriendlyName: PCWSTR,
        pSetup: *mut IUnknown,
        pEvidence: *mut IUnknown,
        pAppDomain: *mut *mut IUnknown,
    ) -> HRESULT,

    /// Creates a setup configuration for an application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pAppDomainSetup` - Pointer to where the setup configuration is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CreateDomainSetup: unsafe extern "system" fn(
        this: *mut c_void,
        pAppDomainSetup: *mut *mut IUnknown,
    ) -> HRESULT,

    /// Creates an evidence object for an application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pEvidence` - Pointer to where the evidence object is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CreateEvidence: unsafe extern "system" fn(
        this: *mut c_void, 
        pEvidence: *mut *mut IUnknown
    ) -> HRESULT,

    /// Unloads the specified application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pAppDomain` - Pointer to the application domain to unload.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub UnloadDomain: unsafe extern "system" fn(
        this: *mut c_void, 
        pAppDomain: *mut IUnknown
    ) -> HRESULT,

    /// Retrieves the current application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pAppDomain` - Pointer to where the current application domain is stored.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    pub CurrentDomain: unsafe extern "system" fn(
        this: *mut c_void, 
        pAppDomain: *mut *mut IUnknown
    ) -> HRESULT,
}
