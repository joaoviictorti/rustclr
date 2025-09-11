use alloc::{string::String, vec};
use core::{ffi::c_void, ops::Deref};
use windows_core::{GUID, IUnknown, Interface, PWSTR};
use windows_sys::core::HRESULT;
use crate::{Result, error::ClrError};

/// This struct represents the COM `ICLRAssemblyIdentityManager` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct ICLRAssemblyIdentityManager(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `ICLRAssemblyIdentityManager` methods.
impl ICLRAssemblyIdentityManager {
    /// Extracts the textual identity of an assembly from a binary stream.
    ///
    /// # Arguments
    ///
    /// * `pstream` - Pointer to a `IStream` containing the assembly data.
    /// * `dwFlags` - Flags to control the extraction behavior.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The string representation of the assembly's identity.
    /// * `Err(ClrError)` - If the operation fails or returns an error HRESULT.
    pub fn get_identity_stream(&self, pstream: *mut c_void, dwFlags: u32) -> Result<String> {
        let mut buffer = vec![0; 2048];
        let mut size = buffer.len() as u32;

        self.GetBindingIdentityFromStream(pstream, dwFlags, PWSTR(buffer.as_mut_ptr()), &mut size)?;
        Ok(String::from_utf16_lossy(&buffer[..size as usize - 1]))
    }

    /// Creates an `ICLRAssemblyIdentityManager` instance from a raw COM interface pointer.
    ///
    /// # Arguments
    ///
    /// * `raw` - A raw pointer to an `IUnknown` COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(ICLRAssemblyIdentityManager)` - Wraps the given COM interface as `ICLRAssemblyIdentityManager`.
    /// * `Err(ClrError)` - If casting fails, returns a `ClrError`.
    #[inline(always)]
    pub fn from_raw(raw: *mut c_void) -> Result<ICLRAssemblyIdentityManager> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<ICLRAssemblyIdentityManager>()
            .map_err(|_| ClrError::CastingError("ICLRAssemblyIdentityManager"))
    }
}

/// Implementation of the original `ICLRAssemblyIdentityManager` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl ICLRAssemblyIdentityManager {
    /// Retrieves the binding identity from a binary stream representing an assembly.
    ///
    /// # Arguments
    ///
    /// * `pstream` - Pointer to a `IStream` with assembly contents.
    /// * `dwFlags` - Control flags.
    /// * `pwzBuffer` - Buffer that receives the resulting identity string.
    /// * `pcchbuffersize` - Input/output buffer size.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the operation succeeds.
    /// * `Err(ClrError)` - On failure.
    pub fn GetBindingIdentityFromStream(
        &self,
        pstream: *mut c_void,
        dwFlags: u32,
        pwzBuffer: PWSTR,
        pcchbuffersize: *mut u32,
    ) -> Result<()> {
        let hr = unsafe {
            (Interface::vtable(self).GetBindingIdentityFromStream)(
                Interface::as_raw(self),
                pstream,
                dwFlags,
                pwzBuffer,
                pcchbuffersize,
            )
        };
        if hr == 0 {
            Ok(())
        } else {
            Err(ClrError::ApiError("GetBindingIdentityFromStream", hr))
        }
    }
}

unsafe impl Interface for ICLRAssemblyIdentityManager {
    type Vtable = ICLRAssemblyIdentityManager_Vtbl;

    /// The interface identifier (IID) for the `ICLRAssemblyIdentityManager` COM interface.
    ///
    /// This GUID is used to identify the `ICLRAssemblyIdentityManager` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `ICLRAssemblyIdentityManager` interface.
    const IID: GUID = GUID::from_u128(0x15f0a9da_3ff6_4393_9da9_fdfd284e6972);
}

impl Deref for ICLRAssemblyIdentityManager {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `ICLRAssemblyIdentityManager` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `ICLRAssemblyIdentityManager` interface.
#[repr(C)]
pub struct ICLRAssemblyIdentityManager_Vtbl {
    base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub GetCLRAssemblyReferenceList: *const c_void,
    pub GetBindingIdentityFromFile: *const c_void,
    pub GetBindingIdentityFromStream: unsafe extern "system" fn(
        this: *mut c_void,
        pstream: *mut c_void,
        dwFlags: u32,
        pwzBuffer: PWSTR,
        pcchbuffersize: *mut u32,
    ) -> HRESULT,
    pub GetReferencedAssembliesFromFile: *const c_void,
    pub GetReferencedAssembliesFromStream: *const c_void,
    pub GetProbingAssembliesFromReference: *const c_void,
    pub IsStronglyNamed: *const c_void,
}
