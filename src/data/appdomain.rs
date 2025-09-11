use alloc::{string::String, vec::Vec};
use core::{ffi::c_void, ops::Deref, ptr::null_mut};

use windows_core::{GUID, IUnknown, Interface};
use windows_sys::{
    core::{BSTR, HRESULT},
    Win32::System::{
        Com::SAFEARRAY,
        Ole::{
            SafeArrayGetElement, 
            SafeArrayGetLBound, 
            SafeArrayGetUBound
        },
    },
};

use super::{_Assembly, _Type};
use crate::{
    Result, ComString, 
    create_safe_array_buffer, 
    error::ClrError
};

pub type PVOID = *const c_void;

/// This struct represents the COM `_AppDomain` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _AppDomain(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `_AppDomain` methods.
impl _AppDomain {
    /// Loads an assembly into the current application domain from a byte slice.
    ///
    /// This method creates a `SAFEARRAY` from the given byte buffer and loads it using
    /// the `Load_3` method.
    ///
    /// # Arguments
    ///
    /// * `buffer` - A slice of bytes representing the raw assembly data.
    ///
    /// # Returns
    ///
    /// * `Ok(_Assembly)` - If successful, returns an `_Assembly` instance.
    /// * `Err(ClrError)` - If loading fails, returns a `ClrError`.
    pub fn load_bytes(&self, buffer: &[u8]) -> Result<_Assembly> {
        let safe_array = create_safe_array_buffer(buffer)?;
        self.Load_3(safe_array)
    }

    /// Loads an assembly by its name in the current application domain.
    ///
    /// This method converts the assembly name to a `BSTR` and uses the `Load_2` method.
    ///
    /// # Arguments
    ///
    /// * `name` - The name of the assembly as a string slice.
    ///
    /// # Returns
    ///
    /// * `Ok(_Assembly)` - If successful, returns an `_Assembly` instance.
    /// * `Err(ClrError)` - If loading fails, returns a `ClrError`.
    pub fn load_name(&self, name: &str) -> Result<_Assembly> {
        let lib_name = name.to_bstr();
        self.Load_2(lib_name)
    }

    /// Creates an `_AppDomain` instance from a raw COM interface pointer.
    ///
    /// # Arguments
    ///
    /// * `raw` - A raw pointer to an `IUnknown` COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - Wraps the given COM interface as `_AppDomain`.
    /// * `Err(ClrError)` - If casting fails, returns a `ClrError`.
    #[inline(always)]
    pub fn from_raw(raw: *mut c_void) -> Result<_AppDomain> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_AppDomain>()
            .map_err(|_| ClrError::CastingError("_AppDomain"))
    }

    /// Searches for an assembly by name within the current AppDomain.
    ///
    /// # Arguments
    ///
    /// * `assembly_name` – A substring to look for in the assembly's full display name.
    ///
    /// # Returns
    ///
    /// * `Ok(_Assembly)` – If an assembly is found matching the name.
    /// * `Err(ClrError)` – If no matching assembly is found.
    pub fn get_assembly(&self, assembly_name: &str) -> Result<_Assembly> {
        let assemblies = self.assemblies()?;
        for (name, assembly) in assemblies {
            if name.contains(assembly_name) {
                return Ok(assembly);
            }
        }

        Err(ClrError::GenericError("Assembly Not Found"))
    }

    /// Retrieves all assemblies currently loaded in the AppDomain.
    ///
    /// # Returns
    ///
    /// * `Ok(Vec<(String, _Assembly)>)` – A list of loaded assemblies and their display names.
    /// * `Err(ClrError)` – If any error occurs during retrieval or conversion.
    pub fn assemblies(&self) -> Result<Vec<(String, _Assembly)>> {
        let sa_assemblies = self.GetAssemblies()?;
        if sa_assemblies.is_null() {
            return Err(ClrError::NullPointerError("GetAssemblies"));
        }

        let mut assemblies = Vec::new();
        let mut lbound = 0;
        let mut ubound = 0;
        unsafe {
            SafeArrayGetLBound(sa_assemblies, 1, &mut lbound);
            SafeArrayGetUBound(sa_assemblies, 1, &mut ubound);

            for i in lbound..=ubound {
                let mut p_assembly = null_mut::<_Assembly>();
                let hr =
                    SafeArrayGetElement(sa_assemblies, &i, &mut p_assembly as *mut _ as *mut _);
                if hr != 0 || p_assembly.is_null() {
                    return Err(ClrError::ApiError("SafeArrayGetElement", hr));
                }

                let _assembly = _Assembly::from_raw(p_assembly as *mut c_void)?;
                let assembly_name = _assembly.ToString()?;
                assemblies.push((assembly_name, _assembly));
            }
        }

        Ok(assemblies)
    }
}

/// Implementation of the original `_AppDomain` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl _AppDomain {
    /// Calls the `Load_3` method from the vtable of the `_AppDomain` interface.
    ///
    /// # Arguments
    ///
    /// * `rawAssembly` - The raw assembly data as a `SAFEARRAY` pointer.
    ///
    /// # Returns
    ///
    /// * `Ok(_Assembly)` - If successful, returns a `_Assembly` instance.
    /// * `Err(ClrError)` - If loading fails, returns a `ClrError`.
    pub fn Load_3(&self, rawAssembly: *mut SAFEARRAY) -> Result<_Assembly> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).Load_3)(Interface::as_raw(self), rawAssembly, &mut result)
        };
        if hr == 0 {
            _Assembly::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("Load_3", hr))
        }
    }

    /// Calls the `Load_2` method from the vtable of the `_AppDomain` interface.
    ///
    /// # Arguments
    ///
    /// * `rawAssembly` - The raw assembly data as a `SAFEARRAY` pointer.
    ///
    /// # Returns
    ///
    /// * `Ok(_Assembly)` - If successful, returns a `_Assembly` instance.
    /// * `Err(ClrError)` - If loading fails, returns a `ClrError`.
    pub fn Load_2(&self, assemblyString: BSTR) -> Result<_Assembly> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).Load_2)(Interface::as_raw(self), assemblyString, &mut result)
        };
        if hr == 0 {
            _Assembly::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("Load_2", hr))
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_AppDomain` interface.
    ///
    /// # Returns
    ///
    /// * `Ok(u32)` - Returns a 32-bit unsigned integer representing the hash code.
    /// * `Err(ClrError)` - If the call fails, returns a `ClrError`.
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr = unsafe { 
            (Interface::vtable(self).GetHashCode)(Interface::as_raw(self), &mut result) 
        };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetHashCode", hr))
        }
    }

    /// Retrieves the primary type associated with the current app domain.
    ///
    /// # Returns
    ///
    /// * `Ok(_Type)` - On success, returns the `_Type` associated with the app domain.
    /// * `Err(ClrError)` - If the type cannot be retrieved, returns a `ClrError`.
    pub fn GetType(&self) -> Result<_Type> {
        let mut result = null_mut();
        let hr  = unsafe { 
            (Interface::vtable(self).GetType)(Interface::as_raw(self), &mut result) 
        };
        if hr == 0 {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetType", hr))
        }
    }

    /// Retrieves the assemblies currently loaded into the current AppDomain.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut SAFEARRAY)` – Pointer to a COM SAFEARRAY of `_Assembly` references.
    /// * `Err(ClrError)` – If the COM call fails or returns an error HRESULT.
    pub fn GetAssemblies(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr: i32 = unsafe {
            (Interface::vtable(self).GetAssemblies)(Interface::as_raw(self), &mut result)
        };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetAssemblies", hr))
        }
    }
}

unsafe impl Interface for _AppDomain {
    type Vtable = _AppDomainVtbl;

    /// The interface identifier (IID) for the `_AppDomain` COM interface.
    ///
    /// This GUID is used to identify the `_AppDomain` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `_AppDomain` interface.
    const IID: GUID = GUID::from_u128(0x05F696DC_2B29_3663_AD8B_C4389CF2A713);
}

impl Deref for _AppDomain {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_AppDomain` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `_AppDomain` interface.
#[repr(C)]
pub struct _AppDomainVtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // IDispatch methods
    GetTypeInfoCount: PVOID,
    GetTypeInfo: PVOID,
    GetIDsOfNames: PVOID,
    Invoke: PVOID,

    // Methods specific to the COM interface
    get_ToString: PVOID,
    Equals: PVOID,
    GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut u32) -> HRESULT,
    GetType: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _Type) -> HRESULT,
    InitializeLifetimeService: PVOID,
    GetLifetimeService: PVOID,
    get_Evidence: PVOID,
    add_DomainUnload: PVOID,
    remove_DomainUnload: PVOID,
    add_AssemblyLoad: PVOID,
    remove_AssemblyLoad: PVOID,
    add_ProcessExit: PVOID,
    remove_ProcessExit: PVOID,
    add_TypeResolve: PVOID,
    remove_TypeResolve: PVOID,
    add_ResourceResolve: PVOID,
    remove_ResourceResolve: PVOID,
    add_AssemblyResolve: PVOID,
    remove_AssemblyResolve: PVOID,
    add_UnhandledException: PVOID,
    remove_UnhandledException: PVOID,
    DefineDynamicAssembly: PVOID,
    DefineDynamicAssembly_2: PVOID,
    DefineDynamicAssembly_3: PVOID,
    DefineDynamicAssembly_4: PVOID,
    DefineDynamicAssembly_5: PVOID,
    DefineDynamicAssembly_6: PVOID,
    DefineDynamicAssembly_7: PVOID,
    DefineDynamicAssembly_8: PVOID,
    DefineDynamicAssembly_9: PVOID,
    CreateInstance: PVOID,
    CreateInstanceFrom: PVOID,
    CreateInstance_2: PVOID,
    CreateInstanceFrom_2: PVOID,
    CreateInstance_3: PVOID,
    CreateInstanceFrom_3: PVOID,
    Load: PVOID,
    Load_2: unsafe extern "system" fn(
        this: *mut c_void,
        assemblyString: BSTR,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,
    Load_3: unsafe extern "system" fn(
        this: *mut c_void,
        rawAssembly: *mut SAFEARRAY,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,
    Load_4: PVOID,
    Load_5: PVOID,
    Load_6: PVOID,
    Load_7: PVOID,
    ExecuteAssembly: PVOID,
    ExecuteAssembly_2: PVOID,
    ExecuteAssembly_3: PVOID,
    get_FriendlyName: PVOID,
    get_BaseDirectory: PVOID,
    get_RelativeSearchPath: PVOID,
    get_ShadowCopyFiles: PVOID,
    GetAssemblies: unsafe extern "system" fn(
        this: *mut c_void, 
        pRetVal: *mut *mut SAFEARRAY
    ) -> HRESULT,
    AppendPrivatePath: PVOID,
    ClearPrivatePath: PVOID,
    SetShadowCopyPath: PVOID,
    ClearShadowCopyPath: PVOID,
    SetCachePath: PVOID,
    SetData: PVOID,
    GetData: PVOID,
    SetAppDomainPolicy: PVOID,
    SetThreadPrincipal: PVOID,
    SetPrincipalPolicy: PVOID,
    DoCallBack: PVOID,
    get_DynamicDirectory: PVOID,
}
