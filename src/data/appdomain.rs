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

#[repr(C)]
pub struct _AppDomainVtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    ///
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Placeholder for the methods. Not used directly.
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,
    get_ToString: *const c_void,
    Equals: *const c_void,

    /// Implementation of the `GetHashCode` method.
    ///
    /// This method returns the hash code of the current application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a variable that receives the hash code.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut u32) -> HRESULT,

    /// Implementation of the `GetType` method.
    ///
    /// This method retrieves the type of the current application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a variable that receives the `_Type` object.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    GetType: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _Type) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
    InitializeLifetimeService: *const c_void,
    GetLifetimeService: *const c_void,
    get_Evidence: *const c_void,
    add_DomainUnload: *const c_void,
    remove_DomainUnload: *const c_void,
    add_AssemblyLoad: *const c_void,
    remove_AssemblyLoad: *const c_void,
    add_ProcessExit: *const c_void,
    remove_ProcessExit: *const c_void,
    add_TypeResolve: *const c_void,
    remove_TypeResolve: *const c_void,
    add_ResourceResolve: *const c_void,
    remove_ResourceResolve: *const c_void,
    add_AssemblyResolve: *const c_void,
    remove_AssemblyResolve: *const c_void,
    add_UnhandledException: *const c_void,
    remove_UnhandledException: *const c_void,
    DefineDynamicAssembly: *const c_void,
    DefineDynamicAssembly_2: *const c_void,
    DefineDynamicAssembly_3: *const c_void,
    DefineDynamicAssembly_4: *const c_void,
    DefineDynamicAssembly_5: *const c_void,
    DefineDynamicAssembly_6: *const c_void,
    DefineDynamicAssembly_7: *const c_void,
    DefineDynamicAssembly_8: *const c_void,
    DefineDynamicAssembly_9: *const c_void,
    CreateInstance: *const c_void,
    CreateInstanceFrom: *const c_void,
    CreateInstance_2: *const c_void,
    CreateInstanceFrom_2: *const c_void,
    CreateInstance_3: *const c_void,
    CreateInstanceFrom_3: *const c_void,
    Load: *const c_void,

    /// Implementation of the `Load_2` method.
    ///
    /// This method loads an assembly into the current application domain by its name.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `assemblyString` - The name of the assembly to load, as a `BSTR`.
    /// * `pRetVal` - Pointer to a variable that receives the loaded `_Assembly`.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    Load_2: unsafe extern "system" fn(
        this: *mut c_void,
        assemblyString: BSTR,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,

    /// Implementation of the `Load_3` method.
    ///
    /// This method loads an assembly into the current application domain from raw byte data.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `rawAssembly` - Pointer to a `SAFEARRAY` containing the raw assembly data.
    /// * `pRetVal` - Pointer to a variable that receives the loaded `_Assembly`.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    Load_3: unsafe extern "system" fn(
        this: *mut c_void,
        rawAssembly: *mut SAFEARRAY,
        pRetVal: *mut *mut _Assembly,
    ) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
    Load_4: *const c_void,
    Load_5: *const c_void,
    Load_6: *const c_void,
    Load_7: *const c_void,
    ExecuteAssembly: *const c_void,
    ExecuteAssembly_2: *const c_void,
    ExecuteAssembly_3: *const c_void,
    get_FriendlyName: *const c_void,
    get_BaseDirectory: *const c_void,
    get_RelativeSearchPath: *const c_void,
    get_ShadowCopyFiles: *const c_void,

    /// Implementation of the `GetAssemblies` method.
    ///
    /// This method retrieves all assemblies currently loaded into the current application domain.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a variable that receives a `SAFEARRAY` of `_Assembly` references.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    GetAssemblies:unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
    AppendPrivatePath: *const c_void,
    ClearPrivatePath: *const c_void,
    SetShadowCopyPath: *const c_void,
    ClearShadowCopyPath: *const c_void,
    SetCachePath: *const c_void,
    SetData: *const c_void,
    GetData: *const c_void,
    SetAppDomainPolicy: *const c_void,
    SetThreadPrincipal: *const c_void,
    SetPrincipalPolicy: *const c_void,
    DoCallBack: *const c_void,
    get_DynamicDirectory: *const c_void,
}
