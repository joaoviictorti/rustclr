use alloc::{string::String, vec::Vec};
use core::{
    ffi::c_void,
    ops::Deref,
    ptr::{null, null_mut},
};

use obfstr::obfstr as s;
use windows_core::{GUID, IUnknown, Interface};
use windows_sys::{
    core::{BSTR, HRESULT},
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::SAFEARRAY,
            Variant::VARIANT,
            Ole::{
                SafeArrayGetElement, 
                SafeArrayGetLBound, 
                SafeArrayGetUBound
            },
        },
    },
};

use super::{_MethodInfo, _Type};
use crate::string::ComString;
use crate::error::{ClrError, Result};

/// This struct represents the COM `_Assembly` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _Assembly(windows_core::IUnknown);

impl _Assembly {
    /// Resolves a type by name within the assembly.
    #[inline]
    pub fn resolve_type(&self, name: &str) -> Result<_Type> {
        let type_name = name.to_bstr();
        self.GetType_2(type_name)
    }

    /// Executes the entry point of the assembly.
    ///
    /// The `run` method identifies the main entry point of the assembly and attempts
    /// to invoke it. It distinguishes between `Main()` and `Main(System.String[])` entry points,
    /// allowing optional arguments to be passed when the latter is detected.
    #[inline]
    pub fn run(&self, args: *mut SAFEARRAY) -> Result<VARIANT> {
        let entrypoint = self.get_EntryPoint()?;
        let str = entrypoint.ToString()?;
        match str.as_str() {
            str if str.ends_with(s!("Main()")) => entrypoint.invoke(None, None),
            str if str.ends_with(s!("Main(System.String[])")) => entrypoint.invoke(None, Some(args)),
            _ => Err(ClrError::MethodNotFound),
        }
    }

    /// Creates an instance of a type within the assembly.
    #[inline]
    pub fn create_instance(&self, name: &str) -> Result<VARIANT> {
        let type_name = name.to_bstr();
        self.CreateInstance(type_name)
    }

    /// Retrieves all types within the assembly.
    #[inline]
    pub fn types(&self) -> Result<Vec<String>> {
        let sa_types = self.GetTypes()?;
        if sa_types.is_null() {
            return Err(ClrError::NullPointerError("GetTypes"));
        }

        let mut types = Vec::new();
        let mut lbound = 0;
        let mut ubound = 0;
        unsafe {
            SafeArrayGetLBound(sa_types, 1, &mut lbound);
            SafeArrayGetUBound(sa_types, 1, &mut ubound);

            for i in lbound..=ubound {
                let mut p_type = null_mut::<_Type>();
                let hr = SafeArrayGetElement(sa_types, &i, &mut p_type as *mut _ as *mut _);
                if hr != 0 || p_type.is_null() {
                    return Err(ClrError::ApiError("SafeArrayGetElement", hr));
                }

                let _type = _Type::from_raw(p_type as *mut c_void)?;
                let type_name = _type.ToString()?;
                types.push(type_name);
            }
        }

        Ok(types)
    }

    /// Creates an `_Assembly` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_Assembly> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_Assembly>()
            .map_err(|_| ClrError::CastingError("_Assembly"))
    }

    /// Retrieves the string representation of the assembly.
    #[inline]
    pub fn ToString(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("ToString", hr))
            }
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_Assembly` interface.
    #[inline]
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr =
            unsafe { (Interface::vtable(self).GetHashCode)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetHashCode", hr))
        }
    }

    /// Retrieves the entry point method of the assembly.
    #[inline]
    pub fn get_EntryPoint(&self) -> Result<_MethodInfo> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).get_EntryPoint)(Interface::as_raw(self), &mut result)
        };
        if hr == 0 {
            _MethodInfo::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("get_EntryPoint", hr))
        }
    }

    /// Resolves a specific type by name within the assembly.
    #[inline]
    pub fn GetType_2(&self, name: BSTR) -> Result<_Type> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetType_2)(Interface::as_raw(self), name, &mut result)
        };
        if hr == 0 {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetType_2", hr))
        }
    }

    /// Retrieves all types defined within the assembly as a `SAFEARRAY`.
    #[inline]
    pub fn GetTypes(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr =
            unsafe { (Interface::vtable(self).GetTypes)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetTypes", hr))
        }
    }

    /// Creates an instance of a type using its name as a `BSTR`.
    #[inline]
    pub fn CreateInstance(&self, typeName: BSTR) -> Result<VARIANT> {
        let mut result = unsafe { core::mem::zeroed::<VARIANT>() };
        let hr = unsafe {
            (Interface::vtable(self).CreateInstance)(Interface::as_raw(self), typeName, &mut result)
        };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("CreateInstance", hr))
        }
    }

    /// Retrieves the main type associated with the assembly.
    #[inline]
    pub fn GetType(&self) -> Result<_Type> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).GetType)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            _Type::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetType", hr))
        }
    }

    /// Retrieves the assembly's codebase as a URI.
    #[inline]
    pub fn get_CodeBase(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_CodeBase)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("get_CodeBase", hr))
            }
        }
    }

    /// Retrieves the escaped codebase of the assembly as a URI.
    #[inline]
    pub fn get_EscapedCodeBase(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr =
                (Interface::vtable(self).get_EscapedCodeBase)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("get_EscapedCodeBase", hr))
            }
        }
    }

    /// Retrieves the name of the assembly.
    #[inline]
    pub fn GetName(&self) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetName)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetName", hr))
            }
        }
    }

    /// Retrieves the name of the assembly, with an option to copy the name.
    #[inline]
    pub fn GetName_2(&self, copiedName: VARIANT_BOOL) -> Result<*mut c_void> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetName_2)(
                Interface::as_raw(self),
                copiedName,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetName_2", hr))
            }
        }
    }

    /// Retrieves the full name of the assembly.
    #[inline]
    pub fn get_FullName(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_FullName)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("get_FullName", hr))
            }
        }
    }

    /// Retrieves the file location of the assembly.
    #[inline]
    pub fn get_Location(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_Location)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("get_Location", hr))
            }
        }
    }
}

unsafe impl Interface for _Assembly {
    type Vtable = _Assembly_Vtbl;

    /// The interface identifier (IID) for the `_Assembly` COM interface.
    ///
    /// This GUID is used to identify the `_Assembly` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `_Assembly` interface.
    const IID: GUID = GUID::from_u128(0x17156360_2f1a_384a_bc52_fde93c215c5b);
}

impl Deref for _Assembly {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_Assembly` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `_Assembly` interface.
#[repr(C)]
pub struct _Assembly_Vtbl {
    base__: windows_core::IUnknown_Vtbl,

    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    Equals: *const c_void,
    GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut u32) -> HRESULT,
    GetType: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut _Type) -> HRESULT,
    get_CodeBase: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    get_EscapedCodeBase: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    GetName: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut c_void) -> HRESULT,
    GetName_2: unsafe extern "system" fn(
        this: *mut c_void,
        copiedName: VARIANT_BOOL,
        pRetVal: *mut *mut c_void,
    ) -> HRESULT,
    get_FullName: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    get_EntryPoint: unsafe extern "system" fn(
        this: *mut c_void, 
        pRetVal: *mut *mut _MethodInfo
    ) -> HRESULT,
    GetType_2: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR,
        pRetVal: *mut *mut _Type,
    ) -> HRESULT,
    GetType_3: *const c_void,
    GetExportedTypes: *const c_void,
    GetTypes: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    GetManifestResourceStream: *const c_void,
    GetManifestResourceStream_2: *const c_void,
    GetFile: *const c_void,
    GetFiles: *const c_void,
    GetFiles_2: *const c_void,
    GetManifestResourceNames: *const c_void,
    GetManifestResourceInfo: *const c_void,
    get_Location: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    get_Evidence: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    GetObjectData: *const c_void,
    add_ModuleResolve: *const c_void,
    remove_ModuleResolve: *const c_void,
    GetType_4: *const c_void,
    GetSatelliteAssembly: *const c_void,
    GetSatelliteAssembly_2: *const c_void,
    LoadModule: *const c_void,
    LoadModule_2: *const c_void,
    CreateInstance: unsafe extern "system" fn(
        this: *mut c_void,
        typeName: BSTR,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    CreateInstance_2: *const c_void,
    CreateInstance_3: *const c_void,
    GetLoadedModules: *const c_void,
    GetLoadedModules_2: *const c_void,
    GetModules: *const c_void,
    GetModules_2: *const c_void,
    GetModule: *const c_void,
    GetReferencedAssemblies: *const c_void,
    get_GlobalAssemblyCache: *const c_void,
}
