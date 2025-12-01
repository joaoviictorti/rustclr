use alloc::string::String;
use core::{
    ffi::c_void,
    ops::Deref,
    ptr::{null, null_mut},
};

use windows_core::{GUID, IUnknown, Interface};
use windows_sys::{
    core::{BSTR, HRESULT},
    Win32::System::{
        Com::SAFEARRAY,
        Variant::{VARIANT, VariantClear},
    },
};

use super::_Type;
use crate::error::{ClrError, Result};

/// This struct represents the COM `_MethodInfo` interface.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _MethodInfo(windows_core::IUnknown);

impl _MethodInfo {
    /// Invokes the method represented by this `_MethodInfo` instance.
    #[inline]
    pub fn invoke(
        &self,
        obj: Option<VARIANT>,
        parameters: Option<*mut SAFEARRAY>,
    ) -> Result<VARIANT> {
        let variant_obj = unsafe { obj.unwrap_or(core::mem::zeroed::<VARIANT>()) };
        self.Invoke_3(variant_obj, parameters.unwrap_or(null_mut()))
    }

    /// Creates an `_MethodInfo` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_MethodInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_MethodInfo>()
            .map_err(|_| ClrError::CastingError("_MethodInfo"))
    }

    /// Retrieves the string representation of the method (equivalent to `ToString` in .NET).
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

    /// Retrieves the name of the method.
    #[inline]
    pub fn get_name(&self) -> Result<String> {
        unsafe {
            let mut result = null::<u16>();
            let hr = (Interface::vtable(self).get_name)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }

                let slice = core::slice::from_raw_parts(result, len);
                Ok(String::from_utf16_lossy(slice))
            } else {
                Err(ClrError::ApiError("get_name", hr))
            }
        }
    }

    /// Internal invocation method for the method, used by `invoke`.
    #[inline]
    pub fn Invoke_3(&self, obj: VARIANT, parameters: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).Invoke_3)(
                Interface::as_raw(self),
                obj,
                parameters,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                VariantClear(&mut result);
                Err(ClrError::ApiError("Invoke_3", hr))
            }
        }
    }

    /// Retrieves the parameters of the method as a `SAFEARRAY`.
    #[inline]
    pub fn GetParameters(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetParameters)(Interface::as_raw(self), &mut result)
        };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetParameters", hr))
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_MethodInfo` interface.
    #[inline]
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr = unsafe { 
            (Interface::vtable(self).GetHashCode)(
                Interface::as_raw(self), 
                &mut result
            ) 
        };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetHashCode", hr))
        }
    }

    /// Calls the `GetBaseDefinition` method from the vtable of the `_MethodInfo` interface.
    #[inline]
    pub fn GetBaseDefinition(&self) -> Result<_MethodInfo> {
        let mut result = null_mut();
        let hr = unsafe {
            (Interface::vtable(self).GetBaseDefinition)(Interface::as_raw(self), &mut result)
        };
        if hr == 0 {
            _MethodInfo::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetBaseDefinition", hr))
        }
    }

    /// Retrieves the main type associated with the method.
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
}

unsafe impl Interface for _MethodInfo {
    type Vtable = _MethodInfo_Vtbl;

    /// The interface identifier (IID) for the `_MethodInfo` COM interface.
    ///
    /// This GUID is used to identify the `_MethodInfo` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `_MethodInfo` interface.
    const IID: GUID = GUID::from_u128(0xffcc1b5d_ecb8_38dd_9b01_3dc8abc2aa5f);
}

impl Deref for _MethodInfo {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_MethodInfo` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `_MethodInfo` interface.
#[repr(C)]
pub struct _MethodInfo_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

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
    get_MemberType: *const c_void,
    get_name: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
    get_DeclaringType: *const c_void,
    get_ReflectedType: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    GetParameters: unsafe extern "system" fn(
        this: *mut c_void, 
        pRetVal: *mut *mut SAFEARRAY
    ) -> HRESULT,
    GetMethodImplementationFlags: *const c_void,
    get_MethodHandle: *const c_void,
    get_Attributes: *const c_void,
    get_CallingConvention: *const c_void,
    Invoke_2: *const c_void,
    get_IsPublic: *const c_void,
    get_IsPrivate: *const c_void,
    get_IsFamily: *const c_void,
    get_IsAssembly: *const c_void,
    get_IsFamilyAndAssembly: *const c_void,
    get_IsFamilyOrAssembly: *const c_void,
    get_IsStatic: *const c_void,
    get_IsFinal: *const c_void,
    get_IsVirtual: *const c_void,
    get_IsHideBySig: *const c_void,
    get_IsAbstract: *const c_void,
    get_IsSpecialName: *const c_void,
    get_IsConstructor: *const c_void,
    Invoke_3: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        parameters: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    get_returnType: *const c_void,
    get_ReturnTypeCustomAttributes: *const c_void,
    GetBaseDefinition: unsafe extern "system" fn(
        this: *mut c_void, 
        pRetVal: *mut *mut _MethodInfo
    ) -> HRESULT,
}
