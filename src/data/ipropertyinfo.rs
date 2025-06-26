use alloc::{string::String, vec::Vec};
use core::{
    ffi::c_void,
    ops::Deref,
    ptr::{null, null_mut},
};

use windows_core::{GUID, IUnknown, Interface};
use windows_sys::{
    Win32::System::{Com::SAFEARRAY, Variant::VARIANT},
    core::{BSTR, HRESULT},
};

use crate::create_safe_args;
use crate::{Result, error::ClrError};

/// This struct represents the COM `_PropertyInfo` interface,
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _PropertyInfo(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `_PropertyInfo` methods.
impl _PropertyInfo {
    /// Retrieves the value of the property.
    ///
    /// # Arguments
    ///
    /// * `instance` - An optional `VARIANT` representing the target object on which
    ///   the property is being accessed. If `None`, a default zeroed variant is used.
    /// * `args` - An optional `Vec<VARIANT>` containing index arguments, such as for indexer properties.
    ///
    /// # Returns
    ///
    /// * `Ok(VARIANT)` - The value of the property if successfully retrieved.
    /// * `Err(ClrError)` - If the call to `GetValue` fails or argument conversion fails.
    pub fn value(&self, instance: Option<VARIANT>, args: Option<Vec<VARIANT>>) -> Result<VARIANT> {
        let args = args
            .as_ref()
            .map_or_else(|| Ok(null_mut()), |args| create_safe_args(args.to_vec()))?;

        let instance = instance.unwrap_or(unsafe { core::mem::zeroed::<VARIANT>() });
        self.GetValue(instance, args)
    }

    /// Creates an `_PropertyInfo` instance from a raw COM interface pointer.
    ///
    /// # Arguments
    ///
    /// * `raw` - A raw pointer to an `IUnknown` COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(_PropertyInfo)` - Wraps the given COM interface as `_PropertyInfo`.
    /// * `Err(ClrError)` - If casting fails, returns a `ClrError`.
    #[inline(always)]
    pub fn from_raw(raw: *mut c_void) -> Result<_PropertyInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_PropertyInfo>()
            .map_err(|_| ClrError::CastingError("_PropertyInfo"))
    }
}

/// Implementation of the original `_PropertyInfo` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl _PropertyInfo {
    /// Retrieves the string representation of the method (equivalent to `ToString` in .NET).
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The string representation of the method.
    /// * `Err(ClrError)` - Returns an error if the method retrieval fails.
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

    /// Retrieves a method by name.
    ///
    /// # Arguments
    ///
    /// * `name` - The name of the method as a `BSTR`.
    ///
    /// # Returns
    ///
    /// * `Ok(_MethodInfo)` - On success, returns the `_MethodInfo` for the method.
    /// * `Err(ClrError)` - On failure, returns a `ClrError`.
    pub fn GetValue(&self, instance: VARIANT, args: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).GetValue)(
                Interface::as_raw(self),
                instance,
                args,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetValue", hr))
            }
        }
    }
}

unsafe impl Interface for _PropertyInfo {
    type Vtable = _PropertyInfo_Vtbl;

    /// The interface identifier (IID) for the `_PropertyInfo` COM interface.
    ///
    /// This GUID is used to identify the `_PropertyInfo` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `_PropertyInfo` interface.
    const IID: GUID = GUID::from_u128(0xF59ED4E4_E68F_3218_BD77_061AA82824BF);
}

impl Deref for _PropertyInfo {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_PropertyInfo` to be used as an `IUnknown`
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`,
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct _PropertyInfo_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    ///
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Placeholder for the methods .Not used directly.
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,

    /// Retrieves the string representation of the Property.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a `BSTR` that receives the string result.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,

    /// Placeholder for the methods .Not used directly.
    Equals: *const c_void,
    GetHashCode: *const c_void,
    GetType: *const c_void,
    get_MemberType: *const c_void,
    get_name: *const c_void,
    get_DeclaringType: *const c_void,
    get_ReflectedType: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,
    get_PropertyType: *const c_void,

    /// Retrieves the value of the property for a given object instance and optional index parameters.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `obj` - The `VARIANT` representing the object whose property value is to be retrieved.
    /// * `index` - Pointer to a `SAFEARRAY` of `VARIANT`s representing optional index parameters (or `NULL`).
    /// * `pRetVal` - Pointer to a `VARIANT` that will receive the property value.
    ///
    /// # Returns
    ///
    /// * Returns an HRESULT indicating success or failure.
    GetValue: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        index: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,

    /// Placeholder for the methods .Not used directly.
    GetValue_2: *const c_void,
    SetValue: *const c_void,
    SetValue_2: *const c_void,
    GetAccessors: *const c_void,
    GetGetMethod: *const c_void,
    GetSetMethod: *const c_void,
    GetIndexParameters: *const c_void,
    get_Attributes: *const c_void,
    get_CanRead: *const c_void,
    get_CanWrite: *const c_void,
    GetAccessors_2: *const c_void,
    GetGetMethod_2: *const c_void,
    GetSetMethod_2: *const c_void,
    get_IsSpecialName: *const c_void,
}
