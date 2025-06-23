use core::{
    ffi::c_void, 
    ops::Deref, 
    ptr::{null, null_mut}
}; 
use {
    super::_Type, crate::Result,
    crate::error::ClrError, 
};
use alloc::string::String;
use windows_core::{IUnknown, Interface, GUID};
use windows_sys::{
    core::{BSTR, HRESULT}, 
    Win32::System::{
        Com::SAFEARRAY,
        Variant::{VariantClear, VARIANT}
    }
};

/// This struct represents the COM `_MethodInfo` interface, 
/// a .NET assembly in the CLR environment.
#[repr(C)]
#[derive(Debug, Clone)]
pub struct _MethodInfo(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `_MethodInfo` methods.
impl _MethodInfo {
    /// Invokes the method represented by this `_MethodInfo` instance.
    ///
    /// # Arguments
    /// 
    /// * `obj` - An optional `VARIANT` representing the target object for instance methods.
    /// * `parameters` - An optional pointer to a `SAFEARRAY` containing the parameters for the method.
    ///
    /// # Returns
    ///
    /// * `Ok(VARIANT)` - On successful invocation, returns the result as a `VARIANT`.
    /// * `Err(ClrError)` - Returns an error if the entry point cannot be resolved or invoked.
    pub fn invoke(&self, obj: Option<VARIANT>, parameters: Option<*mut SAFEARRAY>) -> Result<VARIANT> {
        let variant_obj = unsafe { obj.unwrap_or(core::mem::zeroed::<VARIANT>()) };
        self.Invoke_3(variant_obj, parameters.unwrap_or(null_mut()))
    }

    /// Creates an `_MethodInfo` instance from a raw COM interface pointer.
    ///
    /// # Arguments
    ///
    /// * `raw` - A raw pointer to an `IUnknown` COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(_MethodInfo)` - Wraps the given COM interface as `_MethodInfo`.
    /// * `Err(ClrError)` - If casting fails, returns a `ClrError`.
    #[inline(always)]
    pub fn from_raw(raw: *mut c_void) -> Result<_MethodInfo> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown.cast::<_MethodInfo>().map_err(|_| ClrError::CastingError("_MethodInfo"))
    }
}

/// Implementation of the original `_MethodInfo` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl _MethodInfo {
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

    /// Retrieves the name of the method.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The name of the method.
    /// * `Err(ClrError)` - Returns an error if the method name retrieval fails.
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
    ///
    /// # Arguments
    /// 
    /// * `obj` - A `VARIANT` representing the target instance or null for static methods.
    /// * `parameters` - A pointer to a `SAFEARRAY` containing the parameters for the method.
    ///
    /// # Returns
    ///
    /// * `Ok(VARIANT)` - The result of the method invocation.
    /// * `Err(ClrError)` - Returns an error if the invocation fails.
    pub fn Invoke_3(&self, obj: VARIANT, parameters: *mut SAFEARRAY) -> Result<VARIANT> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).Invoke_3)(Interface::as_raw(self), obj, parameters, &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                VariantClear(&mut result);
                Err(ClrError::ApiError("Invoke_3", hr))
            }
        }
    }

    /// Retrieves the parameters of the method as a `SAFEARRAY`.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut SAFEARRAY)` - A pointer to the `SAFEARRAY` containing the method's parameters.
    /// * `Err(ClrError)` - Returns an error if the parameters cannot be retrieved.
    pub fn GetParameters(&self) -> Result<*mut SAFEARRAY> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).GetParameters)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetParameters", hr))
        }
    }

    /// Calls the `GetHashCode` method from the vtable of the `_MethodInfo` interface.
    ///
    /// # Returns
    ///
    /// * `Ok(u32)` - Returns a 32-bit unsigned integer representing the hash code.
    /// * `Err(ClrError)` - If the call fails, returns a `ClrError`.
    pub fn GetHashCode(&self) -> Result<u32> {
        let mut result = 0;
        let hr = unsafe { (Interface::vtable(self).GetHashCode)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            Ok(result)
        } else {
            Err(ClrError::ApiError("GetHashCode", hr))
        }
    }

    /// Calls the `GetBaseDefinition` method from the vtable of the `_MethodInfo` interface.
    ///
    /// This method retrieves the base definition of the current method, 
    /// which represents the original declaration of the method in the inheritance chain.
    ///
    /// # Returns
    ///
    /// * `Ok(_MethodInfo)` - Returns an instance of `_MethodInfo`, representing the base method.
    /// * `Err(ClrError)` - Returns a `ClrError` if the call to `GetBaseDefinition` fails.
    pub fn GetBaseDefinition(&self) -> Result<_MethodInfo> {
        let mut result = null_mut();
        let hr = unsafe { (Interface::vtable(self).GetBaseDefinition)(Interface::as_raw(self), &mut result) };
        if hr == 0 {
            _MethodInfo::from_raw(result as *mut c_void)
        } else {
            Err(ClrError::ApiError("GetBaseDefinition", hr))
        }
    }

    /// Retrieves the main type associated with the method.
    ///
    /// # Returns
    ///
    /// * `Ok(_Type)` - On success, returns the `_Type` associated with the method.
    /// * `Err(ClrError)` - If retrieval fails, returns a `ClrError`.
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

#[repr(C)]
pub struct _MethodInfo_Vtbl {
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

    /// Retrieves the string representation of the Method.
    ///
    /// # Arguments
    /// 
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a `BSTR` that receives the string result.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    get_ToString: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut BSTR
    ) -> HRESULT,

    /// Placeholder for the method. Not used directly.
    Equals: *const c_void,

    /// Calculates the hash code for the method.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a `u32` that receives the hash code.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetHashCode: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut u32
    ) -> HRESULT,

    /// Retrieves the type information associated with the method.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to `_Type` where the type information is stored.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetType: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut *mut _Type
    ) -> HRESULT,

    /// Placeholder for the method. Not used directly.
    get_MemberType: *const c_void,

    /// Retrieves the name of the method as a `BSTR`.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to a `BSTR` that receives the method's name.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    get_name: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut BSTR
    ) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
    get_DeclaringType: *const c_void,
    get_ReflectedType: *const c_void,
    GetCustomAttributes: *const c_void,
    GetCustomAttributes_2: *const c_void,
    IsDefined: *const c_void,

    /// Retrieves the method parameters as a `SAFEARRAY`.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// - `pRetVal` - Pointer to a `SAFEARRAY` that receives the parameters.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetParameters: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut *mut SAFEARRAY
    ) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
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

    /// Invokes the method represented by this vtable entry.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `obj` - A `VARIANT` representing the target instance (or null for static methods).
    /// * `parameters` - A pointer to a `SAFEARRAY` of parameters.
    /// * `pRetVal` - Pointer to a `VARIANT` that will hold the result of the invocation.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    Invoke_3: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        parameters: *mut SAFEARRAY,
        pRetVal: *mut VARIANT
    ) -> HRESULT,

    /// Placeholder for the methods. Not used directly.
    get_returnType: *const c_void,
    get_ReturnTypeCustomAttributes: *const c_void,

    /// Retrieves the base definition of the method.
    ///
    /// # Arguments
    ///
    /// * `this` - Pointer to the COM object.
    /// * `pRetVal` - Pointer to `_MethodInfo` that will hold the base definition.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetBaseDefinition: unsafe extern "system" fn(
        this: *mut c_void,
        pRetVal: *mut *mut _MethodInfo
    ) -> HRESULT,
}