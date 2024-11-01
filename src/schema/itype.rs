use {
    windows_core::{IUnknown, Interface, GUID},
    std::{
        ffi::c_void, ptr::{null_mut, null}, 
        ops::{BitOr, Deref}
    },
    windows_sys::{
        core::{BSTR, HRESULT}, 
        Win32::System::{
            Com::SAFEARRAY, 
            Variant::VARIANT,
            Ole::{
                SafeArrayGetElement, SafeArrayGetLBound, 
                SafeArrayGetUBound
            }, 
        }
    }
};

use crate::{
    error::ClrError, schema::_MethodInfo,
    WinStr, create_safe_args, InvocationType,
};

/// Represents the `_Type` COM interface, allowing for the invocation of
/// methods, retrieval of properties, and interaction with .NET types
/// within the CLR environment.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct _Type(windows_core::IUnknown);

/// Implementation of auxiliary methods for convenience.
///
/// These methods provide Rust-friendly wrappers around the original `_Type` methods.
impl _Type {
    /// Retrieves a method by its name from the type.
    ///
    /// # Arguments
    ///
    /// * `name` - A string slice representing the method name.
    ///
    /// # Returns
    ///
    /// * `Ok(_MethodInfo)` - On success, returns the method's `_MethodInfo`.
    /// * `Err(ClrError)` - On failure, returns a `ClrError`.
    pub fn method(&self, name: &str) -> Result<_MethodInfo, ClrError> {
        let method_name = name.to_bstr();
        self.GetMethod_6(method_name)
    }

    /// Finds a method by signature from the type.
    ///
    /// # Arguments
    ///
    /// * `name` - A string slice representing the method signature.
    ///
    /// # Returns
    ///
    /// * `Ok(_MethodInfo)` - On success, returns the matching `_MethodInfo`.
    /// * `Err(ClrError)` - On failure, returns `ClrError::MethodNotFound`.
    pub fn method_signature(&self, name: &str) -> Result<_MethodInfo, ClrError> {
        let methods = self.methods();
        if let Ok(methods) = methods {
            for (method_name, method_info) in methods {
                if method_name == name {
                    return Ok(method_info);
                }
            }
        }

        Err(ClrError::MethodNotFound)
    }

    /// Invokes a method on the type.
    ///
    /// # Arguments
    ///
    /// * `name` - The name of the method to invoke.
    /// * `instance` - An optional `VARIANT` representing the instance.
    /// * `args` - Optional vector of `VARIANT` arguments.
    /// * `invocation_type` - The `InvocationType`, indicating if it's a static or instance method.
    ///
    /// # Returns
    ///
    /// * `Ok(VARIANT)` - On success, returns the result as `VARIANT`.
    /// * `Err(ClrError)` - On failure, returns `ClrError`.
    pub fn invoke(
        &self, 
        name: &str, 
        instance: Option<VARIANT>, 
        args: Option<Vec<VARIANT>>, 
        invocation_type: InvocationType
    ) -> Result<VARIANT, ClrError> {
        let flags = match invocation_type {
            InvocationType::Static => BindingFlags::Public | BindingFlags::Static | BindingFlags::InvokeMethod,
            InvocationType::Instance => BindingFlags::Public | BindingFlags::Instance | BindingFlags::InvokeMethod,
        };

        let method_name = name.to_bstr();
        let args = args.as_ref().map_or_else(
            || Ok(null_mut()),
            |args| create_safe_args(args.to_vec())
        )?;

        let instance = instance.unwrap_or(unsafe { std::mem::zeroed::<VARIANT>() });
        self.InvokeMember_3(method_name, flags, instance, args)
    }

    /// Retrieves all methods of the type.
    ///
    /// # Returns
    ///
    /// * `Ok(Vec<(String, _MethodInfo)>)` - On success, returns a vector of method names and `_MethodInfo`.
    /// * `Err(ClrError)` - On failure, returns a `ClrError`.
    pub fn methods(&self) -> Result<Vec<(String, _MethodInfo)>, ClrError> {
        let binding_flags = BindingFlags::Public | BindingFlags::Instance | 
            BindingFlags::Static | BindingFlags::FlattenHierarchy;
        
        let sa_methods = self.GetMethods(binding_flags)?;
        if sa_methods.is_null() {
            return Err(ClrError::NullPointerError("GetMethods"));
        }

        let mut lbound = 0;
        let mut ubound = 0;
        let mut methods = Vec::new();
        unsafe {
            SafeArrayGetLBound(sa_methods, 1, &mut lbound);
            SafeArrayGetUBound(sa_methods, 1, &mut ubound);
            
            let mut p_method = null_mut::<_MethodInfo>();
            for i in lbound..=ubound {
                let hr = SafeArrayGetElement(sa_methods, &i, &mut p_method as *mut _ as *mut _);
                if hr != 0 || p_method.is_null() {
                    return Err(ClrError::ApiError("SafeArrayGetElement", hr));
                }

                let method = _MethodInfo::from_raw(p_method as *mut c_void)?;
                let method_name = method.ToString()?;
                methods.push((method_name, method));
            }
        }

        Ok(methods)
    }

    /// Creates an `_Type` instance from a raw COM interface pointer.
    ///
    /// # Arguments
    ///
    /// * `raw` - A raw pointer to an `IUnknown` COM interface.
    ///
    /// # Returns
    ///
    /// * `Ok(_Type)` - On success, returns the `_Type` wrapping the COM interface.
    /// * `Err(ClrError)` - If creation fails, returns a `ClrError`.
    #[inline(always)]
    pub fn from_raw(raw: *mut c_void) -> Result<_Type, ClrError> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown.cast::<_Type>().map_err(|_| ClrError::CastingError("_Type"))
    }
}

/// Implementation of the original `_Type` COM interface methods.
///
/// These methods are direct FFI bindings to the corresponding functions in the COM interface.
impl _Type {
    /// Retrieves the string representation of the type.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - On success, returns the type's name as a `String`.
    /// * `Err(ClrError)` - On failure, returns a `ClrError`.
    pub fn ToString(&self) -> Result<String, ClrError> {
        unsafe {
            let mut result= null::<u16>();
            let hr = (Interface::vtable(self).get_ToString)(Interface::as_raw(self), &mut result);
            if hr == 0 {
                let mut len = 0;
                while *result.add(len) != 0 {
                    len += 1;
                }
    
                let slice = std::slice::from_raw_parts(result, len);
                let entrypoint = String::from_utf16_lossy(slice);

                Ok(entrypoint)
            } else {
                Err(ClrError::ApiError("ToString", hr))
            }
        }
    }

    /// Retrieves all methods matching the specified `BindingFlags`.
    ///
    /// # Arguments
    ///
    /// * `bindingAttr` - The `BindingFlags` specifying which methods to retrieve.
    ///
    /// # Returns
    ///
    /// * `Ok(*mut SAFEARRAY)` - On success, returns a pointer to a `SAFEARRAY` of methods.
    /// * `Err(ClrError)` - On failure, returns a `ClrError`.
    pub fn GetMethods(&self, bindingAttr: BindingFlags) -> Result<*mut SAFEARRAY, ClrError> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetMethods)(Interface::as_raw(self), bindingAttr, &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetMethods", hr))
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
    pub fn GetMethod_6(&self, name: BSTR) -> Result<_MethodInfo, ClrError> {
        unsafe {
            let mut result = std::mem::zeroed();
            let hr = (Interface::vtable(self).GetMethod_6)(Interface::as_raw(self), name, &mut result);
            if hr == 0 {
                _MethodInfo::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("GetMethod_6", hr))
            }
        }
    }

    /// Invokes a method (static or instance) by name on the specified type or object.
    ///
    /// # Arguments
    ///
    /// * `name` - The name of the member to invoke, provided as a `BSTR`.
    /// * `invoke_attr` - `BindingFlags` that specify invocation options (such as
    ///   whether to target a static or instance method).
    /// * `instance` - A `VARIANT` representing the object instance on which to invoke
    ///   the member, or a `null`/default value for static members.
    /// * `args` - A pointer to a `SAFEARRAY` containing the arguments for the method invocation.
    ///
    /// # Returns
    ///
    /// * `Ok(VARIANT)` - On success, returns the result of the invocation as a `VARIANT`.
    /// * `Err(ClrError)` - If invocation fails, returns an appropriate `ClrError`.
    pub fn InvokeMember_3(&self, name: BSTR, invoke_attr: BindingFlags, instance: VARIANT, args: *mut SAFEARRAY) -> Result<VARIANT, ClrError> {
        unsafe {
            let mut result = std::mem::zeroed();
            let hr = (Interface::vtable(self).InvokeMember_3)(Interface::as_raw(self), name, invoke_attr, null_mut(), instance, args, &mut result);
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("InvokeMember_3", hr))
            }
        }
    }
}

unsafe impl Interface for _Type {
    type Vtable = _Type_Vtbl;

    /// The interface identifier (IID) for the `_Type` COM interface.
    ///
    /// This GUID is used to identify the `_Type` interface when calling 
    /// COM methods like `QueryInterface`. It is defined based on the standard 
    /// .NET CLR IID for the `_Type` interface.
    const IID: GUID = GUID::from_u128(0xbca8b44d_aad6_3a86_8ab7_03349f4f2da2);
}

impl Deref for _Type {
    type Target = windows_core::IUnknown;

    /// Provides a reference to the underlying `IUnknown` interface.
    ///
    /// This implementation allows `_Type` to be used as an `_Type` 
    /// pointer, enabling access to basic COM methods like `AddRef`, `Release`, 
    /// and `QueryInterface`.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

#[repr(C)]
pub struct _Type_Vtbl {
    /// Base vtable inherited from the `IUnknown` interface.
    /// 
    /// This field contains the basic methods for reference management,
    /// like `AddRef`, `Release`, and `QueryInterface`.
    pub base__: windows_core::IUnknown_Vtbl,

    /// Placeholder for the `GetTypeInfoCount` method. Not used directly.
    GetTypeInfoCount: *const c_void,

    /// Placeholder for the `GetTypeInfo` method. Not used directly.
    GetTypeInfo: *const c_void,

    /// Placeholder for the `GetIDsOfNames` method. Not used directly.
    GetIDsOfNames: *const c_void,

    /// Placeholder for the `Invoke` method. Not used directly.
    Invoke: *const c_void,

    /// Retrieves the string representation of the Method.
    ///
    /// # Arguments
    /// 
    /// * `*mut c_void` - Pointer to the COM object implementing the interface.
    /// * `pRetVal` - Pointer to a `BSTR` that receives the string result.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    get_ToString: unsafe extern "system" fn(
        *mut c_void,
        pRetVal: *mut BSTR
    ) -> HRESULT,

    /// Placeholder for the `Equals` method. Not used directly.
    Equals: *const c_void,

    /// Placeholder for the `GetHashCode` method. Not used directly.
    GetHashCode: *const c_void,
    
    /// Placeholder for the `GetType` method. Not used directly.
    GetType: *const c_void,

    /// Placeholder for the `get_MemberType` method. Not used directly.
    get_MemberType: *const c_void,

    /// Placeholder for the `get_name` method. Not used directly.
    get_name: *const c_void,

    /// Placeholder for the `get_DeclaringType` method. Not used directly.
    get_DeclaringType: *const c_void,

    /// Placeholder for the `get_ReflectedType` method. Not used directly.
    get_ReflectedType: *const c_void,

    /// Placeholder for the `GetCustomAttributes` method. Not used directly.
    GetCustomAttributes: *const c_void,

    /// Placeholder for the `GetCustomAttributes_2` method. Not used directly.
    GetCustomAttributes_2: *const c_void,

    /// Placeholder for the `IsDefined` method. Not used directly.
    IsDefined: *const c_void,

    /// Placeholder for the `get_Guid` method. Not used directly.
    get_Guid: *const c_void,

    /// Placeholder for the `get_Module` method. Not used directly.
    get_Module: *const c_void,

    /// Placeholder for the `get_Assembly` method. Not used directly.
    get_Assembly: *const c_void,

    /// Placeholder for the `get_TypeHandle` method. Not used directly.
    get_TypeHandle: *const c_void,

    /// Placeholder for the `get_FullName` method. Not used directly.
    get_FullName: *const c_void,

    /// Placeholder for the `get_Namespace` method. Not used directly.
    get_Namespace: *const c_void,

    /// Placeholder for the `get_AssemblyQualifiedName` method. Not used directly.
    get_AssemblyQualifiedName: *const c_void,

    /// Placeholder for the `GetArrayRank` method. Not used directly.
    GetArrayRank: *const c_void,

    /// Placeholder for the `get_BaseType` method. Not used directly.
    get_BaseType: *const c_void,

    /// Placeholder for the `GetConstructors` method. Not used directly.
    GetConstructors: *const c_void,
    
    /// Placeholder for the `GetInterface` method. Not used directly.
    GetInterface: *const c_void,

    /// Placeholder for the `GetInterfaces` method. Not used directly.
    GetInterfaces: *const c_void,

    /// Placeholder for the `FindInterfaces` method. Not used directly.
    FindInterfaces: *const c_void,

    /// Placeholder for the `GetEvent` method. Not used directly.
    GetEvent: *const c_void,

    /// Placeholder for the `GetEvents` method. Not used directly.
    GetEvents: *const c_void,

    /// Placeholder for the `GetEvents_2` method. Not used directly.
    GetEvents_2: *const c_void,

    /// Placeholder for the `GetNestedTypes` method. Not used directly.
    GetNestedTypes: *const c_void,

    /// Placeholder for the `GetNestedType` method. Not used directly.
    GetNestedType: *const c_void,

    /// Placeholder for the `GetMember` method. Not used directly.
    GetMember: *const c_void,

    /// Placeholder for the `GetDefaultMembers` method. Not used directly.
    GetDefaultMembers: *const c_void,

    /// Placeholder for the `FindMembers` method. Not used directly.
    FindMembers: *const c_void,

    /// Placeholder for the `GetElementType` method. Not used directly.
    GetElementType: *const c_void,

    /// Placeholder for the `IsSubclassOf` method. Not used directly.
    IsSubclassOf: *const c_void,

    /// Placeholder for the `IsInstanceOfType` method. Not used directly.
    IsInstanceOfType: *const c_void,

    /// Placeholder for the `IsAssignableFrom` method. Not used directly.
    IsAssignableFrom: *const c_void,

    /// Placeholder for the `GetInterfaceMap` method. Not used directly.
    GetInterfaceMap: *const c_void,

    /// Placeholder for the `GetMethod` method. Not used directly.
    GetMethod: *const c_void,

    /// Placeholder for the `GetMethod_2` method. Not used directly.
    GetMethod_2: *const c_void,

    /// Retrieves methods matching the specified `BindingFlags`.
    ///
    /// # Arguments
    /// 
    /// * `*mut c_void` - Pointer to the COM object implementing the interface.
    /// * `bindingAttr` - The `BindingFlags` specifying the methods to retrieve.
    /// * `pRetVal` - A pointer to a `SAFEARRAY` that receives the retrieved methods.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetMethods: unsafe extern "system" fn(
        *mut c_void, 
        bindingAttr: BindingFlags, 
        pRetVal: *mut *mut SAFEARRAY
    ) -> HRESULT,

    /// Placeholder for the `GetField` method. Not used directly.
    GetField: *const c_void,

    /// Placeholder for the `GetFields` method. Not used directly.
    GetFields: *const c_void,

    /// Placeholder for the `GetProperty` method. Not used directly.
    GetProperty: *const c_void,

    /// Placeholder for the `GetProperty_2` method. Not used directly.
    GetProperty_2: *const c_void,

    /// Placeholder for the `GetProperties` method. Not used directly.
    GetProperties: *const c_void,

    /// Placeholder for the `GetMember_2` method. Not used directly.
    GetMember_2: *const c_void,

    /// Placeholder for the `GetMembers` method. Not used directly.
    GetMembers: *const c_void,

    /// Placeholder for the `InvokeMember` method. Not used directly.
    InvokeMember: *const c_void,

    /// Placeholder for the `get_UnderlyingSystemType` method. Not used directly.
    get_UnderlyingSystemType: *const c_void,

    /// Placeholder for the `InvokeMember_2` method. Not used directly.
    InvokeMember_2: *const c_void,

    /// Invokes a method (static or instance) by name on the specified type or object.
    ///
    /// # Arguments
    /// 
    /// * `*mut c_void` - Pointer to the COM object implementing the interface.
    /// * `name` - The name of the member to invoke as a `BSTR`.
    /// * `invokeAttr` - Flags controlling invocation behavior.
    /// * `Binder` - Pointer to binder; typically `null`.
    /// * `Target` - The instance of the type for invocation.
    /// * `args` - Pointer to a `SAFEARRAY` of arguments.
    /// * `pRetVal` - Pointer to receive the invocation result.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    InvokeMember_3: unsafe extern "system" fn(
        *mut c_void, 
        name: BSTR, 
        invokeAttr: BindingFlags, 
        Binder: *mut c_void, 
        Target: VARIANT,
        args: *mut SAFEARRAY,
        pRetVal: *mut VARIANT
    ) -> HRESULT,

    /// Placeholder for the `GetConstructor` method. Not used directly.
    GetConstructor: *const c_void,

    /// Placeholder for the `GetConstructor_2` method. Not used directly.
    GetConstructor_2: *const c_void,

    /// Placeholder for the `GetConstructor_3` method. Not used directly.
    GetConstructor_3: *const c_void,

    /// Placeholder for the `GetConstructors_2` method. Not used directly.
    GetConstructors_2: *const c_void,

    /// Placeholder for the `get_TypeInitializer` method. Not used directly.
    get_TypeInitializer: *const c_void,

    /// Placeholder for the `GetMethod_3` method. Not used directly.
    GetMethod_3: *const c_void,

    /// Placeholder for the `GetMethod_4` method. Not used directly.
    GetMethod_4: *const c_void,

    /// Placeholder for the `GetMethod_5` method. Not used directly.
    GetMethod_5: *const c_void,

    /// Retrieves a method by name.
    ///
    /// # Arguments
    /// 
    /// * `*mut c_void` - Pointer to the COM object implementing the interface. 
    /// * `name` - A `BSTR` representing the method name.
    /// * `pRetVal` - Pointer that receives the `_MethodInfo` object.
    ///
    /// # Returns
    /// 
    /// * Returns an HRESULT indicating success or failure.
    GetMethod_6: unsafe extern "system" fn(
        *mut c_void, 
        name: BSTR, 
        pRetVal: *mut *mut _MethodInfo
    ) -> HRESULT,

    /// Placeholder for the method. Not used directly.
    GetMethods_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetField_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetFields_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetInterface_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetEvent_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetProperty_3: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetProperty_4: *const c_void,
    
    /// Placeholder for the method. Not used directly.
    GetProperty_5: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetProperty_6: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetProperty_7: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetProperties_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetNestedTypes_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetNestedType_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetMember_3: *const c_void,

    /// Placeholder for the method. Not used directly.
    GetMembers_2: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_Attributes: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNotPublic: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsPublic: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedPublic: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedPrivate: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedFamily: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedAssembly: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedFamANDAssem: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsNestedFamORAssem: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsAutoLayout: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsLayoutSequential: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsExplicitLayout: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsClass: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsInterface: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsValueType: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsAbstract: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsSealed: *const c_void,
    
    /// Placeholder for the method. Not used directly.
    get_IsEnum: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsSpecialName: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsImport: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsSerializable: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsAnsiClass: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsUnicodeClass: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsArray: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsByRef: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsPointer: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsPrimitive: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsCOMObject: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_HasElementType: *const c_void,
    
    /// Placeholder for the method. Not used directly.
    get_IsContextful: *const c_void,

    /// Placeholder for the method. Not used directly.
    get_IsMarshalByRef: *const c_void,

    /// Placeholder for the method. Not used directly.
    Equals_2: *const c_void
}

/// Specifies flags that control binding and the way in which members are searched and invoked.
/// 
/// These flags can be combined using bitwise operations to refine the scope of the invocation or search.
/// `BindingFlags` are commonly used in .NET reflection to determine if a method or property is
/// public, static, instance-based, and more.
#[repr(C)]
pub enum BindingFlags {
    /// Default binding, no special options.
    Default = 0,

    /// Ignores case when looking up members.
    IgnoreCase = 1,

    /// Only members declared at the level of the supplied type's hierarchy should be considered.
    DeclaredOnly = 2,

    /// Specifies instance members.
    Instance = 4,

    /// Specifies static members.
    Static = 8,

    /// Specifies public members.
    Public = 16,

    /// Specifies non-public members.
    NonPublic = 32,

    /// Includes inherited members in the search.
    FlattenHierarchy = 64,

    /// Specifies that the member to invoke is a method.
    InvokeMethod = 256,

    /// Creates an instance of the object.
    CreateInstance = 512,

    /// Specifies that the member to retrieve is a field.
    GetField = 1024,

    /// Specifies that the member to set is a field.
    SetField = 2048,

    /// Specifies that the member to retrieve is a property.
    GetProperty = 4096,

    /// Specifies that the member to set is a property.
    SetProperty = 8192,

    /// Sets a COM object property.
    PutDispProperty = 16384,

    /// Sets a COM object reference property.
    PutRefDispProperty = 32768,

    /// Uses the most precise match during binding.
    ExactBinding = 65536,

    /// Suppresses coercion of argument types during method invocation.
    SuppressChangeType = 131072,

    /// Allows binding to optional parameters.
    OptionalParamBinding = 262144,

    /// Ignores the return value of a method.
    IgnoreReturn = 16777216,
}

impl BitOr for BindingFlags {
    type Output = Self;

    /// Enables combining multiple `BindingFlags` using bitwise OR.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let flags = BindingFlags::Public | BindingFlags::Instance;
    /// ```
    fn bitor(self, rhs: Self) -> Self::Output {
        unsafe { std::mem::transmute::<u32, BindingFlags>(self as u32 | rhs as u32) }
    }
}