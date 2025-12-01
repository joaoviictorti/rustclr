use alloc::{string::String, vec::Vec};
use bitflags::bitflags;
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
        Variant::VARIANT,
        Ole::{SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound},
    },
};

use crate::com::{_MethodInfo, _PropertyInfo};
use crate::error::{ClrError, Result};
use crate::string::ComString;
use crate::variant::create_safe_args;
use crate::Invocation;

/// This struct represents the COM `_Type` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct _Type(windows_core::IUnknown);

impl _Type {
    /// Retrieves a method by its name from the type.
    #[inline]
    pub fn method(&self, name: &str) -> Result<_MethodInfo> {
        let method_name = name.to_bstr();
        self.GetMethod_6(method_name)
    }

    /// Finds a method by signature from the type.
    #[inline]
    pub fn method_signature(&self, name: &str) -> Result<_MethodInfo> {
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

    /// Finds a property by signature from the type.
    #[inline]
    pub fn property_signature(&self, name: &str) -> Result<_PropertyInfo> {
        let properties = self.properties();
        if let Ok(properties) = properties {
            for (property_name, property_info) in properties {
                if property_name == name {
                    return Ok(property_info);
                }
            }
        }

        Err(ClrError::PropertyNotFound)
    }

    /// Retrieves a property by name from the type.
    #[inline]
    pub fn property(&self, name: &str) -> Result<_PropertyInfo> {
        unsafe {
            let binding_flags = BindingFlags::Public
                | BindingFlags::Instance
                | BindingFlags::Static
                | BindingFlags::FlattenHierarchy
                | BindingFlags::NonPublic;

            let property_name = name.to_bstr();
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetProperty)(
                Interface::as_raw(self),
                property_name,
                binding_flags,
                &mut result,
            );

            if hr == 0 && !result.is_null() {
                Ok(_PropertyInfo::from_raw(result)?)
            } else {
                Err(ClrError::ApiError("GetProperty", hr))
            }
        }
    }

    /// Invokes a method on the type.
    #[inline]
    pub fn invoke(
        &self,
        name: &str,
        instance: Option<VARIANT>,
        args: Option<Vec<VARIANT>>,
        invocation_type: Invocation,
    ) -> Result<VARIANT> {
        let flags = match invocation_type {
            Invocation::Static => {
                BindingFlags::NonPublic
                    | BindingFlags::Public
                    | BindingFlags::Static
                    | BindingFlags::InvokeMethod
            }
            Invocation::Instance => {
                BindingFlags::NonPublic
                    | BindingFlags::Public
                    | BindingFlags::Instance
                    | BindingFlags::InvokeMethod
            }
        };

        let method_name = name.to_bstr();
        let args = args
            .as_ref()
            .map_or_else(|| Ok(null_mut()), |args| create_safe_args(args.to_vec()))?;

        let instance = instance.unwrap_or(unsafe { core::mem::zeroed::<VARIANT>() });
        self.InvokeMember_3(method_name, flags, instance, args)
    }

    /// Retrieves all methods of the type.
    #[inline]
    pub fn methods(&self) -> Result<Vec<(String, _MethodInfo)>> {
        let binding_flags = BindingFlags::Public
            | BindingFlags::Instance
            | BindingFlags::Static
            | BindingFlags::FlattenHierarchy
            | BindingFlags::NonPublic;

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

    /// Retrieves all properties of the type.
    #[inline]
    pub fn properties(&self) -> Result<Vec<(String, _PropertyInfo)>> {
        let binding_flags = BindingFlags::Public
            | BindingFlags::Instance
            | BindingFlags::Static
            | BindingFlags::FlattenHierarchy
            | BindingFlags::NonPublic;

        let sa_properties = self.GetProperties(binding_flags)?;
        if sa_properties.is_null() {
            return Err(ClrError::NullPointerError("GetProperties"));
        }

        let mut lbound = 0;
        let mut ubound = 0;
        let mut properties = Vec::new();
        unsafe {
            SafeArrayGetLBound(sa_properties, 1, &mut lbound);
            SafeArrayGetUBound(sa_properties, 1, &mut ubound);

            let mut p_property = null_mut::<_PropertyInfo>();
            for i in lbound..=ubound {
                let hr =
                    SafeArrayGetElement(sa_properties, &i, &mut p_property as *mut _ as *mut _);
                if hr != 0 || p_property.is_null() {
                    return Err(ClrError::ApiError("SafeArrayGetElement", hr));
                }

                let property = _PropertyInfo::from_raw(p_property as *mut c_void)?;
                let name = property.ToString()?;
                properties.push((name, property));
            }
        }

        Ok(properties)
    }

    /// Creates an `_Type` instance from a raw COM interface pointer.
    #[inline]
    pub fn from_raw(raw: *mut c_void) -> Result<_Type> {
        let iunknown = unsafe { IUnknown::from_raw(raw) };
        iunknown
            .cast::<_Type>()
            .map_err(|_| ClrError::CastingError("_Type"))
    }

    /// Retrieves the string representation of the type.
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

    /// Retrieves all properties matching the specified `BindingFlags`.
    #[inline]
    pub fn GetProperties(&self, bindingAttr: BindingFlags) -> Result<*mut SAFEARRAY> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetProperties)(
                Interface::as_raw(self),
                bindingAttr,
                &mut result,
            );

            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetProperties", hr))
            }
        }
    }

    /// Retrieves all methods matching the specified `BindingFlags`.
    #[inline]
    pub fn GetMethods(&self, bindingAttr: BindingFlags) -> Result<*mut SAFEARRAY> {
        unsafe {
            let mut result = null_mut();
            let hr = (Interface::vtable(self).GetMethods)(
                Interface::as_raw(self),
                bindingAttr,
                &mut result,
            );
            if hr == 0 {
                Ok(result)
            } else {
                Err(ClrError::ApiError("GetMethods", hr))
            }
        }
    }

    /// Retrieves a method by name.
    #[inline]
    pub fn GetMethod_6(&self, name: BSTR) -> Result<_MethodInfo> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).GetMethod_6)(Interface::as_raw(self), name, &mut result);
            if hr == 0 {
                _MethodInfo::from_raw(result as *mut c_void)
            } else {
                Err(ClrError::ApiError("GetMethod_6", hr))
            }
        }
    }

    /// Invokes a method (static or instance) by name on the specified type or object.
    #[inline]
    pub fn InvokeMember_3(
        &self,
        name: BSTR,
        invoke_attr: BindingFlags,
        instance: VARIANT,
        args: *mut SAFEARRAY,
    ) -> Result<VARIANT> {
        unsafe {
            let mut result = core::mem::zeroed();
            let hr = (Interface::vtable(self).InvokeMember_3)(
                Interface::as_raw(self),
                name,
                invoke_attr,
                null_mut(),
                instance,
                args,
                &mut result,
            );
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

bitflags! {
    /// Flags that control binding and the way in which members are searched and invoked.
    ///
    /// Uses a bitflag-backed type to avoid invalid enum transmute panics on Nightly.
    #[repr(transparent)]
    pub struct BindingFlags: u32 {
        /// Default binding, no special options.
        const Default = 0;
        /// Ignores case when looking up members.
        const IgnoreCase = 1;
        /// Only members declared at the level of the supplied type's hierarchy should be considered.
        const DeclaredOnly = 2;
        /// Specifies instance members.
        const Instance = 4;
        /// Specifies static members.
        const Static = 8;
        /// Specifies public members.
        const Public = 16;
        /// Specifies non-public members.
        const NonPublic = 32;
        /// Includes inherited members in the search.
        const FlattenHierarchy = 64;
        /// Specifies that the member to invoke is a method.
        const InvokeMethod = 256;
        /// Creates an instance of the object.
        const CreateInstance = 512;
        /// Specifies that the member to retrieve is a field.
        const GetField = 1024;
        /// Specifies that the member to set is a field.
        const SetField = 2048;
        /// Specifies that the member to retrieve is a property.
        const GetProperty = 4096;
        /// Specifies that the member to set is a property.
        const SetProperty = 8192;
        /// Sets a COM object property.
        const PutDispProperty = 16384;
        /// Sets a COM object reference property.
        const PutRefDispProperty = 32768;
        /// Uses the most precise match during binding.
        const ExactBinding = 65536;
        /// Suppresses coercion of argument types during method invocation.
        const SuppressChangeType = 131072;
        /// Allows binding to optional parameters.
        const OptionalParamBinding = 262144;
        /// Ignores the return value of a method.
        const IgnoreReturn = 16777216;
    }
}

/// Raw COM vtable for the `_Type` interface.
#[repr(C)]
pub struct _Type_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    
    // IDispatch methods
    GetTypeInfoCount: *const c_void,
    GetTypeInfo: *const c_void,
    GetIDsOfNames: *const c_void,
    Invoke: *const c_void,
    
    // Methods specific to the COM interface
    get_ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut BSTR) -> HRESULT,
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
    get_Guid: *const c_void,
    get_Module: *const c_void,
    get_Assembly: *const c_void,
    get_TypeHandle: *const c_void,
    get_FullName: *const c_void,
    get_Namespace: *const c_void,
    get_AssemblyQualifiedName: *const c_void,
    GetArrayRank: *const c_void,
    get_BaseType: *const c_void,
    GetConstructors: *const c_void,
    GetInterface: *const c_void,
    GetInterfaces: *const c_void,
    FindInterfaces: *const c_void,
    GetEvent: *const c_void,
    GetEvents: *const c_void,
    GetEvents_2: *const c_void,
    GetNestedTypes: *const c_void,
    GetNestedType: *const c_void,
    GetMember: *const c_void,
    GetDefaultMembers: *const c_void,
    FindMembers: *const c_void,
    GetElementType: *const c_void,
    IsSubclassOf: *const c_void,
    IsInstanceOfType: *const c_void,
    IsAssignableFrom: *const c_void,
    GetInterfaceMap: *const c_void,
    GetMethod: *const c_void,
    GetMethod_2: *const c_void,
    GetMethods: unsafe extern "system" fn(
        this: *mut c_void,
        bindingAttr: BindingFlags,
        pRetVal: *mut *mut SAFEARRAY,
    ) -> HRESULT,
    GetField: *const c_void,
    GetFields: *const c_void,
    pub GetProperty: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR,
        bindingAttr: BindingFlags,
        result: *mut *mut c_void,
    ) -> HRESULT,
    GetProperty_2: *const c_void,
    GetProperties: unsafe extern "system" fn(
        this: *mut c_void,
        bindingAttr: BindingFlags,
        pRetVal: *mut *mut SAFEARRAY,
    ) -> HRESULT,
    GetMember_2: *const c_void,
    GetMembers: *const c_void,
    InvokeMember: *const c_void,
    get_UnderlyingSystemType: *const c_void,
    InvokeMember_2: *const c_void,
    InvokeMember_3: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR,
        invokeAttr: BindingFlags,
        Binder: *mut c_void,
        Target: VARIANT,
        args: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    GetConstructor: *const c_void,
    GetConstructor_2: *const c_void,
    GetConstructor_3: *const c_void,
    GetConstructors_2: *const c_void,
    get_TypeInitializer: *const c_void,
    GetMethod_3: *const c_void,
    GetMethod_4: *const c_void,
    GetMethod_5: *const c_void,
    GetMethod_6: unsafe extern "system" fn(
        this: *mut c_void,
        name: BSTR,
        pRetVal: *mut *mut _MethodInfo,
    ) -> HRESULT,
    GetMethods_2: *const c_void,
    GetField_2: *const c_void,
    GetFields_2: *const c_void,
    GetInterface_2: *const c_void,
    GetEvent_2: *const c_void,
    GetProperty_3: *const c_void,
    GetProperty_4: *const c_void,
    GetProperty_5: *const c_void,
    GetProperty_6: *const c_void,
    GetProperty_7: *const c_void,
    GetProperties_2: *const c_void,
    GetNestedTypes_2: *const c_void,
    GetNestedType_2: *const c_void,
    GetMember_3: *const c_void,
    GetMembers_2: *const c_void,
    get_Attributes: *const c_void,
    get_IsNotPublic: *const c_void,
    get_IsPublic: *const c_void,
    get_IsNestedPublic: *const c_void,
    get_IsNestedPrivate: *const c_void,
    get_IsNestedFamily: *const c_void,
    get_IsNestedAssembly: *const c_void,
    get_IsNestedFamANDAssem: *const c_void,
    get_IsNestedFamORAssem: *const c_void,
    get_IsAutoLayout: *const c_void,
    get_IsLayoutSequential: *const c_void,
    get_IsExplicitLayout: *const c_void,
    get_IsClass: *const c_void,
    get_IsInterface: *const c_void,
    get_IsValueType: *const c_void,
    get_IsAbstract: *const c_void,
    get_IsSealed: *const c_void,
    get_IsEnum: *const c_void,
    get_IsSpecialName: *const c_void,
    get_IsImport: *const c_void,
    get_IsSerializable: *const c_void,
    get_IsAnsiClass: *const c_void,
    get_IsUnicodeClass: *const c_void,
    get_IsArray: *const c_void,
    get_IsByRef: *const c_void,
    get_IsPointer: *const c_void,
    get_IsPrimitive: *const c_void,
    get_IsCOMObject: *const c_void,
    get_HasElementType: *const c_void,
    get_IsContextful: *const c_void,
    get_IsMarshalByRef: *const c_void,
    Equals_2: *const c_void,
}
