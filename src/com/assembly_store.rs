use core::{ffi::c_void, mem::transmute_copy, ops::Deref};
use windows_core::{GUID, Interface};

/// This struct represents the COM `IHostAssemblyStore` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct IHostAssemblyStore(windows_core::IUnknown);

/// Trait representing the implementation of the `IHostAssemblyStore` interface.
pub trait IHostAssemblyStore_Impl: windows_core::IUnknownImpl {
    /// Provides an assembly image in response to a bind request.
    fn ProvideAssembly(
        &self,
        pbindinfo: *const AssemblyBindInfo,
        passemblyid: *mut u64,
        pcontext: *mut u64,
        ppstmassemblyimage: *mut *mut c_void,
        ppstmpdb: *mut *mut c_void,
    ) -> windows_core::Result<()>;

    /// Provides a module image in response to a bind request.
    fn ProvideModule(
        &self,
        pbindinfo: *const ModuleBindInfo,
        pdwmoduleid: *mut u32,
        ppstmmoduleimage: *mut *mut c_void,
        ppstmpdb: *mut *mut c_void,
    ) -> windows_core::Result<()>;
}

impl IHostAssemblyStore_Vtbl {
    /// Creates a new virtual table for the `IHostAssemblyStore` implementation.
    ///
    /// This table contains function pointers for each method exposed by the interface.
    pub const fn new<Identity: IHostAssemblyStore_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn ProvideAssembly<
            Identity: IHostAssemblyStore_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            pbindinfo: *const AssemblyBindInfo,
            passemblyid: *mut u64,
            pcontext: *mut u64,
            ppstmassemblyimage: *mut *mut c_void,
            ppstmpdb: *mut *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostAssemblyStore_Impl::ProvideAssembly(
                    this,
                    transmute_copy(&pbindinfo),
                    transmute_copy(&passemblyid),
                    transmute_copy(&pcontext),
                    transmute_copy(&ppstmassemblyimage),
                    transmute_copy(&ppstmpdb),
                )
                .into()
            }
        }

        unsafe extern "system" fn ProvideModule<
            Identity: IHostAssemblyStore_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            pbindinfo: *const ModuleBindInfo,
            pdwmoduleid: *mut u32,
            ppstmmoduleimage: *mut *mut c_void,
            ppstmpdb: *mut *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostAssemblyStore_Impl::ProvideModule(
                    this,
                    transmute_copy(&pbindinfo),
                    transmute_copy(&pdwmoduleid),
                    transmute_copy(&ppstmmoduleimage),
                    transmute_copy(&ppstmpdb),
                )
                .into()
            }
        }

        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            ProvideAssembly: ProvideAssembly::<Identity, OFFSET>,
            ProvideModule: ProvideModule::<Identity, OFFSET>,
        }
    }

    /// Verifies if a given interface ID matches `IHostAssemblyStore`.
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IHostAssemblyStore as windows_core::Interface>::IID
    }
}

impl windows_core::RuntimeName for IHostAssemblyStore {}

/// Struct containing metadata needed for binding an assembly in the CLR.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct AssemblyBindInfo {
    /// Identifier of the application domain making the request.
    pub dwAppDomainId: u32,

    /// The identity of the referenced assembly (pre-policy).
    pub lpReferencedIdentity: windows_core::PCWSTR,

    /// The identity of the assembly after policy has been applied.
    pub lpPostPolicyIdentity: windows_core::PCWSTR,

    /// The level of policy applied (e.g., application, machine, etc).
    pub ePolicyLevel: u32,
}

/// Struct containing metadata needed for binding a module in the CLR.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct ModuleBindInfo {
    /// Identifier of the application domain making the request.
    pub dwAppDomainId: u32,

    /// The identity of the containing assembly.
    pub lpAssemblyIdentity: windows_core::PCWSTR,

    /// The name of the module being requested.
    pub lpModuleName: windows_core::PCWSTR,
}

unsafe impl Interface for IHostAssemblyStore {
    type Vtable = IHostAssemblyStore_Vtbl;

    /// The interface identifier (IID) for the `IHostAssemblyStore` COM interface.
    ///
    /// This GUID is used to identify the `IHostAssemblyStore` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostAssemblyStore` interface.
    const IID: GUID = GUID::from_u128(0x613dabd7_62b2_493e_9e65_c1e32a1e0c5e);
}

impl Deref for IHostAssemblyStore {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `IHostAssemblyStore` COM interface.
    ///
    /// This GUID is used to identify the `IHostAssemblyStore` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostAssemblyStore` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `IHostAssemblyStore` interface.
#[repr(C)]
pub struct IHostAssemblyStore_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub ProvideAssembly: unsafe extern "system" fn(
        this: *mut c_void,
        pbindinfo: *const AssemblyBindInfo,
        passemblyid: *mut u64,
        pcontext: *mut u64,
        ppstmassemblyimage: *mut *mut c_void,
        ppstmpdb: *mut *mut c_void,
    ) -> windows_core::HRESULT,
    pub ProvideModule: unsafe extern "system" fn(
        this: *mut c_void,
        pbindinfo: *const ModuleBindInfo,
        pdwmoduleid: *mut u32,
        ppstmmoduleimage: *mut *mut c_void,
        ppstmpdb: *mut *mut c_void,
    ) -> windows_core::HRESULT,
}
