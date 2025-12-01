use core::{ffi::c_void, ops::Deref};
use windows_core::{GUID, Interface};

/// This struct represents the COM `IHostControl` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct IHostControl(windows_core::IUnknown);

/// Trait representing the implementation of the `IHostControl` interface.
pub trait IHostControl_Impl: windows_core::IUnknownImpl {
    /// Requests a host-provided manager object that implements the interface specified by `riid`.
    fn GetHostManager(
        &self,
        riid: *const GUID,
        ppobject: *mut *mut c_void,
    ) -> windows_core::Result<()>;

    /// Notifies the host that the CLR has created an `AppDomainManager` for a new AppDomain.
    fn SetAppDomainManager(
        &self,
        dwappdomainid: u32,
        punkappdomainmanager: windows_core::Ref<'_, windows_core::IUnknown>,
    ) -> windows_core::Result<()>;
}

impl IHostControl_Vtbl {
    /// Creates a new virtual table for the `IHostControl` implementation.
    ///
    /// This table contains function pointers for each method exposed by the interface.
    pub const fn new<Identity: IHostControl_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn GetHostManager<
            Identity: IHostControl_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            riid: *const GUID,
            ppobject: *mut *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostControl_Impl::GetHostManager(
                    this,
                    core::mem::transmute_copy(&riid),
                    core::mem::transmute_copy(&ppobject),
                )
                .into()
            }
        }

        unsafe extern "system" fn SetAppDomainManager<
            Identity: IHostControl_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            dwappdomainid: u32,
            punkappdomainmanager: *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IHostControl_Impl::SetAppDomainManager(
                    this,
                    core::mem::transmute_copy(&dwappdomainid),
                    core::mem::transmute_copy(&punkappdomainmanager),
                )
                .into()
            }
        }

        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            GetHostManager: GetHostManager::<Identity, OFFSET>,
            SetAppDomainManager: SetAppDomainManager::<Identity, OFFSET>,
        }
    }

    /// Verifies if a given interface ID matches `IHostControl`.
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IHostControl as windows_core::Interface>::IID
    }
}

impl windows_core::RuntimeName for IHostControl {}

unsafe impl Interface for IHostControl {
    type Vtable = IHostControl_Vtbl;

    /// The interface identifier (IID) for the `IHostControl` COM interface.
    ///
    /// This GUID is used to identify the `IHostControl` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostControl` interface.
    const IID: GUID = GUID::from_u128(0x02ca073c_7079_4860_880a_c2f7a449c991);
}

impl Deref for IHostControl {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `IHostControl` COM interface.
    ///
    /// This GUID is used to identify the `IHostControl` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostControl` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `IHostControl` interface.
#[repr(C)]
pub struct IHostControl_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub GetHostManager: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const GUID,
        ppobject: *mut *mut c_void,
    ) -> windows_core::HRESULT,
    pub SetAppDomainManager: unsafe extern "system" fn(
        this: *mut c_void,
        dwappdomainid: u32,
        punkappdomainmanager: *mut c_void,
    ) -> windows_core::HRESULT,
}
