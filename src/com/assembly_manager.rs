use core::{ffi::c_void, ops::Deref, ptr::null_mut};
use windows_core::{GUID, Interface};
use crate::com::IHostAssemblyStore;

/// Represents the COM `IHostAssemblyManager` interface.
#[repr(C)]
#[derive(Clone, Debug)]
pub struct IHostAssemblyManager(windows_core::IUnknown);

/// Trait that defines the implementation of the `IHostAssemblyManager` interface.
pub trait IHostAssemblyManager_Impl: windows_core::IUnknownImpl {
    /// Retrieves assemblies not stored in the host store.
    fn GetNonHostStoreAssemblies(&self) -> windows_core::Result<()>;

    /// Retrieves the host's `IHostAssemblyStore`.
    fn GetAssemblyStore(&self) -> windows_core::Result<IHostAssemblyStore>;
}

impl IHostAssemblyManager_Vtbl {
    /// Constructs the virtual function table (vtable) for `IHostAssemblyManager`.
    ///
    /// This binds the trait implementation to the raw function pointers expected by COM.
    pub const fn new<Identity: IHostAssemblyManager_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn GetNonHostStoreAssemblies<
            Identity: IHostAssemblyManager_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            ppreferencelist: *mut *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IHostAssemblyManager_Impl::GetNonHostStoreAssemblies(this) {
                    Ok(_) => {
                        ppreferencelist.write(null_mut());
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }

        unsafe extern "system" fn GetAssemblyStore<
            Identity: IHostAssemblyManager_Impl,
            const OFFSET: isize,
        >(
            this: *mut c_void,
            ppassemblystore: *mut *mut c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IHostAssemblyManager_Impl::GetAssemblyStore(this) {
                    Ok(ok) => {
                        ppassemblystore.write(core::mem::transmute(ok));
                        windows_core::HRESULT(0)
                    }

                    Err(err) => err.into(),
                }
            }
        }

        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            GetNonHostStoreAssemblies: GetNonHostStoreAssemblies::<Identity, OFFSET>,
            GetAssemblyStore: GetAssemblyStore::<Identity, OFFSET>,
        }
    }

    /// Checks if the given IID matches the `IHostAssemblyManager` interface.
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IHostAssemblyManager as windows_core::Interface>::IID
    }
}

impl windows_core::RuntimeName for IHostAssemblyManager {}

unsafe impl Interface for IHostAssemblyManager {
    type Vtable = IHostAssemblyManager_Vtbl;

    /// The interface identifier (IID) for the `IHostAssemblyManager` COM interface.
    ///
    /// This GUID is used to identify the `IHostAssemblyManager` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostAssemblyManager` interface.
    const IID: GUID = GUID::from_u128(0x613dabd7_62b2_493e_9e65_c1e32a1e0c5e);
}

impl Deref for IHostAssemblyManager {
    type Target = windows_core::IUnknown;

    /// The interface identifier (IID) for the `IHostAssemblyManager` COM interface.
    ///
    /// This GUID is used to identify the `IHostAssemblyManager` interface when calling
    /// COM methods like `QueryInterface`. It is defined based on the standard
    /// .NET CLR IID for the `IHostAssemblyManager` interface.
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}

/// Raw COM vtable for the `IHostAssemblyManager` interface.
#[repr(C)]
pub struct IHostAssemblyManager_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,

    // Methods specific to the COM interface
    pub GetNonHostStoreAssemblies: unsafe extern "system" fn(
        this: *mut c_void,
        ppreferencelist: *mut *mut c_void,
    ) -> windows_core::HRESULT,
    pub GetAssemblyStore: unsafe extern "system" fn(
        this: *mut c_void,
        ppassemblystore: *mut *mut c_void,
    ) -> windows_core::HRESULT,
}
