use alloc::string::{String, ToString};
use core::{ffi::c_void, ptr::null_mut};

use obfstr::obfstr as s;
use windows_core::*;
use windows_sys::Win32::UI::Shell::SHCreateMemStream;

use crate::com::*;

/// Implements `IHostControl`.
#[implement(IHostControl)]
pub struct RustClrControl {
    /// Host manager responsible for resolving assemblies.
    manager: IHostAssemblyManager,
}

impl RustClrControl {
    /// Creates a new `RustClrControl` with the target assembly and buffer.
    pub fn new(buffer: &[u8], assembly: &str) -> Self {
        Self {
            manager: RustClrManager::new(buffer, assembly.to_string()).into(),
        }
    }
}

impl IHostControl_Impl for RustClrControl_Impl {
    /// Returns `IHostAssemblyManager` when requested.
    fn GetHostManager(&self, riid: *const GUID, ppobject: *mut *mut c_void) -> Result<()> {
        unsafe {
            if *riid == IHostAssemblyManager::IID {
                *ppobject = self.manager.as_raw();
                return Ok(());
            }

            // IID_IHostTaskManager
            // IID_IHostThreadpoolManager
            // IID_IHostSyncManager
            // IID_IHostAssemblyManager
            // IID_IHostGCManager
            // IID_IHostPolicyManager
            // IHostSecurityManager
            *ppobject = null_mut();
            Err(Error::new(
                HRESULT(0x80004002u32 as i32),
                s!("E_NOINTERFACE"),
            ))
        }
    }

    fn SetAppDomainManager(
        &self,
        _dwappdomainid: u32,
        _punkappdomainmanager: Ref<'_, IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

/// Implements `IHostAssemblyManager`.
#[implement(IHostAssemblyManager)]
pub struct RustClrManager {
    /// Store responsible for resolving and serving assemblies.
    store: IHostAssemblyStore,
}

impl RustClrManager {
    /// Creates a new [`RustClrManager`].
    pub fn new(buffer: &[u8], assembly: String) -> Self {
        Self {
            store: RustClrStore::new(buffer, assembly).into(),
        }
    }
}

impl IHostAssemblyManager_Impl for RustClrManager_Impl {
    fn GetNonHostStoreAssemblies(&self) -> Result<()> {
        Ok(())
    }

    /// Returns the custom assembly store used to resolve in-memory assemblies.
    fn GetAssemblyStore(&self) -> Result<IHostAssemblyStore> {
        Ok(self.store.clone())
    }
}

/// Implements `IHostAssemblyStore`
#[implement(IHostAssemblyStore)]
pub struct RustClrStore<'a> {
    /// Assembly bytes.
    buffer: &'a [u8],

    /// Identity name to match.
    assembly: String,
}

impl<'a> RustClrStore<'a> {
    /// Creates a new `RustClrManager`.
    pub fn new(buffer: &'a [u8], assembly: String) -> Self {
        Self { buffer, assembly }
    }
}

impl IHostAssemblyStore_Impl for RustClrStore_Impl<'_> {
    /// Returns the managed assembly image from memory when the identity matches.
    fn ProvideAssembly(
        &self,
        pbindinfo: *const AssemblyBindInfo,
        passemblyid: *mut u64,
        pcontext: *mut u64,
        ppstmassemblyimage: *mut *mut c_void,
        _ppstmpdb: *mut *mut c_void,
    ) -> Result<()> {
        let identity = unsafe { (*pbindinfo).lpPostPolicyIdentity.to_string() }?;
        if self.assembly == identity {
            let stream = unsafe { 
                SHCreateMemStream(self.buffer.as_ptr(), self.buffer.len() as u32) 
            };
            unsafe { *passemblyid = 800 };
            unsafe { *pcontext = 0 }
            unsafe { *ppstmassemblyimage = stream };
            return Ok(());
        }

        Err(Error::new(
            HRESULT(0x80070002u32 as i32),
            s!("assembly not recognized"),
        ))
    }

    /// Always returns `ERROR_FILE_NOT_FOUND` as this implementation does not
    /// support module resolution.
    fn ProvideModule(
        &self,
        _pbindinfo: *const ModuleBindInfo,
        _pdwmoduleid: *mut u32,
        _ppstmmoduleimage: *mut *mut c_void,
        _ppstmpdb: *mut *mut c_void,
    ) -> Result<()> {
        Err(Error::new(
            HRESULT(0x80070002u32 as i32),
            s!("module not recognized"),
        ))
    }
}
