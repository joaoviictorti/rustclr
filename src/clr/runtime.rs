use core::{
    mem::transmute,
    ffi::c_void, 
    ptr::null_mut
};
use alloc::{
    format,
    string::{String, ToString},
    vec::Vec,
};

use obfstr::obfstr as s;
use dinvk::winapis::{
    NtCurrentProcess, 
    NtProtectVirtualMemory, 
    NT_SUCCESS
};
use windows_core::{IUnknown, Interface, PCWSTR};
use windows_sys::Win32::{
    UI::Shell::SHCreateMemStream,
    System::Memory::PAGE_EXECUTE_READWRITE
};

use super::hosting::RustClrControl;
use crate::{com::*, variant::Variant};
use crate::error::{ClrError, Result};

/// Holds the runtime state and execution configuration for the CLR.
#[derive(Default, Debug, Clone)]
pub struct RustClrRuntime<'a> {
    /// Raw buffer containing the loaded .NET assembly.
    pub buffer: &'a [u8],

    /// Unique identity name of the loaded .NET assembly.
    pub identity_assembly: String,

    /// Version of the .NET runtime to load.
    pub runtime_version: Option<RuntimeVersion>,

    /// Optional name of the application domain to be created.
    pub domain_name: Option<String>,

    /// Current application domain.
    pub app_domain: Option<_AppDomain>,

    /// Runtime host instance used to manage CLR execution.
    pub cor_runtime_host: Option<ICorRuntimeHost>,
}

impl<'a> RustClrRuntime<'a> {
    /// Creates a new `RustClrRuntime`.
    pub fn new(buffer: &'a [u8]) -> Self {
        Self {
            buffer,
            identity_assembly: String::new(),
            runtime_version: None,
            domain_name: None,
            app_domain: None,
            cor_runtime_host: None,
        }
    }

    /// Initializes the CLR environment and prepares it for execution.
    ///
    /// This loads the requested CLR version, resolves the assembly identity,
    /// starts the runtime if needed, and creates the initial application domain.
    pub fn prepare(&mut self) -> Result<()> {
        // Creates the MetaHost to access the available CLR versions
        let meta_host = self.create_meta_host()?;

        // Gets information about the specified (or default) runtime version
        let runtime_info = self.get_runtime_info(&meta_host)?;

        // Get ICLRAssemblyIdentityManager via GetProcAddress
        let addr = runtime_info.GetProcAddress(s!("GetCLRIdentityManager"))?;
        let GetCLRIdentityManager = unsafe { transmute::<*mut c_void, CLRIdentityManagerType>(addr) };
        let mut ptr = null_mut();
        GetCLRIdentityManager(&ICLRAssemblyIdentityManager::IID, &mut ptr);

        // Create a stream for the in-memory assembly and get the identity string from the stream
        let iclr_assembly = ICLRAssemblyIdentityManager::from_raw(ptr)?;
        let stream = unsafe { SHCreateMemStream(self.buffer.as_ptr(), self.buffer.len() as u32) };
        self.identity_assembly = iclr_assembly.get_identity_stream(stream, 0)?;

        // Creates the `ICLRuntimeHost`
        let iclr_runtime_host = self.get_clr_runtime_host(&runtime_info)?;

        // Checks if the runtime is started
        if runtime_info.IsLoadable().is_ok() && !runtime_info.is_started() {
            // Create and register IHostControl with custom assembly and identity
            let host_control: IHostControl = RustClrControl::new(self.buffer, &self.identity_assembly).into();
            iclr_runtime_host.SetHostControl(&host_control)?;
            
            // Starts the CLR runtime
            self.start_runtime(&iclr_runtime_host)?;
        }

        // Creates the `ICorRuntimeHost`
        let cor_runtime_host = self.get_icor_runtime_host(&runtime_info)?;

        // Initializes the specified application domain or the default
        self.init_app_domain(&cor_runtime_host)?;

        // Saves the runtime host for future use
        self.cor_runtime_host = Some(self.get_icor_runtime_host(&runtime_info)?);
        Ok(())
    }

    /// Returns the active application domain.
    pub fn get_app_domain(&mut self) -> Result<_AppDomain> {
        self.app_domain
            .clone()
            .ok_or(ClrError::NoDomainAvailable)
    }

    /// Creates an instance of [`ICLRMetaHost`].
    fn create_meta_host(&self) -> Result<ICLRMetaHost> {
        CLRCreateInstance::<ICLRMetaHost>(&CLSID_CLRMETAHOST)
            .map_err(|e| ClrError::MetaHostCreationError(format!("{e}")))
    }

    /// Retrieves runtime information based on the selected .NET version.
    fn get_runtime_info(&self, meta_host: &ICLRMetaHost) -> Result<ICLRRuntimeInfo> {
        let runtime_version = &self.runtime_version.unwrap_or(RuntimeVersion::V4);
        let version_wide = runtime_version.to_vec();
        let version = PCWSTR(version_wide.as_ptr());
        meta_host
            .GetRuntime::<ICLRRuntimeInfo>(version)
            .map_err(|error| ClrError::RuntimeInfoError(format!("{error}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    fn get_icor_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICorRuntimeHost> {
        runtime_info
            .GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|error| ClrError::RuntimeHostError(format!("{error}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    fn get_clr_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICLRuntimeHost> {
        runtime_info
            .GetInterface::<ICLRuntimeHost>(&CLSID_ICLR_RUNTIME_HOST)
            .map_err(|error| ClrError::RuntimeHostError(format!("{error}")))
    }

    /// Starts the CLR runtime using the provided runtime host.
    fn start_runtime(&self, iclr_runtime_host: &ICLRuntimeHost) -> Result<()> {
        if iclr_runtime_host.Start() != 0 {
            return Err(ClrError::RuntimeStartError);
        }
        Ok(())
    }

    /// Initializes the application domain with the specified domain name or
    /// creates a unique default domain.
    fn init_app_domain(&mut self, cor_runtime_host: &ICorRuntimeHost) -> Result<()> {
        let app_domain = if let Some(domain_name) = &self.domain_name {
            let wide_domain_name = domain_name
                .encode_utf16()
                .chain(Some(0))
                .collect::<Vec<u16>>();

            cor_runtime_host.CreateDomain(PCWSTR(wide_domain_name.as_ptr()), null_mut())?
        } else {
            let uuid = uuid()
                .to_string()
                .encode_utf16()
                .chain(Some(0))
                .collect::<Vec<u16>>();

            cor_runtime_host.CreateDomain(PCWSTR(uuid.as_ptr()), null_mut())?
        };

        // Saves the created application domain
        self.app_domain = Some(app_domain);
        Ok(())
    }

    /// Unloads the current application domain.
    pub fn unload_domain(&self) -> Result<()> {
        if let (Some(cor_runtime_host), Some(app_domain)) =
            (&self.cor_runtime_host, &self.app_domain)
        {
            cor_runtime_host.UnloadDomain(
                app_domain
                    .cast::<windows_core::IUnknown>()
                    .map(|i| i.as_raw().cast())
                    .unwrap_or(null_mut()),
            )?
        }

        Ok(())
    }
}

/// Represents the .NET runtime versions supported by RustClr.
#[derive(Debug, Clone, Copy)]
pub enum RuntimeVersion {
    /// .NET Framework 2.0.
    V2,

    /// .NET Framework 3.0.
    V3,

    /// .NET Framework 4.0.
    V4,

    /// Represents an unsupported .NET runtime version.
    UNKNOWN,
}

impl RuntimeVersion {
    /// Converts the `RuntimeVersion` to a wide string representation as a `Vec<u16>`.
    pub fn to_vec(self) -> Vec<u16> {
        let runtime_version = match self {
            RuntimeVersion::V2 => "v2.0.50727",
            RuntimeVersion::V3 => "v3.0",
            RuntimeVersion::V4 => "v4.0.30319",
            RuntimeVersion::UNKNOWN => "UNKNOWN",
        };

        runtime_version
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>()
    }
}

/// Generates a uuid used to create the AppDomain
pub fn uuid() -> uuid::Uuid {
    let mut buf = [0u8; 16];

    for i in 0..4 {
        let ticks = unsafe { core::arch::x86_64::_rdtsc() };
        buf[i * 4] = ticks as u8;
        buf[i * 4 + 1] = (ticks >> 8) as u8;
        buf[i * 4 + 2] = (ticks >> 16) as u8;
        buf[i * 4 + 3] = (ticks >> 24) as u8;
    }

    uuid::Uuid::from_bytes(buf)
}

/// Patches `System.Environment.Exit` to prevent the CLR from terminating the host process.
///
/// This replaces the first byte of `Environment.Exit` with `0xC3` (`ret`), effectively
/// neutralizing the method.
pub fn patch_exit(mscorlib: &_Assembly) -> Result<()> {
    // Resolve System.Environment type and the Exit method
    let env = mscorlib.resolve_type(s!("System.Environment"))?;
    let exit = env.method(s!("Exit"))?;

    // Resolve System.Reflection.MethodInfo.MethodHandle property
    let method_info = mscorlib.resolve_type(s!("System.Reflection.MethodInfo"))?;
    let method_handle = method_info.property(s!("MethodHandle"))?;

    // Convert the Exit method into a COM IUnknown pointer
    let instance = exit
        .cast::<IUnknown>()
        .map_err(|_| ClrError::Msg("Failed to cast to IUnknown"))?;

    // Call to retrieve the RuntimeMethodHandle
    let method_handle_exit = method_handle.value(Some(instance.to_variant()), None)?;

    // Get the native address of Environment.Exit
    let runtime_method = mscorlib.resolve_type(s!("System.RuntimeMethodHandle"))?;
    let get_function_pointer = runtime_method.method(s!("GetFunctionPointer"))?;
    let ptr = get_function_pointer.invoke(Some(method_handle_exit), None)?;

    // Extract pointer from VARIANT
    let mut addr_exit = unsafe { ptr.Anonymous.Anonymous.Anonymous.byref };
    let mut old = 0;
    let mut size = 1;

    // Change memory protection to RWX for patching
    if !NT_SUCCESS(NtProtectVirtualMemory(
        NtCurrentProcess(),
        &mut addr_exit,
        &mut size,
        PAGE_EXECUTE_READWRITE,
        &mut old,
    )) {
        return Err(ClrError::Msg(
            "failed to change memory protection to RWX",
        ));
    }

    // Overwrite first byte with RET (0xC3)
    unsafe { *(ptr.Anonymous.Anonymous.Anonymous.byref as *mut u8) = 0xC3 };

    // Restore original protection
    if !NT_SUCCESS(NtProtectVirtualMemory(
        NtCurrentProcess(),
        &mut addr_exit,
        &mut size,
        old,
        &mut old,
    )) {
        return Err(ClrError::Msg(
            "failed to restore memory protection",
        ));
    }

    Ok(())
}
