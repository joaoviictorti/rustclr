use core::{ffi::c_void, ptr::null_mut};
use alloc::{
    boxed::Box,
    format,
    string::{String, ToString},
    vec,
    vec::Vec,
};

use obfstr::obfstr as s;
use dinvk::{
    NtCurrentProcess, 
    NtProtectVirtualMemory, 
    NT_SUCCESS
};
use windows_core::{IUnknown, Interface, PCWSTR};
use windows_sys::Win32::{
    UI::Shell::SHCreateMemStream,
    System::{
        Memory::PAGE_EXECUTE_READWRITE,
        Variant::{VARIANT, VariantClear},
    },
};

use super::{com::*, data::*};
use super::{
    error::ClrError,
    host_control::RustClrControl
};
use super::{
    uuid,
    create_safe_array_args,
    file::{read_file, validate_file}
};
use super::{
    Invocation, 
    Result, 
    Variant,
    WinStr,
};

/// Represents a Rust interface to the Common Language Runtime (CLR).
#[derive(Debug, Clone)]
pub struct RustClr<'a> {
    /// Buffer containing the .NET assembly in bytes.
    buffer: &'a [u8],

    /// Flag to indicate if output redirection is enabled.
    redirect_output: bool,

    /// Whether to patch `System.Environment.Exit` to prevent the process from terminating.
    patch_exit: bool,

    /// The identity name of the assembly being loaded or executed.
    identity_assembly: String,

    /// Name of the application domain to create or use.
    domain_name: Option<String>,

    /// .NET runtime version to use.
    runtime_version: Option<RuntimeVersion>,

    /// Arguments to pass to the .NET assembly's `Main` method.
    args: Option<Vec<String>>,

    /// Current application domain where the assembly is loaded.
    app_domain: Option<_AppDomain>,

    /// Host for the CLR runtime.
    cor_runtime_host: Option<ICorRuntimeHost>,
}

impl Default for RustClr<'_> {
    fn default() -> Self {
        Self {
            buffer: &[],
            runtime_version: None,
            redirect_output: false,
            patch_exit: false,
            identity_assembly: String::new(),
            domain_name: None,
            args: None,
            app_domain: None,
            cor_runtime_host: None,
        }
    }
}

impl<'a> RustClr<'a> {
    /// Creates a new [`RustClr`] instance with the specified assembly buffer.
    ///
    /// # Arguments
    ///
    /// * `source` - A value convertible into [`ClrSource`], representing either a file path or a byte buffer.
    ///
    /// # Returns
    ///
    /// * `Ok(Self)` - If the buffer is valid and the [`RustClr`] instance is created successfully.
    /// * `Err(ClrError)` - If the buffer validation fails (e.g., not a valid .NET assembly).
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     // Load a sample .NET assembly into a buffer
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Create a new RustClr instance
    ///     let clr = RustClr::new(&buffer)?;
    ///     println!("RustClr instance created successfully.");
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn new<T: Into<ClrSource<'a>>>(source: T) -> Result<Self> {
        let buffer = match source.into() {
            // Try reading the file
            ClrSource::File(path) => Box::leak(read_file(path)?.into_boxed_slice()),

            // Creates the .NET directly from the buffer
            ClrSource::Buffer(buffer) => buffer,
        };

        // Checks if it is a valid .NET and EXE file
        validate_file(buffer)?;

        // Initializes the default instance and injects the read buffer
        let mut clr = Self::default();
        clr.buffer = buffer;
        Ok(clr)
    }

    /// Sets the .NET runtime version to use.
    pub fn runtime_version(mut self, version: RuntimeVersion) -> Self {
        self.runtime_version = Some(version);
        self
    }

    /// Sets the application domain name to use.
    pub fn domain(mut self, domain_name: &str) -> Self {
        self.domain_name = Some(domain_name.to_string());
        self
    }

    /// Sets the arguments to pass to the .NET assembly's entry point.
    pub fn args(mut self, args: Vec<&str>) -> Self {
        self.args = Some(args.iter().map(|&s| s.to_string()).collect());
        self
    }

    /// Enables or disables output redirection.
    pub fn output(mut self) -> Self {
        self.redirect_output = true;
        self
    }

    /// Enables patching of the `System.Environment.Exit` method in `mscorlib`.
    pub fn exit(mut self) -> Self {
        self.patch_exit = true;
        self
    }

    /// Runs the .NET assembly by loading it into the application domain and invoking its entry point.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The output from the .NET assembly if executed successfully.
    /// * `Err(ClrError)` - If an error occurs during execution.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Create and configure a RustClr instance
    ///     let mut clr = RustClr::new(&buffer)?
    ///         .runtime_version(RuntimeVersion::V4)
    ///         .domain("CustomDomain")
    ///         .args(vec!["arg1", "arg2"])
    ///         .output();
    ///
    ///     // Run the .NET assembly and capture the output
    ///     let output = clr.run()?;
    ///     println!("Output: {}", output);
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn run(&mut self) -> Result<String> {
        // Prepare the CLR environment
        self.prepare()?;

        // Gets the current application domain
        let domain = self.get_app_domain()?;

        // Loads the .NET assembly specified by name
        let assembly = domain.load_name(&self.identity_assembly)?;

        // Prepares the parameters for the `Main` method
        let parameters = self.args.as_ref().map_or_else(
            || Ok(null_mut()),
            |args| create_safe_array_args(args.to_vec()),
        )?;

        // Retrieves the mscorlib library
        let mscorlib = domain.get_assembly(s!("mscorlib"))?;

        // Perform the patch in System.Environment.Exit (If Enabled)
        if self.patch_exit {
            self.patch_exit(&mscorlib)?;
        }

        let output = if self.redirect_output {
            // Redirecting output
            let mut output_manager = ClrOutput::new(&mscorlib);
            output_manager.redirect()?;

            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;
            output_manager.capture()?
        } else {
            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;
            String::new()
        };

        self.unload_domain()?;
        Ok(output)
    }

    /// Prepares the CLR environment by initializing the runtime and application domain.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the environment is successfully prepared.
    /// * `Err(ClrError)` - If any error occurs during the preparation process.
    fn prepare(&mut self) -> Result<()> {
        // Creates the MetaHost to access the available CLR versions
        let meta_host = self.create_meta_host()?;

        // Gets information about the specified (or default) runtime version
        let runtime_info = self.get_runtime_info(&meta_host)?;

        // Get ICLRAssemblyIdentityManager via GetProcAddress
        let addr = runtime_info.GetProcAddress(s!("GetCLRIdentityManager"))?;
        let GetCLRIdentityManager = unsafe { core::mem::transmute::<*mut c_void, CLRIdentityManagerType>(addr) };
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

    /// Patches the `System.Environment.Exit` method to avoid process termination.
    fn patch_exit(&self, mscorlib: &_Assembly) -> Result<()> {
        // Resolve System.Environment type and the Exit method
        let env = mscorlib.resolve_type(s!("System.Environment"))?;
        let exit = env.method(s!("Exit"))?;

        // Resolve System.Reflection.MethodInfo.MethodHandle property
        let method_info = mscorlib.resolve_type(s!("System.Reflection.MethodInfo"))?;
        let method_handle = method_info.property(s!("MethodHandle"))?;

        // Convert the Exit method into a COM IUnknown pointer
        let instance = exit
            .cast::<IUnknown>()
            .map_err(|_| ClrError::GenericError("Failed to cast to IUnknown"))?;

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
            return Err(ClrError::GenericError(
                "Failed to change memory protection to RWX",
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
            return Err(ClrError::GenericError(
                "Failed to restore memory protection",
            ));
        }

        Ok(())
    }

    /// Retrieves the current application domain.
    fn get_app_domain(&mut self) -> Result<_AppDomain> {
        self.app_domain.clone().ok_or(ClrError::NoDomainAvailable)
    }

    /// Creates an instance of `ICLRMetaHost`.
    fn create_meta_host(&self) -> Result<ICLRMetaHost> {
        CLRCreateInstance::<ICLRMetaHost>(&CLSID_CLRMETAHOST)
            .map_err(|e| ClrError::MetaHostCreationError(format!("{e}")))
    }

    /// Retrieves runtime information based on the selected .NET version.
    fn get_runtime_info(&self, meta_host: &ICLRMetaHost) -> Result<ICLRRuntimeInfo> {
        let runtime_version = self.runtime_version.unwrap_or(RuntimeVersion::V4);
        let version_wide = runtime_version.to_vec();
        let version = PCWSTR(version_wide.as_ptr());
        meta_host
            .GetRuntime::<ICLRRuntimeInfo>(version)
            .map_err(|e| ClrError::RuntimeInfoError(format!("{e}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    fn get_icor_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICorRuntimeHost> {
        runtime_info
            .GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    fn get_clr_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICLRuntimeHost> {
        runtime_info
            .GetInterface::<ICLRuntimeHost>(&CLSID_ICLR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))
    }

    /// Starts the CLR runtime using the provided runtime host.
    fn start_runtime(&self, iclr_runtime_host: &ICLRuntimeHost) -> Result<()> {
        if iclr_runtime_host.Start() != 0 {
            return Err(ClrError::RuntimeStartError);
        }
        Ok(())
    }

    /// Initializes the application domain with the specified name or uses the default domain.
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
    fn unload_domain(&self) -> Result<()> {
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

impl Drop for RustClr<'_> {
    fn drop(&mut self) {
        if let Some(cor_runtime_host) = &self.cor_runtime_host {
            cor_runtime_host.Stop();
        }
    }
}

/// Manages output redirection in the CLR.
pub struct ClrOutput<'a> {
    /// The `StringWriter` instance used to capture output.
    string_writer: Option<VARIANT>,

    /// Reference to the `mscorlib` assembly for creating types.
    mscorlib: &'a _Assembly,
}

impl<'a> ClrOutput<'a> {
    /// Creates a new [`ClrOutput`].
    ///
    /// # Arguments
    ///
    /// * `mscorlib` - An instance of the `_Assembly` representing `mscorlib`.
    ///
    /// # Returns
    ///
    /// * A new instance of [`ClrOutput`].
    pub fn new(mscorlib: &'a _Assembly) -> Self {
        Self {
            string_writer: None,
            mscorlib,
        }
    }

    /// Redirects standard output and error streams to a `StringWriter`.
    ///
    /// # Returns
    ///
    /// * `Ok(())` – If redirection succeeds.
    /// * `Err(ClrError)` – If an error occurs while setting the redirection.
    pub fn redirect(&mut self) -> Result<()> {
        let console = self.mscorlib.resolve_type(s!("System.Console"))?;
        let string_writer = self.mscorlib.create_instance(s!("System.IO.StringWriter"))?;

        // Invokes the methods
        console.invoke(
            s!("SetOut"),
            None,
            Some(vec![string_writer]),
            Invocation::Static,
        )?;
        
        console.invoke(
            s!("SetError"),
            None,
            Some(vec![string_writer]),
            Invocation::Static,
        )?;

        // Saves the StringWriter instance to retrieve the output later
        self.string_writer = Some(string_writer);
        Ok(())
    }

    /// Captures the content of the `StringWriter` as a `String`.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The captured output as a string if successful.
    /// * `Err(ClrError)` - If an error occurs while capturing the output.
    pub fn capture(&self) -> Result<String> {
        // Ensure that the StringWriter instance is available
        let mut instance = self.string_writer
            .ok_or(ClrError::GenericError("No StringWriter instance found"))?;

        // Resolve the 'ToString' method on the StringWriter type
        let string_writer = self.mscorlib.resolve_type(s!("System.IO.StringWriter"))?;
        let to_string = string_writer.method(s!("ToString"))?;

        // Invoke 'ToString' on the StringWriter instance
        let result = to_string.invoke(Some(instance), None)?;

        // Extract the BSTR from the result
        let bstr = unsafe { result.Anonymous.Anonymous.Anonymous.bstrVal };

        // Clean Variant
        unsafe { VariantClear(&mut instance as *mut _) };

        // Convert the BSTR to a UTF-8 String
        Ok(bstr.to_string())
    }
}

/// Represents a simplified interface to the CLR components without loading assemblies.
#[derive(Debug)]
pub struct RustClrEnv {
    /// .NET runtime version to use.
    pub runtime_version: RuntimeVersion,

    /// MetaHost for accessing CLR components.
    pub meta_host: ICLRMetaHost,

    /// Runtime information for the specified CLR version.
    pub runtime_info: ICLRRuntimeInfo,

    /// Host for the CLR runtime.
    pub cor_runtime_host: ICorRuntimeHost,

    /// Current application domain.
    pub app_domain: _AppDomain,
}

impl RustClrEnv {
    /// Creates a new `RustClrEnv` instance with the specified runtime version.
    ///
    /// # Arguments
    ///
    /// * `runtime_version` - The .NET runtime version to use.
    ///
    /// # Returns
    ///
    /// * `Ok(Self)` - If the components are initialized successfully.
    /// * `Err(ClrError)` - If initialization fails at any step.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use rustclr::{RustClrEnv, RuntimeVersion};
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     // Create a new RustClrEnv with a specific runtime version
    ///     let clr_env = RustClrEnv::new(Some(RuntimeVersion::V4))?;
    ///
    ///     println!("CLR initialized successfully.");
    ///     Ok(())
    /// }
    /// ```
    pub fn new(runtime_version: Option<RuntimeVersion>) -> Result<Self> {
        // Initialize MetaHost
        let meta_host = CLRCreateInstance::<ICLRMetaHost>(&CLSID_CLRMETAHOST)
            .map_err(|e| ClrError::MetaHostCreationError(format!("{e}")))?;

        // Initialize RuntimeInfo
        let version_str = runtime_version.unwrap_or(RuntimeVersion::V4).to_vec();
        let version = PCWSTR(version_str.as_ptr());

        let runtime_info = meta_host
            .GetRuntime::<ICLRRuntimeInfo>(version)
            .map_err(|e| ClrError::RuntimeInfoError(format!("{e}")))?;

        // Initialize CorRuntimeHost
        let cor_runtime_host = runtime_info
            .GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))?;

        if cor_runtime_host.Start() != 0 {
            return Err(ClrError::RuntimeStartError);
        }

        // Initialize AppDomain
        let uuid = uuid()
            .to_string()
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>();

        let app_domain = cor_runtime_host
            .CreateDomain(PCWSTR(uuid.as_ptr()), null_mut())
            .map_err(|_| ClrError::NoDomainAvailable)?;

        // Return the initialized instance
        Ok(Self {
            runtime_version: runtime_version.unwrap_or(RuntimeVersion::V4),
            meta_host,
            runtime_info,
            cor_runtime_host,
            app_domain,
        })
    }
}

impl Drop for RustClrEnv {
    fn drop(&mut self) {
        if let Err(e) = self.cor_runtime_host.UnloadDomain(
            self.app_domain
                .cast::<windows_core::IUnknown>()
                .map(|i| i.as_raw().cast())
                .unwrap_or(null_mut()),
        ) {
            dinvk::println!("Failed to unload AppDomain: {:?}", e);
        }

        self.cor_runtime_host.Stop();
    }
}

/// Represents the .NET runtime versions supported by RustClr.
#[derive(Debug, Clone, Copy)]
pub enum RuntimeVersion {
    /// .NET Framework 2.0, identified by version `v2.0.50727`.
    V2,

    /// .NET Framework 3.0, identified by version `v3.0`.
    V3,

    /// .NET Framework 4.0, identified by version `v4.0.30319`.
    V4,

    /// Represents an unknown or unsupported .NET runtime version.
    UNKNOWN,
}

impl RuntimeVersion {
    /// Converts the `RuntimeVersion` to a wide string representation as a `Vec<u16>`.
    ///
    /// # Returns
    ///
    /// A `Vec<u16>` containing the .NET runtime version as a null-terminated wide string.
    fn to_vec(self) -> Vec<u16> {
        let runtime_version = match self {
            RuntimeVersion::V2 => "v2.0.50727",
            RuntimeVersion::V3 => "v3.0",
            RuntimeVersion::V4 => "v4.0.30319",
            RuntimeVersion::UNKNOWN => "UNKNOWN",
        };

        runtime_version.encode_utf16().chain(Some(0)).collect::<Vec<u16>>()
    }
}

/// Represents a source of CLR data, which can come from a file path or an in-memory buffer.
#[derive(Debug, Clone)]
pub enum ClrSource<'a> {
    /// File indicated by a string representing the file path.
    File(&'a str),

    /// In-memory buffer containing the data.
    Buffer(&'a [u8]),
}

impl<'a> From<&'a str> for ClrSource<'a> {
    /// Converts a file path (`&'a str`) into a [`ClrSource::File`].
    fn from(file: &'a str) -> Self {
        ClrSource::File(file)
    }
}

impl<'a, const N: usize> From<&'a [u8; N]> for ClrSource<'a> {
    /// Converts a fixed-size byte array into a [`ClrSource::Buffer`].
    fn from(buffer: &'a [u8; N]) -> Self {
        ClrSource::Buffer(buffer)
    }
}

impl<'a> From<&'a [u8]> for ClrSource<'a> {
    /// Converts a byte slice into a [`ClrSource::Buffer`].
    fn from(buffer: &'a [u8]) -> Self {
        ClrSource::Buffer(buffer)
    }
}

impl<'a> From<&'a Vec<u8>> for ClrSource<'a> {
    /// Converts a [`Vec<u8>`] reference into a [`ClrSource::Buffer`].
    fn from(buffer: &'a Vec<u8>) -> Self {
        ClrSource::Buffer(buffer.as_slice())
    }
}
