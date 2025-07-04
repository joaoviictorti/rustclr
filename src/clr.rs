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
    data::NT_SUCCESS
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
    Invocation, Result, Variant,
    WinStr, create_safe_array_args,
    error::ClrError, uuid,
    file::{read_file, validate_file},
    host_control::RustClrControl
};

/// Represents a Rust interface to the Common Language Runtime (CLR).
///
/// This structure allows loading and executing .NET assemblies with specific runtime versions,
/// application domains, and arguments.
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
    /// Provides a default-initialized `RustClr`.
    ///
    /// # Returns
    ///
    /// * A default-initialized `RustClr`.
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
    ///
    /// # Arguments
    ///
    /// * `version` - The `RuntimeVersion` enum representing the .NET version.
    ///
    /// # Returns
    ///
    /// * Returns the modified `RustClr` instance.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Set a specific .NET runtime version
    ///     let clr = RustClr::new(&buffer)?
    ///         .runtime_version(RuntimeVersion::V4);
    ///
    ///     println!("Runtime version set successfully.");
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn runtime_version(mut self, version: RuntimeVersion) -> Self {
        self.runtime_version = Some(version);
        self
    }

    /// Sets the application domain name to use.
    ///
    /// # Arguments
    ///
    /// * `domain_name` - A string representing the name of the application domain.
    ///
    /// # Returns
    ///
    /// * Returns the modified `RustClr` instance.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Set a custom application domain name
    ///     let clr = RustClr::new(&buffer)?
    ///         .domain("CustomDomain");
    ///
    ///     println!("Domain set successfully.");
    ///     Ok(())
    /// }
    /// ```
    pub fn domain(mut self, domain_name: &str) -> Self {
        self.domain_name = Some(domain_name.to_string());
        self
    }

    /// Sets the arguments to pass to the .NET assembly's entry point.
    ///
    /// # Arguments
    ///
    /// * `args` - A vector of strings representing the arguments.
    ///
    /// # Returns
    ///
    /// * Returns the modified `RustClr` instance.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Pass arguments to the .NET assembly's entry point
    ///     let clr = RustClr::new(&buffer)?
    ///         .args(vec!["arg1", "arg2"]);
    ///
    ///     println!("Arguments set successfully.");
    ///     Ok(())
    /// }
    /// ```
    pub fn args(mut self, args: Vec<&str>) -> Self {
        self.args = Some(args.iter().map(|&s| s.to_string()).collect());
        self
    }

    /// Enables or disables output redirection.
    ///
    /// # Arguments
    ///
    /// * `redirect` - A boolean indicating whether to enable output redirection.
    ///
    /// # Returns
    ///
    /// * The modified `RustClr` instance with the updated output redirection setting.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Enable output redirection to capture console output
    ///     let clr = RustClr::new(&buffer)?
    ///         .output();
    ///
    ///     println!("Output redirection enabled.");
    ///     Ok(())
    /// }
    /// ```
    pub fn output(mut self) -> Self {
        self.redirect_output = true;
        self
    }

    /// Enables patching of the `System.Environment.Exit` method in `mscorlib`.
    ///
    /// # Returns
    ///
    /// * The modified `RustClr` instance with the patch setting enabled.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Enable patching of Environment.Exit
    ///     let clr = RustClr::new(&buffer)?
    ///         .exit();
    ///
    ///     println!("Environment.Exit will be patched to prevent termination.");
    ///     Ok(())
    /// }
    /// ```
    pub fn exit(mut self) -> Self {
        self.patch_exit = true;
        self
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

        // Create a stream for the in-memory assembly
        let iclr_assembly = ICLRAssemblyIdentityManager::from_raw(ptr)?;
        let stream = unsafe { SHCreateMemStream(self.buffer.as_ptr(), self.buffer.len() as u32) };

        // Get the identity string from the stream
        self.identity_assembly = iclr_assembly.get_identity_stream(stream, 0)?;

        // Creates the `ICLRuntimeHost`
        let iclr_runtime_host = self.get_clr_runtime_host(&runtime_info)?;

        // Create and register IHostControl with custom assembly and identity
        let host_control: IHostControl = RustClrControl::new(self.buffer, &self.identity_assembly).into();
        iclr_runtime_host.SetHostControl(&host_control)?;

        // Checks if the runtime is started
        if runtime_info.IsLoadable().is_ok() && !runtime_info.is_started() {
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

        // Loads the .NET assembly specified by the buffer
        let assembly = domain.load_name(&self.identity_assembly)?;

        // Prepares the parameters for the `Main` method
        let parameters = self.args.as_ref().map_or_else(
            || Ok(null_mut()),
            |args| create_safe_array_args(args.to_vec()),
        )?;

        // Loads the mscorlib library
        let mscorlib = domain.get_assembly(s!("mscorlib"))?;

        // If the exit patch is enabled, perform the patch in System.Environment.Exit
        if self.patch_exit {
            self.patch_exit(&mscorlib)?;
        }

        // Redirects output if enabled
        let output = if self.redirect_output {
            // Redirecting output
            let mut output_manager = ClrOutput::new(&mscorlib);
            output_manager.redirect()?;

            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;

            // Restores output if redirected
            let output = output_manager.capture()?;
            output_manager.restore()?;
            output
        } else {
            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;

            // Empty output
            String::new()
        };

        // Unload Domain
        self.unload_domain()?;
        Ok(output)
    }

    /// Patches the `System.Environment.Exit` method in `mscorlib` to avoid process termination.
    ///
    /// # Arguments
    ///
    /// * `mscorlib` - The `_Assembly` object representing the loaded `mscorlib.dll`.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the patch was applied successfully.
    /// * `Err(ClrError)` - If any COM error or memory protection error occurs during the process.
    fn patch_exit(&self, mscorlib: &_Assembly) -> Result<()> {
        // Resolve System.Environment type and the Exit method
        let env = mscorlib.resolve_type(s!("System.Environment"))?;
        let exit = env.method(s!("Exit"))?;

        // Resolve System.Reflection.MethodInfo.MethodHandle property
        let method_info = mscorlib.resolve_type(s!("System.Reflection.MethodInfo"))?;
        let method_handle = method_info.property(s!("MethodHandle"))?;

        // Convert the Exit method into a COM IUnknown pointer to pass into MethodHandle
        let instance = exit
            .cast::<IUnknown>()
            .map_err(|_| ClrError::GenericError("Failed to cast to IUnknown"))?;

        // Call MethodHandle.get_Value(instance) to retrieve the RuntimeMethodHandle
        let method_handle_exit = method_handle.value(Some(instance.to_variant()), None)?;

        // Resolve System.RuntimeMethodHandle.GetFunctionPointer
        let runtime_method = mscorlib.resolve_type(s!("System.RuntimeMethodHandle"))?;
        let get_function_pointer = runtime_method.method(s!("GetFunctionPointer"))?;

        // Invoke GetFunctionPointer(handle) to get the native address of Environment.Exit
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

        // Overwrite first byte with RET (0xC3) to effectively no-op the function
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
    ///
    /// # Returns
    ///
    /// * `Ok(_AppDomain)` - If the application domain is available.
    /// * `Err(ClrError)` - If no application domain is available.
    fn get_app_domain(&mut self) -> Result<_AppDomain> {
        self.app_domain.clone().ok_or(ClrError::NoDomainAvailable)
    }

    /// Creates an instance of `ICLRMetaHost`.
    ///
    /// # Returns
    ///
    /// * `Ok(ICLRMetaHost)` - If the instance is created successfully.
    /// * `Err(ClrError)` - If the instance creation fails.
    fn create_meta_host(&self) -> Result<ICLRMetaHost> {
        CLRCreateInstance::<ICLRMetaHost>(&CLSID_CLRMETAHOST)
            .map_err(|e| ClrError::MetaHostCreationError(format!("{e}")))
    }

    /// Retrieves runtime information based on the selected .NET version.
    ///
    /// # Arguments
    ///
    /// * `meta_host` - Reference to the `ICLRMetaHost` instance.
    ///
    /// # Returns
    ///
    /// * `Ok(ICLRRuntimeInfo)` - If runtime information is retrieved successfully.
    /// * `Err(ClrError)` - If the retrieval fails.
    fn get_runtime_info(&self, meta_host: &ICLRMetaHost) -> Result<ICLRRuntimeInfo> {
        let runtime_version = self.runtime_version.unwrap_or(RuntimeVersion::V4);
        let version_wide = runtime_version.to_vec();
        let version = PCWSTR(version_wide.as_ptr());
        meta_host
            .GetRuntime::<ICLRRuntimeInfo>(version)
            .map_err(|e| ClrError::RuntimeInfoError(format!("{e}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    ///
    /// # Arguments
    ///
    /// * `runtime_info` - Reference to the `ICLRRuntimeInfo` instance.
    ///
    /// # Returns
    ///
    /// * `Ok(ICorRuntimeHost)` - If the interface is obtained successfully.
    /// * `Err(ClrError)` - If the retrieval fails.
    fn get_icor_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICorRuntimeHost> {
        runtime_info
            .GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))
    }

    /// Gets the runtime host interface from the provided runtime information.
    ///
    /// # Arguments
    ///
    /// * `runtime_info` - Reference to the `ICLRRuntimeInfo` instance.
    ///
    /// # Returns
    ///
    /// * `Ok(ICorRuntimeHost)` - If the interface is obtained successfully.
    /// * `Err(ClrError)` - If the retrieval fails.
    fn get_clr_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICLRuntimeHost> {
        runtime_info
            .GetInterface::<ICLRuntimeHost>(&CLSID_ICLR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))
    }

    /// Starts the CLR runtime using the provided runtime host.
    ///
    /// # Arguments
    ///
    /// * `iclr_runtime_host` - Reference to the `ICorRuntimeHost` instance.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the runtime starts successfully.
    /// * `Err(ClrError)` - If the runtime fails to start.
    fn start_runtime(&self, iclr_runtime_host: &ICLRuntimeHost) -> Result<()> {
        if iclr_runtime_host.Start() != 0 {
            return Err(ClrError::RuntimeStartError);
        }

        Ok(())
    }

    /// Initializes the application domain with the specified name or uses the default domain.
    ///
    /// # Arguments
    ///
    /// * `cor_runtime_host` - Reference to the `ICorRuntimeHost` instance.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the application domain is successfully initialized.
    /// * `Err(ClrError)` - If the initialization fails.
    fn init_app_domain(&mut self, cor_runtime_host: &ICorRuntimeHost) -> Result<()> {
        // Creates the application domain based on the specified name or uses the default domain
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
    ///
    /// This method is used to properly unload a custom AppDomain created by `RustClr`.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the AppDomain is unloaded or not present.
    /// * `Err(ClrError)` - If unloading the domain fails.
    fn unload_domain(&self) -> Result<()> {
        if let (Some(cor_runtime_host), Some(app_domain)) =
            (&self.cor_runtime_host, &self.app_domain)
        {
            // Attempt to unload the AppDomain, log error if it fails
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

/// Implements the `Drop` trait to release memory when `RustClr` goes out of scope.
impl Drop for RustClr<'_> {
    fn drop(&mut self) {
        if let Some(cor_runtime_host) = &self.cor_runtime_host {
            // Attempt to stop the CLR runtime
            cor_runtime_host.Stop();
        }
    }
}

/// Manages output redirection in the CLR by using a `StringWriter`.
///
/// This struct handles the redirection of standard output and error streams
/// to a `StringWriter` instance, enabling the capture of output produced
/// by the .NET code.
pub struct ClrOutput<'a> {
    /// The `StringWriter` instance used to capture output.
    string_writer: Option<VARIANT>,

    /// Reference to the `mscorlib` assembly for creating types.
    mscorlib: &'a _Assembly,
}

impl<'a> ClrOutput<'a> {
    /// Creates a new `ClrOutput`.
    ///
    /// # Arguments
    ///
    /// * `mscorlib` - An instance of the `_Assembly` representing `mscorlib`.
    ///
    /// # Returns
    ///
    /// * A new instance of `ClrOutput`.
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
    /// * `Ok(())` - If the redirection is successful.
    /// * `Err(ClrError)` - If an error occurs while attempting to redirect the streams.
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

    /// Restores the original standard output and error streams.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the restoration is successful.
    /// * `Err(ClrError)` - If an error occurs while restoring the streams.
    pub fn restore(&mut self) -> Result<()> {
        let console = self.mscorlib.resolve_type(s!("System.Console"))?;
        console.method_signature(s!("Void InitializeStdOutError(Boolean)"))?
            .invoke(
                None,
                Some(crate::create_safe_args(vec![true.to_variant()])?),
            )?;

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
    /// ```ignore
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
        // Attempt to unload the AppDomain, log error if it fails
        if let Err(e) = self.cor_runtime_host.UnloadDomain(
            self.app_domain
                .cast::<windows_core::IUnknown>()
                .map(|i| i.as_raw().cast())
                .unwrap_or(null_mut()),
        ) {
            dinvk::println!("Failed to unload AppDomain: {:?}", e);
        }

        // Attempt to stop the CLR runtime
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

        runtime_version
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>()
    }
}

/// Represents the COFF data source, which can be a file or a memory buffer.
#[derive(Debug, Clone)]
pub enum ClrSource<'a> {
    /// COFF file indicated by a string representing the file path.
    File(&'a str),

    /// Memory buffer containing COFF data.
    Buffer(&'a [u8]),
}

impl<'a> From<&'a str> for ClrSource<'a> {
    /// Converts a file path (`&'a str`) to a COFF source (`ClrSource::File`).
    ///
    /// # Arguments
    ///
    /// * `file` - Path of the COFF file.
    ///
    /// # Returns
    ///
    /// * The input string will be treated as the path of a COFF file.
    fn from(file: &'a str) -> Self {
        ClrSource::File(file)
    }
}

impl<'a, const N: usize> From<&'a [u8; N]> for ClrSource<'a> {
    /// Converts a fixed-size byte array (`&[u8; N]`) to a COFF source (`ClrSource::Buffer`).
    ///
    /// # Arguments
    ///
    /// * `buffer` - A fixed-size byte array representing the COFF file data.
    ///
    /// # Returns
    ///
    /// * The input byte array will be treated as a COFF buffer in memory.
    fn from(buffer: &'a [u8; N]) -> Self {
        ClrSource::Buffer(buffer)
    }
}

impl<'a> From<&'a [u8]> for ClrSource<'a> {
    /// Converts a byte slice (`&[u8]`) to a COFF source (`ClrSource::Buffer`).
    ///
    /// # Arguments
    ///
    /// * `buffer` - A byte slice representing the COFF file data.
    ///
    /// # Returns
    ///
    /// * The input byte slice will be treated as a COFF buffer in memory.
    fn from(buffer: &'a [u8]) -> Self {
        ClrSource::Buffer(buffer)
    }
}

impl<'a> From<&'a Vec<u8>> for ClrSource<'a> {
    /// Converts a byte slice (`&Vec<u8>`) to a COFF source (`ClrSource::Buffer`).
    ///
    /// # Arguments
    ///
    /// * `buffer` - A byte slice representing the COFF file data.
    ///
    /// # Returns
    ///
    /// * The input byte slice will be treated as a COFF buffer in memory.
    fn from(buffer: &'a Vec<u8>) -> Self {
        ClrSource::Buffer(buffer.as_slice())
    }
}
