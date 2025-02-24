use windows_core::PCWSTR;
use windows_sys::Win32::System::Variant::VARIANT;
use crate::{
    WinStr, 
    Invocation,
    error::ClrError, 
    file::validate_file, 
    create_safe_array_args,
};
use crate::com::{
    CLRCreateInstance, 
    CLSID_CLRMETAHOST, 
    CLSID_COR_RUNTIME_HOST
};
use crate::data::{
    _AppDomain, ICLRMetaHost, 
    ICLRRuntimeInfo, ICorRuntimeHost, 
    _Assembly 
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

impl<'a> Default for RustClr<'a> {
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
            domain_name: None,
            args: None, 
            app_domain: None,
            cor_runtime_host: None
        }
    }
}

impl<'a> RustClr<'a> {
    /// Creates a new `RustClr` instance with the specified assembly buffer.
    /// 
    /// # Arguments
    /// 
    /// * `buffer` - A reference to a byte slice representing the .NET assembly.
    /// 
    /// # Returns
    /// 
    /// * `Ok(Self)` - If the buffer is valid and the `RustClr` instance is created successfully.
    /// * `Err(ClrError)` - If the buffer validation fails (e.g., not a valid .NET assembly).
    /// 
    /// # Examples
    /// 
    /// ```ignore
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
    pub fn new(buffer: &'a [u8]) -> Result<Self, ClrError> {
        // Checks if it is a valid .NET and EXE file
        validate_file(buffer)?;

        Ok(Self { 
            buffer, 
            redirect_output: false,
            runtime_version: None,
            domain_name: None, 
            args: None, 
            app_domain: None,
            cor_runtime_host: None
        })
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
    /// # Examples
    /// 
    /// ```ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Set a specific .NET runtime version
    ///     let clr = RustClr::new(&buffer)?
    ///         .with_runtime_version(RuntimeVersion::V4);
    ///
    ///     println!("Runtime version set successfully.");
    /// 
    ///     Ok(())
    /// }
    /// ```
    pub fn with_runtime_version(mut self, version: RuntimeVersion) -> Self {
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
    /// # Examples
    /// 
    /// ```ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Set a custom application domain name
    ///     let clr = RustClr::new(&buffer)?
    ///         .with_domain("CustomDomain");
    ///
    ///     println!("Domain set successfully.");
    ///     Ok(())
    /// }
    /// ```
    pub fn with_domain(mut self, domain_name: &str) -> Self {
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
    /// ```ignore
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Pass arguments to the .NET assembly's entry point
    ///     let clr = RustClr::new(&buffer)?
    ///         .with_args(vec!["arg1", "arg2"]);
    ///
    ///     println!("Arguments set successfully.");
    ///     Ok(())
    /// }
    /// ```
    pub fn with_args(mut self, args: Vec<&str>) -> Self {
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
    /// # Examples
    ///
    /// ```rust
    /// use rustclr::RustClr;
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Enable output redirection to capture console output
    ///     let clr = RustClr::new(&buffer)?
    ///         .with_output_redirection(true);
    ///
    ///     println!("Output redirection enabled.");
    ///     Ok(())
    /// }
    /// ```
    pub fn with_output_redirection(mut self, redirect: bool) -> Self {
        self.redirect_output = redirect;
        self
    }

    /// Prepares the CLR environment by initializing the runtime and application domain.
    /// 
    /// # Returns
    /// 
    /// * `Ok(())` - If the environment is successfully prepared.
    /// * `Err(ClrError)` - If any error occurs during the preparation process.
    fn prepare(&mut self) -> Result<(), ClrError> {
        // Creates the MetaHost to access the available CLR versions
        let meta_host = self.create_meta_host()?;

        // Gets information about the specified (or default) runtime version
        let runtime_info = self.get_runtime_info(&meta_host)?;

        // Creates the runtime host
        let cor_runtime_host = self.get_runtime_host(&runtime_info)?;

        // Checks if the runtime is started
        if runtime_info.IsLoadable().is_ok() && !runtime_info.is_started() {
            // Starts the CLR runtime
            self.start_runtime(&cor_runtime_host)?;
        }

        // Initializes the specified application domain or the default
        self.init_app_domain(&cor_runtime_host)?;

        // Saves the runtime host for future use
        self.cor_runtime_host = Some(cor_runtime_host);

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
    /// ```ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    ///
    /// fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     let buffer = fs::read("examples/sample.exe")?;
    ///
    ///     // Create and configure a RustClr instance
    ///     let mut clr = RustClr::new(&buffer)?
    ///         .with_runtime_version(RuntimeVersion::V4)
    ///         .with_domain("CustomDomain")
    ///         .with_args(vec!["arg1", "arg2"])
    ///         .with_output_redirection(true);
    ///
    ///     // Run the .NET assembly and capture the output
    ///     let output = clr.run()?;
    ///     println!("Output: {}", output);
    /// 
    ///     Ok(())
    /// }
    /// ```
    pub fn run(&mut self) -> Result<String, ClrError> {
        // Prepare the CLR environment
        self.prepare()?;

        // Gets the current application domain
        let domain = self.get_app_domain()?;

        // Loads the .NET assembly specified by the buffer
        let assembly = domain.load_assembly(self.buffer)?;

        // Prepares the parameters for the `Main` method
        let parameters = self.args.as_ref().map_or_else(
            || Ok(std::ptr::null_mut()),
            |args| create_safe_array_args(args.to_vec())
        )?;

        // Redirects output if enabled
        let output = if self.redirect_output {
            // Loads the mscorlib library for output redirection
            let mscorlib = domain.load_lib("mscorlib")?;
            let mut output_manager = ClrOutput::new(&mscorlib);
            
            // Redirecting output
            output_manager.redirect()?;

            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;

            // Captures and restores output if redirected
            let result = output_manager.capture()?;
            output_manager.restore()?;
            result
        } else {
            // Invokes the `Main` method of the assembly
            assembly.run(parameters)?;

            // Empty output
            String::new()
        };

        Ok(output)
    }

    /// Retrieves the current application domain.
    /// 
    /// # Returns
    /// 
    /// * `Ok(_AppDomain)` - If the application domain is available.
    /// * `Err(ClrError)` - If no application domain is available.
    fn get_app_domain(&mut self) -> Result<_AppDomain, ClrError> {
        self.app_domain.clone().ok_or(ClrError::NoDomainAvailable)
    }

    /// Creates an instance of `ICLRMetaHost`.
    /// 
    /// # Returns
    /// 
    /// * `Ok(ICLRMetaHost)` - If the instance is created successfully.
    /// * `Err(ClrError)` - If the instance creation fails.
    fn create_meta_host(&self) -> Result<ICLRMetaHost, ClrError> {
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
    fn get_runtime_info(&self, meta_host: &ICLRMetaHost) -> Result<ICLRRuntimeInfo, ClrError> {
        let runtime_version = self.runtime_version.unwrap_or(RuntimeVersion::V4);
        let version_wide = runtime_version.to_vec();
        let version = PCWSTR(version_wide.as_ptr());

        meta_host.GetRuntime::<ICLRRuntimeInfo>(version)
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
    fn get_runtime_host(&self, runtime_info: &ICLRRuntimeInfo) -> Result<ICorRuntimeHost, ClrError> {
        runtime_info.GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))
    }

    /// Starts the CLR runtime using the provided runtime host.
    /// 
    /// # Arguments
    /// 
    /// * `cor_runtime_host` - Reference to the `ICorRuntimeHost` instance.
    /// 
    /// # Returns
    /// 
    /// * `Ok(())` - If the runtime starts successfully.
    /// * `Err(ClrError)` - If the runtime fails to start.
    fn start_runtime(&self, cor_runtime_host: &ICorRuntimeHost) -> Result<(), ClrError> {
        if cor_runtime_host.Start() != 0 {
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
    fn init_app_domain(&mut self, cor_runtime_host: &ICorRuntimeHost) -> Result<(), ClrError> {
        // Creates the application domain based on the specified name or uses the default domain
        let app_domain = if let Some(domain_name) = &self.domain_name {
            let wide_domain_name = domain_name.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
            cor_runtime_host.CreateDomain(PCWSTR(wide_domain_name.as_ptr()), std::ptr::null_mut())?
        } else {
            cor_runtime_host.GetDefaultDomain()?
        };

        // Saves the created application domain
        self.app_domain = Some(app_domain);

        Ok(())
    }
}

/// Implements the `Drop` trait to release memory when `RustClr` goes out of scope.
impl<'a> Drop for RustClr<'a> {
    fn drop(&mut self) {
        // Ends the CLR runtime
        if let Some(ref cor_runtime_host) = self.cor_runtime_host {
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
    /// Original standard output stream.
    out: Option<VARIANT>,

    /// Original standard error stream.
    error: Option<VARIANT>,

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
            out: None,
            error: None,
            string_writer: None,
            mscorlib
        }
    }

    /// Redirects standard output and error streams to a `StringWriter`.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the redirection is successful.
    /// * `Err(ClrError)` - If an error occurs while attempting to redirect the streams.
    pub fn redirect(&mut self) -> Result<(), ClrError> {
        let console = self.mscorlib.resolve_type("System.Console")?;
        let string_writer =  self.mscorlib.create_instance("System.IO.StringWriter")?;

        // Save the original output and error streams
        self.out = Some(console.invoke("get_Out", None, None, Invocation::Static)?);
        self.error = Some(console.invoke("get_Error", None, None, Invocation::Static)?);

        // Invokes the methods
        console.invoke("SetOut", None, Some(vec![string_writer]), Invocation::Static)?;
        console.invoke("SetError", None, Some(vec![string_writer]), Invocation::Static)?;

        self.string_writer = Some(string_writer);

        Ok(())
    }

    /// Restores the original standard output and error streams.
    ///
    /// # Returns
    ///
    /// * `Ok(())` - If the restoration is successful.
    /// * `Err(ClrError)` - If an error occurs while restoring the streams.
    pub fn restore(&mut self) -> Result<(), ClrError> {
        let console =  self.mscorlib.resolve_type("System.Console")?;

        if let Some(out) = self.out.take() {
            console.invoke("SetOut", None, Some(vec![out]), Invocation::Static)?;
        }

        if let Some(error) = self.error.take() {
            console.invoke("SetError", None, Some(vec![error]), Invocation::Static)?;
        }

        Ok(())
    }

    /// Captures the content of the `StringWriter` as a `String`.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The captured output as a string if successful.
    /// * `Err(ClrError)` - If an error occurs while capturing the output.
    pub fn capture(&self) -> Result<String, ClrError> {
        // Ensure that the StringWriter instance is available
        let instance = self.string_writer.ok_or(ClrError::ErrorClr("No StringWriter instance found"))?;
        
        // Resolve the 'ToString' method on the StringWriter type
        let string_writer = self.mscorlib.resolve_type("System.IO.StringWriter")?;
        let to_string = string_writer.method("ToString")?;
        
        // Invoke 'ToString' on the StringWriter instance
        let result = to_string.invoke(Some(instance), None)?;

        // Extract the BSTR from the result
        let bstr = unsafe { result.Anonymous.Anonymous.Anonymous.bstrVal };

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
    pub fn new(runtime_version: Option<RuntimeVersion>) -> Result<Self, ClrError> {
        // Initialize MetaHost
        let meta_host = CLRCreateInstance::<ICLRMetaHost>(&CLSID_CLRMETAHOST)
            .map_err(|e| ClrError::MetaHostCreationError(format!("{e}")))?;

        // Initialize RuntimeInfo
        let version_str = runtime_version.unwrap_or(RuntimeVersion::V4).to_vec();
        let version = PCWSTR(version_str.as_ptr());

        let runtime_info = meta_host.GetRuntime::<ICLRRuntimeInfo>(version)
            .map_err(|e| ClrError::RuntimeInfoError(format!("{e}")))?;

        // Initialize CorRuntimeHost
        let cor_runtime_host = runtime_info.GetInterface::<ICorRuntimeHost>(&CLSID_COR_RUNTIME_HOST)
            .map_err(|e| ClrError::RuntimeHostError(format!("{e}")))?;
        
        if cor_runtime_host.Start() != 0 {
            return Err(ClrError::RuntimeStartError);
        }

        // Initialize AppDomain
        let app_domain = cor_runtime_host.GetDefaultDomain()
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
