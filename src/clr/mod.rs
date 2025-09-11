use core::ptr::null_mut;
use alloc::{
    boxed::Box,
    format,
    string::{String, ToString},
    vec,
    vec::Vec,
};

use obfstr::obfstr as s;
use windows_core::{Interface, PCWSTR};
use windows_sys::Win32::System::Variant::{
    VARIANT, 
    VariantClear
};

use self::file::{read_file, validate_file};
use self::runtime::{RustClrRuntime, uuid};
use super::com::*;
use super::error::ClrError;
use super::create_safe_array_args;
use super::{
    Result, 
    Variant,
    ComString,
};

mod hosting;
mod runtime;
mod file;

pub use runtime::RuntimeVersion;

/// Represents a Rust interface to the Common Language Runtime (CLR).
#[derive(Default, Debug, Clone)]
pub struct RustClr<'a> {
    /// Encapsulates all runtime-related state and preparation logic.
    runtime: RustClrRuntime<'a>,

    /// Flag to indicate if output redirection is enabled.
    redirect_output: bool,

    /// Whether to patch `System.Environment.Exit` to prevent the process from terminating.
    patch_exit: bool,

    /// Arguments to pass to the .NET assembly's `Main` method.
    args: Option<Vec<String>>,
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
    /// * `Err(ClrError)` - If the buffer validation fails.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    /// 
    /// // Load a sample .NET assembly into a buffer
    /// let buffer = fs::read("examples/sample.exe")?;
    /// let mut clr = RustClr::new(&buffer)?
    ///     .runtime_version(RuntimeVersion::V4)
    ///     .domain("CustomDomain")
    ///     .args(vec!["arg1", "arg2"])
    ///     .output();
    /// 
    /// let output = clr.run()?;
    /// println!("Output: {}", output);
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

        Ok(Self {
            runtime: RustClrRuntime::new(buffer),
            redirect_output: false,
            patch_exit: false,
            args: None,
        })
    }

    /// Sets the .NET runtime version to use.
    pub fn runtime_version(mut self, version: RuntimeVersion) -> Self {
        self.runtime.runtime_version = Some(version);
        self
    }

    /// Sets the application domain name to use.
    pub fn domain(mut self, domain_name: &str) -> Self {
        self.runtime.domain_name = Some(domain_name.to_string());
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

    /// Loads the .NET assembly and runs its entry point.
    ///
    /// # Returns
    ///
    /// * `Ok(String)` - The output from the .NET assembly if executed successfully.
    /// * `Err(ClrError)` - If an error occurs during execution.
    /// 
    /// # Example
    /// 
    /// ```rust,ignore
    /// use rustclr::{RustClr, RuntimeVersion};
    /// use std::fs;
    ///
    /// // Load a sample .NET assembly into a buffer
    /// let buffer = fs::read("examples/sample.exe")?;
    /// let mut clr = RustClr::new(&buffer)?
    ///     .runtime_version(RuntimeVersion::V4)
    ///     .domain("CustomDomain")
    ///     .args(vec!["arg1", "arg2"])
    ///     .output();
    ///
    /// let output = clr.run()?;
    /// println!("Output: {}", output);
    /// ```
    pub fn run(&mut self) -> Result<String> {
        // Prepare the CLR environment
        self.runtime.prepare()?;

        // Gets the current application domain
        let domain = self.runtime.get_app_domain()?;

        // Loads the .NET assembly specified by name
        let assembly = domain.load_name(&self.runtime.identity_assembly)?;

        // Prepares the parameters for the `Main` method
        let parameters = self.args.as_ref().map_or_else(
            || Ok(null_mut()),
            |args| create_safe_array_args(args.to_vec()),
        )?;

        // Retrieves the mscorlib library
        let mscorlib = domain.get_assembly(s!("mscorlib"))?;

        // Disables Environment.Exit if patching is enabled.
        if self.patch_exit {
            runtime::patch_exit(&mscorlib)?;
        }

        // Optional output redirection
        let output_manager = if self.redirect_output {
            let mut manager = ClrOutput::new(&mscorlib);
            manager.redirect()?;
            Some(manager)
        } else {
            None
        };

        // Invokes the `Main` method of the assembly
        assembly.run(parameters)?;

        // Optionally capture redirected output
        let output = match output_manager {
            Some(manager) => manager.capture()?,
            None => String::new(),
        };

        self.runtime.unload_domain()?;
        Ok(output)
    }
}

impl Drop for RustClr<'_> {
    fn drop(&mut self) {
        if let Some(cor_runtime_host) = &self.runtime.cor_runtime_host {
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

/// Specifies the invocation type for a method.
pub enum Invocation {
    /// Indicates that the method to invoke is static.
    Static,

    /// Indicates that the method to invoke is an instance method.
    Instance,
}

/// Represents a source of CLR data.
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

