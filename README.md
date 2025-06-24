# rustclr 🦀

![Rust](https://img.shields.io/badge/made%20with-Rust-red)
![crate](https://img.shields.io/crates/v/rustclr.svg)
![docs](https://docs.rs/rustclr/badge.svg)
![Forks](https://img.shields.io/github/forks/joaoviictorti/rustclr)
![Stars](https://img.shields.io/github/stars/joaoviictorti/rustclr)
![License](https://img.shields.io/github/license/joaoviictorti/rustclr)

`rustclr` is a powerful library for hosting the Common Language Runtime (CLR) and executing .NET binaries directly with Rust, among other operations.

## Features

- ✅ Supports `#[no_std]` environments (with `alloc`).
- ✅ Compatible with `x64` architecture.
- ✅ Run .NET binaries in memory with full control over runtime configurations.
- ✅ Fine-grained control over the CLR environment and runtime initialization.
- ✅ Configure output redirection to capture .NET program output.
- ✅ Patch `System.Environment.Exit()` to prevent .NET from terminating the Rust host process.

## Getting started

Add `rustclr` to your project by updating your `Cargo.toml`:
```bash
cargo add rustclr
```

## Usage

### Running a .NET Assembly with Configured Flags

The following flags provide full control over your CLR environment and the execution of your .NET assemblies:

- **`.runtime_version(RuntimeVersion::V4)`**: Sets the .NET runtime version (e.g., RuntimeVersion::V2, RuntimeVersion::V3, RuntimeVersion::V4). This flag ensures that the assembly runs with the specified CLR version.
- **`.output`**: Redirects the output from the .NET assembly's console to the Rust environment, capturing all console output.
- **`.domain("DomainName")`**: Sets a custom AppDomain name, which is useful for isolating different .NET assemblies.
- **`.args(vec!["arg1", "arg2"])`**: Passes arguments to the .NET application, useful for parameterized entry points in the assembly.
- **`.exit`**: This prevents calls to `System.Environment.Exit()` within the .NET assembly from terminating the host process (your Rust program). Instead, control is maintained on the Rust side, and the .run() method returns normally even if the assembly attempts to terminate the process.
  
Using `rustclr` to load and execute a .NET assembly, redirect its output and customize the CLR runtime environment.

```rs
use std::fs;
use rustclr::{RustClr, RuntimeVersion};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load a sample .NET assembly into a buffer
    let buffer = fs::read("examples/sample.exe")?;

    // Create and configure a RustClr instance with runtime version and output redirection
    let output = RustClr::new(&buffer)?
        .runtime_version(RuntimeVersion::V4) // Specify .NET runtime version
        .output() // Redirect output to capture it in Rust
        .domain("CustomDomain") // Optionally set a custom application domain
        .exit() // Patch Environment.Exit to prevent host process termination
        .args(vec!["arg1", "arg2"]) // Pass arguments to the .NET assembly's entry point
        .run()?; // Execute the assembly

    println!("Captured output: {}", output);

    Ok(())
}
```

### Running PowerShell Commands

`rustclr` also provides a high-level interface to execute `PowerShell` commands from Rust using the built-in .NET `System.Management.Automation` namespace.

```rs
use std::error::Error;
use rustclr::PowerShell;

fn main() -> Result<(), Box<dyn Error>> {
    let pwsh = PowerShell::new()?;
    print!("{}", pwsh.execute("Get-Process | Select-Object -First 3")?);
    print!("{}", pwsh.execute("whoami")?);
    
    Ok(())
}
```

### Configuration with RustClrEnv and ClrOutput

For more fine-grained control, rustclr provides the `RustClrEnv` and `ClrOutput` components:

- **`RustClrEnv`**: Allows for low-level customization and initialization of the .NET runtime environment, which is useful if you need to manually control the CLR version, MetaHost, runtime information, and application domain. This struct provides an alternative way to initialize a CLR environment without executing an assembly immediately.
```rs
use rustclr::{RustClrEnv, RuntimeVersion};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a new environment for .NET with a specific runtime version
    let clr_env = RustClrEnv::new(Some(RuntimeVersion::V4))?;
    println!("CLR environment initialized successfully with version {:?}", clr_env.runtime_version);

    Ok(())
}
```

- **`ClrOutput`**: Manages redirection of standard output and error streams from .NET to Rust. This is especially useful if you need to capture and process all output produced by .NET code within a Rust environment.
```rs
use rustclr::{
    RustClrEnv, ClrOutput, 
    Invocation, Variant
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create and initialize the CLR environment
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.load_lib("mscorlib")?;
    let console = mscorlib.resolve_type("System.Console")?;

    // Set up output redirection
    let mut clr_output = ClrOutput::new(&mscorlib);
    clr_output.redirect()?;

    // Prepare the arguments
    let args = vec!["Hello World".to_variant()];

    // Invoke the WriteLine method
    console.invoke("WriteLine", None, Some(args), Invocation::Static)?;

    // Restore the original output and capture redirected content
    clr_output.restore()?;
    let output = clr_output.capture()?;

    print!("{output}");

    Ok(())
}
```

## Additional Resources

For more examples, check the [examples](/examples) folder in the repository.

## CLI

`rustclr` also includes a command-line interface (CLI) for running .NET assemblies with various configuration options. Below is a description of the available flags and usage examples.

The CLI accepts the following options:

- **`-f, --file`**: Specifies the path to the .NET assembly file to be executed (required).
- **`-i, --inputs`**: Provides string arguments to be passed to the .NET program's entry point. This flag can be repeated to add multiple arguments.
- **`-r, --runtime-version`**: Sets the .NET runtime version to use. Accepted values include `"v2"`, `"v3"`, and `"v4"`. Defaults to `"v4"`.
- **`-d, --domain`**: Allows setting a custom name for the application domain (optional).

### Example Command

```powershell
clr.exe -f Rubeus.exe -i "triage" -i "/consoleoutfile:C:\Path" -r v4 -d "CustomDomain"
```

### CLI Help

```
Host CLR and run .NET binaries using Rust

Usage: clr.exe [OPTIONS] --file <FILE>

Options:
  -f, --file <FILE>                        Path to the .NET assembly file
  -i, --inputs <INPUTS>                    String arguments for the .NET program
  -r, --runtime-version <RUNTIME_VERSION>  Specify .NET runtime version [default: v4]
  -d, --domain <DOMAIN>                    Set custom application domain name
  -h, --help                               Print help
  -V, --version                            Print version
```

## References

I want to express my gratitude to these projects that inspired me to create `rustclr` and contribute with some features:

- [InlineExecute-Assembly](https://github.com/anthemtotheego/InlineExecute-Assembly)
- [Being a good CLR host – Modernizing offensive .NET tradecraft](https://www.ibm.com/think/x-force/being-a-good-clr-host-modernizing-offensive-net-tradecraft)
- [Microsoft - windows-rs](https://github.com/microsoft/windows-rs)

## License

This project is licensed under the MIT License. See the [LICENSE](/LICENSE) file for details.
