# rustclr 🦀

![Rust](https://img.shields.io/badge/made%20with-Rust-red)
![crate](https://img.shields.io/crates/v/rustclr.svg)
![docs](https://docs.rs/rustclr/badge.svg)
[![build](https://github.com/joaoviictorti/rustclr/actions/workflows/ci.yml/badge.svg)](https://github.com/joaoviictorti/rustclr/actions/workflows/ci.yml)
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

- **`with_runtime_version(RuntimeVersion::V4)`**: Sets the .NET runtime version. This flag ensures that the assembly runs with the specified CLR version.
- **`with_output`**: Redirects the output from the .NET assembly's console to the Rust environment, capturing all console output.
- **`with_domain("DomainName")`**: Sets a custom AppDomain name, which is useful for isolating different .NET assemblies.
- **`with_args(vec!["arg1", "arg2"])`**: Passes arguments to the .NET application, useful for parameterized entry points in the assembly.
- **`with_patch_exit`**: This prevents calls to `System.Environment.Exit()` within the .NET assembly from terminating the host process (your Rust program).
  
Using `rustclr` to load and execute a .NET assembly, redirect its output and customize the CLR runtime environment.

```rust
use std::fs;
use rustclr::{RustClr, RuntimeVersion};

fn main() -> ClrResult<(), Box<dyn std::error::Error>> {
    // Load a sample .NET assembly into a buffer
    let buffer = fs::read("examples/sample.exe")?;

    // Create and configure a RustClr instance with runtime version and output redirection
    let output = RustClr::new(&buffer)?
        .with_runtime_version(RuntimeVersion::V4)
        .with_output()
        .with_domain("CustomDomain")
        .with_patch_exit()
        .with_args(vec!["arg1", "arg2"])
        .run()?;

    println!("Captured output: {}", output);

    Ok(())
}
```

### Running PowerShell Commands

`rustclr` also provides a high-level interface to execute `PowerShell` commands from Rust using the built-in .NET `System.Management.Automation` namespace.

```rust
use std::error::Error;
use rustclr::PowerShell;

fn main() -> ClrResult<(), Box<dyn Error>> {
    let pwsh = PowerShell::new()?;
    print!("{}", pwsh.execute("Get-Process | Select-Object -First 3")?);
    print!("{}", pwsh.execute("whoami")?);
    
    Ok(())
}
```

### Configuration with RustClrEnv and ClrOutput

For more fine-grained control, rustclr provides the `RustClrEnv` and `ClrOutput` components:

- **`RustClrEnv`**: Allows for low-level customization and initialization of the .NET runtime environment, which is useful if you need to manually control the CLR version, MetaHost, runtime information, and application domain. This struct provides an alternative way to initialize a CLR environment without executing an assembly immediately.
```rust
use rustclr::{RustClrEnv, RuntimeVersion};

fn main() -> ClrResult<(), Box<dyn std::error::Error>> {
    // Create a new environment for .NET with a specific runtime version
    let clr_env = RustClrEnv::new(Some(RuntimeVersion::V4))?;
    println!("CLR environment initialized successfully with version {:?}", clr_env.runtime_version);

    Ok(())
}
```

- **`ClrOutput`**: Manages redirection of standard output and error streams from .NET to Rust. This is especially useful if you need to capture and process all output produced by .NET code within a Rust environment.
```rust
use rustclr::variant::Variant;
use rustclr::{ClrOutput, Invocation, RustClrEnv};

fn main() -> ClrResult<(), Box<dyn std::error::Error>> {
    // Initialize the CLR environment
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.get_assembly("mscorlib")?;
    let console = mscorlib.resolve_type("System.Console")?;

    // First redirection: captures Console.WriteLine output
    let mut clr_output = ClrOutput::new(&mscorlib);
    clr_output.redirect()?;

    // Call Console.WriteLine("Hello World")
    let args = vec!["Hello World".to_variant()];
    console.invoke("WriteLine", None, Some(args), Invocation::Static)?;

    // Capture output
    let output = clr_output.capture()?;
    print!("{output}");

    Ok(())
}
```

## Additional Resources

For more examples, check the [examples](https://github.com/joaoviictorti/rustclr/tree/main/examples) folder in the repository.

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

This project is licensed under the MIT License. See the [LICENSE](https://github.com/joaoviictorti/rustclr/tree/main/LICENSE) file for details.
