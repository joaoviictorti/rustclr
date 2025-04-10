# rustclr 🦀

![Rust](https://img.shields.io/badge/made%20with-Rust-red)
![crate](https://img.shields.io/crates/v/rustclr.svg)
![docs](https://docs.rs/rustclr/badge.svg)
![Forks](https://img.shields.io/github/forks/joaoviictorti/rustclr)
![Stars](https://img.shields.io/github/stars/joaoviictorti/rustclr)
![License](https://img.shields.io/github/license/joaoviictorti/rustclr)

`rustclr` is a powerful library for hosting the Common Language Runtime (CLR) and executing .NET binaries directly with Rust, among other operations.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
  - [Running a .NET Assembly with Configured Flags](#running-a-net-assembly-with-configured-flags)
  - [Configuration with RustClrEnv and ClrOutput](#configuration-with-rustclrenv-and-clroutput)
  - [Running PowerShell Commands](#running-powershell-commands)
- [Additional Resources](#additional-resources)
- [CLI](#cli)
  - [Example Command](#example-command)
  - [CLI Help](#cli-help)
- [Contributing to rustclr](#contributing-to-rustclr)
- [References](#references)
- [License](#license)

## Features

- ✅ Run .NET binaries in memory with full control over runtime configurations
- ✅ Fine-grained control over the CLR environment and runtime initialization
- ✅ Configure output redirection to capture .NET program output

## Installation

Add `rustclr` to your project by updating your `Cargo.toml`:
```bash
cargo add rustclr
```

Or manually add the dependency:
```toml
[dependencies]
rustclr = "<version>"
```

## Usage

### Running a .NET Assembly with Configured Flags

The following flags provide full control over your CLR environment and the execution of your .NET assemblies:

- **`.with_runtime_version(RuntimeVersion::V4)`**: Sets the .NET runtime version (e.g., RuntimeVersion::V2, RuntimeVersion::V3, RuntimeVersion::V4). This flag ensures that the assembly runs with the specified CLR version.
- **`.with_output_redirection(true)`**: Redirects the output from the .NET assembly's console to the Rust environment, capturing all console output.
- **`.with_domain("DomainName")`**: Sets a custom AppDomain name, which is useful for isolating different .NET assemblies.
- **`.with_args(vec!["arg1", "arg2"])`**: Passes arguments to the .NET application, useful for parameterized entry points in the assembly.
  
Using `rustclr` to load and execute a .NET assembly, redirect its output and customize the CLR runtime environment.

```rs
use std::fs;
use rustclr::{RustClr, RuntimeVersion};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load a sample .NET assembly into a buffer
    let buffer = fs::read("examples/sample.exe")?;

    // Create and configure a RustClr instance with runtime version and output redirection
    let output = RustClr::new(&buffer)?
        .with_runtime_version(RuntimeVersion::V4) // Specify .NET runtime version
        .with_output_redirection(true) // Redirect output to capture it in Rust
        .with_domain("CustomDomain") // Optionally set a custom application domain
        .with_args(vec!["arg1", "arg2"]) // Pass arguments to the .NET assembly's entry point
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

## Contributing to rustclr

To contribute to **rustclr**, follow these steps:

1. Fork this repository.
2. Create a branch: `git checkout -b <branch_name>`.
3. Make your changes and commit them: `git commit -m '<commit_message>'`.
4. Push your changes to your branch: `git push origin <branch_name>`.
5. Create a pull request.

Alternatively, consult the [GitHub documentation](https://docs.github.com/en/pull-requests/collaborating-with-pull-requests) on how to create a pull request.

## References

- <https://github.com/anthemtotheego/InlineExecute-Assembly>
- <https://github.com/microsoft/windows-rs>

## License

This project is licensed under the MIT License. See the [LICENSE](/LICENSE) file for details.