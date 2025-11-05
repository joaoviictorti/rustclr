use std::fs;
use clap::{Parser, ArgAction};
use rustclr::{
    RustClr,
    RuntimeVersion,
    error::ClrError, 
};

/// The main command-line interface struct.
#[derive(Parser)]
#[clap(author="joaoviictorti", about="Host CLR and run .NET binaries using Rust", version="1.0")]
pub struct Cli {
    /// Path to the .NET assembly file to be executed.
    #[arg(short, long, required = true, help = "Path to the .NET assembly file")]
    pub file: String,

    /// Arguments for the .NET program (strings only).
    #[arg(short, long, action = ArgAction::Append, help = "String arguments for the .NET program")]
    pub inputs: Option<Vec<String>>,

    /// Specify the .NET runtime version (e.g., "v2", "v3", "v4").
    #[arg(short, long, default_value = "v4", help = "Specify .NET runtime version")]
    pub runtime_version: String,

    /// Set a custom application domain name.
    #[arg(short = 'd', long, help = "Set custom application domain name")]
    pub domain: Option<String>,
}

fn main() -> Result<(), ClrError> {
    // Parse command-line arguments
    let cli = Cli::parse();

    // Read the .NET assembly file
    let data = fs::read(&cli.file)
        .map_err(|_| ClrError::Msg("failed to read file"))?;

    // Convert version string to RuntimeVersion enum
    let runtime_version = match cli.runtime_version.as_str() {
        "v2" => RuntimeVersion::V2,
        "v3" => RuntimeVersion::V3,
        "v4" => RuntimeVersion::V4,
        _ => RuntimeVersion::UNKNOWN,
    };

    // Initialize and configure the RustClr instance
    let mut clr = RustClr::new(data.as_slice())?
        .with_runtime_version(runtime_version)
        .with_output();

    // Set the custom application domain if provided
    if let Some(domain_name) = cli.domain {
        clr = clr.with_domain(&domain_name);
    }

    // Set the string arguments for the .NET assembly if provided
    if let Some(inputs) = cli.inputs {
        // Convert Vec<String> to Vec<&str>
        let args = inputs.iter().map(|s| s.as_str()).collect::<Vec<&str>>();
        clr = clr.with_args(args);
    }

    // Run the .NET assembly
    match clr.run() {
        Ok(output) => println!("Output: {}", output),
        Err(err) => println!("Error: {err}")
    }
    
    Ok(())
}
