#![allow(unused_imports)]

use rustclr::{RustClr, RuntimeVersion};

#[test]
fn test_create_domain() -> ClrResult<(), Box<dyn std::error::Error>> {
    let buffer = std::fs::read("file").expect("Error reading file");
    let output = RustClr::new(&buffer)?
        .with_domain("CustomDomain")
        .with_output_redirection(true)
        .run()?;

    println!("{output}");

    Ok(())
}

#[test]
fn test_with_args() -> ClrResult<(), Box<dyn std::error::Error>> {
    let buffer = std::fs::read("file").expect("Error reading file");
    let output = RustClr::new(&buffer)?
        .with_args(vec!["test"])
        .with_output_redirection(true)
        .run()?;

    println!("{output}");

    Ok(())
}

#[test]
fn test_with_runtime() -> ClrResult<(), Box<dyn std::error::Error>> {
    let buffer = std::fs::read("file").expect("Error reading file");
    let output = RustClr::new(&buffer)?
        .with_runtime_version(RuntimeVersion::V4)
        .with_output_redirection(true)
        .run()?;

    println!("{output}");

    Ok(())
}

#[test]
fn test_without_args() -> ClrResult<(), Box<dyn std::error::Error>> {
    let buffer = std::fs::read("file").expect("Error reading file");
    let output = RustClr::new(&buffer)?
        .with_output_redirection(true)
        .run()?;

    println!("{output}");

    Ok(())
}
