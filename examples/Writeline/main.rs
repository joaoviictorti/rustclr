use rustclr::variant::Variant;
use rustclr::{ClrOutput, Invocation, RustClrEnv};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Initialize the CLR environment and load the 'mscorlib' assembly
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.get_assembly("mscorlib")?;
    let console = mscorlib.resolve_type("System.Console")?;

    // Create a ClrOutput to intercept stdout via StringWriter
    let mut clr_output = ClrOutput::new(&mscorlib);

    // First redirection: captures Console.WriteLine output
    clr_output.redirect()?;

    // Call Console.WriteLine("Hello World")
    let args = vec!["Hello World".to_variant()];
    console.invoke("WriteLine", None, Some(args), Invocation::Static)?;

    // Capture and print the redirected output
    let output = clr_output.capture()?;
    print!("OUTPUT (1) ====> {output}");

    // Second redirection: resets the internal buffer
    clr_output.redirect()?;

    // Call Console.WriteLine("Hello Victor")
    let args = vec!["Hello Victor".to_variant()];
    console.invoke("WriteLine", None, Some(args), Invocation::Static)?;

    // Capture and print the new output
    let output = clr_output.capture()?;
    print!("OUTPUT (2) ====> {output}");

    Ok(())
}
