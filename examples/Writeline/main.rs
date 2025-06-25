use rustclr::{ClrOutput, Invocation, RustClrEnv, Variant};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create and initialize the CLR environment
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.get_assembly("mscorlib")?;
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
