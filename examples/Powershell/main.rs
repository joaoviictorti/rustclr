use rustclr::PowerShell;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let pwsh = PowerShell::new()?;
    let out = pwsh.execute("whoami")?;
    print!("{}", out);

    Ok(())
}
