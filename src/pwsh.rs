use alloc::{format, string::String, vec};
use obfstr::obfstr as s;

use crate::error::Result;
use crate::com::_Assembly;
use crate::string::ComString;
use crate::{Invocation, RustClrEnv};
use crate::variant::{Variant, create_safe_args};

/// Provides a persistent interface for executing PowerShell commands.
/// 
/// # Example
/// 
/// ```
/// let pwsh = PowerShell::new()?;
/// let out = pwsh.execute("whoami")?;
/// print!("Output: {}", out);
/// ```
pub struct PowerShell {
    /// The loaded .NET automation assembly.
    automation: _Assembly,

    /// CLR environment used to host the .NET runtime.
    #[allow(dead_code)]
    clr: RustClrEnv,
}

impl PowerShell {
    /// Creates a new `PowerShell`.
    pub fn new() -> Result<Self> {
        // Initialize .NET runtime (v4.0).
        let clr = RustClrEnv::new(None)?;

        // Load `mscorlib` and resolve `System.Reflection.Assembly`.
        let mscorlib = clr.app_domain.get_assembly(s!("mscorlib"))?;
        let reflection_assembly = mscorlib.resolve_type(s!("System.Reflection.Assembly"))?;

        // Resolve and invoke `LoadWithPartialName` method.
        let load_partial_name = reflection_assembly.method_signature(s!(
            "System.Reflection.Assembly LoadWithPartialName(System.String)"
        ))?;
        let param = create_safe_args(vec![s!("System.Management.Automation").to_variant()])?;
        let result = load_partial_name.invoke(None, Some(param))?;

        // Convert result to `_Assembly`.
        let automation =
            _Assembly::from_raw(unsafe { result.Anonymous.Anonymous.Anonymous.byref })?;

        Ok(Self { automation, clr })
    }

    /// Executes a PowerShell command and returns its output as a string.
    pub fn execute(&self, command: &str) -> Result<String> {
        // Invoke `CreateRunspace` method.
        let runspace_factory = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.RunspaceFactory")
        )?;
        let create_runspace = runspace_factory.method_signature(s!(
            "System.Management.Automation.Runspaces.Runspace CreateRunspace()"
        ))?;
        let runspace = create_runspace.invoke(None, None)?;

        // Invoke `CreatePipeline` method.
        let assembly_runspace = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.Runspace"),
        )?;
        assembly_runspace.invoke(
            s!("Open"),
            Some(runspace),
            None,
            Invocation::Instance,
        )?;
        
        let create_pipeline = assembly_runspace.method_signature(s!(
            "System.Management.Automation.Runspaces.Pipeline CreatePipeline()"
        ))?;
        let pipe = create_pipeline.invoke(Some(runspace), None)?;

        // Invoke `get_Commands` method.
        let pipeline = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.Pipeline"),
        )?;
        let get_command = pipeline.invoke(
            s!("get_Commands"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        // Invoke `AddScript` method.
        let cmd = vec![format!("{} | {}", command, s!("Out-String")).to_variant()];
        let command_collection = self.automation.resolve_type(
            s!("System.Management.Automation.Runspaces.CommandCollection"),
        )?;
        let add_script = command_collection.method_signature(s!(
            "Void AddScript(System.String)"
        ))?;
        
        add_script.invoke(Some(get_command), Some(create_safe_args(cmd)?))?;

        // Invoke `InvokeAsync` method.
        pipeline.invoke(
            s!("InvokeAsync"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        // Invoke `get_Output` method.
        let get_output = pipeline.invoke(
            s!("get_Output"),
            Some(pipe),
            None,
            Invocation::Instance,
        )?;

        // Invoke `Read` method.
        let pipeline_reader = self.automation.resolve_type(s!(
            "System.Management.Automation.Runspaces.PipelineReader`1[System.Management.Automation.PSObject]"
        ))?;
        let read = pipeline_reader.method_signature(s!(
            "System.Management.Automation.PSObject Read()"
        ))?;
        let ps_object_instance = read.invoke(Some(get_output), None)?;

        // Invoke `ToString` method.
        let ps_object = self.automation.resolve_type(s!("System.Management.Automation.PSObject"))?;
        let to_string = ps_object.method_signature(s!("System.String ToString()"))?;
        let output = to_string.invoke(Some(ps_object_instance), None)?;

        // Close runspace
        assembly_runspace.invoke(
            s!("Close"),
            Some(runspace),
            None,
            Invocation::Instance,
        )?;

        Ok(unsafe {
            output
                .Anonymous
                .Anonymous
                .Anonymous
                .bstrVal
                .to_string()
        })
    }
}

#[cfg(test)]
mod tests {
    use crate::error::Result;
    use super::PowerShell;

    #[test]
    fn test_powershell() -> Result<()> {
        let pwsh = PowerShell::new()?;
        let output = pwsh.execute("whoami /all")?;
        assert!(
            output.contains("\\")
                || output.contains("User")
                || output.contains("Account")
                || output.contains("Authority"),
            "whoami output does not look valid: {output}"
        );

        Ok(())
    }
}