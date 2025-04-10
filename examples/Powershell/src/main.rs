use rustclr::{
    create_safe_args,
    RustClrEnv, Invocation,
    data::_Assembly, Variant,
    WinStr
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = std::env::args().skip(1).collect::<Vec<String>>();
    let command = args.join(" ");

    // Initialize .NET runtime (v4.0).
    let clr = RustClrEnv::new(None)?;

    // Load `mscorlib` and resolve `System.Reflection.Assembly`.
    let mscorlib = clr.app_domain.get_assembly("mscorlib")?;
    let reflection_assembly = mscorlib.resolve_type("System.Reflection.Assembly")?;

    // Resolve and invoke `LoadWithPartialName` method.
    let load_partial_name = reflection_assembly.method_signature("System.Reflection.Assembly LoadWithPartialName(System.String)")?;
    let param = create_safe_args(vec!["System.Management.Automation".to_variant()])?;
    let result = load_partial_name.invoke(None, Some(param))?;

    // Convert result to `_Assembly`.
    let automation = _Assembly::from_raw(unsafe { result.Anonymous.Anonymous.Anonymous.byref })?;

    // Invoke `CreateRunspace` method.
    let runspace_factory = automation.resolve_type("System.Management.Automation.Runspaces.RunspaceFactory")?;
    let create_runspace = runspace_factory.method_signature("System.Management.Automation.Runspaces.Runspace CreateRunspace()")?;
    let runspace = create_runspace.invoke(None, None)?;

    // Invoke `CreatePipeline` method.
    let assembly_runspace = automation.resolve_type("System.Management.Automation.Runspaces.Runspace")?;
    assembly_runspace.invoke("Open", Some(runspace), None, Invocation::Instance)?;
    let create_pipeline = assembly_runspace.method_signature("System.Management.Automation.Runspaces.Pipeline CreatePipeline()")?;
    let pipe = create_pipeline.invoke(Some(runspace), None)?;

    // Invoke `get_Commands` method.
    let pipeline = automation.resolve_type("System.Management.Automation.Runspaces.Pipeline")?;
    let get_command = pipeline.invoke("get_Commands", Some(pipe), None, Invocation::Instance)?;
    
    // Invoke `AddScript` method.
    let command_collection = automation.resolve_type("System.Management.Automation.Runspaces.CommandCollection")?;
    let cmd= vec![format!("{} | Out-String", command).to_variant()];
    let args = create_safe_args(cmd)?;
    let add_script = command_collection.method_signature("Void AddScript(System.String)")?;
    add_script.invoke(Some(get_command), Some(args))?;

    // Invoke `InvokeAsync` method.
    pipeline.invoke("InvokeAsync", Some(pipe), None, Invocation::Instance)?;

    // Invoke `get_Output` method.
    let get_output = pipeline.invoke("get_Output", Some(pipe), None, Invocation::Instance)?;

    // Invoke `Read` method.
    let pipeline_reader = automation.resolve_type("System.Management.Automation.Runspaces.PipelineReader`1[System.Management.Automation.PSObject]")?;
    let read = pipeline_reader.method_signature("System.Management.Automation.PSObject Read()")?;
    let ps_object_instance = read.invoke(Some(get_output), None)?;

    // Invoke `ToString` method.
    let ps_object = automation.resolve_type("System.Management.Automation.PSObject")?;
    let to_string = ps_object.method_signature("System.String ToString()")?;
    let output = to_string.invoke(Some(ps_object_instance), None)?;

    // Read output.
    let str = unsafe { output.Anonymous.Anonymous.Anonymous.bstrVal.to_string() };
    println!("{}", str);

    assembly_runspace.invoke("Close", Some(runspace), None, Invocation::Instance)?;

    Ok(())
}