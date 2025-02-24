use crate::error::ClrError;
use windows_sys::Win32::System::{
    Diagnostics::Debug::{
        IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR, IMAGE_FILE_DLL, 
        IMAGE_FILE_EXECUTABLE_IMAGE, IMAGE_NT_HEADERS64, 
        IMAGE_SUBSYSTEM_NATIVE
    }, 
    SystemServices::{
        IMAGE_DOS_HEADER, 
        IMAGE_DOS_SIGNATURE, 
        IMAGE_NT_SIGNATURE
    }
};

/// Extracts the NT header from the given buffer if it represents a valid PE file.
/// 
/// # Arguments
/// 
/// * `buffer` - A reference to a byte slice representing the potential PE file.
/// 
/// # Returns
/// 
/// * `Some(*const IMAGE_NT_HEADERS64)` - If the buffer contains a valid NT header.
/// * `None` - If the buffer does not represent a valid NT header.
unsafe fn get_nt_header(buffer: &[u8]) -> Option<*const IMAGE_NT_HEADERS64> {
    if buffer.len() < size_of::<IMAGE_DOS_HEADER>() {
        return None;
    }

    let dos_header = buffer.as_ptr() as *const IMAGE_DOS_HEADER;
    if (*dos_header).e_magic != IMAGE_DOS_SIGNATURE {
        return None;
    }

    if (*dos_header).e_lfanew >= (buffer.len() - size_of::<IMAGE_NT_HEADERS64>()) as i32 {
        return None;
    }

    let nt_header = (buffer.as_ptr() as usize + (*dos_header).e_lfanew as usize) as *const IMAGE_NT_HEADERS64;
    if (*nt_header).Signature != IMAGE_NT_SIGNATURE {
        return None;
    }

    Some(nt_header)
}

/// Checks if the given buffer represents a valid PE executable (non-DLL, non-Native).
/// 
/// # Arguments
/// 
/// * `buffer` - A reference to a byte slice representing the potential PE file.
/// 
/// # Returns
/// 
/// * `true` - If the buffer represents a valid PE executable.
/// * `false` - If the buffer is not a valid PE executable.
pub(crate) fn is_exe(buffer: &[u8]) -> bool {
    unsafe {
        if let Some(nt_header) = get_nt_header(buffer) {
            let characteristics = (*nt_header).FileHeader.Characteristics;

            return characteristics & IMAGE_FILE_EXECUTABLE_IMAGE != 0
                && characteristics & IMAGE_FILE_DLL == 0
                && characteristics & IMAGE_SUBSYSTEM_NATIVE == 0;
        }

        false
    }
}

/// Checks if the given buffer represents a valid .NET executable.
/// 
/// # Arguments
/// 
/// * `buffer` - A reference to a byte slice representing the potential .NET assembly.
/// 
/// # Returns
/// 
/// * `true` - If the buffer represents a valid .NET executable.
/// * `false` - If the buffer is not a .NET executable.
pub(crate) fn is_dotnet(buffer: &[u8]) -> bool {
    unsafe {
        if let Some(nt_header) = get_nt_header(buffer) {
            let com_directory = (*nt_header).OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR as usize];
            return com_directory.VirtualAddress != 0 && com_directory.Size != 0;
        }

        false
    }
}

/// Validates if the given buffer represents a valid .NET executable.
///
/// # Arguments
///
/// * `buffer` - A reference to a byte slice representing the potential .NET assembly.
///
/// # Returns
/// 
/// * `Ok(())` - If the environment is successfully prepared.
/// * `Err(ClrError)` - If any error occurs during the preparation process.
pub(crate) fn validate_file(buffer: &[u8]) -> Result<(), ClrError> {
    if !is_exe(buffer) {
        return Err(ClrError::InvalidExecutable);
    }

    if !is_dotnet(buffer) {
        return Err(ClrError::NotDotNet);
    }

    Ok(())
}
