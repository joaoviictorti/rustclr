use alloc::{ffi::CString, vec, vec::Vec};
use core::ptr::null_mut;

use dinvk::{data::IMAGE_NT_HEADERS, parse::PE};
use windows_sys::Win32::System::Diagnostics::Debug::{
    IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR, IMAGE_FILE_DLL, 
    IMAGE_FILE_EXECUTABLE_IMAGE, IMAGE_SUBSYSTEM_NATIVE,
};
use windows_sys::Win32::{
    Foundation::{GENERIC_READ, INVALID_HANDLE_VALUE},
    Storage::FileSystem::{
        CreateFileA, FILE_ATTRIBUTE_NORMAL, FILE_SHARE_READ, 
        GetFileSize, INVALID_FILE_SIZE, OPEN_EXISTING, ReadFile,
    },
};

use crate::{Result, error::ClrError};

/// Checks if the PE headers indicate a valid Windows executable (not DLL, not Native subsystem).
///
/// # Safety
///
/// `nt_header` must be a valid pointer to an `IMAGE_NT_HEADERS` struct.
fn is_valid_executable(nt_header: *const IMAGE_NT_HEADERS) -> bool {
    unsafe {
        let characteristics = (*nt_header).FileHeader.Characteristics;
        (characteristics & IMAGE_FILE_EXECUTABLE_IMAGE != 0)
            && (characteristics & IMAGE_FILE_DLL == 0)
            && (characteristics & IMAGE_SUBSYSTEM_NATIVE == 0)
    }
}

/// Checks if the PE contains a COM Descriptor directory (i.e., is a .NET assembly).
///
/// # Safety
///
/// `nt_header` must be a valid pointer to an `IMAGE_NT_HEADERS` struct.
fn is_dotnet(nt_header: *const IMAGE_NT_HEADERS) -> bool {
    unsafe {
        let com_dir = (*nt_header).OptionalHeader.DataDirectory
            [IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR as usize];
        com_dir.VirtualAddress != 0 && com_dir.Size != 0
    }
}

/// Validates whether the given PE buffer represents a .NET executable.
///
/// # Returns
///
/// * `Ok(())` if the buffer is a valid .NET executable.
/// * `Err(ClrError)` otherwise.
pub(crate) fn validate_file(buffer: &[u8]) -> Result<()> {
    let pe = PE::parse(buffer.as_ptr().cast_mut().cast());
    let Some(nt_header) = pe.nt_header() else {
        return Err(ClrError::InvalidNtHeader);
    };

    if !is_valid_executable(nt_header) {
        return Err(ClrError::InvalidExecutable);
    }

    if !is_dotnet(nt_header) {
        return Err(ClrError::NotDotNet);
    }

    Ok(())
}

/// Reads the entire contents of a file into memory using the Windows API.
///
/// # Arguments
///
/// * `name` - The path to the file as a UTF-8 string.
///
/// # Returns
///
/// Returns `Ok(Vec<u8>)` containing the file's contents if the operation succeeds, or a
/// `ClrError::GenericError` if any step fails.
pub fn read_file(name: &str) -> Result<Vec<u8>> {
    let file_name = CString::new(name).map_err(|_| ClrError::GenericError("Invalid cstring"))?;
    let h_file = unsafe {
        CreateFileA(
            file_name.as_ptr().cast(),
            GENERIC_READ,
            FILE_SHARE_READ,
            null_mut(),
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            null_mut(),
        )
    };

    if h_file == INVALID_HANDLE_VALUE {
        return Err(ClrError::GenericError("Failed to open file"));
    }

    let size = unsafe { GetFileSize(h_file, null_mut()) };
    if size == INVALID_FILE_SIZE {
        return Err(ClrError::GenericError("Invalid file size"));
    }

    let mut out = vec![0; size as usize];
    let mut bytes = 0;
    unsafe {
        ReadFile(
            h_file,
            out.as_mut_ptr(),
            out.len() as u32,
            &mut bytes,
            null_mut(),
        );
    }

    Ok(out)
}
