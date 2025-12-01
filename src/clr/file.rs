use alloc::{ffi::CString, vec, vec::Vec};
use core::ptr::null_mut;

use dinvk::{types::IMAGE_NT_HEADERS, helper::PE};
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

use crate::error::{ClrError, Result};

/// Validates whether the given PE buffer represents a valid .NET executable.
///
/// # Errors
/// 
/// Returns a [`ClrError`] variant if the file is not valid or not a .NET assembly.
pub fn validate_file(buffer: &[u8]) -> Result<()> {
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

/// Reads the entire contents of a file from disk into memory using the Win32 API.
pub fn read_file(name: &str) -> Result<Vec<u8>> {
    let file_name = CString::new(name)
        .map_err(|_| ClrError::Msg("invalid cstring"))?;
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
        return Err(ClrError::Msg("failed to open file"));
    }

    let size = unsafe { GetFileSize(h_file, null_mut()) };
    if size == INVALID_FILE_SIZE {
        return Err(ClrError::Msg("invalid file size"));
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

/// Checks whether the PE headers represent a valid Windows executable.
fn is_valid_executable(nt_header: *const IMAGE_NT_HEADERS) -> bool {
    unsafe {
        let characteristics = (*nt_header).FileHeader.Characteristics;
        (characteristics & IMAGE_FILE_EXECUTABLE_IMAGE != 0)
            && (characteristics & IMAGE_FILE_DLL == 0)
            && (characteristics & IMAGE_SUBSYSTEM_NATIVE == 0)
    }
}

/// Checks if the PE includes a COM Descriptor directory, indicating a .NET assembly.
///
/// The COM descriptor is required for the CLR to recognize and load the assembly.
fn is_dotnet(nt_header: *const IMAGE_NT_HEADERS) -> bool {
    unsafe {
        let com_dir = (*nt_header).OptionalHeader.DataDirectory
            [IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR as usize];
        com_dir.VirtualAddress != 0 && com_dir.Size != 0
    }
}
