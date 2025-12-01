//! Defines the `ComString` trait for converting between Rust strings and BSTRs.

use alloc::{string::String, vec::Vec};
use windows_sys::Win32::Foundation::{SysAllocString, SysStringLen};

/// The `ComString` trait provides methods for working with BSTRs.
pub trait ComString {
    /// Converts a Rust string into a BSTR.
    fn to_bstr(&self) -> *const u16;

    /// Converts a BSTR to a Rust string.
    fn to_string(&self) -> String {
        String::new()
    }
}

impl ComString for &str {
    fn to_bstr(&self) -> *const u16 {
        let utf16_str = self.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        unsafe { SysAllocString(utf16_str.as_ptr()) }
    }
}

impl ComString for String {
    fn to_bstr(&self) -> *const u16 {
        let utf16_str = self.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        unsafe { SysAllocString(utf16_str.as_ptr()) }
    }
}

impl ComString for *const u16 {
    fn to_bstr(&self) -> *const u16 {
        *self
    }

    fn to_string(&self) -> String {
        let len = unsafe { SysStringLen(*self) };
        if len == 0 {
            return String::new();
        }

        let slice = unsafe { core::slice::from_raw_parts(*self, len as usize) };
        String::from_utf16_lossy(slice)
    }
}
