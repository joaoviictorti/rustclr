use alloc::{string::String, vec::Vec};
use windows_sys::Win32::Foundation::{SysAllocString, SysStringLen};

/// Module related to safearray creation
mod safearray;
pub use safearray::*;

/// Module used to validate that the file corresponds to what is expected
pub(crate) mod file;

/// The `WinStr` trait provides methods for working with BSTRs (Binary String),
/// a format commonly used in Windows API. BSTRs are wide strings (UTF-16)
/// with specific memory layouts, used for interoperation with COM
/// (Component Object Model) and other Windows-based APIs.
pub trait WinStr {
    /// Converts a Rust string into a BSTR.
    ///
    /// # Returns
    ///
    /// * `*const u16` - A pointer to the UTF-16 encoded BSTR.
    ///
    /// This method is implemented for `&str` and `String`, converting
    /// them into BSTRs, and for `*const u16` as a passthrough.
    ///
    /// # Example
    ///
    /// ```
    /// use rustclr::WinStr;
    ///
    /// let rust_str = "Hello, World!";
    /// let bstr_ptr = rust_str.to_bstr();
    ///
    /// // Use the BSTR pointer in a COM function...
    /// ```
    fn to_bstr(&self) -> *const u16;

    /// Converts a BSTR (pointer `*const u16`) back to a Rust `String`.
    ///
    /// # Returns
    ///
    /// * `String` - A `String` containing the text from the BSTR if the trait
    ///   is implemented for `*const u16`. For other types, returns an empty `String`.
    ///
    /// # Example
    ///
    /// ```
    /// use rustclr::WinStr;
    ///
    /// let bstr: *const u16 = /* assume a BSTR from COM */;
    /// let rust_string = bstr.to_string();
    /// ```
    fn to_string(&self) -> String {
        String::new()
    }
}

impl WinStr for &str {
    /// Converts a `&str` to a BSTR.
    ///
    /// # Returns
    ///
    /// * `*const u16` - A pointer to the UTF-16 encoded BSTR.
    fn to_bstr(&self) -> *const u16 {
        let utf16_str = self.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        unsafe { SysAllocString(utf16_str.as_ptr()) }
    }
}

impl WinStr for String {
    /// Converts a `String` to a BSTR.
    ///
    /// # Returns
    ///
    /// * `*const u16` - A pointer to the UTF-16 encoded BSTR.
    fn to_bstr(&self) -> *const u16 {
        let utf16_str = self.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        unsafe { SysAllocString(utf16_str.as_ptr()) }
    }
}

impl WinStr for *const u16 {
    /// Passes through the BSTR pointer without modification.
    ///
    /// # Returns
    ///
    /// * `*const u16` - The original BSTR pointer.
    fn to_bstr(&self) -> *const u16 {
        *self
    }

    /// Converts a `*const u16` BSTR to a `String`.
    ///
    /// # Returns
    ///
    /// * `String` - A `String` containing the UTF-16 encoded text from the BSTR.
    fn to_string(&self) -> String {
        let len = unsafe { SysStringLen(*self) };
        if len == 0 {
            return String::new();
        }

        let slice = unsafe { core::slice::from_raw_parts(*self, len as usize) };
        String::from_utf16_lossy(slice)
    }
}

/// Generates a uuid used to create the AppDomain
pub(crate) fn uuid() -> uuid::Uuid {
    let mut buf = [0u8; 16];

    for i in 0..4 {
        let ticks = unsafe { core::arch::x86_64::_rdtsc() };
        buf[i * 4] = ticks as u8;
        buf[i * 4 + 1] = (ticks >> 8) as u8;
        buf[i * 4 + 2] = (ticks >> 16) as u8;
        buf[i * 4 + 3] = (ticks >> 24) as u8;
    }

    uuid::Uuid::from_bytes(buf)
}

/// Specifies the invocation type for a method, indicating if it is static or instance-based.
pub enum Invocation {
    /// Indicates that the method to invoke is static.
    Static,

    /// Indicates that the method to invoke is an instance method.
    Instance,
}
