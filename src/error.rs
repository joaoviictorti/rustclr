//! Error definitions for CLR and COM runtime interaction.

use alloc::string::String;
use thiserror::Error;

/// Result alias for CLR-related operations.
pub type Result<T> = core::result::Result<T, ClrError>;

/// Represents all possible errors that can occur when interacting with the .NET runtime.
#[derive(Debug, Error)]
pub enum ClrError {
    /// File read failure.
    #[error("file read error: {0}")]
    FileReadError(String),

    /// API call failed with HRESULT.
    #[error("{0} failed with HRESULT: {1}")]
    ApiError(&'static str, i32),

    /// Entrypoint expected arguments but none were supplied.
    #[error("entrypoint expected arguments but received none")]
    MissingArguments,

    /// Invalid interface cast.
    #[error("interface cast failed: {0}")]
    CastingError(&'static str),

    /// Provided buffer is not a valid executable.
    #[error("invalid or unsupported executable buffer")]
    InvalidExecutable,

    /// Method not found in target assembly.
    #[error("method not found")]
    MethodNotFound,

    /// Property not found in target assembly.
    #[error("property not found")]
    PropertyNotFound,

    /// Executable is not a .NET assembly.
    #[error("not a .NET application")]
    NotDotNet,

    /// Failed to create CLR MetaHost.
    #[error("metahost creation failed: {0}")]
    MetaHostCreationError(String),

    /// Failed to retrieve runtime information.
    #[error("runtime info retrieval failed: {0}")]
    RuntimeInfoError(String),

    /// Failed to obtain runtime host interface.
    #[error("runtime host retrieval failed: {0}")]
    RuntimeHostError(String),

    /// Runtime start failure.
    #[error("failed to start CLR runtime")]
    RuntimeStartError,

    /// Failed to create AppDomain.
    #[error("domain creation failed: {0}")]
    DomainCreationError(String),

    /// Failed to retrieve default AppDomain.
    #[error("default domain retrieval failed: {0}")]
    DefaultDomainError(String),

    /// No AppDomain available in runtime.
    #[error("no domain available in current runtime")]
    NoDomainAvailable,

    /// Null pointer passed to an API expecting a valid reference.
    #[error("null pointer passed to {0} API")]
    NullPointerError(&'static str),

    /// SafeArray creation failure.
    #[error("safearray creation failed: {0}")]
    SafeArrayError(String),

    /// Unsupported VARIANT type.
    #[error("unsupported VARIANT type")]
    VariantUnsupported,

    /// Generic error with descriptive message.
    #[error("{0}")]
    Msg(&'static str),

    /// Invalid or missing NT header in PE image.
    #[error("invalid PE file: missing or malformed NT header")]
    InvalidNtHeader,
}
