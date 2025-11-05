// Copyright (c) 2025 joaoviictorti
// Licensed under the MIT License. See LICENSE file in the project root for details.

//! Manages specific error types used when interacting with the CLR and COM APIs.

use alloc::string::String;
use thiserror::Error;

/// Result alias for CLR-related operations.
pub type Result<T> = core::result::Result<T, ClrError>;

/// Represents errors that can occur when interacting with the .NET runtime.
#[derive(Debug, Error)]
pub enum ClrError {
    /// Raised when a .NET file cannot be read correctly.
    #[error("The file could not be read: {0}")]
    FileReadError(String),

    /// Raised when an API call fails, returning a specific HRESULT.
    #[error("{0} Failed With HRESULT: {1}")]
    ApiError(&'static str, i32),

    /// Raised when an entry point expects arguments but receives none.
    #[error("Entrypoint is waiting for arguments, but has been supplied with zero")]
    MissingArguments,

    /// Raised when there is an error casting a COM interface to the specified type.
    #[error("Error casting the interface to {0}")]
    CastingError(&'static str),

    /// Raised when the buffer provided does not represent a valid executable file.
    #[error("The buffer does not represent a valid executable")]
    InvalidExecutable,

    /// Raised when a required method is not found in the .NET assembly.
    #[error("Method not found")]
    MethodNotFound,

    /// Raised when a required property is not found in the .NET assembly.
    #[error("Property not found")]
    PropertyNotFound,

    /// Raised when the buffer does not contain a .NET application.
    #[error("The executable is not a .NET application")]
    NotDotNet,

    /// Raised when there is a failure creating the .NET MetaHost.
    #[error("Failed to create the MetaHost: {0}")]
    MetaHostCreationError(String),

    /// Raised when retrieving information about the .NET runtime fails.
    #[error("Failed to retrieve runtime information: {0}")]
    RuntimeInfoError(String),

    /// Raised when the runtime host interface could not be obtained.
    #[error("Failed to obtain runtime host interface: {0}")]
    RuntimeHostError(String),

    /// Raised when the runtime fails to start.
    #[error("Failed to start the runtime")]
    RuntimeStartError,

    /// Raised when there is an error creating a new AppDomain.
    #[error("Failed to create domain: {0}")]
    DomainCreationError(String),

    /// Raised when the default AppDomain cannot be retrieved.
    #[error("Failed to retrieve the default domain: {0}")]
    DefaultDomainError(String),

    /// Raised when no AppDomain is available in the runtime environment.
    #[error("No domain available")]
    NoDomainAvailable,

    /// Raised when a null pointer is passed to an API where a valid reference was expected.
    #[error("The {0} API received a null pointer where a valid reference was expected")]
    NullPointerError(&'static str),

    /// Raised when there is an error creating a SafeArray.
    #[error("Error creating SafeArray: {0}")]
    SafeArrayError(String),

    /// Raised when the type of a VARIANT is unsupported by the current context.
    #[error("Type of VARIANT not supported")]
    VariantUnsupported,

    /// Represents a generic error specific to the CLR.
    #[error("{0}")]
    GenericError(&'static str),

    /// Related error if the PE file used in the loader does not have a valid NT HEADER
    #[error("Invalid PE file: missing or malformed NT header")]
    InvalidNtHeader,
}
