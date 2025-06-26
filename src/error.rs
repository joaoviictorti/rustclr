use alloc::string::String;
use thiserror::Error;

/// Represents errors that can occur when interacting with the .NET runtime
/// or while handling .NET-related operations within an unmanaged application.
#[derive(Debug, Error)]
pub enum ClrError {
    /// Raised when a .NET file cannot be read correctly.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the file read error.
    #[error("The file could not be read: {0}")]
    FileReadError(String),

    /// Raised when an API call fails, returning a specific HRESULT.
    ///
    /// # Arguments
    ///
    /// * `{0}` - The name of the API that failed.
    /// * `{1}` - The HRESULT code returned by the API indicating the specific failure.
    #[error("{0} Failed With HRESULT: {1}")]
    ApiError(&'static str, i32),

    /// Raised when an entry point expects arguments but receives none.
    #[error("Entrypoint is waiting for arguments, but has been supplied with zero")]
    MissingArguments,

    /// Raised when there is an error casting a COM interface to the specified type.
    ///
    /// # Arguments
    ///
    /// * `{0}` - The name of the type to which casting failed.
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
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the failure to create the MetaHost.
    #[error("Failed to create the MetaHost: {0}")]
    MetaHostCreationError(String),

    /// Raised when retrieving information about the .NET runtime fails.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the error in runtime information retrieval.
    #[error("Failed to retrieve runtime information: {0}")]
    RuntimeInfoError(String),

    /// Raised when the runtime host interface could not be obtained.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the failure to obtain the runtime host interface.
    #[error("Failed to obtain runtime host interface: {0}")]
    RuntimeHostError(String),

    /// Raised when the runtime fails to start.
    #[error("Failed to start the runtime")]
    RuntimeStartError,

    /// Raised when there is an error creating a new AppDomain.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the domain creation error.
    #[error("Failed to create domain: {0}")]
    DomainCreationError(String),

    /// Raised when the default AppDomain cannot be retrieved.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the error in retrieving the default AppDomain.
    #[error("Failed to retrieve the default domain: {0}")]
    DefaultDomainError(String),

    /// Raised when no AppDomain is available in the runtime environment.
    #[error("No domain available")]
    NoDomainAvailable,

    /// Raised when a null pointer is passed to an API where a valid reference was expected.
    ///
    /// # Arguments
    ///
    /// * `{0}` - The name of the API that received the null pointer.
    #[error("The {0} API received a null pointer where a valid reference was expected")]
    NullPointerError(&'static str),

    /// Raised when there is an error creating a SafeArray.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message describing the SafeArray creation error.
    #[error("Error creating SafeArray: {0}")]
    SafeArrayError(String),

    /// Raised when the type of a VARIANT is unsupported by the current context.
    #[error("Type of VARIANT not supported")]
    VariantUnsupported,

    /// Represents a generic error specific to the CLR.
    ///
    /// # Arguments
    ///
    /// * `{0}` - A message providing details about the CLR-specific error.
    #[error("{0}")]
    GenericError(&'static str),

    /// Related error if the PE file used in the loader does not have a valid NT HEADER
    #[error("Invalid PE file: missing or malformed NT header")]
    InvalidNtHeader,
}
