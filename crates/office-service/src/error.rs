use std::fmt;

/// An agent edit refusal's code, as `OfficeEditError` names it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EditCode {
    InvalidInput,
    StaleTarget,
    UnavailableTarget,
    UnsupportedOperation,
}

impl EditCode {
    pub fn as_str(self) -> &'static str {
        match self {
            EditCode::InvalidInput => "invalid_input",
            EditCode::StaleTarget => "stale_target",
            EditCode::UnavailableTarget => "unavailable_target",
            EditCode::UnsupportedOperation => "unsupported_operation",
        }
    }
}

/// Display strings are the TS error messages Capy matches on.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Error {
    /// `<code>: <message>`, an agent edit the engine refused.
    Edit { code: EditCode, message: String },
    /// `Office rebase: <message>`, a deterministic refusal.
    Rebase(String),
    /// Bad arguments (a TS `TypeError`).
    Invalid(String),
    /// Anything else an engine or the glue reported (a TS `Error`).
    Engine(String),
    /// A panic inside an engine call (a WebAssembly trap in the TS runtime).
    Panic(String),
}

impl Error {
    pub(crate) fn edit(code: EditCode, message: impl Into<String>) -> Self {
        Error::Edit {
            code,
            message: message.into(),
        }
    }

    pub(crate) fn engine(message: impl fmt::Display) -> Self {
        Error::Engine(message.to_string())
    }

    pub(crate) fn rebase(message: impl Into<String>) -> Self {
        Error::Rebase(message.into())
    }

    /// The message a TS `Error` carried (what `transplant` wraps).
    pub(crate) fn message(&self) -> String {
        match self {
            Error::Edit { code, message } => format!("{}: {message}", code.as_str()),
            Error::Rebase(message) => format!("Office rebase: {message}"),
            Error::Invalid(message) | Error::Engine(message) => message.clone(),
            Error::Panic(message) => format!("Office engine panicked: {message}"),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.message())
    }
}

impl std::error::Error for Error {}

pub type Result<T> = std::result::Result<T, Error>;
