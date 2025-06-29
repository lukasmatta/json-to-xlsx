use serde_json::Error as SerdeJsonError;
use std::{fmt, io};
use zip::result::ZipError;

pub type XlsxExporResult<T> = Result<T, XlsxExportError>;

#[derive(Debug)]
#[non_exhaustive]
pub enum XlsxExportError {
    NotAnArray,
    EmptyArray,
    ExpectedObject,
    JsonError(SerdeJsonError),
    IoError(io::Error),
    ZipError(ZipError),
}

impl fmt::Display for XlsxExportError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XlsxExportError::NotAnArray => write!(f, "The JSON root value is not an array"),
            XlsxExportError::EmptyArray => write!(f, "The JSON array is empty"),
            XlsxExportError::ExpectedObject => {
                write!(f, "Expected each item in the array to be a JSON object")
            }
            XlsxExportError::JsonError(err) => write!(f, "JSON error: {}", err),
            XlsxExportError::IoError(err) => write!(f, "IO error: {}", err),
            XlsxExportError::ZipError(err) => write!(f, "ZIP error: {}", err),
        }
    }
}

impl std::error::Error for XlsxExportError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsxExportError::JsonError(err) => Some(err),
            XlsxExportError::IoError(err) => Some(err),
            XlsxExportError::ZipError(err) => Some(err),
            _ => None,
        }
    }
}

impl From<ZipError> for XlsxExportError {
    fn from(value: ZipError) -> Self {
        Self::ZipError(value)
    }
}

impl From<io::Error> for XlsxExportError {
    fn from(value: io::Error) -> Self {
        Self::IoError(value)
    }
}
