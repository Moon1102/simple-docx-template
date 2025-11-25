use async_zip::error::ZipError;

/// Error type for DOCX operations / DOCX 操作的错误类型
///
/// Wraps errors from XML parsing and ZIP file operations / 包装来自 XML 解析和 ZIP 文件操作的错误
#[derive(Debug)]
pub enum DocxError {
    /// XML parsing error / XML 解析错误
    Xml(quick_xml::Error),

    /// ZIP file operation error / ZIP 文件操作错误
    Zip(ZipError),
}

// Automatic conversion from ZipError / 从 ZipError 自动转换
impl From<ZipError> for DocxError {
    fn from(value: ZipError) -> Self {
        DocxError::Zip(value)
    }
}

// Automatic conversion from XML Error / 从 XML 错误自动转换
impl From<quick_xml::Error> for DocxError {
    fn from(value: quick_xml::Error) -> Self {
        DocxError::Xml(value)
    }
}
