mod core;
mod public;
#[cfg(test)]
mod tests;

pub use public::docx::DOCX;
pub use public::error::DocxError;
pub use public::value_extern::ValueExt;
