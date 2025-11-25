use serde_json::Value;
use std::collections::HashMap;

/// Value extension trait for placeholder replacement / 占位符替换的值扩展 trait
pub trait ValueExt: Send + Sync {
    /// Replace placeholders in cyclic table cells / 替换循环表格单元格中的占位符
    ///
    /// # Arguments / 参数
    /// * `index` - Row index for context / 用于上下文的行索引
    /// * `key` - Placeholder key / 占位符键
    /// * `placeholders` - Value map / 值映射
    fn replace_in_table(
        &self,
        index: usize,
        key: &str,
        placeholders: &HashMap<String, Value>,
    ) -> String;

    /// Replace placeholders in regular text / 替换常规文本中的占位符
    ///
    /// # Arguments / 参数
    /// * `key` - Placeholder key / 占位符键
    /// * `placeholders` - Value map / 值映射
    fn replace(&self, key: &str, placeholders: &HashMap<String, Value>) -> String;
}
