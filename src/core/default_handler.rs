use crate::public::value_extern::ValueExt;
use serde_json::Value;
use std::collections::HashMap;

/// Default implementation of placeholder value handling / 占位符值处理的默认实现
///
/// Provides standard placeholder replacement logic with support for:
/// 提供标准的占位符替换逻辑，支持：
/// - Basic value substitution / 基本值替换
/// - Uppercase transformation (^) / 大写转换 (^)
/// - Image placeholders (@) / 图片占位符 (@)
/// - Index placeholders ($index) / 索引占位符 ($index)
#[derive(Default)]
pub(crate) struct DefaultValueHandler;

impl DefaultValueHandler {
    /// Convert JSON value to string without quotes / 将 JSON 值转换为不带引号的字符串
    ///
    /// # Arguments / 参数
    /// * `value` - JSON value to convert / 要转换的 JSON 值
    ///
    /// # Returns / 返回
    /// String representation of the value / 值的字符串表示
    fn handle_without_quotes(value: &Value) -> String {
        match value {
            // String values returned as-is / 字符串值原样返回
            Value::String(s) => s.to_owned(),

            // Null becomes empty string / Null 变为空字符串
            Value::Null => "".to_string(),

            // Numbers formatted to 2 decimal places / 数字格式化为 2 位小数
            Value::Number(n) => n
                .as_f64()
                .map(|v| format!("{:.2}", v))
                .unwrap_or_else(|| "".to_string()),

            // Other types use default JSON serialization / 其他类型使用默认 JSON 序列化
            _ => value.to_string(),
        }
    }
}

// Implementation of ValueExt trait / ValueExt trait 的实现
impl ValueExt for DefaultValueHandler {
    /// Replace placeholders in table cells / 替换表格单元格中的占位符
    ///
    /// Supports special syntax:
    /// 支持特殊语法：
    /// - `[^key]` - Uppercase value / 大写值
    /// - `[@key]` - Image placeholder / 图片占位符
    /// - `[$index]` - Row index / 行索引
    /// - `[key]` - Normal value / 普通值
    ///
    /// # Arguments / 参数
    /// * `index` - Current row index / 当前行索引
    /// * `key` - Placeholder key with brackets / 带括号的占位符键
    /// * `placeholders` - Value map / 值映射
    fn replace_in_table(
        &self,
        index: usize,
        key: &str,
        placeholders: &HashMap<String, Value>,
    ) -> String {
        let mut result = key.to_string();
        // Remove brackets from key / 从键中移除括号
        let cleaned_key = result.replace("]", "").replace("[", "");

        // Helper to get value from placeholders / 从占位符获取值的辅助函数
        let handle = |cleaned_key: String| -> String {
            if let Some(row) = placeholders.get(&cleaned_key) {
                Self::handle_without_quotes(row)
            } else {
                "".to_string()
            }
        };

        // Handle uppercase transformation / 处理大写转换
        if cleaned_key.contains("^") {
            result = handle(cleaned_key.replace("^", "")).to_uppercase()
        }
        // Handle image placeholder - return base64 value / 处理图片占位符 - 返回 base64 值
        else if cleaned_key.contains("@") {
            result = handle(cleaned_key.replace("@", ""))
        }
        // Handle row index / 处理行索引
        else if cleaned_key == "$index" {
            result = index.to_string();
        }
        // Handle default content / 处理默认内容
        else {
            result = handle(cleaned_key);
        }

        result
    }

    /// Replace placeholders in regular text / 替换常规文本中的占位符
    ///
    /// # Arguments / 参数
    /// * `content` - Text content that may contain placeholders / 可能包含占位符的文本内容
    /// * `placeholders` - Value map / 值映射
    fn replace(&self, content: &str, placeholders: &HashMap<String, Value>) -> String {
        // If content looks like a placeholder, process it / 如果内容看起来像占位符，则处理它
        if content.starts_with("{{") && content.ends_with("}}") {
            return self.replace_in_table(0, content, placeholders);
        }

        // Return original content if no match / 如果没有匹配则返回原始内容
        content.to_string()
    }
}
