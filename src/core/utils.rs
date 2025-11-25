use crate::core::constant::{
    ERR_INVALID_JPG_MARKER, ERR_INVALID_PNG_IHDR, ERR_NO_SOF_MARKER, ERR_SLICE_TOO_SHORT,
    ERR_UNKNOWN_FORMAT, FLATTEN_RECORDS_CAPACITY, JPEG_INITIAL_OFFSET, JPEG_MARKER_DAC,
    JPEG_MARKER_DHT, JPEG_MARKER_JPG, JPEG_MIN_SEGMENT_SIZE, JPEG_SOF_MARKER_END,
    JPEG_SOF_MARKER_START, MIN_IMAGE_DATA_LEN, PNG_IHDR_MARKER, PNG_SIG_BYTE_0, PNG_SIG_BYTE_1,
    PNG_SIG_BYTE_2, PNG_SIG_BYTE_3, REGEX_REL_ID, REL_ID_PREFIX,
};
use regex::Regex;
use serde_json::Value;
use std::collections::HashMap;
use std::sync::LazyLock;

/// Extract image dimensions from PNG or JPEG bytes / 从 PNG 或 JPEG 字节中提取图片尺寸
///
/// Supports PNG and JPEG formats by parsing their headers  / 通过解析头部支持 PNG 和 JPEG 格式
///
/// # Arguments / 参数
/// * `bytes` - Image file bytes / 图片文件字节
///
/// # Returns / 返回
/// * `Ok((width, height))` - Image dimensions in pixels / 图片尺寸（像素）
/// * `Err(msg)` - Error message if format is unsupported / 如果格式不支持则返回错误消息
#[inline]
pub(crate) fn get_image_dimensions<'a>(bytes: &[u8]) -> Result<(f32, f32), &'a str> {
    // Minimum size check / 最小尺寸检查
    if bytes.len() < MIN_IMAGE_DATA_LEN {
        return Err(ERR_SLICE_TOO_SHORT);
    }

    // Check for PNG signature / 检查 PNG 签名
    // PNG signature: 137 80 78 71 13 10 26 10
    if bytes[0] == PNG_SIG_BYTE_0
        && bytes[1] == PNG_SIG_BYTE_1
        && bytes[2] == PNG_SIG_BYTE_2
        && bytes[3] == PNG_SIG_BYTE_3
    {
        // PNG: Check IHDR chunk (skip signature check for perf) / PNG：检查 IHDR 块（跳过签名检查以提升性能）
        if bytes[12] == PNG_IHDR_MARKER[0]
            && bytes[13] == PNG_IHDR_MARKER[1]
            && bytes[14] == PNG_IHDR_MARKER[2]
            && bytes[15] == PNG_IHDR_MARKER[3]
        {
            // Width: bytes 16-19 (big-endian u32) / 宽度
            let width = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]) as f32;
            // Height: bytes 20-23 (big-endian u32) / 高度
            let height = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]) as f32;
            return Ok((width, height));
        }
        return Err(ERR_INVALID_PNG_IHDR);
    }

    // Check for JPEG signature / 检查 JPEG 签名
    if bytes[0] == 0xFF && bytes[1] == 0xD8 {
        // JPEG: Scan for SOF marker with bounds checking / JPEG：带边界检查地扫描 SOF 标记
        let mut offset = JPEG_INITIAL_OFFSET;
        let len = bytes.len();

        while offset + JPEG_MIN_SEGMENT_SIZE < len {
            if bytes[offset] != 0xFF {
                return Err(ERR_INVALID_JPG_MARKER);
            }

            let marker = bytes[offset + 1];
            let segment_len = u16::from_be_bytes([bytes[offset + 2], bytes[offset + 3]]) as usize;

            // SOF markers / SOF 标记
            if (JPEG_SOF_MARKER_START..=JPEG_SOF_MARKER_END).contains(&marker)
                && marker != JPEG_MARKER_DHT
                && marker != JPEG_MARKER_JPG
                && marker != JPEG_MARKER_DAC
            {
                let height = u16::from_be_bytes([bytes[offset + 5], bytes[offset + 6]]) as f32;
                let width = u16::from_be_bytes([bytes[offset + 7], bytes[offset + 8]]) as f32;
                return Ok((width, height));
            }

            offset += segment_len + 2;
        }
        return Err(ERR_NO_SOF_MARKER);
    }

    Err(ERR_UNKNOWN_FORMAT)
}

// Regex to find all rId patterns - compiled once / 正则表达式 - 仅编译一次
static REGEX: LazyLock<Regex> = LazyLock::new(|| Regex::new(REGEX_REL_ID).unwrap());

/// Parse relationship XML content to get next available rId / 解析关系 XML 内容以获取下一个可用的 rId
///
/// Scans all existing rId values and returns the next sequential ID / 扫描所有现有的 rId 值并返回下一个顺序 ID
///
/// # Arguments / 参数
/// * `rels_content` - XML content of .rels file / .rels 文件的 XML 内容
///
/// # Returns / 返回
/// Next available rId number / 下一个可用的 rId 编号
#[inline]
pub(crate) fn parse_next_rid_from_rels(rels_content: &str) -> u32 {
    let mut max_id = 0_u32;

    // Find all rId patterns and track the maximum / 查找所有 rId 模式并跟踪最大值
    for cap in REGEX.captures_iter(rels_content) {
        if let Some(id_match) = cap.get(1) {
            let id_str = id_match.as_str();
            // Fast path: parse directly from "rId xxx" / 快速路径：直接从 "rId xx" 解析
            if let Some(num_str) = id_str.strip_prefix(REL_ID_PREFIX)
                && let Ok(num) = num_str.parse::<u32>()
                && num > max_id
            {
                max_id = num;
            }
        }
    }

    max_id + 1
}

/// Flatten nested JSON structure into flat records / 将嵌套的 JSON 结构展平成扁平记录
///
/// Converts nested objects and arrays into a list of flat key-value maps / 将嵌套对象和数组转换为扁平键值映射列表
///
/// # Example / 示例
/// ```ignore
/// use serde_json::json;
///
/// let value = json!({"user": {"name": "Alice"}});
/// let records = flatten_json(&value);
/// assert_eq!(records.len(), 1);
/// ```
///
/// # Arguments / 参数
/// * `value` - JSON value to flatten / 要展平的 JSON 值
///
/// # Returns / 返回
/// Vector of flattened records / 展平记录的向量
pub(crate) fn flatten_json(value: &Value) -> Vec<HashMap<String, Value>> {
    if let Value::Object(obj) = value {
        // Pre-allocate with estimated capacity / 预分配估计容量
        let mut records = Vec::with_capacity(FLATTEN_RECORDS_CAPACITY);
        records.push(HashMap::with_capacity(obj.len()));

        // Process each key-value pair / 处理每个键值对
        for (key, val) in obj {
            let estimated_size = records.len().saturating_mul(2);
            let mut new_records = Vec::with_capacity(estimated_size.max(FLATTEN_RECORDS_CAPACITY));

            for mut record in records {
                match val {
                    // Arrays and objects are processed recursively / 数组和对象递归处理
                    Value::Array(arr) if !arr.is_empty() => {
                        for item in arr {
                            // Recursively flatten / 递归展平
                            for mut sub_record in flatten_json(item) {
                                merge_record_with_prefix(key, &mut record, &mut sub_record);
                                new_records.push(record.clone());
                            }
                        }
                    }
                    // Objects - recursive flattening / 对象 - 递归展平
                    Value::Object(_) => {
                        for mut sub_record in flatten_json(val) {
                            merge_record_with_prefix(key, &mut record, &mut sub_record);
                            new_records.push(record.clone());
                        }
                    }
                    // Primitive types - direct insert / 基本类型 - 直接插入
                    _ => {
                        record.insert(key.clone(), val.clone());
                        new_records.push(record);
                    }
                }
            }

            records = new_records;
        }

        records
    } else {
        // Fast path for non-objects / 非对象的快速路径
        vec![HashMap::new()]
    }
}

/// Merge record with prefixed keys into base record / 将带前缀键的记录合并到基础记录中
///
/// Moves all key-value pairs from `other` into `base` with a prefix  / 将所有键值对从 `other` 移动到 `base` 并添加前缀
///
/// # Arguments / 参数
/// * `prefix` - Prefix to add to keys / 要添加到键的前缀
/// * `base` - Base record to merge into / 要合并到的基础记录
/// * `other` - Record to merge from (will be drained) / 要合并的记录（将被清空）
#[inline]
fn merge_record_with_prefix(
    prefix: &str,
    base: &mut HashMap<String, Value>,
    other: &mut HashMap<String, Value>,
) {
    // Pre-allocate exact space needed / 预分配所需的精确空间
    let needed_capacity = base.len().saturating_add(other.len());
    if base.capacity() < needed_capacity {
        base.reserve(needed_capacity - base.len());
    }

    // Use with_capacity for string formatting / 为字符串格式化使用 with_capacity
    let prefix_len = prefix.len();

    for (k, v) in other.drain() {
        // Pre-allocate string capacity to avoid reallocations / 预分配字符串容量以避免重新分配
        let mut new_key = String::with_capacity(prefix_len + 1 + k.len());
        new_key.push_str(prefix);
        new_key.push('.');
        new_key.push_str(&k);
        base.insert(new_key, v);
    }
}
