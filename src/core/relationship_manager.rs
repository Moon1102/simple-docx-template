use crate::core::constant::{
    REL_ID_PREFIX, REL_TYPE_IMAGE, REL_XML_BASE_CAPACITY, TYPICAL_IMAGE_COUNT,
};
use crate::core::utils::parse_next_rid_from_rels;
use bytes::{Bytes, BytesMut};
use std::str::from_utf8;

/// Manager for DOCX document relationships (.rels file) / DOCX 文档关系（.rels 文件）管理器
///
/// Handles relationship IDs for images and other resources, and generates updated relationship XML / 处理图片和其他资源的关系 ID，并生成更新的关系 XML
pub(crate) struct RelationshipManager {
    current_rid: u32,      // Next available relationship ID / 下一个可用的关系 ID
    new_rels: Vec<String>, // New relationships to add (pre-allocated) / 要添加的新关系（预分配）
    original_rels_content: Option<Bytes>, // Original .rels file content (zero-copy) / 原始 .rels 文件内容（零拷贝）
}

impl RelationshipManager {
    /// Create new relationship manager / 创建新的关系管理器
    ///
    /// Pre-allocates space for typical number of images / 为典型图片数量预分配空间
    #[inline]
    pub(crate) fn new() -> Self {
        Self {
            current_rid: 1,
            new_rels: Vec::with_capacity(TYPICAL_IMAGE_COUNT),
            original_rels_content: None,
        }
    }

    /// Set initial relationship file content / 设置初始关系文件内容
    ///
    /// Parses existing relationships to determine next available ID / 解析现有关系以确定下一个可用 ID
    ///
    /// # Arguments / 参数
    /// * `content` - Original .rels file bytes / 原始 .rels 文件字节
    #[inline]
    pub(crate) fn set_initial_content(&mut self, content: Bytes) {
        // Fast path: parse existing relationships / 快速路径：解析现有关系
        if let Ok(rels_str) = from_utf8(&content) {
            self.current_rid = parse_next_rid_from_rels(rels_str);
        }
        self.original_rels_content = Some(content);
    }

    /// Add new image relationship / 添加新的图片关系
    ///
    /// Generates unique relationship ID and registers the image  / 生成唯一的关系 ID 并注册图片
    ///
    /// # Arguments / 参数
    /// * `filename` - Image filename (e.g., "image_123.png") / 图片文件名（例如 "image_123.png"）
    ///
    /// # Returns / 返回
    /// * `(rel_id, image_id)` - Relationship ID and numeric ID / 关系 ID 和数字 ID
    #[inline]
    pub(crate) fn add_image_relationship(&mut self, filename: &str) -> (String, u32) {
        let image_id = self.current_rid;

        let mut rel_id = String::with_capacity(8);
        rel_id.push_str(REL_ID_PREFIX);
        rel_id.push_str(&self.current_rid.to_string());

        self.current_rid += 1;

        // Base XML template is ~150 chars + filename length / 基础 XML 模板约 150 字符 + 文件名长度
        let capacity = REL_XML_BASE_CAPACITY + filename.len();
        let mut rel_xml = String::with_capacity(capacity);

        rel_xml.push_str(r#"<Relationship Id=""#);
        rel_xml.push_str(&rel_id);
        rel_xml.push_str(r#"" Type=""#);
        rel_xml.push_str(REL_TYPE_IMAGE);
        rel_xml.push_str(r#"" Target="media/"#);
        rel_xml.push_str(filename);
        rel_xml.push_str(r#""/>"#);

        self.new_rels.push(rel_xml);

        (rel_id, image_id)
    }

    /// Generate final relationship file content / 生成最终的关系文件内容
    ///
    /// Merges new relationships into original content / 将新关系合并到原始内容中
    ///
    /// # Returns / 返回
    /// * `Some(bytes)` - Updated .rels file content (zero-copy) / 更新的 .rels 文件内容（零拷贝）
    /// * `None` - If no original content was set / 如果未设置原始内容
    pub(crate) fn generate_final_rels_content(&self) -> Option<Bytes> {
        let content = self.original_rels_content.as_ref()?;

        // Fast path: if no new relationships, return cloned Bytes (cheap) / 快速路径：如果没有新关系，返回克隆的 Bytes（廉价）
        if self.new_rels.is_empty() {
            return Some(content.clone()); // Bytes::clone is cheap (reference counting)
        }

        let rels_str = from_utf8(content).ok()?;

        // Find insertion point / 查找插入点
        let insert_pos = rels_str.rfind("</Relationships>")?;

        // Calculate exact capacity needed / 计算所需的精确容量
        let new_rels_total_len: usize = self.new_rels.iter().map(|s| s.len() + 5).sum(); // +5 for "\n    "
        let final_capacity = rels_str.len() + new_rels_total_len + 10; // +10 for safety margin

        // Use BytesMut for efficient building, then freeze to Bytes / 使用 BytesMut 高效构建，然后冻结为 Bytes
        let mut buffer = BytesMut::with_capacity(final_capacity);

        // Build final content efficiently / 高效构建最终内容
        buffer.extend_from_slice(&rels_str.as_bytes()[..insert_pos]);
        buffer.extend_from_slice(b"\n    ");

        for rel in &self.new_rels {
            buffer.extend_from_slice(rel.as_bytes());
            buffer.extend_from_slice(b"\n    ");
        }

        buffer.extend_from_slice(&rels_str.as_bytes()[insert_pos..]);

        Some(buffer.freeze())
    }
}
