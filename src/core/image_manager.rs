use crate::core::constant::{
    COORD_ZERO, DEFAULT_HEIGHT_EMU, DEFAULT_WIDTH_EMU, DRAWING_DIST_BOTTOM, DRAWING_DIST_LEFT,
    DRAWING_DIST_RIGHT, DRAWING_DIST_TOP, DRAWING_XML_CAPACITY, EFFECT_EXTENT_BOTTOM,
    EFFECT_EXTENT_LEFT, EFFECT_EXTENT_RIGHT, EFFECT_EXTENT_TOP, EMU_PER_INCH, ERR_BASE64_DECODE,
    IMAGE_EXT_JPEG, IMAGE_EXT_PNG, IMAGE_FILENAME_CAPACITY, IMAGE_FILENAME_PREFIX, MAX_EMU,
    NO_CHANGE_ASPECT, TYPICAL_IMAGE_COUNT, XMLNS_DRAWINGML, XMLNS_PICTURE,
};
use crate::core::relationship_manager::RelationshipManager;
use crate::core::utils::get_image_dimensions;
use base64::Engine;
use base64::engine::general_purpose;
use bytes::Bytes;
use std::collections::HashMap;
use uuid::Uuid;

/// Manager for handling images in DOCX documents / DOCX 文档中图片处理的管理器
///
/// Tracks all images to be embedded, generates unique filenames, and creates XML markup for image display / 跟踪所有要嵌入的图片，生成唯一文件名，并创建图片显示的 XML 标记
pub(crate) struct ImageManager<'a> {
    dpi: f32,                                  // DPI for size calculation / 用于尺寸计算的 DPI
    images: HashMap<String, (Bytes, &'a str)>, // Pre-allocated hashmap (zero-copy) / 预分配的哈希映射（零拷贝）
}

impl<'a> ImageManager<'a> {
    /// Create new image manager / 创建新的图片管理器
    ///
    /// Pre-allocates space for typical number of images / 为典型图片数量预分配空间
    ///
    /// # Arguments / 参数
    /// * `dpi` - DPI for image size calculation / 用于图片尺寸计算的 DPI
    #[inline]
    pub(crate) fn new(dpi: f32) -> Self {
        Self {
            dpi,
            images: HashMap::with_capacity(TYPICAL_IMAGE_COUNT),
        }
    }

    /// Get all managed images / 获取所有管理的图片
    #[inline]
    pub(crate) fn get_images(&self) -> &HashMap<String, (Bytes, &'a str)> {
        &self.images
    }

    /// Process base64 image data and prepare for embedding / 处理 base64 图片数据并准备嵌入
    ///
    /// Decodes base64, detects format, generates unique filename, calculates dimensions, and registers with relationship manager / 解码 base64，检测格式，生成唯一文件名，计算尺寸，并在关系管理器中注册
    ///
    /// # Arguments / 参数
    /// * `base64_data` - Base64 encoded image data / Base64 编码的图片数据
    /// * `rel_manager` - Relationship manager / 关系管理器
    ///
    /// # Returns / 返回
    /// * `Ok((rel_id, image_id, width_emu, height_emu))` - Image info / 图片信息
    /// * `Err` - If base64 decode fails / 如果 base64 解码失败
    pub(crate) fn process_base64(
        &mut self,
        base64_data: &str,
        rel_manager: &mut RelationshipManager,
    ) -> Result<(String, u32, u32, u32), quick_xml::Error> {
        let image_bytes = general_purpose::STANDARD.decode(base64_data).map_err(|_| {
            quick_xml::errors::IllFormedError::UnmatchedEndTag(ERR_BASE64_DECODE.to_string())
        })?;

        // Fast format detection / 快速格式检测
        let extension = if image_bytes.len() >= 4
            && image_bytes[0] == 0x89
            && image_bytes[1] == b'P'
            && image_bytes[2] == b'N'
            && image_bytes[3] == b'G'
        {
            IMAGE_EXT_PNG
        } else if image_bytes.len() >= 3
            && image_bytes[0] == 0xFF
            && image_bytes[1] == 0xD8
            && image_bytes[2] == 0xFF
        {
            IMAGE_EXT_JPEG
        } else {
            IMAGE_EXT_PNG // Safe default / 安全默认值
        };

        // Generate unique filename / 生成唯一文件名
        let uuid = Uuid::now_v7();
        let mut filename = String::with_capacity(IMAGE_FILENAME_CAPACITY);
        filename.push_str(IMAGE_FILENAME_PREFIX);
        filename.push_str(&uuid.to_string());
        filename.push('.');
        filename.push_str(extension);

        // Register image in relationship manager / 在关系管理器中注册图片
        let (rel_id, image_id) = rel_manager.add_image_relationship(&filename);

        // Calculate image dimensions with fast path / 使用快速路径计算图片尺寸
        let (mut width_emu, mut height_emu) = match get_image_dimensions(&image_bytes) {
            Ok((width_px, height_px)) => {
                let dpi_inv = 1.0 / self.dpi;
                (
                    width_px * EMU_PER_INCH * dpi_inv,
                    height_px * EMU_PER_INCH * dpi_inv,
                )
            }
            Err(_) => (DEFAULT_WIDTH_EMU, DEFAULT_HEIGHT_EMU),
        };

        // Scale down if needed / 如果需要缩小
        let scale = (width_emu / MAX_EMU).max(height_emu / MAX_EMU);
        if scale > 1.0 {
            let scale_inv = 1.0 / scale;
            width_emu *= scale_inv;
            height_emu *= scale_inv;
        }

        // Store image bytes (zero-copy via Bytes) / 存储图片字节（通过 Bytes 零拷贝）
        self.images.insert(filename, (Bytes::from(image_bytes), ""));

        Ok((
            rel_id,
            image_id,
            width_emu.round() as u32,
            height_emu.round() as u32,
        ))
    }

    /// Generate OOXML markup for inline image / 生成内联图片的 OOXML 标记
    ///
    /// Creates complete XML structure for displaying an image inline in the document / 创建用于在文档中内联显示图片的完整 XML 结构
    ///
    /// # Arguments / 参数
    /// * `relationship_id` - Relationship ID (e.g., "rId5") / 关系 ID（例如 "rId5"）
    /// * `image_id` - Unique image ID / 唯一图片 ID
    /// * `width` - Width in EMU / 宽度（EMU）
    /// * `height` - Height in EMU / 高度（EMU）
    /// * `name` - Image name / 图片名称
    /// * `descr` - Image description / 图片描述
    ///
    /// # Returns / 返回
    /// Complete XML string for the image / 图片的完整 XML 字符串
    #[inline]
    pub(crate) fn generate_xml_drawing_inner(
        relationship_id: &str,
        image_id: u32,
        width: u32,
        height: u32,
        name: &str,
        descr: &str,
    ) -> String {
        let doc_pr_id = image_id;

        let capacity =
            DRAWING_XML_CAPACITY + relationship_id.len() + name.len() * 2 + descr.len() * 2;
        let mut xml = String::with_capacity(capacity);

        // Build XML string efficiently / 高效构建 XML 字符串
        xml.push_str(r#"<w:r><w:drawing><wp:inline distT=""#);
        xml.push_str(DRAWING_DIST_TOP);
        xml.push_str(r#"" distB=""#);
        xml.push_str(DRAWING_DIST_BOTTOM);
        xml.push_str(r#"" distL=""#);
        xml.push_str(DRAWING_DIST_LEFT);
        xml.push_str(r#"" distR=""#);
        xml.push_str(DRAWING_DIST_RIGHT);
        xml.push_str(r#""><wp:extent cx=""#);
        xml.push_str(&width.to_string());
        xml.push_str(r#"" cy=""#);
        xml.push_str(&height.to_string());
        xml.push_str(r#""/><wp:effectExtent l=""#);
        xml.push_str(EFFECT_EXTENT_LEFT);
        xml.push_str(r#"" t=""#);
        xml.push_str(EFFECT_EXTENT_TOP);
        xml.push_str(r#"" r=""#);
        xml.push_str(EFFECT_EXTENT_RIGHT);
        xml.push_str(r#"" b=""#);
        xml.push_str(EFFECT_EXTENT_BOTTOM);
        xml.push_str(r#""/><wp:docPr id=""#);
        xml.push_str(&doc_pr_id.to_string());
        xml.push_str(r#"" name=""#);
        xml.push_str(name);
        xml.push_str(r#"" descr=""#);
        xml.push_str(descr);
        xml.push_str(r#""/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=""#);
        xml.push_str(XMLNS_DRAWINGML);
        xml.push_str(r#"" noChangeAspect=""#);
        xml.push_str(NO_CHANGE_ASPECT);
        xml.push_str(r#""/></wp:cNvGraphicFramePr><a:graphic xmlns:a=""#);
        xml.push_str(XMLNS_DRAWINGML);
        xml.push_str(r#""><a:graphicData uri=""#);
        xml.push_str(XMLNS_PICTURE);
        xml.push_str(r#""><pic:pic xmlns:pic=""#);
        xml.push_str(XMLNS_PICTURE);
        xml.push_str(r#""><pic:nvPicPr><pic:cNvPr id=""#);
        xml.push_str(&doc_pr_id.to_string());
        xml.push_str(r#"" name=""#);
        xml.push_str(name);
        xml.push_str(r#"" descr=""#);
        xml.push_str(descr);
        xml.push_str(r#""/><pic:cNvPicPr><a:picLocks noChangeAspect=""#);
        xml.push_str(NO_CHANGE_ASPECT);
        xml.push_str(r#""/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=""#);
        xml.push_str(relationship_id);
        xml.push_str(
            r#""/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=""#,
        );
        xml.push_str(COORD_ZERO);
        xml.push_str(r#"" y=""#);
        xml.push_str(COORD_ZERO);
        xml.push_str(r#""/><a:ext cx=""#);
        xml.push_str(&width.to_string());
        xml.push_str(r#"" cy=""#);
        xml.push_str(&height.to_string());
        xml.push_str(r#""/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#);

        xml
    }
}
