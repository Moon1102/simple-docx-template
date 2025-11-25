use crate::core::constant::*;
use crate::core::default_handler::DefaultValueHandler;
use crate::core::docx_processor::DocxProcessor;
use crate::core::image_manager::ImageManager;
use crate::core::relationship_manager::RelationshipManager;
use crate::public::value_extern::ValueExt;
use async_zip::error::ZipError;
use async_zip::tokio::read::seek::ZipFileReader;
use async_zip::tokio::write::ZipFileWriter;
use async_zip::{Compression, ZipEntryBuilder};
use bytes::Bytes;
use serde_json::Value;
use std::collections::HashMap;
use std::env::temp_dir;
use std::marker::PhantomData;
use std::path::{Path, PathBuf};
use tokio::fs::{File as AsyncFile, create_dir_all, remove_file};
use tokio::io::{AsyncReadExt, BufReader, BufWriter};
use tokio_util::compat::{FuturesAsyncReadCompatExt, FuturesAsyncWriteCompatExt};
use uuid::Uuid;

/// Main DOCX processor struct / 主 DOCX 处理器结构体
pub struct DOCX<'a> {
    // DPI (dots per inch) for image rendering / 图片渲染的 DPI（每英寸点数）
    dpi: f32,

    // Custom cell value handler for placeholder replacement / 用于占位符替换的自定义单元格值处理器
    cell_handler: Option<Box<dyn ValueExt + Send>>,

    // Flag to skip w:t events during image processing / 在图片处理期间跳过 w:t 事件的标志
    skip_w_t_events: bool,

    // Phantom data for lifetime parameter / 生命周期参数的幽灵数据
    _marker: PhantomData<&'a ()>,
}

impl<'a> Default for DOCX<'a> {
    fn default() -> Self {
        Self {
            // Use default value handler / 使用默认值处理器
            cell_handler: Some(Box::new(DefaultValueHandler)),

            // Use default DPI constant / 使用默认 DPI 常量
            dpi: DEFAULT_DPI,

            // Initially not skipping w:t events / 初始时不跳过 w:t 事件
            skip_w_t_events: false,

            _marker: PhantomData,
        }
    }
}

impl<'a> DOCX<'a> {
    // Set custom DPI for image rendering / 设置图片渲染的自定义 DPI
    pub fn set_dpi(&mut self, dpi: f32) {
        self.dpi = dpi;
    }

    /// Set custom cell value handler / 设置自定义单元格值处理器
    /// # Arguments / 参数
    ///  * `handler` - Custom cell value handle / 自定义单元格处理器
    ///
    /// see [`DefaultValueHandler`]
    pub fn set_cell_handler(&mut self, handler: Box<dyn ValueExt + Send>) {
        self.cell_handler = Some(handler);
    }

    /// Single-pass processing of the DOCX file / DOCX 文件的单次处理
    ///
    /// Reads from input, processes XML, handles images, and writes to output / 从输入读取，处理 XML，处理图片，并写入输出
    ///
    /// # Arguments / 参数
    /// * `input_path` - Path to input DOCX file / 输入 DOCX 文件路径
    /// * `output_path` - Path to output DOCX file / 输出 DOCX 文件路径
    /// * `placeholders` - HashMap of placeholder values / 占位符值的 HashMap
    ///
    /// # Returns / 返回
    /// * `Result<(), ZipError>` - Success or zip error / 成功或 zip 错误
    pub async fn generate(
        &mut self,
        input_path: &str,
        output_path: &str,
        placeholders: &HashMap<String, Value>,
    ) -> Result<(), ZipError> {
        // Ensure output directory exists / 确保输出目录存在
        if let Some(parent_dir) = Path::new(output_path).parent() {
            create_dir_all(parent_dir).await?;
        }

        // Open input DOCX file as zip stream / 将输入 DOCX 文件作为 zip 流打开
        let input_file = AsyncFile::open(input_path).await?;
        let reader = BufReader::new(input_file);
        let mut zip_stream = ZipFileReader::with_tokio(reader).await?;

        // Create output DOCX file writer with buffering / 创建带缓冲的输出 DOCX 文件写入器
        let output_file = AsyncFile::create(output_path).await?;
        // // Wrap in BufWriter to optimize zip metadata writes / 包装在 BufWriter 中以优化 zip 元数据写入
        let buffered_output = BufWriter::new(output_file);
        let mut writer = ZipFileWriter::with_tokio(buffered_output);

        // Initialize managers for relationships and images / 初始化关系和图片管理器
        let mut rel_manager = RelationshipManager::new();
        let mut img_manager = ImageManager::new(self.dpi);

        // Store path to temporary document.xml file / 存储临时 document.xml 文件的路径
        let mut temp_doc_xml_path: Option<PathBuf> = None;

        // Process all entries in the input zip / 处理输入 zip 中的所有条目
        let entries_len = zip_stream.file().entries().len();
        for index in 0..entries_len {
            let entry = &zip_stream.file().entries()[index];
            let filename_owned = entry.filename().as_str()?.to_string();
            let filename_str = filename_owned.as_str();
            let entry_reader = zip_stream.reader_with_entry(index).await?;
            // Handle document relationships file / 处理文档关系文件
            if filename_str == RELS_PATH {
                let mut content = Vec::with_capacity(DEFAULT_BUFFER_SIZE);
                entry_reader.compat().read_to_end(&mut content).await?;
                // Store relationships for later processing (Bytes for zero-copy) / 存储关系以供后续处理（Bytes 实现零拷贝）
                rel_manager.set_initial_content(Bytes::from(content));
            } else if filename_str == DOCUMENT_XML_PATH {
                // Buffer to temp file to process later / 缓冲到临时文件以便后续处理
                let uuid = Uuid::now_v7().to_string();
                let tmp_path = temp_dir().join(format!(
                    "{}{}{}",
                    TEMP_FILE_PREFIX, uuid, TEMP_FILE_EXTENSION
                ));
                let mut tmp_file = AsyncFile::create(&tmp_path).await?;
                tokio::io::copy(&mut entry_reader.compat(), &mut tmp_file).await?;
                temp_doc_xml_path = Some(tmp_path);
            } else {
                // Write other files immediately (pass-through) / 立即写入其他文件（透传）
                // Load into memory to ensure correct decompression / 加载到内存以确保正确解压
                let mut content = Vec::with_capacity(DEFAULT_BUFFER_SIZE);
                entry_reader.compat().read_to_end(&mut content).await?;

                let options = ZipEntryBuilder::new(filename_owned.into(), Compression::Deflate);
                writer.write_entry_whole(options, &content).await?;
            }
        }

        // Now process document.xml if we found it / 如果找到了 document.xml，现在处理它
        if let Some(tmp_path) = temp_doc_xml_path {
            let options = ZipEntryBuilder::new(DOCUMENT_XML_PATH.into(), Compression::Deflate);
            let entry_writer = writer.write_entry_stream(options).await?;

            // Take ownership of cell handler / 获取单元格处理器的所有权
            let cell_handler = self
                .cell_handler
                .take()
                .unwrap_or(Box::from(DefaultValueHandler));

            let mut processor = DocxProcessor {
                cell_handler,
                skip_w_t_events: self.skip_w_t_events,
            };

            // Open temp file asynchronously for reading / 异步打开临时文件进行读取
            let file = AsyncFile::open(&tmp_path).await?;
            let mut buf_reader = BufReader::new(file);

            // Process XML events directly / 直接处理 XML 事件
            // Use compat_write() to convert futures AsyncWrite to tokio AsyncWrite if needed
            let mut compat_writer = entry_writer.compat_write();

            processor
                .process_xml_events(
                    &mut compat_writer,
                    &mut buf_reader,
                    placeholders,
                    &mut rel_manager,
                    &mut img_manager,
                )
                .await
                .map_err(|_| ZipError::FeatureNotSupported("XML processing failed"))?;

            // Restore cell handler / 恢复单元格处理器
            self.cell_handler = Some(processor.cell_handler);

            // Get back entry_writer and close it
            compat_writer.into_inner().close().await?;

            // Cleanup temp file after successful processing / 成功处理后清理临时文件
            remove_file(&tmp_path).await?;
        }

        // Write updated relationship file / 写入更新后的关系文件
        if let Some(rels_content) = rel_manager.generate_final_rels_content() {
            let options = ZipEntryBuilder::new(RELS_PATH.into(), Compression::Deflate);
            writer.write_entry_whole(options, &rels_content).await?;
        }

        // Write all new images to media folder / 将所有新图片写入媒体文件夹
        for (filename, (bytes, _)) in img_manager.get_images() {
            let path = format!("{}{}", MEDIA_PATH_PREFIX, filename);
            let options = ZipEntryBuilder::new(path.into(), Compression::Stored);
            writer.write_entry_whole(options, bytes).await?;
        }

        // Close output zip file / 关闭输出 zip 文件
        writer.close().await?;
        Ok(())
    }
}
