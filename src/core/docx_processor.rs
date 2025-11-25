use crate::core::constant::{
    DEFAULT_BUFFER_SIZE, DEFAULT_IMAGE_DESCRIPTION, ERR_NESTED_TABLE, ERR_PICTURE_NAME,
    IMAGE_NAME_PREFIX, JPEG_BASE64_SIGNATURE, LOOP_END_MARKER, LOOP_START_MARKER, MERGE_CONTINUE,
    MERGE_RESTART, MERGE_TYPE_CONTINUE, MERGE_TYPE_RESTART, PICTURE_NAME_CAPACITY,
    PNG_BASE64_SIGNATURE, PREVIEW_BUFFER_SIZE, REGEX_PLACEHOLDER, TYPICAL_COLUMN_COUNT,
    TYPICAL_DATA_ROW_COUNT, TYPICAL_HEADER_ROW_COUNT, TYPICAL_OTHER_EVENT_COUNT,
    TYPICAL_ROW_EVENT_COUNT, XML_TABLE, XML_TABLE_CELL, XML_TABLE_CELL_PROPERTIES,
    XML_TABLE_MERGE_TAG, XML_TABLE_ROW, XML_TEXT,
};
use crate::core::image_manager::ImageManager;
use crate::core::relationship_manager::RelationshipManager;
use crate::core::utils::flatten_json;
use crate::public::value_extern::ValueExt;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use regex::Regex;
use serde_json::Value;
use std::collections::HashMap;
use std::fmt::Write;
use std::sync::LazyLock;
use tokio::io::{AsyncBufRead, AsyncWrite, AsyncWriteExt};

/// Regex pattern for placeholder detection / 用于占位符检测的正则表达式模式
///
/// Matches patterns like [key] in text / 匹配文本中的 [key] 模式
pub(crate) static REGEX: LazyLock<Regex> = LazyLock::new(|| Regex::new(REGEX_PLACEHOLDER).unwrap());

/// Table content structure / 表格内容结构
struct TableContent<'a> {
    header_rows: Vec<Vec<Event<'a>>>,
    data_rows: Vec<Event<'a>>,
    other_events: Vec<Event<'a>>,
    first_col: Option<String>, // First column placeholder key / 第一列占位符键
}

/// XML processor running in blocking thread / 在阻塞线程中运行的 XML 处理器
pub(crate) struct DocxProcessor {
    // Custom cell value handler / 自定义单元格值处理器
    pub(crate) cell_handler: Box<dyn ValueExt + Send>,

    // Flag to skip w:t events during image processing / 在图片处理期间跳过 w:t 事件的标志
    pub(crate) skip_w_t_events: bool,
}

impl DocxProcessor {
    /// Process XML events and replace placeholders / 处理 XML 事件并替换占位符
    ///
    /// This is the core XML processing method that streams through the document / 流式处理文档的核心 XML 处理方法
    ///
    /// # Arguments / 参数
    /// * `writer` - XML writer for output / 用于输出的 XML 写入器
    /// * `reader` - Buffered reader for input XML / 用于输入 XML 的缓冲读取器
    /// * `placeholders` - Placeholder values to replace / 要替换的占位符值
    /// * `rel_manager` - Relationship manager / 关系管理器
    /// * `img_manager` - Image manager / 图片管理器
    pub(crate) async fn process_xml_events<'a, W, R>(
        &mut self,
        writer: &mut W,
        reader: &mut R,
        placeholders: &HashMap<String, Value>,
        rel_manager: &mut RelationshipManager,
        img_manager: &mut ImageManager<'a>,
    ) -> Result<(), quick_xml::Error>
    where
        W: AsyncWrite + Unpin,
        R: AsyncBufRead + Unpin,
    {
        // Create XML writer wrapping the output writer / 创建包装输出写入器的 XML 写入器
        let mut xml_writer = Writer::new(writer);
        let mut reader = Reader::from_reader(reader);

        // Buffers for XML event processing / XML 事件处理的缓冲区
        let buf = &mut Vec::with_capacity(DEFAULT_BUFFER_SIZE);
        let preview_buf = &mut Vec::with_capacity(PREVIEW_BUFFER_SIZE);

        // State tracking variables / 状态跟踪变量
        let mut inside_text_tag = false; // Currently inside w:t tag / 当前在 w:t 标签内
        let mut skip_current_event = false; // Skip writing current event / 跳过写入当前事件
        let mut pending_event: Option<Event> = None; // Lookahead event / 前瞻事件

        // Main event processing loop / 主事件处理循环
        loop {
            // Get next event (either pending or read new) / 获取下一个事件（待处理或读取新的）
            let event = if let Some(e) = pending_event.take() {
                e
            } else {
                reader.read_event_into_async(buf).await?
            };

            match event {
                // Start tag event / 开始标签事件
                Event::Start(e) => {
                    // Handle table elements / 处理表格元素
                    if e.name().as_ref() == XML_TABLE.as_bytes() {
                        self.process_table(
                            &mut reader,
                            &mut xml_writer,
                            buf,
                            placeholders,
                            rel_manager,
                            img_manager,
                        )
                        .await?;
                    } else {
                        // Handle text elements / 处理文本元素
                        if e.name().as_ref() == XML_TEXT {
                            // Skip if we're in image processing mode / 如果在图片处理模式则跳过
                            if self.skip_w_t_events {
                                continue;
                            }

                            // Check if text contains base64 image / 检查文本是否包含 base64 图片
                            let mut is_base64_image = false;
                            let mut base64_data = None;
                            preview_buf.clear();
                            {
                                // Peek at next event to check for image / 查看下一个事件以检查图片
                                match reader.read_event_into_async(preview_buf).await {
                                    Ok(Event::Text(text)) => {
                                        // Replace placeholders in text / 替换文本中的占位符
                                        let replaced = self
                                            .cell_handler
                                            .replace(&text.decode()?, placeholders);

                                        // Check for image signatures / 检查图片签名
                                        if replaced.starts_with(PNG_BASE64_SIGNATURE)
                                            || replaced.starts_with(JPEG_BASE64_SIGNATURE)
                                        {
                                            is_base64_image = true;
                                            base64_data = Some(replaced);
                                        } else {
                                            // Not an image, save for later processing / 不是图片，保存以供后续处理
                                            pending_event = Some(Event::Text(text.into_owned()));
                                        }
                                    }
                                    Ok(e) => {
                                        pending_event = Some(e.into_owned());
                                    }
                                    Err(e) => return Err(e),
                                };
                            }

                            // Process base64 image if detected / 如果检测到 base64 图片则处理
                            if is_base64_image {
                                self.skip_w_t_events = true;
                                inside_text_tag = false;
                                if let Some(base64_str) = base64_data {
                                    self.process_base64_image(
                                        &base64_str,
                                        &mut xml_writer,
                                        rel_manager,
                                        img_manager,
                                    )
                                    .await?;
                                }
                                self.skip_w_t_events = false;
                                continue; // Skip normal text processing / 跳过正常文本处理
                            } else {
                                inside_text_tag = true; // Enter text tag / 进入文本标签
                            }
                        }
                        // Write start tag if not skipped / 如果未跳过则写入开始标签
                        if skip_current_event {
                            skip_current_event = false;
                        } else {
                            xml_writer.write_event_async(Event::Start(e)).await?;
                        }
                    }
                }
                // Text content event / 文本内容事件
                Event::Text(text) => {
                    // Skip if in image processing mode / 如果在图片处理模式则跳过
                    if self.skip_w_t_events && inside_text_tag {
                        continue;
                    }
                    // Replace placeholders in text tags / 替换文本标签中的占位符
                    if inside_text_tag {
                        let replaced = self.cell_handler.replace(&text.decode()?, placeholders);
                        xml_writer
                            .write_event_async(Event::Text(BytesText::from_escaped(replaced)))
                            .await?;
                    } else {
                        // Pass through non-text-tag content / 传递非文本标签内容
                        xml_writer.write_event_async(Event::Text(text)).await?;
                    }
                }
                // End tag event / 结束标签事件
                Event::End(e) => {
                    // Reset state when exiting text tag / 退出文本标签时重置状态
                    if e.name().as_ref() == XML_TEXT {
                        inside_text_tag = false;
                        self.skip_w_t_events = false;
                    }
                    // Skip if in image processing mode / 如果在图片处理模式则跳过
                    if self.skip_w_t_events {
                        continue;
                    }
                    xml_writer.write_event_async(Event::End(e)).await?;
                }
                // End of file / 文件结束
                Event::Eof => break,
                // Pass through all other events / 传递所有其他事件
                _ => xml_writer.write_event_async(event).await?,
            }
            buf.clear(); // Clear buffer for next event / 清空缓冲区以处理下一个事件
        }
        Ok(())
    }

    /// Process base64 image and insert into document / 处理 base64 图片并插入文档
    ///
    /// Decodes base64 image data and generates XML drawing elements / 解码 base64 图片数据并生成 XML 绘图元素
    #[inline]
    async fn process_base64_image<'a, W>(
        &mut self,
        base64_data: &str,
        writer: &mut Writer<W>,
        rel_manager: &mut RelationshipManager,
        img_manager: &mut ImageManager<'a>,
    ) -> Result<(), quick_xml::Error>
    where
        W: AsyncWrite + Unpin,
    {
        // Try to process base64 image data / 尝试处理 base64 图片数据
        if let Ok((rel_id, image_id, width, height)) =
            img_manager.process_base64(base64_data, rel_manager)
        {
            let mut name = String::with_capacity(PICTURE_NAME_CAPACITY);
            write!(&mut name, "{}{}", IMAGE_NAME_PREFIX, image_id).map_err(|_e| {
                quick_xml::errors::IllFormedError::UnmatchedEndTag(ERR_PICTURE_NAME.to_string())
            })?;

            // Generate XML drawing markup for the image / 为图片生成 XML 绘图标记
            let xml_inner = ImageManager::generate_xml_drawing_inner(
                &rel_id,
                image_id,
                width,
                height,
                &name,
                DEFAULT_IMAGE_DESCRIPTION,
            );
            // Write XML directly to output / 直接将 XML 写入输出
            writer.get_mut().write_all(xml_inner.as_bytes()).await?;
        }
        Ok(())
    }

    /// Process table element and handle dynamic rows / 处理表格元素并处理动态行
    ///
    /// Tables can contain placeholder arrays that generate multiple rows / 表格可以包含生成多行的占位符数组
    #[inline]
    async fn process_table<'a, R, W>(
        &mut self,
        reader: &mut Reader<R>,
        writer: &mut Writer<W>,
        buf: &mut Vec<u8>,
        placeholders: &HashMap<String, Value>,
        rel_manager: &mut RelationshipManager,
        img_manager: &mut ImageManager<'a>,
    ) -> Result<(), quick_xml::Error>
    where
        R: AsyncBufRead + Unpin,
        W: AsyncWrite + Unpin,
    {
        // Collect all table content (headers, data rows, properties) / 收集所有表格内容（标题、数据行、属性）
        let table_content = Self::collect_table_content(reader, buf).await?;

        // Write table start tag / 写入表格开始标签
        writer
            .write_event_async(Event::Start(BytesStart::new(XML_TABLE)))
            .await?;

        // Write table properties and other non-row elements / 写入表格属性和其他非行元素
        for event in table_content.other_events {
            writer.write_event_async(event).await?;
        }

        let table_key = table_content.first_col;
        // Check if table has dynamic data (array placeholder) / 检查表格是否有动态数据（数组占位符）
        if let Some(table_key) = &table_key
            && let Some(Value::Array(list)) = placeholders.get(table_key)
            && !table_content.data_rows.is_empty()
        {
            // Write header rows / 写入标题行
            for mut header_row in table_content.header_rows {
                for event in header_row.drain(..) {
                    writer.write_event_async(event).await?;
                }
            }

            // Flatten JSON array and generate rows with merging / 展平 JSON 数组并生成带合并的行
            let items = list.iter().flat_map(flatten_json).collect::<Vec<_>>();
            self.write_rows_with_merge(
                writer,
                &table_content.data_rows,
                items.into_iter(),
                rel_manager,
                img_manager,
            )
            .await?;
        } else {
            for mut header_row in table_content.header_rows {
                for event in header_row.drain(..) {
                    match event {
                        Event::Text(text) => {
                            let replaced = self.cell_handler.replace(&text.decode()?, placeholders);
                            if replaced.starts_with("iVBORw0KGgo") || replaced.starts_with("/9j/") {
                                self.process_base64_image(
                                    replaced.as_str(),
                                    writer,
                                    rel_manager,
                                    img_manager,
                                )
                                .await?;
                            } else {
                                writer
                                    .write_event_async(Event::Text(BytesText::from_escaped(
                                        replaced,
                                    )))
                                    .await?;
                            }
                        }
                        _ => writer.write_event_async(event).await?,
                    }
                }
            }
        }

        writer
            .write_event_async(Event::End(BytesEnd::new(XML_TABLE)))
            .await?;
        Ok(())
    }

    /// Collect and categorize table content into headers and data rows / 收集并分类表格内容为标题行和数据行
    ///
    /// Separates rows with placeholders (data rows) from rows without (header rows) / 将包含占位符的行（数据行）与不包含的行（标题行）分离
    #[inline]
    async fn collect_table_content<R>(
        reader: &mut Reader<R>,
        buf: &mut Vec<u8>,
    ) -> Result<TableContent<'static>, quick_xml::Error>
    where
        R: AsyncBufRead + Unpin,
    {
        // Storage for different table components / 不同表格组件的存储
        let mut header_rows = Vec::with_capacity(TYPICAL_HEADER_ROW_COUNT);
        let mut data_rows = Vec::with_capacity(TYPICAL_DATA_ROW_COUNT);
        let mut other_events = Vec::with_capacity(TYPICAL_OTHER_EVENT_COUNT);
        let mut table_key = None; // First column placeholder key / 第一列占位符键

        // Read all table events / 读取所有表格事件
        loop {
            buf.clear();
            match reader.read_event_into_async(buf).await {
                // Nested tables not supported / 不支持嵌套表格
                Ok(Event::Start(e)) if e.name().as_ref() == XML_TABLE.as_bytes() => {
                    return Err(quick_xml::errors::IllFormedError::UnmatchedEndTag(
                        ERR_NESTED_TABLE.to_string(),
                    )
                    .into());
                }
                // Process table row / 处理表格行
                Ok(Event::Start(e)) if e.name().as_ref() == XML_TABLE_ROW => {
                    let start_owned = e.into_owned();
                    let (row_events, has_placeholder) = Self::process_table_row_internal(
                        reader,
                        buf,
                        Event::Start(start_owned),
                        &mut table_key,
                    )
                    .await?;

                    // Categorize row based on placeholder presence / 根据是否包含占位符对行进行分类
                    if has_placeholder {
                        data_rows = row_events; // Data template row / 数据模板行
                    } else {
                        header_rows.push(row_events); // Header row / 标题行
                    }
                }
                // End of table / 表格结束
                Ok(Event::End(e)) if e.name().as_ref() == XML_TABLE.as_bytes() => {
                    break;
                }
                // Collect other events (properties, etc.) / 收集其他事件（属性等）
                Ok(e) => {
                    other_events.push(e.into_owned());
                }
                Err(e) => return Err(e),
            }
        }

        Ok(TableContent {
            header_rows,
            data_rows,
            other_events,
            first_col: table_key,
        })
    }

    /// Process a single table row and detect placeholders / 处理单个表格行并检测占位符
    ///
    /// Returns row events and whether the row contains placeholders / 返回行事件以及该行是否包含占位符
    #[inline]
    async fn process_table_row_internal<R>(
        reader: &mut Reader<R>,
        buf: &mut Vec<u8>,
        start_event: Event<'static>,
        table_key: &mut Option<String>,
    ) -> Result<(Vec<Event<'static>>, bool), quick_xml::Error>
    where
        R: AsyncBufRead + Unpin,
    {
        // Storage for row events and state / 行事件和状态的存储
        let mut row_events = Vec::with_capacity(TYPICAL_ROW_EVENT_COUNT);
        row_events.push(start_event);
        let mut has_placeholder = false; // Track if row contains placeholders / 跟踪行是否包含占位符
        let mut row_depth = 1; // Track nesting depth for nested rows / 跟踪嵌套行的深度
        let mut is_first_text = true; // Track first text element / 跟踪第一个文本元素

        // Process all events in the row / 处理行中的所有事件
        loop {
            buf.clear();
            match reader.read_event_into_async(buf).await {
                // Handle row start tags / 处理行开始标签
                Ok(Event::Start(row_e)) => {
                    if row_e.name().as_ref() == XML_TABLE_ROW {
                        row_depth += 1; // Track nesting / 跟踪嵌套
                    }
                    row_events.push(Event::Start(row_e.into_owned()));
                }
                // Handle row end tags / 处理行结束标签
                Ok(Event::End(row_e)) => {
                    if row_e.name().as_ref() == XML_TABLE_ROW {
                        row_depth -= 1;
                        if row_depth == 0 {
                            // End of this row / 此行结束
                            row_events.push(Event::End(row_e.into_owned()));
                            break;
                        }
                    }
                    row_events.push(Event::End(row_e.into_owned()));
                }
                // Handle text content / 处理文本内容
                Ok(Event::Text(row_e)) => {
                    let text = row_e.decode()?;
                    // Check for placeholder pattern / 检查占位符模式
                    if REGEX.is_match(&text) {
                        has_placeholder = true;
                    }

                    // Extract table key from first text if it's a loop marker / 如果是循环标记，从第一个文本提取表格键
                    if is_first_text
                        && table_key.is_none()
                        && text.starts_with(LOOP_START_MARKER)
                        && let Some(pos) = text.find(LOOP_END_MARKER)
                    {
                        let first_col = &text[..pos + 2];
                        let first = text.replace(first_col, "");
                        *table_key = Some(first_col.to_string());

                        row_events.push(Event::Text(BytesText::from_escaped(first)));
                    } else {
                        row_events.push(Event::Text(row_e.into_owned()));
                    }

                    if is_first_text {
                        is_first_text = false
                    }
                }
                Ok(Event::Eof) => break,
                Ok(row_e) => {
                    row_events.push(row_e.into_owned());
                }
                Err(e) => return Err(e),
            }
        }

        Ok((row_events, has_placeholder))
    }

    /// Write table rows with vertical cell merging / 写入带垂直单元格合并的表格行
    ///
    /// Handles automatic cell merging for consecutive rows with identical values / 处理具有相同值的连续行的自动单元格合并
    #[inline]
    async fn write_rows_with_merge<'a, W, I>(
        &mut self,
        writer: &mut Writer<W>,
        row_template: &[Event<'a>],
        items: I,
        rel_manager: &mut RelationshipManager,
        img_manager: &mut ImageManager<'a>,
    ) -> Result<(), quick_xml::Error>
    where
        W: AsyncWrite + Unpin,
        I: Iterator<Item = HashMap<String, Value>>,
    {
        // Initialize iteration state / 初始化迭代状态
        let mut iter = items.peekable(); // Peekable to look ahead / 可窥视以便前瞻
        let mut prev_row_values: Option<Vec<String>> = None; // Previous row values for comparison / 用于比较的前一行值
        let mut merging_cols: Vec<bool> = Vec::new(); // Track which columns are currently merging / 跟踪当前正在合并的列
        let mut row_index = 0; // Current row index / 当前行索引

        // Process each data row / 处理每个数据行
        while let Some(item) = iter.next() {
            // Compute current row values by replacing placeholders / 通过替换占位符计算当前行值
            // Pre-allocate based on previous row or estimate / 根据前一行或估计预分配
            let capacity = prev_row_values
                .as_ref()
                .map(|v| v.len())
                .unwrap_or(TYPICAL_COLUMN_COUNT);
            let mut current_values = Vec::with_capacity(capacity);
            for event in row_template.iter() {
                if let Event::Text(text) = event {
                    let replaced =
                        self.cell_handler
                            .replace_in_table(row_index, &text.decode()?, &item);
                    current_values.push(replaced);
                }
            }

            // Initialize merging_cols on first row / 在第一行初始化 merging_cols
            if merging_cols.is_empty() {
                merging_cols = vec![false; current_values.len()];
            }

            // Peek next row values for merge detection / 窥视下一行值以检测合并
            let next_values = if let Some(next_item) = iter.peek() {
                // Pre-allocate with known capacity / 使用已知容量预分配
                let mut values = Vec::with_capacity(current_values.len());
                for event in row_template.iter() {
                    if let Event::Text(text) = event {
                        let replaced = self.cell_handler.replace_in_table(
                            row_index + 1,
                            &text.decode()?,
                            next_item,
                        );
                        values.push(replaced);
                    }
                }
                Some(values)
            } else {
                None // No next row / 没有下一行
            };

            // Determine merge info for current row / 确定当前行的合并信息
            // None = no merge, Some(0) = continue merge, Some(1) = restart merge
            // None = 无合并, Some(0) = 继续合并, Some(1) = 重新开始合并
            let mut merge_info = vec![None; current_values.len()];

            // Check each column for merge state / 检查每列的合并状态
            for (col_idx, val) in current_values.iter().enumerate() {
                if col_idx >= merging_cols.len() {
                    break; // Safety check / 安全检查
                }

                // Get previous and next values for comparison / 获取前一个和下一个值进行比较
                let prev_val = prev_row_values.as_ref().and_then(|v| v.get(col_idx));
                let next_val = next_values.as_ref().and_then(|v| v.get(col_idx));

                // Optimized merge state logic with pattern matching / 使用模式匹配优化的合并状态逻辑
                match (merging_cols[col_idx], prev_val, next_val) {
                    // Currently merging and same as previous - continue merge / 当前在合并且与前一个相同 - 继续合并
                    (true, Some(p), _) if p == val => {
                        merge_info[col_idx] = Some(MERGE_CONTINUE);
                        // merging_cols[col_idx] remains true / merging_cols[col_idx] 保持为 true
                    }
                    // Start new merge (when next equals current and not empty) / 开始新合并（当下一个等于当前且非空）
                    (_, _, Some(n)) if n == val && !val.is_empty() => {
                        merge_info[col_idx] = Some(MERGE_RESTART);
                        merging_cols[col_idx] = true;
                    }
                    // No merge / 无合并
                    _ => {
                        merging_cols[col_idx] = false;
                    }
                }
            }

            // Write row with merge information / 使用合并信息写入行
            self.write_row_with_merge_fixed(
                writer,
                row_template,
                &item,
                &merge_info,
                row_index,
                rel_manager,
                img_manager,
            )
            .await?;

            // Update state for next iteration / 更新状态以供下次迭代
            prev_row_values = Some(current_values);
            row_index += 1;
        }

        Ok(())
    }

    /// Write a single row with merge information / 使用合并信息写入单行
    ///
    /// Applies vertical merge markers to cells based on merge state / 根据合并状态将垂直合并标记应用于单元格
    #[inline]
    #[allow(clippy::too_many_arguments)]
    async fn write_row_with_merge_fixed<'a, W>(
        &mut self,
        writer: &mut Writer<W>,
        row: &[Event<'a>],
        item: &HashMap<String, Value>,
        merge_info: &[Option<u32>],
        row_index: usize,
        rel_manager: &mut RelationshipManager,
        img_manager: &mut ImageManager<'a>,
    ) -> Result<(), quick_xml::Error>
    where
        W: AsyncWrite + Unpin,
    {
        // Track cell position and merge state / 跟踪单元格位置和合并状态
        let mut tc_index: i32 = -1; // Current cell index / 当前单元格索引
        let mut in_tc = false; // Inside table cell / 在表格单元格内
        let mut current_tc_is_continue = false; // Current cell is continuation of merge / 当前单元格是合并的延续

        // Process all events in row / 处理行中的所有事件
        for event in row {
            match event {
                // Handle start tags / 处理开始标签
                Event::Start(bytes_start) => {
                    // Borrow from bytes_start instead of cloning event / 从 bytes_start 借用而不是克隆事件
                    writer
                        .write_event_async(Event::Start(bytes_start.borrow()))
                        .await?;

                    // Handle table cell start / 处理表格单元格开始
                    if bytes_start.name().as_ref() == XML_TABLE_CELL {
                        in_tc = true;
                        tc_index += 1;
                        let merge_val = merge_info.get(tc_index as usize).and_then(|&v| v);

                        // Add merge properties if needed / 如果需要添加合并属性
                        if let Some(span) = merge_val {
                            let merge_type = if span == MERGE_RESTART {
                                MERGE_TYPE_RESTART
                            } else {
                                MERGE_TYPE_CONTINUE
                            };
                            let merge_tag =
                                format!(r#"<{}="{}"/>"#, XML_TABLE_MERGE_TAG, merge_type);
                            writer
                                .write_event_async(Event::Start(BytesStart::new(
                                    XML_TABLE_CELL_PROPERTIES,
                                )))
                                .await?;
                            writer.get_mut().write_all(merge_tag.as_bytes()).await?;
                            writer
                                .write_event_async(Event::End(BytesEnd::new(
                                    XML_TABLE_CELL_PROPERTIES,
                                )))
                                .await?;

                            // Mark as continuation cell (skip content) / 标记为延续单元格（跳过内容）
                            if span == MERGE_CONTINUE {
                                current_tc_is_continue = true;
                            }
                        }
                    }
                }
                // Handle text content / 处理文本内容
                Event::Text(text) => {
                    // Skip text in continuation cells / 跳过延续单元格中的文本
                    if in_tc && current_tc_is_continue {
                        // skip
                    } else {
                        // Replace placeholders and handle images / 替换占位符并处理图片
                        let replaced =
                            self.cell_handler
                                .replace_in_table(row_index, &text.decode()?, item);
                        // Check for base64 image / 检查 base64 图片
                        if replaced.starts_with(PNG_BASE64_SIGNATURE)
                            || replaced.starts_with(JPEG_BASE64_SIGNATURE)
                        {
                            self.process_base64_image(
                                replaced.as_str(),
                                writer,
                                rel_manager,
                                img_manager,
                            )
                            .await?;
                        } else {
                            writer
                                .write_event_async(Event::Text(BytesText::from_escaped(replaced)))
                                .await?;
                        }
                    }
                }
                // Handle end tags / 处理结束标签
                Event::End(bytes_end) => {
                    if bytes_end.name().as_ref() == XML_TABLE_CELL {
                        in_tc = false;
                        current_tc_is_continue = false;
                    }
                    // Borrow from bytes_end instead of cloning / 从 bytes_end 借用而不是克隆
                    writer
                        .write_event_async(Event::End(bytes_end.borrow()))
                        .await?;
                }
                // Pass through other events / 传递其他事件
                other => {
                    // For other event types, we need to borrow / 对于其他事件类型，我们需要借用
                    writer.write_event_async(other.borrow()).await?;
                }
            }
        }
        Ok(())
    }
}
