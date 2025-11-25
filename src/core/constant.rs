// ---------- Buffer size constants / 缓冲区大小常量 ----------

// Default buffer size for reading/writing operations (8KB) / 读写操作的默认缓冲区大小（8KB）
pub(crate) const DEFAULT_BUFFER_SIZE: usize = 8192;

// Preview buffer size for lookahead operations (4KB) / 预览缓冲区大小，用于前瞻操作（4KB）
pub(crate) const PREVIEW_BUFFER_SIZE: usize = 4096;

// ---------- Image dimension constants / 图片尺寸常量 ----------

// Minimum image data length / 最小的图片数据长度
pub(crate) const MIN_IMAGE_DATA_LEN: usize = 24;

// Maximum image size: 5cm = 1800000 EMU / 最大图片尺寸：5厘米 = 1800000 EMU
pub(crate) const MAX_EMU: f32 = 1800000.0;

// Default image width: 2cm / 默认图片宽度：2厘米
pub(crate) const DEFAULT_WIDTH_EMU: f32 = 720000.0;

// Default image height: 2.5cm / 默认图片高度：2.5厘米
pub(crate) const DEFAULT_HEIGHT_EMU: f32 = 900000.0;

// EMU (English Metric Units) per inch conversion factor / 每英寸的 EMU（英制公制单位）转换因子
pub(crate) const EMU_PER_INCH: f32 = 914400.0;

// Default DPI (dots per inch) for image rendering / 图片渲染的默认 DPI（每英寸点数）
pub(crate) const DEFAULT_DPI: f32 = 96.0;

// ---------- DOCX file path constants / DOCX 文件路径常量 ----------

// Path to document relationships file / 文档关系文件路径
pub(crate) const RELS_PATH: &str = "word/_rels/document.xml.rels";

// Path to main document XML file / 主文档 XML 文件路径
pub(crate) const DOCUMENT_XML_PATH: &str = "word/document.xml";

// Path prefix for media files / 媒体文件路径前缀
pub(crate) const MEDIA_PATH_PREFIX: &str = "word/media/";

// Temporary file name prefix / 临时文件名前缀
pub(crate) const TEMP_FILE_PREFIX: &str = "docx_";

// Temporary file extension / 临时文件扩展名
pub(crate) const TEMP_FILE_EXTENSION: &str = ".xml";

// ---------- XML element name constants / XML 元素名称常量 ----------

// Table element name / 表格元素名称
pub(crate) const XML_TABLE: &str = "w:tbl";

// Text element name / 文本元素名称
pub(crate) const XML_TEXT: &[u8] = b"w:t";

// Table row element name / 表格行元素名称
pub(crate) const XML_TABLE_ROW: &[u8] = b"w:tr";

// Table cell element name / 表格单元格元素名称
pub(crate) const XML_TABLE_CELL: &[u8] = b"w:tc";

// Table cell properties element name / 表格单元格属性元素名称
pub(crate) const XML_TABLE_CELL_PROPERTIES: &str = "w:tcPr";

// Table cell v_merge tag / 表格合并标记
pub(crate) const XML_TABLE_MERGE_TAG: &str = "w:vMerge w:val";

// ---------- Image format detection constants / 图片格式检测常量 ----------

// PNG image base64 signature / PNG 图片的 base64 签名
pub(crate) const PNG_BASE64_SIGNATURE: &str = "iVBORw0KGgo";

// JPEG image base64 signature / JPEG 图片的 base64 签名
pub(crate) const JPEG_BASE64_SIGNATURE: &str = "/9j/";

// ---------- Merge type constants / 合并类型常量 ----------

// Vertical merge restart value / 垂直合并重新开始值
pub(crate) const MERGE_RESTART: u32 = 1;

// Vertical merge continue value / 垂直合并继续值
pub(crate) const MERGE_CONTINUE: u32 = 0;

// Vertical merge restart type string / 垂直合并重新开始类型字符串
pub(crate) const MERGE_TYPE_RESTART: &str = "restart";

// Vertical merge continue type string / 垂直合并继续类型字符串
pub(crate) const MERGE_TYPE_CONTINUE: &str = "continue";

// ---------- Image format detection constants / 图片格式检测常量（扩展）----------

// PNG file signature bytes / PNG 文件签名字节
#[allow(dead_code)]
pub(crate) const PNG_SIGNATURE: [u8; 4] = [0x89, b'P', b'N', b'G'];

// Alternative PNG signature for numeric check / PNG 签名的数字形式
pub(crate) const PNG_SIG_BYTE_0: u8 = 137;
pub(crate) const PNG_SIG_BYTE_1: u8 = 80;
pub(crate) const PNG_SIG_BYTE_2: u8 = 78;
pub(crate) const PNG_SIG_BYTE_3: u8 = 71;

// JPEG file signature bytes / JPEG 文件签名字节
#[allow(dead_code)]
pub(crate) const JPEG_SIGNATURE: [u8; 3] = [0xFF, 0xD8, 0xFF];

// PNG IHDR chunk marker / PNG IHDR 块标记
pub(crate) const PNG_IHDR_MARKER: [u8; 4] = [b'I', b'H', b'D', b'R'];

// Default image file extensions / 默认图片文件扩展名
pub(crate) const IMAGE_EXT_PNG: &str = "png";
pub(crate) const IMAGE_EXT_JPEG: &str = "jpg";

// Image filename prefix / 图片文件名前缀
pub(crate) const IMAGE_FILENAME_PREFIX: &str = "image_";

// ---------- Capacity hint constants / 容量提示常量 ----------

// Typical number of images in a document / 文档中典型的图片数量
pub(crate) const TYPICAL_IMAGE_COUNT: usize = 8;

// Capacity for image filename / 图片文件名容量
pub(crate) const IMAGE_FILENAME_CAPACITY: usize = 50;

// Capacity for relationship XML / 关系 XML 容量
pub(crate) const REL_XML_BASE_CAPACITY: usize = 150;

// Capacity for drawing XML / 绘图 XML 容量
pub(crate) const DRAWING_XML_CAPACITY: usize = 850;

// Typical table row event count / 典型表格行事件数
pub(crate) const TYPICAL_ROW_EVENT_COUNT: usize = 20;

// Typical header row count / 典型标题行数
pub(crate) const TYPICAL_HEADER_ROW_COUNT: usize = 5;

// Typical data row count / 典型数据行数
pub(crate) const TYPICAL_DATA_ROW_COUNT: usize = 50;

// Typical other events count / 典型其他事件数
pub(crate) const TYPICAL_OTHER_EVENT_COUNT: usize = 20;

// Typical column value count / 典型列值数
pub(crate) const TYPICAL_COLUMN_COUNT: usize = 10;

// Estimated flatten records size / 估计的展平记录大小
pub(crate) const FLATTEN_RECORDS_CAPACITY: usize = 4;

// Picture name capacity / 图片名称容量
pub(crate) const PICTURE_NAME_CAPACITY: usize = 20;

// ---------- XML namespace constants / XML 命名空间常量 ----------

// DrawingML namespace / DrawingML 命名空间
pub(crate) const XMLNS_DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

// Picture namespace / 图片命名空间
pub(crate) const XMLNS_PICTURE: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";

// Image relationship type / 图片关系类型
pub(crate) const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

// ---------- Template marker constants / 模板标记常量 ----------

// Loop start marker / 循环开始标记
pub(crate) const LOOP_START_MARKER: &str = "{{#";

// Loop end marker / 循环结束标记
pub(crate) const LOOP_END_MARKER: &str = "}}";

// Relationship ID prefix / 关系 ID 前缀
pub(crate) const REL_ID_PREFIX: &str = "rId";

// ---------- Drawing XML attribute constants / 绘图 XML 属性常量 ----------

// Drawing distance values / 绘图距离值
pub(crate) const DRAWING_DIST_TOP: &str = "0";
pub(crate) const DRAWING_DIST_BOTTOM: &str = "0";
pub(crate) const DRAWING_DIST_LEFT: &str = "114300";
pub(crate) const DRAWING_DIST_RIGHT: &str = "114300";

// Effect extent values / 效果范围值
pub(crate) const EFFECT_EXTENT_LEFT: &str = "0";
pub(crate) const EFFECT_EXTENT_TOP: &str = "0";
pub(crate) const EFFECT_EXTENT_RIGHT: &str = "24765";
pub(crate) const EFFECT_EXTENT_BOTTOM: &str = "24130";

// Lock attributes / 锁定属性
pub(crate) const NO_CHANGE_ASPECT: &str = "1";

// Coordinate values / 坐标值
pub(crate) const COORD_ZERO: &str = "0";

// ---------- JPEG marker constants / JPEG 标记常量 ----------

// JPEG SOF marker range / JPEG SOF 标记范围
pub(crate) const JPEG_SOF_MARKER_START: u8 = 0xC0;
pub(crate) const JPEG_SOF_MARKER_END: u8 = 0xCF;

// JPEG excluded SOF markers / JPEG 排除的 SOF 标记
pub(crate) const JPEG_MARKER_DHT: u8 = 0xC4; // Define Huffman Table
pub(crate) const JPEG_MARKER_JPG: u8 = 0xC8; // JPG extension
pub(crate) const JPEG_MARKER_DAC: u8 = 0xCC; // Define Arithmetic Coding

// JPEG segment offset / JPEG 段偏移量
pub(crate) const JPEG_INITIAL_OFFSET: usize = 2;
pub(crate) const JPEG_MIN_SEGMENT_SIZE: usize = 9;

// ---------- Error message constants / 错误消息常量 ----------

pub(crate) const ERR_BASE64_DECODE: &str = "Failed convert Base64 data to image";
pub(crate) const ERR_PICTURE_NAME: &str = "Failed generate picture name";
pub(crate) const ERR_NESTED_TABLE: &str = "nested table";
#[allow(dead_code)]
pub(crate) const ERR_XML_PROCESSING: &str = "XML processing failed";
pub(crate) const ERR_SLICE_TOO_SHORT: &str = "Byte slice too short";
pub(crate) const ERR_INVALID_PNG_IHDR: &str = "Invalid PNG IHDR chunk";
pub(crate) const ERR_INVALID_JPG_MARKER: &str = "Invalid JPG marker";
pub(crate) const ERR_NO_SOF_MARKER: &str = "No SOF marker found in JPG";
pub(crate) const ERR_UNKNOWN_FORMAT: &str = "Unknown image format";

// ---------- Regex pattern constants / 正则表达式模式常量 ----------

// Placeholder detection pattern / 占位符检测模式
pub(crate) const REGEX_PLACEHOLDER: &str = r"\S(.+?)]";

// Relationship ID pattern / 关系 ID 模式
pub(crate) const REGEX_REL_ID: &str = r#"Id="(rId\d+)""#;

// ---------- Image description constants / 图片描述常量 ----------

pub(crate) const DEFAULT_IMAGE_DESCRIPTION: &str = "Generated Image";
pub(crate) const IMAGE_NAME_PREFIX: &str = "Picture ";
