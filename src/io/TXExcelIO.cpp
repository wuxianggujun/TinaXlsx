//
// @file TXExcelIO.cpp
// @brief 🚀 Excel文件I/O处理器实现
//

#include <TinaXlsx/io/TXExcelIO.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <pugixml.hpp>
#include <fstream>
#include <filesystem>
#include <algorithm>
#include <sstream>

namespace TinaXlsx {

// ==================== 前向声明 ====================
static TXResult<std::unique_ptr<TXWorkbook>> loadXLSX(
    const std::string& file_path,
    const TXExcelIO::ReadOptions& options);

static TXResult<std::unique_ptr<TXWorkbook>> loadXLSXFromMemory(
    const void* data,
    size_t size,
    const TXExcelIO::ReadOptions& options);

static TXResult<void> saveXLSX(
    const TXWorkbook* workbook,
    const std::string& file_path,
    const TXExcelIO::WriteOptions& options);

static TXResult<TXVector<uint8_t>> saveXLSXToMemory(
    const TXWorkbook* workbook,
    const TXExcelIO::WriteOptions& options);

// ==================== 静态读取方法 ====================

TXResult<std::unique_ptr<TXWorkbook>> TXExcelIO::loadFromFile(
    const std::string& file_path,
    const ReadOptions& options) {
    
    TX_LOG_INFO("开始加载Excel文件: {}", file_path);
    
    // 验证文件路径
    auto path_result = validateFilePath(file_path, false);
    if (path_result.isError()) {
        return TXResult<std::unique_ptr<TXWorkbook>>(path_result.error());
    }
    
    // 检测文件格式
    FileFormat format = (options.read_formulas) ? detectFormat(file_path) : FileFormat::AUTO_DETECT;
    
    try {
        switch (format) {
            case FileFormat::XLSX:
                return loadXLSX(file_path, options);
            case FileFormat::CSV:
                return loadCSV(file_path);
            case FileFormat::AUTO_DETECT: {
                // 自动检测格式
                format = detectFormat(file_path);
                if (format == FileFormat::XLSX) {
                    return loadXLSX(file_path, options);
                } else if (format == FileFormat::CSV) {
                    return loadCSV(file_path);
                } else {
                    return TXResult<std::unique_ptr<TXWorkbook>>(
                        TXError(TXErrorCode::UnsupportedFormat, "不支持的文件格式")
                    );
                }
            }
            default:
                return TXResult<std::unique_ptr<TXWorkbook>>(
                    TXError(TXErrorCode::UnsupportedFormat, "不支持的文件格式")
                );
        }
    } catch (const std::exception& e) {
        return TXResult<std::unique_ptr<TXWorkbook>>(
            TXError(TXErrorCode::FileReadFailed, 
                   std::string("文件读取失败: ") + e.what())
        );
    }
}

TXResult<std::unique_ptr<TXWorkbook>> TXExcelIO::loadFromMemory(
    const void* data,
    size_t size,
    const ReadOptions& options) {
    
    TX_LOG_INFO("开始从内存加载Excel数据: {} 字节", size);
    
    if (!data || size == 0) {
        return TXResult<std::unique_ptr<TXWorkbook>>(
            TXError(TXErrorCode::InvalidArgument, "无效的内存数据")
        );
    }
    
    // 检测格式
    FileFormat format = detectFormat(data, size);
    
    try {
        switch (format) {
            case FileFormat::XLSX:
                return loadXLSXFromMemory(data, size, options);
            default:
                return TXResult<std::unique_ptr<TXWorkbook>>(
                    TXError(TXErrorCode::UnsupportedFormat, "不支持的内存数据格式")
                );
        }
    } catch (const std::exception& e) {
        return TXResult<std::unique_ptr<TXWorkbook>>(
            TXError(TXErrorCode::SerializationError, 
                   std::string("内存数据解析失败: ") + e.what())
        );
    }
}

// ==================== 静态写入方法 ====================

TXResult<void> TXExcelIO::saveToFile(
    const TXWorkbook* workbook,
    const std::string& file_path,
    const WriteOptions& options) {
    
    TX_LOG_INFO("开始保存Excel文件: {}", file_path);
    
    // 验证参数
    auto wb_result = validateWorkbook(workbook);
    if (wb_result.isError()) {
        return wb_result;
    }
    
    auto path_result = validateFilePath(file_path, true);
    if (path_result.isError()) {
        return path_result;
    }
    
    try {
        // 创建备份（如果文件已存在）
        if (std::filesystem::exists(file_path)) {
            auto backup_result = createBackup(file_path);
            if (backup_result.isError()) {
                TX_LOG_WARN("创建备份失败: {}", backup_result.error().getMessage());
            }
        }
        
        switch (options.format) {
            case FileFormat::XLSX:
                return saveXLSX(workbook, file_path, options);
            case FileFormat::CSV:
                return saveCSV(workbook, file_path, 0);
            default:
                return TXResult<void>(
                    TXError(TXErrorCode::UnsupportedFormat, "不支持的输出格式")
                );
        }
    } catch (const std::exception& e) {
        return TXResult<void>(
            TXError(TXErrorCode::FileWriteFailed, 
                   std::string("文件保存失败: ") + e.what())
        );
    }
}

TXResult<TXVector<uint8_t>> TXExcelIO::saveToMemory(
    const TXWorkbook* workbook,
    const WriteOptions& options) {
    
    TX_LOG_INFO("开始保存Excel到内存");
    
    // 验证工作簿
    auto wb_result = validateWorkbook(workbook);
    if (wb_result.isError()) {
        return TXResult<TXVector<uint8_t>>(wb_result.error());
    }
    
    try {
        switch (options.format) {
            case FileFormat::XLSX:
                return saveXLSXToMemory(workbook, options);
            default:
                return TXResult<TXVector<uint8_t>>(
                    TXError(TXErrorCode::UnsupportedFormat, "不支持的内存输出格式")
                );
        }
    } catch (const std::exception& e) {
        return TXResult<TXVector<uint8_t>>(
            TXError(TXErrorCode::SerializationError, 
                   std::string("内存保存失败: ") + e.what())
        );
    }
}

// ==================== 格式检测 ====================

TXExcelIO::FileFormat TXExcelIO::detectFormat(const std::string& file_path) {
    std::string ext = getFileExtension(file_path);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    if (ext == ".xlsx") {
        return FileFormat::XLSX;
    } else if (ext == ".xls") {
        return FileFormat::XLS;
    } else if (ext == ".csv") {
        return FileFormat::CSV;
    }
    
    // 尝试通过文件内容检测
    std::ifstream file(file_path, std::ios::binary);
    if (!file.is_open()) {
        return FileFormat::XLSX; // 默认
    }
    
    // 读取文件头
    char header[8];
    file.read(header, 8);
    
    // XLSX文件是ZIP格式，以PK开头
    if (header[0] == 'P' && header[1] == 'K') {
        return FileFormat::XLSX;
    }
    
    // XLS文件有特定的OLE头
    if (header[0] == '\xD0' && header[1] == '\xCF') {
        return FileFormat::XLS;
    }
    
    return FileFormat::CSV; // 默认为CSV
}

TXExcelIO::FileFormat TXExcelIO::detectFormat(const void* data, size_t size) {
    if (!data || size < 8) {
        return FileFormat::XLSX; // 默认
    }
    
    const char* bytes = static_cast<const char*>(data);
    
    // XLSX文件是ZIP格式，以PK开头
    if (bytes[0] == 'P' && bytes[1] == 'K') {
        return FileFormat::XLSX;
    }
    
    // XLS文件有特定的OLE头
    if (bytes[0] == '\xD0' && bytes[1] == '\xCF') {
        return FileFormat::XLS;
    }
    
    return FileFormat::CSV; // 默认为CSV
}

bool TXExcelIO::isValidExcelFile(const std::string& file_path) {
    if (!std::filesystem::exists(file_path)) {
        return false;
    }
    
    FileFormat format = detectFormat(file_path);
    return (format == FileFormat::XLSX || format == FileFormat::XLS);
}

// ==================== 便捷方法 ====================

TXResult<std::unique_ptr<TXWorkbook>> TXExcelIO::loadCSV(
    const std::string& file_path,
    char delimiter) {
    
    TX_LOG_INFO("开始加载CSV文件: {}", file_path);
    
    try {
        std::ifstream file(file_path);
        if (!file.is_open()) {
            return TXResult<std::unique_ptr<TXWorkbook>>(
                TXError(TXErrorCode::FileReadFailed, "无法打开CSV文件")
            );
        }
        
        // 创建工作簿
        auto workbook = TXWorkbook::create("CSV_Import");
        auto sheet = workbook->getSheet(0);
        
        std::string line;
        uint32_t row = 1;
        
        while (std::getline(file, line) && row <= 1048576) {
            std::stringstream ss(line);
            std::string cell_value;
            uint32_t col = 1;
            
            while (std::getline(ss, cell_value, delimiter) && col <= 16384) {
                if (!cell_value.empty()) {
                    // 尝试解析为数字
                    try {
                        double num_value = std::stod(cell_value);
                        // 🔧 修复：cell()方法期望0基坐标，所以需要减1
                        sheet->cell(row - 1, col - 1).setValue(num_value);
                    } catch (...) {
                        // 作为字符串处理
                        // 🔧 修复：cell()方法期望0基坐标，所以需要减1
                        sheet->cell(row - 1, col - 1).setValue(cell_value);
                    }
                }
                ++col;
            }
            ++row;
        }
        
        TX_LOG_INFO("CSV加载完成: {} 行数据", row - 1);
        return TXResult<std::unique_ptr<TXWorkbook>>(std::move(workbook));
        
    } catch (const std::exception& e) {
        return TXResult<std::unique_ptr<TXWorkbook>>(
            TXError(TXErrorCode::FileReadFailed, 
                   std::string("CSV读取失败: ") + e.what())
        );
    }
}

TXResult<void> TXExcelIO::saveCSV(
    const TXWorkbook* workbook,
    const std::string& file_path,
    size_t sheet_index,
    char delimiter) {
    
    TX_LOG_INFO("开始保存CSV文件: {}", file_path);
    
    auto wb_result = validateWorkbook(workbook);
    if (wb_result.isError()) {
        return wb_result;
    }
    
    if (sheet_index >= workbook->getSheetCount()) {
        return TXResult<void>(
            TXError(TXErrorCode::InvalidArgument, "工作表索引超出范围")
        );
    }
    
    try {
        std::ofstream file(file_path);
        if (!file.is_open()) {
            return TXResult<void>(
                TXError(TXErrorCode::FileWriteFailed, "无法创建CSV文件")
            );
        }
        
        // 获取非const指针以便访问cell方法
        auto sheet = const_cast<TXWorkbook*>(workbook)->getSheet(sheet_index);
        auto range = sheet->getUsedRange();

        if (range.isEmpty()) {
            TX_LOG_INFO("工作表为空，创建空CSV文件");
            return TXResult<void>();
        }

        auto start = range.getStart();
        auto end = range.getEnd();

        TX_LOG_INFO("CSV保存范围: {}:{} 到 {}:{}",
                    start.getRow().index(), start.getCol().index(),
                    end.getRow().index(), end.getCol().index());

        for (uint32_t row = start.getRow().index(); row <= end.getRow().index(); ++row) {
            bool first_col = true;
            for (uint32_t col = start.getCol().index(); col <= end.getCol().index(); ++col) {
                if (!first_col) {
                    file << delimiter;
                }
                first_col = false;

                // 🔧 修复：cell()方法期望0基坐标，所以需要减1
                auto cell_value = sheet->cell(row - 1, col - 1).getValue();
                // TX_LOG_INFO("CSV写入单元格 {}:{} = {} (类型: {})",
                //             row, col,
                //             cell_value.getType() == TXVariant::Type::String ? cell_value.getString() :
                //             (cell_value.getType() == TXVariant::Type::Number ? std::to_string(cell_value.getNumber()) : "Empty"),
                //             static_cast<int>(cell_value.getType()));
                if (cell_value.getType() == TXVariant::Type::String) {
                    // 处理包含分隔符的字符串
                    std::string str_val = cell_value.getString();
                    if (str_val.find(delimiter) != std::string::npos || 
                        str_val.find('"') != std::string::npos ||
                        str_val.find('\n') != std::string::npos) {
                        // 需要引号包围
                        file << '"';
                        for (char c : str_val) {
                            if (c == '"') {
                                file << "\"\""; // 转义引号
                            } else {
                                file << c;
                            }
                        }
                        file << '"';
                    } else {
                        file << str_val;
                    }
                } else if (cell_value.getType() == TXVariant::Type::Number) {
                    file << cell_value.getNumber();
                }
            }
            file << '\n';
        }
        
        TX_LOG_INFO("CSV保存完成");
        return TXResult<void>();

    } catch (const std::exception& e) {
        return TXResult<void>(
            TXError(TXErrorCode::FileWriteFailed,
                   std::string("CSV保存失败: ") + e.what())
        );
    }
}

// ==================== XLSX处理方法 ====================

static TXResult<std::unique_ptr<TXWorkbook>> loadXLSX(
    const std::string& file_path,
    const TXExcelIO::ReadOptions& options) {

    TX_LOG_INFO("开始解析XLSX文件: {}", file_path);

    // TODO: 实现XLSX解析
    // 这里需要使用ZIP库解压XLSX文件，然后解析XML

    // 临时实现：创建一个示例工作簿
    auto workbook = TXWorkbook::create("XLSX_Loaded");
    auto sheet = workbook->getSheet(0);

    // 添加一些示例数据表明文件已"加载"
    sheet->cell("A1").setValue("从XLSX加载的数据");
    sheet->cell("B1").setValue(42.0);

    TX_LOG_WARN("XLSX加载功能暂未完全实现，返回示例数据");
    return TXResult<std::unique_ptr<TXWorkbook>>(std::move(workbook));
}

static TXResult<std::unique_ptr<TXWorkbook>> loadXLSXFromMemory(
    const void* data,
    size_t size,
    const TXExcelIO::ReadOptions& options) {

    TX_LOG_INFO("开始解析XLSX内存数据: {} 字节", size);

    // TODO: 实现从内存解析XLSX
    auto workbook = TXWorkbook::create("XLSX_Memory_Loaded");

    TX_LOG_WARN("XLSX内存加载功能暂未完全实现");
    return TXResult<std::unique_ptr<TXWorkbook>>(std::move(workbook));
}

static TXResult<void> saveXLSX(
    const TXWorkbook* workbook,
    const std::string& file_path,
    const TXExcelIO::WriteOptions& options) {

    TX_LOG_INFO("开始保存XLSX文件: {}", file_path);

    // TODO: 实现XLSX保存
    // 这里需要生成OpenXML格式的文件并压缩为ZIP

    // 临时实现：创建一个简单的文件表明保存成功
    std::ofstream file(file_path);
    if (!file.is_open()) {
        return TXResult<void>(
            TXError(TXErrorCode::FileWriteFailed, "无法创建XLSX文件")
        );
    }

    file << "TinaXlsx XLSX Placeholder - " << workbook->getName() << std::endl;
    file << "工作表数量: " << workbook->getSheetCount() << std::endl;

    TX_LOG_WARN("XLSX保存功能暂未完全实现，创建了占位文件");
    return TXResult<void>();
}

static TXResult<TXVector<uint8_t>> saveXLSXToMemory(
    const TXWorkbook* workbook,
    const TXExcelIO::WriteOptions& options) {

    TX_LOG_INFO("开始保存XLSX到内存");

    auto& memory_manager = GlobalUnifiedMemoryManager::getInstance();
    TXVector<uint8_t> data(memory_manager);

    // TODO: 实现XLSX内存保存
    std::string placeholder = "TinaXlsx XLSX Memory Placeholder";
    data.reserve(placeholder.size());
    for (char c : placeholder) {
        data.push_back(static_cast<uint8_t>(c));
    }

    TX_LOG_WARN("XLSX内存保存功能暂未完全实现");
    return TXResult<TXVector<uint8_t>>(std::move(data));
}

// ==================== 内部辅助方法 ====================

TXResult<void> TXExcelIO::validateFilePath(const std::string& file_path, bool for_writing) {
    if (file_path.empty()) {
        return TXResult<void>(
            TXError(TXErrorCode::InvalidArgument, "文件路径为空")
        );
    }

    if (!for_writing) {
        // 读取模式：检查文件是否存在
        if (!std::filesystem::exists(file_path)) {
            return TXResult<void>(
                TXError(TXErrorCode::FileNotFound, "文件不存在: " + file_path)
            );
        }

        if (!std::filesystem::is_regular_file(file_path)) {
            return TXResult<void>(
                TXError(TXErrorCode::InvalidArgument, "不是有效的文件: " + file_path)
            );
        }
    } else {
        // 写入模式：检查目录是否存在
        std::filesystem::path path(file_path);
        auto parent_dir = path.parent_path();

        if (!parent_dir.empty() && !std::filesystem::exists(parent_dir)) {
            try {
                std::filesystem::create_directories(parent_dir);
            } catch (const std::exception& e) {
                return TXResult<void>(
                    TXError(TXErrorCode::FileWriteFailed,
                           std::string("无法创建目录: ") + e.what())
                );
            }
        }
    }

    return TXResult<void>();
}

TXResult<void> TXExcelIO::validateWorkbook(const TXWorkbook* workbook) {
    if (!workbook) {
        return TXResult<void>(
            TXError(TXErrorCode::InvalidArgument, "工作簿指针为空")
        );
    }

    if (!workbook->isValid()) {
        return TXResult<void>(
            TXError(TXErrorCode::InvalidArgument, "工作簿状态无效")
        );
    }

    if (workbook->isEmpty()) {
        TX_LOG_WARN("工作簿为空");
    }

    return TXResult<void>();
}

std::string TXExcelIO::getFileExtension(const std::string& file_path) {
    std::filesystem::path path(file_path);
    return path.extension().string();
}

TXResult<void> TXExcelIO::createBackup(const std::string& file_path) {
    try {
        std::string backup_path = file_path + ".backup";
        std::filesystem::copy_file(file_path, backup_path,
                                  std::filesystem::copy_options::overwrite_existing);
        TX_LOG_DEBUG("创建备份文件: {}", backup_path);
        return TXResult<void>();
    } catch (const std::exception& e) {
        return TXResult<void>(
            TXError(TXErrorCode::FileWriteFailed,
                   std::string("创建备份失败: ") + e.what())
        );
    }
}

void TXExcelIO::handleError(const std::string& operation, const TXError& error) {
    TX_LOG_ERROR("TXExcelIO操作失败: {} - 错误: {}", operation, error.getMessage());
}

} // namespace TinaXlsx
