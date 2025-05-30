#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <memory>
#include <functional>
#include <mutex>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 错误级别枚举
 */
enum class TXErrorLevel : int {
    Info = 0,           // 信息
    Warning = 1,        // 警告
    Error = 2,          // 错误
    Fatal = 3           // 致命错误
};

/**
 * @brief 错误代码枚举
 */
enum class TXErrorCode : int {
    Success = 0,                    // 成功，无错误
    
    // 通用错误 (1-99)
    Unknown = 1,                    // 未知错误
    InvalidArgument = 2,            // 无效参数
    NullPointer = 3,                // 空指针
    OutOfRange = 4,                 // 超出范围
    InvalidOperation = 5,           // 无效操作
    MemoryAllocation = 6,           // 内存分配失败
    
    // 文件I/O错误 (100-199)
    FileNotFound = 100,             // 文件未找到
    FileOpenFailed = 101,           // 文件打开失败
    FileWriteFailed = 102,          // 文件写入失败
    FileReadFailed = 103,           // 文件读取失败
    InvalidFileFormat = 104,        // 无效文件格式
    
    // 坐标相关错误 (200-299)
    InvalidCoordinate = 200,        // 无效坐标
    InvalidRow = 201,               // 无效行号
    InvalidColumn = 202,            // 无效列号
    InvalidRange = 203,             // 无效范围
    
    // 单元格相关错误 (300-399)
    CellNotFound = 300,             // 单元格不存在
    InvalidCellValue = 301,         // 无效单元格值
    CellTypeConversion = 302,       // 单元格类型转换失败
    
    // 工作表相关错误 (400-499)
    SheetNotFound = 400,            // 工作表不存在
    SheetNameExists = 401,          // 工作表名称已存在
    InvalidSheetName = 402,         // 无效工作表名称
    
    // 样式相关错误 (500-599)
    InvalidStyle = 500,             // 无效样式
    StyleNotFound = 501,            // 样式不存在
    StyleRegistrationFailed = 502,  // 样式注册失败
    
    // XML处理错误 (600-699)
    XmlParseError = 600,            // XML解析错误
    XmlGenerationError = 601,       // XML生成错误
    InvalidXmlStructure = 602,      // 无效XML结构
    
    // ZIP处理错误 (700-799)
    ZipCreateFailed = 700,          // ZIP创建失败
    ZipOpenFailed = 701,            // ZIP打开失败
    ZipExtractFailed = 702,         // ZIP解压失败
    ZipCompressionFailed = 703      // ZIP压缩失败
};

/**
 * @brief 错误信息结构
 */
struct TXErrorInfo {
    TXErrorCode code;
    TXErrorLevel level;
    std::string message;
    std::string context;           // 错误发生的上下文
    std::string function;          // 发生错误的函数名
    std::string file;              // 发生错误的文件名
    int line;                      // 发生错误的行号
    std::chrono::system_clock::time_point timestamp;
    
    TXErrorInfo() 
        : code(TXErrorCode::Success)
        , level(TXErrorLevel::Info)
        , line(0)
        , timestamp(std::chrono::system_clock::now()) {}
        
    TXErrorInfo(TXErrorCode err_code, TXErrorLevel err_level, std::string err_message)
        : code(err_code)
        , level(err_level) 
        , message(std::move(err_message))
        , line(0)
        , timestamp(std::chrono::system_clock::now()) {}
};

/**
 * @brief 错误处理器类型定义
 */
using TXErrorHandler = std::function<void(const TXErrorInfo&)>;

/**
 * @brief 统一错误管理类
 */
class TXError {
public:
    /**
     * @brief 默认构造函数（成功状态）
     */
    TXError() : error_info_() {}
    
    /**
     * @brief 构造错误对象
     * @param code 错误代码
     * @param message 错误消息
     * @param level 错误级别
     */
    TXError(TXErrorCode code, std::string message = "", TXErrorLevel level = TXErrorLevel::Error)
        : error_info_(code, level, std::move(message)) {}
    
    /**
     * @brief 构造错误对象（带详细信息）
     * @param code 错误代码
     * @param message 错误消息
     * @param level 错误级别
     * @param context 上下文信息
     * @param function 函数名
     * @param file 文件名
     * @param line 行号
     */
    TXError(TXErrorCode code, std::string message, TXErrorLevel level,
            std::string context, std::string function, std::string file, int line)
        : error_info_(code, level, std::move(message)) {
        error_info_.context = std::move(context);
        error_info_.function = std::move(function);
        error_info_.file = std::move(file);
        error_info_.line = line;
    }
    
    // ==================== 状态查询 ====================
    
    /**
     * @brief 检查是否成功（无错误）
     * @return 成功返回true，否则返回false
     */
    bool isSuccess() const { return error_info_.code == TXErrorCode::Success; }
    
    /**
     * @brief 检查是否有错误
     * @return 有错误返回true，否则返回false
     */
    bool hasError() const { return error_info_.code != TXErrorCode::Success; }
    
    /**
     * @brief 获取错误代码
     * @return 错误代码
     */
    TXErrorCode getCode() const { return error_info_.code; }
    
    /**
     * @brief 获取错误级别
     * @return 错误级别
     */
    TXErrorLevel getLevel() const { return error_info_.level; }
    
    /**
     * @brief 获取错误消息
     * @return 错误消息
     */
    const std::string& getMessage() const { return error_info_.message; }
    
    /**
     * @brief 获取错误上下文
     * @return 错误上下文
     */
    const std::string& getContext() const { return error_info_.context; }
    
    /**
     * @brief 获取完整错误信息
     * @return 错误信息结构
     */
    const TXErrorInfo& getInfo() const { return error_info_; }
    
    // ==================== 便捷操作符 ====================
    
    /**
     * @brief 布尔转换操作符（成功时返回true）
     * @return 成功返回true，失败返回false
     */
    explicit operator bool() const { return isSuccess(); }
    
    /**
     * @brief 否定操作符（有错误时返回true）
     * @return 有错误返回true，成功返回false  
     */
    bool operator!() const { return hasError(); }
    
    // ==================== 字符串转换 ====================
    
    /**
     * @brief 转换为字符串
     * @return 错误描述字符串
     */
    std::string toString() const;
    
    /**
     * @brief 转换为详细字符串
     * @return 详细的错误描述字符串
     */
    std::string toDetailString() const;
    
    // ==================== 静态工具方法 ====================
    
    /**
     * @brief 创建成功状态
     * @return 成功的错误对象
     */
    static TXError success() { return TXError(); }
    
    /**
     * @brief 创建错误状态
     * @param code 错误代码
     * @param message 错误消息
     * @param level 错误级别
     * @return 错误对象
     */
    static TXError create(TXErrorCode code, const std::string& message = "", 
                         TXErrorLevel level = TXErrorLevel::Error) {
        return TXError(code, message, level);
    }
    
    /**
     * @brief 获取错误代码名称
     * @param code 错误代码
     * @return 错误代码名称
     */
    static std::string getCodeName(TXErrorCode code);
    
    /**
     * @brief 获取错误级别名称
     * @param level 错误级别
     * @return 错误级别名称
     */
    static std::string getLevelName(TXErrorLevel level);
    
    /**
     * @brief 获取默认错误消息
     * @param code 错误代码
     * @return 默认错误消息
     */
    static std::string getDefaultMessage(TXErrorCode code);

private:
    TXErrorInfo error_info_;
};

/**
 * @brief 全局错误管理器
 */
class TXErrorManager {
public:
    /**
     * @brief 获取单例实例
     * @return 错误管理器实例
     */
    static TXErrorManager& getInstance();
    
    /**
     * @brief 设置错误处理器
     * @param handler 错误处理器函数
     */
    void setErrorHandler(TXErrorHandler handler);
    
    /**
     * @brief 报告错误
     * @param error 错误信息
     */
    void reportError(const TXError& error);
    
    /**
     * @brief 获取最后一个错误
     * @return 最后一个错误
     */
    TXError getLastError() const;
    
    /**
     * @brief 清除最后一个错误
     */
    void clearLastError();
    
    /**
     * @brief 获取错误历史记录
     * @return 错误历史记录
     */
    std::vector<TXErrorInfo> getErrorHistory() const;
    
    /**
     * @brief 设置最大历史记录数量
     * @param max_count 最大数量
     */
    void setMaxHistoryCount(size_t max_count);

private:
    TXErrorManager() = default;
    ~TXErrorManager() = default;
    TXErrorManager(const TXErrorManager&) = delete;
    TXErrorManager& operator=(const TXErrorManager&) = delete;
    
    mutable std::mutex mutex_;
    TXErrorHandler error_handler_;
    TXError last_error_;
    std::vector<TXErrorInfo> error_history_;
    size_t max_history_count_ = 100;
};

// ==================== 便捷宏定义 ====================

#define TX_ERROR(code, message) \
    TXError(code, message, TXErrorLevel::Error, "", __FUNCTION__, __FILE__, __LINE__)

#define TX_WARNING(code, message) \
    TXError(code, message, TXErrorLevel::Warning, "", __FUNCTION__, __FILE__, __LINE__)

#define TX_FATAL(code, message) \
    TXError(code, message, TXErrorLevel::Fatal, "", __FUNCTION__, __FILE__, __LINE__)

#define TX_REPORT_ERROR(error) \
    TXErrorManager::getInstance().reportError(error)

} // namespace TinaXlsx 