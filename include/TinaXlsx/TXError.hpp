#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <memory>
#include <functional>
#include <mutex>
#include <chrono>

namespace TinaXlsx
{
    class TXError;

    /**
     * @brief 错误级别枚举
     */
    enum class TXErrorLevel : int
    {
        Info = 0, // 信息
        Warning = 1, // 警告
        Error = 2, // 错误
        Fatal = 3 // 致命错误
    };

    /**
     * @brief 错误代码枚举
     */
    enum class TXErrorCode : int
    {
        Success = 0, // 成功，无错误

        // 通用错误 (1-99)
        Unknown = 1, // 未知错误
        InvalidArgument = 2, // 无效参数
        InvalidParameter = 2, // 无效参数 (别名)
        NullPointer = 3, // 空指针
        OutOfRange = 4, // 超出范围
        InvalidOperation = 5, // 无效操作
        MemoryAllocation = 6, // 内存分配失败
        MemoryError = 6, // 内存错误 (别名)
        OperationFailed = 7, // 操作失败 (通用)
        InvalidData = 8, // 无效数据
        SerializationError = 9, // 序列化错误

        // 文件I/O错误 (100-199)
        FileNotFound = 100, // 文件未找到
        FileOpenFailed = 101, // 文件打开失败
        FileWriteFailed = 102, // 文件写入失败
        FileReadFailed = 103, // 文件读取失败
        InvalidFileFormat = 104, // 无效文件格式
        UnsupportedFormat = 105, // 不支持的文件格式

        // 坐标相关错误 (200-299)
        InvalidCoordinate = 200, // 无效坐标
        InvalidRow = 201, // 无效行号
        InvalidColumn = 202, // 无效列号
        InvalidRange = 203, // 无效范围

        // 单元格相关错误 (300-399)
        CellNotFound = 300, // 单元格不存在
        InvalidCellValue = 301, // 无效单元格值
        CellTypeConversion = 302, // 单元格类型转换失败

        // 工作表相关错误 (400-499)
        SheetNotFound = 400, // 工作表不存在
        SheetNameExists = 401, // 工作表名称已存在
        InvalidSheetName = 402, // 无效工作表名称

        // 样式相关错误 (500-599)
        InvalidStyle = 500, // 无效样式
        StyleNotFound = 501, // 样式不存在
        StyleRegistrationFailed = 502, // 样式注册失败

        // XML处理错误 (600-699)
        XmlParseError = 600, // XML解析错误
        XmlGenerationError = 601, // XML生成错误
        InvalidXmlStructure = 602, // 无效XML结构
        XmlInvalidState = 603, // XML对象无效状态
        XmlXpathError = 604, // XPath错误
        XmlNoRoot = 605, // 没有根节点
        XmlNodeNotFound = 606, // 节点未找到
        XmlAttributeNotFound = 607, // 属性未找到
        XmlCreateError = 608, // XML创建错误

        // ZIP处理错误 (700-799)
        ZipCreateFailed = 700, // ZIP创建失败
        ZipOpenFailed = 701, // ZIP打开失败
        ZipExtractFailed = 702, // ZIP解压失败
        ZipCompressionFailed = 703, // ZIP压缩失败
        ZipReadEntryFailed = 704, // 读取ZIP条目数据失败
        ZipWriteEntryFailed = 705, // 写入ZIP条目数据失败
        ZipEntryNotFound = 706, // 在ZIP中未找到指定条目 (不同于文件系统中的 FileNotFound)
        ZipInvalidState = 707, // ZIP对象处于无效状态执行操作 (例如未打开时读取)
        ZipWriteError = 708, // ZIP写入错误
    };

    /**
     * @brief 错误信息结构
     */
    struct TXErrorInfo
    {
        TXErrorCode code;
        TXErrorLevel level;
        std::string message;
        std::string context; // 错误发生的上下文
        std::string function; // 发生错误的函数名
        std::string file; // 发生错误的文件名
        int line; // 发生错误的行号
        std::chrono::system_clock::time_point timestamp;
        std::unique_ptr<TXError> cause; // 导致此错误的错误

        TXErrorInfo()
            : code(TXErrorCode::Success)
              , level(TXErrorLevel::Info)
              , line(0)
              , timestamp(std::chrono::system_clock::now())
              , cause(nullptr)
        {
        }

        TXErrorInfo(TXErrorCode err_code, TXErrorLevel err_level, std::string err_message)
            : code(err_code)
              , level(err_level)
              , message(std::move(err_message))
              , line(0)
              , timestamp(std::chrono::system_clock::now())
              , cause(nullptr)
        {
        }

        // 为了支持unique_ptr的移动，需要移动构造和移动赋值
        TXErrorInfo(TXErrorInfo&& other) noexcept
            : code(other.code), level(other.level), message(std::move(other.message)),
              context(std::move(other.context)), function(std::move(other.function)),
              file(std::move(other.file)), line(other.line), timestamp(other.timestamp),
              cause(std::move(other.cause))
        {
        }

        TXErrorInfo& operator=(TXErrorInfo&& other) noexcept
        {
            if (this != &other)
            {
                code = other.code;
                level = other.level;
                message = std::move(other.message);
                context = std::move(other.context);
                function = std::move(other.function);
                file = std::move(other.file);
                line = other.line;
                timestamp = other.timestamp;
                cause = std::move(other.cause);
            }
            return *this;
        }

        // 禁止拷贝，因为 unique_ptr 不可拷贝
        TXErrorInfo(const TXErrorInfo&) = delete;
        TXErrorInfo& operator=(const TXErrorInfo&) = delete;
    };


    /**
     * @brief 统一错误管理类
     */
    class TXError
    {
    public:
        /**
         * @brief 默认构造函数（成功状态）
         */
        TXError() : error_info_()
        {
        }

        /**
         * @brief 构造错误对象
         * @param code 错误代码
         * @param message 错误消息
         * @param level 错误级别
         */
        explicit TXError(TXErrorCode code, std::string message = "", TXErrorLevel level = TXErrorLevel::Error)
            : error_info_(code, level, std::move(message))
        {
        }

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
            : error_info_(code, level, std::move(message))
        {
            error_info_.context = std::move(context);
            error_info_.function = std::move(function);
            error_info_.file = std::move(file);
            error_info_.line = line;
        }

        // 支持移动构造和赋值
        TXError(TXError&& other) noexcept = default; // 使用默认的移动构造函数
        TXError& operator=(TXError&& other) noexcept = default; // 使用默认的移动赋值操作符


        // 由于 TXErrorInfo 包含 unique_ptr，拷贝构造和赋值需要特殊处理或禁用
        // 如果需要拷贝，则要深拷贝 cause
        TXError(const TXError& other)
            : error_info_(other.error_info_.code, other.error_info_.level, other.error_info_.message)
        {
            error_info_.context = other.error_info_.context;
            error_info_.function = other.error_info_.function;
            error_info_.file = other.error_info_.file;
            error_info_.line = other.error_info_.line;
            error_info_.timestamp = other.error_info_.timestamp;
            if (other.error_info_.cause)
            {
                error_info_.cause = std::make_unique<TXError>(*other.error_info_.cause); // 深拷贝 cause
            }
        }

        TXError& operator=(const TXError& other)
        {
            if (this != &other)
            {
                error_info_.code = other.error_info_.code;
                error_info_.level = other.error_info_.level;
                error_info_.message = other.error_info_.message;
                error_info_.context = other.error_info_.context;
                error_info_.function = other.error_info_.function;
                error_info_.file = other.error_info_.file;
                error_info_.line = other.error_info_.line;
                error_info_.timestamp = other.error_info_.timestamp;
                if (other.error_info_.cause)
                {
                    error_info_.cause = std::make_unique<TXError>(*other.error_info_.cause); // 深拷贝 cause
                }
                else
                {
                    error_info_.cause.reset();
                }
            }
            return *this;
        }


        // ==================== 状态查询 ====================

        /**
         * @brief 检查是否成功（无错误）
         * @return 成功返回true，否则返回false
         */
        [[nodiscard]] bool isSuccess() const { return error_info_.code == TXErrorCode::Success; }

        /**
         * @brief 检查是否有错误
         * @return 有错误返回true，否则返回false
         */
        [[nodiscard]] bool hasError() const { return error_info_.code != TXErrorCode::Success; }

        /**
         * @brief 获取错误代码
         * @return 错误代码
         */
        [[nodiscard]] TXErrorCode getCode() const { return error_info_.code; }

        /**
         * @brief 获取错误级别
         * @return 错误级别
         */
        [[nodiscard]] TXErrorLevel getLevel() const { return error_info_.level; }

        /**
         * @brief 获取错误消息
         * @return 错误消息
         */
        [[nodiscard]] const std::string& getMessage() const { return error_info_.message; }

        /**
         * @brief 获取错误上下文
         * @return 错误上下文
         */
        [[nodiscard]] const std::string& getContext() const { return error_info_.context; }

        /**
         * @brief 获取完整错误信息
         * @return 错误信息结构
         */
        [[nodiscard]] const TXErrorInfo& getInfo() const { return error_info_; }


        // ==================== 上下文和错误链操作 ====================

        void appendContext(const std::string& message)
        {
            if (error_info_.context.empty())
            {
                error_info_.context = message;
            }
            else
            {
                error_info_.context = message + " (原上下文: " + error_info_.context + ")";
            }
        }

        void setCause(TXError&& cause_error)
        {
            if (cause_error.hasError())
            {
                error_info_.cause = std::make_unique<TXError>(std::move(cause_error));
            }
        }

        void setCause(const TXError& cause_error)
        {
            if (cause_error.hasError())
            {
                error_info_.cause = std::make_unique<TXError>(cause_error);
            }
        }

        [[nodiscard]] const TXError* getCause() const { return error_info_.cause.get(); }


        // ==================== 便捷操作符 ====================

        /**
         * @brief 布尔转换操作符（成功时返回true）
         * @return 成功返回true，失败返回false
         */
        [[nodiscard]] explicit operator bool() const { return isSuccess(); }

        /**
         * @brief 否定操作符（有错误时返回true）
         * @return 有错误返回true，成功返回false  
         */
        [[nodiscard]] bool operator!() const { return hasError(); }

        // ==================== 字符串转换 ====================

        /**
         * @brief 转换为字符串
         * @return 错误描述字符串
         */
        [[nodiscard]] std::string toString() const;

        /**
         * @brief 转换为详细字符串
         * @return 详细的错误描述字符串
         */
        [[nodiscard]] std::string toDetailString() const;

        // ==================== 静态工具方法 ====================

        /**
         * @brief 创建成功状态
         * @return 成功的错误对象
         */
        [[nodiscard]] static TXError success() { return {}; }

        /**
         * @brief 创建错误状态
         * @param code 错误代码
         * @param message 错误消息
         * @param level 错误级别
         * @return 错误对象
         */
        [[nodiscard]] static TXError create(TXErrorCode code, const std::string& message = "",
                                            TXErrorLevel level = TXErrorLevel::Error)
        {
            return TXError(code, message, level);
        }

        /**
         * @brief 获取错误代码名称
         * @param code 错误代码
         * @return 错误代码名称
         */
        [[nodiscard]] static std::string getCodeName(TXErrorCode code);

        /**
         * @brief 获取错误级别名称
         * @param level 错误级别
         * @return 错误级别名称
         */
        [[nodiscard]] static std::string getLevelName(TXErrorLevel level);

        /**
         * @brief 获取默认错误消息
         * @param code 错误代码
         * @return 默认错误消息
         */
        [[nodiscard]] static std::string getDefaultMessage(TXErrorCode code);

    private:
        TXErrorInfo error_info_;
    };

    // ==================== 便捷宏定义 ====================
#define TX_ERROR_CREATE(code, message) \
    TinaXlsx::TXError(code, message, TinaXlsx::TXErrorLevel::Error, "", __FUNCTION__, __FILE__, __LINE__)

#define TX_WARNING_CREATE(code, message) \
    TinaXlsx::TXError(code, message, TinaXlsx::TXErrorLevel::Warning, "", __FUNCTION__, __FILE__, __LINE__)

#define TX_FATAL_CREATE(code, message) \
    TinaXlsx::TXError(code, message, TinaXlsx::TXErrorLevel::Fatal, "", __FUNCTION__, __FILE__, __LINE__)
} // namespace TinaXlsx 
