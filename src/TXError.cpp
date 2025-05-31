#include "TinaXlsx/TXError.hpp"
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <map>

namespace TinaXlsx
{
    // ==================== TXError 实现 ====================

    std::string TXError::toString() const
    {
        if (isSuccess())
        {
            return "Success";
        }

        std::ostringstream oss;
        oss << "[" << static_cast<int>(error_info_.code) << "] ";

        if (!error_info_.message.empty())
        {
            oss << error_info_.message;
        }
        else
        {
            oss << getDefaultMessage(error_info_.code);
        }
        if (!error_info_.context.empty())
        {
            oss << " (Context: " << error_info_.context << ")";
        }
        return oss.str();
    }

    std::string TXError::toDetailString() const
    {
        if (isSuccess())
        {
            return "Success (No error details)";
        }

        std::ostringstream oss;
        const TXError* current_error = this;
        int indent_level = 0;
        bool first_error = true;

        while (current_error != nullptr)
        {
            std::string indent(indent_level * 2, ' '); // 每层缩进2个空格

            if (!first_error)
            {
                oss << "\n" << indent << "Caused by:\n";
            }

            const TXErrorInfo& info = current_error->getInfo(); // 使用 getInfo()

            oss << indent << "Error Code: " << static_cast<int>(info.code)
                << " (" << getCodeName(info.code) << ")\n";
            oss << indent << "Level: " << getLevelName(info.level) << "\n";
            oss << indent << "Message: "
                << (info.message.empty() ? getDefaultMessage(info.code) : info.message) << "\n";

            if (!info.context.empty())
            {
                oss << indent << "Context: " << info.context << "\n";
            }
            if (!info.function.empty())
            {
                oss << indent << "Function: " << info.function << "\n";
            }
            if (!info.file.empty())
            {
                oss << indent << "Location: " << info.file << ":" << info.line << "\n";
            }

            // 时间戳的格式化可以根据需要添加
            auto time_point = info.timestamp;
            std::time_t tt = std::chrono::system_clock::to_time_t(time_point);
            std::tm tm_info;
#ifdef _WIN32
            localtime_s(&tm_info, &tt);
#else
            localtime_r(&tt, &tm_info);
#endif
            oss << indent << "Timestamp: " << std::put_time(&tm_info, "%Y-%m-%d %H:%M:%S") << "\n";

            current_error = current_error->getCause();
            indent_level++;
            first_error = false;
        }
        return oss.str();
    }

    std::string TXError::getCodeName(TXErrorCode code)
    {
        // 使用静态映射表来存储错误码和它们的名称
        // 这样可以避免每次调用函数时都构建 switch 语句或大量 if-else
        static const std::map<TXErrorCode, std::string> code_names = {
            {TXErrorCode::Success, "Success"},

            // 通用错误
            {TXErrorCode::Unknown, "Unknown"},
            {TXErrorCode::InvalidArgument, "InvalidArgument"},
            {TXErrorCode::NullPointer, "NullPointer"},
            {TXErrorCode::OutOfRange, "OutOfRange"},
            {TXErrorCode::InvalidOperation, "InvalidOperation"},
            {TXErrorCode::MemoryAllocation, "MemoryAllocation"},
            {TXErrorCode::OperationFailed, "OperationFailed"},

            // 文件I/O错误
            {TXErrorCode::FileNotFound, "FileNotFound"},
            {TXErrorCode::FileOpenFailed, "FileOpenFailed"},
            {TXErrorCode::FileWriteFailed, "FileWriteFailed"},
            {TXErrorCode::FileReadFailed, "FileReadFailed"},
            {TXErrorCode::InvalidFileFormat, "InvalidFileFormat"},

            // 坐标相关错误
            {TXErrorCode::InvalidCoordinate, "InvalidCoordinate"},
            {TXErrorCode::InvalidRow, "InvalidRow"},
            {TXErrorCode::InvalidColumn, "InvalidColumn"},
            {TXErrorCode::InvalidRange, "InvalidRange"},

            // 单元格相关错误
            {TXErrorCode::CellNotFound, "CellNotFound"},
            {TXErrorCode::InvalidCellValue, "InvalidCellValue"},
            {TXErrorCode::CellTypeConversion, "CellTypeConversion"},

            // 工作表相关错误
            {TXErrorCode::SheetNotFound, "SheetNotFound"},
            {TXErrorCode::SheetNameExists, "SheetNameExists"},
            {TXErrorCode::InvalidSheetName, "InvalidSheetName"},

            // 样式相关错误
            {TXErrorCode::InvalidStyle, "InvalidStyle"},
            {TXErrorCode::StyleNotFound, "StyleNotFound"},
            {TXErrorCode::StyleRegistrationFailed, "StyleRegistrationFailed"},

            // XML处理错误
            {TXErrorCode::XmlParseError, "XmlParseError"},
            {TXErrorCode::XmlGenerationError, "XmlGenerationError"},
            {TXErrorCode::InvalidXmlStructure, "InvalidXmlStructure"},

            // ZIP处理错误
            {TXErrorCode::ZipCreateFailed, "ZipCreateFailed"},
            {TXErrorCode::ZipOpenFailed, "ZipOpenFailed"},
            {TXErrorCode::ZipExtractFailed, "ZipExtractFailed"},
            {TXErrorCode::ZipCompressionFailed, "ZipCompressionFailed"},
            {TXErrorCode::ZipReadEntryFailed, "ZipReadEntryFailed"},
            {TXErrorCode::ZipWriteEntryFailed, "ZipWriteEntryFailed"},
            {TXErrorCode::ZipEntryNotFound, "ZipEntryNotFound"},
            {TXErrorCode::ZipInvalidState, "ZipInvalidState"}
        };

        auto it = code_names.find(code);
        if (it != code_names.end())
        {
            return it->second;
        }
        return "UndefinedErrorCode"; // 如果错误码未在映射中找到
    }

    std::string TXError::getLevelName(TXErrorLevel level)
    {
        switch (level)
        {
        case TXErrorLevel::Info: return "Info";
        case TXErrorLevel::Warning: return "Warning";
        case TXErrorLevel::Error: return "Error";
        case TXErrorLevel::Fatal: return "Fatal";
        default: return "UnknownLevel";
        }
    }

    std::string TXError::getDefaultMessage(TXErrorCode code)
    {
        // 使用静态映射表来存储错误码和它们的默认英文消息
        static const std::map<TXErrorCode, std::string> default_messages = {
            {TXErrorCode::Success, "Operation completed successfully."},

            // 通用错误 (General Errors)
            {TXErrorCode::Unknown, "An unknown error occurred."},
            {TXErrorCode::InvalidArgument, "An invalid argument was provided."},
            {TXErrorCode::NullPointer, "A null pointer was encountered unexpectedly."},
            {TXErrorCode::OutOfRange, "Access was attempted outside of a valid range."},
            {TXErrorCode::InvalidOperation, "The operation is invalid in the current state or context."},
            {TXErrorCode::MemoryAllocation, "Failed to allocate necessary memory."},
            {TXErrorCode::OperationFailed, "The operation did not complete successfully."},

            // 文件I/O错误 (File I/O Errors)
            {TXErrorCode::FileNotFound, "The specified file was not found."},
            {TXErrorCode::FileOpenFailed, "Failed to open the specified file."},
            {TXErrorCode::FileWriteFailed, "Failed to write to the file."},
            {TXErrorCode::FileReadFailed, "Failed to read from the file."},
            {TXErrorCode::InvalidFileFormat, "The file format is invalid or unsupported."},

            // 坐标相关错误 (Coordinate Errors)
            {TXErrorCode::InvalidCoordinate, "The provided coordinate is invalid."},
            {TXErrorCode::InvalidRow, "The row number is invalid or out of range."},
            {TXErrorCode::InvalidColumn, "The column identifier (number or name) is invalid or out of range."},
            {TXErrorCode::InvalidRange, "The specified cell range is invalid."},

            // 单元格相关错误 (Cell Errors)
            {TXErrorCode::CellNotFound, "No cell was found at the specified location."},
            {TXErrorCode::InvalidCellValue, "The cell value is invalid or does not match the expected type."},
            {TXErrorCode::CellTypeConversion, "Failed to convert the cell's data type."},

            // 工作表相关错误 (Worksheet Errors)
            {TXErrorCode::SheetNotFound, "The worksheet with the specified name or index does not exist."},
            {TXErrorCode::SheetNameExists, "A worksheet with the same name already exists."},
            {
                TXErrorCode::InvalidSheetName,
                "The worksheet name is invalid (e.g., contains illegal characters or is too long)."
            },

            // 样式相关错误 (Style Errors)
            {TXErrorCode::InvalidStyle, "The provided style information is invalid."},
            {TXErrorCode::StyleNotFound, "The requested style was not found."},
            {TXErrorCode::StyleRegistrationFailed, "Failed to register the style."},

            // XML处理错误 (XML Processing Errors)
            {TXErrorCode::XmlParseError, "An error occurred while parsing XML data."},
            {TXErrorCode::XmlGenerationError, "An error occurred while generating XML data."},
            {TXErrorCode::InvalidXmlStructure, "The XML structure does not conform to the expected schema."},

            // ZIP处理错误 (ZIP Processing Errors)
            {TXErrorCode::ZipCreateFailed, "Failed to create the ZIP archive file."},
            {TXErrorCode::ZipOpenFailed, "Failed to open the ZIP archive file."},
            {TXErrorCode::ZipExtractFailed, "Failed to extract data from the ZIP archive file."},
            {TXErrorCode::ZipCompressionFailed, "ZIP compression or decompression operation failed."},
            {TXErrorCode::ZipReadEntryFailed, "Failed to read an entry from the ZIP archive file."},
            {TXErrorCode::ZipWriteEntryFailed, "Failed to write an entry to the ZIP archive file."},
            {TXErrorCode::ZipEntryNotFound, "The specified entry was not found in the ZIP archive file."},
            {TXErrorCode::ZipInvalidState, "The ZIP archive file is in an invalid state."},
        };

        const auto it = default_messages.find(code);
        if (it != default_messages.end())
        {
            return it->second;
        }
        return "No default message available for this error code."; // 通用回退消息
    }
}
