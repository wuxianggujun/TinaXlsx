#include "TinaXlsx/TXError.hpp"
#include <sstream>
#include <iomanip>

namespace TinaXlsx
{
    // ==================== TXError 实现 ====================

    std::string TXError::toString() const {
        if (isSuccess()) {
            return "Success";
        }
    
        std::ostringstream oss;
        oss << "[" << static_cast<int>(error_info_.code) << "] ";
    
        if (!error_info_.message.empty()) {
            oss << error_info_.message;
        } else {
            oss << getDefaultMessage(error_info_.code);
        }
    
        return oss.str();
    }

    std::string TXError::toDetailString() const {
        if (isSuccess()) {
            return "Success";
        }
    
        std::ostringstream oss;
        oss << "Error Details:\n";
        oss << "  Code: " << static_cast<int>(error_info_.code) << " (" << getCodeName(error_info_.code) << ")\n";
        oss << "  Level: " << getLevelName(error_info_.level) << "\n";
        oss << "  Message: " << (error_info_.message.empty() ? getDefaultMessage(error_info_.code) : error_info_.message) << "\n";
    
        if (!error_info_.context.empty()) {
            oss << "  Context: " << error_info_.context << "\n";
        }
    
        if (!error_info_.function.empty()) {
            oss << "  Function: " << error_info_.function << "\n";
        }
    
        if (!error_info_.file.empty()) {
            oss << "  File: " << error_info_.file;
            if (error_info_.line > 0) {
                oss << ":" << error_info_.line;
            }
            oss << "\n";
        }
    
        return oss.str();
    }
}
