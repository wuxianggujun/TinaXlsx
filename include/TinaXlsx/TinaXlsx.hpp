#pragma once

/**
 * @file TinaXlsx.hpp
 * @brief TinaXlsx库的主头文件
 * 
 * 这个文件包含了TinaXlsx库的所有公开接口，
 * 用户只需要包含这一个头文件即可使用库的全部功能。
 */

#include "TXWorkbook.hpp"
#include "TXSheet.hpp"
#include "TXCell.hpp"
#include "TXZipHandler.hpp"
#include "TXXmlHandler.hpp"

/**
 * @namespace TinaXlsx
 * @brief TinaXlsx库的命名空间
 * 
 * 所有TinaXlsx库的类和函数都位于这个命名空间中。
 */
namespace TinaXlsx {

/**
 * @brief 库版本信息
 */
struct Version {
    static constexpr int MAJOR = 1;     ///< 主版本号
    static constexpr int MINOR = 0;     ///< 次版本号
    static constexpr int PATCH = 0;     ///< 补丁版本号
    
    /**
     * @brief 获取版本字符串
     * @return 版本字符串，格式为"major.minor.patch"
     */
    static std::string getString() {
        return std::to_string(MAJOR) + "." + 
               std::to_string(MINOR) + "." + 
               std::to_string(PATCH);
    }
};

/**
 * @brief 初始化库
 * 
 * 在使用库的功能之前调用此函数进行初始化。
 * 这个函数是线程安全的，可以多次调用。
 * 
 * @return 成功返回true，失败返回false
 */
bool initialize();

/**
 * @brief 清理库资源
 * 
 * 在程序结束前调用此函数清理库使用的资源。
 * 这个函数是线程安全的，可以多次调用。
 */
void cleanup();

/**
 * @brief 获取库的版本信息
 * @return 版本字符串
 */
std::string getVersion();

/**
 * @brief 获取构建信息
 * @return 构建信息字符串
 */
std::string getBuildInfo();

} // namespace TinaXlsx 