/**
 * @file Workbook.hpp
 * @brief Excel工作簿类 - 统一的读写接口
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include "Reader.hpp"
#include "Writer.hpp"
#include <memory>
#include <string>

namespace TinaXlsx {

/**
 * @brief Excel工作簿类
 * 
 * 提供统一的Excel文件读写接口，可以根据需要自动选择读取或写入模式
 */
class Workbook {
public:
    /**
     * @brief 工作簿模式枚举
     */
    enum class Mode {
        Read,       // 只读模式
        Write,      // 只写模式
        ReadWrite   // 读写模式（暂不支持）
    };

private:
    std::string filePath_;
    Mode mode_;
    std::unique_ptr<Reader> reader_;
    std::unique_ptr<Writer> writer_;

public:
    /**
     * @brief 构造函数
     * @param filePath Excel文件路径
     * @param mode 工作簿模式
     */
    explicit Workbook(const std::string& filePath, Mode mode = Mode::Write);
    
    /**
     * @brief 移动构造函数
     */
    Workbook(Workbook&& other) noexcept;
    
    /**
     * @brief 移动赋值操作符
     */
    Workbook& operator=(Workbook&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~Workbook();
    
    // 禁用拷贝构造和拷贝赋值
    Workbook(const Workbook&) = delete;
    Workbook& operator=(const Workbook&) = delete;
    
    /**
     * @brief 获取文件路径
     * @return std::string 文件路径
     */
    [[nodiscard]] const std::string& getFilePath() const { return filePath_; }
    
    /**
     * @brief 获取工作簿模式
     * @return Mode 工作簿模式
     */
    [[nodiscard]] Mode getMode() const { return mode_; }
    
    /**
     * @brief 切换到读取模式
     * @return Reader& 读取器引用
     */
    [[nodiscard]] Reader& getReader();
    
    /**
     * @brief 切换到写入模式
     * @return Writer& 写入器引用
     */
    [[nodiscard]] Writer& getWriter();
    
    /**
     * @brief 检查是否支持读取
     * @return bool 是否支持读取
     */
    [[nodiscard]] bool canRead() const;
    
    /**
     * @brief 检查是否支持写入
     * @return bool 是否支持写入
     */
    [[nodiscard]] bool canWrite() const;
    
    /**
     * @brief 保存工作簿（仅写入模式）
     * @return bool 是否成功保存
     */
    bool save();
    
    /**
     * @brief 关闭工作簿
     * @return bool 是否成功关闭
     */
    bool close();
    
    /**
     * @brief 检查工作簿是否已关闭
     * @return bool 是否已关闭
     */
    [[nodiscard]] bool isClosed() const;
    
    /**
     * @brief 从现有文件创建只读工作簿
     * @param filePath 文件路径
     * @return std::unique_ptr<Workbook> 工作簿智能指针
     */
    [[nodiscard]] static std::unique_ptr<Workbook> openForRead(const std::string& filePath);
    
    /**
     * @brief 创建新的只写工作簿
     * @param filePath 文件路径
     * @return std::unique_ptr<Workbook> 工作簿智能指针
     */
    [[nodiscard]] static std::unique_ptr<Workbook> createForWrite(const std::string& filePath);
};

} // namespace TinaXlsx 