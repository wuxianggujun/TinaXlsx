//
// @file TXCoordUtils.hpp
// @brief 统一的坐标转换工具 - 消除重复实现，提供高性能坐标转换
//

#pragma once

#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXResult.hpp"
#include <string>
#include <string_view>
#include <vector>
#include <utility>

namespace TinaXlsx {

/**
 * @brief 🚀 统一的Excel坐标转换工具
 * 
 * 整合并替代以下重复实现：
 * - TXExcelCoordParser (TXSheetAPI)
 * - TXExcelCoordinates (TXXMLTemplates)
 * - TXCoordinate构造函数中的解析逻辑
 * 
 * 设计目标：
 * - 高性能：避免不必要的对象构造
 * - 类型安全：完善的错误处理
 * - 功能完整：支持单个坐标、范围、批量操作
 */
class TXCoordUtils {
public:
    // ==================== 单个坐标转换 ====================
    
    /**
     * @brief 🚀 解析Excel格式坐标为TXCoordinate对象
     * @param excel_coord Excel格式坐标 (如 "A1", "B2", "AA100")
     * @return TXCoordinate对象或错误信息
     */
    static TXResult<TXCoordinate> parseCoord(std::string_view excel_coord);
    
    /**
     * @brief 🚀 解析Excel格式坐标为行列索引 (高性能版本)
     * @param excel_coord Excel格式坐标
     * @return {row, col} 0-based索引对，解析失败返回{UINT32_MAX, UINT32_MAX}
     */
    static std::pair<uint32_t, uint32_t> parseCoordFast(std::string_view excel_coord) noexcept;
    
    /**
     * @brief 🚀 TXCoordinate转换为Excel格式字符串
     * @param coord 坐标对象
     * @return Excel格式字符串 (如 "A1", "B2")
     */
    static std::string coordToExcel(const TXCoordinate& coord);
    
    /**
     * @brief 🚀 行列索引转换为Excel格式字符串 (高性能版本)
     * @param row 行索引 (0-based)
     * @param col 列索引 (0-based)
     * @return Excel格式字符串
     */
    static std::string coordToExcel(uint32_t row, uint32_t col);

    // ==================== 范围转换 ====================
    
    /**
     * @brief 🚀 解析Excel格式范围
     * @param excel_range Excel格式范围 (如 "A1:B2", "C3:D10")
     * @return 起始和结束坐标对
     */
    static TXResult<std::pair<TXCoordinate, TXCoordinate>> parseRange(std::string_view excel_range);
    
    /**
     * @brief 🚀 坐标对转换为Excel格式范围字符串
     * @param start 起始坐标
     * @param end 结束坐标
     * @return Excel格式范围字符串 (如 "A1:B2")
     */
    static std::string rangeToExcel(const TXCoordinate& start, const TXCoordinate& end);

    // ==================== 批量转换 (高性能) ====================
    
    /**
     * @brief 🚀 批量转换坐标为Excel格式字符串
     * @param coords 坐标数组
     * @param count 坐标数量
     * @param output 输出字符串向量
     */
    static void coordsBatchToExcel(const TXCoordinate* coords, size_t count, 
                                   std::vector<std::string>& output);
    
    /**
     * @brief 🚀 批量转换打包坐标为Excel格式字符串
     * @param packed_coords 打包坐标数组 (高16位=行，低16位=列)
     * @param count 坐标数量
     * @param output 输出字符串向量
     */
    static void packedCoordsBatchToExcel(const uint32_t* packed_coords, size_t count,
                                         std::vector<std::string>& output);

    // ==================== 列转换工具 ====================
    
    /**
     * @brief 🚀 列字母转换为列索引
     * @param col_letters 列字母 (如 "A", "B", "AA")
     * @return 列索引 (0-based)，失败返回UINT32_MAX
     */
    static uint32_t columnLettersToIndex(std::string_view col_letters) noexcept;
    
    /**
     * @brief 🚀 列索引转换为列字母
     * @param col_index 列索引 (0-based)
     * @return 列字母字符串
     */
    static std::string columnIndexToLetters(uint32_t col_index);

    // ==================== 验证工具 ====================
    
    /**
     * @brief 🚀 验证Excel坐标格式是否有效
     * @param excel_coord Excel格式坐标
     * @return 有效返回true
     */
    static bool isValidExcelCoord(std::string_view excel_coord) noexcept;
    
    /**
     * @brief 🚀 验证Excel范围格式是否有效
     * @param excel_range Excel格式范围
     * @return 有效返回true
     */
    static bool isValidExcelRange(std::string_view excel_range) noexcept;

    // ==================== 常量定义 ====================
    
    /// 无效坐标标记
    static constexpr uint32_t INVALID_INDEX = UINT32_MAX;
    
    /// Excel最大行数 (1-based)
    static constexpr uint32_t MAX_EXCEL_ROWS = 1048576;
    
    /// Excel最大列数 (1-based)  
    static constexpr uint32_t MAX_EXCEL_COLS = 16384;

private:
    // ==================== 内部优化方法 ====================
    
    /**
     * @brief 内部高性能列字母解析
     */
    static uint32_t parseColumnLettersInternal(std::string_view letters) noexcept;
    
    /**
     * @brief 内部高性能行号解析
     */
    static uint32_t parseRowNumberInternal(std::string_view numbers) noexcept;
    
    /**
     * @brief 内部高性能列索引转字母
     */
    static void columnIndexToLettersInternal(uint32_t col_index, std::string& result);
};

// 🚀 直接使用 TXCoordUtils，无需别名和宏

} // namespace TinaXlsx
