#pragma once

#include "TXTypes.hpp"
#include "TXFormula.hpp"
#include "TXNumberFormat.hpp"
#include <string>
#include <variant>
#include <memory>
#include <utility>

namespace TinaXlsx
{
    // 前向声明
    class TXSheet; // TXCell::evaluateFormula 需要 TXSheet

    /**
     * @brief Excel单元格类
     *
     * 单元格数据的抽象，提供读写单个单元格的能力，包括公式、格式化和合并功能。
     */
    class TXCell
    {
    public:
        // ==================== 类型别名 ====================
        using CellValue = cell_value_t; // cell_value_t 来自 TXTypes.hpp

        /**
         * @brief 单元格类型枚举
         */
        enum class CellType
        {
            Empty, ///< 空单元格
            String, ///< 字符串
            Number, ///< 数字
            Integer, ///< 整数
            Boolean, ///< 布尔值
            Formula, ///< 公式
            Error ///< 错误值
        };

    public:
        TXCell();
        explicit TXCell(const CellValue& value);
        ~TXCell();

        // 支持拷贝构造和赋值
        TXCell(const TXCell& other);
        TXCell& operator=(const TXCell& other);

        // 支持移动构造和赋值
        TXCell(TXCell&& other) noexcept;
        TXCell& operator=(TXCell&& other) noexcept;

        /**
         * @brief 获取单元格的原始值 (std::variant)
         * @return 单元格的原始值 (CellValue)
         */
        [[nodiscard]] const CellValue& getValue() const;

        /**
         * @brief 设置单元格的原始值 (std::variant)
         * @param value 要设置的单元格值 (CellValue)
         */
        void setValue(const CellValue& value);

        /**
         * @brief 获取单元格的当前类型
         * @return 单元格类型 (CellType)
         */
        [[nodiscard]] CellType getType() const;

        /**
         * @brief 检查单元格是否为空 (值为std::monostate且无公式)
         * @return 如果单元格为空则返回true，否则返回false
         */
        [[nodiscard]] bool isEmpty() const;

        /**
         * @brief 获取单元格的字符串值
         * @return 如果值为字符串则返回该字符串，否则尝试转换为字符串或返回空字符串。
         * 对于公式单元格，这通常返回公式的缓存结果（如果已计算）或值的字符串形式。
         */
        [[nodiscard]] std::string getStringValue() const;

        /**
         * @brief 获取单元格的数字值 (double)
         * @return 如果值为数字或可转换为数字，则返回该数字，否则返回0.0。
         */
        [[nodiscard]] double getNumberValue() const;

        /**
         * @brief 获取单元格的整数值 (int64_t)
         * @return 如果值为整数或可转换为整数，则返回该整数，否则返回0。
         */
        [[nodiscard]] int64_t getIntegerValue() const;

        /**
         * @brief 获取单元格的布尔值
         * @return 如果值为布尔型，则返回该布尔值，否则返回false。
         */
        [[nodiscard]] bool getBooleanValue() const;

        /**
         * @brief 将单元格值设置为字符串
         * @param value 字符串值
         */
        void setStringValue(const std::string& value);

        /**
         * @brief 将单元格值设置为数字 (double)
         * @param value 数字值
         */
        void setNumberValue(double value);

        /**
         * @brief 将单元格值设置为整数 (int64_t)
         * @param value 整数值
         */
        void setIntegerValue(int64_t value);

        /**
         * @brief 将单元格值设置为布尔型
         * @param value 布尔值
         */
        void setBooleanValue(bool value);

        // ==================== 公式功能 ====================

        /**
         * @brief 获取单元格的公式字符串 (不包含等号)
         * @return 如果单元格包含公式，则返回公式字符串；否则返回空字符串。
         */
        [[nodiscard]] std::string getFormula() const;

        /**
         * @brief 设置单元格的公式
         * @param formula_str 公式字符串 (不包含等号)。如果为空字符串，则清除公式。
         */
        void setFormula(const std::string& formula_str);

        /**
         * @brief 检查单元格是否包含公式
         * @return 如果单元格是一个公式单元格，则返回true，否则返回false。
         */
        [[nodiscard]] bool isFormula() const;

        /**
         * @brief 获取单元格的TXFormula对象 (只读)
         * @return 如果存在公式对象，则返回其常量指针；否则返回nullptr。
         */
        [[nodiscard]] const TXFormula* getFormulaObject() const;

        /**
         * @brief 设置单元格的TXFormula对象
         * @param formula_ptr 指向TXFormula对象的unique_ptr。所有权将转移。
         */
        void setFormulaObject(std::unique_ptr<TXFormula> formula_ptr);

        /**
         * @brief 计算单元格公式的值 (如果它是公式单元格)
         * @param sheet 当前单元格所在的工作表指针 (用于解析单元格引用)。
         * @param currentRow 当前单元格的行号。
         * @param currentCol 当前单元格的列号。
         * @return 计算得到的CellValue。如果不是公式或计算失败，可能返回monostate或错误字符串。
         * @note 此方法可能会修改单元格内部缓存的公式结果(如果适用)，但通常TXCell的value_用于存储。
         */
        CellValue evaluateFormula(const TXSheet* sheet, row_t currentRow, column_t currentCol);


        // ==================== 数字格式化功能 ====================
        /**
         * @brief 设置单元格的自定义数字格式。
         * @param format_string 自定义格式字符串 (例如 "0.00%", "#,##0", "yyyy-mm-dd")。
         */
        void setCustomFormat(const std::string& format_string);

        /**
         * @brief 获取单元格的数字格式化对象 (只读)。
         * @return 如果设置了格式化对象，则返回其常量指针；否则返回nullptr。
         */
        [[nodiscard]] const TXNumberFormat* getNumberFormatObject() const;

        /**
         * @brief 设置单元格的数字格式化对象。
         * @param number_format_ptr 指向TXNumberFormat对象的unique_ptr。所有权将转移。
         */
        void setNumberFormatObject(std::unique_ptr<TXNumberFormat> number_format_ptr);

        /**
         * @brief 获取单元格格式化后的显示值。
         * @return 根据单元格的数字格式规则格式化后的字符串值。
         */
        [[nodiscard]] std::string getFormattedValue() const;

        /**
         * @brief 设置单元格的预定义数字格式。
         * @param type 预定义的格式类型 (来自 TXNumberFormat::FormatType)。
         * @param decimalPlaces 小数位数 (如果适用)。
         * @param useThousandSeparator 是否使用千位分隔符 (如果适用)。
         */
        void setPredefinedFormat(TXNumberFormat::FormatType type, int decimalPlaces = 2, bool useThousandSeparator = true);


        // ==================== 合并单元格功能 ====================
        [[nodiscard]] bool isMerged() const;
        void setMerged(bool merged);
        [[nodiscard]] bool isMasterCell() const;
        void setMasterCell(bool master);
        [[nodiscard]] std::pair<row_t::index_t, column_t::index_t> getMasterCellPosition() const; // 返回原始索引类型
        void setMasterCellPosition(row_t::index_t row_idx, column_t::index_t col_idx);


        // ==================== 样式方法 ====================
        [[nodiscard]] bool hasStyle() const;
        [[nodiscard]] u32 getStyleIndex() const; // u32 来自 TXTypes.hpp
        void setStyleIndex(u32 index);


        // ==================== 工具方法 ====================
        /**
         * @brief 清空单元格内容、公式、格式和样式信息。
         */
        void clear();

        /**
         * @brief 将单元格内容转换为通用字符串表示 (通常等同于getFormattedValue或getStringValue)。
         * @return 单元格内容的字符串表示。
         */
        [[nodiscard]] std::string toString() const;

        /**
         * @brief 从字符串解析并设置单元格值。
         * @param str 要解析的字符串。
         * @param auto_detect_type 是否尝试自动检测类型 (例如，数字、布尔值)。
         * 如果为false，则直接设置为字符串类型。
         * @return 如果解析和设置成功，则返回true。
         */
        bool fromString(const std::string& str, bool auto_detect_type = true);

        /**
         * @brief 创建并返回当前单元格的深拷贝。
         * @return 指向新创建的TXCell副本的unique_ptr。
         */
        [[nodiscard]] std::unique_ptr<TXCell> clone() const;

        /**
         * @brief 将当前单元格的格式信息 (数字格式、样式索引) 复制到目标单元格。
         * @param target 要将格式复制到的目标TXCell对象。
         */
        void copyFormatTo(TXCell& target) const;

        /**
         * @brief 比较当前单元格的值是否与另一个单元格的值相等 (忽略格式和样式)。
         * @param other 要比较的另一个TXCell对象。
         * @return 如果值相等，则返回true。
         */
        [[nodiscard]] bool isValueEqual(const TXCell& other) const;

        // ==================== 保护功能 ====================
        
        /**
         * @brief 设置单元格锁定状态
         * @param locked 是否锁定
         */
        void setLocked(bool locked);
        
        /**
         * @brief 检查单元格是否锁定
         * @return 锁定返回true，否则返回false
         */
        [[nodiscard]] bool isLocked() const;
        
        /**
         * @brief 检查单元格是否有公式
         * @return 有公式返回true，否则返回false
         */
        [[nodiscard]] bool hasFormula() const;


        // ==================== 类型转换操作符 ====================
        explicit operator std::string() const { return getFormattedValue(); }
        explicit operator double() const { return getNumberValue(); }
        explicit operator int64_t() const { return getIntegerValue(); }
        explicit operator bool() const { return getBooleanValue(); }


        // ==================== 赋值操作符 ====================
        TXCell& operator=(const std::string& value_str);
        TXCell& operator=(const char* value_cstr);
        TXCell& operator=(double value_dbl);
        TXCell& operator=(int64_t value_i64);
        TXCell& operator=(int value_int); // 转换为int64_t
        TXCell& operator=(bool value_bool);


        // ==================== 比较操作符 (基于值) ====================
        bool operator==(const TXCell& other) const;
        bool operator!=(const TXCell& other) const;
        // 注意: <, <=, >, >= 的比较对于std::variant可能没有直观意义，除非定义了自定义比较逻辑
        // 这里提供的比较将依赖于 std::variant 的默认比较行为，它会比较 index() 和具体类型的值
        bool operator<(const TXCell& other) const;
        bool operator<=(const TXCell& other) const;
        bool operator>(const TXCell& other) const;
        bool operator>=(const TXCell& other) const;

    private:
        CellValue value_;                              ///< 存储单元格的实际数据 (std::variant)
        CellType type_;                                ///< 单元格的当前类型
        std::unique_ptr<TXFormula> formula_object_;    ///< 指向公式对象的智能指针 (如果单元格是公式)
        std::unique_ptr<TXNumberFormat> number_format_object_; ///< 指向数字格式化对象的智能指针

        // 合并单元格相关状态
        bool is_merged_ = false;                       ///< 是否是合并单元格的一部分
        bool is_master_cell_ = false;                  ///< 是否是合并区域的左上角主单元格
        row_t::index_t master_row_idx_ = 0;            ///< 主单元格的行索引 (如果is_merged_且非is_master_cell_)
        column_t::index_t master_col_idx_ = 0;         ///< 主单元格的列索引

        // 样式相关状态
        bool has_style_ = false;                       ///<单元格是否有显式样式
        u32 style_index_ = 0;                          ///< 应用于此单元格的样式索引 (来自样式管理器)
        
        // 保护相关状态
        bool is_locked_ = true;                        ///< 单元格是否锁定（默认锁定）

        /**
         * @brief 根据value_和formula_object_的内容更新type_成员。
         */
        void updateType();
    };

} // namespace TinaXlsx
