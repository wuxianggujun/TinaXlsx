#pragma once

#include <cstdint>
#include <string>
#include <limits>
#include <variant>

namespace TinaXlsx {

// ==================== 基础类型定义 ====================

/// 所有基础数值类型别名
using u8 = std::uint8_t;
using u16 = std::uint16_t;
using u32 = std::uint32_t;
using u64 = std::uint64_t;
using i8 = std::int8_t;
using i16 = std::int16_t;
using i32 = std::int32_t;
using i64 = std::int64_t;
using f32 = float;
using f64 = double;

// ==================== 行类型定义 ====================

/**
 * @brief 行索引类型，封装行号的所有操作
 * 支持 1-based 索引，提供完整的行操作接口
 */
class row_t {
public:
    /// 内部索引类型
    using index_t = u32;
    
    /// Excel 最大行数
    static constexpr index_t MAX_ROWS = 1048576;
    
    /// 无效行号
    static constexpr index_t INVALID_ROW = 0;
    
    /**
     * @brief 默认构造函数，指向第一行
     */
    row_t() : index_(1) {}
    
    /**
     * @brief 从行号构造
     * @param row_index 行号 (1-based)
     */
    explicit row_t(index_t row_index) : index_(row_index) {}
    
    /**
     * @brief 获取行索引
     * @return 行索引 (1-based)
     */
    index_t index() const { return index_; }
    
    /**
     * @brief 转换操作符，允许显式转换为 index_t
     * @return 行索引 (1-based)
     */
    explicit operator index_t() const { return index_; }
    
    /**
     * @brief 检查行号是否有效
     * @return 有效返回 true
     */
    bool is_valid() const {
        return index_ >= 1 && index_ <= MAX_ROWS;
    }
    
    /**
     * @brief 转换为字符串
     * @return 行号字符串
     */
    std::string to_string() const {
        return std::to_string(index_);
    }
    
    // ==================== 运算符重载 ====================
    
    bool operator==(const row_t& other) const { return index_ == other.index_; }
    bool operator!=(const row_t& other) const { return index_ != other.index_; }
    bool operator<(const row_t& other) const { return index_ < other.index_; }
    bool operator<=(const row_t& other) const { return index_ <= other.index_; }
    bool operator>(const row_t& other) const { return index_ > other.index_; }
    bool operator>=(const row_t& other) const { return index_ >= other.index_; }
    
    bool operator==(index_t other) const { return index_ == other; }
    bool operator!=(index_t other) const { return index_ != other; }
    bool operator<(index_t other) const { return index_ < other; }
    bool operator<=(index_t other) const { return index_ <= other; }
    bool operator>(index_t other) const { return index_ > other; }
    bool operator>=(index_t other) const { return index_ >= other; }
    
    /// 前置递增
    row_t& operator++() { ++index_; return *this; }
    /// 前置递减
    row_t& operator--() { --index_; return *this; }
    /// 后置递增
    row_t operator++(int) { row_t tmp(*this); ++index_; return tmp; }
    /// 后置递减
    row_t operator--(int) { row_t tmp(*this); --index_; return tmp; }
    
    /// 加法运算
    row_t operator+(index_t offset) const { return row_t(index_ + offset); }
    /// 减法运算
    row_t operator-(index_t offset) const { return row_t(index_ - offset); }
    /// 复合赋值
    row_t& operator+=(index_t offset) { index_ += offset; return *this; }
    row_t& operator-=(index_t offset) { index_ -= offset; return *this; }
    
    /// 友元运算符
    friend row_t operator+(index_t lhs, const row_t& rhs) { return row_t(lhs + rhs.index_); }
    friend bool operator==(index_t lhs, const row_t& rhs) { return lhs == rhs.index_; }
    friend bool operator!=(index_t lhs, const row_t& rhs) { return lhs != rhs.index_; }
    friend bool operator<(index_t lhs, const row_t& rhs) { return lhs < rhs.index_; }
    friend bool operator<=(index_t lhs, const row_t& rhs) { return lhs <= rhs.index_; }
    friend bool operator>(index_t lhs, const row_t& rhs) { return lhs > rhs.index_; }
    friend bool operator>=(index_t lhs, const row_t& rhs) { return lhs >= rhs.index_; }
    
    /// 创建静态函数
    static row_t first() { return row_t(1); }
    static row_t last() { return row_t(MAX_ROWS); }

private:
    index_t index_;
};

// ==================== 列类型定义 ====================

/**
 * @brief 列索引类型，封装列号和列名的相互转换
 * 支持 1-based 索引和 A-Z,AA-ZZ,AAA-ZZZ 格式的列名
 */
class column_t {
public:
    /// 内部索引类型
    using index_t = u32;
    
    /// Excel 最大列数
    static constexpr index_t MAX_COLUMNS = 16384;
    
    /// 无效列号
    static constexpr index_t INVALID_COLUMN = 0;
    
    /**
     * @brief 将列名转换为列索引
     * @param column_string 列名 (如 "A", "B", "AA")
     * @return 列索引 (1-based)
     */
    static index_t column_index_from_string(const std::string& column_string);
    
    /**
     * @brief 将列索引转换为列名
     * @param column_index 列索引 (1-based)
     * @return 列名 (如 "A", "B", "AA")
     */
    static std::string column_string_from_index(index_t column_index);
    
    /**
     * @brief 默认构造函数，指向 A 列
     */
    column_t() : index_(1) {}
    
    /**
     * @brief 从列索引构造
     * @param column_index 列索引 (1-based)
     */
    explicit column_t(index_t column_index) : index_(column_index) {}
    
    /**
     * @brief 从列名构造
     * @param column_string 列名 (如 "A", "B", "AA")
     */
    explicit column_t(const std::string& column_string) 
        : index_(column_index_from_string(column_string)) {}
    
    /**
     * @brief 从 C 字符串构造
     * @param column_string 列名
     */
    explicit column_t(const char* column_string) 
        : index_(column_index_from_string(std::string(column_string))) {}
    
    /**
     * @brief 获取列索引
     * @return 列索引 (1-based)
     */
    index_t index() const { return index_; }
    
    /**
     * @brief 转换操作符，允许显式转换为 index_t
     * @return 列索引 (1-based)
     */
    explicit operator index_t() const { return index_; }
    
    /**
     * @brief 获取列名
     * @return 列名字符串
     */
    std::string column_string() const {
        return column_string_from_index(index_);
    }
    
    /**
     * @brief 检查列号是否有效
     * @return 有效返回 true
     */
    bool is_valid() const {
        return index_ >= 1 && index_ <= MAX_COLUMNS;
    }
    
    /**
     * @brief 转换为字符串（等同于 column_string）
     * @return 列名字符串
     */
    std::string to_string() const { return column_string(); }
    
    // ==================== 赋值运算符 ====================
    
    column_t& operator=(const std::string& rhs) {
        index_ = column_index_from_string(rhs);
        return *this;
    }
    
    column_t& operator=(const char* rhs) {
        index_ = column_index_from_string(std::string(rhs));
        return *this;
    }
    
    // ==================== 比较运算符 ====================
    
    bool operator==(const column_t& other) const { return index_ == other.index_; }
    bool operator!=(const column_t& other) const { return index_ != other.index_; }
    bool operator<(const column_t& other) const { return index_ < other.index_; }
    bool operator<=(const column_t& other) const { return index_ <= other.index_; }
    bool operator>(const column_t& other) const { return index_ > other.index_; }
    bool operator>=(const column_t& other) const { return index_ >= other.index_; }
    
    bool operator==(index_t other) const { return index_ == other; }
    bool operator!=(index_t other) const { return index_ != other; }
    bool operator<(index_t other) const { return index_ < other; }
    bool operator<=(index_t other) const { return index_ <= other; }
    bool operator>(index_t other) const { return index_ > other; }
    bool operator>=(index_t other) const { return index_ >= other; }
    
    bool operator==(const std::string& other) const {
        return column_string() == other;
    }
    bool operator!=(const std::string& other) const {
        return column_string() != other;
    }
    
    bool operator==(const char* other) const {
        return column_string() == std::string(other);
    }
    bool operator!=(const char* other) const {
        return column_string() != std::string(other);
    }
    
    // ==================== 递增递减运算符 ====================
    
    /// 前置递增
    column_t& operator++() { ++index_; return *this; }
    /// 前置递减
    column_t& operator--() { --index_; return *this; }
    /// 后置递增
    column_t operator++(int) { column_t tmp(*this); ++index_; return tmp; }
    /// 后置递减
    column_t operator--(int) { column_t tmp(*this); --index_; return tmp; }
    
    // ==================== 算术运算符 ====================
    
    column_t operator+(const column_t& rhs) const { return column_t(index_ + rhs.index_); }
    column_t operator-(const column_t& rhs) const { return column_t(index_ - rhs.index_); }
    column_t operator+(index_t offset) const { return column_t(index_ + offset); }
    column_t operator-(index_t offset) const { return column_t(index_ - offset); }
    
    column_t& operator+=(const column_t& rhs) { index_ += rhs.index_; return *this; }
    column_t& operator-=(const column_t& rhs) { index_ -= rhs.index_; return *this; }
    column_t& operator+=(index_t offset) { index_ += offset; return *this; }
    column_t& operator-=(index_t offset) { index_ -= offset; return *this; }
    
    /// 友元运算符
    friend column_t operator+(index_t lhs, const column_t& rhs) { return column_t(lhs + rhs.index_); }
    friend bool operator==(index_t lhs, const column_t& rhs) { return lhs == rhs.index_; }
    friend bool operator!=(index_t lhs, const column_t& rhs) { return lhs != rhs.index_; }
    friend bool operator<(index_t lhs, const column_t& rhs) { return lhs < rhs.index_; }
    friend bool operator<=(index_t lhs, const column_t& rhs) { return lhs <= rhs.index_; }
    friend bool operator>(index_t lhs, const column_t& rhs) { return lhs > rhs.index_; }
    friend bool operator>=(index_t lhs, const column_t& rhs) { return lhs >= rhs.index_; }
    
    /// 创建静态函数
    static column_t first() { return column_t(1); }
    static column_t last() { return column_t(MAX_COLUMNS); }

private:
    index_t index_;
};

// ==================== 其他类型定义 ====================

/// 工作表索引类型
using sheet_index_t = u32;

/// 颜色值类型 (ARGB 格式)
using color_value_t = u32;

/// 字体大小类型
using font_size_t = u32;

/// 边框宽度类型
using border_width_t = u32;

/// 单元格数值类型
using cell_double_t = f64;
using cell_integer_t = i64;

/// 统一的值类型，用于单元格、公式计算和格式化
using cell_value_t = std::variant<std::monostate, std::string, f64, i64, bool>;

// ==================== 常量定义 ====================

/// 工作表名称最大长度
constexpr std::size_t MAX_SHEET_NAME = 31;

/// 无效值常量
constexpr sheet_index_t INVALID_SHEET = std::numeric_limits<sheet_index_t>::max();

/// 默认值常量
constexpr font_size_t DEFAULT_FONT_SIZE = 11;
constexpr color_value_t DEFAULT_COLOR = 0xFF000000; // 黑色
constexpr border_width_t DEFAULT_BORDER_WIDTH = 1;

// ==================== 文件格式常量 ====================

/// 支持的文件扩展名
constexpr const char* XLSX_EXTENSION = ".xlsx";
constexpr const char* XLS_EXTENSION = ".xls";
constexpr const char* CSV_EXTENSION = ".csv";

/// MIME 类型
constexpr const char* XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
constexpr const char* XLS_MIME_TYPE = "application/vnd.ms-excel";
constexpr const char* CSV_MIME_TYPE = "text/csv";

// ==================== 工具函数 ====================

/**
 * @brief 检查坐标是否有效
 * @param row 行对象
 * @param col 列对象
 * @return 有效返回 true
 */
inline bool is_valid_coordinate(const row_t& row, const column_t& col) {
    return row.is_valid() && col.is_valid();
}

/**
 * @brief 检查字体大小是否有效
 * @param size 字体大小
 * @return 有效返回 true
 */
constexpr bool is_valid_font_size(font_size_t size) {
    return size >= 1 && size <= 72;
}

/**
 * @brief 检查工作表名称是否有效
 * @param name 工作表名称
 * @return 有效返回 true
 */
bool is_valid_sheet_name(const std::string& name);

/**
 * @brief 计算曼哈顿距离
 * @param row1 起始行
 * @param col1 起始列
 * @param row2 结束行
 * @param col2 结束列
 * @return 曼哈顿距离
 */
inline std::size_t manhattan_distance(const row_t& row1, const column_t& col1,
                                     const row_t& row2, const column_t& col2) {
    const auto r1 = row1.index(), c1 = col1.index();
    const auto r2 = row2.index(), c2 = col2.index();
    return static_cast<std::size_t>(
        (r1 > r2 ? r1 - r2 : r2 - r1) + 
        (c1 > c2 ? c1 - c2 : c2 - c1)
    );
}

/**
 * @brief 获取文件扩展名
 * @param filename 文件名
 * @return 扩展名（含点号）
 */
std::string get_file_extension(const std::string& filename);

/**
 * @brief 检查是否为 Excel 文件
 * @param filename 文件名
 * @return 是 Excel 文件返回 true
 */
bool is_excel_file(const std::string& filename);

} // namespace TinaXlsx

// ==================== 哈希支持 ====================

namespace std {

/// row_t 哈希支持
template <>
struct hash<TinaXlsx::row_t> {
    std::size_t operator()(const TinaXlsx::row_t& r) const {
        return std::hash<TinaXlsx::row_t::index_t>()(r.index());
    }
};

/// column_t 哈希支持
template <>
struct hash<TinaXlsx::column_t> {
    std::size_t operator()(const TinaXlsx::column_t& c) const {
        return std::hash<TinaXlsx::column_t::index_t>()(c.index());
    }
};

} // namespace std 
