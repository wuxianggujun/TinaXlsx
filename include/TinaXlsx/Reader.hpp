/**
 * @file Reader.hpp
 * @brief High performance Excel reader
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <memory>
#include <string>
#include <vector>
#include <functional>
#include <optional>

namespace TinaXlsx {

/**
 * @brief High performance Excel reader
 * 
 * High performance reader based on minizip-ng and expat, replacing xlsxio
 * Supports streaming and batch reading, uses RAII for resource management, ensures exception safety
 */
class Reader {
public:
    /**
     * @brief Row reading callback function type
     * @param rowIndex Row index (0-based)
     * @param rowData Row data
     * @return bool true to continue reading, false to stop
     */
    using RowCallback = std::function<bool(RowIndex rowIndex, const RowData& rowData)>;
    
    /**
     * @brief Cell reading callback function type
     * @param position Cell position
     * @param value Cell value
     * @return bool true to continue reading, false to stop
     */
    using CellCallback = std::function<bool(const CellPosition& position, const CellValue& value)>;

private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;

public:
    /**
     * @brief Constructor
     * @param filePath Excel file path
     * @throws FileException if file cannot be opened
     */
    explicit Reader(const std::string& filePath);
    
    /**
     * @brief Move constructor
     */
    Reader(Reader&& other) noexcept;
    
    /**
     * @brief Move assignment operator
     */
    Reader& operator=(Reader&& other) noexcept;
    
    /**
     * @brief Destructor
     */
    ~Reader();
    
    // Disable copy constructor and copy assignment
    Reader(const Reader&) = delete;
    Reader& operator=(const Reader&) = delete;
    
    /**
     * @brief Get list of worksheet names
     * @return std::vector<std::string> List of worksheet names
     */
    [[nodiscard]] std::vector<std::string> getSheetNames() const;
    
    /**
     * @brief Open specified worksheet
     * @param sheetName Worksheet name
     * @return bool Whether successfully opened
     */
    bool openSheet(const std::string& sheetName);
    
    /**
     * @brief Open worksheet by index
     * @param sheetIndex Worksheet index (0-based)
     * @return bool Whether successfully opened
     */
    bool openSheet(SheetIndex sheetIndex);
    
    /**
     * @brief Get row count of current worksheet (estimated)
     * @return std::optional<RowIndex> Row count, empty if cannot determine
     */
    [[nodiscard]] std::optional<RowIndex> getRowCount() const;
    
    /**
     * @brief Get column count of current worksheet (estimated)
     * @return std::optional<ColumnIndex> Column count, empty if cannot determine
     */
    [[nodiscard]] std::optional<ColumnIndex> getColumnCount() const;
    
    /**
     * @brief Streaming read next row data
     * @param rowData Output row data
     * @param maxColumns Maximum number of columns, 0 means unlimited
     * @return bool Whether successfully read data
     */
    bool readNextRow(RowData& rowData, ColumnIndex maxColumns = 0);
    
    /**
     * @brief Read data of specified row
     * @param rowIndex Row index (0-based)
     * @param maxColumns Maximum number of columns, 0 means unlimited
     * @return std::optional<RowData> Row data, empty if row doesn't exist
     */
    [[nodiscard]] std::optional<RowData> readRow(RowIndex rowIndex, ColumnIndex maxColumns = 0);
    
    /**
     * @brief Read value of specified cell
     * @param position Cell position
     * @return std::optional<CellValue> Cell value, empty if cell doesn't exist
     */
    [[nodiscard]] std::optional<CellValue> readCell(const CellPosition& position);
    
    /**
     * @brief Read data of specified range
     * @param range Cell range
     * @return TableData Data within the range
     */
    [[nodiscard]] TableData readRange(const CellRange& range);
    
    /**
     * @brief Read entire worksheet data
     * @param maxRows Maximum number of rows, 0 means unlimited
     * @param maxColumns Maximum number of columns, 0 means unlimited
     * @param skipEmptyRows Whether to skip empty rows
     * @return TableData Worksheet data
     */
    [[nodiscard]] TableData readAll(RowIndex maxRows = 0, ColumnIndex maxColumns = 0, bool skipEmptyRows = true);
    
    /**
     * @brief Streaming read all rows using callback function
     * @param callback Row reading callback function
     * @param maxColumns Maximum number of columns, 0 means unlimited
     * @param skipEmptyRows Whether to skip empty rows
     * @return RowIndex Actual number of rows read
     */
    RowIndex readAllRows(const RowCallback& callback, ColumnIndex maxColumns = 0, bool skipEmptyRows = true);
    
    /**
     * @brief Streaming read all cells using callback function
     * @param callback Cell reading callback function
     * @param range Reading range, if empty then read entire worksheet
     * @return size_t Actual number of cells read
     */
    size_t readAllCells(const CellCallback& callback, const std::optional<CellRange>& range = std::nullopt);
    
    /**
     * @brief Reset reading position to worksheet beginning
     */
    void reset();
    
    /**
     * @brief Check if reached end of worksheet
     * @return bool Whether reached end
     */
    [[nodiscard]] bool isAtEnd() const;
    
    /**
     * @brief Get current reading position
     * @return RowIndex Current row index
     */
    [[nodiscard]] RowIndex getCurrentRowIndex() const;
    
    /**
     * @brief Check if specified row is empty
     * @param rowData Row data
     * @return bool Whether is empty row
     */
    [[nodiscard]] static bool isEmptyRow(const RowData& rowData);
    
    /**
     * @brief Check if specified cell is empty
     * @param value Cell value
     * @return bool Whether is empty
     */
    [[nodiscard]] static bool isEmptyCell(const CellValue& value);
    
    /**
     * @brief Convert string to CellValue
     * @param str String
     * @return CellValue Converted value
     */
    [[nodiscard]] static CellValue stringToCellValue(const std::string& str);
    
    /**
     * @brief Convert CellValue to string
     * @param value Cell value
     * @return std::string String representation
     */
    [[nodiscard]] static std::string cellValueToString(const CellValue& value);
};

} // namespace TinaXlsx 