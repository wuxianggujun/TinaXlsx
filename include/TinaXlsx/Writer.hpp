/**
 * @file Writer.hpp
 * @brief High performance Excel writer
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include "Format.hpp"
#include <memory>
#include <string>
#include <vector>
#include <optional>

// Forward declarations to avoid including xlsxwriter.h in public header
struct lxw_workbook;

namespace TinaXlsx {

// Forward declaration
class Worksheet;

/**
 * @brief High performance Excel writer
 * 
 * High performance writer based on libxlsxwriter, supports batch writing and streaming writing
 * Uses RAII for resource management, ensures exception safety
 */
class Writer {
private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;

public:
    /**
     * @brief Constructor
     * @param filePath Excel file path
     * @throws FileException if file cannot be created
     */
    explicit Writer(const std::string& filePath);
    
    /**
     * @brief Move constructor
     */
    Writer(Writer&& other) noexcept;
    
    /**
     * @brief Move assignment operator
     */
    Writer& operator=(Writer&& other) noexcept;
    
    /**
     * @brief Destructor
     */
    ~Writer();
    
    // Disable copy constructor and copy assignment
    Writer(const Writer&) = delete;
    Writer& operator=(const Writer&) = delete;
    
    /**
     * @brief Create new worksheet
     * @param sheetName Worksheet name, if empty then use default name
     * @return std::shared_ptr<Worksheet> Worksheet object
     */
    [[nodiscard]] std::shared_ptr<Worksheet> createWorksheet(const std::string& sheetName = "");
    
    /**
     * @brief Get worksheet by name
     * @param sheetName Worksheet name
     * @return std::shared_ptr<Worksheet> Worksheet object, nullptr if not found
     */
    [[nodiscard]] std::shared_ptr<Worksheet> getWorksheet(const std::string& sheetName);
    
    /**
     * @brief Get worksheet by index
     * @param index Worksheet index (0-based)
     * @return std::shared_ptr<Worksheet> Worksheet object, nullptr if not found
     */
    [[nodiscard]] std::shared_ptr<Worksheet> getWorksheet(SheetIndex index);
    
    /**
     * @brief Get all worksheet names
     * @return std::vector<std::string> List of worksheet names
     */
    [[nodiscard]] std::vector<std::string> getWorksheetNames() const;
    
    /**
     * @brief Get worksheet count
     * @return size_t Number of worksheets
     */
    [[nodiscard]] size_t getWorksheetCount() const;
    
    /**
     * @brief Create format object
     * @return std::unique_ptr<Format> Format object
     */
    [[nodiscard]] std::unique_ptr<Format> createFormat();
    
    /**
     * @brief Get format builder
     * @return FormatBuilder Format builder
     */
    [[nodiscard]] FormatBuilder getFormatBuilder();
    
    /**
     * @brief Set document properties
     * @param title Title
     * @param subject Subject
     * @param author Author
     * @param manager Manager
     * @param company Company
     * @param category Category
     * @param keywords Keywords
     * @param comments Comments
     */
    void setDocumentProperties(
        const std::string& title = "",
        const std::string& subject = "",
        const std::string& author = "",
        const std::string& manager = "",
        const std::string& company = "",
        const std::string& category = "",
        const std::string& keywords = "",
        const std::string& comments = "");
    
    /**
     * @brief Set custom property (string)
     * @param name Property name
     * @param value Property value
     */
    void setCustomProperty(const std::string& name, const std::string& value);
    
    /**
     * @brief Set custom property (number)
     * @param name Property name
     * @param value Property value
     */
    void setCustomProperty(const std::string& name, double value);
    
    /**
     * @brief Set custom property (boolean)
     * @param name Property name
     * @param value Property value
     */
    void setCustomProperty(const std::string& name, bool value);
    
    /**
     * @brief Save file
     * @return bool Whether successfully saved
     */
    bool save();
    
    /**
     * @brief Close file (automatically calls save)
     * @return bool Whether successfully closed
     */
    bool close();
    
    /**
     * @brief Get internal workbook object pointer (for internal use only)
     * @return lxw_workbook* Internal workbook object pointer
     */
    [[nodiscard]] lxw_workbook* getInternalWorkbook() const;
    
    /**
     * @brief Check if file is closed
     * @return bool Whether closed
     */
    [[nodiscard]] bool isClosed() const;
    
    /**
     * @brief Set default date format
     * @param format Date format string
     */
    void setDefaultDateFormat(const std::string& format);
    
    /**
     * @brief Define named range
     * @param name Range name
     * @param sheetName Worksheet name
     * @param range Cell range
     */
    void defineNamedRange(const std::string& name, const std::string& sheetName, const CellRange& range);
};

} // namespace TinaXlsx 