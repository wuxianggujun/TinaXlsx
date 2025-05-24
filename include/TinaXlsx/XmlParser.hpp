/**
 * @file XmlParser.hpp
 * @brief XML解析器组件 - 封装expat库，提供Excel XML解析能力
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <string>
#include <vector>
#include <functional>
#include <unordered_map>
#include <memory>

// 前向声明，避免直接包含expat.h
struct XML_ParserStruct;
typedef struct XML_ParserStruct* XML_Parser;

namespace TinaXlsx {

// Forward declarations for callback function types
using XML_Char = char;

// Callback function types
using ElementStartCallback = std::function<void(const std::string&, const std::vector<std::pair<std::string, std::string>>&)>;
using ElementEndCallback = std::function<void(const std::string&)>;
using CharacterDataCallback = std::function<void(const std::string&)>;

/**
 * @brief XML解析状态枚举
 */
enum class XmlParseState {
    None,
    Workbook,
    SheetData,
    Row,
    Cell,
    Value,
    InlineString,
    SharedString,
    Relationships,
    ContentTypes
};

/**
 * @brief XML解析上下文
 */
struct XmlParseContext {
    XmlParseState state = XmlParseState::None;
    XmlParseState previousState = XmlParseState::None;
    std::string currentValue;
    std::string currentElementName;
    
    void reset() {
        state = XmlParseState::None;
        previousState = XmlParseState::None;
        currentValue.clear();
        currentElementName.clear();
    }
};

/**
 * @brief XML解析器基类
 * 封装expat的使用，提供统一的XML解析接口
 */
class XmlParser {
public:
    XML_Parser parser_ = nullptr;
    XmlParseContext context_;
    ElementStartCallback startCallback_;
    ElementEndCallback endCallback_;
    CharacterDataCallback dataCallback_;

public:
    /**
     * @brief 构造函数
     */
    XmlParser();
    
    /**
     * @brief 虚析构函数
     */
    virtual ~XmlParser();
    
    // 禁用拷贝
    XmlParser(const XmlParser&) = delete;
    XmlParser& operator=(const XmlParser&) = delete;
    
    // 启用移动
    XmlParser(XmlParser&& other) noexcept;
    XmlParser& operator=(XmlParser&& other) noexcept;
    
    /**
     * @brief 设置元素开始回调
     */
    void setElementStartCallback(ElementStartCallback callback) {
        startCallback_ = std::move(callback);
    }
    
    /**
     * @brief 设置元素结束回调
     */
    void setElementEndCallback(ElementEndCallback callback) {
        endCallback_ = std::move(callback);
    }
    
    /**
     * @brief 设置字符数据回调
     */
    void setCharacterDataCallback(CharacterDataCallback callback) {
        dataCallback_ = std::move(callback);
    }
    
    /**
     * @brief 解析XML内容
     * @param content XML内容
     * @param isFinal 是否是最后一块数据
     * @return bool 解析是否成功
     */
    bool parse(const std::string& content, bool isFinal = true);
    
    /**
     * @brief 获取当前解析上下文
     */
    const XmlParseContext& getContext() const { return context_; }
    
    /**
     * @brief 重置解析器
     */
    void reset();

protected:
    /**
     * @brief 初始化解析器
     */
    void initializeParser();
    
    /**
     * @brief 清理解析器
     */
    void cleanupParser();

private:
    static void startElementHandler(void* userData, const XML_Char* name, const XML_Char** atts);
    static void endElementHandler(void* userData, const XML_Char* name);
    static void characterDataHandler(void* userData, const XML_Char* s, int len);
};

/**
 * @brief Excel工作簿XML解析器
 * 专门解析workbook.xml文件
 */
class WorkbookXmlParser : public XmlParser {
public:
    struct SheetInfo {
        std::string name;
        std::string relationId;
        RowIndex sheetId;
    };
    
private:
    std::vector<SheetInfo> sheets_;

public:
    /**
     * @brief 构造函数
     */
    WorkbookXmlParser();
    
    /**
     * @brief 解析工作簿XML
     * @param content XML内容
     * @return std::vector<SheetInfo> 工作表信息列表
     */
    std::vector<SheetInfo> parseWorkbook(const std::string& content);
    
    /**
     * @brief 获取解析到的工作表信息
     */
    const std::vector<SheetInfo>& getSheets() const { return sheets_; }

private:
    void handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes);
    void handleElementEnd(const std::string& name);
};

/**
 * @brief 共享字符串XML解析器
 * 专门解析sharedStrings.xml文件
 */
class SharedStringsXmlParser : public XmlParser {
private:
    std::vector<std::string> sharedStrings_;
    std::string currentString_;

public:
    /**
     * @brief 构造函数
     */
    SharedStringsXmlParser();
    
    /**
     * @brief 解析共享字符串XML
     * @param content XML内容
     * @return std::vector<std::string> 共享字符串列表
     */
    std::vector<std::string> parseSharedStrings(const std::string& content);
    
    /**
     * @brief 获取解析到的共享字符串
     */
    const std::vector<std::string>& getSharedStrings() const { return sharedStrings_; }

private:
    void handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes);
    void handleElementEnd(const std::string& name);
    void handleCharacterData(const std::string& data);
};

/**
 * @brief 关系XML解析器
 * 专门解析.rels文件
 */
class RelationshipsXmlParser : public XmlParser {
public:
    struct Relationship {
        std::string id;
        std::string type;
        std::string target;
    };

private:
    std::vector<Relationship> relationships_;

public:
    /**
     * @brief 构造函数
     */
    RelationshipsXmlParser();
    
    /**
     * @brief 解析关系XML
     * @param content XML内容
     * @return std::vector<Relationship> 关系列表
     */
    std::vector<Relationship> parseRelationships(const std::string& content);
    
    /**
     * @brief 获取解析到的关系
     */
    const std::vector<Relationship>& getRelationships() const { return relationships_; }

private:
    void handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes);
};

/**
 * @brief 工作表XML解析器
 * 专门解析worksheet.xml文件
 */
class WorksheetXmlParser : public XmlParser {
public:
    using CellCallback = std::function<void(const CellPosition& position, const CellValue& value)>;
    using RowCallback = std::function<void(RowIndex rowIndex, const RowData& rowData)>;

private:
    CellCallback cellCallback_;
    RowCallback rowCallback_;
    RowData currentRow_;
    RowIndex currentRowIndex_ = 0;
    std::vector<std::string>* sharedStrings_ = nullptr;

public:
    /**
     * @brief 构造函数
     */
    WorksheetXmlParser();
    
    /**
     * @brief 设置共享字符串
     */
    void setSharedStrings(std::vector<std::string>* sharedStrings) {
        sharedStrings_ = sharedStrings;
    }
    
    /**
     * @brief 设置单元格回调
     */
    void setCellCallback(CellCallback callback) {
        cellCallback_ = std::move(callback);
    }
    
    /**
     * @brief 设置行回调
     */
    void setRowCallback(RowCallback callback) {
        rowCallback_ = std::move(callback);
    }
    
    /**
     * @brief 解析工作表XML
     * @param content XML内容
     */
    void parseWorksheet(const std::string& content);

private:
    void handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes);
    void handleElementEnd(const std::string& name);
    void handleCharacterData(const std::string& data);
    
    CellPosition parseCellReference(const std::string& cellRef);
    CellValue parseSharedString(const std::string& index);
};

} // namespace TinaXlsx 