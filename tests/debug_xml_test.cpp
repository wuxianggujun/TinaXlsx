//
// XML Handler Debug Test
// 用于调试XML声明问题
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXXmlHandler.hpp"
#include <iostream>

TEST(DebugXmlTest, CheckXmlDeclaration) {
    TinaXlsx::TXXmlHandler xml;
    
    std::string test_xml = R"(<?xml version="1.0" encoding="UTF-8"?>
<root>
    <header title="Test Document" version="1.0">
        <author>Test Author</author>
        <date>2024-01-01</date>
    </header>
</root>)";

    std::cout << "=== XML Declaration Debug Test ===" << std::endl;
    
    // 解析XML
    bool parsed = xml.parseFromString(test_xml);
    std::cout << "Parse result: " << (parsed ? "SUCCESS" : "FAILED") << std::endl;
    
    if (parsed) {
        // 保存为字符串（不格式化）
        std::string saved = xml.saveToString(false);
        std::cout << "Saved XML (raw):" << std::endl;
        std::cout << "'" << saved << "'" << std::endl;
        std::cout << "Length: " << saved.length() << std::endl;
        
        // 检查是否包含XML声明
        bool has_declaration = saved.find("<?xml version=\"1.0\" encoding=\"UTF-8\"?>") != std::string::npos;
        std::cout << "Has XML declaration: " << (has_declaration ? "YES" : "NO") << std::endl;
        
        // 保存为字符串（格式化）
        std::string formatted = xml.saveToString(true);
        std::cout << "\nFormatted XML:" << std::endl;
        std::cout << "'" << formatted << "'" << std::endl;
        
        bool formatted_has_declaration = formatted.find("<?xml version=\"1.0\" encoding=\"UTF-8\"?>") != std::string::npos;
        std::cout << "Formatted has XML declaration: " << (formatted_has_declaration ? "YES" : "NO") << std::endl;
    }
    
    std::cout << "=== End XML Debug Test ===" << std::endl;
} 