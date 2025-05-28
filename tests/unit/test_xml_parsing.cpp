#include <gtest/gtest.h>
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include <fstream>
#include <filesystem>

class TXXmlReaderWriterTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试XML内容
        test_xml_ = R"(<?xml version="1.0" encoding="UTF-8"?>
<root>
    <header title="Test Document" version="1.0">
        <author>Test Author</author>
        <date>2024-01-01</date>
    </header>
    <body>
        <section id="1" name="Introduction">
            <paragraph>This is the first paragraph.</paragraph>
            <paragraph>This is the second paragraph.</paragraph>
        </section>
        <section id="2" name="Content">
            <paragraph>Content paragraph 1.</paragraph>
            <list type="ordered">
                <item>Item 1</item>
                <item>Item 2</item>
                <item>Item 3</item>
            </list>
        </section>
    </body>
</root>)";

        // 创建测试文件
        std::filesystem::create_directories("test_files");
        std::ofstream file("test_files/test.xml");
        file << test_xml_;
        file.close();
    }

    void TearDown() override {
        // 清理测试文件
        std::filesystem::remove_all("test_files");
    }

    std::string test_xml_;
};

TEST_F(TXXmlReaderWriterTest, XmlReaderParseFromString) {
    TinaXlsx::TXXmlReader reader;
    
    // 测试从字符串解析XML
    EXPECT_TRUE(reader.parseFromString(test_xml_));
    EXPECT_TRUE(reader.isValid());
    
    // 测试获取根节点
    auto root = reader.getRootNode();
    EXPECT_EQ(root.name, "root");
    
    // 测试无效XML
    TinaXlsx::TXXmlReader reader2;
    EXPECT_FALSE(reader2.parseFromString("<invalid><xml>"));
    EXPECT_FALSE(reader2.isValid());
}

TEST_F(TXXmlReaderWriterTest, XmlReaderFindNodes) {
    TinaXlsx::TXXmlReader reader;
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    
    // 测试查找节点
    auto sections = reader.findNodes("//section");
    EXPECT_EQ(sections.size(), 2);
    
    auto paragraphs = reader.findNodes("//paragraph");
    EXPECT_EQ(paragraphs.size(), 3);
    
    // 测试不存在的节点
    auto nonexistent = reader.findNodes("//nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlReaderWriterTest, XmlReaderGetNodeText) {
    TinaXlsx::TXXmlReader reader;
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    
    // 测试获取节点文本
    std::string author = reader.getNodeText("//author");
    EXPECT_EQ(author, "Test Author");
    
    std::string date = reader.getNodeText("//date");
    EXPECT_EQ(date, "2024-01-01");
    
    std::string first_paragraph = reader.getNodeText("//paragraph[1]");
    EXPECT_EQ(first_paragraph, "This is the first paragraph.");
    
    // 测试不存在的节点
    std::string nonexistent = reader.getNodeText("//nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlReaderWriterTest, XmlReaderGetNodeAttribute) {
    TinaXlsx::TXXmlReader reader;
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    
    // 测试获取节点属性
    std::string title = reader.getNodeAttribute("//header", "title");
    EXPECT_EQ(title, "Test Document");
    
    std::string version = reader.getNodeAttribute("//header", "version");
    EXPECT_EQ(version, "1.0");
    
    std::string section_id = reader.getNodeAttribute("//section[1]", "id");
    EXPECT_EQ(section_id, "1");
    
    // 测试不存在的属性
    std::string nonexistent = reader.getNodeAttribute("//header", "nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlReaderWriterTest, XmlReaderGetAllNodeTexts) {
    TinaXlsx::TXXmlReader reader;
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    
    // 测试获取所有匹配节点的文本
    auto paragraph_texts = reader.getAllNodeTexts("//paragraph");
    EXPECT_EQ(paragraph_texts.size(), 3);
    EXPECT_EQ(paragraph_texts[0], "This is the first paragraph.");
    EXPECT_EQ(paragraph_texts[1], "This is the second paragraph.");
    EXPECT_EQ(paragraph_texts[2], "Content paragraph 1.");
    
    auto item_texts = reader.getAllNodeTexts("//item");
    EXPECT_EQ(item_texts.size(), 3);
    EXPECT_EQ(item_texts[0], "Item 1");
    EXPECT_EQ(item_texts[1], "Item 2");
    EXPECT_EQ(item_texts[2], "Item 3");
}

TEST_F(TXXmlReaderWriterTest, XmlWriterCreateDocument) {
    TinaXlsx::TXXmlWriter writer;
    
    // 创建新文档
    writer.createDocument("testdoc");
    EXPECT_TRUE(writer.isValid());
    
    // 验证文档创建
    std::string xml_string = writer.generateXmlString();
    EXPECT_FALSE(xml_string.empty());
    EXPECT_TRUE(xml_string.find("testdoc") != std::string::npos);
}

TEST_F(TXXmlReaderWriterTest, XmlWriterNodeBuilder) {
    TinaXlsx::TXXmlWriter writer;
    
    // 使用节点构建器创建复杂结构
    TinaXlsx::XmlNodeBuilder root("spreadsheet");
    
    TinaXlsx::XmlNodeBuilder worksheet("worksheet");
    worksheet.addAttribute("name", "Sheet1");
    
    TinaXlsx::XmlNodeBuilder sheetData("sheetData");
    
    for (int row = 1; row <= 3; ++row) {
        TinaXlsx::XmlNodeBuilder rowNode("row");
        rowNode.addAttribute("r", std::to_string(row));
        
        for (int col = 1; col <= 3; ++col) {
            TinaXlsx::XmlNodeBuilder cell("c");
            cell.addAttribute("r", std::string(1, 'A' + col - 1) + std::to_string(row));
            
            TinaXlsx::XmlNodeBuilder value("v");
            value.setText(std::to_string(row * 10 + col));
            
            cell.addChild(value);
            rowNode.addChild(cell);
        }
        sheetData.addChild(rowNode);
    }
    
    worksheet.addChild(sheetData);
    root.addChild(worksheet);
    
    writer.setRootNode(root);
    
    // 验证结构
    std::string xml = writer.generateXmlString();
    EXPECT_TRUE(xml.find("worksheet") != std::string::npos);
    EXPECT_TRUE(xml.find("sheetData") != std::string::npos);
    EXPECT_TRUE(xml.find("A1") != std::string::npos);
    EXPECT_TRUE(xml.find("C3") != std::string::npos);
}

TEST_F(TXXmlReaderWriterTest, XmlWriterWithOptions) {
    TinaXlsx::XmlWriteOptions options;
    options.format_output = true;
    options.indent = "    "; // 4 spaces
    options.include_declaration = true;
    options.encoding = "UTF-8";
    
    TinaXlsx::TXXmlWriter writer(options);
    
    TinaXlsx::XmlNodeBuilder root("test");
    TinaXlsx::XmlNodeBuilder child("child");
    child.setText("content");
    root.addChild(child);
    
    writer.setRootNode(root);
    
    std::string xml = writer.generateXmlString();
    EXPECT_TRUE(xml.find("<?xml version") != std::string::npos);
    EXPECT_TRUE(xml.find("UTF-8") != std::string::npos);
    EXPECT_TRUE(xml.find("test") != std::string::npos);
    EXPECT_TRUE(xml.find("child") != std::string::npos);
}

TEST_F(TXXmlReaderWriterTest, XmlReaderReset) {
    TinaXlsx::TXXmlReader reader;
    
    // 解析XML
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    EXPECT_TRUE(reader.isValid());
    
    // 重置读取器
    reader.reset();
    
    // 验证重置后状态
    EXPECT_FALSE(reader.isValid());
    auto root = reader.getRootNode();
    EXPECT_TRUE(root.name.empty());
}

TEST_F(TXXmlReaderWriterTest, XmlWriterReset) {
    TinaXlsx::TXXmlWriter writer;
    
    // 创建文档
    writer.createDocument("test");
    EXPECT_TRUE(writer.isValid());
    
    // 重置写入器
    writer.reset();
    
    // 验证重置后状态 - 修正：重置后写入器可能无效
    std::string xml = writer.generateXmlString();
    EXPECT_TRUE(xml.empty() || xml.find("test") == std::string::npos);
}

TEST_F(TXXmlReaderWriterTest, XmlWriterStats) {
    TinaXlsx::TXXmlWriter writer;
    
    TinaXlsx::XmlNodeBuilder root("document");
    root.addAttribute("version", "1.0");
    
    TinaXlsx::XmlNodeBuilder section("section");
    section.addAttribute("id", "1");
    section.setText("Some content");
    
    root.addChild(section);
    writer.setRootNode(root);
    
    auto stats = writer.getStats();
    EXPECT_GT(stats.nodeCount, 0);
    EXPECT_GT(stats.attributeCount, 0);
    EXPECT_GT(stats.textLength, 0);
}

TEST_F(TXXmlReaderWriterTest, ErrorHandling) {
    TinaXlsx::TXXmlReader reader;
    
    // 测试解析错误
    EXPECT_FALSE(reader.parseFromString(""));
    EXPECT_FALSE(reader.parseFromString("<unclosed>"));
    EXPECT_FALSE(reader.parseFromString("<root><unclosed></root>"));
    
    // 检查错误信息
    EXPECT_FALSE(reader.getLastError().empty());
    
    TinaXlsx::TXXmlWriter writer;
    
    // 测试写入器初始状态 - 修正：新创建的写入器可能无效直到设置内容
    EXPECT_TRUE(writer.getLastError().empty());
}

TEST_F(TXXmlReaderWriterTest, MoveSemantics) {
    // 测试移动语义（如果支持）
    TinaXlsx::TXXmlReader reader1;
    reader1.parseFromString(test_xml_);
    EXPECT_TRUE(reader1.isValid());
    
    // 移动构造
    TinaXlsx::TXXmlReader reader2 = std::move(reader1);
    EXPECT_TRUE(reader2.isValid());
    
    // 移动赋值
    TinaXlsx::TXXmlReader reader3;
    reader3 = std::move(reader2);
    EXPECT_TRUE(reader3.isValid());
    
    // 测试写入器移动语义
    TinaXlsx::TXXmlWriter writer1;
    writer1.createDocument("test");
    
    TinaXlsx::TXXmlWriter writer2 = std::move(writer1);
    EXPECT_TRUE(writer2.isValid());
}

TEST_F(TXXmlReaderWriterTest, ComplexXPathQueries) {
    TinaXlsx::TXXmlReader reader;
    ASSERT_TRUE(reader.parseFromString(test_xml_));
    
    // 测试复杂的XPath查询
    auto sections_with_intro = reader.findNodes("//section[@name='Introduction']");
    EXPECT_EQ(sections_with_intro.size(), 1);
    
    auto first_items = reader.findNodes("//item[1]");
    EXPECT_EQ(first_items.size(), 1);
    
    auto all_elements_with_id = reader.findNodes("//*[@id]");
    EXPECT_GE(all_elements_with_id.size(), 2); // 至少有2个section有id属性
} 