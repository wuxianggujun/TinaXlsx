#include <gtest/gtest.h>
#include "TinaXlsx/TXXmlHandler.hpp"
#include <fstream>
#include <filesystem>

class TXXmlHandlerTest : public ::testing::Test {
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

TEST_F(TXXmlHandlerTest, ParseFromString) {
    TinaXlsx::TXXmlHandler xml;
    
    // 测试解析XML字符串
    EXPECT_TRUE(xml.parseFromString(test_xml_));
    EXPECT_TRUE(xml.isValid());
    EXPECT_EQ(xml.getRootName(), "root");
    
    // 测试无效XML
    EXPECT_FALSE(xml.parseFromString("<invalid><xml>"));
    EXPECT_FALSE(xml.getLastError().empty());
}

TEST_F(TXXmlHandlerTest, ParseFromFile) {
    TinaXlsx::TXXmlHandler xml;
    
    // 测试从文件解析
    EXPECT_TRUE(xml.parseFromFile("test_files/test.xml"));
    EXPECT_TRUE(xml.isValid());
    EXPECT_EQ(xml.getRootName(), "root");
    
    // 测试不存在的文件
    EXPECT_FALSE(xml.parseFromFile("nonexistent.xml"));
    EXPECT_FALSE(xml.getLastError().empty());
}

TEST_F(TXXmlHandlerTest, SaveToString) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试保存为字符串
    std::string saved_xml = xml.saveToString(false);
    EXPECT_FALSE(saved_xml.empty());
    EXPECT_TRUE(saved_xml.find("<?xml version=\"1.0\" encoding=\"UTF-8\"?>") != std::string::npos);
    
    // 测试格式化输出
    std::string formatted_xml = xml.saveToString(true);
    EXPECT_FALSE(formatted_xml.empty());
    EXPECT_TRUE(formatted_xml.find("  <header") != std::string::npos); // 应该有缩进
}

TEST_F(TXXmlHandlerTest, SaveToFile) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试保存到文件
    EXPECT_TRUE(xml.saveToFile("test_files/output.xml"));
    EXPECT_TRUE(std::filesystem::exists("test_files/output.xml"));
    
    // 验证保存的文件可以被重新解析
    TinaXlsx::TXXmlHandler xml2;
    EXPECT_TRUE(xml2.parseFromFile("test_files/output.xml"));
    EXPECT_EQ(xml2.getRootName(), "root");
}

TEST_F(TXXmlHandlerTest, FindNodes) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试查找节点
    auto sections = xml.findNodes("//section");
    EXPECT_EQ(sections.size(), 2);
    
    auto paragraphs = xml.findNodes("//paragraph");
    EXPECT_EQ(paragraphs.size(), 3);
    
    // 测试查找单个节点
    auto header = xml.findNode("//header");
    EXPECT_EQ(header.name, "header");
    EXPECT_EQ(header.attributes["title"], "Test Document");
    EXPECT_EQ(header.attributes["version"], "1.0");
    
    // 测试不存在的节点
    auto nonexistent = xml.findNodes("//nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlHandlerTest, GetNodeText) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试获取节点文本
    std::string author = xml.getNodeText("//author");
    EXPECT_EQ(author, "Test Author");
    
    std::string date = xml.getNodeText("//date");
    EXPECT_EQ(date, "2024-01-01");
    
    std::string first_paragraph = xml.getNodeText("//paragraph[1]");
    EXPECT_EQ(first_paragraph, "This is the first paragraph.");
    
    // 测试不存在的节点
    std::string nonexistent = xml.getNodeText("//nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlHandlerTest, GetNodeAttribute) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试获取节点属性
    std::string title = xml.getNodeAttribute("//header", "title");
    EXPECT_EQ(title, "Test Document");
    
    std::string version = xml.getNodeAttribute("//header", "version");
    EXPECT_EQ(version, "1.0");
    
    std::string section_id = xml.getNodeAttribute("//section[1]", "id");
    EXPECT_EQ(section_id, "1");
    
    // 测试不存在的属性
    std::string nonexistent = xml.getNodeAttribute("//header", "nonexistent");
    EXPECT_TRUE(nonexistent.empty());
}

TEST_F(TXXmlHandlerTest, SetNodeText) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试设置节点文本
    EXPECT_TRUE(xml.setNodeText("//author", "New Author"));
    std::string author = xml.getNodeText("//author");
    EXPECT_EQ(author, "New Author");
    
    EXPECT_TRUE(xml.setNodeText("//date", "2024-12-31"));
    std::string date = xml.getNodeText("//date");
    EXPECT_EQ(date, "2024-12-31");
    
    // 测试设置不存在节点的文本
    EXPECT_FALSE(xml.setNodeText("//nonexistent", "value"));
}

TEST_F(TXXmlHandlerTest, SetNodeAttribute) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试设置现有属性
    EXPECT_TRUE(xml.setNodeAttribute("//header", "title", "Updated Title"));
    std::string title = xml.getNodeAttribute("//header", "title");
    EXPECT_EQ(title, "Updated Title");
    
    // 测试添加新属性
    EXPECT_TRUE(xml.setNodeAttribute("//header", "language", "en"));
    std::string language = xml.getNodeAttribute("//header", "language");
    EXPECT_EQ(language, "en");
    
    // 测试设置不存在节点的属性
    EXPECT_FALSE(xml.setNodeAttribute("//nonexistent", "attr", "value"));
}

TEST_F(TXXmlHandlerTest, AddChildNode) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试添加子节点
    EXPECT_TRUE(xml.addChildNode("//header", "description", "This is a test document"));
    
    // 验证节点已添加
    std::string description = xml.getNodeText("//description");
    EXPECT_EQ(description, "This is a test document");
    
    // 测试添加空文本的节点
    EXPECT_TRUE(xml.addChildNode("//header", "empty_node"));
    auto empty_nodes = xml.findNodes("//empty_node");
    EXPECT_EQ(empty_nodes.size(), 1);
    
    // 测试添加到不存在的父节点
    EXPECT_FALSE(xml.addChildNode("//nonexistent", "child"));
}

TEST_F(TXXmlHandlerTest, RemoveNodes) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试删除节点
    auto items_before = xml.findNodes("//item");
    EXPECT_EQ(items_before.size(), 3);
    
    std::size_t removed = xml.removeNodes("//item[2]");
    EXPECT_EQ(removed, 1);
    
    auto items_after = xml.findNodes("//item");
    EXPECT_EQ(items_after.size(), 2);
    
    // 测试删除多个节点
    removed = xml.removeNodes("//paragraph");
    EXPECT_EQ(removed, 3);
    
    auto paragraphs_after = xml.findNodes("//paragraph");
    EXPECT_TRUE(paragraphs_after.empty());
}

TEST_F(TXXmlHandlerTest, BatchOperations) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试批量查找节点
    std::vector<std::string> xpaths = {"//author", "//date", "//section"};
    auto batch_results = xml.batchFindNodes(xpaths);
    
    EXPECT_EQ(batch_results.size(), 3);
    EXPECT_EQ(batch_results["//author"].size(), 1);
    EXPECT_EQ(batch_results["//date"].size(), 1);
    EXPECT_EQ(batch_results["//section"].size(), 2);
    
    // 测试批量设置节点文本
    std::unordered_map<std::string, std::string> text_updates;
    text_updates["//author"] = "Batch Author";
    text_updates["//date"] = "Batch Date";
    
    std::size_t updated = xml.batchSetNodeText(text_updates);
    EXPECT_EQ(updated, 2);
    
    EXPECT_EQ(xml.getNodeText("//author"), "Batch Author");
    EXPECT_EQ(xml.getNodeText("//date"), "Batch Date");
}

TEST_F(TXXmlHandlerTest, CreateDocument) {
    TinaXlsx::TXXmlHandler xml;
    
    // 测试创建新文档
    EXPECT_TRUE(xml.createDocument("books", "UTF-8"));
    EXPECT_TRUE(xml.isValid());
    EXPECT_EQ(xml.getRootName(), "books");
    
    // 添加一些内容
    EXPECT_TRUE(xml.addChildNode("/books", "book"));
    EXPECT_TRUE(xml.setNodeAttribute("//book", "id", "1"));
    EXPECT_TRUE(xml.addChildNode("//book", "title", "Test Book"));
    EXPECT_TRUE(xml.addChildNode("//book", "author", "Test Author"));
    
    // 验证创建的文档
    std::string title = xml.getNodeText("//title");
    EXPECT_EQ(title, "Test Book");
    
    std::string book_id = xml.getNodeAttribute("//book", "id");
    EXPECT_EQ(book_id, "1");
}

TEST_F(TXXmlHandlerTest, DocumentStats) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    
    // 测试获取文档统计信息
    auto stats = xml.getDocumentStats();
    
    EXPECT_GT(stats.total_nodes, 0);
    EXPECT_GT(stats.total_attributes, 0);
    EXPECT_GT(stats.max_depth, 1);
    EXPECT_GT(stats.document_size, 0);
    
    // 具体验证一些统计信息
    EXPECT_GE(stats.total_attributes, 4); // 至少有title, version, id, name等属性
}

TEST_F(TXXmlHandlerTest, ParseOptions) {
    TinaXlsx::TXXmlHandler::ParseOptions options;
    options.preserve_whitespace = true;
    options.trim_pcdata = false;
    
    TinaXlsx::TXXmlHandler xml(options);
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    EXPECT_TRUE(xml.isValid());
}

TEST_F(TXXmlHandlerTest, MoveSemantics) {
    TinaXlsx::TXXmlHandler xml1;
    ASSERT_TRUE(xml1.parseFromString(test_xml_));
    
    // 测试移动构造
    TinaXlsx::TXXmlHandler xml2 = std::move(xml1);
    EXPECT_TRUE(xml2.isValid());
    EXPECT_EQ(xml2.getRootName(), "root");
    
    // 测试移动赋值
    TinaXlsx::TXXmlHandler xml3;
    xml3 = std::move(xml2);
    EXPECT_TRUE(xml3.isValid());
    EXPECT_EQ(xml3.getRootName(), "root");
}

TEST_F(TXXmlHandlerTest, Reset) {
    TinaXlsx::TXXmlHandler xml;
    ASSERT_TRUE(xml.parseFromString(test_xml_));
    EXPECT_TRUE(xml.isValid());
    
    // 测试重置
    xml.reset();
    EXPECT_FALSE(xml.isValid());
    EXPECT_TRUE(xml.getRootName().empty());
    EXPECT_TRUE(xml.getLastError().empty());
} 