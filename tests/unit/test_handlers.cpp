//
// Created by wuxianggujun on 2025/5/29.
//

#include <gtest/gtest.h>
#include <cstdio>
#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXWorkbookXmlHandler.hpp"
#include "TinaXlsx/TXStylesXmlHandler.hpp"
#include "TinaXlsx/TXDocumentPropertiesXmlHandler.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <memory>
#include <fstream>

// 测试工作表XML处理器
class WorksheetXmlHandlerTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建一个简单的工作簿和工作表
        workbook = std::make_unique<TinaXlsx::TXWorkbook>();
        auto sheet = std::make_unique<TinaXlsx::TXSheet>("TestSheet", workbook.get());
        
        // 添加一些测试数据
        sheet->setCellValue("A1", std::string("Hello"));
        sheet->setCellValue("B1", 123.45);
        sheet->setCellValue("C1", static_cast<int64_t>(100));
        sheet->setCellValue("A2", std::string("World"));
        sheet->setCellValue("B2", true);
        
        test_sheet = sheet.get();
        workbook->addSheet(std::move(sheet));
    }

    std::unique_ptr<TinaXlsx::TXWorkbook> workbook;
    TinaXlsx::TXSheet* test_sheet;
};

TEST_F(WorksheetXmlHandlerTest, GenerateWorksheetXml) {
    TinaXlsx::TXWorksheetXmlHandler handler(0);
    
    // 创建工作簿上下文
    TinaXlsx::TXWorkbookContext context{
        workbook->getSheets(),
        workbook->getStyleManager(),        workbook->getComponentManager()
    };
    
    // 测试保存功能 - 创建临时ZIP文件
    TinaXlsx::TXZipArchiveWriter zipWriter;
    std::string tempFile = "test_worksheet.xlsx";
    EXPECT_TRUE(zipWriter.open(tempFile, false));
    
    bool result = handler.save(zipWriter, context);
    EXPECT_TRUE(result) << "Failed to save worksheet XML: " << handler.lastError();
    zipWriter.close();
    
    // 用Reader读取验证
    TinaXlsx::TXZipArchiveReader zipReader;
    EXPECT_TRUE(zipReader.open(tempFile));
    auto xmlData = zipReader.read(std::string(handler.partName()));
    EXPECT_FALSE(xmlData.empty()) << "Worksheet XML data is empty";
    zipReader.close();
    std::remove(tempFile.c_str());
    
    // 转换为字符串并检查内容
    std::string xmlContent(xmlData.begin(), xmlData.end());
    EXPECT_TRUE(xmlContent.find("worksheet") != std::string::npos) << "XML should contain worksheet element";
    EXPECT_TRUE(xmlContent.find("sheetData") != std::string::npos) << "XML should contain sheetData element";
    EXPECT_TRUE(xmlContent.find("Hello") != std::string::npos) << "XML should contain test data";
}

// 测试工作簿XML处理器
TEST(WorkbookXmlHandlerTest, GenerateWorkbookXml) {
    auto workbook = std::make_unique<TinaXlsx::TXWorkbook>();
    auto sheet1 = std::make_unique<TinaXlsx::TXSheet>("Sheet1", workbook.get());
    auto sheet2 = std::make_unique<TinaXlsx::TXSheet>("Sheet2", workbook.get());
    
    workbook->addSheet(std::move(sheet1));
    workbook->addSheet(std::move(sheet2));
    
    TinaXlsx::TXWorkbookXmlHandler handler;
    TinaXlsx::TXWorkbookContext context{
        workbook->getSheets(),
        workbook->getStyleManager(),
        workbook->getComponentManager()
    };
    
    TinaXlsx::TXZipArchiveWriter zipWriter;
    std::string tempFile = "test_workbook.xlsx";
    EXPECT_TRUE(zipWriter.open(tempFile, false));
    
    bool result = handler.save(zipWriter, context);
    EXPECT_TRUE(result) << "Failed to save workbook XML: " << handler.lastError();
    zipWriter.close();
    
    TinaXlsx::TXZipArchiveReader zipReader;
    EXPECT_TRUE(zipReader.open(tempFile));
    auto xmlData = zipReader.read(std::string(handler.partName()));
    EXPECT_FALSE(xmlData.empty()) << "Workbook XML data is empty";
    zipReader.close();
    std::remove(tempFile.c_str());
    
    std::string xmlContent(xmlData.begin(), xmlData.end());
    EXPECT_TRUE(xmlContent.find("workbook") != std::string::npos);
    EXPECT_TRUE(xmlContent.find("sheets") != std::string::npos);
    EXPECT_TRUE(xmlContent.find("Sheet1") != std::string::npos);
    EXPECT_TRUE(xmlContent.find("Sheet2") != std::string::npos);
}// 测试样式XML处理器
TEST(StylesXmlHandlerTest, GenerateStylesXml) {
    auto workbook = std::make_unique<TinaXlsx::TXWorkbook>();
    
    TinaXlsx::StylesXmlHandler handler;
    TinaXlsx::TXWorkbookContext context{
        workbook->getSheets(),
        workbook->getStyleManager(),
        workbook->getComponentManager()
    };
    
    TinaXlsx::TXZipArchiveWriter zipWriter;
    std::string tempFile = "test_styles.xlsx";
    EXPECT_TRUE(zipWriter.open(tempFile, false));
    
    bool result = handler.save(zipWriter, context);
    EXPECT_TRUE(result) << "Failed to save styles XML: " << handler.lastError();
    zipWriter.close();
    
    TinaXlsx::TXZipArchiveReader zipReader;
    EXPECT_TRUE(zipReader.open(tempFile));
    auto xmlData = zipReader.read(std::string(handler.partName()));
    EXPECT_FALSE(xmlData.empty()) << "Styles XML data is empty";
    zipReader.close();
    std::remove(tempFile.c_str());
    
    std::string xmlContent(xmlData.begin(), xmlData.end());
    EXPECT_TRUE(xmlContent.find("styleSheet") != std::string::npos);
}

// 测试文档属性XML处理器
TEST(DocumentPropertiesXmlHandlerTest, GenerateDocumentPropertiesXml) {
    auto workbook = std::make_unique<TinaXlsx::TXWorkbook>();
    
    TinaXlsx::TXDocumentPropertiesXmlHandler handler;
    TinaXlsx::TXWorkbookContext context{
        workbook->getSheets(),
        workbook->getStyleManager(),
        workbook->getComponentManager()
    };
    
    TinaXlsx::TXZipArchiveWriter zipWriter;
    std::string tempFile = "test_docprops.xlsx";
    EXPECT_TRUE(zipWriter.open(tempFile, false));
    
    bool result = handler.save(zipWriter, context);
    EXPECT_TRUE(result) << "Failed to save document properties XML: " << handler.lastError();
    zipWriter.close();
    
    TinaXlsx::TXZipArchiveReader zipReader;
    EXPECT_TRUE(zipReader.open(tempFile));
    
    // 检查是否创建了核心属性文件
    auto coreData = zipReader.read("docProps/core.xml");
    EXPECT_FALSE(coreData.empty()) << "Core properties XML is empty";
    
    // 检查是否创建了应用程序属性文件
    auto appData = zipReader.read("docProps/app.xml");
    EXPECT_FALSE(appData.empty()) << "App properties XML is empty";
    
    zipReader.close();
    std::remove(tempFile.c_str());
}
