#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <cstdio>
#include <fstream>

using namespace TinaXlsx;

class CellStyleTest : public TestWithFileGeneration<CellStyleTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<CellStyleTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<CellStyleTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
};

// 综合样式测试 - 将所有样式测试合并到一个文件中
TEST_F(CellStyleTest, ComprehensiveStyleTest) {
    // ==================== Sheet 1: 字体样式测试 ====================
    auto* fontSheet = workbook->addSheet("Font Styles");
    ASSERT_NE(fontSheet, nullptr);
    
    // 测试1: 基本字体样式
    TXFont font1;
    font1.setName("Arial");
    font1.setSize(12);
    font1.setBold(true);
    font1.setItalic(true);
    font1.setColor(TXColor(255, 0, 0)); // 红色
    
    TXCellStyle style1;
    style1.setFont(font1);
    
    fontSheet->setCellValue(row_t(1), column_t(1), std::string("Bold Red Arial 12pt"));
    EXPECT_TRUE(fontSheet->setCellStyle(row_t(1), column_t(1), style1));
    
    // 测试2: 不同字体大小
    TXFont font2;
    font2.setName("Times New Roman");
    font2.setSize(16);
    font2.setUnderline(UnderlineStyle::Single);
    font2.setColor(TXColor(0, 0, 255)); // 蓝色
    
    TXCellStyle style2;
    style2.setFont(font2);
    
    fontSheet->setCellValue(row_t(2), column_t(1), std::string("Underlined Blue Times 16pt"));
    EXPECT_TRUE(fontSheet->setCellStyle(row_t(2), column_t(1), style2));
    
    // 测试3: 删除线字体
    TXFont font3;
    font3.setName("Calibri");
    font3.setSize(10);
    font3.setStrikethrough(true);
    font3.setColor(TXColor(128, 128, 128)); // 灰色
    
    TXCellStyle style3;
    style3.setFont(font3);
    
    fontSheet->setCellValue(row_t(3), column_t(1), std::string("Strikethrough Gray Calibri 10pt"));
    EXPECT_TRUE(fontSheet->setCellStyle(row_t(3), column_t(1), style3));

    // ==================== Sheet 2: 边框样式测试 ====================
    auto* borderSheet = workbook->addSheet("Border Styles");
    ASSERT_NE(borderSheet, nullptr);
    
    // 测试1: 四边不同边框
    TXBorder border1;
    border1.setLeftBorder(BorderStyle::Thin, TXColor(0, 0, 255));   // 蓝色细线
    border1.setRightBorder(BorderStyle::Thick, TXColor(0, 255, 0)); // 绿色粗线
    border1.setTopBorder(BorderStyle::Double, TXColor(255, 0, 0));  // 红色双线
    border1.setBottomBorder(BorderStyle::Dashed, TXColor(0, 0, 0)); // 黑色虚线
    
    TXCellStyle borderStyle1;
    borderStyle1.setBorder(border1);
    
    borderSheet->setCellValue(row_t(1), column_t(1), std::string("Mixed Borders"));
    EXPECT_TRUE(borderSheet->setCellStyle(row_t(1), column_t(1), borderStyle1));
    
    // 测试2: 统一边框
    TXBorder border2;
    border2.setAllBorders(BorderStyle::Medium, TXColor(128, 0, 128)); // 紫色中等边框
    
    TXCellStyle borderStyle2;
    borderStyle2.setBorder(border2);
    
    borderSheet->setCellValue(row_t(1), column_t(3), std::string("Uniform Purple Border"));
    EXPECT_TRUE(borderSheet->setCellStyle(row_t(1), column_t(3), borderStyle2));
    
    // 测试3: 对角线边框 - 移除不存在的对角线方法
    TXBorder border3;
    border3.setAllBorders(BorderStyle::Thin, TXColor(255, 165, 0)); // 橙色边框
    
    TXCellStyle borderStyle3;
    borderStyle3.setBorder(border3);
    
    borderSheet->setCellValue(row_t(1), column_t(5), std::string("Orange Borders"));
    EXPECT_TRUE(borderSheet->setCellStyle(row_t(1), column_t(5), borderStyle3));

    // ==================== Sheet 3: 填充样式测试 ====================
    auto* fillSheet = workbook->addSheet("Fill Styles");
    ASSERT_NE(fillSheet, nullptr);
    
    // 测试1: 纯色填充
    TXFill fill1;
    fill1.setPattern(FillPattern::Solid);
    fill1.setForegroundColor(TXColor(255, 255, 0)); // 黄色背景
    
    TXCellStyle fillStyle1;
    fillStyle1.setFill(fill1);
    
    fillSheet->setCellValue(row_t(1), column_t(1), std::string("Yellow Background"));
    EXPECT_TRUE(fillSheet->setCellStyle(row_t(1), column_t(1), fillStyle1));
    
    // 测试2: 渐变填充
    TXFill fill2;
    fill2.setPattern(FillPattern::Gray125);
    fill2.setForegroundColor(TXColor(0, 255, 0)); // 绿色
    fill2.setBackgroundColor(TXColor(255, 255, 255)); // 白色
    
    TXCellStyle fillStyle2;
    fillStyle2.setFill(fill2);
    
    fillSheet->setCellValue(row_t(1), column_t(3), std::string("Green Pattern"));
    EXPECT_TRUE(fillSheet->setCellStyle(row_t(1), column_t(3), fillStyle2));
    
    // 测试3: 点状填充 - 使用存在的枚举值
    TXFill fill3;
    fill3.setPattern(FillPattern::Gray0625);
    fill3.setForegroundColor(TXColor(255, 0, 255)); // 洋红色
    
    TXCellStyle fillStyle3;
    fillStyle3.setFill(fill3);
    
    fillSheet->setCellValue(row_t(1), column_t(5), std::string("Magenta Pattern"));
    EXPECT_TRUE(fillSheet->setCellStyle(row_t(1), column_t(5), fillStyle3));

    // ==================== Sheet 4: 对齐样式测试 ====================
    auto* alignSheet = workbook->addSheet("Alignment Styles");
    ASSERT_NE(alignSheet, nullptr);
    
    // 移除不存在的setColumnWidth方法调用
    
    // 测试1: 左对齐
    TXAlignment align1;
    align1.setHorizontal(HorizontalAlignment::Left);
    align1.setVertical(VerticalAlignment::Top);
    
    TXCellStyle alignStyle1;
    alignStyle1.setAlignment(align1);
    
    alignSheet->setCellValue(row_t(1), column_t(1), std::string("Left-Top"));
    EXPECT_TRUE(alignSheet->setCellStyle(row_t(1), column_t(1), alignStyle1));
    
    // 测试2: 居中对齐
    TXAlignment align2;
    align2.setHorizontal(HorizontalAlignment::Center);
    align2.setVertical(VerticalAlignment::Middle);
    align2.setWrapText(true);
    
    TXCellStyle alignStyle2;
    alignStyle2.setAlignment(align2);
    
    alignSheet->setCellValue(row_t(1), column_t(2), std::string("Center-Middle with Wrap"));
    EXPECT_TRUE(alignSheet->setCellStyle(row_t(1), column_t(2), alignStyle2));
    
    // 测试3: 右对齐
    TXAlignment align3;
    align3.setHorizontal(HorizontalAlignment::Right);
    align3.setVertical(VerticalAlignment::Bottom);
    align3.setIndent(2);
    
    TXCellStyle alignStyle3;
    alignStyle3.setAlignment(align3);
    
    alignSheet->setCellValue(row_t(1), column_t(3), std::string("Right-Bottom Indented"));
    EXPECT_TRUE(alignSheet->setCellStyle(row_t(1), column_t(3), alignStyle3));

    // ==================== Sheet 5: 组合样式测试 ====================
    auto* comboSheet = workbook->addSheet("Combined Styles");
    ASSERT_NE(comboSheet, nullptr);
    
    // 测试1: 表头样式
    TXFont headerFont;
    headerFont.setName("Arial");
    headerFont.setSize(14);
    headerFont.setBold(true);
    headerFont.setColor(TXColor(255, 255, 255)); // 白色文字
    
    TXFill headerFill;
    headerFill.setPattern(FillPattern::Solid);
    headerFill.setForegroundColor(TXColor(0, 0, 128)); // 深蓝色背景
    
    TXBorder headerBorder;
    headerBorder.setAllBorders(BorderStyle::Medium, TXColor(0, 0, 0)); // 黑色边框
    
    TXAlignment headerAlign;
    headerAlign.setHorizontal(HorizontalAlignment::Center);
    headerAlign.setVertical(VerticalAlignment::Middle);
    
    TXCellStyle headerStyle;
    headerStyle.setFont(headerFont);
    headerStyle.setFill(headerFill);
    headerStyle.setBorder(headerBorder);
    headerStyle.setAlignment(headerAlign);
    
    // 应用表头样式到多个单元格
    std::vector<std::string> headers = {"Name", "Age", "Department", "Salary"};
    for (size_t i = 0; i < headers.size(); ++i) {
        comboSheet->setCellValue(row_t(1), column_t(static_cast<column_t::index_t>(i + 1)), headers[i]);
        EXPECT_TRUE(comboSheet->setCellStyle(row_t(1), column_t(static_cast<column_t::index_t>(i + 1)), headerStyle));
    }
    
    // 测试2: 数据行样式（交替颜色）
    TXFill evenRowFill;
    evenRowFill.setPattern(FillPattern::Solid);
    evenRowFill.setForegroundColor(TXColor(240, 240, 240)); // 浅灰色
    
    TXFill oddRowFill;
    oddRowFill.setPattern(FillPattern::Solid);
    oddRowFill.setForegroundColor(TXColor(255, 255, 255)); // 白色
    
    TXBorder dataBorder;
    dataBorder.setAllBorders(BorderStyle::Thin, TXColor(128, 128, 128)); // 灰色细边框
    
    TXCellStyle evenRowStyle;
    evenRowStyle.setFill(evenRowFill);
    evenRowStyle.setBorder(dataBorder);
    
    TXCellStyle oddRowStyle;
    oddRowStyle.setFill(oddRowFill);
    oddRowStyle.setBorder(dataBorder);
    
    // 添加示例数据
    std::vector<std::vector<std::string>> data = {
        {"John Doe", "30", "Engineering", "$75000"},
        {"Jane Smith", "28", "Marketing", "$65000"},
        {"Bob Johnson", "35", "Sales", "$70000"},
        {"Alice Brown", "32", "HR", "$60000"}
    };
    
    for (size_t row = 0; row < data.size(); ++row) {
        bool isEvenRow = (row % 2 == 0);
        TXCellStyle& rowStyle = isEvenRow ? evenRowStyle : oddRowStyle;
        
        for (size_t col = 0; col < data[row].size(); ++col) {
            comboSheet->setCellValue(row_t(static_cast<row_t::index_t>(row + 2)), 
                                   column_t(static_cast<column_t::index_t>(col + 1)), data[row][col]);
            EXPECT_TRUE(comboSheet->setCellStyle(row_t(static_cast<row_t::index_t>(row + 2)), 
                                                column_t(static_cast<column_t::index_t>(col + 1)), rowStyle));
        }
    }

    // ==================== Sheet 6: 范围样式测试 ====================
    auto* rangeSheet = workbook->addSheet("Range Styles");
    ASSERT_NE(rangeSheet, nullptr);
    
    // 创建一个5x5的数据表格
    for (int row = 1; row <= 5; ++row) {
        for (int col = 1; col <= 5; ++col) {
            rangeSheet->setCellValue(row_t(static_cast<row_t::index_t>(row)), 
                                   column_t(static_cast<column_t::index_t>(col)), 
                                   std::string("R") + std::to_string(row) + "C" + std::to_string(col));
        }
    }
    
    // 应用范围样式
    TXFont rangeFont;
    rangeFont.setBold(true);
    rangeFont.setSize(10);
    
    TXBorder rangeBorder;
    rangeBorder.setAllBorders(BorderStyle::Thin, TXColor(0, 0, 0));
    
    TXCellStyle rangeStyle;
    rangeStyle.setFont(rangeFont);
    rangeStyle.setBorder(rangeBorder);
    
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(5), column_t(5)));
    std::size_t styledCount = rangeSheet->setRangeStyle(range, rangeStyle);
    EXPECT_EQ(styledCount, 25); // 5x5 = 25个单元格
    
    // 验证所有单元格都有样式
    for (int row = 1; row <= 5; ++row) {
        for (int col = 1; col <= 5; ++col) {
            auto* cell = rangeSheet->getCell(row_t(static_cast<row_t::index_t>(row)), 
                                           column_t(static_cast<column_t::index_t>(col)));
            ASSERT_NE(cell, nullptr);
            EXPECT_TRUE(cell->hasStyle());
        }
    }

    // ==================== 保存综合测试文件 ====================
    // 添加测试信息到第一个工作表
    addTestInfo(fontSheet, "ComprehensiveStyleTest", "综合样式测试 - 包含字体、边框、填充、对齐、组合和范围样式");

    EXPECT_TRUE(saveWorkbook(workbook, "ComprehensiveStyleTest"));

    // 验证文件确实被创建
    std::string filePath = getFilePath("ComprehensiveStyleTest");
    std::ifstream file(filePath, std::ios::binary);
    ASSERT_TRUE(file.good());
    file.close();
    
    // 验证工作簿包含所有预期的工作表
    EXPECT_EQ(workbook->getSheetCount(), 6);
    EXPECT_EQ(workbook->getSheet("Font Styles"), fontSheet);
    EXPECT_EQ(workbook->getSheet("Border Styles"), borderSheet);
    EXPECT_EQ(workbook->getSheet("Fill Styles"), fillSheet);
    EXPECT_EQ(workbook->getSheet("Alignment Styles"), alignSheet);
    EXPECT_EQ(workbook->getSheet("Combined Styles"), comboSheet);
    EXPECT_EQ(workbook->getSheet("Range Styles"), rangeSheet);
}

// 保留一个简单的单独测试用于快速验证
TEST_F(CellStyleTest, QuickStyleTest) {
    auto* sheet = workbook->addSheet("Quick Test");
    ASSERT_NE(sheet, nullptr);
    
    // 简单的字体样式测试
    TXFont font;
    font.setName("Arial");
    font.setSize(12);
    font.setBold(true);
    font.setColor(TXColor(255, 0, 0));
    
    TXCellStyle style;
    style.setFont(font);
    
    sheet->setCellValue(row_t(1), column_t(1), std::string("Quick Test"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(1), column_t(1), style));
    
    auto* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 添加测试信息
    addTestInfo(sheet, "QuickStyleTest", "快速样式测试");

    EXPECT_TRUE(saveWorkbook(workbook, "QuickStyleTest"));
} 
