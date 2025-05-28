//
// 单元格格式设置功能完整测试程序
//

#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXConditionalFormat.hpp"
#include "TinaXlsx/TXStyleTemplate.hpp"
#include "TinaXlsx/TXBatchFormat.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <memory>

using namespace TinaXlsx;

class CellFormatFeaturesTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("TestSheet");
    }
    
    void TearDown() override {
        workbook.reset();
    }
    
    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 基本格式设置功能测试
TEST_F(CellFormatFeaturesTest, BasicFontFormatting) {
    TXCellStyle style;
    
    // 字体名称和大小
    style.setFont("Arial", 14);
    EXPECT_EQ(style.getFont().name, "Arial");
    EXPECT_EQ(style.getFont().size, 14);
    
    // 字体颜色
    style.setFontColor(ColorConstants::RED);
    EXPECT_EQ(style.getFont().color.getValue(), ColorConstants::RED);
    
    // 字体样式
    style.setFontStyle(FontStyle::Bold);
    EXPECT_TRUE(style.getFont().isBold());
}

TEST_F(CellFormatFeaturesTest, AlignmentFormatting) {
    TXCellStyle style;
    
    // 水平对齐
    style.setHorizontalAlignment(HorizontalAlignment::Center);
    EXPECT_EQ(style.getAlignment().horizontal, HorizontalAlignment::Center);
    
    // 垂直对齐
    style.setVerticalAlignment(VerticalAlignment::Middle);
    EXPECT_EQ(style.getAlignment().vertical, VerticalAlignment::Middle);
    
    // 文本换行
    style.getAlignment().setWrapText(true);
    EXPECT_TRUE(style.getAlignment().wrapText);
    
    // 缩小以适应
    style.getAlignment().setShrinkToFit(true);
    EXPECT_TRUE(style.getAlignment().shrinkToFit);
    
    // 文本旋转
    style.getAlignment().setTextRotation(45);
    EXPECT_EQ(style.getAlignment().textRotation, 45);
    
    // 缩进
    style.getAlignment().setIndent(2);
    EXPECT_EQ(style.getAlignment().indent, 2);
}

TEST_F(CellFormatFeaturesTest, BorderFormatting) {
    TXCellStyle style;
    
    // 设置所有边框
    style.setAllBorders(BorderStyle::Thin, ColorConstants::BLACK);
    EXPECT_EQ(style.getBorder().leftStyle, BorderStyle::Thin);
    EXPECT_EQ(style.getBorder().rightStyle, BorderStyle::Thin);
    EXPECT_EQ(style.getBorder().topStyle, BorderStyle::Thin);
    EXPECT_EQ(style.getBorder().bottomStyle, BorderStyle::Thin);
    
    // 设置单独边框
    style.getBorder().setLeftBorder(BorderStyle::Thick, ColorConstants::RED);
    EXPECT_EQ(style.getBorder().leftStyle, BorderStyle::Thick);
    EXPECT_EQ(style.getBorder().leftColor.getValue(), ColorConstants::RED);
    
    // 设置对角线边框
    style.getBorder().setDiagonalBorder(BorderStyle::Dashed, ColorConstants::BLUE, true, false);
    EXPECT_EQ(style.getBorder().diagonalStyle, BorderStyle::Dashed);
    EXPECT_TRUE(style.getBorder().diagonalUp);
    EXPECT_FALSE(style.getBorder().diagonalDown);
}

TEST_F(CellFormatFeaturesTest, FillFormatting) {
    TXCellStyle style;
    
    // 背景颜色
    style.setBackgroundColor(ColorConstants::YELLOW);
    EXPECT_EQ(style.getFill().pattern, FillPattern::Solid);
    EXPECT_EQ(style.getFill().foregroundColor.getValue(), ColorConstants::YELLOW);
    
    // 渐变填充
    style.getFill().setPattern(FillPattern::Gray50);
    EXPECT_EQ(style.getFill().pattern, FillPattern::Gray50);
}

// ==================== 2. 条件格式测试 ====================

TEST_F(CellFormatFeaturesTest, ConditionalFormatCellValue) {
    // 创建条件格式管理器
    TXConditionalFormatManager manager;
    
    // 创建单元格值规则
    TXCellStyle highlightStyle;
    highlightStyle.setBackgroundColor(ColorConstants::RED);
    
    auto rule = TXConditionalFormatManager::createCellValueRule(
        ConditionalOperator::Greater, 
        static_cast<double>(100.0), 
        highlightStyle
    );
    
    EXPECT_EQ(rule->getType(), ConditionalFormatType::CellValue);
    
    manager.addRule(std::move(rule));
    EXPECT_EQ(manager.getRuleCount(), 1);
}

TEST_F(CellFormatFeaturesTest, ConditionalFormatColorScale) {
    TXConditionalFormatManager manager;
    
    // 创建二色阶规则
    auto colorScaleRule = TXConditionalFormatManager::createTwoColorScale(
        ColorConstants::RED, 
        ColorConstants::GREEN
    );
    
    EXPECT_EQ(colorScaleRule->getType(), ConditionalFormatType::ColorScale);
    
    manager.addRule(std::move(colorScaleRule));
    EXPECT_EQ(manager.getRuleCount(), 1);
}

TEST_F(CellFormatFeaturesTest, ConditionalFormatDataBar) {
    TXConditionalFormatManager manager;
    
    // 创建数据条规则
    auto dataBarRule = TXConditionalFormatManager::createDataBarRule(
        ColorConstants::BLUE, 
        true
    );
    
    EXPECT_EQ(dataBarRule->getType(), ConditionalFormatType::DataBar);
    
    manager.addRule(std::move(dataBarRule));
    EXPECT_EQ(manager.getRuleCount(), 1);
}

TEST_F(CellFormatFeaturesTest, ConditionalFormatIconSet) {
    TXConditionalFormatManager manager;
    
    // 创建图标集规则
    auto iconSetRule = TXConditionalFormatManager::createIconSetRule(
        IconSetType::ThreeArrows, 
        true
    );
    
    EXPECT_EQ(iconSetRule->getType(), ConditionalFormatType::IconSet);
    
    manager.addRule(std::move(iconSetRule));
    EXPECT_EQ(manager.getRuleCount(), 1);
}

// ==================== 3. 样式模板测试 ====================

TEST_F(CellFormatFeaturesTest, StyleTemplateBasic) {
    StyleTemplateInfo info("test_template", "Test Template", StyleTemplateCategory::Custom);
    TXStyleTemplate styleTemplate(info);
    
    EXPECT_EQ(styleTemplate.getId(), "test_template");
    EXPECT_EQ(styleTemplate.getName(), "Test Template");
    
    // 设置基础样式
    TXCellStyle baseStyle;
    baseStyle.setFont("Arial", 12);
    styleTemplate.setBaseStyle(baseStyle);
    
    EXPECT_EQ(styleTemplate.getBaseStyle().getFont().name, "Arial");
    EXPECT_EQ(styleTemplate.getBaseStyle().getFont().size, 12);
}

TEST_F(CellFormatFeaturesTest, StyleTemplateNamedStyles) {
    TXStyleTemplate styleTemplate;
    
    // 添加命名样式
    TXCellStyle headerStyle;
    headerStyle.setFont("Arial", 14).setFontStyle(FontStyle::Bold);
    styleTemplate.addNamedStyle("header", headerStyle);
    
    TXCellStyle dataStyle;
    dataStyle.setFont("Arial", 11);
    styleTemplate.addNamedStyle("data", dataStyle);
    
    // 验证命名样式
    auto retrievedHeaderStyle = styleTemplate.getNamedStyle("header");
    EXPECT_NE(retrievedHeaderStyle, nullptr);
    EXPECT_EQ(retrievedHeaderStyle->getFont().size, 14);
    EXPECT_TRUE(retrievedHeaderStyle->getFont().isBold());
    
    auto retrievedDataStyle = styleTemplate.getNamedStyle("data");
    EXPECT_NE(retrievedDataStyle, nullptr);
    EXPECT_EQ(retrievedDataStyle->getFont().size, 11);
    
    // 获取所有命名样式名称
    auto styleNames = styleTemplate.getNamedStyleNames();
    EXPECT_EQ(styleNames.size(), 2);
}

TEST_F(CellFormatFeaturesTest, StyleTemplateManager) {
    auto& manager = TXStyleTemplateManager::getInstance();
    
    // 创建并注册模板
    StyleTemplateInfo info("business_template", "Business Template", StyleTemplateCategory::Data);
    TXStyleTemplate businessTemplate(info);
    
    EXPECT_TRUE(manager.registerTemplate(businessTemplate));
    EXPECT_TRUE(manager.hasTemplate("business_template"));
    
    // 获取模板
    auto retrievedTemplate = manager.getTemplate("business_template");
    EXPECT_NE(retrievedTemplate, nullptr);
    EXPECT_EQ(retrievedTemplate->getName(), "Business Template");
    
    // 注销模板
    EXPECT_TRUE(manager.unregisterTemplate("business_template"));
    EXPECT_FALSE(manager.hasTemplate("business_template"));
}

// ==================== 4. 预定义样式测试 ====================

TEST_F(CellFormatFeaturesTest, PredefinedStyles) {
    // 测试预定义样式
    auto headerStyle = Styles::createHeaderStyle();
    EXPECT_EQ(headerStyle.getFont().size, 14);
    EXPECT_TRUE(headerStyle.getFont().isBold());
    EXPECT_EQ(headerStyle.getAlignment().horizontal, HorizontalAlignment::Center);
    
    auto dataStyle = Styles::createDataStyle();
    EXPECT_EQ(dataStyle.getFont().size, 11);
    EXPECT_EQ(dataStyle.getAlignment().horizontal, HorizontalAlignment::Left);
    
    auto numberStyle = Styles::createNumberStyle();
    EXPECT_EQ(numberStyle.getAlignment().horizontal, HorizontalAlignment::Right);
    
    auto highlightStyle = Styles::createHighlightStyle(ColorConstants::YELLOW);
    EXPECT_TRUE(highlightStyle.getFont().isBold());
    
    auto tableStyle = Styles::createTableStyle();
    EXPECT_EQ(tableStyle.getBorder().leftStyle, BorderStyle::Thin);
}

// ==================== 5. 综合测试 ====================

TEST_F(CellFormatFeaturesTest, ComprehensiveFormatTest) {
    // 创建一个复杂的格式样式
    TXCellStyle complexStyle;
    
    // 字体设置
    complexStyle.setFont("Times New Roman", 16)
                .setFontColor(ColorConstants::DARK_BLUE)
                .setFontStyle(FontStyle::BoldItalic);
    
    // 对齐设置
    complexStyle.setHorizontalAlignment(HorizontalAlignment::Center)
                .setVerticalAlignment(VerticalAlignment::Middle);
    complexStyle.getAlignment().setWrapText(true);
    complexStyle.getAlignment().setTextRotation(15);
    
    // 边框设置
    complexStyle.setAllBorders(BorderStyle::Double, ColorConstants::BLACK);
    complexStyle.getBorder().setLeftBorder(BorderStyle::Thick, ColorConstants::RED);
    
    // 填充设置
    complexStyle.setBackgroundColor(ColorConstants::LIGHT_GRAY);
    
    // 验证所有设置
    EXPECT_EQ(complexStyle.getFont().name, "Times New Roman");
    EXPECT_EQ(complexStyle.getFont().size, 16);
    EXPECT_EQ(complexStyle.getFont().color.getValue(), ColorConstants::DARK_BLUE);
    EXPECT_TRUE(complexStyle.getFont().isBold());
    EXPECT_TRUE(complexStyle.getFont().isItalic());
    
    EXPECT_EQ(complexStyle.getAlignment().horizontal, HorizontalAlignment::Center);
    EXPECT_EQ(complexStyle.getAlignment().vertical, VerticalAlignment::Middle);
    EXPECT_TRUE(complexStyle.getAlignment().wrapText);
    EXPECT_EQ(complexStyle.getAlignment().textRotation, 15);
    
    EXPECT_EQ(complexStyle.getBorder().rightStyle, BorderStyle::Double);
    EXPECT_EQ(complexStyle.getBorder().leftStyle, BorderStyle::Thick);
    EXPECT_EQ(complexStyle.getBorder().leftColor.getValue(), ColorConstants::RED);
    
    EXPECT_EQ(complexStyle.getFill().pattern, FillPattern::Solid);
    EXPECT_EQ(complexStyle.getFill().foregroundColor.getValue(), ColorConstants::LIGHT_GRAY);
} // 测试结束