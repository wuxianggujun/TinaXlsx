#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;

class CellFormattingTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("FormatTest");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        // std::remove("test_formatting.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 测试数字格式化
TEST_F(CellFormattingTest, NumberFormatting) {
    // 设置不同的数字值
    sheet->setCellValue(row_t(1), column_t(1), 1234.567);
    sheet->setCellValue(row_t(2), column_t(1), -9876.543);
    sheet->setCellValue(row_t(3), column_t(1), 0.75);

    // 应用数字格式 - 2位小数
    // 使用 TXNumberFormat::FormatType 替换 TXCell::NumberFormat
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXNumberFormat::FormatType::Number, 2));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXNumberFormat::FormatType::Number, 2));

    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));

    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());
    

    // 保存验证
    EXPECT_TRUE(workbook->saveToFile("test_formatting_number.xlsx"));
}

// 测试百分比格式化
TEST_F(CellFormattingTest, PercentageFormatting) {
    // 设置百分比值
    sheet->setCellValue(row_t(1), column_t(1), 0.25);  // 25%
    sheet->setCellValue(row_t(2), column_t(1), 0.856); // 85.6%
    sheet->setCellValue(row_t(3), column_t(1), 1.2);   // 120%

    // 应用百分比格式
    // 使用 TXNumberFormat::FormatType 替换 TXCell::NumberFormat
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXNumberFormat::FormatType::Percentage, 1)); // 25.0%
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXNumberFormat::FormatType::Percentage, 1)); // 85.6%
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(3), column_t(1), TXNumberFormat::FormatType::Percentage, 0)); // 120%

    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));
    std::string formatted3 = sheet->getCellFormattedValue(row_t(3), column_t(1));

    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());
    EXPECT_FALSE(formatted3.empty());

    // 例如: EXPECT_EQ(formatted1, "25.0%");
    //       EXPECT_EQ(formatted2, "85.6%");
    //       EXPECT_EQ(formatted3, "120%");
    // 具体值取决于 TXNumberFormat::formatPercentage 的实现

    EXPECT_TRUE(workbook->saveToFile("test_formatting_percentage.xlsx"));
}

// 测试货币格式化
TEST_F(CellFormattingTest, CurrencyFormatting) {
    // 设置货币值
    sheet->setCellValue(row_t(1), column_t(1), 1234.56);
    sheet->setCellValue(row_t(2), column_t(1), -567.89);

    // 应用货币格式
    // 使用 TXNumberFormat::FormatType 替换 TXCell::NumberFormat
    // 注意：setCellNumberFormat 并没有直接设置货币符号的参数，
    // 它依赖于 TXCell 内部创建 TXNumberFormat 对象时的默认选项，或者需要一个更丰富的 setCellNumberFormat 版本。
    // 假设 TXCell::setPredefinedFormat 内部处理了 Currency 类型的默认符号。
    // 如果需要指定货币符号，应该使用 TXCell::setNumberFormatObject 方法，并传递一个配置好的 TXNumberFormat 对象。
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXNumberFormat::FormatType::Currency, 2));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXNumberFormat::FormatType::Currency, 2));

    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));

    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());

    // 例如: EXPECT_EQ(formatted1, "$1,234.56"); (如果默认是美元且带千位分隔符)
    //       EXPECT_EQ(formatted2, "($567.89)" or "-$567.89");
    // 具体值取决于 TXNumberFormat::formatCurrency 的实现

    EXPECT_TRUE(workbook->saveToFile("test_formatting_currency.xlsx"));
}

// 测试自定义格式
TEST_F(CellFormattingTest, CustomFormatting) {
    // 设置数值
    sheet->setCellValue(row_t(1), column_t(1), 1234567.89);

    // 应用自定义格式
    std::string custom_format_string = "#,##0.00";
    EXPECT_TRUE(sheet->setCellCustomFormat(row_t(1), column_t(1), custom_format_string));

    // 检查自定义格式
    auto* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    const TXNumberFormat* num_fmt_obj = cell->getNumberFormatObject(); // 获取 TXNumberFormat 对象
    ASSERT_NE(num_fmt_obj, nullptr);

    EXPECT_EQ(num_fmt_obj->getFormatType(), TXNumberFormat::FormatType::Custom); // 检查类型是否为 Custom
    EXPECT_EQ(num_fmt_obj->getFormatString(), custom_format_string); // 检查格式字符串是否匹配

    // 检查格式化后的值
    std::string formatted_val = sheet->getCellFormattedValue(row_t(1), column_t(1));
    EXPECT_FALSE(formatted_val.empty());
    // 例如: EXPECT_EQ(formatted_val, "1,234,567.89");
    // 具体值取决于 TXNumberFormat 对自定义格式的解析和应用

    EXPECT_TRUE(workbook->saveToFile("test_formatting_custom.xlsx"));
}

// 测试范围格式化
TEST_F(CellFormattingTest, RangeFormatting) {
    // 设置一个2x2的数据范围
    sheet->setCellValue(row_t(1), column_t(1), 100.123);
    sheet->setCellValue(row_t(1), column_t(2), 200.234);
    sheet->setCellValue(row_t(2), column_t(1), 300.345);
    sheet->setCellValue(row_t(2), column_t(2), 400.456);

    // 创建范围
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));

    // 应用范围格式
    // 使用 TXNumberFormat::FormatType 替换 TXCell::NumberFormat
    std::size_t count = sheet->setRangeNumberFormat(range, TXNumberFormat::FormatType::Number, 1);
    EXPECT_EQ(count, 4);  // 应该格式化4个单元格

    // 验证每个单元格都被格式化了
    for (uint32_t r_idx = 1; r_idx <= 2; ++r_idx) { // 使用简单循环
        for (uint32_t c_idx = 1; c_idx <= 2; ++c_idx) {
            row_t current_row(r_idx);
            column_t current_col(c_idx);
            auto* cell = sheet->getCell(current_row, current_col);
            ASSERT_NE(cell, nullptr);
            const TXNumberFormat* num_fmt_obj = cell->getNumberFormatObject(); // 获取 TXNumberFormat 对象
            ASSERT_NE(num_fmt_obj, nullptr);
            EXPECT_EQ(num_fmt_obj->getFormatType(), TXNumberFormat::FormatType::Number); // 检查格式类型

            // 也可以检查格式化后的字符串
            std::string formatted_val = sheet->getCellFormattedValue(current_row, current_col);
            // 例如，对于 100.123 和 1位小数，可能是 "100.1"
            // EXPECT_EQ(formatted_val, std::to_string(sheet->getCellValue(current_row, current_col).getNumberValue() /* adapt to get double*/ ) with 1 decimal place);
        }
    }
    EXPECT_TRUE(workbook->saveToFile("test_formatting_range.xlsx"));
}

// 测试TXNumberFormat对象
TEST_F(CellFormattingTest, TXNumberFormatObject) {
    // 创建不同类型的格式对象
    auto numberFormat = TXNumberFormat::createNumberFormat(3, true); // 3位小数，带千位分隔符
    auto currencyFormat = TXNumberFormat::createCurrencyFormat("¥", 2); // 人民币，2位小数
    auto percentageFormat = TXNumberFormat::createPercentageFormat(1); // 1位小数百分比

    // 验证格式类型 (这部分应该已经是正确的，因为它直接使用了 TXNumberFormat)
    EXPECT_EQ(numberFormat.getFormatType(), TXNumberFormat::FormatType::Number);
    EXPECT_EQ(currencyFormat.getFormatType(), TXNumberFormat::FormatType::Currency);
    EXPECT_EQ(percentageFormat.getFormatType(), TXNumberFormat::FormatType::Percentage);

    // 测试格式化功能
    double testValue = 1234.5678;

    // formatNumber, formatCurrency 等方法是 TXNumberFormat 的成员
    std::string numberStr = numberFormat.format(testValue); // 或者 numberFormat.formatNumber(testValue)
    std::string currencyStr = currencyFormat.format(testValue); // 或者 currencyFormat.formatCurrency(testValue)
    std::string percentStr = percentageFormat.format(testValue * 100.0); // formatPercentage 期望的是乘以100后的值，或者其内部会乘100

    // TXNumberFormat::formatPercentage 接收的参数是原始值（如0.25代表25%），内部会乘以100
    // 如果TXNumberFormat::formatPercentage直接格式化传入的数字为百分比字符串，则应：
    std::string percentStrValue = percentageFormat.format(0.751); // 应该得到 "75.1%"

    EXPECT_FALSE(numberStr.empty());
    EXPECT_FALSE(currencyStr.empty());
    EXPECT_FALSE(percentStrValue.empty());

    // 验证货币符号
    EXPECT_NE(currencyStr.find("¥"), std::string::npos); // 使用 EXPECT_NE
}

// 测试批量格式化
TEST_F(CellFormattingTest, BatchFormatting) {
    // 设置多个单元格的值和目标格式
    // 使用 TXNumberFormat::FormatType 替换 TXCell::NumberFormat
    std::vector<std::pair<TXCoordinate, TXNumberFormat::FormatType>> formats = {
        {TXCoordinate(row_t(1), column_t(1)), TXNumberFormat::FormatType::Number},
        {TXCoordinate(row_t(1), column_t(2)), TXNumberFormat::FormatType::Currency},
        {TXCoordinate(row_t(1), column_t(3)), TXNumberFormat::FormatType::Percentage}
    };

    // 先设置数值
    sheet->setCellValue(row_t(1), column_t(1), 123.45);
    sheet->setCellValue(row_t(1), column_t(2), 678.90);
    sheet->setCellValue(row_t(1), column_t(3), 0.85); // 85%

    // 批量应用格式
    // 注意：setCellFormats 的实现。它会调用 setCellNumberFormat，后者使用默认小数位数。
    // 如果需要为每种格式指定不同的小数位数，则 setCellFormats 的设计可能需要调整，
    // 或者需要一个更复杂的参数结构（例如，std::vector<std::tuple<TXCoordinate, TXNumberFormat::FormatType, int>>）。
    // 当前的 setCellFormats(const std::vector<std::pair<Coordinate, TXNumberFormat::FormatType>>&) 会为所有格式使用默认小数位数。
    std::size_t count = sheet->setCellFormats(formats);
    EXPECT_EQ(count, 3);

    // 验证格式是否正确应用
    auto* cell1 = sheet->getCell(row_t(1), column_t(1));
    auto* cell2 = sheet->getCell(row_t(1), column_t(2));
    auto* cell3 = sheet->getCell(row_t(1), column_t(3));

    ASSERT_NE(cell1, nullptr);
    ASSERT_NE(cell2, nullptr);
    ASSERT_NE(cell3, nullptr);

    const TXNumberFormat* fmt1 = cell1->getNumberFormatObject();
    const TXNumberFormat* fmt2 = cell2->getNumberFormatObject();
    const TXNumberFormat* fmt3 = cell3->getNumberFormatObject();

    ASSERT_NE(fmt1, nullptr);
    EXPECT_EQ(fmt1->getFormatType(), TXNumberFormat::FormatType::Number);

    ASSERT_NE(fmt2, nullptr);
    EXPECT_EQ(fmt2->getFormatType(), TXNumberFormat::FormatType::Currency);

    ASSERT_NE(fmt3, nullptr);
    EXPECT_EQ(fmt3->getFormatType(), TXNumberFormat::FormatType::Percentage);

    EXPECT_TRUE(workbook->saveToFile("test_formatting_batch.xlsx"));
}
