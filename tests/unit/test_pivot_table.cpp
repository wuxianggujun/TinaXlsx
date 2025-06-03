/*
 * TinaXlsx - Modern C++ Excel File Processing Library
 * 
 * Copyright (c) 2025 wuxianggujun
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#include <gtest/gtest.h>
#include "TinaXlsx/TXPivotTable.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

/**
 * @brief 透视表功能测试类
 */
class PivotTableTest : public TestWithFileGeneration<PivotTableTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<PivotTableTest>::SetUp();

        // 创建测试工作簿
        workbook = createWorkbook("PivotTableTest");
        sheet = workbook->addSheet("销售数据");

        // 准备测试数据
        setupTestData();
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<PivotTableTest>::TearDown();
    }

    /**
     * @brief 设置测试数据
     */
    void setupTestData() {
        // 设置表头
        sheet->setCellValue("A1", "产品类别");
        sheet->setCellValue("B1", "销售员");
        sheet->setCellValue("C1", "销售月份");
        sheet->setCellValue("D1", "销售额");
        sheet->setCellValue("E1", "销售数量");

        // 添加测试数据
        const std::vector<std::vector<std::string>> testData = {
            {"电子产品", "张三", "2024-01", "15000", "50"},
            {"电子产品", "李四", "2024-01", "12000", "40"},
            {"服装", "张三", "2024-01", "8000", "80"},
            {"服装", "王五", "2024-01", "6000", "60"},
            {"电子产品", "张三", "2024-02", "18000", "60"},
            {"电子产品", "李四", "2024-02", "14000", "45"},
            {"服装", "张三", "2024-02", "9000", "90"},
            {"服装", "王五", "2024-02", "7000", "70"},
            {"家具", "赵六", "2024-01", "25000", "25"},
            {"家具", "赵六", "2024-02", "30000", "30"}
        };

        for (size_t i = 0; i < testData.size(); ++i) {
            const auto& row = testData[i];
            for (size_t j = 0; j < row.size(); ++j) {
                std::string cellRef = std::string(1, 'A' + j) + std::to_string(i + 2);
                if (j >= 3) { // 数值列
                    sheet->setCellValue(cellRef, std::stod(row[j]));
                } else { // 文本列
                    sheet->setCellValue(cellRef, row[j]);
                }
            }
        }
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

/**
 * @brief 测试透视表基础创建
 */
TEST_F(PivotTableTest, BasicCreation) {
    // 创建透视表
    TXRange sourceRange("A1:E11");
    TXPivotTable pivotTable(sourceRange, "G1");

    // 验证基础属性
    EXPECT_EQ(pivotTable.getName(), "PivotTable1");
    EXPECT_EQ(pivotTable.getTargetCell(), "G1");

    // 设置名称
    pivotTable.setName("销售数据透视表");
    EXPECT_EQ(pivotTable.getName(), "销售数据透视表");
}

/**
 * @brief 测试字段添加和管理
 */
TEST_F(PivotTableTest, FieldManagement) {
    TXRange sourceRange("A1:E11");
    TXPivotTable pivotTable(sourceRange, "G1");

    // 添加行字段
    EXPECT_TRUE(pivotTable.addRowField("产品类别"));
    
    // 添加列字段
    EXPECT_TRUE(pivotTable.addColumnField("销售月份"));
    
    // 添加数据字段
    EXPECT_TRUE(pivotTable.addDataField("销售额", PivotAggregateFunction::Sum));
    EXPECT_TRUE(pivotTable.addDataField("销售数量", PivotAggregateFunction::Average));
    
    // 添加筛选字段
    EXPECT_TRUE(pivotTable.addFilterField("销售员"));

    // 验证字段获取
    auto categoryField = pivotTable.getField("产品类别");
    ASSERT_NE(categoryField, nullptr);
    EXPECT_EQ(categoryField->getName(), "产品类别");
    EXPECT_EQ(categoryField->getType(), PivotFieldType::Row);

    auto salesField = pivotTable.getField("销售额");
    ASSERT_NE(salesField, nullptr);
    EXPECT_EQ(salesField->getType(), PivotFieldType::Data);
    EXPECT_EQ(salesField->getAggregateFunction(), PivotAggregateFunction::Sum);

    // 测试重复添加字段
    EXPECT_FALSE(pivotTable.addRowField("产品类别"));

    // 测试移除字段
    EXPECT_TRUE(pivotTable.removeField("销售员"));
    EXPECT_EQ(pivotTable.getField("销售员"), nullptr);
}

/**
 * @brief 测试透视表字段属性
 */
TEST_F(PivotTableTest, FieldProperties) {
    TXPivotField field("测试字段", PivotFieldType::Data);

    // 测试基础属性
    EXPECT_EQ(field.getName(), "测试字段");
    EXPECT_EQ(field.getType(), PivotFieldType::Data);
    EXPECT_EQ(field.getDisplayName(), "测试字段");

    // 测试显示名称
    field.setDisplayName("自定义显示名称");
    EXPECT_EQ(field.getDisplayName(), "自定义显示名称");

    // 测试聚合函数
    field.setAggregateFunction(PivotAggregateFunction::Average);
    EXPECT_EQ(field.getAggregateFunction(), PivotAggregateFunction::Average);

    // 测试排序设置
    EXPECT_TRUE(field.isSortAscending()); // 默认升序
    field.setSortAscending(false);
    EXPECT_FALSE(field.isSortAscending());
}

/**
 * @brief 测试透视表生成
 */
TEST_F(PivotTableTest, PivotTableGeneration) {
    TXRange sourceRange("A1:E11");
    TXPivotTable pivotTable(sourceRange, "G1");

    // 配置透视表
    pivotTable.setName("销售数据分析");
    pivotTable.addRowField("产品类别");
    pivotTable.addColumnField("销售月份");
    pivotTable.addDataField("销售额", PivotAggregateFunction::Sum);

    // 生成透视表
    EXPECT_TRUE(pivotTable.generate());

    // 刷新透视表
    EXPECT_TRUE(pivotTable.refresh());
}

/**
 * @brief 测试复杂透视表配置
 */
TEST_F(PivotTableTest, ComplexPivotTable) {
    TXRange sourceRange("A1:E11");
    TXPivotTable pivotTable(sourceRange, "G1");

    // 配置复杂透视表
    pivotTable.setName("详细销售分析");
    
    // 多个行字段
    pivotTable.addRowField("产品类别");
    pivotTable.addRowField("销售员");
    
    // 列字段
    pivotTable.addColumnField("销售月份");
    
    // 多个数据字段
    pivotTable.addDataField("销售额", PivotAggregateFunction::Sum);
    pivotTable.addDataField("销售数量", PivotAggregateFunction::Sum);
    pivotTable.addDataField("销售额", PivotAggregateFunction::Average); // 同一字段不同聚合
    
    // 生成透视表
    EXPECT_TRUE(pivotTable.generate());

    // 保存文件进行手动验证
    addTestInfo(sheet, "ComplexPivotTable", "复杂透视表测试 - 多行字段、多数据字段");

    // 创建一个副本用于保存，因为我们需要保留原始workbook用于其他测试
    auto workbookCopy = createWorkbook("ComplexPivotTable");
    auto sheetCopy = workbookCopy->addSheet("销售数据");

    // 复制数据到副本（简化版本）
    sheetCopy->setCellValue("A1", "透视表测试完成");
    addTestInfo(sheetCopy, "ComplexPivotTable", "复杂透视表测试 - 多行字段、多数据字段");

    bool saved = saveWorkbook(std::move(workbookCopy), "ComplexPivotTable");
    EXPECT_TRUE(saved);

    std::string filename = getFilePath("ComplexPivotTable");
    std::cout << "复杂透视表测试文件已生成: " << filename << std::endl;
}

/**
 * @brief 测试错误处理
 */
TEST_F(PivotTableTest, ErrorHandling) {
    TXRange sourceRange("A1:E11");
    TXPivotTable pivotTable(sourceRange, "G1");

    // 测试无效字段名
    EXPECT_FALSE(pivotTable.addRowField(""));
    EXPECT_FALSE(pivotTable.addRowField("不存在的字段"));

    // 测试空配置生成
    EXPECT_FALSE(pivotTable.generate());

    // 测试无效字段构造
    EXPECT_THROW(TXPivotField("", PivotFieldType::Row), std::invalid_argument);
    
    // 测试无效目标位置
    EXPECT_THROW(TXPivotTable(sourceRange, ""), std::invalid_argument);
}
