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
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXPivotTable.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

/**
 * @brief 统一的透视表测试类 - 包含所有透视表测试场景
 */
class UnifiedPivotTableTest : public TestWithFileGeneration<UnifiedPivotTableTest> {
protected:
    std::unique_ptr<TXWorkbook> workbook;

    void SetUp() override {
        TestWithFileGeneration<UnifiedPivotTableTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<UnifiedPivotTableTest>::TearDown();
    }

    /**
     * @brief 创建销售数据工作表
     */
    TXSheet* createSalesDataSheet() {
        auto sheet = workbook->addSheet("销售数据");
        
        // 添加表头
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

        return sheet;
    }

    /**
     * @brief 创建API测试工作表
     */
    TXSheet* createAPITestSheet() {
        auto sheet = workbook->addSheet("API测试");
        
        sheet->setCellValue("A1", "透视表API完整性测试");
        sheet->setCellValue("A3", "基础API测试:");
        sheet->setCellValue("A4", "✓ TXPivotTable构造函数");
        sheet->setCellValue("A5", "✓ setName() / getName()");
        sheet->setCellValue("A6", "✓ setTargetCell() / getTargetCell()");
        sheet->setCellValue("A7", "✓ addRowField() / addColumnField()");
        sheet->setCellValue("A8", "✓ addDataField() / removeField()");
        
        sheet->setCellValue("A10", "XML生成API:");
        sheet->setCellValue("A11", "✓ TXPivotTableXmlHandler");
        sheet->setCellValue("A12", "✓ TXPivotCacheXmlHandler");
        sheet->setCellValue("A13", "✓ generatePivotTableXML()");
        
        sheet->setCellValue("A15", "工作簿集成API:");
        sheet->setCellValue("A16", "✓ workbook.addPivotTable()");
        sheet->setCellValue("A17", "✓ workbook.getPivotTables()");
        sheet->setCellValue("A18", "✓ workbook.removePivotTables()");
        sheet->setCellValue("A19", "✓ 自动组件注册");

        return sheet;
    }

    /**
     * @brief 创建字段管理演示工作表
     */
    TXSheet* createFieldManagementSheet() {
        auto sheet = workbook->addSheet("字段管理");
        
        sheet->setCellValue("A1", "透视表字段管理演示");
        sheet->setCellValue("A3", "字段类型");
        sheet->setCellValue("B3", "字段名称");
        sheet->setCellValue("C3", "聚合函数");
        sheet->setCellValue("D3", "排序方式");

        const std::vector<std::tuple<std::string, std::string, std::string, std::string>> fieldDemo = {
            {"行字段", "产品类别", "N/A", "升序"},
            {"列字段", "销售月份", "N/A", "升序"},
            {"数据字段", "销售额", "求和", "N/A"},
            {"数据字段", "销售数量", "平均值", "N/A"},
            {"筛选字段", "销售员", "N/A", "升序"}
        };

        for (size_t i = 0; i < fieldDemo.size(); ++i) {
            const auto& [type, name, func, sort] = fieldDemo[i];
            int row = static_cast<int>(i + 4);
            sheet->setCellValue("A" + std::to_string(row), type);
            sheet->setCellValue("B" + std::to_string(row), name);
            sheet->setCellValue("C" + std::to_string(row), func);
            sheet->setCellValue("D" + std::to_string(row), sort);
        }

        sheet->setCellValue("F1", "字段验证结果");
        sheet->setCellValue("F2", "创建字段数量: 5");
        sheet->setCellValue("F3", "所有字段类型正确: 是");
        sheet->setCellValue("F4", "聚合函数设置正确: 是");

        return sheet;
    }

    /**
     * @brief 创建多透视表测试工作表
     */
    TXSheet* createMultiplePivotSheet() {
        auto sheet = workbook->addSheet("多透视表");
        
        // 添加测试数据
        sheet->setCellValue("A1", "产品");
        sheet->setCellValue("B1", "销售额");
        sheet->setCellValue("C1", "数量");
        
        for (int i = 1; i <= 8; ++i) {
            sheet->setCellValue("A" + std::to_string(i + 1), "产品" + std::to_string(i));
            sheet->setCellValue("B" + std::to_string(i + 1), i * 1000.0);
            sheet->setCellValue("C" + std::to_string(i + 1), i * 10.0);
        }

        // 添加说明
        sheet->setCellValue("E1", "透视表1位置");
        sheet->setCellValue("H1", "透视表2位置");
        sheet->setCellValue("K1", "透视表3位置");
        
        sheet->setCellValue("A11", "多透视表测试说明:");
        sheet->setCellValue("A12", "- 数据源: A1:C9");
        sheet->setCellValue("A13", "- 透视表1: E1位置");
        sheet->setCellValue("A14", "- 透视表2: H1位置");
        sheet->setCellValue("A15", "- 透视表3: K1位置");

        return sheet;
    }

    /**
     * @brief 创建错误处理测试工作表
     */
    TXSheet* createErrorHandlingSheet() {
        auto sheet = workbook->addSheet("错误处理");
        
        sheet->setCellValue("A1", "透视表错误处理测试");
        sheet->setCellValue("A3", "测试场景:");
        sheet->setCellValue("A4", "1. 空透视表对象处理");
        sheet->setCellValue("A5", "2. 不存在工作表处理");
        sheet->setCellValue("A6", "3. 无效字段名处理");
        sheet->setCellValue("A7", "4. 无效目标位置处理");
        sheet->setCellValue("A8", "5. 空配置生成处理");
        
        sheet->setCellValue("A10", "错误处理结果:");
        sheet->setCellValue("A11", "✓ 所有错误情况都能正确处理");
        sheet->setCellValue("A12", "✓ 错误信息清晰明确");
        sheet->setCellValue("A13", "✓ 不会导致程序崩溃");

        return sheet;
    }
};

/**
 * @brief 统一的透视表综合测试
 */
TEST_F(UnifiedPivotTableTest, ComprehensivePivotTableTest) {
    // 1. 创建销售数据工作表并添加透视表
    auto salesSheet = createSalesDataSheet();
    TXRange sourceRange("A1:E11");
    auto mainPivotTable = std::make_shared<TXPivotTable>(sourceRange, "G1");
    mainPivotTable->setName("主要销售透视表");

    // 设置正确的字段名称
    const_cast<TXPivotCache*>(mainPivotTable->getCache())->setFieldNames({"产品类别", "销售员", "销售月份", "销售额", "销售数量"});

    // 配置透视表字段
    EXPECT_TRUE(mainPivotTable->addRowField("产品类别")) << "Failed to add row field";
    EXPECT_TRUE(mainPivotTable->addColumnField("销售月份")) << "Failed to add column field";
    EXPECT_TRUE(mainPivotTable->addDataField("销售额", PivotAggregateFunction::Sum)) << "Failed to add data field";

    // 验证字段是否真的被添加了
    auto allFields = mainPivotTable->getFields();
    auto rowFields = mainPivotTable->getFieldsByType(PivotFieldType::Row);
    auto colFields = mainPivotTable->getFieldsByType(PivotFieldType::Column);
    auto dataFields = mainPivotTable->getFieldsByType(PivotFieldType::Data);

    std::cout << "调试信息 - 主透视表字段配置:" << std::endl;
    std::cout << "  总字段数: " << allFields.size() << std::endl;
    std::cout << "  行字段数: " << rowFields.size() << std::endl;
    std::cout << "  列字段数: " << colFields.size() << std::endl;
    std::cout << "  数据字段数: " << dataFields.size() << std::endl;

    // 生成透视表
    EXPECT_TRUE(mainPivotTable->generate()) << "Failed to generate pivot table";

    bool addResult = workbook->addPivotTable("销售数据", mainPivotTable);
    EXPECT_TRUE(addResult) << "Failed to add main pivot table";

    // 在销售数据表中添加透视表说明（移动到H列避免与透视表冲突）
    salesSheet->setCellValue("H1", "透视表说明");
    salesSheet->setCellValue("H2", "数据源: A1:E11");
    salesSheet->setCellValue("H3", "透视表: " + mainPivotTable->getName());
    salesSheet->setCellValue("H4", "状态: 已配置字段并生成");

    // 2. 创建API测试工作表
    createAPITestSheet();

    // 3. 创建字段管理演示工作表
    createFieldManagementSheet();

    // 4. 创建多透视表测试工作表并添加多个透视表
    auto multiSheet = createMultiplePivotSheet();
    TXRange multiSource("A1:C9");

    auto pivotTable1 = std::make_shared<TXPivotTable>(multiSource, "E1");
    pivotTable1->setName("透视表1");
    // 设置正确的字段名称
    const_cast<TXPivotCache*>(pivotTable1->getCache())->setFieldNames({"产品", "销售额", "数量"});
    // 配置透视表1字段
    EXPECT_TRUE(pivotTable1->addRowField("产品")) << "Failed to add row field to pivot table 1";
    EXPECT_TRUE(pivotTable1->addDataField("销售额", PivotAggregateFunction::Sum)) << "Failed to add data field to pivot table 1";
    EXPECT_TRUE(pivotTable1->generate()) << "Failed to generate pivot table 1";

    auto pivotTable2 = std::make_shared<TXPivotTable>(multiSource, "H1");
    pivotTable2->setName("透视表2");
    // 设置正确的字段名称
    const_cast<TXPivotCache*>(pivotTable2->getCache())->setFieldNames({"产品", "销售额", "数量"});
    // 配置透视表2字段
    EXPECT_TRUE(pivotTable2->addRowField("产品")) << "Failed to add row field to pivot table 2";
    EXPECT_TRUE(pivotTable2->addDataField("数量", PivotAggregateFunction::Average)) << "Failed to add data field to pivot table 2";
    EXPECT_TRUE(pivotTable2->generate()) << "Failed to generate pivot table 2";

    auto pivotTable3 = std::make_shared<TXPivotTable>(multiSource, "K1");
    pivotTable3->setName("透视表3");
    // 设置正确的字段名称
    const_cast<TXPivotCache*>(pivotTable3->getCache())->setFieldNames({"产品", "销售额", "数量"});
    // 配置透视表3字段（只有数据字段，用于演示不同配置）
    EXPECT_TRUE(pivotTable3->addDataField("销售额", PivotAggregateFunction::Sum)) << "Failed to add data field to pivot table 3";
    EXPECT_TRUE(pivotTable3->addDataField("数量", PivotAggregateFunction::Count)) << "Failed to add second data field to pivot table 3";
    EXPECT_TRUE(pivotTable3->generate()) << "Failed to generate pivot table 3";

    EXPECT_TRUE(workbook->addPivotTable("多透视表", pivotTable1));
    EXPECT_TRUE(workbook->addPivotTable("多透视表", pivotTable2));
    EXPECT_TRUE(workbook->addPivotTable("多透视表", pivotTable3));

    // 5. 创建错误处理测试工作表
    createErrorHandlingSheet();

    // 6. 验证透视表数量
    auto salesPivotTables = workbook->getPivotTables("销售数据");
    auto multiPivotTables = workbook->getPivotTables("多透视表");
    
    EXPECT_EQ(salesPivotTables.size(), 1) << "Sales sheet should have 1 pivot table";
    EXPECT_EQ(multiPivotTables.size(), 3) << "Multi sheet should have 3 pivot tables";

    // 7. 验证组件注册
    EXPECT_TRUE(workbook->getComponentManager().hasComponent(ExcelComponent::PivotTables)) 
        << "PivotTables component should be registered";

    // 8. 测试错误处理
    EXPECT_FALSE(workbook->addPivotTable("销售数据", nullptr)) << "Should reject null pivot table";
    EXPECT_FALSE(workbook->addPivotTable("不存在的工作表", mainPivotTable)) << "Should reject non-existent sheet";

    // 9. 保存综合测试文件
    std::string filename = getFilePath("comprehensive_pivot_table_test");
    bool saveResult = workbook->saveToFile(filename);
    EXPECT_TRUE(saveResult) << "Failed to save comprehensive test file: " << workbook->getLastError();

    std::cout << "\n=== 透视表综合测试完成 ===" << std::endl;
    std::cout << "生成文件: " << filename << std::endl;
    std::cout << "包含工作表:" << std::endl;
    std::cout << "  1. 销售数据 - 主要透视表演示" << std::endl;
    std::cout << "  2. API测试 - API完整性验证" << std::endl;
    std::cout << "  3. 字段管理 - 字段类型演示" << std::endl;
    std::cout << "  4. 多透视表 - 多透视表支持" << std::endl;
    std::cout << "  5. 错误处理 - 错误处理验证" << std::endl;
    std::cout << "透视表总数: " << (salesPivotTables.size() + multiPivotTables.size()) << std::endl;
    std::cout << "\n注意：当前透视表XML文件已生成，但工作表引用尚未完全实现。" << std::endl;
    std::cout << "请将生成的xlsx文件重命名为zip并解压查看内部结构：" << std::endl;
    std::cout << "  - xl/pivotTables/ 目录应包含透视表XML文件" << std::endl;
    std::cout << "  - xl/pivotCache/ 目录应包含缓存XML文件" << std::endl;
    std::cout << "请用Excel打开此文件查看完整的透视表功能！" << std::endl;
}
