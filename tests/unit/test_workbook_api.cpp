#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;

class WorkbookAPITest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        std::remove("test_api.xlsx");
        std::remove("test_complex.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
};

// 测试工作表管理
TEST_F(WorkbookAPITest, SheetManagement) {
    // 添加工作表
    auto* sheet1 = workbook->addSheet("销售数据");
    auto* sheet2 = workbook->addSheet("统计报表");
    auto* sheet3 = workbook->addSheet("图表");
    
    ASSERT_NE(sheet1, nullptr);
    ASSERT_NE(sheet2, nullptr);
    ASSERT_NE(sheet3, nullptr);
    
    // 验证工作表名称
    EXPECT_EQ(sheet1->getName(), "销售数据");
    EXPECT_EQ(sheet2->getName(), "统计报表");
    EXPECT_EQ(sheet3->getName(), "图表");
    
    // 验证工作表数量
    EXPECT_EQ(workbook->getSheetCount(), 3);
    
    // 按名称获取工作表
    auto* foundSheet = workbook->getSheet("统计报表");
    ASSERT_NE(foundSheet, nullptr);
    EXPECT_EQ(foundSheet->getName(), "统计报表");
    
    // 按索引获取工作表
    auto* indexSheet = workbook->getSheet(1); // 第二个工作表 (0-based)
    ASSERT_NE(indexSheet, nullptr);
    EXPECT_EQ(indexSheet->getName(), "统计报表");
}

// 测试批量数据操作
TEST_F(WorkbookAPITest, BatchDataOperations) {
    auto* sheet = workbook->addSheet("BatchTest");
    ASSERT_NE(sheet, nullptr);
    
    // 准备批量数据
    std::vector<std::pair<TXCoordinate, TXCell::CellValue>> batchData = {
        {TXCoordinate(row_t(1), column_t(1)), std::string("产品名称")},
        {TXCoordinate(row_t(1), column_t(2)), std::string("单价")},
        {TXCoordinate(row_t(1), column_t(3)), std::string("数量")},
        {TXCoordinate(row_t(1), column_t(4)), std::string("总额")},
        
        {TXCoordinate(row_t(2), column_t(1)), std::string("苹果")},
        {TXCoordinate(row_t(2), column_t(2)), 5.50},
        {TXCoordinate(row_t(2), column_t(3)), int64_t(100)},
        {TXCoordinate(row_t(2), column_t(4)), 550.0},
        
        {TXCoordinate(row_t(3), column_t(1)), std::string("香蕉")},
        {TXCoordinate(row_t(3), column_t(2)), 3.20},
        {TXCoordinate(row_t(3), column_t(3)), int64_t(80)},
        {TXCoordinate(row_t(3), column_t(4)), 256.0}
    };
    
    // 批量设置值
    std::size_t setCount = sheet->setCellValues(batchData);
    EXPECT_EQ(setCount, 12);
    
    // 验证数据
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(1)), TXCell::CellValue(std::string("产品名称")));
    EXPECT_EQ(sheet->getCellValue(row_t(2), column_t(2)), TXCell::CellValue(5.50));
    EXPECT_EQ(sheet->getCellValue(row_t(3), column_t(3)), TXCell::CellValue(int64_t(80)));
    
    // 批量获取值
    std::vector<TXCoordinate> coords = {
        TXCoordinate(row_t(1), column_t(1)),
        TXCoordinate(row_t(2), column_t(2)),
        TXCoordinate(row_t(3), column_t(4))
    };
    
    auto values = sheet->getCellValues(coords);
    EXPECT_EQ(values.size(), 3);
    EXPECT_EQ(values[0].second, TXCell::CellValue(std::string("产品名称")));
    EXPECT_EQ(values[1].second, TXCell::CellValue(5.50));
    EXPECT_EQ(values[2].second, TXCell::CellValue(256.0));
}

// 测试范围操作
TEST_F(WorkbookAPITest, RangeOperations) {
    auto* sheet = workbook->addSheet("RangeTest");
    ASSERT_NE(sheet, nullptr);
    
    // 准备2x3的数据矩阵
    std::vector<std::vector<TXCell::CellValue>> matrixData = {
        {std::string("A1"), std::string("B1"), std::string("C1")},
        {100.0, 200.0, 300.0}
    };
    
    // 定义范围
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(3)));
    
    // 设置范围数据
    EXPECT_TRUE(sheet->setRangeValues(range, matrixData));
    
    // 获取范围数据
    auto retrievedData = sheet->getRangeValues(range);
    EXPECT_EQ(retrievedData.size(), 2);
    EXPECT_EQ(retrievedData[0].size(), 3);
    EXPECT_EQ(retrievedData[1].size(), 3);
    
    // 验证数据
    EXPECT_EQ(retrievedData[0][0], TXCell::CellValue(std::string("A1")));
    EXPECT_EQ(retrievedData[0][2], TXCell::CellValue(std::string("C1")));
    EXPECT_EQ(retrievedData[1][1], TXCell::CellValue(200.0));
    
    // 获取已使用范围
    auto usedRange = sheet->getUsedRange();
    EXPECT_TRUE(usedRange.isValid());
    EXPECT_EQ(usedRange.getStart().getRow(), row_t(1));
    EXPECT_EQ(usedRange.getStart().getCol(), column_t(1));
    EXPECT_EQ(usedRange.getEnd().getRow(), row_t(2));
    EXPECT_EQ(usedRange.getEnd().getCol(), column_t(3));
}

// 测试行列操作
TEST_F(WorkbookAPITest, RowColumnOperations) {
    auto* sheet = workbook->addSheet("RowColTest");
    ASSERT_NE(sheet, nullptr);
    
    // 设置初始数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("A1"));
    sheet->setCellValue(row_t(2), column_t(1), std::string("A2"));
    sheet->setCellValue(row_t(3), column_t(1), std::string("A3"));
    sheet->setCellValue(row_t(1), column_t(2), std::string("B1"));
    sheet->setCellValue(row_t(2), column_t(2), std::string("B2"));
    sheet->setCellValue(row_t(3), column_t(2), std::string("B3"));
    
    // 插入行（在第2行前插入）
    EXPECT_TRUE(sheet->insertRows(row_t(2), row_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(1)), TXCell::CellValue(std::string("A1"))); // 不变
    EXPECT_EQ(sheet->getCellValue(row_t(3), column_t(1)), TXCell::CellValue(std::string("A2"))); // 原A2移动到第3行
    EXPECT_EQ(sheet->getCellValue(row_t(4), column_t(1)), TXCell::CellValue(std::string("A3"))); // 原A3移动到第4行
    
    // 插入列（在第2列前插入）
    EXPECT_TRUE(sheet->insertColumns(column_t(2), column_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(1)), TXCell::CellValue(std::string("A1"))); // 不变
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(3)), TXCell::CellValue(std::string("B1"))); // 原B1移动到第3列
    
    // 删除行
    EXPECT_TRUE(sheet->deleteRows(row_t(2), row_t(1))); // 删除插入的空行
    
    // 删除列
    EXPECT_TRUE(sheet->deleteColumns(column_t(2), column_t(1))); // 删除插入的空列
    
    // 验证数据恢复到接近原状态
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(1)), TXCell::CellValue(std::string("A1")));
    EXPECT_EQ(sheet->getCellValue(row_t(2), column_t(1)), TXCell::CellValue(std::string("A2")));
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(2)), TXCell::CellValue(std::string("B1")));
}

// 测试复杂表格创建
TEST_F(WorkbookAPITest, ComplexTableCreation) {
    auto* sheet = workbook->addSheet("复杂表格");
    ASSERT_NE(sheet, nullptr);
    
    // 创建表头
    sheet->setCellValue("A1", std::string("员工编号"));
    sheet->setCellValue("B1", std::string("姓名"));
    sheet->setCellValue("C1", std::string("部门"));
    sheet->setCellValue("D1", std::string("基本工资"));
    sheet->setCellValue("E1", std::string("绩效奖金"));
    sheet->setCellValue("F1", std::string("总工资"));
    
    // 表头样式
    TXFont headerFont;
    headerFont.setBold(true);
    headerFont.setSize(12);
    headerFont.setColor(TXColor(255, 255, 255)); // 白色
    
    TXFill headerFill;
    headerFill.setPattern(FillPattern::Solid);
    headerFill.setForegroundColor(TXColor(79, 129, 189)); // 蓝色
    
    TXCellStyle headerStyle;
    headerStyle.setFont(headerFont);
    headerStyle.setFill(headerFill);
    
    // 应用表头样式
    TXRange headerRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(1), column_t(6)));
    sheet->setRangeStyle(headerRange, headerStyle);
    
    // 添加数据
    std::vector<std::vector<TXCell::CellValue>> employeeData = {
        {std::string("E001"), std::string("张三"), std::string("技术部"), 8000.0, 2000.0, 10000.0},
        {std::string("E002"), std::string("李四"), std::string("销售部"), 6000.0, 3000.0, 9000.0},
        {std::string("E003"), std::string("王五"), std::string("财务部"), 7000.0, 1500.0, 8500.0},
        {std::string("E004"), std::string("赵六"), std::string("人事部"), 6500.0, 1000.0, 7500.0}
    };
    
    TXRange dataRange(TXCoordinate(row_t(2), column_t(1)), TXCoordinate(row_t(5), column_t(6)));
    EXPECT_TRUE(sheet->setRangeValues(dataRange, employeeData));
    
    // 设置货币格式
    TXRange salaryRange(TXCoordinate(row_t(2), column_t(4)), TXCoordinate(row_t(5), column_t(6)));
    sheet->setRangeNumberFormat(salaryRange, TXCell::NumberFormat::Currency, 2);
    
    // 合并标题行中的某些单元格作为示例
    sheet->setCellValue("G1", std::string("备注信息"));
    sheet->mergeCells("G1:H1");
    
    // 保存复杂表格
    EXPECT_TRUE(workbook->saveToFile("test_complex.xlsx"));
}

// 测试组件管理
TEST_F(WorkbookAPITest, ComponentManagement) {
    // 检查默认组件
    auto& componentManager = workbook->getComponentManager();
    
    // SharedStrings应该默认注册
    EXPECT_TRUE(componentManager.hasComponent(ExcelComponent::SharedStrings));
    
    // 添加样式时应该自动注册Styles组件
    auto* sheet = workbook->addSheet("StyleSheet");
    TXCellStyle style;
    TXFont font;
    font.setBold(true);
    style.setFont(font);
    
    sheet->setCellValue(row_t(1), column_t(1), std::string("Styled Text"));
    sheet->setCellStyle(row_t(1), column_t(1), style);
    
    EXPECT_TRUE(componentManager.hasComponent(ExcelComponent::Styles));
    
    // 添加合并单元格时应该自动注册MergedCells组件
    sheet->mergeCells(row_t(2), column_t(1), row_t(3), column_t(2));
    EXPECT_TRUE(componentManager.hasComponent(ExcelComponent::MergedCells));
}

// 测试错误处理
TEST_F(WorkbookAPITest, ErrorHandling) {
    auto* sheet = workbook->addSheet("ErrorTest");
    ASSERT_NE(sheet, nullptr);
    
    // 测试无效坐标
    EXPECT_FALSE(sheet->setCellValue(row_t(0), column_t(1), std::string("Invalid"))); // 行号从1开始
    
    // 测试重复工作表名称
    auto* duplicateSheet = workbook->addSheet("ErrorTest"); // 重复名称
    // 具体行为取决于实现，这里假设允许重复名称或自动重命名
    
    // 测试无效范围
    TXRange invalidRange(TXCoordinate(row_t(5), column_t(5)), TXCoordinate(row_t(1), column_t(1))); // 起始>结束
    std::vector<std::vector<TXCell::CellValue>> data = {{std::string("test")}};
    EXPECT_FALSE(sheet->setRangeValues(invalidRange, data));
    
    // 测试错误信息
    std::string lastError = sheet->getLastError();
    EXPECT_FALSE(lastError.empty()); // 应该有错误信息
}

// 测试性能相关的大数据操作
TEST_F(WorkbookAPITest, LargeDataHandling) {
    auto* sheet = workbook->addSheet("LargeData");
    ASSERT_NE(sheet, nullptr);
    
    // 生成大量数据（100行x5列）
    std::vector<std::pair<TXCoordinate, TXCell::CellValue>> largeData;
    largeData.reserve(500);
    
    for (u32 row = 1; row <= 100; ++row) {
        for (u32 col = 1; col <= 5; ++col) {
            // 明确使用构造函数创建对象
            auto coord = TXCoordinate{row_t(row), column_t(col)};
            TXCell::CellValue value;
            
            if (col == 1) {
                value = std::string("Row") + std::to_string(row);
            } else {
                value = static_cast<double>(row * col * 10.5);
            }
            
            // 使用initializer_list语法避免类型推导问题
            largeData.push_back({coord, value});
        }
    }
    
    // 批量设置数据
    std::size_t setCount = sheet->setCellValues(largeData);
    EXPECT_EQ(setCount, 500);
    
    // 验证使用范围
    auto usedRange = sheet->getUsedRange();
    EXPECT_EQ(usedRange.getEnd().getRow(), row_t(100));
    EXPECT_EQ(usedRange.getEnd().getCol(), column_t(5));
    
    // 验证部分数据
    EXPECT_EQ(sheet->getCellValue(row_t(50), column_t(1)), TXCell::CellValue(std::string("Row50")));
    EXPECT_EQ(sheet->getCellValue(row_t(50), column_t(3)), TXCell::CellValue(50.0 * 3.0 * 10.5));
    
    // 保存大数据文件
    EXPECT_TRUE(workbook->saveToFile("test_api.xlsx"));
} 