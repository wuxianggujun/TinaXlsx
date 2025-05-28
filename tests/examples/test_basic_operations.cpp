//
// Basic Operations Test
// 测试基础的工作簿、工作表和单元格操作
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <filesystem>
#include <iostream>

class BasicOperationsTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 确保输出目录存在
        std::filesystem::create_directories("output");
    }

    void TearDown() override {
        // 测试完成后清理文件
        std::filesystem::remove("output/basic_test.xlsx");
    }
};

/**
 * @brief 测试创建基本工作簿和工作表
 */
TEST_F(BasicOperationsTest, CreateBasicWorkbookAndSheet) {
    std::cout << "\n=== 基础工作簿和工作表创建测试 ===\n";
    
    // 创建工作簿
    TinaXlsx::TXWorkbook workbook;
    EXPECT_TRUE(workbook.isEmpty());
    EXPECT_EQ(0, workbook.getSheetCount());
    
    // 添加工作表
    auto* sheet = workbook.addSheet("测试工作表");
    ASSERT_NE(sheet, nullptr);
    EXPECT_EQ(1, workbook.getSheetCount());
    EXPECT_FALSE(workbook.isEmpty());
    
    // 验证工作表名称
    EXPECT_EQ("测试工作表", sheet->getName());
    EXPECT_TRUE(workbook.hasSheet("测试工作表"));
    
    // 获取工作表
    auto* retrieved_sheet = workbook.getSheet("测试工作表");
    EXPECT_EQ(sheet, retrieved_sheet);
    
    auto* sheet_by_index = workbook.getSheet(0);
    EXPECT_EQ(sheet, sheet_by_index);
    
    std::cout << "基础工作簿创建测试通过！\n";
}

/**
 * @brief 测试不同数据类型的单元格操作
 */
TEST_F(BasicOperationsTest, CellDataTypes) {
    std::cout << "\n=== 单元格数据类型测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("数据类型测试");
    ASSERT_NE(sheet, nullptr);
    
    // 测试字符串
    EXPECT_TRUE(sheet->setCellValue("A1", std::string("Hello, TinaXlsx!")));
    auto string_value = sheet->getCellValue("A1");
    EXPECT_TRUE(std::holds_alternative<std::string>(string_value));
    if (std::holds_alternative<std::string>(string_value)) {
        EXPECT_EQ("Hello, TinaXlsx!", std::get<std::string>(string_value));
    }
    
    // 测试整数
    EXPECT_TRUE(sheet->setCellValue("A2", static_cast<int64_t>(42)));
    auto int_value = sheet->getCellValue("A2");
    EXPECT_TRUE(std::holds_alternative<int64_t>(int_value));
    if (std::holds_alternative<int64_t>(int_value)) {
        EXPECT_EQ(42, std::get<int64_t>(int_value));
    }
    
    // 测试浮点数
    EXPECT_TRUE(sheet->setCellValue("A3", 3.14159));
    auto double_value = sheet->getCellValue("A3");
    EXPECT_TRUE(std::holds_alternative<double>(double_value));
    if (std::holds_alternative<double>(double_value)) {
        EXPECT_DOUBLE_EQ(3.14159, std::get<double>(double_value));
    }
    
    // 测试布尔值
    EXPECT_TRUE(sheet->setCellValue("A4", true));
    auto bool_value = sheet->getCellValue("A4");
    EXPECT_TRUE(std::holds_alternative<bool>(bool_value));
    if (std::holds_alternative<bool>(bool_value)) {
        EXPECT_TRUE(std::get<bool>(bool_value));
    }
    
    // 测试负数
    EXPECT_TRUE(sheet->setCellValue("A5", static_cast<int64_t>(-100)));
    EXPECT_TRUE(sheet->setCellValue("A6", -2.718));
    
    // 保存文件
    bool saved = workbook.saveToFile("output/basic_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists("output/basic_test.xlsx"));
    
    std::cout << "单元格数据类型测试通过！\n";
}

/**
 * @brief 测试坐标系统
 */
TEST_F(BasicOperationsTest, CoordinateSystem) {
    std::cout << "\n=== 坐标系统测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("坐标测试");
    ASSERT_NE(sheet, nullptr);
    
    // 测试不同的坐标表示方法
    EXPECT_TRUE(sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(1), std::string("A1 坐标")));
    EXPECT_TRUE(sheet->setCellValue("B2", std::string("B2 地址")));
    
    // 验证坐标转换
    TinaXlsx::TXCoordinate coord1(TinaXlsx::row_t(1), TinaXlsx::column_t(1));
    EXPECT_EQ("A1", coord1.toAddress());
    
    TinaXlsx::TXCoordinate coord2 = TinaXlsx::TXCoordinate::fromAddress("B2");
    EXPECT_EQ(2U, coord2.getRow().index());
    EXPECT_EQ(2U, coord2.getCol().index());
    
    // 测试列名转换
    EXPECT_EQ("A", TinaXlsx::column_t(1).column_string());
    EXPECT_EQ("Z", TinaXlsx::column_t(26).column_string());
    EXPECT_EQ("AA", TinaXlsx::column_t(27).column_string());
    
    EXPECT_EQ(1U, TinaXlsx::column_t("A").index());
    EXPECT_EQ(26U, TinaXlsx::column_t("Z").index());
    EXPECT_EQ(27U, TinaXlsx::column_t("AA").index());
    
    std::cout << "坐标系统测试通过！\n";
}

/**
 * @brief 测试批量操作
 */
TEST_F(BasicOperationsTest, BatchOperations) {
    std::cout << "\n=== 批量操作测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("批量操作");
    ASSERT_NE(sheet, nullptr);
    
    // 准备批量数据
    std::vector<std::pair<TinaXlsx::TXCoordinate, TinaXlsx::TXSheet::CellValue>> batch_data;
    for (int i = 1; i <= 10; ++i) {
        batch_data.emplace_back(
            TinaXlsx::TXCoordinate(TinaXlsx::row_t(i), TinaXlsx::column_t(1)),
            std::string("批量数据_") + std::to_string(i)
        );
        batch_data.emplace_back(
            TinaXlsx::TXCoordinate(TinaXlsx::row_t(i), TinaXlsx::column_t(2)),
            static_cast<double>(i * 10)
        );
    }
    
    // 执行批量操作
    size_t success_count = sheet->setCellValues(batch_data);
    EXPECT_EQ(20, success_count); // 10行 × 2列 = 20个单元格
    
    // 验证批量设置的数据
    auto first_cell = sheet->getCellValue("A1");
    EXPECT_TRUE(std::holds_alternative<std::string>(first_cell));
    if (std::holds_alternative<std::string>(first_cell)) {
        EXPECT_EQ("批量数据_1", std::get<std::string>(first_cell));
    }
    
    auto last_cell = sheet->getCellValue("B10");
    EXPECT_TRUE(std::holds_alternative<double>(last_cell));
    if (std::holds_alternative<double>(last_cell)) {
        EXPECT_DOUBLE_EQ(100.0, std::get<double>(last_cell));
    }
    
    // 保存文件
    bool saved = workbook.saveToFile("output/basic_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "批量操作测试通过！\n";
}

/**
 * @brief 测试工作表管理
 */
TEST_F(BasicOperationsTest, SheetManagement) {
    std::cout << "\n=== 工作表管理测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    
    // 添加多个工作表
    auto* sheet1 = workbook.addSheet("Sheet1");
    auto* sheet2 = workbook.addSheet("Sheet2");
    auto* sheet3 = workbook.addSheet("Sheet3");
    
    ASSERT_NE(sheet1, nullptr);
    ASSERT_NE(sheet2, nullptr);
    ASSERT_NE(sheet3, nullptr);
    
    EXPECT_EQ(3, workbook.getSheetCount());
    
    // 测试工作表名称列表
    auto names = workbook.getSheetNames();
    EXPECT_EQ(3, names.size());
    EXPECT_EQ("Sheet1", names[0]);
    EXPECT_EQ("Sheet2", names[1]);
    EXPECT_EQ("Sheet3", names[2]);
    
    // 测试重命名
    EXPECT_TRUE(workbook.renameSheet("Sheet2", "重命名的工作表"));
    EXPECT_TRUE(workbook.hasSheet("重命名的工作表"));
    EXPECT_FALSE(workbook.hasSheet("Sheet2"));
    
    // 测试删除工作表
    EXPECT_TRUE(workbook.removeSheet("Sheet3"));
    EXPECT_EQ(2, workbook.getSheetCount());
    EXPECT_FALSE(workbook.hasSheet("Sheet3"));
    
    // 在剩余工作表中添加数据
    sheet1->setCellValue("A1", std::string("第一个工作表"));
    auto* renamed_sheet = workbook.getSheet("重命名的工作表");
    ASSERT_NE(renamed_sheet, nullptr);
    renamed_sheet->setCellValue("A1", std::string("重命名后的工作表"));
    
    // 保存文件
    bool saved = workbook.saveToFile("output/basic_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "工作表管理测试通过！\n";
} 