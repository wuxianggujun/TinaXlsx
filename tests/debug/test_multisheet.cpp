//
// 多工作表调试测试 (GTest版本)
//
#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <iostream>

class MultisheetTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前准备
    }

    void TearDown() override {
        // 测试后清理
    }
};

TEST_F(MultisheetTest, AddMultipleSheets) {
    TinaXlsx::TXWorkbook workbook;
    
    int initial_count = workbook.getSheetCount();
    EXPECT_GE(initial_count, 0) << "初始工作表数量应该 >= 0";
    
    // 添加多个工作表
    const int SHEET_COUNT = 5;
    for (int i = 1; i <= SHEET_COUNT; ++i) {
        std::string sheet_name = "DebugSheet_" + std::to_string(i);
        
        auto sheet = workbook.addSheet(sheet_name);
        ASSERT_NE(sheet, nullptr) << "工作表 " << sheet_name << " 创建失败";
        EXPECT_EQ(sheet->getName(), sheet_name);
        
        int expected_count = initial_count + i;
        EXPECT_EQ(workbook.getSheetCount(), expected_count) 
            << "添加工作表 " << i << " 后，工作表数量不正确";
        
        // 在工作表中添加测试数据
        sheet->setCellValue("A1", std::string("工作表: ") + sheet_name);
        sheet->setCellValue("A2", static_cast<int64_t>(i * 100));
        sheet->setCellValue("A3", static_cast<double>(i * 3.14));
        
        // 验证数据是否正确设置
        auto val_a1 = sheet->getCellValue("A1");
        auto val_a2 = sheet->getCellValue("A2");
        auto val_a3 = sheet->getCellValue("A3");
        
        EXPECT_TRUE(std::holds_alternative<std::string>(val_a1));
        EXPECT_TRUE(std::holds_alternative<int64_t>(val_a2));
        EXPECT_TRUE(std::holds_alternative<double>(val_a3));
        
        if (std::holds_alternative<std::string>(val_a1)) {
            EXPECT_EQ(std::get<std::string>(val_a1), std::string("工作表: ") + sheet_name);
        }
        if (std::holds_alternative<int64_t>(val_a2)) {
            EXPECT_EQ(std::get<int64_t>(val_a2), i * 100);
        }
        if (std::holds_alternative<double>(val_a3)) {
            EXPECT_DOUBLE_EQ(std::get<double>(val_a3), i * 3.14);
        }
    }
    
    // 最终验证
    EXPECT_EQ(workbook.getSheetCount(), initial_count + SHEET_COUNT);
}

TEST_F(MultisheetTest, AccessSheetsByIndex) {
    TinaXlsx::TXWorkbook workbook;
    
    // 创建几个工作表
    std::vector<std::string> sheet_names = {"First", "Second", "Third"};
    
    for (const auto& name : sheet_names) {
        auto sheet = workbook.addSheet(name);
        ASSERT_NE(sheet, nullptr) << "工作表 " << name << " 创建失败";
    }
    
    // 通过索引访问工作表
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        auto sheet = workbook.getSheet(static_cast<int>(i));
        if (sheet) {
            // 注意：如果有默认工作表，索引可能会偏移
            std::cout << "索引 " << i << " 的工作表名称: " << sheet->getName() << std::endl;
        }
    }
    
    // 验证总数
    EXPECT_GE(workbook.getSheetCount(), 3) << "工作表总数应该至少为3";
}

TEST_F(MultisheetTest, AccessSheetsByName) {
    TinaXlsx::TXWorkbook workbook;
    
    // 创建具有特定名称的工作表
    std::vector<std::string> sheet_names = {"销售数据", "财务报表", "库存统计"};
    
    for (const auto& name : sheet_names) {
        auto sheet = workbook.addSheet(name);
        ASSERT_NE(sheet, nullptr) << "工作表 " << name << " 创建失败";
        
        // 在每个工作表添加标识数据
        sheet->setCellValue("A1", name);
    }
    
    // 通过名称访问工作表并验证
    for (const auto& name : sheet_names) {
        auto sheet = workbook.getSheet(name);
        ASSERT_NE(sheet, nullptr) << "无法通过名称获取工作表: " << name;
        EXPECT_EQ(sheet->getName(), name);
        
        // 验证标识数据
        auto val = sheet->getCellValue("A1");
        EXPECT_TRUE(std::holds_alternative<std::string>(val));
        if (std::holds_alternative<std::string>(val)) {
            EXPECT_EQ(std::get<std::string>(val), name);
        }
    }
}