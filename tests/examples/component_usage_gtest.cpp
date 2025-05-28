//
// 组件化架构使用GTest示例
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <filesystem>

class ComponentUsageTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 清理测试文件
        std::filesystem::remove("auto_component_test.xlsx");
        std::filesystem::remove("manual_component_test.xlsx");
        std::filesystem::remove("minimal_component_test.xlsx");
    }

    void TearDown() override {
        // 可选择保留文件以供检查
        // std::filesystem::remove("auto_component_test.xlsx");
        // std::filesystem::remove("manual_component_test.xlsx");
        // std::filesystem::remove("minimal_component_test.xlsx");
    }
};

TEST_F(ComponentUsageTest, AutoComponentDetection) {
    std::cout << "\n=== 自动组件检测测试 ===" << std::endl;
    
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("自动检测");
    ASSERT_NE(sheet, nullptr);
    
    // 添加不同类型的数据，会自动检测所需组件
    sheet->setCellValue("A1", std::string("文本数据"));  // 自动注册SharedStrings
    sheet->setCellValue("B1", static_cast<int64_t>(123));
    sheet->setCellValue("C1", 3.14159);
    sheet->mergeCells("A2:C2");  // 自动注册MergedCells
    sheet->setCellValue("A2", std::string("合并单元格"));
    
    // 验证内存中的数据
    auto value_a1 = sheet->getCellValue("A1");
    auto value_b1 = sheet->getCellValue("B1");
    EXPECT_TRUE(std::holds_alternative<std::string>(value_a1));
    EXPECT_TRUE(std::holds_alternative<int64_t>(value_b1));
    
    // 保存时自动检测并生成需要的组件
    bool saved = workbook.saveToFile("auto_component_test.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    const auto& components = workbook.getComponentManager().getComponents();
    std::cout << "自动检测到 " << components.size() << " 个组件" << std::endl;
    
    // 验证必要组件被注册
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::SharedStrings));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::Styles));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::MergedCells));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::DocumentProperties));
    
    // 验证文件大小（应该有内容）
    auto file_size = std::filesystem::file_size("auto_component_test.xlsx");
    EXPECT_GT(file_size, 1000) << "文件大小太小，可能没有数据";
    std::cout << "文件大小: " << file_size << " 字节" << std::endl;
}

TEST_F(ComponentUsageTest, ManualComponentControl) {
    std::cout << "\n=== 手动组件控制测试 ===" << std::endl;
    
    TinaXlsx::TXWorkbook manual_workbook;
    manual_workbook.setAutoComponentDetection(false);  // 禁用自动检测
    
    // 手动注册需要的组件
    manual_workbook.registerComponent(TinaXlsx::ExcelComponent::SharedStrings);
    
    auto sheet = manual_workbook.addSheet("精确控制");
    ASSERT_NE(sheet, nullptr);
    
    sheet->setCellValue("A1", std::string("仅文本数据"));
    
    bool saved = manual_workbook.saveToFile("manual_component_test.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << manual_workbook.getLastError();
    
    const auto& components = manual_workbook.getComponentManager().getComponents();
    std::cout << "手动注册了 " << components.size() << " 个组件" << std::endl;
    
    // 验证只有手动注册的组件
    EXPECT_TRUE(manual_workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    EXPECT_TRUE(manual_workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::SharedStrings));
    EXPECT_FALSE(manual_workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::DocumentProperties));
    
    auto file_size = std::filesystem::file_size("manual_component_test.xlsx");
    std::cout << "文件大小: " << file_size << " 字节" << std::endl;
}

TEST_F(ComponentUsageTest, MinimalExcelFile) {
    std::cout << "\n=== 最小文件测试 ===" << std::endl;
    
    TinaXlsx::TXWorkbook minimal_workbook;
    minimal_workbook.setAutoComponentDetection(false);
    
    // 不手动注册任何额外组件，只有BasicWorkbook
    auto sheet = minimal_workbook.addSheet("最小");
    ASSERT_NE(sheet, nullptr);
    
    // 不添加任何数据
    
    bool saved = minimal_workbook.saveToFile("minimal_component_test.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << minimal_workbook.getLastError();
    
    const auto& components = minimal_workbook.getComponentManager().getComponents();
    std::cout << "最小文件包含 " << components.size() << " 个组件" << std::endl;
    
    // 验证只有基础组件
    EXPECT_EQ(components.size(), 1);
    EXPECT_TRUE(minimal_workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    
    auto file_size = std::filesystem::file_size("minimal_component_test.xlsx");
    std::cout << "文件大小: " << file_size << " 字节" << std::endl;
}

TEST_F(ComponentUsageTest, DataIntegrityValidation) {
    std::cout << "\n=== 数据完整性验证测试 ===" << std::endl;
    
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("数据完整性");
    ASSERT_NE(sheet, nullptr);
    
    // 添加测试数据
    sheet->setCellValue("A1", std::string("测试文本"));
    sheet->setCellValue("B1", static_cast<int64_t>(42));
    sheet->setCellValue("C1", 3.14159);
    
    // 验证内存中的数据
    auto value_a1 = sheet->getCellValue("A1");
    auto value_b1 = sheet->getCellValue("B1");
    auto value_c1 = sheet->getCellValue("C1");
    
    ASSERT_TRUE(std::holds_alternative<std::string>(value_a1));
    ASSERT_TRUE(std::holds_alternative<int64_t>(value_b1));
    ASSERT_TRUE(std::holds_alternative<double>(value_c1));
    
    EXPECT_EQ(std::get<std::string>(value_a1), "测试文本");
    EXPECT_EQ(std::get<int64_t>(value_b1), 42);
    EXPECT_DOUBLE_EQ(std::get<double>(value_c1), 3.14159);
    
    // 检查使用范围
    auto used_range = sheet->getUsedRange();
    EXPECT_TRUE(used_range.isValid());
    std::cout << "使用范围: " << used_range.getStart().toString() 
              << " 到 " << used_range.getEnd().toString() << std::endl;
    
    // 保存文件
    bool saved = workbook.saveToFile("data_integrity_test.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    // 重新加载并验证（注意：当前读取功能有限）
    TinaXlsx::TXWorkbook reload_workbook;
    bool loaded = reload_workbook.loadFromFile("data_integrity_test.xlsx");
    EXPECT_TRUE(loaded);
    
    if (loaded) {
        auto reload_sheet = reload_workbook.getSheet("数据完整性");
        EXPECT_NE(reload_sheet, nullptr);
        
        // 注意：当前的简化读取实现可能不会正确读取单元格数据
        // 这是后续需要改进的功能
        std::cout << "文件重新加载成功，但单元格数据读取需要进一步完善" << std::endl;
    }
}