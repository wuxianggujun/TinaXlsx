//
// Component Architecture Test
// Tests the new component-based architecture for Excel generation
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <filesystem>

class ComponentArchitectureTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 清理测试文件
        std::filesystem::remove("test_auto_components.xlsx");
        std::filesystem::remove("test_manual_components.xlsx");
        std::filesystem::remove("test_minimal_components.xlsx");
    }

    void TearDown() override {
        // 清理测试文件
        std::filesystem::remove("test_auto_components.xlsx");
        std::filesystem::remove("test_manual_components.xlsx");
        std::filesystem::remove("test_minimal_components.xlsx");
    }
};

TEST_F(ComponentArchitectureTest, AutoComponentDetection) {
    TinaXlsx::TXWorkbook workbook;
    
    // 验证默认启用自动检测
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    
    auto sheet = workbook.addSheet("测试工作表");
    
    // 添加不同类型的数据
    sheet->setCellValue("A1", std::string("文本数据"));  // 应触发SharedStrings
    sheet->setCellValue("B1", 42.5);                   // 应触发Styles
    sheet->mergeCells("A2:B2");                        // 应触发MergedCells
    sheet->setCellValue("A2", std::string("合并单元格"));
    
    // 保存触发组件检测
    bool saved = workbook.saveToFile("test_auto_components.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    // 验证自动检测到的组件
    const auto& components = workbook.getComponentManager().getComponents();
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::SharedStrings));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::Styles));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::MergedCells));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::DocumentProperties));
    
    std::cout << "自动检测到 " << components.size() << " 个组件" << std::endl;
}

TEST_F(ComponentArchitectureTest, ManualComponentControl) {
    TinaXlsx::TXWorkbook workbook;
    
    // 禁用自动检测
    workbook.setAutoComponentDetection(false);
    
    // 手动注册最少必需的组件
    workbook.registerComponent(TinaXlsx::ExcelComponent::Styles);
    // 故意不注册SharedStrings和DocumentProperties
    
    auto sheet = workbook.addSheet("手动控制");
    sheet->setCellValue("A1", std::string("精简Excel"));
    
    bool saved = workbook.saveToFile("test_manual_components.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    // 验证只有手动注册的组件
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::Styles));
    EXPECT_FALSE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::DocumentProperties));
    
    const auto& components = workbook.getComponentManager().getComponents();
    std::cout << "手动控制：" << components.size() << " 个组件" << std::endl;
}

TEST_F(ComponentArchitectureTest, MinimalExcelFile) {
    TinaXlsx::TXWorkbook workbook;
    
    // 禁用自动检测，创建最小Excel文件
    workbook.setAutoComponentDetection(false);
    
    auto sheet = workbook.addSheet("最小文件");
    sheet->setCellValue("A1", 123);  // 只有数字，不需要SharedStrings
    
    bool saved = workbook.saveToFile("test_minimal_components.xlsx");
    ASSERT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    // 验证最少组件
    const auto& components = workbook.getComponentManager().getComponents();
    EXPECT_EQ(components.size(), 1);  // 只有BasicWorkbook
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(TinaXlsx::ExcelComponent::BasicWorkbook));
    
    std::cout << "最小文件：" << components.size() << " 个组件" << std::endl;
    
    // 验证文件可以重新加载
    TinaXlsx::TXWorkbook verify_workbook;
    bool loaded = verify_workbook.loadFromFile("test_minimal_components.xlsx");
    EXPECT_TRUE(loaded) << "加载失败: " << verify_workbook.getLastError();
    
    auto verify_sheet = verify_workbook.getSheet("最小文件");
    ASSERT_NE(verify_sheet, nullptr);
    
    // 注意：当前的简化读取实现还不支持单元格数据读取
    // 这里主要验证文件结构的正确性
    // TODO: 实现完整的单元格数据读取功能
    auto value = verify_sheet->getCellValue("A1");
    // EXPECT_TRUE(std::holds_alternative<int64_t>(value) || std::holds_alternative<double>(value));
} 