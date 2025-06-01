#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <memory>
#include <iostream>

using namespace TinaXlsx;

class ProtectionDebugTest : public TestWithFileGeneration<ProtectionDebugTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<ProtectionDebugTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("保护调试");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<ProtectionDebugTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(ProtectionDebugTest, DebugProtectionStatus) {
    std::cout << "\n=== 保护状态调试测试 ===" << std::endl;
    
    // 添加一些测试数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{"锁定单元格"});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{"未锁定单元格"});
    
    // 设置锁定状态
    std::cout << "设置单元格锁定状态..." << std::endl;
    bool result1 = sheet->setCellLocked(row_t(2), column_t(1), true);
    bool result2 = sheet->setCellLocked(row_t(3), column_t(1), false);
    std::cout << "设置B2锁定: " << (result1 ? "成功" : "失败") << std::endl;
    std::cout << "设置B3未锁定: " << (result2 ? "成功" : "失败") << std::endl;
    
    // 检查锁定状态
    bool locked1 = sheet->isCellLocked(row_t(2), column_t(1));
    bool locked2 = sheet->isCellLocked(row_t(3), column_t(1));
    std::cout << "B2锁定状态: " << (locked1 ? "锁定" : "未锁定") << std::endl;
    std::cout << "B3锁定状态: " << (locked2 ? "锁定" : "未锁定") << std::endl;
    
    // 检查保护状态（保护前）
    auto& protectionManager = sheet->getProtectionManager();
    std::cout << "保护前工作表保护状态: " << (protectionManager.isSheetProtected() ? "已保护" : "未保护") << std::endl;
    
    // 保护工作表
    std::cout << "保护工作表..." << std::endl;
    bool protectResult = sheet->protectSheet("test123");
    std::cout << "保护工作表结果: " << (protectResult ? "成功" : "失败") << std::endl;
    
    // 检查保护状态（保护后）
    std::cout << "保护后工作表保护状态: " << (protectionManager.isSheetProtected() ? "已保护" : "未保护") << std::endl;
    
    // 获取保护设置详情
    const auto& protection = protectionManager.getSheetProtection();
    std::cout << "保护设置详情:" << std::endl;
    std::cout << "  isProtected: " << protection.isProtected << std::endl;
    std::cout << "  passwordHash: " << protection.passwordHash << std::endl;
    std::cout << "  selectLockedCells: " << protection.selectLockedCells << std::endl;
    std::cout << "  selectUnlockedCells: " << protection.selectUnlockedCells << std::endl;
    std::cout << "  formatCells: " << protection.formatCells << std::endl;
    
    // 检查单元格可编辑性
    bool editable1 = protectionManager.isCellEditable(TXCoordinate(row_t(2), column_t(1)), sheet->getCellManager());
    bool editable2 = protectionManager.isCellEditable(TXCoordinate(row_t(3), column_t(1)), sheet->getCellManager());
    std::cout << "B2可编辑性: " << (editable1 ? "可编辑" : "不可编辑") << std::endl;
    std::cout << "B3可编辑性: " << (editable2 ? "可编辑" : "不可编辑") << std::endl;
    
    // 添加说明到文件中
    sheet->setCellValue(row_t(5), column_t(1), cell_value_t{"保护状态调试信息:"});
    sheet->setCellValue(row_t(6), column_t(1), cell_value_t{"工作表已保护: " + std::string(protectionManager.isSheetProtected() ? "是" : "否")});
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"密码哈希: " + protection.passwordHash});
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"B2锁定: " + std::string(locked1 ? "是" : "否")});
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B3锁定: " + std::string(locked2 ? "是" : "否")});
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"B2可编辑: " + std::string(editable1 ? "是" : "否")});
    sheet->setCellValue(row_t(11), column_t(1), cell_value_t{"B3可编辑: " + std::string(editable2 ? "是" : "否")});
    
    saveWorkbook(workbook, "DebugProtectionStatus");
    
    std::cout << "=== 保护状态调试测试完成 ===" << std::endl;
}
