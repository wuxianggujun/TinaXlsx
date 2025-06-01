#include <gtest/gtest.h>
#include "TinaXlsx/TXSheetProtectionManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class TXSheetProtectionManagerTest : public TestWithFileGeneration<TXSheetProtectionManagerTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<TXSheetProtectionManagerTest>::SetUp();
        protectionManager = std::make_unique<TXSheetProtectionManager>();
        cellManager = std::make_unique<TXCellManager>();

        // 为文件生成创建工作簿和工作表
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("保护管理器测试");

        // 设置一些测试单元格
        cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
        cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string("B1"));
        cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("A2"));
        cellManager->setCellValue(TXCoordinate(row_t(2), column_t(2)), std::string("B2"));
    }

    void TearDown() override {
        protectionManager.reset();
        cellManager.reset();
        workbook.reset();
        TestWithFileGeneration<TXSheetProtectionManagerTest>::TearDown();
    }

    std::unique_ptr<TXSheetProtectionManager> protectionManager;
    std::unique_ptr<TXCellManager> cellManager;
    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

// ==================== 基本保护功能测试 ====================

TEST_F(TXSheetProtectionManagerTest, BasicProtection) {
    // 初始状态应该是未保护的
    EXPECT_FALSE(protectionManager->isSheetProtected());
    
    // 保护工作表（无密码）
    EXPECT_TRUE(protectionManager->protectSheet());
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    // 取消保护
    EXPECT_TRUE(protectionManager->unprotectSheet());
    EXPECT_FALSE(protectionManager->isSheetProtected());
}

TEST_F(TXSheetProtectionManagerTest, PasswordProtection) {
    const std::string password = "test123";
    
    // 使用密码保护
    EXPECT_TRUE(protectionManager->protectSheet(password));
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    // 验证密码
    EXPECT_TRUE(protectionManager->verifyPassword(password));
    EXPECT_FALSE(protectionManager->verifyPassword("wrong"));
    
    // 使用错误密码取消保护应该失败
    EXPECT_FALSE(protectionManager->unprotectSheet("wrong"));
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    // 使用正确密码取消保护应该成功
    EXPECT_TRUE(protectionManager->unprotectSheet(password));
    EXPECT_FALSE(protectionManager->isSheetProtected());
}

TEST_F(TXSheetProtectionManagerTest, ProtectionOptions) {
    auto protection = TXSheetProtectionManager::SheetProtection::createStrictProtection();
    
    EXPECT_TRUE(protectionManager->protectSheet("", protection));
    
    // 验证保护选项
    const auto& currentProtection = protectionManager->getSheetProtection();
    EXPECT_TRUE(currentProtection.isProtected);
    EXPECT_TRUE(currentProtection.selectLockedCells);
    EXPECT_TRUE(currentProtection.selectUnlockedCells);
    EXPECT_FALSE(currentProtection.formatCells);
    EXPECT_FALSE(currentProtection.insertRows);
    EXPECT_FALSE(currentProtection.deleteRows);
}

TEST_F(TXSheetProtectionManagerTest, LooseProtection) {
    auto protection = TXSheetProtectionManager::SheetProtection::createLooseProtection();
    
    EXPECT_TRUE(protectionManager->protectSheet("", protection));
    
    const auto& currentProtection = protectionManager->getSheetProtection();
    EXPECT_TRUE(currentProtection.isProtected);
    EXPECT_TRUE(currentProtection.formatCells);
    EXPECT_TRUE(currentProtection.formatColumns);
    EXPECT_TRUE(currentProtection.formatRows);
    EXPECT_TRUE(currentProtection.sort);
    EXPECT_TRUE(currentProtection.autoFilter);
}

// ==================== 单元格锁定测试 ====================

TEST_F(TXSheetProtectionManagerTest, CellLocking) {
    TXCoordinate coord(row_t(1), column_t(1));
    
    // 默认情况下单元格应该是锁定的
    EXPECT_TRUE(protectionManager->isCellLocked(coord, *cellManager));
    
    // 解锁单元格
    EXPECT_TRUE(protectionManager->setCellLocked(coord, false, *cellManager));
    EXPECT_FALSE(protectionManager->isCellLocked(coord, *cellManager));
    
    // 重新锁定单元格
    EXPECT_TRUE(protectionManager->setCellLocked(coord, true, *cellManager));
    EXPECT_TRUE(protectionManager->isCellLocked(coord, *cellManager));
}

TEST_F(TXSheetProtectionManagerTest, RangeLocking) {
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    
    // 解锁范围内的所有单元格
    std::size_t count = protectionManager->setRangeLocked(range, false, *cellManager);
    EXPECT_EQ(count, 4);
    
    // 验证所有单元格都被解锁
    EXPECT_FALSE(protectionManager->isCellLocked(TXCoordinate(row_t(1), column_t(1)), *cellManager));
    EXPECT_FALSE(protectionManager->isCellLocked(TXCoordinate(row_t(1), column_t(2)), *cellManager));
    EXPECT_FALSE(protectionManager->isCellLocked(TXCoordinate(row_t(2), column_t(1)), *cellManager));
    EXPECT_FALSE(protectionManager->isCellLocked(TXCoordinate(row_t(2), column_t(2)), *cellManager));
}

TEST_F(TXSheetProtectionManagerTest, BatchCellLocking) {
    std::vector<TXCoordinate> coords = {
        TXCoordinate(row_t(1), column_t(1)),
        TXCoordinate(row_t(1), column_t(2)),
        TXCoordinate(row_t(2), column_t(1))
    };
    
    // 批量解锁
    std::size_t count = protectionManager->setCellsLocked(coords, false, *cellManager);
    EXPECT_EQ(count, 3);
    
    // 验证解锁状态
    for (const auto& coord : coords) {
        EXPECT_FALSE(protectionManager->isCellLocked(coord, *cellManager));
    }
}

// ==================== 权限验证测试 ====================

TEST_F(TXSheetProtectionManagerTest, OperationPermissions) {
    // 未保护时所有操作都应该被允许
    EXPECT_TRUE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::FormatCells));
    EXPECT_TRUE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::InsertRows));
    EXPECT_TRUE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::DeleteColumns));
    
    // 使用严格保护
    auto protection = TXSheetProtectionManager::SheetProtection::createStrictProtection();
    protectionManager->protectSheet("", protection);
    
    // 大部分操作应该被禁止
    EXPECT_FALSE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::FormatCells));
    EXPECT_FALSE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::InsertRows));
    EXPECT_FALSE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::DeleteColumns));
    
    // 但选择操作应该被允许
    EXPECT_TRUE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::SelectLockedCells));
    EXPECT_TRUE(protectionManager->isOperationAllowed(TXSheetProtectionManager::OperationType::SelectUnlockedCells));
}

TEST_F(TXSheetProtectionManagerTest, StringOperationPermissions) {
    auto protection = TXSheetProtectionManager::SheetProtection::createLooseProtection();
    protectionManager->protectSheet("", protection);
    
    // 测试字符串版本的权限检查
    EXPECT_TRUE(protectionManager->isOperationAllowed("formatCells"));
    EXPECT_TRUE(protectionManager->isOperationAllowed("sort"));
    EXPECT_FALSE(protectionManager->isOperationAllowed("insertRows"));
    EXPECT_FALSE(protectionManager->isOperationAllowed("deleteRows"));
}

// ==================== 单元格可编辑性测试 ====================

TEST_F(TXSheetProtectionManagerTest, CellEditability) {
    TXCoordinate coord(row_t(1), column_t(1));
    
    // 未保护时单元格应该可编辑
    EXPECT_TRUE(protectionManager->isCellEditable(coord, *cellManager));
    
    // 保护工作表
    protectionManager->protectSheet();
    
    // 锁定的单元格不可编辑
    EXPECT_FALSE(protectionManager->isCellEditable(coord, *cellManager));
    
    // 解锁单元格后应该可编辑
    protectionManager->setCellLocked(coord, false, *cellManager);
    EXPECT_TRUE(protectionManager->isCellEditable(coord, *cellManager));
}

TEST_F(TXSheetProtectionManagerTest, RangeEditability) {
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    
    // 未保护时范围应该可编辑
    EXPECT_TRUE(protectionManager->isRangeEditable(range, *cellManager));
    
    // 保护工作表
    protectionManager->protectSheet();
    
    // 锁定的范围不可编辑
    EXPECT_FALSE(protectionManager->isRangeEditable(range, *cellManager));
    
    // 解锁部分单元格
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(1)), false, *cellManager);
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(2)), false, *cellManager);
    
    // 范围仍然不可编辑（因为有些单元格仍然锁定）
    EXPECT_FALSE(protectionManager->isRangeEditable(range, *cellManager));
    
    // 解锁所有单元格
    protectionManager->setRangeLocked(range, false, *cellManager);
    
    // 现在范围应该可编辑
    EXPECT_TRUE(protectionManager->isRangeEditable(range, *cellManager));
}

// ==================== 保护状态查询测试 ====================

TEST_F(TXSheetProtectionManagerTest, LockedCellsQuery) {
    // 解锁一些单元格
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(1)), false, *cellManager);
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(2)), false, *cellManager);
    
    // 获取锁定的单元格
    auto lockedCells = protectionManager->getLockedCells(*cellManager);
    EXPECT_EQ(lockedCells.size(), 2);  // A2, B2应该仍然锁定
    
    // 获取未锁定的单元格
    auto unlockedCells = protectionManager->getUnlockedCells(*cellManager);
    EXPECT_EQ(unlockedCells.size(), 2);  // A1, B1应该被解锁
}

TEST_F(TXSheetProtectionManagerTest, ProtectionStats) {
    // 保护工作表
    protectionManager->protectSheet("password");
    
    // 解锁一些单元格
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(1)), false, *cellManager);
    
    auto stats = protectionManager->getProtectionStats(*cellManager);
    
    EXPECT_TRUE(stats.isProtected);
    EXPECT_TRUE(stats.hasPassword);
    EXPECT_EQ(stats.lockedCellCount, 3);
    EXPECT_EQ(stats.unlockedCellCount, 1);
    EXPECT_GT(stats.allowedOperationCount, 0);  // 至少选择操作是允许的

    // 生成测试文件
    addTestInfo(sheet, "ProtectionStats", "测试工作表保护统计功能");

    // 演示保护统计信息
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"保护统计项目"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"值"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"说明"});

    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"工作表保护状态"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{stats.isProtected ? "已保护" : "未保护"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"工作表是否启用保护"});

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"密码保护"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{stats.hasPassword ? "有密码" : "无密码"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"是否设置了保护密码"});

    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"锁定单元格数量"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{static_cast<double>(stats.lockedCellCount)});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{"被锁定的单元格总数"});

    sheet->setCellValue(row_t(11), column_t(1), cell_value_t{"未锁定单元格数量"});
    sheet->setCellValue(row_t(11), column_t(2), cell_value_t{static_cast<double>(stats.unlockedCellCount)});
    sheet->setCellValue(row_t(11), column_t(3), cell_value_t{"未锁定的单元格总数"});

    sheet->setCellValue(row_t(12), column_t(1), cell_value_t{"允许的操作数量"});
    sheet->setCellValue(row_t(12), column_t(2), cell_value_t{static_cast<double>(stats.allowedOperationCount)});
    sheet->setCellValue(row_t(12), column_t(3), cell_value_t{"保护状态下允许的操作数量"});

    // 添加示例数据
    sheet->setCellValue(row_t(14), column_t(1), cell_value_t{"示例数据:"});
    sheet->setCellValue(row_t(15), column_t(1), cell_value_t{"锁定单元格"});
    sheet->setCellValue(row_t(15), column_t(2), cell_value_t{"重要数据"});
    sheet->setCellValue(row_t(15), column_t(3), cell_value_t{"此单元格被锁定"});
    sheet->setCellLocked(row_t(15), column_t(2), true);

    sheet->setCellValue(row_t(16), column_t(1), cell_value_t{"未锁定单元格"});
    sheet->setCellValue(row_t(16), column_t(2), cell_value_t{"可编辑数据"});
    sheet->setCellValue(row_t(16), column_t(3), cell_value_t{"此单元格未锁定"});
    sheet->setCellLocked(row_t(16), column_t(2), false);

    // 保护工作表
    sheet->protectSheet("password");

    saveWorkbook(workbook, "ProtectionStats");
}

// ==================== 清空和重置测试 ====================

TEST_F(TXSheetProtectionManagerTest, ClearAndReset) {
    // 设置保护
    protectionManager->protectSheet("password");
    protectionManager->setCellLocked(TXCoordinate(row_t(1), column_t(1)), false, *cellManager);
    
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    // 清空
    protectionManager->clear();
    
    EXPECT_FALSE(protectionManager->isSheetProtected());
    
    // 重置
    protectionManager->protectSheet("test");
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    protectionManager->reset();
    EXPECT_FALSE(protectionManager->isSheetProtected());
}

// ==================== 边界条件测试 ====================

TEST_F(TXSheetProtectionManagerTest, InvalidOperations) {
    TXCoordinate invalidCoord;  // 无效坐标
    
    // 无效坐标的操作应该失败或返回默认值
    EXPECT_TRUE(protectionManager->isCellLocked(invalidCoord, *cellManager));  // 默认锁定
    
    // 创建真正的无效范围
    TXRange invalidRange(TXCoordinate(row_t(static_cast<row_t::index_t>(0)), column_t(static_cast<column_t::index_t>(0))),
                        TXCoordinate(row_t(static_cast<row_t::index_t>(0)), column_t(static_cast<column_t::index_t>(0))));
    EXPECT_EQ(protectionManager->setRangeLocked(invalidRange, false, *cellManager), 0);
}

TEST_F(TXSheetProtectionManagerTest, EmptyPassword) {
    // 空密码保护
    EXPECT_TRUE(protectionManager->protectSheet(""));
    EXPECT_TRUE(protectionManager->isSheetProtected());
    
    // 空密码验证
    EXPECT_TRUE(protectionManager->verifyPassword(""));
    
    // 空密码取消保护
    EXPECT_TRUE(protectionManager->unprotectSheet(""));
    EXPECT_FALSE(protectionManager->isSheetProtected());
}

// ==================== 预定义保护模板测试 ====================

TEST_F(TXSheetProtectionManagerTest, StrictProtectionTemplate) {
    auto protection = TXSheetProtectionManager::SheetProtection::createStrictProtection();
    
    EXPECT_TRUE(protection.isProtected);
    EXPECT_TRUE(protection.selectLockedCells);
    EXPECT_TRUE(protection.selectUnlockedCells);
    EXPECT_FALSE(protection.formatCells);
    EXPECT_FALSE(protection.formatColumns);
    EXPECT_FALSE(protection.formatRows);
    EXPECT_FALSE(protection.insertColumns);
    EXPECT_FALSE(protection.insertRows);
    EXPECT_FALSE(protection.deleteColumns);
    EXPECT_FALSE(protection.deleteRows);
    EXPECT_FALSE(protection.sort);
    EXPECT_FALSE(protection.autoFilter);
}

TEST_F(TXSheetProtectionManagerTest, LooseProtectionTemplate) {
    auto protection = TXSheetProtectionManager::SheetProtection::createLooseProtection();
    
    EXPECT_TRUE(protection.isProtected);
    EXPECT_TRUE(protection.selectLockedCells);
    EXPECT_TRUE(protection.selectUnlockedCells);
    EXPECT_TRUE(protection.formatCells);
    EXPECT_TRUE(protection.formatColumns);
    EXPECT_TRUE(protection.formatRows);
    EXPECT_FALSE(protection.insertColumns);  // 仍然禁止结构性修改
    EXPECT_FALSE(protection.insertRows);
    EXPECT_FALSE(protection.deleteColumns);
    EXPECT_FALSE(protection.deleteRows);
    EXPECT_TRUE(protection.sort);
    EXPECT_TRUE(protection.autoFilter);
}
