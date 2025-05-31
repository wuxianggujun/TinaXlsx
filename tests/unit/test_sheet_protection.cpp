#include <gtest/gtest.h>
#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <memory>

using namespace TinaXlsx;

class SheetProtectionTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("保护测试");
    }

    void TearDown() override {
        workbook.reset();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(SheetProtectionTest, BasicProtection) {
    // 初始状态应该是未保护的
    EXPECT_FALSE(sheet->isSheetProtected());
    
    // 测试无密码保护
    EXPECT_TRUE(sheet->protectSheet());
    EXPECT_TRUE(sheet->isSheetProtected());
    
    // 测试无密码解除保护
    EXPECT_TRUE(sheet->unprotectSheet());
    EXPECT_FALSE(sheet->isSheetProtected());
}

TEST_F(SheetProtectionTest, PasswordProtection) {
    const std::string password = "test123";
    
    // 测试密码保护
    EXPECT_TRUE(sheet->protectSheet(password));
    EXPECT_TRUE(sheet->isSheetProtected());
    
    // 测试错误密码解除保护
    EXPECT_FALSE(sheet->unprotectSheet("wrongpassword"));
    EXPECT_TRUE(sheet->isSheetProtected());
    
    // 测试正确密码解除保护
    EXPECT_TRUE(sheet->unprotectSheet(password));
    EXPECT_FALSE(sheet->isSheetProtected());
}

TEST_F(SheetProtectionTest, CustomProtectionOptions) {
    TXSheet::SheetProtection protection;
    protection.formatCells = false;
    protection.insertRows = false;
    protection.deleteRows = false;
    protection.selectLockedCells = true;
    protection.selectUnlockedCells = true;
    
    // 应用自定义保护选项
    EXPECT_TRUE(sheet->protectSheet("password", protection));
    EXPECT_TRUE(sheet->isSheetProtected());
    
    // 验证保护选项
    const auto& retrievedProtection = sheet->getSheetProtection();
    EXPECT_TRUE(retrievedProtection.isProtected);
    EXPECT_FALSE(retrievedProtection.formatCells);
    EXPECT_FALSE(retrievedProtection.insertRows);
    EXPECT_FALSE(retrievedProtection.deleteRows);
    EXPECT_TRUE(retrievedProtection.selectLockedCells);
    EXPECT_TRUE(retrievedProtection.selectUnlockedCells);
}

TEST_F(SheetProtectionTest, CellLocking) {
    // 创建单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    // 测试默认锁定状态（新单元格默认应该是锁定的）
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 测试解锁单元格
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), false));
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 测试重新锁定单元格
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), true));
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(SheetProtectionTest, RangeLocking) {
    // 创建范围数据
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(3), column_t(3)));
    
    // 在范围内设置一些数据
    for (int r = 1; r <= 3; ++r) {
        for (int c = 1; c <= 3; ++c) {
            sheet->setCellValue(row_t(r), column_t(c), cell_value_t{"数据"});
        }
    }
    
    // 测试范围解锁
    std::size_t unlockedCount = sheet->setRangeLocked(range, false);
    EXPECT_EQ(unlockedCount, 9); // 3x3 = 9个单元格
    
    // 验证范围内的单元格都被解锁
    for (int r = 1; r <= 3; ++r) {
        for (int c = 1; c <= 3; ++c) {
            EXPECT_FALSE(sheet->isCellLocked(row_t(r), column_t(c)));
        }
    }
    
    // 测试范围重新锁定
    std::size_t lockedCount = sheet->setRangeLocked(range, true);
    EXPECT_EQ(lockedCount, 9);
    
    // 验证范围内的单元格都被锁定
    for (int r = 1; r <= 3; ++r) {
        for (int c = 1; c <= 3; ++c) {
            EXPECT_TRUE(sheet->isCellLocked(row_t(r), column_t(c)));
        }
    }
}