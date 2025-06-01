#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXSha512.hpp"
#include "test_file_generator.hpp"
#include <memory>
#include <iomanip>

using namespace TinaXlsx;

class CellLockingTest : public TestWithFileGeneration<CellLockingTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<CellLockingTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("锁定测试");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<CellLockingTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(CellLockingTest, DefaultLockingState) {
    // 创建新单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    TXCell* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    // 新单元格默认应该是锁定的
    EXPECT_TRUE(cell->isLocked());
    
    // 通过工作表接口也应该返回锁定状态
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(CellLockingTest, SetCellLocking) {
    // 创建单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    TXCell* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    // 测试解锁
    cell->setLocked(false);
    EXPECT_FALSE(cell->isLocked());
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 测试重新锁定
    cell->setLocked(true);
    EXPECT_TRUE(cell->isLocked());
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(CellLockingTest, SetCellLockingViaSheet) {
    // ==================== 第一个工作表：基本锁定测试 ====================
    sheet->setName("基本锁定测试");

    // 添加标题行
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"工作表1: 基本单元格锁定功能测试"});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{"测试日期: 2024-01-15"});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{"保护密码: test123"});

    // 调试：显示密码哈希信息
    auto& protectionManager = sheet->getProtectionManager();
    std::cout << "\n=== 密码哈希调试信息 ===" << std::endl;

    // 测试Base64编码/解码
    std::cout << "\n--- Base64测试 ---" << std::endl;
    std::string testData = "Hello World";
    std::vector<uint8_t> testBytes(testData.begin(), testData.end());
    std::string encoded = TinaXlsx::TXBase64::encode(testBytes);
    std::vector<uint8_t> decoded = TinaXlsx::TXBase64::decode(encoded);
    std::string decodedStr(decoded.begin(), decoded.end());
    std::cout << "原始: " << testData << std::endl;
    std::cout << "编码: " << encoded << std::endl;
    std::cout << "解码: " << decodedStr << std::endl;
    std::cout << "Base64测试: " << (testData == decodedStr ? "通过" : "失败") << std::endl;

    // 测试UTF-16编码
    std::cout << "\n--- UTF-16编码测试 ---" << std::endl;
    std::string testPassword = "test";
    auto utf16Bytes = TinaXlsx::TXExcelPasswordHash::passwordToUtf16(testPassword);
    std::cout << "密码: " << testPassword << std::endl;
    std::cout << "UTF-16字节: ";
    for (auto byte : utf16Bytes) {
        std::cout << std::hex << std::setw(2) << std::setfill('0') << static_cast<int>(byte) << " ";
    }
    std::cout << std::dec << std::endl;

    // 测试多个密码的哈希值
    std::cout << "\n--- 密码哈希测试 ---" << std::endl;
    std::vector<std::string> testPasswords = {"test", "test123", "password", "123456", "abc"};
    for (const auto& pwd : testPasswords) {
        // 创建临时保护管理器来测试哈希
        TXSheetProtectionManager tempManager;
        tempManager.protectSheet(pwd);
        const auto& tempProtection = tempManager.getSheetProtection();
        std::cout << "密码: '" << pwd << "' -> 哈希: " << tempProtection.passwordHash << std::endl;
    }

    // 添加测试数据
    sheet->setCellValue(row_t(5), column_t(1), cell_value_t{"单元格"});
    sheet->setCellValue(row_t(5), column_t(2), cell_value_t{"锁定状态"});
    sheet->setCellValue(row_t(5), column_t(3), cell_value_t{"内容"});
    sheet->setCellValue(row_t(5), column_t(4), cell_value_t{"说明"});

    // 锁定的单元格
    sheet->setCellValue(row_t(6), column_t(1), cell_value_t{"C6"});
    sheet->setCellValue(row_t(6), column_t(2), cell_value_t{"锁定"});
    sheet->setCellValue(row_t(6), column_t(3), cell_value_t{"重要数据"});
    sheet->setCellValue(row_t(6), column_t(4), cell_value_t{"此单元格被锁定，保护时无法编辑"});
    sheet->setCellLocked(row_t(6), column_t(3), true);

    // 未锁定的单元格
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"C7"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"可编辑数据"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"此单元格未锁定，保护时仍可编辑"});
    sheet->setCellLocked(row_t(7), column_t(3), false);

    // 混合状态的行
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"C8"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"锁定"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"标题"});
    sheet->setCellValue(row_t(8), column_t(4), cell_value_t{"标题通常需要锁定"});
    sheet->setCellLocked(row_t(8), column_t(3), true);

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"C9"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"输入区域"});
    sheet->setCellValue(row_t(9), column_t(4), cell_value_t{"输入区域通常不锁定"});
    sheet->setCellLocked(row_t(9), column_t(3), false);

    // 保护工作表以使锁定生效
    std::cout << "保护工作表前..." << std::endl;
    bool protectResult = sheet->protectSheet("test123");
    std::cout << "保护工作表结果: " << (protectResult ? "成功" : "失败") << std::endl;

    // 获取密码哈希信息
    const auto& protection = protectionManager.getSheetProtection();
    std::cout << "密码哈希: " << protection.passwordHash << std::endl;
    std::cout << "工作表保护状态: " << (protection.isProtected ? "已保护" : "未保护") << std::endl;
    std::cout << "=== 密码哈希调试信息结束 ===" << std::endl;

    saveWorkbook(workbook, "PasswordHashTest");
}

TEST_F(CellLockingTest, AutoCreateCellForLocking) {
    // 对不存在的单元格设置锁定状态应该自动创建单元格
    EXPECT_TRUE(sheet->setCellLocked(row_t(5), column_t(5), false));
    
    // 验证单元格被创建并且锁定状态正确
    TXCell* cell = sheet->getCell(row_t(5), column_t(5));
    ASSERT_NE(cell, nullptr);
    EXPECT_FALSE(cell->isLocked());
}

TEST_F(CellLockingTest, NonExistentCellDefaultLocked) {
    // 不存在的单元格应该返回默认锁定状态
    EXPECT_TRUE(sheet->isCellLocked(row_t(10), column_t(10)));
}

TEST_F(CellLockingTest, CellCopyPreservesLocking) {
    // 创建原始单元格并设置为解锁
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"原始数据"});
    TXCell* originalCell = sheet->getCell(row_t(1), column_t(1));
    originalCell->setLocked(false);
    
    // 测试拷贝构造
    TXCell copiedCell(*originalCell);
    EXPECT_FALSE(copiedCell.isLocked());
    EXPECT_EQ(copiedCell.getStringValue(), "原始数据");
    
    // 测试拷贝赋值
    TXCell assignedCell;
    assignedCell = *originalCell;
    EXPECT_FALSE(assignedCell.isLocked());
    EXPECT_EQ(assignedCell.getStringValue(), "原始数据");
}

TEST_F(CellLockingTest, CellMovePreservesLocking) {
    // 创建原始单元格并设置为解锁
    TXCell originalCell(cell_value_t{"移动数据"});
    originalCell.setLocked(false);
    
    // 测试移动构造
    TXCell movedCell(std::move(originalCell));
    EXPECT_FALSE(movedCell.isLocked());
    EXPECT_EQ(movedCell.getStringValue(), "移动数据");
    
    // 测试移动赋值
    TXCell anotherCell(cell_value_t{"另一个数据"});
    anotherCell.setLocked(false);
    
    TXCell assignedCell;
    assignedCell = std::move(anotherCell);
    EXPECT_FALSE(assignedCell.isLocked());
    EXPECT_EQ(assignedCell.getStringValue(), "另一个数据");
}

TEST_F(CellLockingTest, HasFormulaMethod) {
    // 创建普通单元格
    TXCell normalCell(cell_value_t{"普通数据"});
    EXPECT_FALSE(normalCell.hasFormula());

    // 创建公式单元格
    TXCell formulaCell;
    formulaCell.setFormula("A1+B1");
    EXPECT_TRUE(formulaCell.hasFormula());

    // 清除公式
    formulaCell.setFormula("");
    EXPECT_FALSE(formulaCell.hasFormula());
}

TEST_F(CellLockingTest, MultiSheetProtectionTest) {
    // ==================== 第一个工作表：基本锁定测试 ====================
    sheet->setName("基本锁定测试");

    // 添加标题行
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"工作表1: 基本单元格锁定功能测试"});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{"测试日期: 2024-01-15"});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{"保护密码: test123"});

    // 添加测试数据
    sheet->setCellValue(row_t(5), column_t(1), cell_value_t{"单元格"});
    sheet->setCellValue(row_t(5), column_t(2), cell_value_t{"锁定状态"});
    sheet->setCellValue(row_t(5), column_t(3), cell_value_t{"内容"});
    sheet->setCellValue(row_t(5), column_t(4), cell_value_t{"说明"});

    // 锁定的单元格
    sheet->setCellValue(row_t(6), column_t(1), cell_value_t{"C6"});
    sheet->setCellValue(row_t(6), column_t(2), cell_value_t{"锁定"});
    sheet->setCellValue(row_t(6), column_t(3), cell_value_t{"重要数据"});
    sheet->setCellValue(row_t(6), column_t(4), cell_value_t{"此单元格被锁定，保护时无法编辑"});
    sheet->setCellLocked(row_t(6), column_t(3), true);

    // 未锁定的单元格
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"C7"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"可编辑数据"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"此单元格未锁定，保护时仍可编辑"});
    sheet->setCellLocked(row_t(7), column_t(3), false);

    // 保护工作表以使锁定生效
    sheet->protectSheet("test123");

    // ==================== 第二个工作表：不同密码的保护测试 ====================
    auto* sheet2 = workbook->addSheet("不同密码保护");

    sheet2->setCellValue(row_t(1), column_t(1), cell_value_t{"工作表2: 不同密码保护测试"});
    sheet2->setCellValue(row_t(2), column_t(1), cell_value_t{"保护密码: password456"});
    sheet2->setCellValue(row_t(3), column_t(1), cell_value_t{"测试目的: 验证不同工作表可以使用不同密码"});

    // 添加测试数据
    sheet2->setCellValue(row_t(5), column_t(1), cell_value_t{"数据类型"});
    sheet2->setCellValue(row_t(5), column_t(2), cell_value_t{"锁定状态"});
    sheet2->setCellValue(row_t(5), column_t(3), cell_value_t{"值"});

    // 财务数据（锁定）
    sheet2->setCellValue(row_t(6), column_t(1), cell_value_t{"收入"});
    sheet2->setCellValue(row_t(6), column_t(2), cell_value_t{"锁定"});
    sheet2->setCellValue(row_t(6), column_t(3), cell_value_t{"100000"});
    sheet2->setCellLocked(row_t(6), column_t(3), true);

    sheet2->setCellValue(row_t(7), column_t(1), cell_value_t{"支出"});
    sheet2->setCellValue(row_t(7), column_t(2), cell_value_t{"锁定"});
    sheet2->setCellValue(row_t(7), column_t(3), cell_value_t{"75000"});
    sheet2->setCellLocked(row_t(7), column_t(3), true);

    // 备注区域（未锁定）
    sheet2->setCellValue(row_t(8), column_t(1), cell_value_t{"备注"});
    sheet2->setCellValue(row_t(8), column_t(2), cell_value_t{"未锁定"});
    sheet2->setCellValue(row_t(8), column_t(3), cell_value_t{"可以修改此备注"});
    sheet2->setCellLocked(row_t(8), column_t(3), false);

    // 保护工作表
    sheet2->protectSheet("password456");

    // ==================== 第三个工作表：无密码保护测试 ====================
    auto* sheet3 = workbook->addSheet("无密码保护");

    sheet3->setCellValue(row_t(1), column_t(1), cell_value_t{"工作表3: 无密码保护测试"});
    sheet3->setCellValue(row_t(2), column_t(1), cell_value_t{"保护密码: 无"});
    sheet3->setCellValue(row_t(3), column_t(1), cell_value_t{"测试目的: 验证无密码保护功能"});

    // 添加测试数据
    sheet3->setCellValue(row_t(5), column_t(1), cell_value_t{"配置项"});
    sheet3->setCellValue(row_t(5), column_t(2), cell_value_t{"锁定状态"});
    sheet3->setCellValue(row_t(5), column_t(3), cell_value_t{"值"});

    // 系统配置（锁定）
    sheet3->setCellValue(row_t(6), column_t(1), cell_value_t{"系统版本"});
    sheet3->setCellValue(row_t(6), column_t(2), cell_value_t{"锁定"});
    sheet3->setCellValue(row_t(6), column_t(3), cell_value_t{"v1.0.0"});
    sheet3->setCellLocked(row_t(6), column_t(3), true);

    // 用户设置（未锁定）
    sheet3->setCellValue(row_t(7), column_t(1), cell_value_t{"用户名"});
    sheet3->setCellValue(row_t(7), column_t(2), cell_value_t{"未锁定"});
    sheet3->setCellValue(row_t(7), column_t(3), cell_value_t{"admin"});
    sheet3->setCellLocked(row_t(7), column_t(3), false);

    // 无密码保护
    sheet3->protectSheet("");

    // ==================== 第四个工作表：未保护测试 ====================
    auto* sheet4 = workbook->addSheet("未保护工作表");

    sheet4->setCellValue(row_t(1), column_t(1), cell_value_t{"工作表4: 未保护测试"});
    sheet4->setCellValue(row_t(2), column_t(1), cell_value_t{"保护状态: 未保护"});
    sheet4->setCellValue(row_t(3), column_t(1), cell_value_t{"测试目的: 验证未保护工作表中锁定设置不生效"});

    // 添加测试数据
    sheet4->setCellValue(row_t(5), column_t(1), cell_value_t{"数据项"});
    sheet4->setCellValue(row_t(5), column_t(2), cell_value_t{"锁定设置"});
    sheet4->setCellValue(row_t(5), column_t(3), cell_value_t{"值"});
    sheet4->setCellValue(row_t(5), column_t(4), cell_value_t{"实际效果"});

    // 设置了锁定但工作表未保护
    sheet4->setCellValue(row_t(6), column_t(1), cell_value_t{"测试数据1"});
    sheet4->setCellValue(row_t(6), column_t(2), cell_value_t{"设置为锁定"});
    sheet4->setCellValue(row_t(6), column_t(3), cell_value_t{"数据1"});
    sheet4->setCellValue(row_t(6), column_t(4), cell_value_t{"仍可编辑（工作表未保护）"});
    sheet4->setCellLocked(row_t(6), column_t(3), true);

    sheet4->setCellValue(row_t(7), column_t(1), cell_value_t{"测试数据2"});
    sheet4->setCellValue(row_t(7), column_t(2), cell_value_t{"设置为未锁定"});
    sheet4->setCellValue(row_t(7), column_t(3), cell_value_t{"数据2"});
    sheet4->setCellValue(row_t(7), column_t(4), cell_value_t{"可编辑（工作表未保护）"});
    sheet4->setCellLocked(row_t(7), column_t(3), false);

    // 注意：这个工作表故意不调用protectSheet()

    // ==================== 第五个工作表：总结测试 ====================
    auto* sheet5 = workbook->addSheet("测试总结");

    sheet5->setCellValue(row_t(1), column_t(1), cell_value_t{"多工作表保护功能测试总结"});
    sheet5->setCellValue(row_t(2), column_t(1), cell_value_t{"测试日期: 2024-01-15"});

    // 添加测试总结
    sheet5->setCellValue(row_t(4), column_t(1), cell_value_t{"工作表名称"});
    sheet5->setCellValue(row_t(4), column_t(2), cell_value_t{"保护状态"});
    sheet5->setCellValue(row_t(4), column_t(3), cell_value_t{"密码"});
    sheet5->setCellValue(row_t(4), column_t(4), cell_value_t{"测试目的"});

    sheet5->setCellValue(row_t(5), column_t(1), cell_value_t{"基本锁定测试"});
    sheet5->setCellValue(row_t(5), column_t(2), cell_value_t{"已保护"});
    sheet5->setCellValue(row_t(5), column_t(3), cell_value_t{"test123"});
    sheet5->setCellValue(row_t(5), column_t(4), cell_value_t{"基本锁定功能验证"});

    sheet5->setCellValue(row_t(6), column_t(1), cell_value_t{"不同密码保护"});
    sheet5->setCellValue(row_t(6), column_t(2), cell_value_t{"已保护"});
    sheet5->setCellValue(row_t(6), column_t(3), cell_value_t{"password456"});
    sheet5->setCellValue(row_t(6), column_t(4), cell_value_t{"不同密码验证"});

    sheet5->setCellValue(row_t(7), column_t(1), cell_value_t{"无密码保护"});
    sheet5->setCellValue(row_t(7), column_t(2), cell_value_t{"已保护"});
    sheet5->setCellValue(row_t(7), column_t(3), cell_value_t{"无"});
    sheet5->setCellValue(row_t(7), column_t(4), cell_value_t{"无密码保护验证"});

    sheet5->setCellValue(row_t(8), column_t(1), cell_value_t{"未保护工作表"});
    sheet5->setCellValue(row_t(8), column_t(2), cell_value_t{"未保护"});
    sheet5->setCellValue(row_t(8), column_t(3), cell_value_t{"无"});
    sheet5->setCellValue(row_t(8), column_t(4), cell_value_t{"未保护状态验证"});

    sheet5->setCellValue(row_t(9), column_t(1), cell_value_t{"测试总结"});
    sheet5->setCellValue(row_t(9), column_t(2), cell_value_t{"未保护"});
    sheet5->setCellValue(row_t(9), column_t(3), cell_value_t{"无"});
    sheet5->setCellValue(row_t(9), column_t(4), cell_value_t{"总结页面"});

    // 这个工作表不保护，方便查看总结信息

    saveWorkbook(workbook, "MultiSheetProtectionTest");
}