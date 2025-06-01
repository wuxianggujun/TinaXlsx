#include <gtest/gtest.h>
#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <memory>

using namespace TinaXlsx;

class EnhancedFormulasTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("公式测试");
    }

    void TearDown() override {
        workbook.reset();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(EnhancedFormulasTest, FormulaCalculationOptions) {
    // 测试默认选项
    const auto& defaultOptions = sheet->getFormulaCalculationOptions();
    EXPECT_TRUE(defaultOptions.autoCalculate);
    EXPECT_FALSE(defaultOptions.iterativeCalculation);
    EXPECT_EQ(defaultOptions.maxIterations, 100);
    EXPECT_DOUBLE_EQ(defaultOptions.maxChange, 0.001);
    EXPECT_FALSE(defaultOptions.precisionAsDisplayed);
    EXPECT_FALSE(defaultOptions.use1904DateSystem);
    
    // 测试自定义选项
    TXSheet::FormulaCalculationOptions customOptions;
    customOptions.autoCalculate = false;
    customOptions.iterativeCalculation = true;
    customOptions.maxIterations = 200;
    customOptions.maxChange = 0.0001;
    customOptions.precisionAsDisplayed = true;
    customOptions.use1904DateSystem = true;
    
    sheet->setFormulaCalculationOptions(customOptions);
    
    const auto& retrievedOptions = sheet->getFormulaCalculationOptions();
    EXPECT_FALSE(retrievedOptions.autoCalculate);
    EXPECT_TRUE(retrievedOptions.iterativeCalculation);
    EXPECT_EQ(retrievedOptions.maxIterations, 200);
    EXPECT_DOUBLE_EQ(retrievedOptions.maxChange, 0.0001);
    EXPECT_TRUE(retrievedOptions.precisionAsDisplayed);
    EXPECT_TRUE(retrievedOptions.use1904DateSystem);
}

TEST_F(EnhancedFormulasTest, NamedRanges) {
    // 创建测试范围
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(3), column_t(3)));
    
    // 测试添加命名范围
    EXPECT_TRUE(sheet->addNamedRange("测试范围", range, "这是一个测试范围"));
    
    // 测试获取命名范围
    TXRange retrievedRange = sheet->getNamedRange("测试范围");
    EXPECT_TRUE(retrievedRange.isValid());
    EXPECT_EQ(retrievedRange.getStart().getRow().index(), 1);
    EXPECT_EQ(retrievedRange.getStart().getCol().index(), 1);
    EXPECT_EQ(retrievedRange.getEnd().getRow().index(), 3);
    EXPECT_EQ(retrievedRange.getEnd().getCol().index(), 3);
    
    // 测试获取不存在的命名范围
    TXRange nonExistentRange = sheet->getNamedRange("不存在的范围");
    EXPECT_FALSE(nonExistentRange.isValid());
    
    // 测试获取所有命名范围
    auto allRanges = sheet->getAllNamedRanges();
    EXPECT_EQ(allRanges.size(), 1);
    EXPECT_TRUE(allRanges.find("测试范围") != allRanges.end());
    
    // 测试删除命名范围
    EXPECT_TRUE(sheet->removeNamedRange("测试范围"));
    EXPECT_FALSE(sheet->removeNamedRange("测试范围")); // 再次删除应该失败
    
    allRanges = sheet->getAllNamedRanges();
    EXPECT_EQ(allRanges.size(), 0);
}

TEST_F(EnhancedFormulasTest, CircularReferenceDetection) {
    // 设置一些正常的公式
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{10.0});
    sheet->setCellValue(row_t(1), column_t(2), cell_value_t{20.0});
    sheet->setCellFormula(row_t(1), column_t(3), "A1+B1");
    
    // 此时应该没有循环引用
    EXPECT_FALSE(sheet->detectCircularReferences());
    
    // 创建循环引用：A2引用B2，B2引用A2
    sheet->setCellFormula(row_t(2), column_t(1), "=B2+1");  // A2 = B2+1
    sheet->setCellFormula(row_t(2), column_t(2), "=A2+1");  // B2 = A2+1
    
    // 现在应该检测到循环引用
    EXPECT_TRUE(sheet->detectCircularReferences());
}

TEST_F(EnhancedFormulasTest, FormulaDependencies) {
    // 设置一些有依赖关系的公式
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{10.0});
    sheet->setCellValue(row_t(1), column_t(2), cell_value_t{20.0});
    sheet->setCellFormula(row_t(1), column_t(3), "=A1+B1");  // C1 = A1+B1
    sheet->setCellFormula(row_t(2), column_t(1), "=C1*2");   // A2 = C1*2
    
    // 获取公式依赖关系
    auto dependencies = sheet->getFormulaDependencies();
    
    // 验证依赖关系
    EXPECT_GT(dependencies.size(), 0);
    
    // C1 (1,3) 应该依赖于 A1 (1,1) 和 B1 (1,2)
    auto c1Coord = TXCoordinate(row_t(1), column_t(3));
    if (dependencies.find(c1Coord) != dependencies.end()) {
        const auto& c1Deps = dependencies[c1Coord];
        EXPECT_GE(c1Deps.size(), 2); // 至少依赖两个单元格
    }
}

TEST_F(EnhancedFormulasTest, InvalidNamedRanges) {
    // 测试无效的命名范围
    TXRange invalidRange(TXCoordinate(row_t(static_cast<row_t::index_t>(0)), column_t(static_cast<column_t::index_t>(0))),
                        TXCoordinate(row_t(static_cast<row_t::index_t>(0)), column_t(static_cast<column_t::index_t>(0)))); // 真正的无效范围
    EXPECT_FALSE(sheet->addNamedRange("", invalidRange)); // 空名称
    EXPECT_FALSE(sheet->addNamedRange("测试", invalidRange)); // 无效范围
    
    // 测试空名称
    TXRange validRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    EXPECT_FALSE(sheet->addNamedRange("", validRange));
}

TEST_F(EnhancedFormulasTest, MultipleNamedRanges) {
    // 添加多个命名范围
    TXRange range1(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(3), column_t(3)));
    TXRange range2(TXCoordinate(row_t(5), column_t(1)), TXCoordinate(row_t(7), column_t(3)));
    TXRange range3(TXCoordinate(row_t(1), column_t(5)), TXCoordinate(row_t(3), column_t(7)));
    
    EXPECT_TRUE(sheet->addNamedRange("范围1", range1));
    EXPECT_TRUE(sheet->addNamedRange("范围2", range2));
    EXPECT_TRUE(sheet->addNamedRange("范围3", range3));
    
    // 验证所有范围都被添加
    auto allRanges = sheet->getAllNamedRanges();
    EXPECT_EQ(allRanges.size(), 3);
    
    // 验证每个范围都可以正确获取
    EXPECT_TRUE(sheet->getNamedRange("范围1").isValid());
    EXPECT_TRUE(sheet->getNamedRange("范围2").isValid());
    EXPECT_TRUE(sheet->getNamedRange("范围3").isValid());
    
    // 删除中间的范围
    EXPECT_TRUE(sheet->removeNamedRange("范围2"));
    allRanges = sheet->getAllNamedRanges();
    EXPECT_EQ(allRanges.size(), 2);
    EXPECT_FALSE(sheet->getNamedRange("范围2").isValid());
}