#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXDataFilter.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class DataFilterTest : public TestWithFileGeneration<DataFilterTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<DataFilterTest>::SetUp();
        TinaXlsx::initialize();
    }

    void TearDown() override {
        TinaXlsx::cleanup();
        TestWithFileGeneration<DataFilterTest>::TearDown();
    }
};

TEST_F(DataFilterTest, AutoFilterBasicTest) {
    std::cout << "=== 自动筛选基础测试 ===" << std::endl;

    // 创建工作簿和工作表
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("数据筛选测试");
    ASSERT_NE(sheet, nullptr);

    // 创建测试数据表格
    std::cout << "创建测试数据..." << std::endl;
    
    // 标题行
    sheet->setCellValue(row_t(1), column_t(1), std::string("产品名称"));
    sheet->setCellValue(row_t(1), column_t(2), std::string("价格"));
    sheet->setCellValue(row_t(1), column_t(3), std::string("类别"));
    sheet->setCellValue(row_t(1), column_t(4), std::string("库存"));
    
    // 数据行
    sheet->setCellValue(row_t(2), column_t(1), std::string("笔记本电脑"));
    sheet->setCellValue(row_t(2), column_t(2), 5999.0);
    sheet->setCellValue(row_t(2), column_t(3), std::string("电子产品"));
    sheet->setCellValue(row_t(2), column_t(4), 50);
    
    sheet->setCellValue(row_t(3), column_t(1), std::string("台式机"));
    sheet->setCellValue(row_t(3), column_t(2), 3999.0);
    sheet->setCellValue(row_t(3), column_t(3), std::string("电子产品"));
    sheet->setCellValue(row_t(3), column_t(4), 30);
    
    sheet->setCellValue(row_t(4), column_t(1), std::string("办公椅"));
    sheet->setCellValue(row_t(4), column_t(2), 899.0);
    sheet->setCellValue(row_t(4), column_t(3), std::string("办公用品"));
    sheet->setCellValue(row_t(4), column_t(4), 100);
    
    sheet->setCellValue(row_t(5), column_t(1), std::string("办公桌"));
    sheet->setCellValue(row_t(5), column_t(2), 1299.0);
    sheet->setCellValue(row_t(5), column_t(3), std::string("办公用品"));
    sheet->setCellValue(row_t(5), column_t(4), 80);
    
    sheet->setCellValue(row_t(6), column_t(1), std::string("手机"));
    sheet->setCellValue(row_t(6), column_t(2), 2999.0);
    sheet->setCellValue(row_t(6), column_t(3), std::string("电子产品"));
    sheet->setCellValue(row_t(6), column_t(4), 200);

    // 启用自动筛选
    std::cout << "启用自动筛选..." << std::endl;
    TXRange dataRange = TXRange::fromAddress("A1:D6");
    TXAutoFilter* autoFilter = sheet->enableAutoFilter(dataRange);
    
    ASSERT_NE(autoFilter, nullptr);
    EXPECT_TRUE(sheet->hasAutoFilter());
    EXPECT_EQ(autoFilter->getRange().toAddress(), "A1:D6");
    std::cout << "✅ 自动筛选启用成功，范围: " << autoFilter->getRange().toAddress() << std::endl;

    // 添加筛选条件
    std::cout << "添加筛选条件..." << std::endl;
    
    // 筛选类别为"电子产品"的记录
    autoFilter->setTextFilter(2, "电子产品", FilterOperator::Equal);
    std::cout << "✅ 添加文本筛选: 类别 = 电子产品" << std::endl;
    
    // 筛选价格大于3000的记录
    autoFilter->setNumberFilter(1, 3000, FilterOperator::GreaterThan);
    std::cout << "✅ 添加数值筛选: 价格 > 3000" << std::endl;
    
    // 验证筛选条件
    const auto& conditions = autoFilter->getFilterConditions();
    EXPECT_EQ(conditions.size(), 2);
    std::cout << "筛选条件数量: " << conditions.size() << std::endl;

    // 保存文件
    std::string fullPath = getFilePath("data_filter_test");
    bool saved = workbook.saveToFile(fullPath);
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "📁 文件已保存: " << fullPath << std::endl;
    }

    if (saved) {
        std::cout << "✅ 数据筛选文件生成测试通过" << std::endl;
        std::cout << "生成了包含自动筛选功能的文件" << std::endl;
        std::cout << "🔍 验证方法:" << std::endl;
        std::cout << "- 解压xlsx文件，查看xl/worksheets/sheet1.xml" << std::endl;
        std::cout << "- 应该能看到<autoFilter>节点" << std::endl;
        std::cout << "- 用Excel打开，A1:D6范围应该显示绿色筛选按钮" << std::endl;
        std::cout << "- 数据内容:" << std::endl;
        std::cout << "  * A1:D1: 标题行（产品名称、价格、类别、库存）" << std::endl;
        std::cout << "  * A2:D6: 5行产品数据" << std::endl;
        std::cout << "- 预设筛选条件（需在Excel中手动应用）:" << std::endl;
        std::cout << "  * 类别列(C): 等于 '电子产品'" << std::endl;
        std::cout << "  * 价格列(B): 大于 3000" << std::endl;
        std::cout << "📌 注意: 筛选条件已定义，但需要在Excel中点击筛选按钮来应用" << std::endl;
    }
}

TEST_F(DataFilterTest, AutoFilterAdvancedTest) {
    std::cout << "\n=== 自动筛选高级测试 ===" << std::endl;

    // 创建工作簿和工作表
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("高级筛选测试");
    ASSERT_NE(sheet, nullptr);

    // 创建更复杂的测试数据
    std::cout << "创建复杂测试数据..." << std::endl;
    
    // 标题行
    sheet->setCellValue(row_t(1), column_t(1), std::string("员工姓名"));
    sheet->setCellValue(row_t(1), column_t(2), std::string("部门"));
    sheet->setCellValue(row_t(1), column_t(3), std::string("薪资"));
    sheet->setCellValue(row_t(1), column_t(4), std::string("绩效评级"));
    sheet->setCellValue(row_t(1), column_t(5), std::string("入职年份"));
    
    // 数据行
    const std::vector<std::vector<std::string>> employeeData = {
        {"张三", "技术部", "15000", "A", "2020"},
        {"李四", "销售部", "12000", "B", "2021"},
        {"王五", "技术部", "18000", "A", "2019"},
        {"赵六", "人事部", "10000", "C", "2022"},
        {"钱七", "销售部", "14000", "A", "2020"},
        {"孙八", "技术部", "16000", "B", "2021"},
        {"周九", "财务部", "13000", "B", "2020"},
        {"吴十", "技术部", "20000", "A", "2018"}
    };
    
    for (size_t i = 0; i < employeeData.size(); ++i) {
        const auto& employee = employeeData[i];
        row_t row = row_t(static_cast<u32>(i + 2));  // 从第2行开始
        
        sheet->setCellValue(row, column_t(1), employee[0]);  // 姓名
        sheet->setCellValue(row, column_t(2), employee[1]);  // 部门
        sheet->setCellValue(row, column_t(3), std::stod(employee[2]));  // 薪资（数值）
        sheet->setCellValue(row, column_t(4), employee[3]);  // 绩效评级
        sheet->setCellValue(row, column_t(5), std::stoi(employee[4]));  // 入职年份（数值）
    }

    // 启用自动筛选
    TXRange dataRange = TXRange::fromAddress("A1:E9");
    TXAutoFilter* autoFilter = sheet->enableAutoFilter(dataRange);
    
    ASSERT_NE(autoFilter, nullptr);
    std::cout << "✅ 自动筛选启用成功，范围: " << autoFilter->getRange().toAddress() << std::endl;

    // 添加多个筛选条件
    std::cout << "添加多个筛选条件..." << std::endl;
    
    // 筛选技术部员工
    autoFilter->setTextFilter(1, "技术部", FilterOperator::Equal);
    std::cout << "✅ 添加部门筛选: 技术部" << std::endl;
    
    // 筛选薪资在15000-20000之间
    autoFilter->setRangeFilter(2, 15000, 20000);
    std::cout << "✅ 添加薪资范围筛选: 15000-20000" << std::endl;
    
    // 筛选绩效评级为A的员工
    autoFilter->setTextFilter(3, "A", FilterOperator::Equal);
    std::cout << "✅ 添加绩效筛选: A级" << std::endl;

    // 验证筛选条件
    const auto& conditions = autoFilter->getFilterConditions();
    EXPECT_GE(conditions.size(), 3);  // 范围筛选会产生2个条件，所以至少3个
    std::cout << "筛选条件数量: " << conditions.size() << std::endl;

    // 保存文件
    std::string fullPath = getFilePath("advanced_filter_test");
    bool saved = workbook.saveToFile(fullPath);
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "📁 文件已保存: " << fullPath << std::endl;
    }

    if (saved) {
        std::cout << "✅ 高级数据筛选文件生成测试通过" << std::endl;
        std::cout << "生成了包含多重筛选条件的文件" << std::endl;
        std::cout << "🔍 验证方法:" << std::endl;
        std::cout << "- 用Excel打开，A1:E9范围应该显示绿色筛选按钮" << std::endl;
        std::cout << "- 预设筛选条件（需在Excel中手动应用）:" << std::endl;
        std::cout << "  * 部门列(B): 等于 '技术部'" << std::endl;
        std::cout << "  * 薪资列(C): >= 15000 且 <= 20000" << std::endl;
        std::cout << "  * 绩效列(D): 等于 'A'" << std::endl;
        std::cout << "- 应用筛选后符合条件的员工: 张三、王五、吴十" << std::endl;
        std::cout << "📌 注意: 当前所有数据都可见，需要在Excel中应用筛选条件" << std::endl;
    }
}
