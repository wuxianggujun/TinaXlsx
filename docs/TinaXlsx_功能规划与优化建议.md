# TinaXlsx 功能规划与优化建议

## 📊 当前状态
- ✅ 基础Excel文件读写功能
- ✅ 工作表和单元格操作
- ✅ 基础图表功能（柱状图、折线图、饼图、散点图）
- ✅ 工作表和工作簿保护
- ✅ 文件加密功能

## 🔧 代码优化建议

### 1. 图表XML生成器重构

#### 当前问题
- `TXChartXmlHandler` 职责过重，承担了所有图表类型的XML生成
- 数据范围计算逻辑重复
- 图表样式硬编码

#### 重构方案
```cpp
// 基础图表XML生成器
class TXChartXmlHandler;

// 专门的系列数据处理器
class TXChartSeriesBuilder {
    virtual XmlNodeBuilder buildSeries(const TXChart* chart) = 0;
};

// 不同图表类型的系列构建器
class TXColumnSeriesBuilder : public TXChartSeriesBuilder;
class TXLineSeriesBuilder : public TXChartSeriesBuilder;
class TXPieSeriesBuilder : public TXChartSeriesBuilder;
class TXScatterSeriesBuilder : public TXChartSeriesBuilder;

// 坐标轴构建器
class TXAxisBuilder {
    XmlNodeBuilder buildCategoryAxis();
    XmlNodeBuilder buildValueAxis();
};

// 数据范围格式化工具
class TXRangeFormatter {
    static std::string formatCategoryRange(const TXRange& range, const std::string& sheetName);
    static std::string formatValueRange(const TXRange& range, const std::string& sheetName);
    static std::string formatScatterXRange(const TXRange& range, const std::string& sheetName);
    static std::string formatScatterYRange(const TXRange& range, const std::string& sheetName);
};

// 图表样式配置
class TXChartStyle {
    std::string primaryColor = "4F81BD";
    std::string secondaryColor = "F79646";
    int lineWidth = 25400;
};
```

### 2. 架构优化

#### 图表工厂模式
```cpp
class TXChartFactory {
    static std::unique_ptr<TXChart> createChart(ChartType type, const std::string& title);
    static std::unique_ptr<TXChartXmlHandler> createXmlHandler(const TXChart* chart, u32 index);
};
```

#### 图表配置分离
```cpp
class TXChartConfig {
    bool showLegend = true;
    bool showDataLabels = false;
    bool showGridlines = true;
    TXChartStyle style;
};
```

## 🚀 功能开发路线图

### 🎯 优先级1：图表增强功能

#### 1.1 图表样式和主题
```cpp
enum class ChartTheme {
    Office,
    Colorful,
    Monochromatic,
    Custom
};

chart->setTheme(ChartTheme::Colorful);
chart->setColors({"#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4"});
```

#### 1.2 图表标题和标签增强
```cpp
chart->setTitle("销售报表", TXFont("Arial", 16, true));
chart->setAxisTitle("月份", true);  // X轴
chart->setAxisTitle("销售额(万元)", false); // Y轴
chart->setDataLabelFormat("{value:F2}万元");
```

#### 1.3 多系列图表支持
```cpp
auto* chart = sheet->addColumnChart("销售对比", dataRange);
chart->addSeries("销售额", "B2:B7", "#4F81BD");
chart->addSeries("利润", "C2:C7", "#F79646");
chart->addSeries("成本", "D2:D7", "#9CBB58");
```

### 🎯 优先级2：数据处理增强

#### 2.1 公式和函数支持增强
```cpp
// 统计函数
sheet->setCellFormula(row_t(10), column_t(2), "=AVERAGE(B2:B7)");
sheet->setCellFormula(row_t(11), column_t(2), "=MAX(B2:B7)");
sheet->setCellFormula(row_t(12), column_t(2), "=MIN(B2:B7)");

// 条件函数
sheet->setCellFormula(row_t(13), column_t(2), "=SUMIF(A2:A7,\"一月\",B2:B7)");
sheet->setCellFormula(row_t(14), column_t(2), "=COUNTIF(B2:B7,\">1000\")");
```

#### 2.2 数据验证
```cpp
TXDataValidation validation;
validation.setType(TXDataValidation::Type::List);
validation.setFormula1("优秀,良好,一般,差");
validation.setShowDropDown(true);
validation.setErrorMessage("请选择有效的评级");
sheet->setDataValidation("D2:D10", validation);
```

#### 2.3 条件格式
```cpp
TXConditionalFormat format;
format.setType(TXConditionalFormat::Type::CellValue);
format.setOperator(TXConditionalFormat::Operator::GreaterThan);
format.setValue(1000);
format.setFormat(TXCellFormat().setBackgroundColor(0xFF90EE90));
sheet->addConditionalFormat("B2:B7", format);
```

### 🎯 优先级3：高级功能

#### 3.1 数据透视表
```cpp
auto* pivotTable = sheet->addPivotTable("数据透视表", "A1:D100");
pivotTable->addRowField("产品类别");
pivotTable->addColumnField("月份");
pivotTable->addDataField("销售额", TXPivotTable::Function::Sum);
pivotTable->addDataField("利润", TXPivotTable::Function::Average);
```

#### 3.2 图片和形状
```cpp
// 插入图片
sheet->insertImage("logo.png", "A1", 100, 50);

// 插入形状
auto* shape = sheet->addShape(TXShape::Type::Rectangle, "B5:D8");
shape->setText("重要提示");
shape->setFillColor(0xFFFFFF00);
shape->setBorderColor(0xFF000000);
```

#### 3.3 工作表保护增强
```cpp
TXSheetProtection protection;
protection.setPassword("123456");
protection.setAllowSelectLockedCells(true);
protection.setAllowSelectUnlockedCells(true);
protection.setAllowFormatCells(false);
protection.setAllowInsertRows(false);
protection.setAllowDeleteRows(false);
protection.setAllowSort(true);
protection.setAllowFilter(true);
sheet->setProtection(protection);
```

## 📋 实施计划

### 第一阶段：代码重构（当前）
1. 重构图表XML生成器
2. 提取系列构建器
3. 创建范围格式化工具
4. 实现图表样式配置

### 第二阶段：图表增强
1. 实现图表主题系统
2. 支持多系列图表
3. 增强标题和标签功能
4. 添加图表动画支持

### 第三阶段：数据功能
1. 扩展公式函数库
2. 实现数据验证
3. 添加条件格式
4. 支持数据排序和筛选

### 第四阶段：高级功能
1. 数据透视表
2. 图片和形状插入
3. 宏和VBA支持（可选）
4. 协作功能（批注、修订）

## 🎯 性能优化目标

1. **内存使用**：大文件处理时内存占用 < 500MB
2. **处理速度**：10万行数据处理时间 < 5秒
3. **文件大小**：生成的XLSX文件压缩率 > 80%
4. **兼容性**：与Excel 2016+、WPS、LibreOffice 100%兼容

## 📊 质量指标

1. **代码覆盖率**：> 90%
2. **单元测试**：每个功能模块都有对应测试
3. **文档完整性**：API文档覆盖率 > 95%
4. **示例代码**：每个主要功能都有使用示例

---

**最后更新**：2024年12月
**状态**：图表基础功能已完成，准备进入重构阶段
