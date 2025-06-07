# TinaXlsx 用户层API设计文档

## 🏗️ 架构概览

TinaXlsx采用分层架构设计，为用户提供简洁易用的接口，同时保持底层的高性能实现。

```
┌─────────────────────────────────────────────────────────────┐
│                    用户层 (User Layer)                      │
├─────────────────────────────────────────────────────────────┤
│  TXWorkbook  │  TXSheet  │  TXRange  │  TXCell  │  TXStyle  │
├─────────────────────────────────────────────────────────────┤
│                   底层实现 (Implementation)                  │
├─────────────────────────────────────────────────────────────┤
│ TXInMemorySheet │ TXSheetAPI │ TXBatchSIMDProcessor │ ...   │
└─────────────────────────────────────────────────────────────┘
```

## 📚 用户层类设计

### 1. TXWorkbook - 工作簿类

**作用**: Excel工作簿的主要入口点，管理多个工作表

**核心功能**:
```cpp
class TXWorkbook {
public:
    // 构造和析构
    TXWorkbook();
    explicit TXWorkbook(const std::string& filename);
    ~TXWorkbook();

    // 工作表管理
    TXSheet& createSheet(const std::string& name = "");
    TXSheet& getSheet(const std::string& name);
    TXSheet& getSheet(size_t index);
    bool removeSheet(const std::string& name);
    bool removeSheet(size_t index);
    
    // 工作表查询
    size_t getSheetCount() const;
    std::vector<std::string> getSheetNames() const;
    bool hasSheet(const std::string& name) const;
    
    // 文件操作
    TXResult<void> save();
    TXResult<void> saveAs(const std::string& filename);
    TXResult<void> load(const std::string& filename);
    
    // 属性设置
    void setAuthor(const std::string& author);
    void setTitle(const std::string& title);
    void setSubject(const std::string& subject);
    
private:
    std::vector<std::unique_ptr<TXSheet>> sheets_;
    std::string filename_;
    // 其他元数据...
};
```

**与其他类的关系**:
- **拥有** 多个 `TXSheet` 对象
- **使用** 底层的文件I/O和序列化功能
- **管理** 工作簿级别的元数据和设置

### 2. TXSheet - 工作表类

**作用**: 单个工作表的用户接口，提供便捷的单元格和范围操作

**核心功能**:
```cpp
class TXSheet {
public:
    // 构造
    explicit TXSheet(const std::string& name);
    
    // 单元格操作 - Excel格式坐标
    TXCell cell(const std::string& coord);           // "A1", "B2"
    TXCell cell(uint32_t row, uint32_t col);         // 0-based索引
    
    // 范围操作 - Excel格式
    TXRange range(const std::string& range);         // "A1:B10"
    TXRange range(const std::string& start, const std::string& end);
    TXRange range(uint32_t start_row, uint32_t start_col, 
                  uint32_t end_row, uint32_t end_col);
    
    // 便捷设置方法
    TXSheet& setValue(const std::string& coord, double value);
    TXSheet& setValue(const std::string& coord, const std::string& value);
    TXSheet& setValue(const std::string& coord, bool value);
    
    // 批量操作
    TXSheet& setValues(const std::string& range, 
                       const std::vector<std::vector<TXVariant>>& data);
    
    // 工作表属性
    const std::string& getName() const;
    void setName(const std::string& name);
    
    // 查询操作
    TXRange getUsedRange() const;
    size_t getCellCount() const;
    
    // 格式化
    TXSheet& setColumnWidth(uint32_t col, double width);
    TXSheet& setRowHeight(uint32_t row, double height);
    
private:
    std::string name_;
    std::unique_ptr<TXInMemorySheet> impl_;  // 底层实现
};
```

**与其他类的关系**:
- **属于** `TXWorkbook`
- **创建** `TXCell` 和 `TXRange` 对象
- **委托** 底层 `TXInMemorySheet` 进行实际操作
- **使用** `TXSheetAPI` 进行坐标转换

### 3. TXRange - 范围类

**作用**: 表示一个单元格范围，支持批量操作

**核心功能**:
```cpp
class TXRange {
public:
    // 构造
    TXRange(TXSheet& sheet, const TXCoordinate& start, const TXCoordinate& end);
    
    // 范围信息
    TXCoordinate getStart() const;
    TXCoordinate getEnd() const;
    size_t getRowCount() const;
    size_t getColumnCount() const;
    size_t getCellCount() const;
    
    // 批量设置
    TXRange& setValue(double value);                    // 填充单一值
    TXRange& setValue(const std::string& value);
    TXRange& setValues(const std::vector<std::vector<TXVariant>>& data);
    
    // 批量获取
    std::vector<std::vector<TXVariant>> getValues() const;
    
    // 数学运算
    TXRange& add(double value);
    TXRange& multiply(double value);
    TXRange& fillSequence(double start = 1.0, double step = 1.0);
    
    // 格式化
    TXRange& setStyle(const TXStyle& style);
    TXRange& setBorder(const TXBorder& border);
    TXRange& setFont(const TXFont& font);
    
    // 查找和替换
    std::vector<TXCell> find(const TXVariant& value) const;
    TXRange& replace(const TXVariant& old_value, const TXVariant& new_value);
    
    // 统计函数
    double sum() const;
    double average() const;
    double min() const;
    double max() const;
    
private:
    TXSheet& sheet_;
    TXCoordinate start_;
    TXCoordinate end_;
};
```

**与其他类的关系**:
- **属于** `TXSheet`
- **包含** 多个 `TXCell` 的逻辑集合
- **使用** 底层批量操作API提高性能

### 4. TXCell - 单元格类

**作用**: 单个单元格的用户接口

**核心功能**:
```cpp
class TXCell {
public:
    // 构造
    TXCell(TXSheet& sheet, const TXCoordinate& coord);
    
    // 值操作
    TXCell& setValue(double value);
    TXCell& setValue(const std::string& value);
    TXCell& setValue(bool value);
    TXCell& setFormula(const std::string& formula);
    
    TXVariant getValue() const;
    std::string getFormula() const;
    TXCellType getType() const;
    
    // 坐标信息
    TXCoordinate getCoordinate() const;
    std::string getAddress() const;                     // 返回"A1"格式
    uint32_t getRow() const;
    uint32_t getColumn() const;
    
    // 格式化
    TXCell& setStyle(const TXStyle& style);
    TXCell& setFont(const TXFont& font);
    TXCell& setBorder(const TXBorder& border);
    TXCell& setAlignment(TXAlignment alignment);
    
    // 数据验证
    TXCell& setValidation(const TXDataValidation& validation);
    
    // 注释
    TXCell& setComment(const std::string& comment);
    std::string getComment() const;
    
private:
    TXSheet& sheet_;
    TXCoordinate coord_;
};
```

**与其他类的关系**:
- **属于** `TXSheet` 和 `TXRange`
- **使用** `TXStyle`、`TXFont` 等格式化类
- **委托** 底层实现进行实际操作

### 5. TXStyle - 样式类

**作用**: 单元格格式化和样式设置

**核心功能**:
```cpp
class TXStyle {
public:
    // 字体设置
    TXStyle& setFont(const TXFont& font);
    TXStyle& setFontSize(double size);
    TXStyle& setFontColor(const TXColor& color);
    TXStyle& setBold(bool bold = true);
    TXStyle& setItalic(bool italic = true);
    
    // 背景和边框
    TXStyle& setBackgroundColor(const TXColor& color);
    TXStyle& setBorder(const TXBorder& border);
    
    // 对齐方式
    TXStyle& setHorizontalAlignment(TXHorizontalAlignment align);
    TXStyle& setVerticalAlignment(TXVerticalAlignment align);
    
    // 数字格式
    TXStyle& setNumberFormat(const std::string& format);
    
private:
    TXFont font_;
    TXColor background_color_;
    TXBorder border_;
    TXAlignment alignment_;
    std::string number_format_;
};
```

## 🔄 类之间的关系图

```
TXWorkbook (1) ──────── (n) TXSheet
    │                        │
    │                        ├── (n) TXRange
    │                        │      │
    │                        │      └── (n) TXCell (逻辑)
    │                        │
    │                        └── (n) TXCell
    │                               │
    │                               └── (1) TXStyle
    │
    └── 底层实现委托
            │
            ├── TXInMemorySheet
            ├── TXSheetAPI  
            └── TXBatchSIMDProcessor
```

## 🎯 设计原则

### 1. **简洁易用**
- 用户层API简洁直观，支持链式调用
- 支持Excel格式坐标（"A1", "B2:D10"）
- 提供常用操作的便捷方法

### 2. **高性能**
- 用户层委托底层高性能实现
- 批量操作自动使用SIMD优化
- 智能缓存和延迟计算

### 3. **类型安全**
- 强类型坐标系统
- 编译时类型检查
- 明确的错误处理

### 4. **可扩展性**
- 插件式的格式化系统
- 可自定义的数据验证
- 支持用户自定义函数

## 📝 使用示例

```cpp
// 创建工作簿
TXWorkbook workbook;

// 创建工作表
auto& sheet = workbook.createSheet("销售数据");

// 设置标题
sheet.setValue("A1", "产品名称")
     .setValue("B1", "销售额")
     .setValue("C1", "利润率");

// 填充数据
auto data_range = sheet.range("A2:C10");
data_range.setValues(sales_data);

// 计算总和
auto total_cell = sheet.cell("B11");
total_cell.setFormula("=SUM(B2:B10)");

// 格式化
TXStyle header_style;
header_style.setBold(true)
           .setBackgroundColor(TXColor::LIGHT_BLUE);

sheet.range("A1:C1").setStyle(header_style);

// 保存文件
workbook.saveAs("sales_report.xlsx");
```

这个设计确保了用户层的简洁性，同时充分利用了底层的高性能实现。

## 🚀 性能优化策略

### 1. **智能批量操作**
用户层会自动检测批量操作并委托给底层的高性能实现：

```cpp
// 用户写法（简单）
for (int i = 0; i < 10000; ++i) {
    sheet.setValue("A" + std::to_string(i+1), i * 0.1);
}

// 底层自动优化为批量操作
// 内部会收集操作，然后调用 setBatchNumbers()
```

### 2. **延迟计算**
- 格式化操作延迟到保存时应用
- 公式计算按需进行
- 索引更新批量进行

### 3. **内存管理**
- 自动使用统一内存管理器
- 智能预分配策略
- 及时释放临时对象

## 🔧 实现优先级

### Phase 1: 核心功能
1. **TXWorkbook** - 基本工作簿管理
2. **TXSheet** - 基本工作表操作
3. **TXCell** - 单元格读写
4. **TXRange** - 基本范围操作

### Phase 2: 增强功能
1. **TXStyle** - 格式化系统
2. **公式支持** - 基本公式计算
3. **数据验证** - 输入验证
4. **图表支持** - 基本图表

### Phase 3: 高级功能
1. **数据透视表**
2. **条件格式**
3. **宏支持**
4. **插件系统**

## 📋 API设计规范

### 1. **命名约定**
- 类名：`TXClassName`
- 方法名：`camelCase`
- 常量：`UPPER_CASE`
- 私有成员：`member_name_`

### 2. **错误处理**
- 使用 `TXResult<T>` 返回类型
- 提供详细的错误信息
- 支持异常和错误码两种模式

### 3. **链式调用**
- 设置方法返回引用支持链式调用
- 保持API的流畅性

### 4. **向后兼容**
- 新版本保持API兼容性
- 废弃的API提供迁移指南

## 🧪 测试策略

### 1. **单元测试**
- 每个类都有对应的测试文件
- 覆盖所有公共API
- 性能基准测试

### 2. **集成测试**
- 完整的工作流测试
- 与Excel兼容性测试
- 大数据量测试

### 3. **性能测试**
- 批量操作性能
- 内存使用测试
- 并发安全测试
