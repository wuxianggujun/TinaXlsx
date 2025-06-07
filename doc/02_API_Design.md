# 🎯 TinaXlsx API设计文档

## 📋 **设计原则**

### **核心原则**
1. **简洁易用** - 用户API应该直观易懂
2. **性能优先** - 所有操作都应该高性能
3. **内存安全** - 使用RAII和智能指针
4. **错误处理** - 使用TXResult进行错误处理
5. **链式调用** - 支持流畅的API调用

### **性能承诺**
- 所有API调用都应该是O(1)或接近O(1)
- 批量操作使用SIMD加速
- 内存分配使用自定义内存管理器
- 零拷贝数据传输

## 🚀 **核心API设计**

### **1. TXWorkbook - 工作簿类**

```cpp
class TXWorkbook {
public:
    // ==================== 创建和加载 ====================
    
    /// 创建新工作簿
    static std::unique_ptr<TXWorkbook> create(const std::string& name = "Workbook");
    
    /// 从文件加载工作簿 (高性能)
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromFile(const std::string& file_path);
    
    /// 从内存加载工作簿
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromMemory(const void* data, size_t size);
    
    // ==================== 工作表管理 ====================
    
    /// 创建新工作表
    TXSheet* createSheet(const std::string& name);
    
    /// 获取工作表 (按索引)
    TXSheet* getSheet(size_t index);
    
    /// 获取工作表 (按名称)
    TXSheet* getSheet(const std::string& name);
    
    /// 获取工作表数量
    size_t getSheetCount() const;
    
    /// 删除工作表
    bool removeSheet(size_t index);
    bool removeSheet(const std::string& name);
    
    // ==================== 保存和导出 ====================
    
    /// 保存到文件 (高性能)
    TXResult<void> saveToFile(const std::string& file_path);
    
    /// 保存到内存 (零拷贝)
    TXResult<TXVector<uint8_t>> saveToMemory();
    
    /// 导出为CSV
    TXResult<void> exportToCSV(const std::string& file_path, size_t sheet_index = 0);
    
    // ==================== 属性访问 ====================
    
    /// 获取/设置工作簿名称
    const std::string& getName() const;
    void setName(const std::string& name);
    
    /// 获取内存使用统计
    MemoryStats getMemoryStats() const;
    
    /// 获取性能统计
    PerformanceStats getPerformanceStats() const;
};
```

### **2. TXSheet - 工作表类**

```cpp
class TXSheet {
public:
    // ==================== 单元格访问 ====================
    
    /// 获取单元格 (A1格式)
    TXCell cell(const std::string& address);
    
    /// 获取单元格 (行列索引，0基)
    TXCell cell(uint32_t row, uint32_t col);
    
    /// 获取单元格范围
    TXRange range(const std::string& range_address);  // "A1:C10"
    TXRange range(uint32_t start_row, uint32_t start_col, 
                  uint32_t end_row, uint32_t end_col);
    
    // ==================== 批量操作 (高性能) ====================
    
    /// 批量设置数值 (SIMD加速)
    TXResult<void> batchSetNumbers(const TXRange& range, 
                                  const std::vector<double>& values);
    
    /// 批量设置字符串
    TXResult<void> batchSetStrings(const TXRange& range, 
                                  const std::vector<std::string>& values);
    
    /// 批量填充
    TXResult<void> fill(const TXRange& range, const TXVariant& value);
    
    /// 批量清空
    TXResult<void> clear(const TXRange& range);
    
    // ==================== 数据分析 (SIMD加速) ====================
    
    /// 计算统计信息
    TXResult<Statistics> calculateStatistics(const TXRange& range);
    
    /// 查找数据
    TXResult<std::vector<TXCoordinate>> find(const TXVariant& value);
    
    /// 排序数据
    TXResult<void> sort(const TXRange& range, SortOrder order = SortOrder::Ascending);
    
    // ==================== 工作表属性 ====================
    
    /// 获取/设置工作表名称
    const std::string& getName() const;
    void setName(const std::string& name);
    
    /// 获取使用范围
    TXRange getUsedRange() const;
    
    /// 获取最大行列
    uint32_t getMaxRow() const;
    uint32_t getMaxColumn() const;
    
    /// 获取单元格数量
    size_t getCellCount() const;
    
    // ==================== 行列操作 ====================
    
    /// 插入行
    TXResult<void> insertRows(uint32_t row, uint32_t count = 1);
    
    /// 删除行
    TXResult<void> deleteRows(uint32_t row, uint32_t count = 1);
    
    /// 插入列
    TXResult<void> insertColumns(uint32_t col, uint32_t count = 1);
    
    /// 删除列
    TXResult<void> deleteColumns(uint32_t col, uint32_t count = 1);
};
```

### **3. TXCell - 单元格类**

```cpp
class TXCell {
public:
    // ==================== 值操作 ====================
    
    /// 设置值
    TXCell& setValue(const TXVariant& value);
    TXCell& setValue(double value);
    TXCell& setValue(const std::string& value);
    TXCell& setValue(const char* value);
    
    /// 获取值
    TXVariant getValue() const;
    
    /// 类型检查
    bool isEmpty() const;
    bool isNumber() const;
    bool isString() const;
    
    /// 类型转换
    double asNumber() const;
    std::string asString() const;
    
    // ==================== 坐标信息 ====================
    
    /// 获取坐标
    TXCoordinate getCoordinate() const;
    
    /// 获取地址字符串 (如 "A1")
    std::string getAddress() const;
    
    /// 获取行列索引
    uint32_t getRow() const;
    uint32_t getColumn() const;
    
    // ==================== 样式操作 (简化) ====================
    
    /// 设置字体
    TXCell& setFont(const TXFont& font);
    
    /// 设置背景色
    TXCell& setBackgroundColor(const TXColor& color);
    
    /// 设置边框
    TXCell& setBorder(const TXBorder& border);
    
    /// 设置对齐方式
    TXCell& setAlignment(TXAlignment alignment);
    
    // ==================== 链式调用支持 ====================
    
    /// 支持链式调用
    TXCell& operator=(const TXVariant& value) { return setValue(value); }
    TXCell& operator=(double value) { return setValue(value); }
    TXCell& operator=(const std::string& value) { return setValue(value); }
};
```

### **4. TXRange - 范围类**

```cpp
class TXRange {
public:
    // ==================== 构造 ====================
    
    TXRange(const TXCoordinate& start, const TXCoordinate& end);
    TXRange(const std::string& range_address);  // "A1:C10"
    
    // ==================== 范围信息 ====================
    
    /// 获取起始/结束坐标
    TXCoordinate getStart() const;
    TXCoordinate getEnd() const;
    
    /// 获取范围大小
    uint32_t getRowCount() const;
    uint32_t getColumnCount() const;
    size_t getCellCount() const;
    
    /// 范围检查
    bool isEmpty() const;
    bool contains(const TXCoordinate& coord) const;
    
    // ==================== 批量操作 ====================
    
    /// 批量填充
    TXResult<void> fill(const TXVariant& value);
    
    /// 批量清空
    TXResult<void> clear();
    
    /// 批量设置数值 (SIMD加速)
    TXResult<void> setNumbers(const std::vector<double>& values);
    
    /// 批量设置字符串
    TXResult<void> setStrings(const std::vector<std::string>& values);
    
    /// 获取所有值
    TXResult<std::vector<TXVariant>> getValues() const;
    
    // ==================== 数据分析 ====================
    
    /// 计算统计信息 (SIMD加速)
    TXResult<Statistics> calculateStatistics() const;
    
    /// 求和 (SIMD加速)
    TXResult<double> sum() const;
    
    /// 平均值 (SIMD加速)
    TXResult<double> average() const;
    
    /// 最大/最小值
    TXResult<double> max() const;
    TXResult<double> min() const;
    
    // ==================== 迭代器支持 ====================
    
    class iterator {
    public:
        TXCell operator*();
        iterator& operator++();
        bool operator!=(const iterator& other) const;
    };
    
    iterator begin();
    iterator end();
};
```

## 🚀 **高性能特性API**

### **1. 批量处理API**

```cpp
namespace TXBatch {
    /// 批量导入数据 (SIMD加速)
    TXResult<void> importFromCSV(TXSheet* sheet, const std::string& csv_file);
    
    /// 批量数值计算 (SIMD加速)
    TXResult<void> batchCalculate(TXRange& range, 
                                 std::function<double(double)> func);
    
    /// 批量数据转换
    TXResult<void> batchTransform(TXRange& range, 
                                 std::function<TXVariant(const TXVariant&)> func);
}
```

### **2. 内存优化API**

```cpp
namespace TXMemory {
    /// 内存使用统计
    struct MemoryStats {
        size_t total_allocated;
        size_t total_used;
        size_t peak_usage;
        size_t cell_buffer_size;
        size_t string_pool_size;
    };
    
    /// 获取全局内存统计
    MemoryStats getGlobalStats();
    
    /// 内存优化
    TXResult<void> optimizeMemoryLayout(TXWorkbook* workbook);
    
    /// 垃圾回收
    TXResult<void> collectGarbage();
}
```

### **3. 性能监控API**

```cpp
namespace TXPerformance {
    /// 性能统计
    struct PerformanceStats {
        double load_time_ms;
        double save_time_ms;
        double calculation_time_ms;
        size_t simd_operations_count;
        size_t memory_allocations_count;
    };
    
    /// 获取性能统计
    PerformanceStats getStats(TXWorkbook* workbook);
    
    /// 重置统计
    void resetStats();
    
    /// 性能基准测试
    TXResult<BenchmarkResult> runBenchmark(const std::string& test_name);
}
```

## 🎯 **使用示例**

### **基础使用**

```cpp
// 创建工作簿
auto workbook = TXWorkbook::create("MyWorkbook");
auto sheet = workbook->createSheet("Data");

// 设置数据
sheet->cell("A1").setValue("Name");
sheet->cell("B1").setValue("Age");
sheet->cell("A2").setValue("Alice");
sheet->cell("B2").setValue(25.0);

// 保存文件
workbook->saveToFile("output.xlsx");
```

### **高性能批量处理**

```cpp
// 加载大文件
auto workbook = TXWorkbook::loadFromFile("huge_data.xlsx");
auto sheet = workbook->getSheet(0);

// 批量计算 (SIMD加速)
auto range = sheet->range("B:B");
auto stats = range->calculateStatistics();

// 批量填充
auto target_range = sheet->range("C1:C1000000");
target_range->fill(100.0);

// 零拷贝保存
workbook->saveToFile("processed_data.xlsx");
```

### **链式调用**

```cpp
auto workbook = TXWorkbook::create("ChainExample");
workbook->createSheet("Sheet1")
        ->cell("A1").setValue("Hello")
        .setFont(TXFont::bold())
        .setBackgroundColor(TXColor::yellow());
```

这个API设计确保了简洁易用的同时，充分展示了TinaXlsx的高性能特性。
