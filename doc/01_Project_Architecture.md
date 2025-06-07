# 🚀 TinaXlsx 项目架构设计文档

## 📋 **项目概述**

### **核心理念**
```
第三方库读取 → 导入到高性能内存系统 → 内存中高效处理 → 一次性输出
```

### **核心价值主张**
TinaXlsx 不是又一个Excel文件读写库，而是一个**高性能Excel数据处理引擎**：
- **内存效率提升3-5倍** - 自定义内存管理器
- **处理速度提升10倍** - SIMD加速批量处理
- **零拷贝序列化** - 超高效数据转换
- **大数据处理能力** - 处理百万行数据不卡顿

## 🏗️ **整体架构设计**

### **三层架构**

```
┌─────────────────────────────────────────────────────────────┐
│                    🎯 用户API层                              │
│  TXWorkbook, TXSheet, TXCell - 简洁易用的用户接口            │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                🚀 高性能处理层 (核心优势)                    │
│  • TXUnifiedMemoryManager - 统一内存管理                    │
│  • TXBatchSIMDProcessor - SIMD批量处理                      │
│  • TXZeroCopySerializer - 零拷贝序列化                      │
│  • TXCompactCellBuffer - 紧凑数据存储                       │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                📁 文件I/O层 (第三方集成)                     │
│  • minizip-ng - ZIP文件处理                                │
│  • pugixml - XML解析                                       │
│  • 标准文件I/O                                             │
└─────────────────────────────────────────────────────────────┘
```

### **数据流向**

```
XLSX文件 → minizip解压 → pugixml解析 → 导入TXCompactCellBuffer 
    ↓
SIMD批量处理 → 内存优化 → 用户API操作 → 零拷贝序列化 → 输出文件
```

## 🚀 **核心组件设计**

### **1. 内存管理层**

#### **TXUnifiedMemoryManager**
```cpp
class TXUnifiedMemoryManager {
public:
    // 统一内存分配接口
    void* allocate(size_t size, size_t alignment = 8);
    void deallocate(void* ptr, size_t size);
    
    // 内存池管理
    TXSlabAllocator& getSlabAllocator();
    TXChunkAllocator& getChunkAllocator();
    
    // 性能监控
    size_t getTotalMemoryUsage() const;
    MemoryStats getStats() const;
};
```

#### **TXCompactCellBuffer**
```cpp
class TXCompactCellBuffer {
    // 紧凑存储设计
    uint32_t* coordinates;     // 坐标压缩存储
    double* number_values;     // 数值数据
    uint32_t* string_indices;  // 字符串索引
    uint8_t* cell_types;       // 单元格类型
    uint16_t* style_indices;   // 样式索引
    
public:
    // 高性能操作
    void reserve(size_t capacity);
    void addCell(uint32_t row, uint32_t col, const TXVariant& value);
    TXVariant getCell(uint32_t row, uint32_t col) const;
};
```

### **2. SIMD处理层**

#### **TXBatchSIMDProcessor**
```cpp
class TXBatchSIMDProcessor {
public:
    // 批量数据处理
    static void batchImportCells(TXCompactCellBuffer& buffer, 
                                const ParsedCellData& data);
    
    // SIMD加速计算
    static Statistics calculateStatistics(const TXCompactCellBuffer& buffer);
    static void batchNumberConversion(TXCompactCellBuffer& buffer);
    
    // 内存布局优化
    static void optimizeMemoryLayout(TXCompactCellBuffer& buffer);
};
```

### **3. 零拷贝序列化层**

#### **TXZeroCopySerializer**
```cpp
class TXZeroCopySerializer {
public:
    // 零拷贝序列化
    TXResult<TXVector<uint8_t>> serialize(const TXCompactCellBuffer& buffer);
    TXResult<void> deserialize(const TXVector<uint8_t>& data, 
                              TXCompactCellBuffer& buffer);
    
    // 流式处理
    TXResult<void> serializeToStream(const TXCompactCellBuffer& buffer, 
                                    std::ostream& stream);
};
```

## 🎯 **用户API设计**

### **简洁的用户接口**

```cpp
// 创建工作簿
auto workbook = TXWorkbook::create("MyWorkbook");

// 创建工作表
auto sheet = workbook->createSheet("Sheet1");

// 设置单元格值
sheet->cell("A1").setValue("Hello");
sheet->cell("B1").setValue(42.0);

// 批量操作
auto range = sheet->range("A1:C10");
range.fill(100.0);

// 高性能保存
workbook->saveToFile("output.xlsx");
```

### **高性能特性展示**

```cpp
// 大数据处理
auto workbook = TXWorkbook::loadFromFile("huge_file.xlsx");  // 1秒加载100万行

// SIMD加速计算
auto stats = workbook->calculateStatistics();  // SIMD加速统计

// 零拷贝保存
workbook->saveOptimized("output.xlsx");  // 瞬间保存
```

## 📁 **文件I/O集成策略**

### **第三方库集成**

```cpp
class TXHighPerformanceXLSXReader {
private:
    // 第三方库组件
    MinizipReader zip_reader_;      // 使用minizip-ng
    PugiXMLParser xml_parser_;      // 使用pugixml
    
    // 我们的核心优势
    TXUnifiedMemoryManager& memory_manager_;
    TXBatchSIMDProcessor simd_processor_;
    TXZeroCopySerializer serializer_;
    
public:
    TXResult<TXWorkbook> loadXLSX(const std::string& file_path) {
        // 1. 第三方库读取
        auto xml_data = zip_reader_.extractWorksheet(file_path);
        auto parsed_cells = xml_parser_.parse(xml_data);
        
        // 2. 🚀 导入到我们的高性能系统
        TXCompactCellBuffer buffer(memory_manager_, parsed_cells.size());
        TXBatchSIMDProcessor::batchImportCells(buffer, parsed_cells);
        
        // 3. 🚀 创建用户API对象
        return TXWorkbook::fromBuffer(std::move(buffer));
    }
};
```

## 🎯 **性能目标**

### **基准测试目标**
- **内存使用** - 比标准库减少60%内存占用
- **加载速度** - 100万行数据1秒内加载完成
- **计算性能** - SIMD加速统计计算比标准实现快10倍
- **保存速度** - 零拷贝序列化比标准方法快5倍

### **支持规模**
- **最大行数** - 支持1000万行数据
- **最大列数** - 支持16384列（Excel标准）
- **内存效率** - 1GB内存处理100万行数据
- **并发处理** - 支持多线程并行处理

## 🔧 **技术栈**

### **核心依赖**
- **C++17** - 现代C++特性
- **minizip-ng** - ZIP文件处理
- **pugixml** - 轻量级XML解析
- **xsimd** - 跨平台SIMD抽象
- **fmt** - 高性能格式化

### **可选依赖**
- **fast_float** - 快速浮点数解析
- **GoogleTest** - 单元测试框架

## 📊 **项目结构**

```
TinaXlsx/
├── include/TinaXlsx/           # 公共头文件
│   ├── core/                   # 核心组件
│   ├── memory/                 # 内存管理
│   ├── simd/                   # SIMD处理
│   └── api/                    # 用户API
├── src/                        # 源文件实现
├── third_party/                # 第三方库
├── tests/                      # 测试代码
├── doc/                        # 文档
├── examples/                   # 示例代码
└── benchmarks/                 # 性能基准测试
```

## 🚀 **下一步计划**

1. **API设计** - 详细设计用户API接口
2. **核心实现** - 实现高性能内存管理和SIMD处理
3. **第三方集成** - 集成minizip-ng和pugixml
4. **性能优化** - 针对性能目标进行优化
5. **测试验证** - 全面的单元测试和性能测试

这个架构确保我们专注于核心优势，避免重新发明轮子，同时提供卓越的性能和易用性。
