# 🏗️ TinaXlsx 类关系和依赖设计

## 📊 **整体类关系图**

```
                    🎯 用户API层
    ┌─────────────────────────────────────────────────┐
    │  TXWorkbook ←→ TXSheet ←→ TXCell                │
    │       ↓           ↓         ↓                   │
    │  TXRange ←→ TXCoordinate ←→ TXVariant           │
    └─────────────────────────────────────────────────┘
                            ↓
                🚀 高性能处理层 (核心)
    ┌─────────────────────────────────────────────────┐
    │  TXUnifiedMemoryManager ←→ TXCompactCellBuffer  │
    │           ↓                        ↓            │
    │  TXSlabAllocator              TXBatchSIMDProcessor│
    │  TXChunkAllocator             TXZeroCopySerializer│
    │           ↓                        ↓            │
    │  TXGlobalStringPool ←→ TXHighPerformanceLogger  │
    └─────────────────────────────────────────────────┘
                            ↓
                📁 文件I/O层 (第三方集成)
    ┌─────────────────────────────────────────────────┐
    │  TXXLSXReader ←→ minizip-ng ←→ pugixml         │
    │  TXXLSXWriter ←→ 标准文件I/O                   │
    └─────────────────────────────────────────────────┘
```

## 🎯 **用户API层类关系**

### **核心类依赖关系**

```cpp
// 主要依赖关系
TXWorkbook {
    std::vector<std::unique_ptr<TXSheet>> sheets_;
    TXUnifiedMemoryManager& memory_manager_;
    std::string name_;
}

TXSheet {
    TXCompactCellBuffer cell_buffer_;
    TXWorkbook* parent_workbook_;
    std::string name_;
    TXBatchSIMDProcessor* simd_processor_;
}

TXCell {
    TXSheet& parent_sheet_;
    TXCoordinate coordinate_;
    // 注意：TXCell不存储值，值存储在TXCompactCellBuffer中
}

TXRange {
    TXSheet& parent_sheet_;
    TXCoordinate start_;
    TXCoordinate end_;
}
```

### **值类型关系**

```cpp
// 值类型和坐标类型
TXVariant {
    enum Type { Empty, Number, String, Boolean, Error };
    union {
        double number_value_;
        uint32_t string_index_;  // 指向字符串池的索引
        bool boolean_value_;
        TXErrorCode error_code_;
    };
}

TXCoordinate {
    row_t row_;      // 强类型行号
    column_t col_;   // 强类型列号
}
```

## 🚀 **高性能处理层类关系**

### **内存管理层次结构**

```cpp
// 内存管理器层次
TXUnifiedMemoryManager {
    std::unique_ptr<TXSlabAllocator> slab_allocator_;
    std::unique_ptr<TXChunkAllocator> chunk_allocator_;
    TXGlobalStringPool& string_pool_;
    MemoryConfig config_;
    MemoryStats stats_;
}

TXSlabAllocator {
    struct SlabPool {
        void* memory_block_;
        size_t block_size_;
        std::vector<bool> allocation_bitmap_;
    };
    std::array<SlabPool, MAX_SLAB_SIZES> slab_pools_;
}

TXChunkAllocator {
    struct Chunk {
        void* data_;
        size_t size_;
        size_t used_;
        Chunk* next_;
    };
    Chunk* current_chunk_;
    size_t chunk_size_;
}
```

### **数据存储层次结构**

```cpp
// 紧凑数据存储
TXCompactCellBuffer {
    // 所有数据使用统一内存管理器分配
    TXUnifiedMemoryManager& memory_manager_;
    
    // 紧凑存储数组
    uint32_t* coordinates_;      // 压缩坐标 (row << 16 | col)
    double* number_values_;      // 数值数据
    uint32_t* string_indices_;   // 字符串池索引
    uint8_t* cell_types_;        // 单元格类型
    uint16_t* style_indices_;    // 样式索引
    
    // 容量管理
    size_t size_;
    size_t capacity_;
    
    // 性能优化
    bool is_sorted_;             // 是否已排序
    uint32_t* sorted_indices_;   // 排序索引
}

TXGlobalStringPool {
    TXUnifiedMemoryManager& memory_manager_;
    std::unordered_map<std::string, uint32_t> string_to_index_;
    TXVector<std::string> index_to_string_;
    mutable std::shared_mutex mutex_;  // 线程安全
}
```

### **SIMD处理器关系**

```cpp
// SIMD处理器
TXBatchSIMDProcessor {
    // 静态方法，无状态设计
    static void batchImportCells(TXCompactCellBuffer& buffer, 
                                const ParsedCellData& data);
    static Statistics calculateStatistics(const TXCompactCellBuffer& buffer);
    static void optimizeMemoryLayout(TXCompactCellBuffer& buffer);
    
    // SIMD特化实现
    template<typename T>
    static void vectorizedOperation(T* data, size_t count, 
                                   std::function<T(T)> operation);
}

// 零拷贝序列化器
TXZeroCopySerializer {
    TXUnifiedMemoryManager& memory_manager_;
    
    // 序列化格式定义
    struct SerializationHeader {
        uint32_t magic_number_;
        uint32_t version_;
        uint32_t cell_count_;
        uint32_t string_count_;
        uint64_t checksum_;
    };
}
```

## 📁 **文件I/O层集成关系**

### **第三方库集成**

```cpp
// XLSX读取器
class TXXLSXReader {
private:
    // 第三方库封装
    class MinizipWrapper {
        mz_zip_file* zip_handle_;
    public:
        TXResult<TXVector<uint8_t>> extractFile(const std::string& path);
    };
    
    class PugiXMLWrapper {
        pugi::xml_document doc_;
    public:
        TXResult<ParsedCellData> parseCells(const TXVector<uint8_t>& xml_data);
    };
    
    // 我们的高性能组件
    TXUnifiedMemoryManager& memory_manager_;
    TXBatchSIMDProcessor simd_processor_;
    
public:
    TXResult<std::unique_ptr<TXWorkbook>> loadFromFile(const std::string& file_path);
};

// XLSX写入器
class TXXLSXWriter {
private:
    TXZeroCopySerializer& serializer_;
    MinizipWrapper zip_writer_;
    
public:
    TXResult<void> saveToFile(const TXWorkbook& workbook, 
                             const std::string& file_path);
};
```

## 🔄 **数据流和生命周期**

### **对象创建流程**

```
1. TXWorkbook::create()
   ↓
2. 创建TXUnifiedMemoryManager实例
   ↓
3. 初始化TXGlobalStringPool
   ↓
4. TXWorkbook::createSheet()
   ↓
5. 创建TXCompactCellBuffer (使用内存管理器)
   ↓
6. TXSheet构造完成
```

### **数据设置流程**

```
1. sheet->cell("A1").setValue("Hello")
   ↓
2. TXCell::setValue() 调用
   ↓
3. 字符串添加到TXGlobalStringPool
   ↓
4. 在TXCompactCellBuffer中添加单元格记录
   ↓
5. 如果需要扩容，使用TXUnifiedMemoryManager分配内存
```

### **文件保存流程**

```
1. workbook->saveToFile("output.xlsx")
   ↓
2. TXZeroCopySerializer序列化TXCompactCellBuffer
   ↓
3. TXXLSXWriter生成XML结构
   ↓
4. MinizipWrapper压缩为ZIP文件
   ↓
5. 写入磁盘
```

## 🧵 **线程安全设计**

### **线程安全策略**

```cpp
// 线程安全级别
class ThreadSafetyLevels {
public:
    // 完全线程安全
    TXGlobalStringPool;          // 使用shared_mutex
    TXUnifiedMemoryManager;      // 使用mutex保护分配
    
    // 读取线程安全，写入需要外部同步
    TXCompactCellBuffer;         // 读取操作线程安全
    
    // 非线程安全 (用户负责同步)
    TXWorkbook;                  // 用户API层
    TXSheet;
    TXCell;
    TXRange;
};
```

### **并发访问模式**

```cpp
// 推荐的并发使用模式
void concurrentProcessing() {
    auto workbook = TXWorkbook::loadFromFile("data.xlsx");
    
    // 多线程读取 - 安全
    std::vector<std::thread> readers;
    for (int i = 0; i < 4; ++i) {
        readers.emplace_back([&workbook, i]() {
            auto sheet = workbook->getSheet(i);
            auto stats = sheet->calculateStatistics();  // 线程安全
        });
    }
    
    // 等待所有读取完成
    for (auto& t : readers) t.join();
    
    // 单线程写入 - 用户负责同步
    auto sheet = workbook->getSheet(0);
    sheet->cell("A1").setValue("Updated");
}
```

## 📊 **内存布局优化**

### **缓存友好设计**

```cpp
// 内存布局优化
struct CacheFriendlyLayout {
    // 热数据放在一起
    struct HotData {
        uint32_t* coordinates_;    // 经常访问
        uint8_t* cell_types_;      // 经常访问
        double* number_values_;    // 数值计算时访问
    } hot_data_;
    
    // 冷数据分离
    struct ColdData {
        uint32_t* string_indices_; // 字符串操作时访问
        uint16_t* style_indices_;  // 样式操作时访问
    } cold_data_;
};
```

这个设计确保了清晰的类关系，高效的内存使用，以及良好的性能特性。
