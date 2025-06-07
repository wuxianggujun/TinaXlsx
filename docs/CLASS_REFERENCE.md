# TinaXlsx 类参考文档

## 🏗️ 类分层架构

### API 层类 (Application Layer)

#### TinaXlsx 命名空间
**文件**: `include/TinaXlsx/TinaXlsx.hpp`
**作用**: 主入口命名空间，提供库的初始化、版本管理和全局接口

**核心功能**:
- `bool initialize()`: 库初始化，线程安全
- `void cleanup()`: 资源清理，线程安全  
- `std::string getVersion()`: 获取版本信息
- `std::string getBuildInfo()`: 获取构建信息

**🚀 无别名设计**:
```cpp
// 直接使用完整类名，TX前缀防止命名冲突：
// - TXWorkbook (用户层工作簿)
// - TXSheet (用户层工作表)
// - TXCell (用户层单元格)
// - TXCoordinate (坐标类)
// - TXColor (颜色类)
```

#### TXInMemoryWorkbook
**文件**: `include/TinaXlsx/TXInMemorySheet.hpp`
**作用**: 内存优先工作簿，顶层容器类

**核心功能**:
- 工作表管理 (创建、删除、查找)
- 全局内存管理器集成
- 全局字符串池管理
- 自动保存功能
- 批量操作支持

**关键方法**:
- `createSheet(name)`: 创建工作表
- `getSheet(name/index)`: 获取工作表
- `save(filename)`: 保存到文件
- `getMemoryStats()`: 获取内存统计

### 核心业务层类 (Core Business Layer)

#### TXInMemorySheet
**文件**: `include/TinaXlsx/TXInMemorySheet.hpp`
**作用**: 内存优先工作表，核心高性能组件

**设计理念**:
- 完全内存中操作，最后一次性序列化
- SIMD 批量处理，极致性能
- 零拷贝设计，最小内存开销
- 智能内存布局，优化缓存命中

**核心功能**:
- 单元格批量操作
- SIMD 优化的数据处理
- 智能内存布局优化
- 性能统计和监控

**关键方法**:
- `batchSetNumbers()`: 批量设置数值
- `batchSetStrings()`: 批量设置字符串
- `optimizeLayout()`: 优化内存布局
- `getPerformanceStats()`: 获取性能统计

#### TXBatchSIMDProcessor
**文件**: `include/TinaXlsx/TXBatchSIMDProcessor.hpp`
**作用**: SIMD 批量处理器，性能核心组件

**核心功能**:
- 批量单元格创建 (SIMD 优化)
- 批量数据转换 (SIMD 优化)
- 批量坐标处理 (SIMD 优化)
- 批量统计计算 (SIMD 优化)

**关键方法**:
- `batchCreateNumberCells()`: 批量创建数值单元格
- `batchCreateStringCells()`: 批量创建字符串单元格
- `batchConvertCoordinates()`: 批量坐标转换
- `batchCalculateStats()`: 批量统计计算

#### TXZeroCopySerializer
**文件**: `include/TinaXlsx/TXZeroCopySerializer.hpp`
**作用**: 零拷贝序列化器，极速 XML 生成核心

**设计目标**:
- 零拷贝：直接在内存中构建 XML，无中间拷贝
- 模板化：预编译 XML 模板，运行时只填充数据
- 流式处理：边生成边压缩，不等待完整 XML
- SIMD 优化：批量字符串操作和数值转换

**核心功能**:
- XML 模板缓存系统
- 零拷贝内存操作
- 流式压缩处理
- 性能统计监控

#### TXCompactCell
**文件**: `include/TinaXlsx/TXCompactCell.hpp`
**作用**: 紧凑单元格，内存优化设计

**内存优化**:
- 16 字节超紧凑设计
- 智能数据类型压缩
- 字符串池索引引用
- 样式信息压缩存储

### 内存管理层类 (Memory Management Layer)

#### TXUnifiedMemoryManager
**文件**: `include/TinaXlsx/TXUnifiedMemoryManager.hpp`
**作用**: 统一内存管理器，整合所有内存组件

**核心组件整合**:
- TXSlabAllocator: 小对象分配（≤8KB）
- TXChunkAllocator: 大对象分配（>8KB）
- TXSmartMemoryManager: 智能监控和管理

**关键功能**:
- `allocate(size)`: 智能分配，自动选择最优分配器
- `deallocate(ptr)`: 智能释放，自动识别分配器类型
- `getUnifiedStats()`: 获取综合统计信息
- `optimizeMemory()`: 内存优化和整理

#### TXSlabAllocator
**文件**: `include/TinaXlsx/TXSlabAllocator.hpp`
**作用**: Slab 分配器，小对象高效分配

**技术特点**:
- 预分配固定大小内存块
- 空闲链表管理
- 多种对象大小支持 (16B, 32B, 64B, 128B, 256B, 512B, 1KB, 2KB)
- 智能回收机制

**性能优势**:
- O(1) 分配和释放
- 减少内存碎片
- 高缓存命中率
- 线程安全设计

#### TXChunkAllocator
**文件**: `include/TinaXlsx/TXChunkAllocator.hpp`
**作用**: 块分配器，大对象分配

**设计特点**:
- 大块内存预分配
- 线性分配策略
- 批量分配支持
- 自动压缩功能

**适用场景**:
- 大数据缓冲区
- 临时工作内存
- 批量对象分配

#### TXSmartMemoryManager
**文件**: `include/TinaXlsx/TXSmartMemoryManager.hpp`
**作用**: 智能内存监控和管理

**监控功能**:
- 实时内存使用监控
- 阈值告警系统
- 自动清理策略
- 内存趋势分析

**管理策略**:
- 预警阈值: 3GB
- 临界阈值: 3.5GB  
- 紧急阈值: 3.75GB
- 自动清理和压缩

#### TXGlobalStringPool
**文件**: `include/TinaXlsx/TXGlobalStringPool.hpp`
**作用**: 全局字符串池，字符串去重优化

**优化技术**:
- 字符串去重存储
- 引用计数管理
- 快速查找索引
- 内存压缩优化

### 基础支撑层类 (Foundation Layer)

#### TXVariant
**文件**: `include/TinaXlsx/TXVariant.hpp`
**作用**: 通用数据类型，支持多种 Excel 数据类型

**支持类型**:
- 数值 (double)
- 字符串 (std::string)
- 布尔值 (bool)
- 日期时间
- 公式
- 错误值

#### TXCoordinate
**文件**: `include/TinaXlsx/TXCoordinate.hpp`
**作用**: 坐标系统，A1/R1C1 格式转换

**功能特点**:
- A1 格式支持 (A1, B2, AA100)
- R1C1 格式支持
- 坐标运算操作
- 范围计算支持

#### TXError & TXResult
**文件**: `include/TinaXlsx/TXError.hpp`, `include/TinaXlsx/TXResult.hpp`
**作用**: 错误处理系统和结果包装

**设计模式**:
- Result<T> 模式，安全的返回值处理
- 详细的错误信息
- 错误码分类管理
- 异常安全保证

### 样式和功能层类 (Style and Feature Layer)

#### TXStyle
**文件**: `include/TinaXlsx/TXStyle.hpp`
**作用**: 完整样式系统

**样式组件**:
- 字体样式 (TXFont)
- 对齐方式 (TXAlignment)
- 边框样式 (TXBorder)
- 填充样式 (TXFill)
- 数字格式 (TXNumberFormat)

#### TXFormula
**文件**: `include/TinaXlsx/TXFormula.hpp`
**作用**: 公式处理系统

**功能支持**:
- 公式解析和验证
- 依赖关系分析
- 计算引擎集成
- 公式优化

#### 高级功能类
- **TXMergedCells**: 合并单元格管理
- **TXDataValidation**: 数据验证规则
- **TXConditionalFormat**: 条件格式化
- **TXDataFilter**: 数据筛选功能

## 🔗 类关系总结

### 依赖关系
1. **API 层** → **核心业务层** → **内存管理层** → **基础支撑层**
2. **样式功能层** 横跨多个层次，提供专业功能支持
3. **测试层** 独立于业务逻辑，提供全面测试覆盖

### 核心交互模式
1. **TXInMemoryWorkbook** 管理多个 **TXInMemorySheet**
2. **TXInMemorySheet** 使用 **TXBatchSIMDProcessor** 进行批量处理
3. **TXZeroCopySerializer** 负责最终的 XML 序列化
4. **TXUnifiedMemoryManager** 为所有组件提供内存管理服务
5. **TXGlobalStringPool** 为整个系统提供字符串优化
