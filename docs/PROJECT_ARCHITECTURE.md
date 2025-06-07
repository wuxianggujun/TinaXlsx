# TinaXlsx 项目架构文档

## 🎯 项目概述

TinaXlsx 是一个高性能的 C++ Excel 文件处理库，采用内存优先架构设计，专注于极致性能和内存效率。项目版本 2.1，基于现代 C++17 标准构建。

### 核心设计理念
- **内存优先**: 完全内存中操作，最后一次性序列化
- **SIMD 优化**: 批量处理，利用现代 CPU 向量指令
- **零拷贝设计**: 最小化内存复制，提升性能
- **模块化架构**: 清晰的分层设计，易于维护和扩展

## 🏗️ 五层架构设计

### 第一层：API 层 (Application Programming Interface Layer)
**职责**: 为用户提供简洁易用的接口

#### 核心组件
- **TinaXlsx.hpp**: 主入口头文件，统一包含所有必要组件
- **QuickExcel**: 快速 API，提供最简单的使用方式
- **TXInMemoryWorkbook**: 内存优先工作簿，顶层容器
- **别名系统**: 提供向后兼容的类型别名

#### 设计特点
- 统一的初始化和清理接口
- 版本信息管理
- 简化的快速使用别名
- 线程安全的全局状态管理

### 第二层：核心业务层 (Core Business Layer)
**职责**: Excel 文档的核心数据结构和业务逻辑

#### 主要组件
- **TXInMemorySheet**: 内存优先工作表，核心高性能组件
- **TXBatchSIMDProcessor**: SIMD 批量处理器，性能核心
- **TXZeroCopySerializer**: 零拷贝序列化器，极速 XML 生成
- **TXCompactCell**: 紧凑单元格，内存优化设计
- **TXSheet/TXWorkbook**: 传统兼容类，向后兼容

#### 性能特性
- 完全内存中操作
- SIMD 批量处理优化
- 智能内存布局优化
- 零拷贝序列化技术

### 第三层：内存管理层 (Memory Management Layer)
**职责**: 高效内存分配和管理

#### 核心组件
- **TXUnifiedMemoryManager**: 统一内存管理器，整合所有内存组件
- **TXSlabAllocator**: Slab 分配器，小对象高效分配 (≤8KB)
- **TXChunkAllocator**: 块分配器，大对象分配 (>8KB)
- **TXSmartMemoryManager**: 智能内存监控和管理
- **TXGlobalStringPool**: 全局字符串池，字符串去重优化

#### 技术特点
- 多级分配策略
- 智能内存监控
- 自动回收机制
- 高效的字符串池管理

### 第四层：基础支撑层 (Foundation Layer)
**职责**: 基础数据类型和工具支持

#### 基础类型系统
- **TXTypes.hpp**: 统一类型定义
- **TXVariant.hpp**: 通用数据类型，支持多种 Excel 数据类型
- **TXError.hpp**: 错误处理系统
- **TXResult.hpp**: 结果包装，安全的返回值处理
- **TXColor.hpp**: 颜色处理类
- **TXCoordinate.hpp**: 坐标系统，A1/R1C1 格式转换

#### 工具类
- **TXRange.hpp**: 范围操作
- **TXNumberUtils.cpp**: 数字工具
- **TXComponentManager.hpp**: 组件管理器

### 第五层：样式和功能层 (Style and Feature Layer)
**职责**: Excel 样式系统和高级功能

#### 样式系统
- **TXStyle.hpp**: 完整样式系统
- **TXFont.cpp**: 字体处理
- **TXNumberFormat.hpp**: 数字格式化
- **TXStyleTemplate.hpp**: 样式模板系统

#### 业务功能
- **TXFormula.hpp**: 公式处理
- **TXMergedCells.hpp**: 合并单元格管理
- **TXDataValidation.cpp**: 数据验证
- **TXConditionalFormat.cpp**: 条件格式化
- **TXDataFilter.cpp**: 数据筛选

## 🔧 依赖管理

### 核心依赖
- **fmt**: 高性能格式化库
- **xsimd**: 跨平台 SIMD 支持
- **minizip-ng**: ZIP 压缩处理
- **pugixml**: XML 处理（最小化使用）

### 测试依赖
- **GoogleTest**: 单元测试框架
- **自定义测试工具**: TestUtils.cmake 统一测试配置

## 📊 性能目标

### 核心性能指标
- **目标**: 2ms 生成 10k 单元格
- **内存效率**: >90% 内存利用率
- **处理速度**: >300K 单元格/秒
- **内存占用**: 控制在 4GB 以内

### 性能分解
- **0.5ms**: 内存操作 (SIMD 批量处理)
- **0.8ms**: XML 序列化 (零拷贝+模板)
- **0.5ms**: ZIP 压缩 (流式压缩)
- **0.2ms**: 文件写入 (一次性写入)

## 🧪 测试体系

### 测试分类
- **单元测试**: 组件级别的功能验证
- **性能测试**: 关键路径性能验证
- **集成测试**: 端到端工作流测试
- **功能测试**: API 功能完整性测试

### 测试工具
- **TestUtils.cmake**: 统一测试配置管理
- **性能基准**: 2ms 挑战测试
- **内存监控**: 内存泄漏检测
- **跨平台验证**: Windows/Linux/macOS 支持

## 📈 项目特色

### 技术创新
1. **内存优先架构**: 完全内存操作，避免频繁 I/O
2. **SIMD 批量优化**: 利用现代 CPU 向量指令
3. **零拷贝序列化**: 直接内存构建 XML，无中间拷贝
4. **智能内存管理**: 多级分配器，自动监控和回收

### 工程实践
1. **模块化设计**: 清晰的分层架构
2. **统一配置管理**: CMake 工具函数简化配置
3. **全面测试覆盖**: 功能、性能、集成多维度测试
4. **文档驱动开发**: 完整的架构和 API 文档

## 🚀 使用建议

### 推荐使用方式
1. **新项目**: 使用内存优先 API (TXInMemoryWorkbook)
2. **高性能场景**: 使用批量 SIMD 处理器
3. **大数据处理**: 启用智能内存管理
4. **兼容性需求**: 使用传统 API (TXWorkbook)

### 最佳实践
1. 调用 `TinaXlsx::initialize()` 初始化库
2. 使用批量操作提升性能
3. 合理配置内存管理参数
4. 在程序结束前调用 `TinaXlsx::cleanup()`
