# TinaXlsx 架构设计与解析流程文档

## 一、项目架构概述

### 1.1 设计理念

TinaXlsx 是一个高性能的 C++ Excel (XLSX) 文件处理库，采用现代 C++17 设计模式：

- **pImpl 模式**：隐藏实现细节，提供稳定的ABI
- **RAII 资源管理**：自动管理内存和文件资源
- **移动语义**：优化大对象的传递性能
- **类型安全**：使用 std::variant 实现类型安全的单元格数据

### 1.2 核心组件架构

```
TinaXlsx 核心架构
├── TXWorkbook（工作簿管理器）
│   ├── 管理多个工作表
│   ├── 文件读写协调
│   └── XLSX 结构生成
├── TXSheet（工作表处理器）
│   ├── 单元格集合管理
│   ├── 坐标系统转换（A1 <-> 数字）
│   └── 行列操作
├── TXCell（单元格数据容器）
│   ├── std::variant 多类型支持
│   ├── 自动类型转换
│   └── 公式与格式化
├── TXZipHandler（ZIP 压缩处理器）
│   ├── minizip-ng 封装
│   ├── 批量文件操作
│   └── 压缩级别控制
└── TXXmlHandler（XML 解析处理器）
    ├── pugixml 封装
    ├── XPath 查询优化
    └── 批量节点操作
```

## 二、XLSX 文件解析流程

### 2.1 XLSX 文件结构

XLSX 文件本质上是一个 ZIP 压缩包，包含以下关键文件：

```
example.xlsx (ZIP 文件)
├── [Content_Types].xml          # MIME 类型定义
├── _rels/
│   └── .rels                    # 主关系文件
└── xl/
    ├── workbook.xml             # 工作簿定义
    ├── _rels/
    │   └── workbook.xml.rels    # 工作簿关系
    ├── worksheets/
    │   ├── sheet1.xml           # 工作表数据
    │   └── sheet2.xml
    ├── sharedStrings.xml        # 共享字符串表
    └── styles.xml               # 样式定义
```

### 2.2 读取流程（loadFromFile）

```
开始读取 XLSX 文件
     ↓
┌─────────────────────┐
│   打开 ZIP 文件      │ ← TXZipHandler::open()
│   (TXZipHandler)    │
└─────────────────────┘
     ↓
┌─────────────────────┐
│   验证文件有效性     │ ← 检查 [Content_Types].xml
│                    │
└─────────────────────┘
     ↓
┌─────────────────────┐
│   解析工作簿结构     │ ← parseWorkbook()
│   (workbook.xml)   │   使用 TXXmlHandler
└─────────────────────┘
     ↓
┌─────────────────────┐
│   提取工作表信息     │ ← XPath: "//sheet"
│   创建 TXSheet 对象  │
└─────────────────────┘
     ↓
┌─────────────────────┐
│   读取单元格数据     │ ← 解析 worksheets/sheetN.xml
│   (可选，延迟加载)   │   填充 TXCell 对象
└─────────────────────┘
     ↓
┌─────────────────────┐
│   完成，关闭文件     │
└─────────────────────┘
```

### 2.3 写入流程（saveToFile）

```
开始保存 XLSX 文件
     ↓
┌─────────────────────┐
│   创建 ZIP 文件      │ ← TXZipHandler::open(Write)
└─────────────────────┘
     ↓
┌─────────────────────┐
│   生成 Content Types │ ← writeXlsxStructure()
│   [Content_Types].xml│
└─────────────────────┘
     ↓
┌─────────────────────┐
│   写入主关系文件     │ ← _rels/.rels
└─────────────────────┘
     ↓
┌─────────────────────┐
│   生成工作簿 XML     │ ← writeWorkbook()
│   xl/workbook.xml   │   使用 TXXmlHandler
└─────────────────────┘
     ↓
┌─────────────────────┐
│   写入工作簿关系     │ ← xl/_rels/workbook.xml.rels
└─────────────────────┘
     ↓
┌─────────────────────┐
│   遍历所有工作表     │ ← writeWorksheet() for each sheet
└─────────────────────┘
     ↓
┌─────────────────────┐
│   转换单元格数据     │ ← TXCell → XML 格式
│   写入工作表 XML     │   xl/worksheets/sheetN.xml
└─────────────────────┘
     ↓
┌─────────────────────┐
│   关闭并保存文件     │ ← TXZipHandler::close()
└─────────────────────┘
```

## 三、核心类详细设计

### 3.1 TXCell 数据类型系统

```cpp
// TXCell 内部数据表示
std::variant<
    std::monostate,    // 空值
    std::string,       // 字符串
    double,           // 数字（浮点）
    int64_t,          // 整数
    bool,             // 布尔值
    std::string       // 公式（以 = 开头）
> data_;

// 类型转换矩阵
空值    → ""    | 0.0   | 0     | false
字符串  → 原值  | 解析  | 解析   | 非空为true
数字    → 格式化 | 原值  | 转换   | 非零为true  
整数    → 格式化 | 转换  | 原值   | 非零为true
布尔    → "TRUE"/"FALSE" | 1.0/0.0 | 1/0 | 原值
公式    → 原值  | 计算结果 | 计算结果 | 计算结果
```

### 3.2 TXSheet 坐标系统

```cpp
// A1 样式坐标 ↔ 数字坐标转换
struct Coordinate {
    std::size_t row;    // 从 1 开始
    std::size_t col;    // 从 1 开始
};

// 转换示例
A1  → {1, 1}
B5  → {5, 2} 
Z99 → {99, 26}
AA1 → {1, 27}
```

### 3.3 性能优化策略

#### 3.3.1 内存管理优化
- 使用 `std::unique_ptr` 管理对象生命周期
- pImpl 模式减少头文件依赖
- 移动语义避免不必要的拷贝

#### 3.3.2 I/O 优化
- 批量ZIP文件操作减少系统调用
- XML 流式解析降低内存占用
- 延迟加载策略按需读取数据

#### 3.3.3 查询优化
- XPath 查询缓存
- 单元格坐标索引
- 稀疏矩阵存储（只存储非空单元格）

## 四、错误处理机制

### 4.1 错误传播链

```
低层错误（minizip-ng, pugixml）
     ↓
Impl 类错误收集
     ↓  
public 接口错误码
     ↓
用户层错误处理
```

### 4.2 典型错误场景

| 错误类型 | 检测位置 | 处理策略 |
|---------|----------|----------|
| 文件不存在 | TXZipHandler | 返回false + 错误信息 |
| ZIP损坏 | minizip-ng | 捕获并转换为友好信息 |
| XML格式错误 | pugixml | 解析位置 + 错误描述 |
| 内存不足 | 各层 | 异常安全，资源清理 |
| 类型转换失败 | TXCell | 返回默认值 + 警告 |

## 五、扩展点设计

### 5.1 插件化架构预留

```cpp
// 未来可扩展的接口
class TXFormulaEngine {  // 公式计算引擎
public:
    virtual double evaluate(const std::string& formula) = 0;
};

class TXStyleManager {   // 样式管理器
public:
    virtual void applyStyle(TXCell& cell, const Style& style) = 0;
};
```

### 5.2 格式支持扩展

- CSV 导入导出支持
- ODS 格式支持  
- 图表和图片支持
- 数据透视表支持

## 六、使用示例

### 6.1 基本读写操作

```cpp
#include "TinaXlsx/TinaXlsx.hpp"

// 读取 Excel 文件
TinaXlsx::TXWorkbook workbook;
if (workbook.loadFromFile("input.xlsx")) {
    auto* sheet = workbook.getSheet("Sheet1");
    if (sheet) {
        auto* cell = sheet->getCell("A1");
        if (cell) {
            std::cout << "A1的值: " << cell->getStringValue() << std::endl;
        }
    }
}

// 创建新的 Excel 文件
TinaXlsx::TXWorkbook new_workbook;
auto* new_sheet = new_workbook.addSheet("数据表");
new_sheet->getCell("A1")->setValue("Hello World");
new_sheet->getCell("B1")->setValue(42.5);
new_workbook.saveToFile("output.xlsx");
```

### 6.2 批量操作示例

```cpp
// 批量设置单元格值
auto* sheet = workbook.getSheet("Sheet1");
for (int row = 1; row <= 100; ++row) {
    for (int col = 1; col <= 10; ++col) {
        auto coord = TXSheet::Coordinate{row, col};
        sheet->getCell(coord)->setValue(
            "R" + std::to_string(row) + "C" + std::to_string(col)
        );
    }
}
```

## 七、测试策略

### 7.1 单元测试覆盖

- **TXCell**: 类型转换、边界条件、性能测试
- **TXSheet**: 坐标转换、范围操作、大数据集测试  
- **TXZipHandler**: 压缩级别、批量操作、错误恢复
- **TXXmlHandler**: XPath查询、批量修改、编码处理
- **TXWorkbook**: 完整文件读写、多工作表管理

### 7.2 集成测试场景

- 大文件处理性能测试
- 内存使用优化验证
- 多线程安全性测试
- 与 Excel 兼容性测试

---

**文档版本**: 1.0  
**创建日期**: 2024年12月  
**维护团队**: TinaXlsx 开发组 