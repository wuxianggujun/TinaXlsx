# TinaXlsx 已知问题清单

本文档记录了 TinaXlsx 项目中当前已知但尚未完全解决的问题。

## 📋 问题分类

- 🔴 **严重问题** - 影响核心功能
- 🟡 **中等问题** - 影响用户体验
- 🟢 **轻微问题** - 不影响主要功能
- 🔵 **功能缺失** - 计划中但未实现的功能

---

## 🔴 严重问题

### 1. 透视表兼容性问题

**问题描述**：
透视表功能虽然能够生成技术上正确的XML结构，但生成的XLSX文件无法被Excel/WPS等主流办公软件正确识别为包含透视表的文件。

**影响范围**：
- `TXPivotTable` 类的所有功能
- `TXPivotCache` 类的数据缓存功能
- 透视表相关的XML处理器
- `UnifiedPivotTableTest.ComprehensivePivotTableTest` 测试用例

**复现步骤**：
1. 运行 `UnifiedPivotTableTests.exe`
2. 打开生成的 `comprehensive_pivot_table_test.xlsx`
3. 发现Excel/WPS无法识别文件中的透视表
4. 文件显示为普通数据表，而非透视表

**已尝试的解决方案**：
- ✅ 修复了工作簿XML缺少 `<pivotCaches>` 元素的问题
- ✅ 修复了透视表定义缺少必需属性的问题（`autoFormatId`, `dataCaption`, `updatedVersion` 等）
- ✅ 修复了缓存记录格式错误（从直接值改为索引引用格式）
- ✅ 修复了 `pivotField` 缺少 `<items>` 子元素的问题
- ✅ 修复了 `rowItems` 和 `colItems` 结构问题
- ✅ 修复了 `dataField` 缺少 `baseField` 和 `baseItem` 属性的问题
- ✅ 完善了 `sharedItems` 的数据类型和统计信息
- ✅ 修复了 cacheId 从0开始的问题
- ❌ Excel仍无法识别生成的透视表

**当前状态**：🔴 **暂时搁置**

**可能原因**：
1. Excel可能对透视表XML有未公开的内部验证规则或校验和
2. 可能缺少特定的Excel内部标识符或元数据
3. 数据类型推断可能不够精确，需要更复杂的类型系统
4. 某些微妙的XML格式差异（如命名空间、属性顺序、编码等）
5. Excel可能需要特定的统计信息或缓存优化数据
6. 可能需要OLE复合文档格式而非纯XML格式

**解决方案建议**：
- **短期**：暂时跳过透视表功能，专注于其他更稳定的Excel功能
- **长期**：如需继续开发，建议深入研究Excel的OLE复合文档格式或使用逆向工程方法

**相关文件**：
- `include/TinaXlsx/TXPivotTable.hpp`（已添加TODO注释）
- `src/TXPivotTable.cpp`
- `src/TXPivotTableXmlHandler.cpp`
- `tests/unit/test_pivot_table_unified.cpp`

---

### 2. Excel数值筛选无法点击确定应用筛选

**问题描述**：
在生成的Excel文件中，数值列（如价格、薪资）的筛选条件虽然能正确显示绿色筛选按钮，但点击筛选按钮后无法点击"确定"来应用筛选条件。文本筛选工作正常。

**影响范围**：
- `DataFilterTest::AutoFilterBasicTest` - 价格列筛选
- `DataFilterTest::AutoFilterAdvancedTest` - 薪资列筛选

**复现步骤**：
1. 运行 `DataFilterTests.exe`
2. 打开生成的 `data_filter_test.xlsx` 或 `advanced_filter_test.xlsx`
3. 点击价格列或薪资列的筛选按钮
4. 尝试设置数值筛选条件
5. 发现无法点击"确定"按钮

**已尝试的解决方案**：
- ✅ 修复了数值格式不一致问题（数据单元格 vs 筛选条件）
- ✅ 集成了 fast_float 库统一数值格式化
- ✅ 确保XML中数值格式一致（如 `3000` 而不是 `3000.000000`）
- ❌ 问题仍然存在

**当前状态**：🔴 **未解决**

**可能原因**：
1. Excel XML格式规范理解不完整
2. 数值筛选的XML结构可能需要额外的属性或节点
3. 可能需要特定的数据类型声明
4. Excel版本兼容性问题

**下一步调查方向**：
- 对比标准Excel生成的数值筛选XML格式
- 研究Microsoft Office Open XML规范中的AutoFilter部分
- 测试不同Excel版本的兼容性
- 检查是否需要额外的XML命名空间或属性

---

## 🟡 中等问题

### 2. 编译时UTF-8编码警告

**问题描述**：
在Windows MSVC环境下编译时出现大量C4819警告，提示文件包含不能在当前代码页中表示的字符。

**状态**：🟢 **已修复**（通过CompilerConfig.cmake）

### 3. 测试文件无法在IDE中直接运行

**问题描述**：
之前的CMake自定义目标无法在IDE中直接点击运行，影响开发效率。

**状态**：🟢 **已修复**（创建了独立的测试可执行文件）

---

## 🟢 轻微问题

### 4. fast_float库CMake版本警告

**问题描述**：
fast_float库的CMakeLists.txt使用了较旧的CMake版本要求，产生弃用警告。

**警告信息**：
```
CMake Deprecation Warning at third_party/fast_float/CMakeLists.txt:1 (cmake_minimum_required):
  Compatibility with CMake < 3.10 will be removed from a future version of CMake.
```

**影响**：不影响功能，仅产生警告信息

**解决方案**：可以忽略，或向fast_float项目提交PR修复

---

## 🔵 功能缺失

### 5. 数据验证功能不完整

**问题描述**：
数据验证功能已有基础框架，但某些验证类型可能不完整。

**缺失功能**：
- 复杂的自定义验证规则
- 错误消息的完整本地化
- 验证失败时的详细错误信息

**优先级**：中等

### 6. 条件格式功能基础

**问题描述**：
条件格式功能处于早期开发阶段，可能存在兼容性问题。

**需要验证**：
- 与不同Excel版本的兼容性
- 复杂条件格式规则的支持
- 性能优化

**优先级**：低

### 7. 图表功能扩展

**问题描述**：
图表功能已实现基础功能，但还有许多高级特性未实现。

**缺失功能**：
- 3D图表支持
- 复杂的图表动画
- 图表模板系统
- 图表数据的动态更新

**优先级**：低

---

## 🔧 调试和诊断信息

### Excel数值筛选问题的详细分析

**生成的XML结构**：
```xml
<autoFilter ref="$A$1:$D$6">
    <filterColumn colId="1">
        <customFilters>
            <customFilter operator="greaterThan" val="3000" />
        </customFilters>
    </filterColumn>
</autoFilter>
```

**数据单元格格式**：
```xml
<c r="B2">
    <v>5999</v>
</c>
```

**已验证的方面**：
- ✅ 列索引正确（B列 = colId="1"）
- ✅ 数值格式一致（都是整数格式）
- ✅ XML结构符合基本规范
- ✅ 文本筛选工作正常

**仍需调查**：
- 是否需要数据类型属性（如 `t="n"` 表示数值）
- 是否需要特定的XML命名空间
- 是否需要额外的元数据节点

---

## 📊 问题统计

| 严重程度 | 数量 | 已解决 | 未解决 |
|---------|------|--------|--------|
| 🔴 严重  | 2    | 0      | 2      |
| 🟡 中等  | 2    | 2      | 0      |
| 🟢 轻微  | 1    | 0      | 1      |
| 🔵 功能缺失 | 3 | 0      | 3      |
| **总计** | **8** | **2** | **6** |

---

## 🎯 优先级排序

1. **最高优先级**：Excel数值筛选问题（影响核心功能）
2. **暂时搁置**：透视表兼容性问题（技术复杂度过高，建议跳过）
3. **高优先级**：数据验证功能完善
4. **中优先级**：条件格式功能验证
5. **低优先级**：图表功能扩展、CMake警告修复

---

## 📝 报告新问题

如果发现新的问题，请按以下格式添加到本文档：

```markdown
### X. 问题标题

**问题描述**：详细描述问题现象

**影响范围**：列出受影响的功能或测试

**复现步骤**：
1. 步骤1
2. 步骤2
3. ...

**已尝试的解决方案**：
- ✅ 已尝试的方案1
- ❌ 失败的方案2

**当前状态**：🔴/🟡/🟢 **状态描述**

**可能原因**：列出可能的原因

**下一步调查方向**：具体的调查计划
```

---

## 🔬 深度技术分析

### Excel数值筛选问题的深度分析

**问题核心**：Excel无法识别我们生成的数值筛选条件为有效的筛选规则。

**技术细节**：

#### 1. XML结构对比
**我们生成的格式**：
```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="2">
            <c r="B2"><v>5999</v></c>
        </row>
    </sheetData>
    <autoFilter ref="A1:D6">
        <filterColumn colId="1">
            <customFilters>
                <customFilter operator="greaterThan" val="3000"/>
            </customFilters>
        </filterColumn>
    </autoFilter>
</worksheet>
```

**可能需要的格式**（待验证）：
```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="2">
            <c r="B2" t="n"><v>5999</v></c>  <!-- 注意 t="n" 数值类型 -->
        </row>
    </sheetData>
    <autoFilter ref="A1:D6">
        <filterColumn colId="1">
            <customFilters>
                <customFilter operator="greaterThan" val="3000"/>
            </customFilters>
        </filterColumn>
    </autoFilter>
</worksheet>
```

#### 2. 可能的解决方向

**方向A：数据类型声明**
- 在数值单元格中添加 `t="n"` 属性
- 确保Excel识别数据为数值类型

**方向B：筛选条件格式**
- 研究是否需要特定的 `val` 属性格式
- 检查是否需要额外的筛选元数据

**方向C：XML命名空间**
- 验证是否使用了正确的XML命名空间
- 检查是否缺少必要的schema声明

#### 3. 调试建议

1. **创建标准Excel文件**：
   - 在Excel中手动创建相同的数据和筛选
   - 保存为.xlsx文件
   - 解压并对比XML结构

2. **逐步验证**：
   - 先测试最简单的数值筛选
   - 逐步增加复杂度
   - 找出最小可工作的XML格式

3. **工具验证**：
   - 使用Excel的XML验证工具
   - 检查生成的文件是否符合Office Open XML标准

---

## 📞 联系信息

- **项目维护者**：wuxianggujun
- **问题跟踪**：本文档
- **代码仓库**：TinaXlsx项目
- **技术讨论**：可通过代码注释或文档更新进行

---

**最后更新**：2025年1月（文档重构，添加深度技术分析）
**文档版本**：v1.2
