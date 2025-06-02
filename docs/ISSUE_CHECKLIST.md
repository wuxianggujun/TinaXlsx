# TinaXlsx 问题检查清单

快速参考清单，用于验证和调试常见问题。

## 🔍 数值筛选问题检查清单

### 立即检查项目

- [ ] **编译成功**：项目能否正常编译
- [ ] **测试运行**：`DataFilterTests.exe` 能否正常运行
- [ ] **文件生成**：Excel文件是否正确生成
- [ ] **文件可打开**：Excel/WPS能否正常打开生成的文件

### Excel文件验证

- [ ] **筛选按钮显示**：数值列是否显示绿色筛选按钮
- [ ] **筛选菜单打开**：点击筛选按钮是否能打开筛选菜单
- [ ] **文本筛选工作**：文本列筛选是否能正常点击确定
- [ ] **数值筛选失败**：数值列筛选是否无法点击确定

### XML结构检查

#### 解压Excel文件
```bash
# 重命名为zip文件
copy data_filter_test.xlsx data_filter_test.zip

# 解压
Expand-Archive -Path data_filter_test.zip -DestinationPath extracted
```

#### 检查关键XML节点

**1. 工作表XML (`xl/worksheets/sheet1.xml`)**
- [ ] **数值单元格格式**：
  ```xml
  <c r="B2"><v>5999</v></c>  <!-- 当前格式 -->
  <c r="B2" t="n"><v>5999</v></c>  <!-- 可能需要的格式 -->
  ```

- [ ] **autoFilter节点存在**：
  ```xml
  <autoFilter ref="A1:D6">
  ```

- [ ] **filterColumn节点正确**：
  ```xml
  <filterColumn colId="1">  <!-- B列 = colId 1 -->
  ```

- [ ] **customFilter节点格式**：
  ```xml
  <customFilter operator="greaterThan" val="3000"/>
  ```

**2. 共享字符串XML (`xl/sharedStrings.xml`)**
- [ ] **文本内容正确**：检查中文字符是否正确显示

**3. 工作簿XML (`xl/workbook.xml`)**
- [ ] **工作表引用正确**：确保工作表正确注册

### 对比测试

#### 创建标准Excel文件进行对比
1. **手动创建相同数据**：
   - 在Excel中输入相同的数据
   - 设置相同的筛选条件
   - 保存为 `reference.xlsx`

2. **对比XML结构**：
   - 解压两个文件
   - 对比 `xl/worksheets/sheet1.xml`
   - 找出差异点

3. **关键差异检查**：
   - [ ] 数值单元格是否有 `t="n"` 属性
   - [ ] autoFilter的ref属性格式是否一致
   - [ ] customFilter的val属性格式是否一致
   - [ ] XML命名空间是否完整

## 🧪 测试验证流程

### 快速验证脚本

```bash
# 1. 编译测试
cmake --build cmake-build-debug --target DataFilterTests

# 2. 运行测试
cd cmake-build-debug
tests\unit\DataFilterTests.exe

# 3. 检查输出
dir tests\unit\test_output\DataFilterTest\

# 4. 验证Excel文件
start tests\unit\test_output\DataFilterTest\data_filter_test.xlsx
```

### 详细验证步骤

1. **基础功能验证**：
   - [ ] 文件能在Excel中打开
   - [ ] 数据显示正确
   - [ ] 中文字符显示正常

2. **筛选功能验证**：
   - [ ] 所有列都显示筛选按钮
   - [ ] 文本列筛选正常工作
   - [ ] 数值列筛选按钮可点击
   - [ ] 数值列筛选条件可设置
   - [ ] 数值列筛选"确定"按钮状态

3. **兼容性验证**：
   - [ ] Microsoft Excel 2016+
   - [ ] WPS Office
   - [ ] LibreOffice Calc

## 🔧 常见修复尝试

### 修复1：添加数值类型属性

**位置**：`src/TXWorksheetXmlHandler.cpp`

**修改**：在数值单元格中添加类型属性
```cpp
// 当前代码
cellNode.addChild(vNode);

// 可能的修复
if (std::holds_alternative<double>(value)) {
    cellNode.addAttribute("t", "n");  // 数值类型
    cellNode.addChild(vNode);
}
```

### 修复2：检查列索引

**验证**：确保列索引正确
- A列 = 0
- B列 = 1  
- C列 = 2
- D列 = 3

### 修复3：验证数值格式

**检查**：筛选条件中的数值格式
```cpp
// 确保格式一致
数据单元格: <v>5999</v>
筛选条件: val="5999"
```

## 📊 问题状态跟踪

| 检查项目 | 状态 | 备注 |
|---------|------|------|
| 编译成功 | ✅ | UTF-8编码已修复 |
| 测试运行 | ✅ | 独立测试可执行文件 |
| 文件生成 | ✅ | Excel文件正常生成 |
| 文本筛选 | ✅ | 工作正常 |
| 数值筛选 | ❌ | 无法点击确定 |
| XML格式 | ❓ | 需要进一步验证 |

## 🎯 下一步行动

### 优先级1：XML格式验证
- [ ] 创建标准Excel文件对比
- [ ] 检查数值单元格类型属性
- [ ] 验证autoFilter XML结构

### 优先级2：代码修复
- [ ] 添加数值类型属性 `t="n"`
- [ ] 验证列索引计算
- [ ] 测试修复效果

### 优先级3：兼容性测试
- [ ] 测试不同Excel版本
- [ ] 验证WPS Office兼容性
- [ ] 检查LibreOffice支持

## 📝 记录模板

发现新问题时使用此模板：

```
## 新问题：[问题标题]

**发现时间**：YYYY-MM-DD
**影响功能**：[具体功能]
**复现步骤**：
1. 步骤1
2. 步骤2

**预期结果**：[期望的行为]
**实际结果**：[实际发生的情况]
**临时解决方案**：[如果有的话]
**状态**：🔴未解决/🟡调查中/🟢已解决
```

---

**文档版本**：v1.1
**最后更新**：2025年1月
