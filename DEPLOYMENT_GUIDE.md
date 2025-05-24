# TinaXlsx GitHub Actions 自动化部署指南

## 🎯 概述

本项目已配置完整的 GitHub Actions 工作流，可以自动化完成以下任务：
- 多平台编译（Windows、Linux、macOS）
- 自动化测试
- 代码质量检查
- 自动发布到 GitHub Releases

## 🚀 如何触发自动发布

### 方法1：标签发布（推荐）
```bash
# 1. 创建版本标签
git tag -a v1.0.0 -m "Release version 1.0.0"

# 2. 推送标签到GitHub
git push origin v1.0.0
```

### 方法2：手动触发
1. 进入 GitHub 仓库页面
2. 点击 "Actions" 标签
3. 选择 "Build and Release" 工作流
4. 点击 "Run workflow" 按钮

## 📁 工作流文件说明

### `.github/workflows/build.yml`
主要的构建和发布工作流：
- **触发条件**: 标签推送、主分支推送、手动触发
- **功能**: 
  - 多平台编译
  - 运行测试
  - 创建发布包
  - 自动发布到 GitHub Releases

### `.github/workflows/ci.yml`
持续集成工作流：
- **触发条件**: Pull Request、主分支推送
- **功能**:
  - 快速编译测试
  - 代码质量检查

## 🛠️ 发布包内容

每个平台的发布包都包含：

```
TinaXlsx-{Platform}-x64/
├── lib/                     # 静态库文件
│   ├── TinaXlsx.lib/.a     # 主库
│   ├── xlsxwriter.lib/.a   # Excel写入库
│   ├── xlsxio_*.lib/.a     # Excel I/O库
│   ├── zlib*.lib/.a        # 压缩库
│   ├── expat.lib/.a        # XML解析库
│   └── pkgconfig/          # pkg-config配置(Linux/macOS)
├── include/                 # 头文件
│   └── TinaXlsx/
├── README.md               # 项目说明
├── LICENSE                 # 许可证
└── README.txt              # 使用说明
```

## 🔄 版本发布流程

### 1. 准备发布
```bash
# 确保代码已提交
git add .
git commit -m "Prepare for release v1.0.0"
git push origin main
```

### 2. 创建发布标签
```bash
# 创建带注释的标签
git tag -a v1.0.0 -m "Release v1.0.0

新特性:
- 高性能类型转换系统
- 类型安全枚举
- 跨平台兼容性改进

修复:
- 修复字符串转换性能问题
- 解决编译器兼容性问题"

# 推送标签
git push origin v1.0.0
```

### 3. 监控构建过程
1. 打开 GitHub Actions 页面
2. 查看 "Build and Release" 工作流状态
3. 等待所有平台编译完成

### 4. 验证发布
1. 检查 Releases 页面是否出现新版本
2. 下载测试各平台的发布包
3. 验证包内容完整性

## 📋 发布检查清单

发布前请确保：

- [ ] 所有测试通过
- [ ] 版本号已更新（CMakeLists.txt）
- [ ] 更新日志已准备（CHANGELOG.md）
- [ ] 文档已更新
- [ ] 示例代码已验证
- [ ] 许可证信息正确

## 🐛 常见问题

### Q: 工作流失败怎么办？
A: 
1. 查看 Actions 日志定位错误原因
2. 常见原因：编译错误、测试失败、权限问题
3. 修复后重新推送或重新运行工作流

### Q: 如何修改发布说明？
A: 
1. 编辑 `.github/workflows/build.yml` 中的 `body` 字段
2. 或发布后在 GitHub Releases 页面手动编辑

### Q: 如何添加新的平台支持？
A:
1. 在 `build.yml` 的 `matrix` 中添加新平台
2. 添加对应的依赖安装步骤
3. 调整打包逻辑

### Q: 如何自定义发布包内容？
A:
修改 `build.yml` 中的 "Package" 步骤：
```yaml
- name: Package (Windows)
  if: matrix.os == 'windows-latest'
  run: |
    # 自定义打包逻辑
    mkdir custom-package
    # ... 复制文件
```

## 🔧 高级配置

### 自定义编译选项
在 `build.yml` 中修改 CMake 配置：
```yaml
- name: Configure CMake
  run: |
    cmake -B build \
      -DCMAKE_BUILD_TYPE=${{ matrix.build_type }} \
      -DTINAXLSX_BUILD_SHARED=OFF \
      -DTINAXLSX_CUSTOM_OPTION=ON \  # 添加自定义选项
      # ...
```

### 添加代码签名（可选）
```yaml
- name: Sign binaries (Windows)
  if: matrix.os == 'windows-latest'
  run: |
    # 添加代码签名逻辑
    signtool sign /f certificate.pfx /p password binary.exe
```

### 集成外部服务
```yaml
- name: Notify external service
  run: |
    curl -X POST "https://api.example.com/notify" \
      -d '{"version": "${{ github.ref_name }}", "status": "released"}'
```

## 🎉 成功！

配置完成后，您的项目将拥有：
- ✅ 自动化多平台编译
- ✅ 自动化测试
- ✅ 自动化发布
- ✅ 专业的发布包
- ✅ 用户友好的下载体验

每次推送版本标签时，GitHub Actions 将自动为您完成所有构建和发布工作！ 