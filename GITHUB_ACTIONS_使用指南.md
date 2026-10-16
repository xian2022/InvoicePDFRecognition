# GitHub Actions 自动打包指南

本项目已配置 GitHub Actions，可自动在云端打包 macOS 和 Windows 两个版本的可执行文件。

## 触发条件

以下情况会自动触发打包：
- ✅ 推送到 `main` 或 `master` 分支
- ✅ 创建以 `v` 开头的标签（如 `v1.0.0`）
- ✅ 手动触发（在 GitHub Actions 页面点击 "Run workflow"）

## 使用方法

### 方式一：推送代码自动打包

```bash
# 提交代码并推送
git add .
git commit -m "更新功能"
git push origin main
```

推送后 GitHub 会自动开始打包，约 5-10 分钟完成。

### 方式二：打标签发布正式版

```bash
# 创建标签
git tag v1.0.0
git push origin v1.0.0
```

这会自动打包并创建 Release，可直接下载发布的可执行文件。

### 方式三：手动触发

1. 进入 GitHub 仓库页面
2. 点击 "Actions" 标签
3. 选择 "Build Invoice Processor"
4. 点击 "Run workflow"

## 下载打包结果

### 开发版（每次推送）

1. 进入 GitHub 仓库 → Actions
2. 找到对应的运行记录
3. 点击运行记录，在页面底部 "Artifacts" 区域下载：
   - `invoice-processor-macos` - macOS 版本
   - `invoice-processor-windows` - Windows 版本

### 正式版（打标签）

1. 进入 GitHub 仓库 → Releases
2. 找到对应版本
3. 在 Assets 区域下载：
   - `发票提取工具-macOS` - macOS 版本
   - `发票提取工具-Windows.exe` - Windows 版本

## 注意事项

1. **首次使用**需要先将项目上传到 GitHub：
   ```bash
   # 初始化 Git 仓库
   git init
   git add .
   git commit -m "Initial commit"
   
   # 添加远程仓库（替换为你的仓库地址）
   git remote add origin https://github.com/你的用户名/invoice-processor.git
   git push -u origin main
   ```

2. **Artifacts 保留期限**：30 天后自动删除，建议及时下载或通过 Release 永久保存

3. **打包时间**：macOS 和 Windows 并行打包，总计约 5-10 分钟

4. **费用**：GitHub Actions 对公开仓库免费，私有仓库每月有 2000 分钟免费额度

## 本地打包（备用）

如果不想用 GitHub Actions，也可以在本地打包：

- **macOS**: `./build.sh`
- **Windows**: 双击 `build.bat`（需在 Windows 上执行）
