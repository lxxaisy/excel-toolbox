# Excel Toolbox (BigBaby)

这是一个基于 Electron Forge + Vite 开发的桌面工具箱应用，旨在逐步替代 VBA Excel 功能。

## 🚀 本地开发 (Mac/Windows)

### 1. 安装依赖
由于我们使用 `npm` 来管理依赖，请直接运行：

```bash
# 安装项目依赖
npm install
```

### 2. 启动开发模式
```bash
npm start
```

### 3. 打包应用
```bash
npm run make
```
打包产物位于 `out/make` 目录下。

---

## 📦 如何构建 Windows 版本 (exe)

由于项目开发环境是 macOS，直接打包无法生成 Windows `.exe` 安装包。我们使用 **GitHub Actions** 自动化构建 Windows 版本。

### 步骤说明：

1. **提交代码到 GitHub**
   将本项目代码推送到 GitHub 仓库的 `main` 或 `master` 分支。

2. **自动触发构建**
   推送代码后，GitHub Actions 会自动开始构建。
   - 点击仓库顶部的 **"Actions"** 标签页。
   - 你会看到名为 **"Build & Release"** 的工作流正在运行。

3. **下载构建产物**
   - 等待工作流（Workflow）运行完成（通常显示为绿色 ✅）。
   - 点击该次运行记录。
   - 在页面底部的 **"Artifacts"** 区域，点击 **"windows-build"** 进行下载。
   - 下载的是一个 `.zip` 压缩包，解压后即可看到 Windows 安装程序（如 `setup.exe` 或 `.nupkg` 文件）。

### 手动触发构建
如果你不想提交代码，也可以手动触发：
1. 进入 **Actions** 页面。
2. 选择左侧的 **"Build & Release"**。
3. 点击右侧的 **"Run workflow"** 按钮。

---

## 🛠 技术栈
- **Electron Forge**: 应用打包与脚手架
- **Vite**: 极速构建工具
- **Vue 3**: 前端 UI 框架
- **Node.js**: 后端逻辑 (Main Process)
- **xlsx**: Excel 处理库

## 📂 目录结构
- `src/main.js`: 主进程（Node.js 逻辑，Excel 处理）
- `src/preload.js`: 预加载脚本（安全 API 桥接）
- `src/renderer.js`: 渲染进程（Vue 入口）
- `src/App.vue`: Vue 根组件（包含侧边栏布局）
- `src/components/`: 功能组件目录
- `.github/workflows/`: CI/CD 自动化配置
