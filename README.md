# AutoSyncVBA 插件使用说明书

## 插件简介

AutoSyncVBA（曾用名：VBA Sync）是一个 VS Code 扩展，用于在 Office 文档（Excel、Word、PowerPoint）和 IDE 之间实现双向 VBA 代码同步。它可以帮助开发者在 VS Code 中编辑 VBA 代码，并自动同步到 Office 文档中，同时也可以从文档中导出 VBA 代码到 IDE 中。

## 安装说明

### 从源码构建

1. 克隆或下载源码到本地
2. 打开终端，进入 `dev` 目录
3. 运行 `npm install` 安装依赖
4. 运行 `npm run build` 构建插件
5. 运行 `npm run package` 生成 VSIX 文件
6. 在 VS Code 中，点击左侧边栏的扩展图标（Extensions）
7. 点击右上角的三个点，选择 "Install from VSIX..."
8. 选择生成的 VSIX 文件
9. 安装完成后，重启 VS Code

## 环境要求

1. **VS Code**：版本 1.85.0 或更高
2. **Python**：版本 3.7 或更高
3. **Office**：支持 Excel、Word、PowerPoint（2010 或更高版本）
4. **Python 依赖**：
   - `win32com.client`（用于与 Office 交互）
   - `pywin32`（Windows 平台）

## 使用方法

### 启用同步

1. 在 VS Code 中打开一个工作区
2. 在文件资源管理器中，右键点击 Office 文档（.xlsm、.docm、.pptm 等）
3. 选择 "VBA: 启用同步"
4. 插件会自动：
   - 打开 Office 文档（后台运行）
   - 在工作区中创建 `vba/文档名称` 文件夹
   - 导出文档中的所有 VBA 模块到该文件夹
   - 开始监控文件变化

### 同步文件到文档

1. 在 VS Code 中编辑 VBA 文件（.bas、.cls）
2. 保存文件后，插件会自动同步到 Office 文档
3. 或者，右键点击 VBA 文件，选择 "VBA: 同步到文档"

### 从文档同步到 IDE

1. 当 Office 文档中的 VBA 代码被修改时，插件会检测到变更
2. 弹出提示询问是否同步到 IDE
3. 点击 "同步" 按钮，将变更同步到 VS Code 中的文件

### 其他命令

- **VBA: 禁用同步**：停止同步并关闭 Office 文档
- **VBA: 新建模块**：在当前同步的文档中创建新的 VBA 模块
- **VBA: 立即同步所有**：手动同步所有启用的文档
- **VBA: 打开调试窗口**：显示 Office 文档的 VBA 编辑器
- **VBA: 最小化到后台**：将 Office 文档窗口最小化到后台
- **VBA: 导出所有模块到 IDE**：手动导出文档中的所有 VBA 模块
- **VBA: 查看同步状态**：查看当前同步状态和统计信息
- **VBA: 重置同步计数**：重置待调试的变更计数
- **VBA: 打开配置**：打开插件配置文件

## 配置选项

插件的配置选项可以通过 VS Code 的设置界面修改：

1. 点击 VS Code 左下角的设置图标
2. 选择 "设置"
3. 在搜索框中输入 "VBA Sync"

主要配置选项：

| 选项 | 默认值 | 描述 |
|------|--------|------|
| `vbaSync.defaultFolder` | `vba` | 默认导出文件夹名 |
| `vbaSync.autoSync` | `true` | 是否自动同步 |
| `vbaSync.conflictStrategy` | `ask` | 冲突策略（ask、ide-priority、document-priority） |
| `vbaSync.showNotifications` | `true` | 是否显示通知 |
| `vbaSync.syncDelay` | `500` | 文件变化后延迟同步时间（毫秒） |
| `vbaSync.pythonPath` | `"` | 自定义 Python 路径 |
| `vbaSync.excludePatterns` | `["**/node_modules/**", "**/.git/**"]` | 排除监控的文件模式 |
| `vbaSync.debugPrompt.onNewSubFunction` | `true` | 有新增过程时提示 |
| `vbaSync.debugPrompt.pendingChangesThreshold` | `5` | 累积 N 次变更时提示 |
| `vbaSync.debugPrompt.timeSinceLastDebug` | `1800000` | 距上次调试超时（毫秒） |

## 故障排除

### 常见问题

1. **Python 环境未就绪**
   - 确保已安装 Python 3.7 或更高版本
   - 确保已安装 `pywin32` 包：`pip install pywin32`

2. **无法打开 Office 文档**
   - 确保 Office 已安装且可以正常运行
   - 确保文档未被其他进程锁定
   - 检查文档是否有密码保护

3. **同步失败**
   - 检查 Python 路径是否正确
   - 检查 Office 文档是否已打开
   - 查看 VS Code 输出面板中的错误信息

4. **文件路径安全验证失败**
   - 确保 VBA 文件在指定的同步文件夹内
   - 避免使用相对路径或路径遍历

### 查看日志

1. 在 VS Code 中，点击 "查看" -> "输出"
2. 在输出面板的下拉菜单中选择 "VBA Sync"
3. 查看详细的日志信息，帮助排查问题

## 性能优化

1. **减少同步延迟**：在配置中调整 `syncDelay` 值
2. **排除不必要的文件**：在 `excludePatterns` 中添加不需要监控的文件模式
3. **关闭自动同步**：如果不需要实时同步，可以将 `autoSync` 设置为 `false`

## 注意事项

1. **Office 文档兼容性**：仅支持 .xlsm、.xlsb、.xltm、.pptm、.potm、.ppsm、.docm、.dotm 格式
2. **文件大小限制**：大型 VBA 项目可能会影响同步性能
3. **网络环境**：在网络驱动器上的文件可能会导致同步延迟
4. **版本控制**：建议使用 Git 等版本控制系统管理 VBA 代码

## 版本历史

### v2.0.0
- 完全重写插件架构，采用模块化设计
- 新增配置验证功能
- 优化文件比较算法
- 改进 Python 进程启动逻辑
- 增强错误处理和用户友好的提示
- 添加文件路径安全验证

## 联系与支持

如果您遇到任何问题或有任何建议，请通过以下方式联系我们：

- **GitHub**：在插件仓库中提交 issue
- **电子邮件**：[support@vba-sync.com]

---

**版权所有 © 2026 AutoSyncVBA 团队**
**保留所有权利**