# AI 写作助手 Word 插件

> 基于 AI 驱动的 Microsoft Word 写作辅助插件

---

## ✨ 功能菜单

| 功能 | 说明 | 子功能 |
|------|------|--------|
| ⚡ 触发 AI 行内补全 | 快捷键 Ctrl+J | — |
| ✏️ 编辑 | AI 辅助编辑 | 优化写作 / 修复拼写与语法 / 缩短内容 / 扩展内容 / 简化表达 |
| 💬 语气 | 调整文字语气 | 专业语气 / 口语语气 |
| 🌐 翻译 | 多语言智能翻译 | 中 / 英 / 日 / 韩 / 法 / 德 / 西 / 俄 |
| 🎯 生成 | AI 内容生成 | 解释这段内容 / 总结内容 / 继续写作 |
| 🪄 AI 自定义改写 | 按指令改写内容 | — |
| ⚙️ API 设置 | 配置 AI 服务 | — |

> ⚠️ 已按需求移除「AI 百宝箱」菜单项

---

## 🚀 安装步骤

### 方法一：开发调试安装（推荐用于测试）

#### 前提条件
- Node.js 18+ 已安装
- Microsoft Word 已安装（Microsoft 365 版本）

#### 步骤

**1. 安装依赖**
```powershell
cd "E:\vscode\word AI 插件"
npm install
```

**2. 安装开发证书**
```powershell
npx office-addin-dev-certs install
```

**3. 启动开发服务器**
```powershell
npm start
```

**4. 在 Word 中侧载插件**

打开 Word → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录

添加目录路径：
```
E:\vscode\word AI 插件
```

或者使用命令行自动侧载：
```powershell
npx office-addin-debugging start manifest.xml
```

---

### 方法二：生产部署安装

**1. 构建生产版本**
```powershell
npm run build
```

**2. 将 `dist/` 目录部署到 HTTPS Web 服务器**

**3. 更新 `manifest.xml` 中的 URL**  
将所有 `https://localhost:3000` 替换为您的服务器地址

**4. 通过组织共享（IT 管理员操作）**
- Microsoft 365 管理中心 → 设置 → 集成应用
- 上传 `manifest.xml`
- 分配给用户

---

### 方法三：网络共享文件夹（最简单）

1. 将项目放到共享网络文件夹（或本地文件夹）
2. Word → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录
3. 添加文件夹路径，勾选「在菜单中显示」
4. 重启 Word → 插入 → 我的加载项 → 共享文件夹 → 选择「AI 写作助手」

---

## ⚙️ API 配置

点击插件菜单中的 **⚙️ API 设置** 按钮，配置以下参数：

| 参数 | 说明 |
|------|------|
| **AI 服务提供商** | OpenAI / Azure OpenAI / 自定义兼容接口 |
| **API Key** | 您的 API 密钥，加密存储在本地 |
| **Base URL** | API 端点地址（默认：https://api.openai.com/v1）|
| **模型** | GPT-4o（推荐）/ GPT-4o Mini / GPT-4 / 自定义 |
| **创造性** | Temperature 值（0=严谨，1=创意，默认 0.7）|
| **最大输出** | Max Tokens（默认 2048）|

### 支持的 AI 服务

- **OpenAI**：使用 https://platform.openai.com 获取 API Key
- **Azure OpenAI**：需要 Azure 订阅和 OpenAI 资源
- **自定义**：任何兼容 OpenAI API 格式的服务（如 DeepSeek、文心一言等）

---

## 📁 项目结构

```
word AI 插件/
├── manifest.xml              # 插件清单（Word 注册文件）
├── package.json              # Node.js 依赖配置
├── webpack.config.js         # 构建配置
├── assets/
│   ├── icon-16.png           # 工具栏图标 (16x16)
│   ├── icon-32.png           # 工具栏图标 (32x32)
│   └── icon-80.png           # 工具栏图标 (80x80)
└── src/
    ├── taskpane/
    │   ├── taskpane.html     # 主任务面板界面
    │   ├── taskpane.css      # 深色玻璃态样式
    │   └── taskpane.js       # 主逻辑（AI调用+文档操作）
    ├── settings/
    │   ├── settings.html     # API 设置界面
    │   ├── settings.css      # 设置页样式
    │   └── settings.js       # 设置保存/加载逻辑
    └── commands/
        ├── commands.html     # 功能区命令文件
        └── commands.js       # 行内补全等直接命令
```

---

## 🛠️ 常见问题

**Q: Word 显示「加载项不受信任」**  
A: 需要安装开发证书：`npx office-addin-dev-certs install`

**Q: 功能无响应**  
A: 确保开发服务器正在运行（`npm start`），且 Word 可以访问 `https://localhost:3000`

**Q: AI 返回错误**  
A: 检查 API Key 是否正确，可在「API 设置」中点击「测试连接」验证
