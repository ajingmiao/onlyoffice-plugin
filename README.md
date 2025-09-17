# OnlyOffice 插件 Demo

本项目为 OnlyOffice 插件的极简示例，展示了插件的基本结构与核心功能实现，适合二次开发和学习参考。

## 目录结构

```text
sdkjs-plugins/demo-clean/
├─ config.json
├─ index.html
├─ plugin.js                  // 入口：仅负责装配与启动
├─ core/
│  ├─ logger.js               // 统一日志
│  ├─ constants.js            // 命令字、事件名、Key等
│  ├─ utils.js                // 小工具（防抖、tryWrap等）
├─ bridge/
│  ├─ host-bridge.js          // 宿主通信（Common.Gateway）
│  └─ plugin-bridge.js        // Asc 插件运行时适配（init/callCommand/事件）
├─ app/
│  ├─ command-bus.js          // 命令总线（派发 insertText 等）
│  ├─ event-bus.js            // 事件总线（选择变化等，供 UI/宿主订阅）
│  ├─ container.js            // 依赖注入/装配（极简）
│  └─ lifecycle.js            // 启动/销毁流程
├─ services/
│  ├─ editor-service.js       // 文档通用操作（GetDocument/Range等）
│  ├─ sdt-service.js          // 内容控件相关（创建/查找/标记）
│  └─ selection-service.js    // 选择变更感知与信息提取
└─ handlers/
   ├─ insert-text.js          // 用例：插入占位符（行内控件）
   └─ report-active-sdt.js    // 用例：回传当前激活SDT信息
```

## 主要模块说明

- **plugin.js**：插件入口，负责装配与启动。
- **core/**：核心工具与常量，包括日志、命令字、工具函数等。
- **bridge/**：与宿主环境和插件运行时的通信桥接。
- **app/**：命令与事件总线、依赖注入、生命周期管理。
- **services/**：文档操作、内容控件、选择感知等服务。
- **handlers/**：具体业务用例处理，如插入占位符、回传激活信息。

## 快速开始

1. 克隆本仓库到本地。
2. 按需修改 `config.json` 和相关 JS 文件。
3. 参考 OnlyOffice 官方文档，将插件部署到 OnlyOffice 环境中。

## 参考文档
- [OnlyOffice 插件开发文档](https://api.onlyoffice.com/)

---
如需定制开发或有任何疑问，欢迎 issue 交流。

