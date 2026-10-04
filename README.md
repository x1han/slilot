# AI PPT 助手（自建 PowerPoint 侧载加载项）

在 PowerPoint 侧边栏里与 AI 模型对话，模型通过 Office.js 工具**实时读写当前打开的演示文稿**。支持生图插图；本地服务内置格式翻译层，上游可接任意提供 `/v1/messages`、`/v1/responses` 或 `/v1/chat/completions` 的服务商。**不绑定任何厂商。**

## 架构

```
PowerPoint 任务窗格 (taskpane.html/js)
    │  同源请求 https://localhost:3010/api/forward（相对地址）
    ▼
本地服务 server.js (Node ≥18, 零依赖, HTTPS + 微软开发证书)
    │  通用转发 + 格式翻译（按 x-upstream-url / x-api-format 头路由）
    ▼
任意 https 上游
    ├─ Anthropic /v1/messages        透传
    ├─ OpenAI   /v1/responses        请求/响应双向翻译
    └─ OpenAI   /v1/chat/completions 请求/响应双向翻译
```

- 面板内部统一 Anthropic 方言，工具调用回路与上游格式解耦
- 模型收到的工具：读概览 / 读某页 / 改文本 / 改样式 / 插文本框 / 加页 / 删形状 / 生图 / 插图
- 工具由 taskpane.js 用 Office.js 在 PowerPoint 进程内执行，改的就是文档本体（可撤销、可保存）
- API Key 只保存在面板的 localStorage，不写入代码、不落盘

## 文件

| 文件 | 作用 |
|---|---|
| `manifest.xml` | 加载项清单（AppSource 兼容格式） |
| `server.js` | 本地静态服务 + 通用反向代理 + 格式翻译层（端口 3010） |
| `public/taskpane.*` | 聊天界面、工具实现、设置/测试界面 |
| `public/office.js 等` | Office.js CDN 失效时的本地兜底副本 |
| `autostart-hidden.vbs` | 隐藏窗口启动服务（复制到用户"启动"文件夹即开机自启） |
| `start-addin.bat` | 手动启动（可见窗口，调试用） |
| `install-sideload.ps1` | 写入 HKCU 开发者注册表完成侧载（免管理员） |
| `install-shared-catalog-admin.ps1` | 备用侧载方案（需管理员跑一次） |
| `logs/client-log.txt` | 面板"黑匣子"日志（页面 JS 错误自动回传，已 gitignore） |

## 使用

1. 本地服务开机自启（把 `autostart-hidden.vbs` 放进 `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`），或手动双击 `start-addin.bat`。
2. 首次运行 `npx office-addin-dev-certs install` 信任 localhost 开发证书。
3. 侧载：运行 `install-sideload.ps1`（免管理员），重启 PowerPoint，「开始」选项卡最右侧出现 **AI → AI 助手**。
4. 打开面板 ⚙，填入上游地址 / 接口格式 / API Key / 模型 → 点「测试」→ 两项全绿后「保存」。
5. 直接用中文提需求，例如"读完这份 PPT，在最后加一页总结"、"给第 1 页生成一张封面图，16:9"。

### 接入示例

| 厂商 | 上游 Base | 接口格式 |
|---|---|---|
| Anthropic 兼容网关 | `https://<host>/<anthropic路径>` | Anthropic /v1/messages |
| OpenAI Responses 风格 | `https://<host>` | OpenAI /v1/responses |
| OpenAI Chat 风格（DeepSeek/Moonshot/智谱等） | `https://<host>` | OpenAI /v1/chat/completions |

生图独立配置（`imageApiUrl` / `imageModel`，在 taskpane.js 顶部改），默认指向已验证可用的图像生成端点；不需要生图时可忽略（「测试」会提示该项不可用，聊天不受影响）。

## 故障排查

- **按钮没出现**：确认 3010 服务在跑（`https://localhost:3010/healthz` 返回 ok）→ 重启 PowerPoint → 仍没有则以管理员运行 `install-shared-catalog-admin.ps1` 后重启。
- **面板没反应**：看 `logs/client-log.txt`，页面错误都会记录在那里。
- **测试不通过**：按提示区分——聊天失败查地址/格式/Key；生图失败通常是该 Key 未开通图像生成。
- **更新代码后没生效**：把 `manifest.xml` 的 `<Version>` 加一位并重启 PowerPoint（Office 缓存清单）；共享目录副本 `C:\addin-catalog\manifest.xml` 记得同步。
- **卸载**：删除注册表值 `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` 下以清单 Id 命名的项。

## 能力边界

PowerPointApi 覆盖不到的操作（动画、切换、母版设计、SmartArt 内部、图表数据）做不了；
面板顶部会显示宿主实际支持的 API 集版本，工具按能力自动裁剪。
