[English](./README.md) | **简体中文**

<div align="center">

<img src="public/icons/icon.svg" width="112" alt="Slilot logo"/>

# Slilot

**自托管的 PowerPoint AI 助手 —— 对话即改稿，直接读写你当前打开的演示文稿**

![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20PowerPoint%20%E6%A1%8C%E9%9D%A2%E7%89%88-blue)
![Node](https://img.shields.io/badge/Node.js-%E2%89%A5%2018-green)
![License](https://img.shields.io/badge/license-MIT-green)
![PRs](https://img.shields.io/badge/PRs-welcome-orange)

在 PowerPoint 侧边栏里用自然语言对话，模型直接在**当前打开的演示文稿**上建页、写文本、生成并插入配图，再逐页截图自检修复。
整个插件跑在你自己的电脑上，接你自己选的大模型服务，文档与 API Key 都不出本机。

<!-- TODO: 录一段 10 秒演示动图（对话 → 实时改稿）放到 docs/demo.gif，取消下行注释
![演示](docs/demo.gif)
-->

</div>

## 特性

- **实时读写文档本体**：模型通过 Office.js / PowerPoint COM 直接操作打开的演示文稿，改的就是文件本身，全程可在 PowerPoint 里撤销、保存
- **11 个精细工具**：读概览 / 读某页 / 改文本 / 改样式 / 插文本框 / 加页 / 删形状 / 生图 / 插图 / 截图 / 逐页视觉审查
- **真实截图审查（RIP 循环）**：每页做完自动用 PowerPoint 渲染成截图让模型"亲眼看"，检查重叠、溢出、变形、风格一致性后当场修复，而不是纯公式脑补；并有确定性数值自检（文本截断/溢出/越界），未清零不得通过
- **插图不变形**：模型给的是摆放区域，插图按原图宽高比等比缩放并在区域内居中
- **中英双语界面**：设置内一键切换（默认中文）
- **上游随便换**：面板统一 Anthropic 工具协议，本地服务内置翻译层，接任意提供 `/v1/messages` 或 `/v1/responses` 的服务商，不绑定任何厂商
- **上下文自动管理**：长任务超限时自动把最早的工具明细折叠成交接摘要，关键数据（面积、台数、负责人）保留，任务不中断

## 前置条件

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（侧载、COM 渲染、开机脚本均依赖 Windows） |
| Office | **桌面版 PowerPoint**（Office 2021 / Microsoft 365 已验证）。网页版、Mac 版不支持 |
| Node.js | ≥ 18（[nodejs.org](https://nodejs.org) 下载 LTS，终端 `node -v` 确认）。零依赖，**无需 npm install** |
| 模型服务 | 自备任意 LLM API Key（需要生图时该 Key 须开通图像生成） |
| 管理员权限 | 全程不需要（仅备用共享目录方案例外） |

## 安装

### 方式一（推荐）：让 AI agent 帮你装

把仓库 clone 到本地，对电脑里的 AI agent 说：

> 给我的 PowerPoint 安装这个插件：https://github.com/x1han/slilot

agent 会按仓库里的 README 与脚本依次完成：

1. 检查 Node ≥ 18 与桌面版 PowerPoint；
2. 运行 `npx office-addin-dev-certs install --days 3650` 信任 localhost 开发证书（有系统弹窗，选信任）。`--days` 很重要：默认有效期只有 30 天；
3. 运行 `install-sideload.ps1`（写入 HKCU 开发者注册表，免管理员）；
4. 启动 `node server.js`，并用 `https://localhost:3010/healthz` 返回 ok 验证；
5. 可选：在 `autostart-hidden.vbs` 上创建快捷方式，把快捷方式放进启动文件夹（Win+R → `shell:startup`）实现开机自启。

**Agent 也能自动完成配置**——本地服务暴露了一组小接口，Agent 可以自行探索并填写设置，无需手动填表：

| 本机接口（https://localhost:3010） | 用途 |
|---|---|
| GET /api/settings | 读取当前配置 |
| POST /api/settings | 合并写入配置（upstreamBase / apiFormat / apiKey / model / imageBase / imageFormat / imageKey / imageModel） |
| POST /api/probe/models | 列出服务商可用的模型（body: base, key） |
| POST /api/probe/chat | 聊天连通性（body: base, format, key, model） |
| POST /api/probe/vision | 识图真实性——服务端随机生成图形并与真值比对 |
| POST /api/probe/image | 生图连通性（自动探测两种生图端点） |

对 Agent 说一句即可：通过 localhost:3010 的探测接口探索我的服务商支持哪些模型与能力，选出文本识图模型与生图模型，三项探测通过后把配置写入 /api/settings。

面板打开时与每 30 秒会自动从本地服务同步配置——Agent 写入后，重开面板（或稍等片刻）即生效。注意：GET /api/settings 会返回 API Key；整组接口仅绑定 127.0.0.1、仅信任本机调用。

### 方式二：手动安装

```powershell
# 1. 信任开发证书（先于服务启动，否则 server.js 会因缺证书直接退出；
#    默认有效期只有 30 天，--days 3650 免去每月续期）
npx office-addin-dev-certs install --days 3650

# 2. 启动本地服务（或双击 start-addin.bat）
node server.js

# 3. 侧载注册（免管理员），然后重启 PowerPoint
powershell -ExecutionPolicy Bypass -File install-sideload.ps1

# 4.（可选）开机自启：右键 autostart-hidden.vbs 创建快捷方式，
#    把快捷方式移入启动文件夹（不要移动 vbs 本身，它按自身位置定位仓库）
```

重启 PowerPoint 后，「开始」选项卡最右侧出现 **Slilot** 按钮。

## 首次配置

打开面板右下角 ⚙，填入你的模型服务信息：

| 接口格式 | 拼接后的完整地址 |
|---|---|
| Anthropic — /v1/messages | `<base>/v1/messages` |
| OpenAI — /v1/responses | `<base>/v1/responses` |

（`<base>` 即上方填写的上游 Base。）

- 设置对话框分两块。**文本识图模型**（聊天与截图审查用）：上游 Base / 接口格式 / API Key / 文本识图模型 id。**生图模型**（插图用）：生图 Base / 生图 API Key——两者留空自动沿用文本块的——外加必填的生图模型 id（与文本识图模型相互独立，如 MiniMax 需 `image-01`）。
- 面板界面本身中英双语：点「设置」标题旁的按钮切换（默认中文）。
- 所有输入框的灰字都是示例占位（api.example.com / your-model-id / your-image-model），没有任何预配置——测试前请填入你自己的服务商。
- 点「**测试**」会并行验证三项：**聊天 / 识图 / 生图**，每行句首转圈、完成后变绿点（通过）或红点（失败）；三项全绿后才能「保存」（改任何字段都需重新测试）。识图检查会随机生成形状/颜色/位置并与真值比对，无法靠猜通过。
- 生图端点自动探测：先 `<base>/v1/images/generations`（OpenAI 标准），404 再试 `<base>/v1/image_generation`（MiniMax 风格），成功端点会被记住；生图 Base 也可以直接填完整端点（识别后不再拼接）。慢生图模型可用：测试限时 5 分钟、实际生图 10 分钟。

配置好后直接用中文提需求，例如：

> 读完这份 PPT，在最后加一页总结，风格保持一致
> 给第 1 页生成一张 16:9 封面图，简约风
> 把第 3 页的三个要点改成 2×3 网格布局，配一张示意图

## 工作原理

```
PowerPoint 任务窗格 (public/taskpane.html/js)
    │  同源请求 https://localhost:3010/api/forward（相对地址）
    ▼
本地服务 server.js (Node ≥ 18, 零依赖, HTTPS + 开发证书)
    │  通用转发 + 格式翻译（Anthropic ↔ OpenAI 双向）
    ▼
任意 https 上游
    ├─ Anthropic /v1/messages        透传
    └─ OpenAI   /v1/responses        请求/响应双向翻译
```

- **工具回路**：面板内是统一的 Anthropic 工具协议 Agent 回路，与上游格式解耦；工具经 Office.js（文本、建页、读形状）与 COM（精确插图、整稿截图）两条通道执行。
- **截图审查**：`scripts/export-slides.ps1` 把当前文稿逐页导出为 JPG，模型逐页查看真实渲染效果后修复问题（占位符删除不了会自动安全降级）。
- **长任务不爆上下文**：截图只保留最近 1 张；历史超阈值自动折叠最早的工具明细为交接摘要；熔断保护兜底。

<details>
<summary>目录结构</summary>

| 文件 | 作用 |
|---|---|
| `manifest.xml` | 标准 Office 加载项清单（TaskPaneApp） |
| `server.js` | 本地静态服务 + 通用反向代理 + 格式翻译层（端口 3010） |
| `public/taskpane.*` | 聊天界面、11 个工具实现、设置与三项测试 |
| `public/blank.pptx` | `add_slides` 的最小模板（insertSlidesFromBase64 用） |
| `public/office.js 等` | Office.js CDN 失效时的本地兜底副本（微软官方文件） |
| `scripts/*.ps1` | 截图渲染 / 精确插图 / 图标生成（PowerPoint COM） |
| `install-*.ps1` | 侧载注册（HKCU 免管理员；备用共享目录方案需管理员） |
| `autostart-hidden.vbs` | 开机自启（隐藏窗口运行本地服务） |
| `start-addin.bat` | 手动启动（可见窗口，调试用） |
| `tests/` | 格式翻译层的手工验证 payload |
| `logs/` | 诊断日志（面板错误自动回传，已 gitignore） |

</details>

## 安全与隐私

- API Key 只保存在本机——面板浏览器存储与本目录的 `settings.json`（已 gitignore）——**不写入仓库代码、不上传**；本地服务仅内存转发、无任何遥测。
- 你的演示文稿内容只发送到**你自己填写的上游地址**（本地服务强制 https）；换模型就是换服务商。
- `server.js` 数百行、零依赖，欢迎自行审计；仅监听 `127.0.0.1`，并拒绝跨源请求。
- 建议使用可信的模型服务商；备用共享目录方案会创建 SMB 共享，装完可执行 `net share addincatalog /delete` 清理。

## 故障排查

| 现象 | 处理 |
|---|---|
| PowerPoint 里没有 Slilot 按钮 | 确认服务在跑（`https://localhost:3010/healthz` 返回 ok）→ 重启 PowerPoint → 仍没有则以管理员运行 `install-shared-catalog-admin.ps1` 后重启 |
| 面板打不开 / 一直初始化 | 看 `logs/client-log.txt`；多为本地服务没启动或证书未信任 |
| “开始”选项卡启动时没有按钮 | 永久版 Office 2021 开发者侧载的已知限制：命令注册发生在本会话首次打开加载项时从【加载项】菜单打开一次即可。（图标托管在 GitHub Pages，需能访问 x1han.github.io） |
| 证书过期（面板突然打不开） | 开发证书默认有效期只有 30 天。刷新：`npx office-addin-dev-certs install --days 3650`，然后重启本地服务 |
| GitHub 连不上（clone 失败） | 部分网络环境直连 github.com 被阻断——用 SSH clone（`git clone git@github.com:x1han/slilot.git`，需在 GitHub 账号配置 SSH key）或代理。其余环节均不依赖 GitHub：图标/FunctionFile 在 x1han.github.io（可直连）、office.js 有本地回退、模型上游是你自己的服务商 |
| 测试不通过 | 聊天失败查上游地址 / 格式 / Key；识图失败说明该模型不支持图片输入，换模型；生图失败通常是 Key 未开通图像生成 |
| 改了代码没生效 | 需重启本地服务，并把面板 × 掉重开；若改了 `manifest.xml`，把其中 `<Version>` 加一位再重启 PowerPoint |
| 插图 / 截图报 COM 错误 | 确认 PowerPoint 打开的是目标文稿（COM 附着"当前活动演示文稿"）；关闭 PowerPoint 里阻塞的弹窗后重试 |

## 卸载

- 删除注册表值 `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` 下以清单 Id（`a7f3d9e2-…`）命名的项，重启 PowerPoint 按钮即消失；
- 删除启动文件夹里的 Slilot 快捷方式，停掉 node 进程；
- 用过共享目录方案的话，执行 `net share addincatalog /delete` 并删除 `C:\addin-catalog`。

## 能力边界

- 仅支持 **Windows 桌面版 PowerPoint**；网页版 / Mac / WPS 不可用。
- PowerPoint API 覆盖不到的操作（动画、切换、母版设计、SmartArt 内部、图表数据编辑）做不了。
- 面板会探测宿主 JS 能力并自动裁剪工具集；逐页审查基于真实截图，要求所用模型支持图片输入（设置里可测试）。

## 许可证

本项目基于 [MIT](./LICENSE) 发布。

---

觉得有用的话点个 Star ⭐ ；欢迎提 Issue 与 PR。
