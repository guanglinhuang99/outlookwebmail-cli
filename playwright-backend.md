# Playwright Backend 方案与实现说明

> 状态：方案 1 已实现。本文保留设计决策、模块边界和验收标准；实际安装、命令和运行模式请以 [README.md](README.md) 为准。

## 1. 目标与边界

本方案为 `webmail-cli` 增加 Playwright 浏览器后端，并保留 Ego Lite 作为可选后端。目标是：

- 不要求安装 Ego Lite App；
- 支持 Windows、macOS、Linux；
- 复用浏览器登录态，但不读取密码、不导出 Cookie、不依赖 Outlook 私有 API；
- 保持现有 `OutlookService`、CLI 命令和标准化邮件数据结构不变；
- 让 Playwright 后端和现有 `BrowserBackend` 接口可互换；
- 对删除、移动、回复、发送等有副作用的操作继续使用现有确认、草稿和 request-id 安全机制。

当前实现默认使用 `auto` 后端：优先使用 Playwright，启动失败时尝试 Ego Lite。需要强制整个 CLI 或 MCP 使用 Ego Lite 时，设置 `WEBMAIL_MODE=egolite`，或在单次 CLI 命令前加 `--mode egolite`；该模式会覆盖 Playwright、CDP、Edge 共享和 Profile 配置。

## 2. 技术决策

### 2.1 依赖选择

方案 1 使用 `playwright-core`，而不是完整的 `playwright`：

```text
playwright-core + 用户已经安装的 Chrome / Edge / Chromium
```

这样 npm 包不会自动下载一套数百 MB 的浏览器，适合本项目的轻量、跨平台目标。代价是运行机器必须已经安装受支持的浏览器，并且需要配置或自动发现浏览器可执行文件。

如果后续希望让项目自带固定版本的 Chromium，才改用完整的 `playwright`，并执行 `npx playwright install chromium`。这会降低系统浏览器差异，但增加安装包体积和升级维护成本。

当前依赖：

```bash
npm install playwright-core
```

Node.js 版本继续遵循项目当前的 `engines` 约束，不另起一套运行时。

### 2.2 浏览器与登录态

默认使用可见浏览器窗口（`headless: false`），因为首次运行需要用户在 Outlook 页面手工完成登录、验证码或企业认证。

登录态保存在项目专用的持久化 Profile 中：

```text
WEBMAIL_PROFILE_DIR=<本地专用目录>
```

Playwright 通过 `launchPersistentContext(profileDir, options)` 打开该目录。首次运行用户登录一次，后续运行复用该 Profile 中的 Cookie 和 localStorage。Profile 只用于本 CLI，不指向用户日常 Chrome 的默认 Profile。

不建议直接打开日常浏览器 Profile，原因是：

- Chrome/Edge 可能因 Profile 锁造成启动失败；
- CLI 退出、崩溃或升级时可能影响用户日常浏览器；
- 企业策略、扩展和浏览历史会把无关状态带入自动化；
- 误用默认 Profile 会扩大 Cookie 和隐私数据的暴露范围。

如确实需要接入已经打开的 Chromium 浏览器，另提供可选的 CDP 模式（`connectOverCDP`）。该模式要求浏览器以远程调试端口启动，只支持 Chromium 系列，自动化能力和上下文隔离能力也弱于正常 Playwright 启动，因此不作为默认路径。

## 3. 用户配置

当前配置（环境变量和 CLI 全局参数）：

| 配置 | 默认值 | 说明 |
| --- | --- | --- |
| `WEBMAIL_MODE` | `default` | `default` 保持当前配置；`egolite` 强制全程使用 Ego Lite |
| `WEBMAIL_BACKEND` | 当前兼容默认值 | `ego-lite` 或 `playwright` |
| `WEBMAIL_BROWSER` | `auto` | `edge`、`chrome`、`chromium` 或 `auto` |
| `WEBMAIL_EXECUTABLE_PATH` | 空 | 明确指定浏览器可执行文件，优先级最高 |
| `WEBMAIL_PROFILE_DIR` | 平台用户数据目录下的 `webmail-cli` | Playwright 专用持久化 Profile |
| `WEBMAIL_HEADLESS` | `false` | 调试和无人值守场景可设为 `true` |
| `WEBMAIL_CDP_ENDPOINT` | 空 | 可选的 Chromium CDP 地址；设置后不再使用 executable/profile 启动模式 |
| `WEBMAIL_SHARE_EDGE` | `false` | 复用已开启远程调试的日常 Edge；不与 `WEBMAIL_MODE=egolite` 同时生效 |
| `WEBMAIL_BROWSER_TIMEOUT_MS` | 30000 | 页面动作和等待的统一超时 |
| `WEBMAIL_URL` | `https://partner.outlook.cn/mail/` | 仅允许 Outlook 目标域名 |

配置校验规则：

1. `WEBMAIL_MODE=egolite` 优先级最高，忽略 Playwright、CDP、Edge 共享和 Profile 相关选择。
2. 默认模式按 `WEBMAIL_BACKEND` 选择后端；`auto` 优先 Playwright，失败后回退 Ego Lite。
3. `WEBMAIL_CDP_ENDPOINT` 与 `WEBMAIL_EXECUTABLE_PATH`、持久化启动模式互斥。
4. `WEBMAIL_PROFILE_DIR` 必须是专用目录，不能是 Chrome/Edge 默认用户数据根目录，也不能是仓库目录。
5. 指定的可执行文件必须存在且可执行；自动发现失败时给出安装和配置提示。
6. `WEBMAIL_URL` 只能是配置允许的 Outlook 域名，禁止把后端变成任意网站浏览器。

## 4. 跨平台浏览器发现

发现顺序如下：

1. `WEBMAIL_EXECUTABLE_PATH`；
2. 平台默认安装路径；
3. `PATH` 中的命令（Linux 优先）；
4. 全部失败时返回明确的 `BROWSER_NOT_FOUND` 错误。

建议的候选路径：

| 平台 | Edge | Chrome / Chromium |
| --- | --- | --- |
| Windows | `C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe`、`C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe` | `C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe`、`C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe` |
| macOS | `/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge` | `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`、`/Applications/Chromium.app/Contents/MacOS/Chromium` |
| Linux | `microsoft-edge`、`microsoft-edge-stable` | `google-chrome`、`google-chrome-stable`、`chromium`、`chromium-browser` |

实现时不要只硬编码一个路径：使用 `process.platform`、`fs.access` 和 PATH 查找，并把最终选择的浏览器名称（不包含敏感 Profile 内容）放入 `status --json`。

## 5. 架构与模块边界

```text
CLI
  -> OutlookService
      -> BrowserBackend 接口
          -> EgoLiteBackend（现有）
          -> PlaywrightBackend（新增）
```

实际模块：

```text
src/browser/playwright.ts          # PlaywrightBackend 生命周期与动作
src/browser/playwright-config.ts   # 配置、Profile 和可执行文件发现
src/browser/browser-factory.ts     # 根据 WEBMAIL_BACKEND 创建后端
src/browser/playwright-runner.ts   # Playwright Page/Context 操作与复用
src/browser/managed-cdp.ts         # 本地 CDP 常驻浏览器管理
src/browser/shared-edge.ts         # 日常 Edge 远程调试端口发现与连接
src/outlook/inbox-parser.ts        # 列表、目录和附件 DOM 解析
src/outlook/message-parser.ts      # 邮件详情 DOM 解析
```

`PlaywrightBackend` 只负责浏览器生命周期、页面动作、下载和 DOM 探针；邮件筛选、日期处理、目录语义、request-id 去重、审计记录仍由 `OutlookService` 和现有安全模块负责。

## 6. Playwright 生命周期

### 6.1 正常启动

伪代码如下：

```ts
const browserPath = await resolveBrowserExecutable(config);
const context = await chromium.launchPersistentContext(profileDir, {
  executablePath: browserPath,
  headless: config.headless,
  acceptDownloads: true,
  timeout: config.timeoutMs,
});

const page = context.pages()[0] ?? await context.newPage();
await page.goto(config.outlookUrl, { waitUntil: 'domcontentloaded' });
```

实际实现需要：

- 复用 Profile 中已经打开的 Outlook Page；
- 没有可用页面时新建页面并打开 `/mail/`；
- 对导航、弹窗、页面崩溃和上下文关闭注册统一错误处理；
- 使用一个后端实例管理一个 Context，避免每个命令创建多个浏览器进程；
- 在 CLI 进程结束时正常关闭 Context，但保留“需要用户手工发送”的草稿窗口。

### 6.2 登录状态

`status()` 和 `inspect()` 必须能够区分：

- 浏览器未安装；
- 浏览器已启动但 Outlook 页面未打开；
- Outlook 页面已打开且已登录；
- 页面已打开但需要登录；
- 页面加载中或 UI 尚未稳定。

如果检测到需要登录：

1. 保持可见窗口；
2. 输出“请在浏览器中完成登录，完成后回到 CLI 重试”的提示；
3. 不读取或记录密码、验证码、Cookie 或 localStorage 原文；
4. 可提供 `webmail doctor --wait-login`，轮询登录状态直到超时。

本方案不自动填写账号密码，也不绕过 MFA、验证码或企业安全策略。

### 6.3 CDP 可选模式

当设置 `WEBMAIL_CDP_ENDPOINT` 时，使用 `chromium.connectOverCDP(endpoint)` 连接已经启动的 Chromium 浏览器。该模式只作为高级选项：

- 需要用户自行以远程调试参数启动 Chrome/Edge；
- 远程调试端口必须只绑定本机或受控网络；
- 不允许在日志中输出完整 endpoint（尤其是带 token 的 URL）；
- 无法保证使用专用 Profile，安全提示必须更明显；
- 若连接失败，错误应提示关闭 CDP 配置或改用持久化 Profile 模式。

## 7. `BrowserBackend` 接口映射

以下方法保持现有返回类型和业务语义，Playwright 只替换底层执行方式。

### 7.1 基础页面动作

| 接口 | Playwright 实现 |
| --- | --- |
| `status()` | `page.url()`、`page.title()`、登录探针、Context/Page 状态 |
| `snapshot()` | `page.locator('body').innerText()`，必要时限制长度 |
| `eval<T>()` | `page.evaluate()`，仅允许内部固定探针，不接受任意用户脚本 |
| `click()` | 解析稳定 locator 后调用 `locator.click()` |
| `clickAndWait()` | click 后等待导航、网络空闲或指定状态变化 |
| `fill()` | `locator.fill()`；富文本编辑器使用 `locator.pressSequentially()` 或 DOM 事件 |
| `fillAndPress()` | fill 后调用 `locator.press(key)`，再等待页面状态 |
| `press()` | 当前页面 `keyboard.press()` |
| `scrollBy()` | `page.mouse.wheel()` 或页面内滚动容器 |
| `wheel()` / `wheelAndEval()` | 目标 locator 定位后滚动，等待虚拟列表渲染，再执行固定探针 |
| `wait()` | `page.waitForTimeout()` 仅作兜底；优先等待可观察条件 |

不允许把 CLI 参数直接拼入选择器或 JavaScript。所有 locator 都应经过固定的选择器表和文本精确匹配，避免误点相似邮件或按钮。

### 7.2 检查、列表和阅读

| 接口 | 实施要点 |
| --- | --- |
| `inspect()` | 用 `page.evaluate()` 执行页面状态探针，返回与现有 `InspectResult` 相同的字段 |
| `inspectMessage()` | 在当前邮件详情页执行正文、发件人、时间、附件探针 |
| `openAndExtractMessage(locator)` | 先按稳定消息 locator 打开，再等待详情区域，复用现有解析器 |
| `listInboxFolders()` | 打开 Inbox 文件夹树，读取名称、路径和子目录关系 |
| `selectInboxFolder(directory)` | 按规范化目录路径逐级点击，并验证当前目录标题/URL |
| 搜索/日期/目录过滤 | 继续由 `OutlookService` 传入标准化参数；后端只负责页面查询和翻页 |

目录为空时仍映射为 Inbox，日期为空时由服务层解析为今天。Playwright 后端不能自行改变这些默认值，否则 Ego Lite 和 Playwright 的结果会不一致。

### 7.3 附件和 Obsidian 导出

`downloadAttachment()` 使用 Playwright 下载事件：

```ts
const downloadPromise = page.waitForEvent('download');
await attachmentLocator.click();
const download = await downloadPromise;
await download.saveAs(safeOutputPath);
```

实施要求：

- 输出目录由服务层校验并创建；
- 文件名必须经过路径穿越、非法字符和重名处理；
- 保存后检查文件存在、大小和最终路径；
- 返回现有 `AttachmentDownloadResult`，不把临时下载路径暴露给用户；
- Obsidian Markdown 仍由现有导出层生成，附件链接使用导出目录内的相对路径；
- 如果下载按钮触发新窗口或弹出菜单，需要先等待对应 locator，再监听 `download`。

### 7.4 写操作、草稿和回复

以下方法沿用现有安全语义：

```text
composeMessage
replyMessage
forwardMessage
openDraft
updateDraft
sendDraft
discardDraft
deleteMessage
moveMessage
setReadState
setFlagState
setCategoryState
```

Playwright 侧统一遵循“动作前精确定位、动作后验证结果”：

1. 根据消息 locator 打开目标邮件；
2. 验证主题、发件人或稳定 DOM 标识与目标一致；
3. 点击回复、全部回复、转发或菜单项；
4. 填写内容并验证编辑器值；
5. `draft=true` 时保存草稿并确认草稿状态，不点击发送；
6. `draft=false` 时在服务层完成确认和 request-id 检查，再点击发送；
7. 发送后等待成功提示、草稿消失或已发送状态，并返回标准化结果。

`replyAll` 只改变按钮选择和收件人验证逻辑，不改变公共接口。不要通过猜测页面上第一个按钮执行 reply/reply-all。

“手工发送”场景不能像 Ego Lite task space 一样依赖外部任务空间。建议返回：

```json
{
  "ok": true,
  "draft": true,
  "manualActionRequired": true,
  "message": "草稿已打开，请在浏览器窗口中检查并手工发送"
}
```

CLI 进程在该模式下应保持 Context 存活，直到用户按回车确认、超时或主动取消；如果当前 CLI 设计为命令立即退出，则至少要保证草稿已经持久化到 Outlook，不能声称已发送。

## 8. 错误模型

新增后端错误应映射到公共错误结构，不把 Playwright 原始堆栈直接返回给用户。建议错误分类：

| 错误码 | 场景 | 用户提示 |
| --- | --- | --- |
| `BROWSER_NOT_FOUND` | 没有发现 Chrome/Edge/Chromium | 安装浏览器或设置 `WEBMAIL_EXECUTABLE_PATH` |
| `PROFILE_LOCKED` | 专用 Profile 正被另一个 CLI 占用 | 关闭其他 webmail-cli 实例后重试 |
| `AUTH_REQUIRED` | Outlook 需要登录 | 在可见窗口完成登录 |
| `PLAYWRIGHT_TIMEOUT` | locator、导航或下载超时 | 检查网络、登录态和 UI 状态 |
| `UI_CHANGED` | 选择器或页面结构不匹配 | 保存诊断快照后更新探针 |
| `DOWNLOAD_FAILED` | 附件下载未完成或文件校验失败 | 重试并检查输出目录 |
| `OPERATION_FAILED` | 写操作未达到后置条件 | 不报告成功，保留审计上下文 |

错误日志可以包含 URL、动作名、locator 逻辑名和超时信息，但不得包含邮件正文、附件内容、密码、Cookie 或完整 CDP token。

## 9. 安全与副作用控制

- 仅允许访问 `partner.outlook.cn` 及其必要的同源页面；
- Profile 目录创建为用户私有权限（Unix 下建议 `0700`），并加入 `.gitignore`；
- 启动前拒绝默认 Chrome/Edge 用户数据根目录；
- 删除、移动、归档、标记、发送仍由现有确认策略控制；
- `draft=true` 永远不能点击发送；
- `draft=false` 必须经过现有 request-id 幂等和审计流程；
- 下载文件名和 Markdown 链接必须防止路径穿越；
- 不实现 Cookie 导出、密码抓取或绕过 MFA；
- 进程收到 SIGINT/SIGTERM 时先停止新动作，再关闭页面/Context，避免半完成写操作。

## 10. 测试方案

### 10.1 单元测试

- 浏览器配置解析：默认值、环境变量覆盖、非法 Profile、CDP 互斥；
- Windows/macOS/Linux 浏览器路径发现（使用临时文件和 mock，不依赖真实安装）；
- 错误映射：Playwright Timeout、TargetClosed、下载失败分别映射到公共错误码；
- locator 选择器和 DOM 解析器：列表、文件夹、详情、附件、编辑器；
- `BrowserBackend` 契约测试：同一组行为断言同时跑 Ego Lite 和 Playwright 的 fake adapter。

### 10.2 本地集成测试

用仓库内静态 HTML fixture 模拟 Outlook 的列表、详情、文件夹树、回复编辑器和附件下载，不连接真实邮箱。使用 Playwright 的 `page.setContent` 或本地 HTTP fixture 验证：

- 列出 Inbox 和指定目录邮件；
- 按日期列出邮件；
- 打开邮件并解析正文/附件；
- 下载附件并生成相对链接；
- 保存草稿、回复全部、转发；
- 删除/移动/标记动作的后置条件和失败回滚提示。

### 10.3 手工冒烟矩阵

至少覆盖：

| 平台 | 浏览器 |
| --- | --- |
| Windows 11 | Edge、Chrome |
| macOS | Chrome、Edge |
| Linux | Chromium 或 Chrome |

真实 Outlook 验证以只读操作为主。写操作只使用测试邮件和测试草稿：`draft=true` 需要打开并手工检查，自动发送不使用真实外部联系人。

## 11. 分阶段实施

### 已完成：配置、工厂、生命周期和业务接入

- 新增 Playwright 配置、浏览器发现和 Profile 校验；
- 增加 `createBrowserBackend()`；
- 保留 Ego Lite 后端，并支持 `WEBMAIL_MODE=egolite` 强制选择；
- 增加 `status --json` 中的后端和浏览器诊断字段。

验收：配置、工厂、Ego Lite/Playwright 路由及错误提示均有自动化测试覆盖。

### 已完成：生命周期、登录和只读探针

实现内容包括：

- 实现持久化 Context、Page 复用、登录检测；
- 实现 `status`、`snapshot`、`inspect`、`inspectMessage`；
- 完成错误映射和优雅关闭。

验收：首次运行能打开可见 Outlook 页面并提示手工登录，第二次运行复用专用 Profile。

### 已完成：邮件读取能力

实现内容包括：

- 实现目录树、Inbox/目录选择、列表、日期过滤、搜索；
- 实现邮件详情、附件列表和单附件下载；
- 验证 Obsidian 导出中的附件相对链接。

验收：CLI 现有 list/read/search/export/download 命令不需要改业务参数即可切换后端。

### 已完成：写操作和安全策略接入

实现内容包括：

- 实现 compose、reply/replyAll、forward、drafts；
- 实现删除、移动、读状态、旗标、分类和归档；
- 接入现有确认、request-id、审计和手工发送等待逻辑。

验收：草稿模式不会发送；自动发送只有在成功提示/状态变化验证后才返回成功。

### 发布与后续验证

- 规划并执行 Windows/macOS/Linux 冒烟矩阵；
- 更新 README、安装诊断和故障排查；
- 当前默认已经是 `auto`，优先 Playwright；`WEBMAIL_MODE=egolite` 用于兼容和显式切换。
- 发布包由 `npm run package:release` 生成，包含 npm 包、macOS/Windows ZIP 和 SHA-256 校验文件。
- Windows、macOS、Linux 的真实浏览器冒烟仍应在对应机器执行，不能由本地单元测试替代。

## 12. 验收标准

实现完成后应满足：

1. Windows、macOS、Linux 均可在安装 Chrome/Edge/Chromium 后运行，不依赖 Ego Lite App；
2. 首次运行只需要用户在可见窗口手工登录，后续运行复用专用 Profile；
3. 目录清单、目录邮件清单、指定日期邮件、搜索、阅读、附件下载和 Obsidian 导出结果与现有后端保持一致；
4. 回复支持 `replyAll`，`draft=true` 不发送，`draft=false` 经过确认并验证发送结果；
5. 删除、移动、标记和归档不会绕过现有安全策略；
6. 浏览器缺失、Profile 锁、登录过期、UI 变化和下载失败都有可操作的错误提示；
7. 单元测试、本地 fixture 集成测试和跨平台手工冒烟全部通过；
8. Git 仓库中不出现 Profile、Cookie、附件、邮件正文或调试快照。

## 13. 风险与应对

| 风险 | 应对 |
| --- | --- |
| Outlook Web UI 改版或虚拟列表导致 DOM 变化 | 集中维护探针/locator，保留 UI_CHANGED 诊断快照，契约测试先失败再更新 |
| 企业浏览器策略禁止自动化或下载 | 在 `doctor` 中提前检测，提示管理员策略，不能通过绕过安全策略解决 |
| Profile 被多个实例占用 | 启动锁、明确错误和专用 Profile；不抢占默认 Profile |
| 系统浏览器版本差异 | 支持 executable override，CI 只固定 API/探针契约，不依赖单一浏览器版本 |
| CDP 远程调试暴露控制权 | CDP 仅本机高级选项，默认关闭并隐藏 token |
| 手工发送时 CLI 过早退出 | 草稿先验证持久化；需要手工操作时保持 Context 或明确提示用户继续在浏览器中操作 |

## 14. 计划中的命令与文档更新

实现和发布前执行：

```bash
npm install playwright-core
npm run check
npm test
npm run build
git diff --check
```

建议增加的使用示例：

```bash
# 使用系统 Edge，首次运行会打开窗口让用户登录
WEBMAIL_BACKEND=playwright WEBMAIL_BROWSER=edge webmail status

# 使用专用 Profile
WEBMAIL_BACKEND=playwright \
WEBMAIL_PROFILE_DIR="$HOME/.local/share/webmail-cli/profile" \
webmail list --dir inbox --date 2026-08-20

# 明确指定浏览器路径
WEBMAIL_BACKEND=playwright \
WEBMAIL_EXECUTABLE_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
webmail status
```

README 已说明：方案 1 不要求安装 Ego Lite，但要求系统存在受支持的浏览器；`playwright-core` 不会自动下载浏览器；首次登录由用户在可见窗口完成；专用 Profile 与日常浏览器登录态是隔离的。若用户要求直接复用已经打开的日常 Edge，则需开启远程调试并设置 `WEBMAIL_SHARE_EDGE=true`；如需全程回到 Ego Lite，使用 `--mode egolite` 或 `WEBMAIL_MODE=egolite`。

## 15. 官方资料

- [Playwright Browsers](https://playwright.dev/docs/browsers)：支持的浏览器、浏览器通道及安装方式。
- [Playwright BrowserType API](https://playwright.dev/docs/api/class-browsertype)：`launchPersistentContext` 和 `connectOverCDP` 的生命周期与限制。
- [Playwright Installation](https://playwright.dev/docs/intro)：Windows、macOS、Linux 的安装与系统要求。
- [Playwright WebView2](https://playwright.dev/docs/webview2)：Windows 上使用 Edge/WebView2 的补充说明。
