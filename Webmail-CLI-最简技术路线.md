# Webmail CLI 最简技术路线（Ego Lite 优先）

> 版本：v1.0
> 目标：把已经登录的 `https://partner.outlook.cn/mail/` Webmail 变成一个本地 CLI，供 Codex / Claude Code / OpenCode 等 AI 通过 shell 直接调用。
> 核心路线：**Node.js CLI + Ego Lite (`ego-browser`) + DOM 解析**。
> 设计原则：**单用户、本地使用、简单优先，不做浏览器插件、不做 MCP、不做 REST、不做审计、多用户、权限系统和数据库。**

---

## 1. 最终形态

```text
Codex / Claude Code / OpenCode
            │
          shell
            ▼
         webmail CLI
            │
      OutlookService
            │
      EgoLiteBackend
            │
      ego-browser nodejs
            │
       ego lite Space
            │
 partner.outlook.cn/mail/
```

最终命令：

```bash
webmail status
webmail inbox
webmail inbox --unread
webmail inbox --limit 20 --json
webmail search "风险报告"
webmail search "from:张三 风险报告" --json
webmail read 3
webmail read 3 --json
webmail attachments 3
webmail attachments 3 --json
```

v0.2 再增加：

```bash
webmail mark-read 3
webmail reply 3 "收到，我看一下" --draft
```

AI 统一通过 `--json` 使用：

```json
{
  "ok": true,
  "data": {
    "messages": [
      {
        "id": "1",
        "sender": {"name": "张三", "address": null},
        "subject": "风险报告",
        "received_at": "2026-08-19T14:32:00+08:00",
        "received_at_text": "14:32",
        "preview": "请查看附件……",
        "unread": true,
        "has_attachments": true
      }
    ]
  }
}
```

---

## 2. 范围与原则

v0.1 只支持：

- macOS
- Ego Lite
- `partner.outlook.cn`
- 收件箱列表
- Outlook 自带搜索
- 打开并读取邮件
- 附件元数据
- JSON 输出

明确不做：

- Graph
- EWS
- IMAP / SMTP
- Outlook REST v2
- OWA `service.svc`
- Bearer Token / Cookie / Storage 截获
- 浏览器插件
- MCP / REST Server
- 数据库
- 自动发送
- 删除邮件
- 后台 daemon
- 多账号
- 多用户

核心规则：

> **Snapshot 用于理解和操作页面，`js()` 用于读取真实 DOM；没有真实 DOM 勘察结果前，不得猜 Outlook selector。**

---

## 3. 技术栈

```text
Node.js 24 LTS
TypeScript
ESM
Commander
Zod
Vitest
```

项目结构：

```text
webmail-cli/
├── package.json
├── package-lock.json
├── tsconfig.json
├── README.md
├── src/
│   ├── cli.ts
│   ├── browser/
│   │   ├── backend.ts
│   │   ├── ego-runner.ts
│   │   └── ego-lite.ts
│   ├── outlook/
│   │   ├── service.ts
│   │   ├── state.ts
│   │   ├── dom-probes.ts
│   │   ├── inbox-parser.ts
│   │   ├── message-parser.ts
│   │   └── selectors.ts
│   ├── session/
│   │   └── session-store.ts
│   ├── types/
│   │   ├── mail.ts
│   │   └── result.ts
│   └── util/
│       ├── errors.ts
│       └── text.ts
└── tests/
    ├── inbox-parser.test.ts
    ├── message-parser.test.ts
    └── fixtures/
```

---

## 4. package.json

```json
{
  "name": "webmail-cli",
  "version": "0.1.0",
  "type": "module",
  "private": true,
  "bin": {
    "webmail": "./dist/cli.js"
  },
  "scripts": {
    "build": "tsc -p tsconfig.json",
    "dev": "tsx src/cli.ts",
    "test": "vitest run",
    "check": "tsc -p tsconfig.json --noEmit"
  },
  "dependencies": {
    "commander": "^14.0.0",
    "zod": "^4.0.0"
  },
  "devDependencies": {
    "@types/node": "^24.0.0",
    "tsx": "^4.20.0",
    "typescript": "^5.9.0",
    "vitest": "^3.2.0"
  }
}
```

正式编码时以 npm 实际兼容版本为准，并提交 `package-lock.json`。

`tsconfig.json`：

```json
{
  "compilerOptions": {
    "target": "ES2023",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "outDir": "dist",
    "rootDir": "src",
    "sourceMap": true
  },
  "include": ["src/**/*.ts"]
}
```

---

## 5. BrowserBackend

`src/browser/backend.ts`

```ts
export interface BrowserBackend {
  status(): Promise<{
    connected: boolean;
    url: string | null;
    title: string | null;
  }>;

  snapshot(): Promise<string>;
  eval<T>(script: string): Promise<T>;
  click(refOrLocator: string): Promise<void>;
  fill(refOrLocator: string, text: string): Promise<void>;
  press(key: string): Promise<void>;
  scrollBy(x: number, y: number): Promise<void>;
  wait(ms: number): Promise<void>;
}
```

v0.1 只实现：

```text
EgoLiteBackend
```

未来 Windows 只需要增加：

```text
PlaywrightChromeBackend
```

CLI 和 Outlook Parser 不改。

---

## 6. EgoRunner

所有 `ego-browser` 子进程调用必须集中到：

```text
src/browser/ego-runner.ts
```

使用：

```ts
import { spawn } from 'node:child_process';
```

启动：

```ts
spawn('ego-browser', ['nodejs'], {
  stdio: ['pipe', 'pipe', 'pipe'],
});
```

禁止业务代码自己拼 shell 命令，也不写临时 `.js` 文件。

接口：

```ts
export interface EgoRunResult<T> {
  stdout: string;
  stderr: string;
  value: T;
}

export class EgoRunner {
  async run<T>(body: string, timeoutMs = 30000): Promise<EgoRunResult<T>> {
    // 1. spawn('ego-browser', ['nodejs'])
    // 2. body 写入 stdin
    // 3. 收集 stdout/stderr
    // 4. 超时 kill
    // 5. 解析最后一个 __webmail_result__
    // 6. 非 0 退出码转成 EGO_BROWSER_ERROR
    throw new Error('TODO');
  }
}
```

每段脚本最终只通过：

```js
cliLog(JSON.stringify({
  __webmail_result__: true,
  result: ...
}));
```

返回结果。

Node 端扫描 stdout，从后向前找到最后一个能解析成：

```json
{"__webmail_result__":true,"result":...}
```

的 JSON 行。

默认超时：

```text
status 10 秒
普通读取 30 秒
search/read 45 秒
```

---

## 7. EgoLiteBackend

固定常量：

```ts
const TASK_SPACE = 'webmail-cli';
const OUTLOOK_URL = 'https://partner.outlook.cn/mail/';
```

每次脚本都先：

```js
await useOrCreateTaskSpace('webmail-cli');

await openOrReuseTab(
  'https://partner.outlook.cn/mail/',
  {
    wait: true,
    timeout: 30
  }
);
```

Ego Lite 的 heredoc Node.js 进程不会保留内存状态，因此每次都重新 `useOrCreateTaskSpace()`；Task Space 和里面的 Tab 可以继续复用。

`status()` 示例：

```js
await useOrCreateTaskSpace('webmail-cli');

await openOrReuseTab(
  'https://partner.outlook.cn/mail/',
  {
    wait: true,
    timeout: 30
  }
);

const info = await pageInfo();

cliLog(JSON.stringify({
  __webmail_result__: true,
  result: info
}));
```

如果最终顶层 URL 不属于允许的 Outlook 页面，或明显进入登录页面，返回：

```text
AUTH_REQUIRED
```

---

## 8. Mail 数据模型

`src/types/mail.ts`

```ts
export interface MailAddress {
  name: string | null;
  address: string | null;
}

export interface MailSummary {
  id: string;
  sender: MailAddress;
  subject: string;
  receivedAt: string | null;
  receivedAtText: string | null;
  preview: string | null;
  unread: boolean | null;
  hasAttachments: boolean | null;
}

export interface AttachmentSummary {
  id: string;
  filename: string;
  sizeText: string | null;
}

export interface MailMessage {
  id: string;
  subject: string;
  from: MailAddress;
  to: MailAddress[];
  cc: MailAddress[];
  receivedAt: string | null;
  receivedAtText: string | null;
  bodyText: string;
  attachments: AttachmentSummary[];
}
```

统一 JSON Envelope：

```ts
export type CliResult<T> =
  | { ok: true; data: T }
  | {
      ok: false;
      error: {
        code: string;
        message: string;
      };
    };
```

错误码：

```text
EGO_BROWSER_NOT_FOUND
EGO_BROWSER_ERROR
AUTH_REQUIRED
OUTLOOK_NOT_READY
UI_CHANGED
MESSAGE_NOT_FOUND
AMBIGUOUS_MESSAGE
INVALID_ARGUMENT
TIMEOUT
```

---

## 9. Session Store

不使用数据库。

文件：

```text
~/.webmail-cli/session.json
```

内容：

```json
{
  "updatedAt": "2026-08-19T19:00:00+08:00",
  "source": "inbox",
  "messages": {
    "1": {
      "subject": "风险报告",
      "senderName": "张三",
      "senderAddress": null,
      "receivedAtText": "14:32",
      "preview": "请查看附件……"
    }
  }
}
```

每次：

```bash
webmail inbox
webmail search ...
```

重新生成短 ID：

```text
1
2
3
...
```

`webmail read 3` 根据 session 保存的：

```text
subject
sender
receivedAtText
preview
```

重新定位邮件。

不依赖 Outlook 是否暴露稳定内部 message ID。

如果候选不唯一：

```text
AMBIGUOUS_MESSAGE
```

不能自动打开第一封。

---

## 10. 页面解析总策略

正式逻辑：

```text
snapshotText()
    ↓
判断当前页面状态 / 找交互控件
    ↓
js()
    ↓
直接从真实 DOM 提取结构化字段
```

不要只依赖 Snapshot 文本，也不要只依赖 DOM class。

选择器优先级：

```text
1. role + aria-label
2. 稳定 data-* 属性
3. 语义 DOM 结构
4. title
5. class 最后才用
```

禁止：

```text
随机 hash class
绝对 DOM path
nth-child
固定 Snapshot ref
```

Snapshot ref 页面变化后立即失效，不能保存到 session。

---

## 11. 第一阶段先实现 `webmail inspect`

这是整个项目最关键的开发步骤。

在没有真实 `partner.outlook.cn` DOM 输出前，编码 AI **不得实现最终 InboxParserV1**。

命令：

```bash
webmail inspect
webmail inspect --json
```

至少返回：

```text
pageInfo
snapshotText
DOM inventory
list candidates
message row candidates
scroll container candidates
iframes
```

---

## 12. DOM Inventory Probe

`src/outlook/dom-probes.ts`

```ts
export const PAGE_DOM_INVENTORY = String.raw`
(() => {
  function visible(el) {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);

    return (
      r.width > 0 &&
      r.height > 0 &&
      s.display !== 'none' &&
      s.visibility !== 'hidden'
    );
  }

  function selectedAttrs(el) {
    return Object.fromEntries(
      Array.from(el.attributes)
        .filter(a =>
          a.name.startsWith('aria-') ||
          a.name.startsWith('data-') ||
          ['role', 'title', 'datetime', 'contenteditable'].includes(a.name)
        )
        .map(a => [a.name, a.value])
    );
  }

  return Array.from(document.querySelectorAll('*'))
    .filter(visible)
    .map((el, index) => {
      const r = el.getBoundingClientRect();
      const text = (el.innerText || '').trim();

      return {
        index,
        tag: el.tagName,
        role: el.getAttribute('role'),
        ariaLabel: el.getAttribute('aria-label'),
        title: el.getAttribute('title'),
        datetime: el.getAttribute('datetime'),
        contentEditable: el.getAttribute('contenteditable'),
        text: text.slice(0, 500),
        textLength: text.length,
        attrs: selectedAttrs(el),
        childCount: el.children.length,
        scrollHeight: el.scrollHeight,
        clientHeight: el.clientHeight,
        rect: {
          x: Math.round(r.x),
          y: Math.round(r.y),
          width: Math.round(r.width),
          height: Math.round(r.height)
        }
      };
    })
    .filter(x =>
      x.role ||
      x.ariaLabel ||
      x.title ||
      x.datetime ||
      Object.keys(x.attrs).length > 0 ||
      x.textLength > 0
    )
    .slice(0, 5000);
})()
`;
```

仅供：

```text
inspect
调试
UI 改版排查
```

正式 `inbox/read` 不能每次全页扫 5000 个节点。

---

## 13. 邮件列表候选 Probe

```ts
export const LIST_CANDIDATES = String.raw`
(() => {
  const selector = [
    '[role="list"]',
    '[role="listbox"]',
    '[role="grid"]',
    '[role="treegrid"]',
    '[role="feed"]'
  ].join(',');

  return Array.from(document.querySelectorAll(selector))
    .map((el, index) => {
      const r = el.getBoundingClientRect();
      const text = (el.innerText || '').trim();

      return {
        index,
        tag: el.tagName,
        role: el.getAttribute('role'),
        ariaLabel: el.getAttribute('aria-label'),
        text: text.slice(0, 3000),
        childCount: el.children.length,
        scrollHeight: el.scrollHeight,
        clientHeight: el.clientHeight,
        rect: {
          x: Math.round(r.x),
          y: Math.round(r.y),
          width: Math.round(r.width),
          height: Math.round(r.height)
        }
      };
    });
})()
`;
```

用途：

```text
找出哪个区域才是真正的 Outlook 邮件列表
```

---

## 14. 邮件行候选 Probe

```ts
export const MESSAGE_ROW_CANDIDATES = String.raw`
(() => {
  const candidates = Array.from(document.querySelectorAll(
    '[role="option"], [role="listitem"], [role="row"]'
  ));

  return candidates
    .map((row, index) => ({
      index,
      tag: row.tagName,
      role: row.getAttribute('role'),
      ariaLabel: row.getAttribute('aria-label'),
      text: (row.innerText || '').trim().slice(0, 2000),
      attrs: Object.fromEntries(
        Array.from(row.attributes)
          .filter(a =>
            a.name.startsWith('aria-') ||
            a.name.startsWith('data-') ||
            ['role', 'title'].includes(a.name)
          )
          .map(a => [a.name, a.value])
      ),
      children: Array.from(row.querySelectorAll('*'))
        .slice(0, 100)
        .map((el, childIndex) => ({
          childIndex,
          tag: el.tagName,
          role: el.getAttribute('role'),
          ariaLabel: el.getAttribute('aria-label'),
          title: el.getAttribute('title'),
          datetime: el.getAttribute('datetime'),
          text: (el.innerText || '').trim().slice(0, 500),
          attrs: Object.fromEntries(
            Array.from(el.attributes)
              .filter(a =>
                a.name.startsWith('aria-') ||
                a.name.startsWith('data-')
              )
              .map(a => [a.name, a.value])
          )
        }))
    }))
    .filter(x => x.text.length > 0)
    .slice(0, 100);
})()
`;
```

真实输出必须用于确认：

```text
sender
subject
time
preview
unread
attachment
```

各字段实际在 DOM 的哪里。

---

## 15. InboxParserV1

在用户把真实：

```bash
webmail inspect --json > outlook-inbox-inspect.json
```

输出交回后，再写真正 Parser。

接口：

```ts
export interface RawMessageRow {
  stableHint: string | null;
  senderName: string | null;
  senderAddress: string | null;
  subject: string;
  receivedAtText: string | null;
  preview: string | null;
  unread: boolean | null;
  hasAttachments: boolean | null;
}

export interface InboxParser {
  extract(): Promise<RawMessageRow[]>;
}
```

要求：

```text
只在已确认邮件列表 root 内取行
每个字段只从本行子树取
字段缺失返回 null
不从相邻邮件补字段
列表非空但解析为 0 行时返回 UI_CHANGED
```

---

## 16. 页面状态

`src/outlook/state.ts`

```ts
export type OutlookState =
  | 'AUTH_REQUIRED'
  | 'INBOX'
  | 'SEARCH_RESULTS'
  | 'MESSAGE_OPEN'
  | 'COMPOSE_OPEN'
  | 'UNKNOWN';
```

判断依据：

```text
pageInfo().url
+
snapshotText()
+
少量 DOM 信号
```

写操作后必须重新 Snapshot，不复用旧 ref。

---

## 17. `webmail status`

流程：

```text
useOrCreateTaskSpace
→ openOrReuseTab Outlook
→ pageInfo
→ snapshot
→ detect state
→ 输出
```

JSON：

```json
{
  "ok": true,
  "data": {
    "backend": "ego-lite",
    "url": "https://partner.outlook.cn/mail/...",
    "state": "INBOX"
  }
}
```

---

## 18. `webmail inbox`

流程：

```text
ensure Outlook
→ detect state
→ 若不是 INBOX，点击 Inbox
→ 新 Snapshot
→ InboxParserV1
→ unread/limit 本地过滤
→ 生成 1..N 短 ID
→ 保存 session.json
→ 输出
```

参数：

```bash
webmail inbox
webmail inbox --limit 20
webmail inbox --unread
webmail inbox --json
```

默认：

```text
limit = 20
```

---

## 19. 虚拟滚动

v0.1 不同步整个邮箱。

如果当前 DOM 行数不足 `limit`：

```text
scrollBy(0, 700)
→ wait
→ 重新解析 DOM
→ fingerprint 去重
```

最多：

```text
5 次滚动
100 封邮件
```

fingerprint：

```text
normalize(sender)
+
normalize(subject)
+
receivedAtText
+
normalize(preview.slice(0,80))
```

达到 limit 立即停止。

---

## 20. `webmail search`

直接操作 Outlook 自带搜索框。

不自己实现搜索 DSL。

流程：

```text
snapshot
→ 找唯一搜索 textbox
→ fillInput
→ pressKey Enter
→ wait
→ 新 snapshot
→ detect SEARCH_RESULTS
→ InboxParserV1
→ 保存 session
→ 返回
```

例如：

```bash
webmail search "风险报告"
webmail search "from:张三 风险报告"
```

输入原样交给 Outlook。

搜索框优先匹配 accessible name：

```text
搜索
Search
搜索邮件和人员
```

真实名称以 inspect 输出为准。

---

## 21. `webmail read <id>`

从 session 取出邮件指纹。

重新定位策略：

```text
A. 当前列表中完整 fingerprint 匹配
↓失败
B. Outlook 搜索 subject
↓
C. 搜索候选内比较 sender + time + preview
```

结果：

```text
唯一 → click
0 → MESSAGE_NOT_FOUND
>1 → AMBIGUOUS_MESSAGE
```

点击后：

```text
wait
→ 新 Snapshot
→ detect MESSAGE_OPEN
→ MessageParserV1
```

---

## 22. `webmail inspect-message`

在开发 MessageParser 前先实现：

```bash
webmail inspect-message --json
```

用户先手工打开一封无敏感内容的测试邮件。

Probe：

```ts
export const MESSAGE_BODY_CANDIDATES = String.raw`
(() => {
  const selector = [
    '[role="document"]',
    '[role="main"]',
    'article',
    'iframe'
  ].join(',');

  return Array.from(document.querySelectorAll(selector))
    .map((el, index) => {
      const r = el.getBoundingClientRect();
      const text = (el.innerText || '').trim();

      return {
        index,
        tag: el.tagName,
        role: el.getAttribute('role'),
        ariaLabel: el.getAttribute('aria-label'),
        textLength: text.length,
        text: text.slice(0, 5000),
        rect: {
          x: Math.round(r.x),
          y: Math.round(r.y),
          width: Math.round(r.width),
          height: Math.round(r.height)
        }
      };
    })
    .filter(x => x.textLength > 50 || x.tag === 'IFRAME');
})()
`;
```

没有这个真实输出前不得宣称 MessageParser 已完成。

---

## 23. MessageParserV1

接口：

```ts
export interface RawMessage {
  subject: string;
  fromName: string | null;
  fromAddress: string | null;
  to: Array<{name: string | null; address: string | null}>;
  cc: Array<{name: string | null; address: string | null}>;
  receivedAtText: string | null;
  bodyText: string;
  attachments: Array<{
    filename: string;
    sizeText: string | null;
  }>;
}
```

要求：

```text
只取真正正文区域
不混入左侧邮件列表
不混入按钮/导航文字
保留段落换行
正文最大 100 KB
```

如果正文位于 iframe：

```text
先观察 Ego Lite snapshot 能否完整看到
js() 直接访问不了时再考虑 cdp()
```

v0.1 只需覆盖当前真实邮箱常见邮件。

---

## 24. `webmail attachments`

流程：

```text
read/open
→ MessageParserV1
→ 返回附件元数据
```

v0.1：

```text
只列附件，不下载
```

---

## 25. v0.2：mark-read / reply draft

`mark-read`：

```text
重新定位邮件
→ 最新 Snapshot 找 Mark as read
→ click
→ 新 Snapshot / DOM 验证结果
```

`reply --draft`：

```text
open message
→ 最新 Snapshot 找 Reply
→ click
→ 新 Snapshot
→ 找 compose body
→ fillInput
→ 不点击 Send
```

v0.2 仍不做自动发送。

---

## 26. CLI

`src/cli.ts`

```ts
import { Command } from 'commander';

const program = new Command();

program
  .name('webmail')
  .description('Outlook Web CLI through Ego Lite');

program.command('status');

program
  .command('inbox')
  .option('-n, --limit <number>')
  .option('--unread')
  .option('--json');

program
  .command('search')
  .argument('<query>')
  .option('-n, --limit <number>')
  .option('--json');

program
  .command('read')
  .argument('<id>')
  .option('--json');

program
  .command('attachments')
  .argument('<id>')
  .option('--json');

program.command('inspect').option('--json');
program.command('inspect-message').option('--json');

await program.parseAsync();
```

`--json` 模式：

```text
stdout 只能输出 JSON
日志全部 stderr
```

这样：

```bash
webmail inbox --json | jq
```

必须正常。

---

## 27. 登录处理

假设用户已经在 Ego Lite 登录：

```text
https://partner.outlook.cn/mail/
```

程序不填写账号、密码、MFA。

登录过期时：

```json
{
  "ok": false,
  "error": {
    "code": "AUTH_REQUIRED",
    "message": "请在 Ego Lite 中重新登录 Outlook，然后再次执行命令。"
  }
}
```

---

## 28. 开发顺序

### Task 1：骨架

实现：

```text
package.json
tsconfig
CLI
types
errors
```

验收：

```bash
npm run check
npm run build
```

### Task 2：EgoRunner

完成：

```text
spawn
stdin
stdout JSON parse
stderr
timeout
error mapping
```

### Task 3：EgoLiteBackend

完成：

```text
status
snapshot
eval
click
fill
press
scrollBy
wait
```

### Task 4：inspect

完成：

```bash
webmail inspect --json
```

至少输出：

```text
page
snapshot
domInventory
listCandidates
messageRowCandidates
scrollCandidates
iframes
```

**到这里暂停。**

用户执行：

```bash
webmail inspect --json > outlook-inbox-inspect.json
```

把文件交给编码 AI。

### Task 5：InboxParserV1

只根据真实 inspect 输出实现。

### Task 6：inbox + session

完成：

```bash
webmail inbox
webmail inbox --json
```

### Task 7：search

操作 Outlook 搜索 UI。

### Task 8：inspect-message

用户打开测试邮件，执行：

```bash
webmail inspect-message --json > outlook-message-inspect.json
```

### Task 9：MessageParserV1

只根据真实 inspect-message 输出实现。

### Task 10：read + attachments

完成 v0.1。

### Task 11：可选 v0.2

```text
mark-read
reply --draft
```

---

## 29. 测试

只做必要测试。

单元测试：

```text
EgoRunner stdout parser
timeout
session store
InboxParser fixture
MessageParser fixture
fingerprint
```

真实烟雾测试：

```bash
webmail status
webmail inbox --limit 5
webmail search "测试"
webmail read 1
webmail attachments 1
```

每个 Task 完成后运行：

```bash
npm run check
npm test
npm run build
```

---

## 30. v0.1 验收

```text
[ ] webmail status 可识别已登录 Outlook
[ ] webmail inspect 能输出真实页面勘察 JSON
[ ] inbox 能返回真实邮件列表
[ ] inbox/search 生成稳定短 ID
[ ] search 使用 Outlook 自带搜索
[ ] read 能重新定位并读取唯一邮件
[ ] 同主题歧义时不误开
[ ] 正文是可读纯文本
[ ] attachments 返回附件元数据
[ ] --json stdout 是合法单一 JSON
[ ] 不使用 Graph / EWS / IMAP / SMTP / REST v2
[ ] 不截获 token/cookie
[ ] npm run check && npm test && npm run build 通过
```

---

## 31. 直接交给编码 AI 的 Prompt

```text
请严格按本 Markdown 实现 webmail-cli。

目标：
在 macOS 上使用 Ego Lite 的 ego-browser，控制已经登录的
https://partner.outlook.cn/mail/
通过 DOM + UI 自动化实现一个供 AI shell 调用的 webmail CLI。

硬性要求：

1. Node.js + TypeScript ESM。
2. 第一版只支持 Ego Lite。
3. 不做浏览器插件、MCP、REST、数据库、后台服务。
4. 所有 ego-browser 调用集中在 EgoRunner / EgoLiteBackend。
5. 固定复用 Task Space：webmail-cli。
6. 使用 snapshotText() 判断页面和寻找交互控件。
7. 使用 js() 从真实 DOM 提取字段。
8. 不使用 Graph、EWS、IMAP、SMTP、Outlook REST v2、service.svc。
9. 不捕获 Bearer Token、Cookie、Storage。
10. 不根据网上的 Outlook DOM 或截图猜 Selector。
11. 第一轮必须只做到 Task 4：webmail inspect 可以真实运行。
12. 在拿到真实 inspect 输出之前：
    - 完成项目骨架
    - EgoRunner
    - EgoLiteBackend
    - DOM probes
    - session store
    - Parser interface
    - 测试框架
    但不得伪造 InboxParserV1。
13. 第一轮结束时明确让我执行：
    webmail inspect --json > outlook-inbox-inspect.json
14. 我把真实 inspect JSON 给你后，再实现 InboxParserV1。
15. MessageParser 同样先实现 inspect-message，再根据真实输出编码。
16. 所有 JSON 命令统一返回：
    {"ok":true,"data":...}
    或
    {"ok":false,"error":{"code":"...","message":"..."}}
17. --json 模式 stdout 只能输出 JSON，日志写 stderr。
18. 页面无法识别返回 UI_CHANGED。
19. 无法唯一定位邮件返回 AMBIGUOUS_MESSAGE。
20. 绝不能“找不到就点第一条”。
21. 每完成一个 Task 都实际运行：
    npm run check
    npm test
    npm run build
22. 不要只输出代码片段，要实际创建完整工程文件。
23. 每轮报告：
    - 修改文件
    - 执行命令
    - 测试结果
    - 阻塞项
    - 下一步
24. 第一轮完成 Task 1→4 后停止，等待我的真实 Outlook DOM inspect 输出。
```

---

## 32. 第一轮开发后的实际操作

编码 AI 完成 Task 1—4：

```bash
cd webmail-cli
npm install
npm run check
npm test
npm run build
npm link
```

然后：

```bash
webmail --help
webmail status
```

确认 Ego Lite 已经登录 Outlook。

执行：

```bash
webmail inspect --json > outlook-inbox-inspect.json
```

把 `outlook-inbox-inspect.json` 给编码 AI。

第二轮：

```text
真实 DOM
→ InboxParserV1
→ inbox
→ search
```

然后手工打开一封测试邮件：

```bash
webmail inspect-message --json > outlook-message-inspect.json
```

第三轮：

```text
真实正文 DOM
→ MessageParserV1
→ read
→ attachments
```

---

## 33. 最终路线

```text
第一轮：
CLI + Ego Backend + inspect

第二轮：
真实 Inbox DOM → inbox/search

第三轮：
真实 Message DOM → read/attachments
```

完成后的 AI 使用方式只有：

```bash
webmail inbox --json
webmail search "xxx" --json
webmail read 3 --json
webmail attachments 3 --json
```

**这是 v0.1 的全部目标。**

---

## 34. Ego Lite API 约束备注

实现时以当前 `ego-browser` 官方 Skill 为准：

- `ego-browser nodejs <<'EOF' ... EOF` 是标准调用方式。
- helpers 在 heredoc Node.js 运行时预加载。
- `snapshotText()` 适合语义树观察。
- `js()` 在浏览器页面上下文运行，可访问 `document/window`。
- `useOrCreateTaskSpace(name)` 用于跨多轮复用同一个 Space。
- `openOrReuseTab(url, {wait:true})` 用于打开或复用页面。
- `cliLog()` 是 heredoc 中向终端输出结果的方式。
- 页面发生点击、搜索、滚动、导航后，应重新 Snapshot，不复用旧 ref。
- `js()` 复杂表达式使用显式 `(() => { ... })()`；需要正则时优先 `String.raw`。

官方参考：

- https://github.com/citrolabs/ego-lite
- https://github.com/citrolabs/ego-lite/blob/main/skills/ego-browser/SKILL.md
