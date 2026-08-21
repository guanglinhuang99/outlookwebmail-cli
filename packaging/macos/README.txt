webmail-cli macOS 安装说明
==========================

要求：
1. macOS 已安装 Node.js 24 或更高版本。
2. 已安装 Edge、Chrome 或 Chromium。
3. 安装时需要联网下载生产依赖。

安装：
1. 解压整个 ZIP，不要只取出其中一个文件。
2. 双击 install.command。
3. 如果 macOS 阻止打开，请右键 install.command，选择“打开”。
4. 打开新的终端，运行：webmail status --json
5. 在浏览器中完成 Outlook 登录后，再次运行同一命令。

命令：
- webmail：CLI
- webmail-mcp：供 MCP 客户端启动的 stdio Server

导出邮件：
- 默认 Obsidian Markdown：webmail export m_xxxxxxxxxxxxxxxxxxxx --output ./exports
- EML（附件嵌入文件）：webmail export m_xxxxxxxxxxxxxxxxxxxx --output ./exports --format eml
- 邮件 ID 来自 webmail list 或 webmail search，建议使用返回的 stableId。

详细用法见项目 README：
https://github.com/guanglinhuang99/outlookwebmail-cli
