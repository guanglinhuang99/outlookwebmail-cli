#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PACKAGE_PATH="$(find "$SCRIPT_DIR" -maxdepth 1 -name 'webmail-cli-*.tgz' -print -quit)"
INSTALL_ROOT="$HOME/.webmail-cli/runtime"
BIN_DIR="$HOME/.local/bin"

finish() {
  printf '\n按回车键关闭窗口。'
  read -r _
}
trap finish EXIT

if [[ -z "$PACKAGE_PATH" ]]; then
  printf '安装包损坏：找不到 webmail-cli-*.tgz。\n' >&2
  exit 1
fi

if ! command -v node >/dev/null 2>&1 || ! command -v npm >/dev/null 2>&1; then
  printf '需要先安装 Node.js 24 或更高版本：https://nodejs.org/\n' >&2
  exit 1
fi

NODE_MAJOR="$(node -p "Number(process.versions.node.split('.')[0])")"
if [[ "$NODE_MAJOR" -lt 24 ]]; then
  printf '当前 Node.js 版本是 %s，需要 24 或更高版本。\n' "$(node --version)" >&2
  exit 1
fi

printf '正在安装 webmail-cli，请保持网络连接……\n'
mkdir -p "$INSTALL_ROOT" "$BIN_DIR"
npm install --prefix "$INSTALL_ROOT" --omit=dev --no-audit --no-fund "$PACKAGE_PATH"

cat > "$BIN_DIR/webmail" <<'EOF'
#!/bin/sh
exec node "$HOME/.webmail-cli/runtime/node_modules/webmail-cli/dist/cli.js" "$@"
EOF
cat > "$BIN_DIR/webmail-mcp" <<'EOF'
#!/bin/sh
exec node "$HOME/.webmail-cli/runtime/node_modules/webmail-cli/dist/mcp/stdio.js" "$@"
EOF
chmod 755 "$BIN_DIR/webmail" "$BIN_DIR/webmail-mcp"

PATH_LINE='export PATH="$HOME/.local/bin:$PATH"'
touch "$HOME/.zprofile"
if ! grep -Fqx "$PATH_LINE" "$HOME/.zprofile"; then
  printf '\n# webmail-cli\n%s\n' "$PATH_LINE" >> "$HOME/.zprofile"
fi
export PATH="$BIN_DIR:$PATH"

printf '\n安装成功：webmail-cli %s\n' "$(webmail --version)"
printf '下一步运行：webmail status --json\n'
printf '如果浏览器提示登录，请完成登录后再次运行该命令。\n'
