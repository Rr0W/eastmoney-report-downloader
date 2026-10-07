#!/usr/bin/env bash
# 东方财富 财报/研报 批量下载器 —— macOS / Linux 启动脚本
# 用法： ./run.sh --code 600519 --period 2026 --zip
#        ./run.sh --name 五粮液 --period 2026H1 --types report
#        ./run.sh --doctor
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP="$DIR/eastmoney_downloader.py"

PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done

if [ -z "$PY" ]; then
  echo "[错误] 未找到 Python。请先安装 Python 3.7 或更高版本：" >&2
  echo "       https://www.python.org/downloads/" >&2
  exit 1
fi

if ! "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 7) else 1)' 2>/dev/null; then
  echo "[错误] 需要 Python 3.7 及以上，当前版本：$("$PY" -V 2>&1)" >&2
  exit 1
fi

exec "$PY" "$APP" "$@"
