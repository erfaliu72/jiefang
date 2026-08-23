#!/bin/zsh

set -u

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT=49165
URL="http://127.0.0.1:${PORT}"
LOG_FILE="/private/tmp/jiefang-server.log"

cd "$PROJECT_DIR" || exit 1

if /usr/bin/curl -fsS --max-time 1 "$URL" >/dev/null; then
  open "$URL"
  exit 0
fi

nohup /usr/bin/env env PYTHONPYCACHEPREFIX=/private/tmp/jiefang-pycache \
  python3 app.py \
  >"$LOG_FILE" 2>&1 </dev/null &

SERVER_PID=$!

for _ in {1..30}; do
  sleep 0.5
  if /usr/bin/curl -fsS --max-time 1 "$URL" >/dev/null; then
    open "$URL"
    exit 0
  fi
done

echo "项目服务启动失败，日志如下："
tail -n 40 "$LOG_FILE"
