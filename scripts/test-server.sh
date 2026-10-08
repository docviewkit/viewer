#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
STATE_DIR="${OFFICE_VIEWER_STATE_DIR:-$ROOT_DIR/.cache/test-server}"
VIEWER_PID_FILE="$STATE_DIR/server.pid"
VIEWER_LOG_FILE="$STATE_DIR/server.log"
VIEWER_SCRIPT="$ROOT_DIR/scripts/serve.mjs"
VIEWER_HOST="${VIEWER_HOST:-${HOST:-0.0.0.0}}"
VIEWER_PORT="${VIEWER_PORT:-${PORT:-4173}}"

server_url() {
  local host="$1"
  local port="$2"
  if [[ "$host" == "0.0.0.0" || "$host" == "::" ]]; then
    host="127.0.0.1"
  fi
  printf 'http://%s:%s/' "$host" "$port"
}

read_pid() {
  local pid_file="$1"
  [[ -f "$pid_file" ]] || return 1
  local pid
  read -r pid < "$pid_file"
  [[ "$pid" =~ ^[0-9]+$ ]] || return 1
  printf '%s' "$pid"
}

is_running() {
  local pid_file="$1"
  local script="$2"
  local pid command
  pid="$(read_pid "$pid_file")" || return 1
  kill -0 "$pid" 2>/dev/null || return 1
  command="$(ps -p "$pid" -o command= 2>/dev/null || true)"
  [[ "$command" == *"$script"* ]]
}

launch_node() {
  local working_dir="$1"
  local script="$2"
  local host="$3"
  local port="$4"
  local log_file="$5"
  node - "$working_dir" "$script" "$host" "$port" "$log_file" <<'NODE'
const { openSync } = require("node:fs");
const { spawn } = require("node:child_process");

const [, , workingDir, script, host, port, logFile] = process.argv;
const log = openSync(logFile, "a");
const args = [script, "--host", host, "--port", port];
const child = spawn(process.execPath, args, {
  cwd: workingDir,
  detached: true,
  env: { ...process.env, HOST: host, PORT: port },
  stdio: ["ignore", log, log],
});
child.unref();
process.stdout.write(String(child.pid));
NODE
}

wait_for_server() {
  local label="$1"
  local pid_file="$2"
  local log_file="$3"
  local url="$4"
  local pid="$5"
  local attempt
  for attempt in {1..50}; do
    if ! kill -0 "$pid" 2>/dev/null; then
      rm -f "$pid_file"
      printf '%s启动失败，最近日志：\n' "$label" >&2
      tail -n 40 "$log_file" >&2
      return 1
    fi
    if curl --fail --silent --show-error --max-time 1 "$url" > /dev/null 2>&1; then
      printf '%s已启动：%s (PID %s)\n' "$label" "$url" "$pid"
      return
    fi
    sleep 0.2
  done

  kill "$pid" 2>/dev/null || true
  rm -f "$pid_file"
  printf '%s健康检查超时，最近日志：\n' "$label" >&2
  tail -n 40 "$log_file" >&2
  return 1
}

start_viewer() {
  local url
  url="$(server_url "$VIEWER_HOST" "$VIEWER_PORT")"
  if is_running "$VIEWER_PID_FILE" "$VIEWER_SCRIPT"; then
    printf 'OfficeViewer Inspector 已运行：%s (PID %s)\n' "$url" "$(read_pid "$VIEWER_PID_FILE")"
    return
  fi

  rm -f "$VIEWER_PID_FILE"
  : > "$VIEWER_LOG_FILE"
  printf '[%s] 编译最新代码\n' "$(date '+%Y-%m-%d %H:%M:%S')" | tee -a "$VIEWER_LOG_FILE"
  (cd "$ROOT_DIR" && npm run build) 2>&1 | tee -a "$VIEWER_LOG_FILE"

  printf '[%s] 启动 Inspector\n' "$(date '+%Y-%m-%d %H:%M:%S')" | tee -a "$VIEWER_LOG_FILE"
  local pid
  pid="$(launch_node "$ROOT_DIR" "$VIEWER_SCRIPT" "$VIEWER_HOST" "$VIEWER_PORT" "$VIEWER_LOG_FILE")"
  printf '%s\n' "$pid" > "$VIEWER_PID_FILE"
  wait_for_server "OfficeViewer Inspector " "$VIEWER_PID_FILE" "$VIEWER_LOG_FILE" "$url" "$pid"
}

start_server() {
  mkdir -p "$STATE_DIR"
  start_viewer
  printf '日志目录：%s\n' "$STATE_DIR"
}

stop_process() {
  local label="$1"
  local pid_file="$2"
  local script="$3"
  if ! is_running "$pid_file" "$script"; then
    rm -f "$pid_file"
    printf '%s未运行。\n' "$label"
    return
  fi

  local pid
  pid="$(read_pid "$pid_file")"
  kill "$pid"
  local attempt
  for attempt in {1..50}; do
    if ! kill -0 "$pid" 2>/dev/null; then
      rm -f "$pid_file"
      printf '%s已停止。\n' "$label"
      return
    fi
    sleep 0.1
  done

  kill -KILL "$pid" 2>/dev/null || true
  rm -f "$pid_file"
  printf '%s已强制停止。\n' "$label"
}

stop_server() {
  stop_process "OfficeViewer Inspector " "$VIEWER_PID_FILE" "$VIEWER_SCRIPT"
}

show_service_status() {
  local label="$1"
  local pid_file="$2"
  local script="$3"
  local url="$4"
  if is_running "$pid_file" "$script"; then
    printf '%s运行中：%s (PID %s)\n' "$label" "$url" "$(read_pid "$pid_file")"
  else
    rm -f "$pid_file"
    printf '%s未运行。\n' "$label"
  fi
}

show_status() {
  show_service_status "OfficeViewer Inspector " "$VIEWER_PID_FILE" "$VIEWER_SCRIPT" "$(server_url "$VIEWER_HOST" "$VIEWER_PORT")"
  printf '日志目录：%s\n' "$STATE_DIR"
}

show_logs() {
  mkdir -p "$STATE_DIR"
  touch "$VIEWER_LOG_FILE"
  local tail_args=(-n "${LINES:-200}")
  if [[ "${1:-}" != "--no-follow" ]]; then
    tail_args+=(-f)
  fi
  tail "${tail_args[@]}" "$VIEWER_LOG_FILE"
}

show_help() {
  cat <<'EOF'
用法：./scripts/test-server.sh <命令>

管理 OfficeViewer Inspector。官网在 docviewkit/website 仓库独立运行。

命令：
  start             编译最新代码并启动 Inspector
  restart           停止后重新编译并启动 Inspector
  logs              查看并持续跟踪 Inspector 日志（Ctrl+C 退出）
  logs --no-follow  查看 Inspector 最近日志后退出
  status            查看 Inspector 状态
  stop              停止 Inspector

默认地址：
  Inspector         监听 0.0.0.0:4173（本机访问 http://127.0.0.1:4173/）

可选环境变量：VIEWER_HOST、VIEWER_PORT、LINES、
              OFFICE_VIEWER_STATE_DIR；HOST、PORT 继续作为 Inspector 的兼容配置。
EOF
}

case "${1:-start}" in
  start)
    start_server
    ;;
  restart)
    stop_server
    start_server
    ;;
  logs|log)
    show_logs "${2:-}"
    ;;
  status)
    show_status
    ;;
  stop)
    stop_server
    ;;
  help|-h|--help)
    show_help
    ;;
  *)
    show_help >&2
    exit 2
    ;;
esac
