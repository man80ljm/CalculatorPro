#!/usr/bin/env bash
# 仅在本备份留下暂停标记时恢复网页，服务超时/中断也会由ExecStopPost调用。
set -euo pipefail
project_dir="${PROJECT_DIR:-/data/projects/calculatorpro}"
backup_dir="${BACKUP_DIR:-/data/backups/calculatorpro}"
marker="$backup_dir/.web-paused-by-backup"
if [[ -f "$marker" ]]; then
  cd "$project_dir"
  docker compose start calculatorpro-web >/dev/null
  rm -f -- "$marker"
  echo '已恢复由备份暂停的网页服务。'
fi
