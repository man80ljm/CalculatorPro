#!/usr/bin/env bash
# 由 systemd 环境文件传入配置；不输出密码、密钥、授权头。
set -euo pipefail
umask 077
project_dir="${PROJECT_DIR:-/data/projects/calculatorpro}"
backup_dir="${BACKUP_DIR:-/data/backups/calculatorpro}"
backup_python="${BACKUP_PYTHON:?请配置备份虚拟环境 Python 路径}"
key_file="${BACKUP_KEY_FILE:?请配置备份密钥文件路径}"
uploads_dir="${BACKUP_UPLOADS_DIR:-/data/uploads/calculatorpro}"
cd "$project_dir"
mkdir -p "$backup_dir"
chmod 700 "$backup_dir"
exec 9>"$backup_dir/.backup.lock"
flock -n 9 || { echo '已有备份正在执行。'; exit 1; }
snapshot_dir="$(mktemp -d "$backup_dir/.snapshot-XXXXXXXX")"
web_stopped=0
cleanup() {
  if [[ "$web_stopped" == 1 ]]; then docker compose start calculatorpro-web >/dev/null; fi
  case "$snapshot_dir" in "$backup_dir"/.snapshot-*) rm -rf -- "$snapshot_dir";; esac
}
trap cleanup EXIT
if [[ -z "$(docker compose ps --status running -q calculatorpro-web)" ]]; then
  echo '网页服务未运行，备份退出；请先检查服务状态。'
  exit 1
fi
# 夜间短暂暂停写入，保证数据库、上传资料和任务输入属于同一份快照。
web_stopped=1
docker compose stop --timeout 30 calculatorpro-web >/dev/null
docker compose exec -T postgres sh -c 'pg_dump -U "$POSTGRES_USER" -d "$POSTGRES_DB" --format=custom' >"$snapshot_dir/database.dump"
docker compose exec -T postgres pg_restore --list <"$snapshot_dir/database.dump" >/dev/null
"$backup_python" scripts/backup_bundle.py pack --dump "$snapshot_dir/database.dump" \
  --uploads "$uploads_dir" --code "$project_dir" --destination "$backup_dir" \
  --key-file "$key_file" --result-file "$snapshot_dir/result.json"
docker compose start calculatorpro-web >/dev/null
web_stopped=0
package_path="$("$backup_python" -c 'import json,sys; print(json.load(open(sys.argv[1]))["package"])' "$snapshot_dir/result.json")"
"$backup_python" scripts/backup_bundle.py upload --package "$package_path"
echo '数据库、课程资料和代码配置已完成加密备份，并通过七牛校验。'
