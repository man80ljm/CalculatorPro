#!/bin/sh
# 备份 Postgres，并提醒同时备份上传目录。
# 在项目根目录执行：sh scripts/backup_db.sh
set -eu
cd "$(dirname "$0")/.."
if [ -f .env ]; then
  set -a
  # shellcheck disable=SC1091
  . ./.env
  set +a
fi
mkdir -p backups
stamp="$(date +%Y%m%d-%H%M%S)"
out="backups/calculatorpro-${stamp}.sql.gz"
docker compose exec -T -e PGPASSWORD="${POSTGRES_PASSWORD:?请在 .env 中设置 POSTGRES_PASSWORD}" postgres \
  pg_dump -U "${POSTGRES_USER:?请在 .env 中设置 POSTGRES_USER}" "${POSTGRES_DB:?请在 .env 中设置 POSTGRES_DB}" \
  | gzip > "$out"
echo "数据库备份已写入 ${out}"
echo "请同时备份上传目录 /data/uploads/calculatorpro"
