#!/bin/sh
# 临时数据库备份。完整加密异地备份请使用 nightly_backup.sh。
# 在项目根目录执行：sh scripts/backup_db.sh
set -eu
umask 077
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
temporary="$(mktemp backups/.database-XXXXXXXX)"
trap 'rm -f -- "$temporary"' EXIT HUP INT TERM
docker compose exec -T -e PGPASSWORD="${POSTGRES_PASSWORD:?请在 .env 中设置 POSTGRES_PASSWORD}" postgres \
  pg_dump -U "${POSTGRES_USER:?请在 .env 中设置 POSTGRES_USER}" "${POSTGRES_DB:?请在 .env 中设置 POSTGRES_DB}" \
  > "$temporary"
test -s "$temporary"
gzip -c "$temporary" > "$out"
echo "数据库备份已写入 ${out}"
echo "请同时备份上传目录 /data/uploads/calculatorpro"
