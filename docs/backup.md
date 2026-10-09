# 加密异地备份（2026-10-09已启用）

备份包括数据库、上传目录（课程成绩、生成资料及任务检查点）、代码和配置。代码目录中的 `.env` 只进入加密包，不会明文上传；32 字节备份解密密钥必须放在代码和上传目录之外，另存一份离线副本。

`scripts/backup_bundle.py` 使用流式 AES-256-GCM 加密，制作完成后重新解密并逐文件检查 SHA256；恢复仅写入尚不存在的新目录，不能直接覆盖运行中的站点。七牛上传使用官方 SDK，上传后核对远端大小和内容哈希，失败时返回非零退出码并保留本地包。测试全部使用合成数据和模拟网络。

部署前由管理员完成这些准备：

1. 建立七牛私有空间，配置只服务于备份的访问凭据。私有访问说明：[七牛设置访问控制](https://developer.qiniu.com/kodo/8502/set-access-control)。
2. 在服务器独立虚拟环境安装 `requirements-backup.txt`。将 `deployment/backup.env.example` 复制到 `/etc/calculatorpro-backup.env`，权限设为 600，并填写路径和凭据。
3. 在 `/data/keys/calculatorpro-backup.key` 保存 32 字节随机密钥，权限 600。密钥不放进备份包，不提交 Git，不通过网页展示。
4. 先手工执行一次 `scripts/nightly_backup.sh`，从七牛下载到隔离目录，用 `backup_bundle.py verify --package <备份包> --key-file <密钥文件> --restore-dir <新目录>` 校验并恢复。使用隔离 PostgreSQL 执行 `pg_restore`，核对账号、课程、学期、文件数，再试下载合成资料。
5. 将 `scripts/backup_bundle.py`、`nightly_backup.sh`、`backup_recover_web.sh` 安装到 `/data/ops/calculatorpro-backup/`，手工验证成功后安装 `deployment/calculatorpro-backup.service`、`.timer` 和 `calculatorpro-backup-recover.service`。服务器systemd每天北京时间02:30执行，最多随机延后60秒；开机保护仅在备份暂停标记存在时恢复网页。失败写入系统日志，本次尚未配置主动消息通知。

夜间脚本先暂停网页容器，保存并验证PostgreSQL自定义格式dump，然后打包上传资料及代码，完成加密校验后立即恢复网页，再通过HTTPS上传七牛。暂停期间站点暂时不可用；这样能让数据库和文件属于同一个快照。正常退出、失败、服务中断和开机时均有基于暂停标记的恢复处理。若今后拆出其他写入资料的后台进程，也必须同步暂停，不能只停网页。

云端对象放在 `calculatorpro-backups/v1/daily/日期/`，日期使用北京时间。日备份的上传策略设置30天后过期删除；当月第一次成功备份另传到 `monthly/月份/`，设置365天后过期删除，首日失败可以在之后成功的任务中补齐月备份。月副本完成后记录本地标记。成功上传后，本地仅清理本工具生成且满7天的加密包；失败时保留本地包。每次上传/下载都会实际查询空间的私有权限，不能仅凭配置标志放行。

备份与限量排队解决不同问题：限量减少服务被刷到崩溃的机会，备份用于机器损坏、误删除或数据损坏后的恢复。每天备份意味着最坏可能丢失最近约一天的数据；备份成功不等于已验证恢复，需要定期在隔离环境演练。

## 本次启用记录

修改文件：`scripts/backup_bundle.py`、`scripts/nightly_backup.sh`、`scripts/backup_recover_web.sh`、`requirements-backup.txt`、`tests/test_backup_bundle.py`、`deployment/backup.env.example`、`deployment/calculatorpro-backup.service`、`deployment/calculatorpro-backup-recover.service`、`README.md`、本说明。凭据、本机恢复密钥和真实验收记录均不纳入Git。

- 七牛空间：`calc-geekhaung-backup`，华南区域，私有；按用户指定拼写创建。
- 凭据：本地忽略配置 `.local/backup/qiniu.env`，服务器 `/etc/calculatorpro-backup.env`（600）。加密密钥在服务器 `/data/keys/calculatorpro-backup.key`（600），本机独立副本为 `C:\Users\huang\.calculatorpro-backup\calc-geekhaung-backup.key`，Windows仅当前用户与SYSTEM可读。密钥不上传七牛、不提交Git，应再复制一份离线保管。
- 首份成功加密包：`calculatorpro-20261009T042034Z-de7d187c.cpbackup`，约4MB。日副本和本月副本均已上传，大小与七牛内容哈希验证通过。
- 云端下载的SHA256与本地包一致；解密、166个清单文件校验通过；恢复到隔离PostgreSQL数据库后，1用户、2课程、3学期、37份资料、2大纲版本、3批资料、5条报告事件与备份前一致。37份数据库引用的资料均存在且大小正确。隔离验证库及明文恢复目录已清理。
- `calculatorpro-backup.timer` 已 `enabled/active`，下一次为2026-10-10北京时间02:30:28；`calculatorpro-backup-recover.service` 已设为开机恢复保护。网站和数据库健康，本次两次制作快照各暂停网页约5秒。
- 网站镜像仍为既有功能版本 `088d25e`；本次独立安装备份工具，没有发布本地稳定性功能批次。
- 最终本地测试：231 passed、1 skipped、1条依赖弃用警告；备份专项7项通过，Shell语法和systemd单元检查通过。

服务器完成记录在 `/data/backups/calculatorpro/latest.json`，首次云端恢复验收在 `restore-verification.json`；失败会保留本地包并进入系统日志。主动失败通知和48小时未更新提醒尚未配置，不会发送邮件或外部消息。

```bash
systemctl list-timers calculatorpro-backup.timer --no-pager
journalctl -u calculatorpro-backup.service --since yesterday --no-pager
```

下载恢复使用专用S3 HTTPS接口，参考[七牛S3服务域名](https://developer.qiniu.com/kodo/4088/s3-access-domainname)，无需公开备份空间或配置公开下载域名。加载备份环境配置后，可执行：

```bash
/data/venvs/calculatorpro-backup/bin/python /data/ops/calculatorpro-backup/backup_bundle.py download --object-key <latest.json中的备份对象键> --destination <新的加密包路径>
/data/venvs/calculatorpro-backup/bin/python /data/ops/calculatorpro-backup/backup_bundle.py verify --package <加密包路径> --key-file /data/keys/calculatorpro-backup.key --restore-dir <新的隔离目录>
```
