# 2026-10-09 网页版 GitHub 同步记录

用户授权将当前已验收的网页版推送到已有仓库 https://github.com/man80ljm/CalculatorPro 。仓库当前为公开仓库，没有调整可见性。

## 分支与历史

- 最新网页版使用 web 分支：https://github.com/man80ljm/CalculatorPro/tree/web 。本地 master 跟踪 origin/web，并设置本仓库 push.default=upstream，后续普通 git push 推送到该分支。
- 远程 main 仍为桌面版 30329c4c4efdd29b4807afb87fdabdfdb61f1fb9，原网页版 cursor/web-calculator-230d 仍为 9fd82a96da4c293cf7c8d824a0d7627c8208d968，没有覆盖这些分支。
- 本地快照与原仓库没有共同根提交。使用保留当前代码树的历史合并，将原网页版历史接入本地 master，合并提交为 f317d7704d4f02d91e3a991b8bcd1609fd47f408。
- 本地原有 40 个提交全部保留；历史合并后共 75 个提交，本记录另作一次文档提交。没有强制推送或重写旧提交。
- 合并前后的代码树均为 3647c9ea7be003c9c8125051f94b3e20e646d475。没有修改网页、桌面代码、计算公式、期望值规则或表 5 数据。

## 推送与检查

Git Credential Manager 已有 man80ljm 账号，可直接完成 Git 推送；Codex 的 GitHub 连接也能读取此仓库。gh CLI 没有登录，不影响此次实际推送。

首次推送后通过 GitHub 接口确认 web 为 f317d77，main 保持原提交，已上线的业务提交 fc492c2d190cf6f03f699d85c2f512680a4275ef 也已能在 GitHub 读取。本记录提交后再次推送并核对本地与远程 SHA。

推送前扫描了相对于原远程仓库新增的 274 个文件内容对象：未发现被检查的私有配置路径或常见密钥格式。新增二进制文件只有合成教学大纲测试夹具 tests/fixtures/synthetic_syllabus.docx；原有二进制资源与远程已存在对象相同。没有提交 .env、.local、数据库、备份、账号凭据或新增真实老师数据。此项是限定规则检查，不等同于全面安全审计。

## 测试与线上状态

最近业务版本完整测试为 251 passed、1 skipped，详见 [历史版本删除上线记录](history-version-delete-deploy-20261009.md)。此次只合并 Git 历史、配置远程分支及增加同步文档，代码树未改变，没有重复运行应用测试。

线上运行代码继续为 fc492c2。本次没有修改服务器、重建镜像、重启服务或再次部署。此前线上源码一致性和资料整理结果见 [线上检查记录](production-cleanliness-20261009.md)。本地检查工具和原始结果留在 Git 忽略的 .local/remote-push-20261009/。
