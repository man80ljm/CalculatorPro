# 启用用户提供的微信收款码

用户提供 ScreenShot_2026-10-10_090316_466.png，用于替换示例并沿用已授权的上线安排。

本地 web_app/static/support/wechat.png 直接复制原始 544 × 576 像素 PNG，不裁剪、重绘或改动二维码。图片依照用户要求用于网站展示，通过独立服务器资源发布。该图片已加入 .gitignore，不提交至公开 GitHub；Git 只记录配置和界面功能。配置不包含密钥或密码。

web_app/static/support.json 改为 preview=false、wechat=/static/support/wechat.png。已有界面据此显示“微信”和“保存收款码”，移除示例及暂不能付款提示。入口、布局、完全自愿说明保持原方案。示例 SVG 保留，当前配置不使用。

提交变更为 .gitignore、web_app/static/support.json 和本说明；图片独立发布。没有修改计算、数据库、后端或桌面版。网站不检查收款到账，也不会显示未经验证的付款成功提示。

完整回归测试：251 passed、1 skipped、1 条已有的 Starlette/httpx 弃用警告，135.07 秒。第一次受沙箱限制，既有夹具无法重建专用 /tmp 测试库；以获得批准的测试命令重跑后通过。

本地独立服务使用合成账号和专用 SQLite，确认正式图片已加载（544 × 576）、保存链接指向原始 PNG 且文件名为“微信收款码.png”，示例和暂不能付款提示已移除。桌面及 390 × 844 手机视口显示完整，没有横向溢出。截图保存在 Git 忽略的 .local/coffee-support-20261010/wechat-desktop.jpg 和 wechat-mobile.jpg。

实际微信扫码和付款须使用微信完成，尚未进行实际付款。线上验收另记发布记录。
