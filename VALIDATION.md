# 验证记录

验证环境：Linux、Python 3.12、PySide6 6.12、openpyxl 3.1.5、python-pptx 1.0.2、LibreOffice Impress 25.2。

执行命令：

```bash
QT_QPA_PLATFORM=offscreen FOODBANK_TEST_SOFFICE=/workspace/foodbank-libreoffice/soffice \
  /workspace/foodbank-venv/bin/python -m unittest discover -s tests -v
```

结果：30 项测试全部通过，0 跳过。

覆盖内容：

- 桌面启动、列映射和连接设置对话框、多行粘贴、自动保存、搜索、重新加载。
- Excel 隐藏稳定 ID、复制 ID 修复、原始公式/数字和其他工作表保留、备份、行删除、外部修改冲突保护。
- Google 模拟 API 响应的精确列映射、原始字段保留、稳定 ID、重复同步去重、更新、不删除本地历史、工作表冲突、无变化时不重写文件。
- PPTX 跨 run 的姓名替换、精确自定义标记、表格单元格标记、缺少标记/模板时提示。
- PDF 缓存和资料变化后重新生成；真实 LibreOffice 导出并用 pdftotext 核对姓名。
- 桌面后台自动处理所有工作表；使用仓库原有志愿者模板和示例捐赠者模板实际生成两份 PDF，困难家庭不生成 PDF。
- 分享按钮使用模拟浏览器及文件夹打开方法验证，没有发送 WhatsApp 信息。

`examples/interface.png` 是用虚构资料运行本系统并实际生成两份 PDF 后截取的界面。

未验证项目：真实 Google 账号/表格（未提供凭证与共享权限）、Windows 上的 PowerPoint COM 转换、真实 WhatsApp 登录和发送、用户尚未提供的正式捐赠者模板。这些不以模拟测试结果代替实际验证。正式部署应在目标 Windows 机器抽查字体、长姓名和模板排版。
