欢迎关注公众号:秋风的安全之路
# maildata-process
# 邮箱附件数据处理

这个脚本是一个基于 POP3 协议的邮箱附件下载器，支持从邮箱中下载附件并保存到本地文件夹。它支持自动登录邮箱，获取邮件列表，并根据邮件的发件人创建文件夹保存附件
本脚本开发的想法是为了方便会计在进行一大堆附件处理的时候不方便整合的痛苦 大大提高了效率

## 功能概述

1. **自动登录**：通过提供邮箱和密码，自动登录到邮箱服务器。
2. **获取邮件列表**：获取邮箱中的邮件列表，并根据邮件 ID 逐一处理。
3. **下载附件**：自动下载邮件中的附件，并保存到本地的文件夹中。
4. **邮件保存**：将邮件的主题和发件人保存到本地文件。
5. **附件文件夹管理**：根据发件人邮箱创建文件夹保存附件，支持清理已下载的附件。
6. **文件名解码**：支持解码附件文件名，确保文件名合法，避免路径错误。

## 安装依赖

本脚本使用 Python 编写，依赖以下库：

- `poplib`：用于连接 POP3 邮件服务器。
- `email`：用于解析邮件内容。
- `os`：用于文件和目录操作。
- `shutil`：用于文件夹操作。
- `re`：用于正则表达式匹配。
- `email.header`：用于解码邮件头部信息（如主题和发件人）。

## 使用方法

1. **配置邮箱**：在脚本中配置你的邮箱和密码，或者通过 `credentials.txt` 文件加载已保存的账号信息。
2. **运行脚本**：
   ```bash
   python maildata-process.py




![图片](https://github.com/user-attachments/assets/c8fb1f03-74db-41fd-8b2f-c2b00b402263)
![1734313197630](https://github.com/user-attachments/assets/a5766684-0e84-4f5a-8a7f-1da0679a6930)

