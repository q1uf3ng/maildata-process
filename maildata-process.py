import os
import poplib
from email.parser import Parser
import re
import email
import shutil
from email.header import decode_header

# 邮箱设置
POP3_SERVER = 'pop.163.com'
CREDENTIALS_FILE = 'credentials.txt'
MAILS_FILE = 'emails.txt'
ATTACHMENT_DIR = './attachments/'  # 将附件保存到当前目录下的 attachments 文件夹

def save_credentials(email, password):
    """保存邮箱和密码到本地文件"""
    with open(CREDENTIALS_FILE, 'w') as f:
        f.write(f"{email}\n{password}")

def load_credentials():
    """从本地文件加载邮箱和密码"""
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, 'r') as f:
            email = f.readline().strip()
            password = f.readline().strip()
            return email, password
    return None, None

def login(email, password):
    """使用 POP3 登录邮箱并获取邮件列表"""
    try:
        # 连接到 POP3 服务器
        server = poplib.POP3(POP3_SERVER)
        server.user(email)
        server.pass_(password)

        # 获取邮件列表
        response, mails, octets = server.list()
        server.quit()

        # 提取邮件 ID
        mail_ids = [mail.decode().split()[0] for mail in mails]
        return mail_ids
    except Exception as e:
        print(f"登录失败: {e}")
        return None

def decode_subject(subject):
    """解码邮件主题"""
    decoded, charset = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(charset or 'utf-8')
    return decoded

def decode_filename(filename):
    """解码附件文件名"""
    decoded_parts = decode_header(filename)
    decoded_filename = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_filename += part.decode(encoding or 'utf-8')
        else:
            decoded_filename += part
    return decoded_filename

def get_sender_email(msg):
    """解析发件人的邮箱"""
    sender = msg['from']
    
    # 解析发件人的邮箱
    match = re.search(r'<(.*?)>', sender)
    if match:
        return match.group(1)
    return None

def save_attachments(msg, sender_email):
    """下载邮件附件并保存到当前文件夹"""
    # 检查邮件是否包含附件
    if msg.is_multipart():
        for part in msg.walk():
            # 获取附件内容类型
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            # 检查是否为附件
            if "attachment" in content_disposition or content_type.startswith("application/"):
                filename = part.get_filename()
                if filename:
                    # 解码文件名
                    decoded_filename = decode_filename(filename)

                    # 创建文件夹以保存附件
                    folder_name = sender_email
                    folder_path = os.path.join(ATTACHMENT_DIR, folder_name)
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)

                    # 确保文件路径合法
                    file_path = os.path.join(folder_path, decoded_filename)
                    try:
                        with open(file_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        print(f"附件 {decoded_filename} 已保存到 {folder_path}")
                    except Exception as e:
                        print(f"保存附件 {decoded_filename} 时出错: {e}")
    else:
        print("没有附件")

def clean_attachments():
    """清理所有新建的文件夹"""
    if os.path.exists(ATTACHMENT_DIR):
        shutil.rmtree(ATTACHMENT_DIR)
        print("所有附件文件夹已清理")

def display_and_save_email(msg, sender_email):
    """格式化输出邮件并保存到本地文件"""
    subject = decode_subject(msg['subject'])  # 解码邮件主题
    sender = msg['from']

    formatted_message = f"邮件主题: {subject}\n发件人: {sender}\n{'-'*40}\n"
    print(formatted_message)

    # 保存邮件到本地文件
    with open(MAILS_FILE, 'a', encoding='utf-8') as f:
        f.write(formatted_message)

def main():
    """主函数"""
    # 清理功能：选择是否清理掉所有新建的文件夹
    clean_choice = input("是否清理所有新建的文件夹？(y/n): ")
    if clean_choice.lower() == 'y':
        clean_attachments()

    # 尝试从本地文件加载账号和密码
    email, password = load_credentials()

    if not email or not password:
        # 如果没有找到保存的账号和密码，提示用户输入
        print("请输入您的邮箱和密码：")
        email = input("邮箱: ")
        password = input("密码: ")

        remember = input("是否记住密码？(y/n): ")
        if remember.lower() == 'y':
            # 只有在登录成功后才保存账号和密码
            mail_ids = login(email, password)
            if mail_ids:
                save_credentials(email, password)
            else:
                print("登录失败，账号和密码不会保存。")
                return

    # 尝试登录并获取邮件 ID
    mail_ids = login(email, password)
    if mail_ids:
        # 获取邮件并处理附件
        for mail_id in mail_ids:
            print(f"正在处理邮件 ID: {mail_id}")  # 打印邮件 ID

            # 获取邮件内容
            try:
                server = poplib.POP3(POP3_SERVER)
                server.user(email)
                server.pass_(password)
                response, lines, octets = server.retr(int(mail_id))  # 使用邮件ID获取邮件
                server.quit()

                msg = Parser().parsestr(b'\r\n'.join(lines).decode())

                # 获取发件人邮箱
                sender_email = get_sender_email(msg)
                if sender_email:
                    # 为每个发件人创建文件夹并保存附件
                    save_attachments(msg, sender_email)
                    display_and_save_email(msg, sender_email)
                else:
                    print(f"无法获取 {msg['from']} 的邮箱")
            except Exception as e:
                print(f"处理邮件 ID {mail_id} 时出错: {e}")
    else:
        print("登录失败，请检查您的邮箱和密码。")

if __name__ == "__main__":
    main()
