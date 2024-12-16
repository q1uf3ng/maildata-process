import os
import poplib
from email.parser import Parser
import re
import shutil
from email.header import decode_header
import ssl


EMAIL_SERVERS = {
    '163': 'pop.163.com',
    'qq': 'pop.qq.com',
    'gmail': 'pop.gmail.com'
}

CREDENTIALS_FILE = 'config.txt'
MAILS_FILE = 'emails_list.txt'
ATTACHMENT_DIR = './file/'

def save_credentials(email, password, server):
    with open(CREDENTIALS_FILE, 'w') as f:
        f.write(f"{email}\n{password}\n{server}")

def load_credentials():
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, 'r') as f:
            email = f.readline().strip()
            password = f.readline().strip()
            server = f.readline().strip()
            return email, password, server
    return None, None, None

def login(email, password, server):
    try:
        if server == 'pop.163.com':
            server_conn = poplib.POP3(server)
            server_conn.stls()  
        else:
            server_conn = poplib.POP3_SSL(server, 995)  

        server_conn.user(email)
        server_conn.pass_(password)
        print("登录成功！")
        return server_conn
    except Exception as e:
        print(f"登录失败: {e}")
        return None

def decode_subject(subject):
    decoded, charset = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(charset or 'utf-8')
    return decoded

def decode_filename(filename):
    decoded_parts = decode_header(filename)
    decoded_filename = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_filename += part.decode(encoding or 'utf-8')
        else:
            decoded_filename += part
    return decoded_filename

def get_sender_email(msg):
    sender = msg['from']
    match = re.search(r'<(.*?)>', sender)
    if match:
        return match.group(1)
    return None

def save_attachments(msg, sender_email):
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if "attachment" in content_disposition or content_type.startswith("application/"):
                filename = part.get_filename()
                if filename:
                    decoded_filename = decode_filename(filename)

                    folder_name = sender_email
                    folder_path = os.path.join(ATTACHMENT_DIR, folder_name)
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)

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
    if os.path.exists(ATTACHMENT_DIR):
        shutil.rmtree(ATTACHMENT_DIR)
        print("所有附件文件夹已清理")

def display_and_save_email(msg, sender_email):
    subject = decode_subject(msg['subject']) 
    sender = msg['from']

    formatted_message = f"邮件主题: {subject}\n发件人: {sender}\n{'-'*40}\n"
    print(formatted_message)

    with open(MAILS_FILE, 'a', encoding='utf-8') as f:
        f.write(formatted_message)

def fetch_emails(server_conn, email, password):
    try:
        response, mails, octets = server_conn.list()
        mail_ids = [mail.decode().split()[0] for mail in mails]

        for mail_id in mail_ids:
            print(f"正在处理邮件 ID: {mail_id}")

            response, lines, octets = server_conn.retr(int(mail_id))
            msg = Parser().parsestr(b'\r\n'.join(lines).decode())

            sender_email = get_sender_email(msg)
            if sender_email:
                save_attachments(msg, sender_email)
                display_and_save_email(msg, sender_email)
            else:
                print(f"无法获取 {msg['from']} 的邮箱")

    except Exception as e:
        print(f"获取邮件时出错: {e}")

def main():
    print("请选择邮箱服务:")
    for i, service in enumerate(EMAIL_SERVERS.keys(), start=1):
        print(f"{i}. {service}")

    choice = input("输入对应数字选择邮箱服务: ")
    try:
        server_key = list(EMAIL_SERVERS.keys())[int(choice) - 1]
        server = EMAIL_SERVERS[server_key]
    except (IndexError, ValueError):
        print("无效选择，程序退出。")
        return

    email, password, saved_server = load_credentials()

    if not email or not password or saved_server != server:
        print("请输入您的邮箱和密码：")
        email = input("邮箱: ")
        password = input("密码: ")

        remember = input("是否记住密码？(y/n): ")
        if remember.lower() == 'y':
            save_credentials(email, password, server)

    clean_choice = input("是否清理所有新建的文件夹？(y/n): ")
    if clean_choice.lower() == 'y':
        clean_attachments()

    server_conn = login(email, password, server)
    if server_conn:
        fetch_emails(server_conn, email, password)
        server_conn.quit()

if __name__ == "__main__":
    main()
