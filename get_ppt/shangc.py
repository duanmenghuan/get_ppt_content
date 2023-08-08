import os
import paramiko

# 配置连接参数
hostname = '8.140.25.143'
port = 22  # 默认SSH端口
username = 'root'
password = 'geleiinfo!123456'
remote_directory = '/root/ppt/'

local_ppt_file = rf'F:\pptx2md\ppt\大气工作总结计划汇报PPT模板.pptx'

# 创建SSH客户端
client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

try:
    # 连接服务器
    client.connect(hostname, port, username, password)

    # 上传文件
    sftp = client.open_sftp()
    remote_path = os.path.join(remote_directory, os.path.basename(local_ppt_file))
    sftp.put(local_ppt_file, remote_path)

    print(f"文件上传成功：{remote_path}")

except Exception as e:
    print(f"上传文件时出现错误：{e}")

finally:
    # 关闭连接
    client.close()
