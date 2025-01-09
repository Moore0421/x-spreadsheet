from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import paramiko
import time
from pathlib import Path, PurePosixPath

class UploadHandler(FileSystemEventHandler):
    def __init__(self, ssh_host, ssh_port, ssh_username, ssh_password, remote_path):
        self.ssh_host = ssh_host
        self.ssh_port = ssh_port
        self.ssh_username = ssh_username
        self.ssh_password = ssh_password
        self.remote_path = remote_path
        self.ssh = None
        self.sftp = None
        self._connect()
    
    def _connect(self):
        try:
            self.ssh = paramiko.SSHClient()
            self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            # 用普通用户登录
            self.ssh.connect(
                self.ssh_host,
                port=self.ssh_port,
                username=self.ssh_username,  # 普通用户名
                password=self.ssh_password   # 普通用户密码
            )
            
            # 获取交互式 shell 并执行 sudo -i
            channel = self.ssh.invoke_shell()
            time.sleep(0.5)
            channel.send('sudo -i\n')
            time.sleep(1)  # 等待切换到 root
            
            # 现在我们已经有了 root 权限，打开 SFTP
            self.sftp = self.ssh.open_sftp()
            
        except Exception as e:
            print(f"连接错误: {e}")

    def on_modified(self, event):
        if not event.is_directory:
            try:
                # 将 Windows 路径转换为 Path 对象
                local_path = Path(event.src_path)
                local_base = Path(LOCAL_PATH)
                
                # 计算相对路径
                relative_path = local_path.relative_to(local_base)
                
                # 将相对路径转换为 Linux 格式
                remote_relative = PurePosixPath(relative_path.as_posix())
                remote_file_path = PurePosixPath(self.remote_path) / remote_relative
                
                # 确保远程目录存在
                remote_dir = str(remote_file_path.parent)
                try:
                    self.ssh.exec_command(f'sudo mkdir -p "{remote_dir}"')
                except:
                    pass
                
                # 先上传到临时位置
                temp_path = f'/tmp/{remote_relative}'
                print(f"正在上传: {local_path} -> {remote_file_path}")
                self.sftp.put(str(local_path), temp_path)
                
                # 使用 sudo 移动文件到目标位置
                self.ssh.exec_command(f'sudo mv "{temp_path}" "{remote_file_path}"')
                # 设置适当的权限
                self.ssh.exec_command(f'sudo chown root:root "{remote_file_path}"')
                self.ssh.exec_command(f'sudo chmod 644 "{remote_file_path}"')
                
            except Exception as e:
                print(f"上传错误: {e}")
    
    def _mkdir_p(self, remote_directory):
        """递归创建远程目录"""
        if remote_directory == '/':
            return
        try:
            self.ssh.exec_command(f'sudo mkdir -p "{remote_directory}"')
        except Exception as e:
            print(f"创建目录错误: {e}")

# 配置信息
LOCAL_PATH = r"D:\vue\x-spreadsheet\dist"  # Windows 本地路径
SSH_HOST = '119.91.209.28'
SSH_PORT = 22
SSH_USERNAME = 'ubuntu'     # 普通用户名
SSH_PASSWORD = 'Jiangge...0421' # 普通用户密码
REMOTE_PATH = '/opt/1panel/apps/openresty/openresty/www/sites/119.91.209.28/index'

if __name__ == "__main__":
    event_handler = UploadHandler(
        SSH_HOST,
        SSH_PORT,
        SSH_USERNAME,
        SSH_PASSWORD,
        REMOTE_PATH
    )
    observer = Observer()
    observer.schedule(event_handler, LOCAL_PATH, recursive=True)
    observer.start()

    try:
        print(f"开始监控文件夹: {LOCAL_PATH}")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        if event_handler.ssh:
            event_handler.ssh.close()
    observer.join()