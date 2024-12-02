import os
import tempfile
import uuid
from pathlib import Path

class TempManager:
    def __init__(self):
        # 在用户临时目录下创建一个唯一的子目录
        self.base_dir = Path(tempfile.gettempdir()) / f"word_formatter_{uuid.uuid4().hex}"
        self.ensure_temp_dir()

    def ensure_temp_dir(self):
        """确保临时目录存在且有正确的权限"""
        try:
            # 如果目录已存在，先删除
            if self.base_dir.exists():
                self.cleanup()
            
            # 创建新的临时目录
            self.base_dir.mkdir(parents=True, exist_ok=True)
            
            # 确保目录有正确的权限
            os.chmod(str(self.base_dir), 0o700)  # 设置目录权限为当前用户可读写执行
            
        except Exception as e:
            print(f"创建临时目录失败: {str(e)}")
            # 如果创建自定义目录失败，回退到系统临时目录
            self.base_dir = Path(tempfile.gettempdir())

    def get_temp_path(self, filename):
        """获取临时文件路径"""
        # 确保文件名是安全的
        safe_filename = Path(filename).name  # 只使用文件名部分
        temp_path = self.base_dir / safe_filename
        
        # 确保父目录存在
        temp_path.parent.mkdir(parents=True, exist_ok=True)
        
        return str(temp_path)

    def cleanup(self):
        """清理临时文件和目录"""
        try:
            if self.base_dir.exists():
                for item in self.base_dir.glob('**/*'):
                    try:
                        if item.is_file():
                            item.unlink()
                        elif item.is_dir():
                            item.rmdir()
                    except Exception as e:
                        print(f"清理临时文件失败: {str(e)}")
                
                # 删除主目录
                if self.base_dir.exists():
                    self.base_dir.rmdir()
        except Exception as e:
            print(f"清理临时目录失败: {str(e)}") 