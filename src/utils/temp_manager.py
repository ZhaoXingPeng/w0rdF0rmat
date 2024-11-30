import os
import tempfile
import atexit
import shutil
from pathlib import Path

class TempManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(TempManager, cls).__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        """初始化临时文件目录"""
        # 在系统临时目录下创建一个固定的子目录
        base_temp = tempfile.gettempdir()
        self.temp_dir = os.path.join(base_temp, 'word_formatter')
        
        # 确保目录存在且为空
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir)
        
        # 注册程序退出时的清理函数
        atexit.register(self.cleanup)
    
    def get_temp_path(self, filename):
        """获取临时文件路径"""
        return os.path.join(self.temp_dir, filename)
    
    def cleanup(self):
        """清理所有临时文件"""
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                print(f"临时文件已清理: {self.temp_dir}")
        except Exception as e:
            print(f"清理临时文件失败: {str(e)}") 