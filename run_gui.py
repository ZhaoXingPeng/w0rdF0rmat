# -*- coding: utf-8 -*-
import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from src.gui.main_window import MainWindow

if __name__ == '__main__':
    # 添加项目根目录到Python路径
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    app = QApplication(sys.argv)
    
    # 设置应用图标
    icon_path = Path(__file__).parent / "src" / "resources" / "icons" / "icon.ico"
    if not icon_path.exists():
        # 尝试使用 app_icon.ico 作为备选
        backup_icon_path = Path(__file__).parent / "src" / "resources" / "icons" / "app_icon.ico"
        if backup_icon_path.exists():
            icon_path = backup_icon_path
        else:
            print(f"警告：找不到图标文件 {icon_path}")
            try:
                # 尝试创建默认图标
                from src.resources.icons.create_icon import create_default_icon
                icon_path = create_default_icon()
                print(f"已创建默认图标：{icon_path}")
            except Exception as e:
                print(f"创建默认图标失败：{e}")
                import traceback
                traceback.print_exc()
    
    window = MainWindow()
    
    if icon_path.exists():
        try:
            icon = QIcon(str(icon_path))
            if icon.isNull():
                print("警告：图标加载失败 - 图标对象为空")
            else:
                app.setWindowIcon(icon)
                window.setWindowIcon(icon)
                print(f"成功设置图标：{icon_path}")
        except Exception as e:
            print(f"设置图标失败：{e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"错误：图标文件不存在：{icon_path}")
    
    window.show()
    sys.exit(app.exec()) 