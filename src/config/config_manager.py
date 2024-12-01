import json
from pathlib import Path
from typing import Dict, Any

class ConfigManager:
    def __init__(self):
        self.config_file = Path.home() / '.w0rdF0rmat' / 'config.json'
        self.config = self._load_config()

    def _load_config(self):
        """加载配置文件"""
        try:
            # 确保配置目录存在
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            # 如果配置文件存在，则加载它
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            
            # 如果配置文件不存在，返回默认配置
            return {}
            
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
            return {}

    def _save_config(self):
        """保存配置到文件"""
        try:
            # 确保配置目录存在
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            # 保存配置
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
                
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")

    def get(self, key, default=None):
        """获取配置值"""
        return self.config.get(key, default)

    def set(self, key, value):
        """设置配置值"""
        self.config[key] = value
        self._save_config()

    def get_format_presets(self):
        """获取格式预设"""
        return self.get('format_presets', {})

    def save_format_preset(self, name, preset):
        """保存格式预设"""
        presets = self.get_format_presets()
        presets[name] = preset
        self.set('format_presets', presets)

    def delete_format_preset(self, name):
        """删除格式预设"""
        presets = self.get_format_presets()
        if name in presets:
            del presets[name]
            self.set('format_presets', presets)
    
    def save_user_template(self, template: Dict[str, Any], project_path: str) -> str:
        """
        保存用户的格式要求为JSON文件
        Args:
            template: 格式要求
            project_path: 项目路径
        Returns:
            保存的文件路径
        """
        template_path = Path(project_path) / "format_template.json"
        try:
            # 确保目录存在
            template_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            
            # 更新配置
            self.config.setdefault("formatting", {})["user_template_path"] = str(template_path)
            self._save_config()
            
            return str(template_path)
        except Exception as e:
            print(f"保存用户模板失败: {str(e)}")
            return None
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, allow_unicode=True)
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")
    
    def is_ai_enabled(self) -> bool:
        """检查是否启用AI功能"""
        return self.config["ai_assistant"]["enabled"]
    
    def get_ai_model(self) -> str:
        """获取AI模型名称"""
        return self.config["ai_assistant"]["model"]
    
    def get_template_path(self) -> str:
        """获取当前使用的模板路径"""
        if not self.config["formatting"]["use_default_template"] and \
           self.config["formatting"]["user_template_path"]:
            return self.config["formatting"]["user_template_path"]
        
        # 使用绝对路径
        default_path = Path(__file__).parent.parent / "core" / "presets" / "default.yaml"
        return str(default_path) 