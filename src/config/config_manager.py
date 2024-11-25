import yaml
import os
from typing import Dict, Any, Optional
from pathlib import Path

class ConfigManager:
    def __init__(self, config_path: str = None):
        """
        初始化配置管理器
        Args:
            config_path: 配置文件路径，如果为None则使用默认路径
        """
        if config_path is None:
            # 使用相对于当前文件的路径
            self.config_path = Path(__file__).parent / "config.yaml"
        else:
            self.config_path = Path(config_path)
            
        # 确保配置文件存在
        if not self.config_path.exists():
            self._create_default_config()
            
        self.config = self._load_config()
    
    def _create_default_config(self):
        """创建默认配置文件"""
        default_config = {
            "ai_assistant": {
                "enabled": False,
                "model": "gpt-3.5-turbo"
            },
            "formatting": {
                "use_default_template": True,
                "template_path": str(Path(__file__).parent.parent / "core" / "presets" / "default.yaml"),
                "user_template_path": None
            }
        }
        
        try:
            # 确保目录存在
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 写入默认配置
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(default_config, f, allow_unicode=True)
            print(f"已创建默认配置文件: {self.config_path}")
        except Exception as e:
            print(f"创建默认配置文件失败: {str(e)}")
            
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            "ai_assistant": {
                "enabled": False,
                "model": "gpt-3.5-turbo"
            },
            "formatting": {
                "use_default_template": True,
                "template_path": str(Path(__file__).parent.parent / "core" / "presets" / "default.yaml"),
                "user_template_path": None
            }
        }
    
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
            self.config["formatting"]["user_template_path"] = str(template_path)
            self.save_config()
            
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
        return self.config["formatting"]["template_path"] 