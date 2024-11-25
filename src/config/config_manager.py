import yaml
import os
from typing import Dict, Any, Optional

class ConfigManager:
    def __init__(self, config_path: str = "src/config/config.yaml"):
        self.config_path = config_path
        self.config = self._load_config()
        
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
                "template_path": "src/core/presets/default.yaml",
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
        template_path = os.path.join(project_path, "format_template.json")
        try:
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            
            # 更新配置
            self.config["formatting"]["user_template_path"] = template_path
            self.save_config()
            
            return template_path
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