import logging
import json
import os

# --- LOGGING CONFIGURATION ---
logging.basicConfig(
    filename='qa_assistant.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("qa_assistant")

# --- CONFIGURATION MANAGER ---
class ConfigManager:
    CONFIG_FILE = "config.json"
    
    DEFAULT_CONFIG = {
        "ORGANIZATION_URL": "https://dev.azure.com/ProjectTestDocumentation",
        "API_TIMEOUT": 10,
        "KEY_FILE": "devops_token.txt",
        "PLACEHOLDERS": {
            "tester": "[Tester]",
            "test_id": "[TC_Number]",
            "us": "[US]",
            "env": "[Environment]",
            "profile": "[Profile]",
            "bugs": "[Bugs]",
            "result": "[Result]"
        }
    }

    @classmethod
    def load_config(cls):
        if not os.path.exists(cls.CONFIG_FILE):
            with open(cls.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(cls.DEFAULT_CONFIG, f, indent=4)
            logger.info("Created default config.json")
            return cls.DEFAULT_CONFIG
        
        try:
            with open(cls.CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load config.json: {e}")
            return cls.DEFAULT_CONFIG

# Global instances
CONFIG = ConfigManager.load_config()
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
ICON_FILE = "ey.ico"
KEY_ICON_FILE = "key.png"