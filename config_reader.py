"""Configuration file reader supporting YAML, JSON, and INI formats."""

import json
import yaml
import tomli
from configparser import ConfigParser
from pathlib import Path
from typing import Dict, Any, Union


class ConfigReader:
    """Reads configuration data from various file formats."""
    
    def __init__(self, config_path: Union[str, Path]):
        """Initialize with config file path."""
        self.config_path = Path(config_path)
        if not self.config_path.exists():
            raise FileNotFoundError(f"Config file not found: {config_path}")
    
    def read(self) -> Dict[str, str]:
        """Read configuration and return as string dictionary."""
        file_extension = self.config_path.suffix.lower()
        
        if file_extension == '.yaml' or file_extension == '.yml':
            return self._read_yaml()
        elif file_extension == '.json':
            return self._read_json()
        elif file_extension == '.ini':
            return self._read_ini()
        elif file_extension == '.toml':
            return self._read_toml()
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")
    
    def _read_yaml(self) -> Dict[str, str]:
        """Read YAML configuration file."""
        with open(self.config_path, 'r', encoding='utf-8') as file:
            data = yaml.safe_load(file)
        return self._flatten_dict(data)
    
    def _read_json(self) -> Dict[str, str]:
        """Read JSON configuration file."""
        with open(self.config_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        return self._flatten_dict(data)
    
    def _read_ini(self) -> Dict[str, str]:
        """Read INI configuration file."""
        config = ConfigParser()
        config.read(self.config_path, encoding='utf-8')
        
        result = {}
        for section in config.sections():
            for key, value in config[section].items():
                # Use section.key format for INI files
                result[f"{section}.{key}"] = value
        
        return result
    
    def _read_toml(self) -> Dict[str, str]:
        """Read TOML configuration file."""
        with open(self.config_path, 'rb') as file:
            data = tomli.load(file)
        return self._flatten_dict(data)
    
    def _flatten_dict(self, data: Dict[str, Any], parent_key: str = '', separator: str = '.') -> Dict[str, str]:
        """Flatten nested dictionary to string values."""
        items = []
        for key, value in data.items():
            new_key = f"{parent_key}{separator}{key}" if parent_key else key
            
            if isinstance(value, dict):
                items.extend(self._flatten_dict(value, new_key, separator).items())
            else:
                items.append((new_key, str(value)))
        
        return dict(items)