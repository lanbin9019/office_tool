import os
from typing import List, Set
from office_tool.config import Config
from office_tool.utils import get_file_extension


class FileManager:
    """文件管理器类"""
    
    def __init__(self):
        self._files: List[str] = []
        self._file_set: Set[str] = set()
        self.config = Config()
    
    def add_file(self, file_path: str) -> bool:
        """
        添加文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否添加成功
        """
        if not os.path.exists(file_path):
            return False
        
        if not os.path.isfile(file_path):
            return False
        
        ext = get_file_extension(file_path)
        if ext not in self.config.SUPPORTED_EXTENSIONS:
            return False
        
        normalized_path = os.path.normpath(file_path)
        if normalized_path in self._file_set:
            return False
        
        self._files.append(normalized_path)
        self._file_set.add(normalized_path)
        return True
    
    def remove_file(self, file_path: str) -> bool:
        """
        移除文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否移除成功
        """
        normalized_path = os.path.normpath(file_path)
        if normalized_path not in self._file_set:
            return False
        
        self._files.remove(normalized_path)
        self._file_set.remove(normalized_path)
        return True
    
    def clear_files(self):
        """清空所有文件"""
        self._files.clear()
        self._file_set.clear()
    
    def get_files(self) -> List[str]:
        """获取所有文件列表"""
        return list(self._files)
    
    def get_files_by_extensions(self, extensions: Set[str]) -> List[str]:
        """
        根据扩展名获取文件列表
        
        Args:
            extensions: 扩展名集合（如 {'.xlsx', '.xls'}）
            
        Returns:
            匹配的文件列表
        """
        return [
            f for f in self._files
            if get_file_extension(f) in extensions
        ]
    
    def get_file_count(self) -> int:
        """获取文件数量"""
        return len(self._files)
    
    def has_files(self) -> bool:
        """检查是否有文件"""
        return len(self._files) > 0
    
    def has_file(self, file_path: str) -> bool:
        """
        检查文件是否已存在
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否存在
        """
        normalized_path = os.path.normpath(file_path)
        return normalized_path in self._file_set
