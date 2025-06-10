import os
import shutil
import atexit
import tempfile
from typing import List
import logging

logger = logging.getLogger(__name__)


class TempFileCleaner:
    """临时文件清理工具（跨线程安全）"""
    _temp_dirs: List[str] = []
    _lock = threading.Lock()

    @classmethod
    def register_temp_dir(cls, dir_path: str):
        """注册需要清理的临时目录"""
        with cls._lock:
            if os.path.exists(dir_path) and dir_path not in cls._temp_dirs:
                cls._temp_dirs.append(dir_path)
                logger.debug(f"注册临时目录: {dir_path}")

    @classmethod
    def cleanup(cls):
        """清理所有注册的临时目录"""
        with cls._lock:
            for dir_path in cls._temp_dirs[:]:
                try:
                    if os.path.exists(dir_path):
                        shutil.rmtree(dir_path)
                        cls._temp_dirs.remove(dir_path)
                        logger.debug(f"已清理临时目录: {dir_path}")
                except Exception as e:
                    logger.error(f"清理临时目录失败 {dir_path}: {str(e)}")


# 注册程序退出时的清理钩子
atexit.register(TempFileCleaner.cleanup)