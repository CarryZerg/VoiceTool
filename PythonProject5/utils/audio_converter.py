import os
import subprocess
import logging
from pathlib import Path
from typing import Optional, Tuple

# 腾讯云支持的音频格式列表 (2023最新)
TENCENT_SUPPORTED_FORMATS = {
    'wav': 'wav',
    'pcm': 'pcm',
    'mp3': 'mp3',
    'aac': 'aac',
    'm4a': 'm4a',
    'flac': 'flac'
}


class AudioConverter:
    @staticmethod
    def get_audio_info(file_path: str) -> Tuple[str, int, int]:
        """获取音频的格式、采样率和声道数"""
        try:
            # 使用ffprobe获取音频信息
            cmd = [
                'ffprobe', '-v', 'error',
                '-select_streams', 'a:0',
                '-show_entries', 'stream=codec_name,sample_rate,channels',
                '-of', 'default=noprint_wrappers=1:nokey=1',
                file_path
            ]
            output = subprocess.check_output(cmd).decode('utf-8').strip().split('\n')

            format_name = output[0].lower()
            sample_rate = int(output[1])
            channels = int(output[2])

            return format_name, sample_rate, channels
        except Exception as e:
            logging.error(f"获取音频信息失败: {str(e)}")
            raise ValueError("无法解析音频文件")

    @staticmethod
    def convert_for_tencent(
            input_path: str,
            output_dir: Optional[str] = None,
            target_format: str = 'wav',
            target_sample_rate: int = 16000,
            target_channels: int = 1
    ) -> str:
        """
        将音频转换为腾讯云兼容格式
        返回转换后的文件路径
        """
        try:
            # 获取原始音频信息
            original_format, original_sr, original_channels = AudioConverter.get_audio_info(input_path)

            # 检查是否需要转换
            need_conversion = (
                    original_format not in TENCENT_SUPPORTED_FORMATS or
                    original_sr != target_sample_rate or
                    original_channels != target_channels
            )

            if not need_conversion:
                return input_path

            # 准备输出路径
            input_file = Path(input_path)
            output_path = str(
                Path(output_dir or input_file.parent) /
                f"{input_file.stem}_converted.{target_format}"
            )

            # FFmpeg转换命令
            cmd = [
                'ffmpeg', '-y', '-i', input_path,
                '-ar', str(target_sample_rate),
                '-ac', str(target_channels),
                '-acodec', 'pcm_s16le' if target_format == 'wav' else 'copy',
                '-f', target_format,
                output_path
            ]

            # 执行转换
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            # 验证输出文件
            if not os.path.exists(output_path):
                raise RuntimeError("转换后的文件未生成")

            return output_path

        except subprocess.CalledProcessError as e:
            logging.error(f"FFmpeg转换失败: {e.stderr.decode('utf-8')}")
            raise
        except Exception as e:
            logging.error(f"音频转换异常: {str(e)}")
            raise