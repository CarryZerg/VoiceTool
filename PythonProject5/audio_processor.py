import os
import wave
import subprocess
import tempfile
from typing import Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class AudioProcessor:
    """音频处理工具类（线程安全版本）"""

    @staticmethod
    def is_valid_wav(path: str, sample_rate: int = 16000) -> bool:
        """检查是否为合规的WAV文件"""
        try:
            with wave.open(path, 'rb') as wf:
                return (
                        wf.getnchannels() == 1 and
                        wf.getframerate() == sample_rate and
                        wf.getsampwidth() == 2  # 16bit
                )
        except (wave.Error, EOFError) as e:
            logger.warning(f"WAV格式检查失败: {str(e)}")
            return False

    @staticmethod
    def convert_to_wav(input_path: str,
                       output_dir: Optional[str] = None,
                       sample_rate: int = 16000) -> Tuple[str, bool]:
        """
        安全转换音频格式

        Args:
            input_path: 原始音频路径
            output_dir: 指定输出目录（None则自动创建临时目录）
            sample_rate: 目标采样率

        Returns:
            (output_path, is_temp) - is_temp指示是否需要后续清理
        """
        if AudioProcessor.is_valid_wav(input_path, sample_rate):
            return input_path, False

        output_dir = output_dir or tempfile.mkdtemp(prefix="asr_temp_")
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(
            output_dir,
            f"{os.path.splitext(os.path.basename(input_path))[0]}_converted.wav"
        )

        try:
            subprocess.run([
                "ffmpeg", "-i", input_path,
                "-ar", str(sample_rate),
                "-ac", "1",
                "-acodec", "pcm_s16le",
                "-loglevel", "error",
                "-y", output_path
            ], check=True, stderr=subprocess.PIPE)
            return output_path, True
        except subprocess.CalledProcessError as e:
            error_msg = f"音频转换失败: {e.stderr.decode().strip()}"
            logger.error(error_msg)
            raise RuntimeError(error_msg) from e

    def validate_tencent_audio(audio_path: str) -> bool:
        """
        验证音频是否符合腾讯云要求:
        - 格式: WAV/PCM
        - 编码: pcm_s16le
        - 采样率: 16000
        - 声道: 单声道
        """
        try:
            import wave
            with wave.open(audio_path, 'rb') as wf:
                return (
                        wf.getnchannels() == 1 and
                        wf.getsampwidth() == 2 and
                        wf.getframerate() == 16000
                )
        except:
            return False

    @staticmethod
    def convert_for_tencent(input_path: str, output_dir: str = None) -> str:
        """
        转换为腾讯云兼容格式
        Args:
            input_path: 原始音频路径
            output_dir: 输出目录 (None则使用临时目录)
        Returns:
            转换后的合规WAV路径
        """
        import tempfile
        import subprocess

        if output_dir is None:
            output_dir = tempfile.mkdtemp(prefix="tencent_")

        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(
            output_dir,
            f"tencent_{os.path.splitext(os.path.basename(input_path))[0]}.wav"
        )

        subprocess.run([
            "ffmpeg", "-i", input_path,
            "-ar", "16000",
            "-ac", "1",
            "-acodec", "pcm_s16le",
            "-loglevel", "error",
            "-y", output_path
        ], check=True)

        return output_path
