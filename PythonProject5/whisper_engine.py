import whisper
import os
import logging
import tempfile

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class WhisperEngine:
    def __init__(self, model_path, lang=None):
        self.model = whisper.load_model(model_path)
        self.lang = lang
        self.logger = logger
        self.logger.info(f"Whisper引擎初始化成功 | 模型: {model_path} | 语言: {lang}")

    def transcribe(self, audio_path):
        if not isinstance(audio_path, str):
            self.logger.error(f"音频路径必须是字符串，但传入的是: {type(audio_path)}")
            return None

        temp_file = None
        try:
            # 确保音频格式兼容
            if not audio_path.lower().endswith('.wav'):
                temp_file = self.convert_to_wav(audio_path)
                if temp_file:
                    audio_path = temp_file

            # 执行转录
            result = self.model.transcribe(audio_path, language=self.lang)
            return result["text"]
        except Exception as e:
            self.logger.error(f"Whisper转录失败: {str(e)}")
            return None
        finally:
            # 清理临时文件
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    self.logger.info(f"已清理临时文件: {os.path.basename(temp_file)}")
                except Exception as e:
                    self.logger.warning(f"临时文件清理失败: {str(e)}")

    def convert_to_wav(self, audio_path):
        """将非WAV音频转换为WAV格式"""
        try:
            import librosa
            import soundfile as sf
            import tempfile

            # 创建临时文件
            temp_dir = tempfile.gettempdir()
            base_name = os.path.basename(audio_path)
            temp_path = os.path.join(temp_dir, f"whisper_temp_{os.path.splitext(base_name)[0]}.wav")

            # 读取并转换音频
            y, sr = librosa.load(audio_path, sr=16000)
            sf.write(temp_path, y, sr)

            self.logger.info(f"音频已转换为WAV格式: {temp_path}")
            return temp_path
        except Exception as e:
            self.logger.error(f"音频转换失败: {str(e)}")
            return None