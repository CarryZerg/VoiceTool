import wave
import os
import json
import logging
import subprocess
import tempfile
import numpy as np
import soundfile as sf
import time
from typing import Optional, Union, Dict, List
from vosk import Model, KaldiRecognizer
import torch
import whisper
import base64
import logging
import os
import re

try:
    from tencentcloud.asr.v20190614 import models as tencent_models
    from tencentcloud.common import credential
    from tencentcloud.common.profile.client_profile import ClientProfile
    from tencentcloud.common.profile.http_profile import HttpProfile
    from tencentcloud.asr.v20190614 import asr_client

    TENCENT_SDK_AVAILABLE = True
except ImportError:
    TENCENT_SDK_AVAILABLE = False



class STTEngine:
    _instance = None
    _initialized = False

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self,
                 model_config: Union[str, Dict, None] = None,
                 lang: str = 'zh',
                 engine_type: str = 'vosk',
                 config: Optional[Dict] = None):

        if self.__class__._initialized:
            return

        # 初始化日志
        self.logger = logging.getLogger(__name__)
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)

        # 参数验证
        if engine_type.lower() in ('tencent', 'microsoft'):
            if not config and not isinstance(model_config, dict):
                raise ValueError("Cloud engines require config dict or file path")
        else:
            if not model_config:
                raise ValueError("Traditional engines require model path")

        self.engine_type = engine_type.lower()
        self.lang = lang
        self.config = config or {}

        # 处理模型配置
        if isinstance(model_config, dict):
            self.model_config = model_config
        elif isinstance(model_config, str):
            if engine_type.lower() in ('tencent', 'microsoft'):
                # 云服务配置文件
                try:
                    with open(model_config, 'r', encoding='utf-8') as f:
                        self.model_config = json.load(f)
                except Exception as e:
                    raise ValueError(f"Failed to load config file: {str(e)}")
            else:
                # 本地模型路径
                self.model_config = model_config
        else:
            raise TypeError("model_config must be str or dict")

        self.__class__._initialized = True
        self._initialize_engine()

    @classmethod
    def reset_engine(cls):
        """重置引擎实例"""
        cls._instance = None
        cls._initialized = False

    def _initialize_engine(self):
        """安全初始化引擎"""
        try:
            self.logger.info(f"Initializing {self.engine_type} engine...")

            if self.engine_type == "vosk":
                self._init_vosk()
            elif self.engine_type == "whisper":
                self._init_whisper()
            elif self.engine_type == "microsoft":
                self._init_microsoft()
            elif self.engine_type == "tencent":
                self._init_tencent()
            elif self.engine_type == "sphinx":
                self._init_sphinx()
            else:
                raise ValueError(f"Unsupported engine type: {self.engine_type}")

        except ImportError as e:
            raise ImportError(f"Missing dependency: {str(e)}") from e
        except Exception as e:
            self.logger.error(f"Engine initialization failed: {str(e)}", exc_info=True)
            raise

    def _init_vosk(self):
        """初始化VOSK引擎"""
        if not os.path.isdir(self.model_config):
            raise ValueError(f"VOSK model must be directory: {self.model_config}")

        self.vosk_model = Model(self.model_config)
        self.logger.info(f"✅ VOSK model loaded | Language: {self.lang}")

    def _init_whisper(self):
        """初始化Whisper引擎"""
        device = "cuda" if torch.cuda.is_available() else "cpu"

        if not (str(self.model_config).endswith('.pt') or os.path.isdir(self.model_config)):
            raise ValueError(f"Whisper model must be .pt file or directory: {self.model_config}")

        self.whisper_model = whisper.load_model(self.model_config, device=device)
        self.logger.info(f"✅ Whisper model loaded | Device: {device}")

    def _init_microsoft(self):
        """初始化Microsoft Azure引擎"""
        self.logger.info("Initializing Microsoft Azure engine...")

        try:
            required_keys = {'api_key', 'region'}
            if missing := required_keys - self.model_config.keys():
                raise ValueError(f"Missing required Microsoft fields: {missing}")

            from microsoft_stt import MicrosoftSTT
            self.microsoft_client = MicrosoftSTT(
                api_key=self.model_config['api_key'],
                region=self.model_config['region'],
                lang=self.lang
            )

            self.logger.info(f"✅ Microsoft client initialized | Region: {self.model_config['region']}")

        except Exception as e:
            self.logger.error("Microsoft initialization failed", exc_info=True)
            raise

    def _init_tencent(self):
        """初始化腾讯云引擎"""
        if not TENCENT_SDK_AVAILABLE:
            raise ImportError("Tencent Cloud SDK not available")

        self.logger.info("Initializing Tencent Cloud engine...")

        try:
            required_keys = {'secret_id', 'secret_key'}
            if missing := required_keys - self.model_config.keys():
                raise ValueError(f"Missing required Tencent fields: {missing}")

            # 初始化凭证
            cred = credential.Credential(
                self.model_config['secret_id'],
                self.model_config['secret_key']
            )

            # 配置HTTP客户端
            http_profile = HttpProfile()
            http_profile.endpoint = "asr.tencentcloudapi.com"
            http_profile.req_timeout = 30

            client_profile = ClientProfile()
            client_profile.httpProfile = http_profile

            # 创建客户端
            self.tencent_client = asr_client.AsrClient(
                cred,
                self.model_config.get('region', 'ap-beijing'),
                client_profile
            )

            self.logger.info(f"✅ Tencent client initialized | Region: {self.model_config.get('region', 'ap-beijing')}")

        except Exception as e:
            self.logger.error("Tencent initialization failed", exc_info=True)
            raise

    def _init_sphinx(self):
        """初始化CMU Sphinx引擎"""
        required_files = {
            'hmm': 'acoustic-model',
            'lm': 'language-model.lm.bin',
            'dict': 'pronounciation-dict.dict'
        }

        self.sphinx_config = {}
        for key, pattern in required_files.items():
            matched = [f for f in os.listdir(self.model_config) if pattern in f.lower()]
            if not matched:
                raise FileNotFoundError(f"Missing {key} file: {pattern}")

            self.sphinx_config[key] = os.path.join(self.model_config, matched[0])

        self.logger.info(f"✅ Sphinx config loaded: {self.sphinx_config}")

    def transcribe(self, audio_path: str) -> str:
        """安全转录入口（添加实时显示功能）"""
        if not os.path.exists(audio_path):
            self.logger.error(f"File not exists: {audio_path}")
            return ""

        temp_path = None
        try:
            # 统一音频预处理
            if not self._is_valid_audio(audio_path):
                temp_path = self._convert_audio(audio_path)
                audio_path = temp_path

            # 路由到对应引擎前显示文件名
            filename = os.path.basename(audio_path)
            print(f"\n[开始识别] {filename}")  # 实时显示开始标记

            # 路由到对应引擎
            if self.engine_type == "vosk":
                result = self._transcribe_with_vosk(audio_path)
            elif self.engine_type == "whisper":
                result = self._transcribe_with_whisper(audio_path)
            elif self.engine_type == "microsoft":
                result = self._transcribe_with_microsoft(audio_path)
            elif self.engine_type == "tencent":
                result = self._transcribe_with_tencent(audio_path)
            elif self.engine_type == "sphinx":
                result = self._transcribe_with_sphinx(audio_path)
            else:
                raise ValueError(f"Unsupported engine type: {self.engine_type}")

            # 实时显示识别结果（核心添加点）
            print(f"[识别结果] {result}")  # 单独一行更清晰
            return result

        except Exception as e:
            self.logger.error(f"Transcription error: {str(e)}", exc_info=True)
            print(f"[识别失败] {os.path.basename(audio_path)}")  # 失败时也显示
            return ""
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass

    def _is_valid_audio(self, path: str) -> bool:
        """检查音频格式"""
        try:
            with wave.open(path, 'rb') as wf:
                if self.engine_type == "tencent":
                    return (wf.getnchannels() == 1 and
                            wf.getframerate() == 16000)
                elif self.engine_type == "vosk":
                    return (wf.getnchannels() == 1 and
                            wf.getframerate() == 16000 and
                            wf.getsampwidth() == 2)
                return True
        except:
            return False

    def _convert_audio(self, input_path: str) -> str:
        """音频格式转换"""
        temp_path = os.path.join(
            tempfile.gettempdir(),
            f"stt_convert_{os.path.basename(input_path)}.wav"
        )

        cmd = ["ffmpeg", "-i", input_path, "-ar", "16000", "-ac", "1"]
        if self.engine_type in ("tencent", "vosk"):
            cmd.extend(["-acodec", "pcm_s16le"])
        cmd.extend(["-y", temp_path])

        subprocess.run(cmd, check=True, capture_output=True)
        return temp_path

    def _transcribe_with_vosk(self, audio_path: str) -> str:
        """VOSK转录"""
        recognizer = KaldiRecognizer(self.vosk_model, 16000)
        result = []

        with open(audio_path, "rb") as f:
            while True:
                data = f.read(4000)
                if not data:
                    break
                if recognizer.AcceptWaveform(data):
                    res = json.loads(recognizer.Result())
                    if res["text"]:
                        result.append(res["text"])

        final_res = json.loads(recognizer.FinalResult())
        if final_res["text"]:
            result.append(final_res["text"])

        return " ".join(result).strip()

    def _transcribe_with_whisper(self, audio_path: str) -> str:
        """Whisper转录"""
        result = self.whisper_model.transcribe(audio_path, language=self.lang)
        return result["text"].strip()

    def _transcribe_with_microsoft(self, audio_path: str) -> str:
        """Microsoft转录"""
        return self.microsoft_client.transcribe(audio_path)

    def _transcribe_with_tencent(self, audio_path: str) -> str:
        """腾讯云转录"""
        try:
            # 1. 读取音频文件
            with open(audio_path, 'rb') as f:
                audio_data = f.read()

            # 2. 创建识别任务请求
            req = tencent_models.CreateRecTaskRequest()
            req.EngineModelType = "16k_zh"
            req.ChannelNum = 1
            req.SourceType = 1  # 1表示语音数据是base64编码
            req.ResTextFormat = 0  # 0表示识别结果文本
            req.Data = base64.b64encode(audio_data).decode('utf-8')

            # 3. 发送请求
            resp = self.tencent_client.CreateRecTask(req)
            task_id = resp.Data.TaskId
            self.logger.info(f"Tencent task created | TaskId: {task_id}")

            # 4. 获取结果
            start_time = time.time()
            while time.time() - start_time < 30:  # 30秒超时
                req = tencent_models.DescribeTaskStatusRequest()
                req.TaskId = task_id
                status = self.tencent_client.DescribeTaskStatus(req)

                if status.Data.Status == 2:  # 成功
                    # ============== 新增内容开始 ==============
                    raw_result = status.Data.Result
                    if not raw_result:
                        return ""
                    # 方案2：正则处理（应对多段情况）

                    clean_result = re.sub(r'\[\d+:\d+\.\d+,\d+:\d+\.\d+\]\s*', '', raw_result)

                    return clean_result
                    # ============== 新增内容结束 ==============

                elif status.Data.Status == 3:  # 失败
                    raise Exception(f"Recognition failed: {status.Data.StatusStr}")

                time.sleep(1)

            raise Exception("Result timeout")

        except Exception as e:
            self.logger.error(f"Tencent transcription failed: {str(e)}")
            return ""

    def _transcribe_with_sphinx(self, audio_path: str) -> str:
        """Sphinx转录"""
        from pocketsphinx import AudioFile
        audio = AudioFile(
            audio_file=audio_path,
            **self.sphinx_config
        )
        return " ".join([str(seg) for seg in audio]).strip()

    def test_model(self, test_audio: Optional[str] = None) -> str:
        """测试模型"""
        if not test_audio:
            test_audio = os.path.join(tempfile.gettempdir(), f"test_{self.lang}.wav")
            t = np.linspace(0, 1, 16000)
            audio_data = 0.5 * np.sin(2 * np.pi * 440 * t)
            sf.write(test_audio, audio_data, 16000)

        self.logger.info(f"🔊 Testing model | File: {os.path.basename(test_audio)}")
        result = self.transcribe(test_audio)

        if test_audio.startswith(tempfile.gettempdir()):
            try:
                os.remove(test_audio)
            except:
                pass

        return result

    @classmethod
    def reset_instance(cls):
        """重置单例实例"""
        cls._instance = None
        cls._initialized = False


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    # 测试示例
    test_cases = [
        ("vosk", "models/vosk/zh-CN"),
        ("whisper", "models/whisper/medium.pt"),
        ("tencent", {
            "secret_id": "your-secret-id",
            "secret_key": "your-secret-key",
            "region": "ap-beijing"
        }),
        ("sphinx", "models/sphinx/zh-CN")
    ]

    for engine_type, model_config in test_cases:
        print(f"\n=== Testing {engine_type.upper()} ===")
        try:
            engine = STTEngine(model_config, engine_type=engine_type)
            print("Test result:", engine.test_model())
        except Exception as e:
            print(f"Test failed: {str(e)}")