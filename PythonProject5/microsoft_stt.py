import os
import requests
import json
import logging
from datetime import datetime, timedelta
import time


class MicrosoftSTT:
    def __init__(self, api_key=None, region="eastus", lang="zh-CN"):
        """
        初始化Microsoft语音识别API

        参数:
            api_key: Azure认知服务API密钥
            region: Azure服务区域(如"eastus")
            lang: 识别语言(如"zh-CN"中文,"en-US"英文)
        """
        self.logger = self._setup_logger()
        self.api_key = api_key
        self.region = region
        self.lang = lang
        self.token = None
        self.token_expiry = None

        # API端点
        self.token_url = f"https://{self.region}.api.cognitive.microsoft.com/sts/v1.0/issueToken"
        self.stt_url = f"https://{self.region}.stt.speech.microsoft.com/speech/recognition/conversation/cognitiveservices/v1"

        self.logger.info(f"Microsoft语音识别API初始化完成 | 区域: {self.region} | 语言: {self.lang}")

    def _setup_logger(self):
        """设置日志记录器"""
        logger = logging.getLogger("MicrosoftSTT")
        logger.setLevel(logging.INFO)

        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)

        logger.propagate = False
        return logger

    def _get_auth_token(self):
        """获取认证token"""
        if self.token and self.token_expiry and datetime.now() < self.token_expiry:
            return self.token

        headers = {
            'Ocp-Apim-Subscription-Key': self.api_key,
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        try:
            response = requests.post(self.token_url, headers=headers)
            response.raise_for_status()

            self.token = response.text
            self.token_expiry = datetime.now() + timedelta(minutes=9)  # token有效期10分钟，提前1分钟更新
            self.logger.info("Microsoft API token获取成功")
            return self.token

        except Exception as e:
            self.logger.error(f"获取Microsoft API token失败: {str(e)}")
            raise RuntimeError(f"获取Microsoft API token失败: {str(e)}")

    def transcribe(self, audio_path):
        """
        转录音频文件

        参数:
            audio_path: 音频文件路径(支持WAV, MP3等格式)

        返回:
            识别文本
        """
        if not os.path.exists(audio_path):
            self.logger.error(f"音频文件不存在: {audio_path}")
            raise FileNotFoundError(f"音频文件不存在: {audio_path}")

        # 获取认证token
        token = self._get_auth_token()

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'audio/wav; codec=audio/pcm; samplerate=16000',
            'Accept': 'application/json'
        }

        params = {
            'language': self.lang,
            'format': 'detailed'  # 获取更详细的结果
        }

        try:
            # 读取音频文件
            with open(audio_path, 'rb') as audio_file:
                audio_data = audio_file.read()

            self.logger.info(f"开始识别: {os.path.basename(audio_path)}")
            start_time = time.time()

            response = requests.post(
                self.stt_url,
                headers=headers,
                params=params,
                data=audio_data
            )

            response.raise_for_status()
            result = response.json()

            duration = time.time() - start_time
            self.logger.info(f"识别完成 | 耗时: {duration:.2f}s | 状态: {result.get('RecognitionStatus')}")

            if result['RecognitionStatus'] == 'Success':
                return result['DisplayText']
            else:
                self.logger.error(f"识别失败: {result.get('RecognitionStatus')}")
                return ""

        except Exception as e:
            self.logger.error(f"识别过程中出错: {str(e)}")
            raise RuntimeError(f"Microsoft语音识别失败: {str(e)}")

    @staticmethod
    def get_supported_languages():
        """返回支持的语言列表"""
        return {
            'zh-CN': '中文(普通话)',
            'en-US': '英文(美国)',
            'ja-JP': '日语',
            'ko-KR': '韩语',
            'fr-FR': '法语',
            'de-DE': '德语',
            'es-ES': '西班牙语',
            'ru-RU': '俄语'
        }