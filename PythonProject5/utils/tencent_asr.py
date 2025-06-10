# tencent_asr.py
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.asr.v20190614 import asr_client, models
import base64
import json


class TencentASR:
    def __init__(self, secret_id, secret_key, region="ap-guangzhou"):
        self.cred = credential.Credential(secret_id, secret_key)
        self.region = region
        self.engine_type = "16k_zh"  # 默认16k中文通用

    def transcribe(self, audio_path, custom_vocab=None):
        try:
            # 1. 读取并编码音频
            with open(audio_path, "rb") as f:
                audio_data = base64.b64encode(f.read()).decode('utf-8')

            # 2. 创建请求对象
            req = models.SentenceRecognitionRequest()
            req.ProjectId = 0  # 默认项目ID
            req.SubServiceType = 2  # 实时语音识别
            req.EngineModelType = self.engine_type
            req.SourceType = 1  # 音频数据(base64)
            req.Data = audio_data
            req.DataLen = len(audio_data)

            if custom_vocab:
                req.HotwordId = custom_vocab  # 使用自定义热词表

            # 3. 发起请求
            http_profile = HttpProfile(endpoint="asr.tencentcloudapi.com")
            client_profile = ClientProfile(httpProfile=http_profile)
            client = asr_client.AsrClient(self.cred, self.region, client_profile)
            resp = client.SentenceRecognition(req)

            return resp.Result if resp.Result else ""

        except Exception as e:
            raise Exception(f"腾讯云ASR调用失败: {str(e)}")