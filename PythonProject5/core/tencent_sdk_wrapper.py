# tencent_sdk_wrapper.py
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.asr.v20190614 import asr_client, models


class TencentSDKWrapper:
    def __init__(self, secret_id, secret_key):
        self.cred = credential.Credential(secret_id, secret_key)
        self.client = self._init_client()

    def _init_client(self, region="ap-guangzhou"):
        """初始化SDK客户端"""
        http_profile = HttpProfile(endpoint="asr.tencentcloudapi.com")
        client_profile = ClientProfile(httpProfile=http_profile)
        return asr_client.AsrClient(self.cred, region, client_profile)

    def recognize(self, audio_data, engine_type="16k_zh"):
        """核心识别方法（同步调用）"""
        req = models.SentenceRecognitionRequest()
        req.EngineModelType = engine_type
        req.SourceType = 1  # 1表示base64音频
        req.Data = base64.b64encode(audio_data).decode('utf-8')
        req.DataLen = len(req.Data)

        try:
            resp = self.client.SentenceRecognition(req)
            return resp.Result
        except TencentCloudSDKException as e:
            raise RuntimeError(f"SDK调用失败: {e.message}")

    # 保留原有热词功能
    def set_hotwords(self, vocab_id):
        """动态设置热词表"""
        self.engine_type = f"{self.engine_type.split('_')[0]}_hotword"  # e.g. 16k_zh_hotword
        self.vocab_id = vocab_id