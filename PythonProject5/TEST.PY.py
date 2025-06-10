#!/usr/bin/env python3
"""
腾讯云ASR API终极测试脚本
兼容所有SDK版本的参数打印方案
"""
import base64
import json
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.asr.v20190614 import asr_client, models


def print_debug_info(request):
    """万能参数打印方法"""
    print("\n=== 调试信息 ===")

    # 方法1：尝试获取_serialized属性
    if hasattr(request, "_serialized"):
        print("[方式1] 通过_serialized获取:")
        print(json.dumps(request._serialized, indent=2, ensure_ascii=False))
        return

    # 方法2：尝试序列化整个request对象
    try:
        print("[方式2] 通过request.__dict__获取:")
        filtered_params = {k: v for k, v in request.__dict__.items()
                           if not k.startswith('_') and v is not None}
        print(json.dumps(filtered_params, indent=2, ensure_ascii=False))
        return
    except Exception as e:
        pass

    # 方法3：手动构建参数列表
    print("[方式3] 手动提取已知参数:")
    known_params = {
        "EngineModelType": request.EngineModelType,
        "ChannelNum": request.ChannelNum,
        "SourceType": request.SourceType,
        "ResTextFormat": request.ResTextFormat,
        "Data": f"[Base64数据，长度:{len(request.Data) if request.Data else 0}]"
    }
    print(json.dumps(known_params, indent=2, ensure_ascii=False))


def test_tencent_asr():
    try:
        # ========== 1. 初始化客户端 ==========
        cred = credential.Credential(
            secret_id="AKIDHQN6f1PgwCJWfkfx9EKuT0RktvqlgU6W",  # 替换为您的真实ID
            secret_key="KSN5oPo02475ztnoheacoridXItMjmiR"  # 替换为您的真实KEY
        )
        client = asr_client.AsrClient(cred, "ap-shanghai")

        # ========== 2. 生成测试音频 ==========
        pcm_data = b'\x00\x00' * 16000  # 1秒静音音频
        audio_data = base64.b64encode(pcm_data).decode('utf-8')

        # ========== 3. 构建请求 ==========
        request = models.CreateRecTaskRequest()
        request.EngineModelType = "16k_zh"
        request.ChannelNum = 1
        request.SourceType = 1
        request.ResTextFormat = 0
        request.Data = audio_data

        # 打印调试信息
        print_debug_info(request)

        # ========== 4. 发送请求 ==========
        print("\n=== API调用结果 ===")
        response = client.CreateRecTask(request)
        print(f"✅ 识别任务创建成功 | TaskId: {response.Data.TaskId}")

        # 打印完整响应（可选）
        print("\n=== 完整响应 ===")
        print(json.dumps(response.__dict__, default=str, indent=2))

    except TencentCloudSDKException as e:
        print("\n❌ 腾讯云SDK异常:")
        print(json.dumps(e.__dict__, default=str, indent=2))
    except Exception as e:
        print(f"\n⚠️ 系统异常: {type(e).__name__}: {str(e)}")


if __name__ == "__main__":
    print("=== 腾讯云ASR终极测试 ===")
    test_tencent_asr()