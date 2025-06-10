"""
腾讯云ASR API终极稳定版（已通过实测验证）
修复所有已知问题，增强错误处理
"""
import base64
import json
import logging
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.asr.v20190614 import asr_client, models

# 配置更详细的日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class TencentSTT:
    def __init__(self, secret_id, secret_key, region="ap-shanghai"):
        """初始化客户端"""
        self.cred = credential.Credential(secret_id, secret_key)
        self.client = asr_client.AsrClient(self.cred, region)
        logging.info("ASR客户端初始化完成 | Region: %s", region)

    @staticmethod
    def _print_debug_info(request):
        """安全打印请求信息（兼容所有SDK版本）"""
        debug_info = {
            "EngineModelType": request.EngineModelType,
            "ChannelNum": request.ChannelNum,
            "SourceType": request.SourceType,
            "ResTextFormat": request.ResTextFormat,
            "DataLength": len(request.Data) if request.Data else 0
        }
        logging.debug("请求参数详情:\n%s", json.dumps(debug_info, indent=2))

    def recognize(self, audio_data, model="16k_zh"):
        """
        执行语音识别
        :param audio_data: PCM音频的bytes对象
        :param model: 引擎模型
        :return: (success, task_id/error_message)
        """
        try:
            # 构建请求
            request = models.CreateRecTaskRequest()
            request.EngineModelType = model
            request.ChannelNum = 1
            request.SourceType = 1  # 1表示语音数据是base64编码
            request.ResTextFormat = 0  # 0表示识别结果文本
            request.Data = base64.b64encode(audio_data).decode('utf-8')

            self._print_debug_info(request)

            # 发送请求
            response = self.client.CreateRecTask(request)

            if hasattr(response, "Data") and hasattr(response.Data, "TaskId"):
                task_id = response.Data.TaskId
                logging.info("识别任务创建成功 | TaskId: %d", task_id)
                return True, task_id
            else:
                error_msg = "响应中缺少TaskId字段"
                logging.error(error_msg)
                return False, error_msg

        except TencentCloudSDKException as e:
            error_msg = f"SDK错误: {e.get_code()} - {e.get_message()}"
            logging.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"系统错误: {type(e).__name__} - {str(e)}"
            logging.error(error_msg)
            return False, error_msg

    def get_result(self, task_id, timeout=30):
        """
        获取识别结果
        :param task_id: 任务ID
        :param timeout: 最长等待时间(秒)
        :return: (success, result/error_message)
        """
        try:
            import time
            start_time = time.time()

            while time.time() - start_time < timeout:
                req = models.DescribeTaskStatusRequest()
                req.TaskId = task_id
                status = self.client.DescribeTaskStatus(req)

                if status.Data.Status == 2:  # 成功
                    return True, status.Data.Result
                elif status.Data.Status == 3:  # 失败
                    return False, "识别任务失败"

                time.sleep(1)  # 每秒检查一次

            return False, "获取结果超时"

        except Exception as e:
            return False, f"查询错误: {str(e)}"


if __name__ == "__main__":
    # 示例用法
    print("=== 腾讯云ASR测试（稳定生产版） ===")

    # 1. 初始化（替换为您的真实密钥）
    asr = TencentSTT(
        secret_id="AKIDHQN6f1PgwCJWfkfx9EKuT0RktvqlgU6W",
        secret_key="KSN5oPo02475ztnoheacoridXItMjmiR"
    )

    # 2. 生成测试音频（1秒静音）
    test_audio = b'\x00\x00' * 16000

    # 3. 执行识别
    success, task_id = asr.recognize(test_audio)

    if success:
        print(f"✅ 任务已提交 | TaskId: {task_id}")

        # 4. 获取结果（实际使用建议异步处理）
        success, result = asr.get_result(task_id)
        if success:
            print(f"识别结果: {result}")
        else:
            print(f"获取结果失败: {result}")
    else:
        print(f"❌ 识别失败: {task_id}")