import matplotlib.pyplot as plt
import numpy as np
import soundfile as sf


class AudioDebugger:
    @staticmethod
    def analyze_audio(audio_path: str):
        """音频分析工具"""
        try:
            data, sample_rate = sf.read(audio_path)

            # 绘制波形图
            plt.figure(figsize=(12, 4))
            plt.plot(np.linspace(0, len(data) / sample_rate, len(data)), data)
            plt.title(f"Audio Waveform: {os.path.basename(audio_path)}")
            plt.xlabel("Time (s)")
            plt.ylabel("Amplitude")
            plt.show()

            # 打印关键信息
            print(f"\n🔍 音频分析报告:")
            print(f"- 采样率: {sample_rate}Hz")
            print(f"- 时长: {len(data) / sample_rate:.2f}秒")
            print(f"- 声道数: {data.shape[1] if len(data.shape) > 1 else 1}")
            print(f"- 最大振幅: {np.max(np.abs(data)):.2f}")

        except Exception as e:
            print(f"分析失败: {str(e)}")