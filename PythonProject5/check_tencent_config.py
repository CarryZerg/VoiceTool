import os
import sys
import json
import logging
from typing import Dict, Optional


def validate_tencent_config(config_path: str) -> bool:
    """增强版腾讯云配置验证工具"""
    required_keys = ['secret_id', 'secret_key']
    optional_keys = ['engine_type', 'filter_dirty', 'filter_mod', 'hotword_id']

    print(f"\n🔍 正在验证腾讯云配置文件: {os.path.basename(config_path)}")

    try:
        # 1. 检查文件是否存在
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        # 2. 检查是否是JSON文件
        if not config_path.lower().endswith('.json'):
            raise ValueError("配置文件必须是.json格式")

        # 3. 解析文件内容
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # 4. 验证必要字段
        missing = [k for k in required_keys if k not in config]
        if missing:
            raise ValueError(f"缺少必要字段: {missing}")

        # 5. 验证字段类型
        if not isinstance(config['secret_id'], str) or not config['secret_id'].startswith('AKID'):
            raise ValueError("secret_id 格式不正确 (应以AKID开头)")

        # 打印验证结果
        print("\n✅ 配置验证通过")
        print("=" * 40)
        print("关键配置信息:")
        print(f"SecretID: {'*' * 8}{config['secret_id'][-4:]}")
        print(f"SecretKey: {'*' * 8}{config['secret_key'][-4:]}")
        print("\n可选配置:")
        for k in optional_keys:
            val = config.get(k, '<未设置>')
            print(f"{k:>12}: {val}")
        print("=" * 40)

        return True

    except Exception as e:
        print(f"\n❌ 验证失败: {str(e)}", file=sys.stderr)
        return False


def interactive_mode():
    """交互式验证模式"""
    print("腾讯云配置验证工具")
    print("=" * 40)

    while True:
        config_path = input("请输入配置文件路径(或输入q退出): ").strip()
        if config_path.lower() == 'q':
            break

        validate_tencent_config(config_path)
        print("\n")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 命令行参数模式
        validate_tencent_config(sys.argv[1])
    else:
        # 交互模式
        interactive_mode()