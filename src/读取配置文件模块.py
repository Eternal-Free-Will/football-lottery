# config.py ✅ 用于统一读取配置文件

import json
import os

def load_config(config_path="配置.json"):
    """
    从指定路径读取 JSON 配置文件，返回一个包含 issue 和 date 的字典。
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件未找到：{config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 校验必填字段
    if "issue" not in cfg or "date" not in cfg:
        raise KeyError("配置文件缺少 issue 或 date 字段")

    return cfg["issue"], cfg["date"]
