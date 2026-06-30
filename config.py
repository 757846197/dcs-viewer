"""
DCS 数据平台 — 集中配置模块
支持运行时配置：exe 旁边的 dcs_config.json 优先于环境变量/内置默认值
Web 页面可通过 /settings 修改配置（密码保护）
"""
import json
import os
import sys
import re
from pathlib import Path

# === 配置文件路径 ===
def _get_exe_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

_EXE_DIR = _get_exe_dir()
_CONFIG_FILE = _EXE_DIR / "dcs_config.json"

# === 内置默认值 ===
_DEFAULTS = {
    "INFLUX_URL": "http://10.56.128.202:8086",
    "INFLUX_TOKEN": "",
    "INFLUX_ORG": "myOrg",
    "INFLUX_BUCKET": "islag",
    "INFLUX_TIMEOUT_MS": 180000,
    "FLASK_HOST": "0.0.0.0",
    "FLASK_PORT": 5000,
    "APP_TOKEN": "dcs2026",
    "SETTINGS_PASSWORD": "123456",
}

# === 加载: dcs_config.json > 环境变量 > 默认值 ===
def _load_config():
    cfg = dict(_DEFAULTS)

    # 1) 环境变量覆盖
    for k in _DEFAULTS:
        env_val = os.environ.get(k)
        if env_val is not None:
            cfg[k] = env_val

    # 2) exe 旁边的 dcs_config.json 最高优先级
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if k in cfg:
                    cfg[k] = v
        except Exception:
            pass

    return cfg

_cfg = _load_config()

# === 导出常量 ===
INFLUX_URL = str(_cfg["INFLUX_URL"])
INFLUX_TOKEN = str(_cfg["INFLUX_TOKEN"])
INFLUX_ORG = str(_cfg["INFLUX_ORG"])
INFLUX_BUCKET = str(_cfg["INFLUX_BUCKET"])
INFLUX_MEASUREMENT = "DCS"
INFLUX_TIMEOUT_MS = int(_cfg["INFLUX_TIMEOUT_MS"])
FLASK_HOST = str(_cfg["FLASK_HOST"])
FLASK_PORT = int(_cfg["FLASK_PORT"])
APP_TOKEN = str(_cfg["APP_TOKEN"])
SETTINGS_PASSWORD = str(_cfg["SETTINGS_PASSWORD"])

# === 业务常量 ===
MAX_QUERY_HOURS = 24
MAX_WEB_ROWS = 50000
MAX_EXCEL_ROWS = 200000
EXPORT_CHUNK_HOURS = 1
LOCAL_TZ_OFFSET_HOURS = 8

# === 运行时配置保存 ===
def save_runtime_config(data: dict):
    """保存配置到 dcs_config.json，只保存非空值"""
    current = {}
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                current = json.load(f)
        except Exception:
            pass
    current.update({k: v for k, v in data.items() if v != "" and v is not None})
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(current, f, indent=2)
    # 立即生效
    global INFLUX_URL, INFLUX_TOKEN, INFLUX_ORG, INFLUX_BUCKET, INFLUX_TIMEOUT_MS, FLASK_HOST, FLASK_PORT, APP_TOKEN, SETTINGS_PASSWORD
    _new = _load_config()
    INFLUX_URL = str(_new["INFLUX_URL"])
    INFLUX_TOKEN = str(_new["INFLUX_TOKEN"])
    INFLUX_ORG = str(_new["INFLUX_ORG"])
    INFLUX_BUCKET = str(_new["INFLUX_BUCKET"])
    INFLUX_TIMEOUT_MS = int(_new["INFLUX_TIMEOUT_MS"])
    FLASK_HOST = str(_new["FLASK_HOST"])
    FLASK_PORT = int(_new["FLASK_PORT"])
    APP_TOKEN = str(_new["APP_TOKEN"])
    SETTINGS_PASSWORD = str(_new["SETTINGS_PASSWORD"])

def get_current_config() -> dict:
    """返回当前所有配置（隐藏 token 和密码）"""
    return {
        "INFLUX_URL": INFLUX_URL,
        "INFLUX_ORG": INFLUX_ORG,
        "INFLUX_BUCKET": INFLUX_BUCKET,
        "INFLUX_TIMEOUT_MS": INFLUX_TIMEOUT_MS,
        "FLASK_HOST": FLASK_HOST,
        "FLASK_PORT": FLASK_PORT,
        "APP_TOKEN": "***" + APP_TOKEN[-4:] if len(APP_TOKEN) > 4 else "***",
    }

# === 参数名校验 ===
_PARAM_NAME_RE = re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')

def validate_param_name(name: str) -> str:
    if not _PARAM_NAME_RE.match(name):
        raise ValueError(f"非法参数名: {name}")
    return name

def validate_params(param_list: list[str]) -> list[str]:
    return [validate_param_name(p) for p in param_list]

def sanitize_param_for_flux(param_list: list[str]) -> str:
    validated = validate_params(param_list)
    return " or ".join([f'r.param == "{p}"' for p in validated])

def check_config() -> list[str]:
    warnings = []
    if not INFLUX_TOKEN:
        warnings.append("[WARN] INFLUX_TOKEN 未设置 — InfluxDB 连接将失败")
    return warnings
