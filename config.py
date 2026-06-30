"""
DCS 数据平台 — 集中配置模块
所有 InfluxDB 连接信息、敏感配置、常量和校验函数统一从此处读取。

使用方式：
    from config import INFLUX_URL, INFLUX_TOKEN, validate_params, APP_TOKEN, ...

环境变量（生产环境通过 .env 或系统环境变量设置）：
    INFLUX_URL      — InfluxDB 地址（默认 http://10.56.128.202:8086）
    INFLUX_TOKEN    — InfluxDB 认证 Token（必填）
    INFLUX_ORG      — InfluxDB 组织（默认 myOrg）
    INFLUX_BUCKET   — InfluxDB Bucket（默认 islag）
    APP_TOKEN       — Web 应用访问 Token（必填，用于 Flask API 认证）
"""
import os
import sys
import re
from pathlib import Path

# 自动加载 .env 文件
# PyInstaller 打包时 exe 目录优先，开发时项目根目录
try:
    from dotenv import load_dotenv
    if getattr(sys, 'frozen', False):
        _ENV_FILE = Path(sys.executable).parent / ".env"
    else:
        _ENV_FILE = Path(__file__).resolve().parent / ".env"
    if _ENV_FILE.exists():
        load_dotenv(_ENV_FILE)
except ImportError:
    pass

# ========================
# InfluxDB 连接配置
# ========================
INFLUX_URL = os.environ.get("INFLUX_URL", "http://10.56.128.202:8086")
INFLUX_TOKEN = os.environ.get("INFLUX_TOKEN", "")
INFLUX_ORG = os.environ.get("INFLUX_ORG", "myOrg")
INFLUX_BUCKET = os.environ.get("INFLUX_BUCKET", "islag")
INFLUX_MEASUREMENT = "DCS"
INFLUX_TIMEOUT_MS = int(os.environ.get("INFLUX_TIMEOUT_MS", "180000"))

# ========================
# Web 应用配置
# ========================
FLASK_HOST = os.environ.get("FLASK_HOST", "0.0.0.0")
FLASK_PORT = int(os.environ.get("FLASK_PORT", "5000"))
APP_TOKEN = os.environ.get("APP_TOKEN", "")

# ========================
# 业务常量
# ========================
MAX_QUERY_HOURS = 24          # Web 查询最大时间跨度（小时）
MAX_WEB_ROWS = 50000          # Web 页面展示行数上限
MAX_EXCEL_ROWS = 200000       # Excel 导出行数上限
EXPORT_CHUNK_HOURS = 1        # 每日导出的分块大小（小时）
LOCAL_TZ_OFFSET_HOURS = 8     # 北京时间 UTC+8

# ========================
# 参数名校验（防 Flux 注入）
# ========================
_PARAM_NAME_RE = re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')

def validate_param_name(name: str) -> str:
    """校验单个参数名仅含合法字符，防止破坏 Flux 查询语法。

    Raises:
        ValueError: 参数名包含非法字符
    """
    if not _PARAM_NAME_RE.match(name):
        raise ValueError(f"非法参数名（仅允许字母/数字/下划线）: {name}")
    return name

def validate_params(param_list: list[str]) -> list[str]:
    """批量校验参数名，任一不合法即抛异常。"""
    return [validate_param_name(p) for p in param_list]

def sanitize_param_for_flux(param_list: list[str]) -> str:
    """将参数列表校验后构建为 Flux filter 片段。

    示例:
        sanitize_param_for_flux(["LT_AA", "LT_BB"])
        → 'r.param == "LT_AA" or r.param == "LT_BB"'
    """
    validated = validate_params(param_list)
    return " or ".join([f'r.param == "{p}"' for p in validated])

# ========================
# 启动时健康检查
# ========================
def check_config() -> list[str]:
    """检查必要配置是否就绪，返回警告消息列表。"""
    warnings = []
    if not INFLUX_TOKEN:
        warnings.append("[WARN] INFLUX_TOKEN 未设置 — InfluxDB 连接将失败")
    if not APP_TOKEN:
        warnings.append("[WARN] APP_TOKEN 未设置 — Web API 认证将放行所有请求（仅内网安全）")
    return warnings
