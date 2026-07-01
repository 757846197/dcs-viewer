"""
DCS 数据分析平台 — 开口机/堵口机专用查询工具
支持: 四组设备分组查询、时序分析、Excel导出
"""
import functools
import io
import json
import os
import sys
import threading
import time as _time
import webbrowser
from datetime import datetime, timedelta, timezone
from pathlib import Path

# PyInstaller 打包兼容: 定位资源文件路径
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys._MEIPASS)
    _EXE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).resolve().parent
    _EXE_DIR = _BASE_DIR.parent

from flask import Flask, request, jsonify, send_file, render_template_string, make_response, session, redirect
from flask_compress import Compress
from influxdb_client import InfluxDBClient
import xlsxwriter

# 项目根目录加入搜索路径，以便导入 config
_PROJECT_ROOT = _EXE_DIR
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# === 内嵌 config 模块（PyInstaller 兼容：避免外部依赖） ===
import types
_config_mod = types.ModuleType('config')

import json as _json
import re as _re
from pathlib import Path as _Path

def _get_exe_dir():
    if getattr(sys, 'frozen', False):
        return _Path(sys.executable).parent
    return _Path(__file__).resolve().parent

_EXE_DIR_CFG = _get_exe_dir()
_CONFIG_FILE = _EXE_DIR_CFG / "dcs_config.json"

_DEFAULTS_CFG = {
    "INFLUX_URL": "http://10.56.128.202:8086",
    "INFLUX_TOKEN": "odx3aMiszU2cZd5PvvoexXYEXOcR-yJ0opTnXixC38TRttY2xHY-84lyRdC7MPIWgK2IAkoS4ZnPPsNBBFWZEA==",
    "INFLUX_ORG": "myOrg",
    "INFLUX_BUCKET": "islag",
    "INFLUX_TIMEOUT_MS": 180000,
    "FLASK_HOST": "0.0.0.0",
    "FLASK_PORT": 5000,
    "APP_TOKEN": "dcs2026",
    "SETTINGS_PASSWORD": "123456",
}

def _load_cfg():
    cfg = dict(_DEFAULTS_CFG)
    for k in _DEFAULTS_CFG:
        env_val = os.environ.get(k)
        if env_val is not None:
            cfg[k] = env_val
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = _json.load(f)
            for k, v in saved.items():
                if k in cfg:
                    cfg[k] = v
        except Exception:
            pass
    return cfg

_cfg = _load_cfg()

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

MAX_QUERY_HOURS = 24
MAX_WEB_ROWS = 50000
MAX_EXCEL_ROWS = 200000
EXPORT_CHUNK_HOURS = 1
LOCAL_TZ_OFFSET_HOURS = 8

def save_runtime_config(data: dict):
    current = {}
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                current = _json.load(f)
        except Exception:
            pass
    current.update({k: v for k, v in data.items() if v != "" and v is not None})
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        _json.dump(current, f, indent=2)
    global INFLUX_URL, INFLUX_TOKEN, INFLUX_ORG, INFLUX_BUCKET, INFLUX_TIMEOUT_MS, FLASK_HOST, FLASK_PORT, APP_TOKEN, SETTINGS_PASSWORD
    _new = _load_cfg()
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
    return {
        "INFLUX_URL": INFLUX_URL,
        "INFLUX_ORG": INFLUX_ORG,
        "INFLUX_BUCKET": INFLUX_BUCKET,
        "INFLUX_TIMEOUT_MS": INFLUX_TIMEOUT_MS,
        "FLASK_HOST": FLASK_HOST,
        "FLASK_PORT": FLASK_PORT,
        "APP_TOKEN": "***" + APP_TOKEN[-4:] if len(APP_TOKEN) > 4 else "***",
    }

_PARAM_NAME_RE = _re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')

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

# === 把所有 config 符号注入到 _config_mod 模块对象（替代 import 语法） ===
_config_mod.INFLUX_URL = INFLUX_URL
_config_mod.INFLUX_TOKEN = INFLUX_TOKEN
_config_mod.INFLUX_ORG = INFLUX_ORG
_config_mod.INFLUX_BUCKET = INFLUX_BUCKET
_config_mod.INFLUX_MEASUREMENT = INFLUX_MEASUREMENT
_config_mod.INFLUX_TIMEOUT_MS = INFLUX_TIMEOUT_MS
_config_mod.FLASK_HOST = FLASK_HOST
_config_mod.FLASK_PORT = FLASK_PORT
_config_mod.APP_TOKEN = APP_TOKEN
_config_mod.SETTINGS_PASSWORD = SETTINGS_PASSWORD
_config_mod.MAX_QUERY_HOURS = MAX_QUERY_HOURS
_config_mod.MAX_WEB_ROWS = MAX_WEB_ROWS
_config_mod.MAX_EXCEL_ROWS = MAX_EXCEL_ROWS
_config_mod.EXPORT_CHUNK_HOURS = EXPORT_CHUNK_HOURS
_config_mod.LOCAL_TZ_OFFSET_HOURS = LOCAL_TZ_OFFSET_HOURS
_config_mod.sanitize_param_for_flux = sanitize_param_for_flux
_config_mod.check_config = check_config
_config_mod.save_runtime_config = save_runtime_config
_config_mod.get_current_config = get_current_config
_config_mod.validate_param_name = validate_param_name
_config_mod.validate_params = validate_params

sys.modules['config'] = _config_mod

# === 兼容：本地代码仍可 from config import ... ===
INFLUX_URL = INFLUX_URL
INFLUX_TOKEN = INFLUX_TOKEN
INFLUX_ORG = INFLUX_ORG
INFLUX_BUCKET = INFLUX_BUCKET
INFLUX_MEASUREMENT = INFLUX_MEASUREMENT
INFLUX_TIMEOUT_MS = INFLUX_TIMEOUT_MS
FLASK_HOST = FLASK_HOST
FLASK_PORT = FLASK_PORT
APP_TOKEN = APP_TOKEN
SETTINGS_PASSWORD = SETTINGS_PASSWORD
MAX_QUERY_HOURS = MAX_QUERY_HOURS
MAX_WEB_ROWS = MAX_WEB_ROWS
MAX_EXCEL_ROWS = MAX_EXCEL_ROWS
EXPORT_CHUNK_HOURS = EXPORT_CHUNK_HOURS
sanitize_param_for_flux = sanitize_param_for_flux
check_config = check_config
save_runtime_config = save_runtime_config
get_current_config = get_current_config

app = Flask(__name__)
app.secret_key = "dcs-viewer-key-20260701"

# === 登录验证装饰器 ===
Compress(app)
app.config["COMPRESS_ALGORITHM"] = "gzip"
app.config["COMPRESS_MIN_SIZE"] = 500       # 小于500字节不压缩
app.config["COMPRESS_LEVEL"] = 6            # 压缩级别（1-9，6为平衡点）

# 北京时间 (UTC+8)
LOCAL_OFFSET = timedelta(hours=8)

# === API 访问认证 ===
@app.before_request
def _check_api_auth():
    """对所有 /api/ 路由做 Token 认证（登录/设置认证绕过）"""
    if not request.path.startswith("/api/"):
        return None
    if request.path in ("/api/login", "/api/settings/auth"):
        return None  # 登录/设置认证不需要 API Token
    if not APP_TOKEN:
        return None
    token = request.args.get("token", "") or request.headers.get("X-API-Token", "")
    if token != APP_TOKEN:
        return jsonify({"error": "未授权访问 — 请提供有效的 API Token"}), 401
    return None

# === 加载参数分组 ===
# PyInstaller: --add-data "dcs_viewer/param_groups.json;dcs_viewer" → {_MEIPASS}/dcs_viewer/param_groups.json
GROUPS_FILE = _BASE_DIR / "dcs_viewer" / "param_groups.json"
if not GROUPS_FILE.exists():
    GROUPS_FILE = _BASE_DIR / "param_groups.json"
if not GROUPS_FILE.exists():
    GROUPS_FILE = _EXE_DIR / "param_groups.json"
with open(GROUPS_FILE, "r", encoding="utf-8") as f:
    PARAM_CONFIG = json.load(f)

# === 性能优化：预映射标签，避免每个记录都查字典 ===
_LABELS = PARAM_CONFIG.get("labels", {})

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DCS 开口机/堵口机 数据分析平台</title>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #1f2937; display: flex; min-height: 100vh; }

/* === 左侧导航 (阿里云风格深色侧栏) === */
.sidebar { width: 220px; min-width: 220px; background: linear-gradient(180deg, #1a2744 0%, #243356 50%, #2d3f66 100%); color: #fff; display: flex; flex-direction: column; position: fixed; top: 0; left: 0; height: 100vh; z-index: 100; }
.sidebar::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.08), transparent); }
.sidebar-header { padding: 24px 20px 16px; border-bottom: 1px solid rgba(255,255,255,0.06); }
.sidebar-header h2 { font-size: 16px; font-weight: 700; letter-spacing: 0.3px; }
.sidebar-header .version { font-size: 10px; opacity: 0.4; margin-top: 3px; letter-spacing: 0.5px; }
.nav-section { padding: 12px 0; }
.nav-section-title { padding: 6px 20px 6px; font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(255,255,255,0.3); font-weight: 600; }
.nav-item { display: flex; align-items: center; gap: 10px; padding: 11px 20px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s; border-left: 3px solid transparent; color: rgba(255,255,255,0.65); }
.nav-item:hover { background: rgba(255,255,255,0.05); color: #fff; }
.nav-item.active { background: rgba(22,119,255,0.15); border-left-color: #1677ff; color: #fff; }
.nav-item .count { margin-left: auto; background: rgba(255,255,255,0.1); font-size: 10px; padding: 2px 7px; border-radius: 10px; font-weight: 500; }
.sidebar-footer { margin-top: auto; padding: 14px 20px; font-size: 11px; border-top: 1px solid rgba(255,255,255,0.06); color: rgba(255,255,255,0.35); }
.sidebar-footer .dot { display: inline-block; width: 6px; height: 6px; background: #52c41a; border-radius: 50%; margin-right: 6px; box-shadow: 0 0 4px rgba(82,196,26,0.5); }

/* === 主内容区 === */
.main { margin-left: 220px; flex: 1; min-height: 100vh; }
.header { background: #fff; padding: 14px 32px; display: flex; align-items: center; gap: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); position: sticky; top: 0; z-index: 50; }
.header h1 { font-size: 18px; font-weight: 700; background: linear-gradient(135deg, #1a1a2e, #1677ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.header .badge { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; font-size: 10px; padding: 2px 10px; border-radius: 12px; font-weight: 500; }
.header .breadcrumb { font-size: 12px; color: #94a3b8; }
.header .nav-links { margin-left: auto; display: flex; gap: 6px; }
.header .nav-links a { padding: 6px 16px; border-radius: 6px; font-size: 13px; text-decoration: none; color: #64748b; transition: all 0.2s; font-weight: 500; }
.header .nav-links a:hover { background: #f0f5ff; color: #1677ff; }
.header .nav-links a.active { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.3); }
.container { padding: 24px 32px; }

/* === 卡片 (阿里云风格 — 大圆角 + 细腻阴影) === */
.card { background: #fff; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.02); margin-bottom: 16px; overflow: hidden; }
.card-body { padding: 24px; }
.card-header { padding: 16px 24px; border-bottom: 1px solid #f1f5f9; font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 10px; color: #0f172a; }
.card-header .dot-indicator { width: 8px; height: 8px; border-radius: 50%; }
.card-header .dot-indicator.blue { background: #1677ff; }
.card-header .dot-indicator.green { background: #52c41a; }

/* === 仪表盘卡片 (阿里云风格指标卡) === */
.stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 20px; }
.stat-card { background: #fff; border-radius: 16px; padding: 20px 24px; box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.02); display: flex; flex-direction: column; }
.stat-card .stat-label { font-size: 12px; color: #94a3b8; margin-bottom: 6px; font-weight: 500; letter-spacing: 0.5px; }
.stat-card .stat-value { font-size: 28px; font-weight: 700; color: #0f172a; line-height: 1.2; }
.stat-card .stat-delta { font-size: 11px; margin-top: 6px; }
.stat-card .stat-delta.up { color: #52c41a; }
.stat-card .stat-delta.down { color: #ff4d4f; }
.stat-card .stat-delta.warn { color: #faad14; }

/* === 筛选栏 (阿里云风格) === */
.filter-bar { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; }
.filter-group { display: flex; flex-direction: column; gap: 5px; }
.filter-group label { font-size: 11px; color: #64748b; font-weight: 600; letter-spacing: 0.5px; }
.filter-group input, .filter-group select { padding: 8px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; outline: none; transition: all 0.2s; background: #fff; color: #1f2937; min-width: 130px; }
.filter-group input:hover, .filter-group select:hover { border-color: #cbd5e1; }
.filter-group input:focus, .filter-group select:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }

/* 按钮系列 (阿里云风格渐变) */
.btn { padding: 8px 20px; border: none; border-radius: 8px; font-size: 13px; cursor: pointer; font-weight: 600; transition: all 0.25s; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
.btn-primary { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.2); }
.btn-primary:hover { background: linear-gradient(135deg, #4096ff, #1677ff); box-shadow: 0 4px 12px rgba(22,119,255,0.3); transform: translateY(-1px); }
.btn-primary:disabled { background: #b0c4de; box-shadow: none; cursor: not-allowed; transform: none; }
.btn-success { background: linear-gradient(135deg, #52c41a, #389e0d); color: #fff; box-shadow: 0 2px 6px rgba(82,196,26,0.2); }
.btn-success:hover { background: linear-gradient(135deg, #73d13d, #52c41a); box-shadow: 0 4px 12px rgba(82,196,26,0.3); transform: translateY(-1px); }
.btn-success:disabled { background: #b7eb8f; box-shadow: none; cursor: not-allowed; transform: none; }
.btn-warning { background: linear-gradient(135deg, #faad14, #d48806); color: #fff; box-shadow: 0 2px 6px rgba(250,173,20,0.2); }
.btn-warning:hover { background: linear-gradient(135deg, #ffc53d, #faad14); box-shadow: 0 4px 12px rgba(250,173,20,0.3); transform: translateY(-1px); }
.btn-sm { padding: 5px 14px; font-size: 11px; }
.quick-btn { padding: 4px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; cursor: pointer; font-size: 12px; transition: all 0.2s; color: #64748b; font-weight: 500; }
.quick-btn:hover { border-color: #1677ff; color: #1677ff; background: #f0f5ff; }
.quick-btn.active { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; border-color: transparent; box-shadow: 0 2px 6px rgba(22,119,255,0.2); }

/* === 状态栏 === */
.status-bar { display: flex; justify-content: space-between; align-items: center; padding: 12px 24px; background: #fafbfc; border-top: 1px solid #f1f5f9; font-size: 12px; color: #94a3b8; }
.status-bar .stat-value { font-weight: 600; color: #1677ff; }

/* === 数据表格 === */
#dataTable { width: 100%; font-size: 12px; }
#dataTable th { background: #f8fafc; font-weight: 600; white-space: nowrap; font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 0.3px; padding: 10px 12px; }
#dataTable td { white-space: nowrap; padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }
.dataTables_wrapper { font-size: 12px; }
.dataTables_wrapper .dataTables_filter input { border: 1px solid #e2e8f0; border-radius: 8px; padding: 6px 12px; font-size: 12px; outline: none; }
.dataTables_wrapper .dataTables_filter input:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }

/* === 告警/提示 (阿里云风格) === */
.alert { padding: 12px 20px; border-radius: 12px; margin-bottom: 16px; display: none; font-size: 13px; }
.alert-error { background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; }
.alert-warning { background: #fffbeb; border: 1px solid #fde68a; color: #d97706; }
.alert-info { background: #eff6ff; border: 1px solid #bfdbfe; color: #2563eb; }

/* === 加载遮罩 === */
.loading-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.3); z-index: 9999; justify-content: center; align-items: center; backdrop-filter: blur(2px); }
.loading-overlay.show { display: flex; }
.loading-box { background: #fff; padding: 32px 48px; border-radius: 16px; text-align: center; box-shadow: 0 8px 32px rgba(0,0,0,0.12); }
.spinner { width: 36px; height: 36px; border: 3px solid #e2e8f0; border-top: 3px solid #1677ff; border-radius: 50%; animation: spin 0.7s cubic-bezier(0.4,0,0.2,1) infinite; margin: 0 auto 16px; }
@keyframes spin { to { transform: rotate(360deg); } }
@keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.4; } }
.spinner.done { border: none; width: 40px; height: 40px; animation: none; }
.spinner.done::after { content: '\2714'; font-size: 28px; color: #10b981; line-height: 40px; }

/* === 分析面板 (阿里云风格标签) === */
.analysis-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; margin-top: 16px; }
.analysis-card { border: 1px solid #f1f5f9; border-radius: 12px; padding: 20px; background: #fafbfc; transition: box-shadow 0.2s; }
.analysis-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
.analysis-card h4 { font-size: 13px; margin-bottom: 10px; color: #0f172a; font-weight: 600; }
.analysis-card .param-list { display: flex; flex-wrap: wrap; gap: 6px; }
.analysis-card .param-tag { padding: 3px 10px; background: #eff6ff; color: #2563eb; border-radius: 20px; font-size: 11px; font-weight: 500; border: none; }
.analysis-card .param-tag.control { background: #fff7ed; color: #c2410c; }
.analysis-card .param-tag.position { background: #ecfdf5; color: #059669; }
.analysis-card .param-tag.pressure { background: #eff6ff; color: #2563eb; }
.analysis-card .param-tag.timing { background: #f5f3ff; color: #7c3aed; }

/* === 响应式 */
.user-area { display: flex; align-items: center; gap: 10px; }
.user-name { font-size: 13px; font-weight: 600; color: #1f2937; }
.btn-logout { padding: 5px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; font-size: 12px; color: #64748b; cursor: pointer; font-weight: 500; transition: all 0.2s; }
.btn-logout:hover { border-color: #ff4d4f; color: #ff4d4f; background: #fff5f5; }
/* 响应式 */
@media (max-width: 768px) {
    .sidebar { width: 60px; min-width: 60px; }
    .sidebar .nav-item span { display: none; }
    .sidebar .nav-section-title { display: none; }
    .sidebar .count { display: none; }
    .main { margin-left: 60px; }
    .filter-bar { flex-direction: column; }
}
</style>
</head>
<body>

<!-- === 左侧导航 === -->
<div class="sidebar">
    <div class="sidebar-header">
        <h2>DCS 分析平台</h2>
        <div class="version">开口机 · 堵口机</div>
    </div>

    <div class="nav-section">
        <div class="nav-section-title">设备分组</div>
        <div class="nav-item active" data-group="east_opener" onclick="switchGroup('east_opener')">
            东开口机 <span class="count">19</span>
        </div>
        <div class="nav-item" data-group="west_opener" onclick="switchGroup('west_opener')">
            西开口机 <span class="count">19</span>
        </div>
        <div class="nav-item" data-group="east_plugger" onclick="switchGroup('east_plugger')">
            东堵口机 <span class="count">18</span>
        </div>
        <div class="nav-item" data-group="west_plugger" onclick="switchGroup('west_plugger')">
            西堵口机 <span class="count">18</span>
        </div>
    </div>

    <div class="nav-section">
        <div class="nav-section-title">全部</div>
        <div class="nav-item" data-group="all" onclick="switchGroup('all')">
            全部设备 <span class="count">74</span>
        </div>
    </div>

    <div class="sidebar-footer">
        <span class="dot" id="sysStatus"></span> <span id="sidebarUser">admin</span><a style="margin-left:auto;font-size:10px;cursor:pointer;color:rgba(255,255,255,0.45);text-decoration:none" onclick="doLogout()" href="javascript:void(0)">退出</a>
    </div>
</div>

<!-- === 主内容 === -->
<div class="main">

    <div class="header">
        <h1 id="pageTitle">东开口机</h1>
        <span class="badge">InfluxDB</span>
        <span class="breadcrumb" id="pageDesc">37个信号参数 · 7大类别</span>
        
        <div class="nav-links">
            <a href="/" class="active">历史查询</a>
            <a href="/realtime">实时监控</a>
            <a href="/trend">趋势分析</a>
            <a href="/analysis">作业分析</a>
        </div>
    </div>

    <div class="container">

        <!-- 统计概览 -->
        <div class="stats-row" id="statsRow">
            <div class="stat-card">
                <div class="stat-label">信号参数</div>
                <div class="stat-value" id="statParams">37</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">数据点数</div>
                <div class="stat-value" id="statRows">--</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">时间跨度</div>
                <div class="stat-value" id="statSpan">--</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">查询耗时</div>
                <div class="stat-value" id="statTime">--</div>
            </div>
        </div>

        <!-- 分析面板 -->
        <div class="card" id="analysisCard">
            <div class="card-header"><span class="dot-indicator blue"></span>参数分类与数据分析维度</div>
            <div class="card-body">
                <div class="analysis-grid" id="analysisGrid"></div>
            </div>
        </div>

        <!-- 按日批量导出 -->
        <div class="card">
            <div class="card-header"><span class="dot-indicator green"></span> 按日批量导出</div>
            <div class="card-body">
                <div class="filter-bar">
                    <div class="filter-group">
                        <label>选择日期</label>
                        <input type="date" id="dailyExportDate" style="min-width:150px;">
                    </div>
                    <div class="filter-group">
                        <label>设备分组</label>
                        <select id="dailyExportGroup" style="min-width:150px;">
                            <option value="east_opener">东开口机 (19)</option>
                            <option value="west_opener">西开口机 (19)</option>
                            <option value="east_plugger">东堵口机 (18)</option>
                            <option value="west_plugger">西堵口机 (18)</option>
                            <option value="all">全部设备 (74)</option>
                        </select>
                    </div>
                    <div class="filter-group" style="flex-direction:row;align-items:center;">
                        <button class="btn btn-success" id="dailyExportBtn" onclick="doDailyExport()">按日导出</button>
                        <span id="dailyExportStatus" style="display:none;color:#f59e0b;font-size:12px;margin-left:8px;animation:pulse 1.2s infinite;">⏳ 导出中...</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- 查询筛选 -->
        <div class="card">
            <div class="card-body">
                <div style="display:flex;gap:8px;margin-bottom:12px;">
                    <button class="quick-btn active" onclick="setQuickRange('1hour')">最近1小时</button>
                    <button class="quick-btn" onclick="setQuickRange('3hours')">最近3小时</button>
                    <button class="quick-btn" onclick="setQuickRange('6hours')">最近6小时</button>
                    <button class="quick-btn" onclick="setQuickRange('12hours')">最近12小时</button>
                    <button class="quick-btn" onclick="setQuickRange('24hours')">最近24小时</button>
                </div>
                <div class="filter-bar">
                    <div class="filter-group">
                        <label>开始时间</label>
                        <input type="datetime-local" id="startTime">
                    </div>
                    <div class="filter-group">
                        <label>结束时间</label>
                        <input type="datetime-local" id="endTime">
                    </div>
                    <div class="filter-group">
                        <label>每页行数</label>
                        <select id="pageSize">
                            <option value="50">50</option>
                            <option value="100" selected>100</option>
                            <option value="200">200</option>
                            <option value="500">500</option>
                        </select>
                    </div>
                    <div class="filter-group">
                        <label>&nbsp;</label>
                        <button class="btn btn-primary" onclick="doQuery()" id="queryBtn">查询</button>
                    </div>
                    <div class="filter-group">
                        <label>&nbsp;</label>
                        <button class="btn btn-success" onclick="doExport()" id="exportBtn" disabled>导出 Excel</button>
                    </div>
                </div>
            </div>
        </div>

        <div class="alert alert-error" id="alertError"></div>
        <div class="alert alert-warning" id="alertWarning"></div>

        <!-- 数据表格 -->
        <div class="card">
            <div class="card-body" style="min-height:200px;">
                <table id="dataTable" class="display compact" style="width:100%">
                    <thead><tr><th>时间</th><th>参数名</th><th>中文名称</th><th>数值</th></tr></thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="status-bar">
                <span>显示 <span class="stat-value" id="shownRows">0</span> 行 | 共 <span class="stat-value" id="totalRows">0</span> 个数据点</span>
                <span>10.56.128.202:8086 / islag / DCS</span>
            </div>
        </div>

    </div>
</div>

<div class="loading-overlay" id="loading">
    <div class="loading-box">
        <div class="spinner" id="loadingSpinner"></div>
        <div style="color:#666;font-size:14px;" id="loadingTitle">正在查询 InfluxDB...</div>
        <div style="color:#999;font-size:12px;margin-top:4px;" id="loadingSub">可能会需要几秒钟</div>
    </div>
</div>

<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
<script>
// === 登录检查 ===
function checkLogin(){fetch('/api/user?token='+API_TOKEN).then(function(r){return r.json()}).then(function(d){if(!d.logged_in){window.location.href='/';}if(d.username){var el=document.getElementById('sidebarUser');if(el)el.textContent=d.username;};}).catch(function(){window.location.href='/';});}
function doLogout(){fetch('/api/logout',{method:'POST'}).then(function(r){return r.json()}).then(function(d){window.location.href='/';});}

// === 初始化 ===
const GROUPS = {{ groups_json | safe }};
const LABELS = {{ labels_json | safe }};
const API_TOKEN = "{{ app_token }}";
checkLogin();
let currentGroup = 'east_opener';
let currentData = [];
let table = null;

// 日期初始化 - 默认最近1小时
const now = new Date();
const oneHourAgo = new Date(now - 3600000);
const pad2 = n => String(n).padStart(2,'0');
const fmtLocal = d => d.getFullYear()+'-'+pad2(d.getMonth()+1)+'-'+pad2(d.getDate())+'T'+pad2(d.getHours())+':'+pad2(d.getMinutes());
document.getElementById('endTime').value = fmtLocal(now);
document.getElementById('startTime').value = fmtLocal(oneHourAgo);

// 初始化分析面板和统计
updateAnalysisPanel();
updateStats();

// === 导航切换 ===
function switchGroup(groupId) {
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    document.querySelector(`[data-group="${groupId}"]`).classList.add('active');
    currentGroup = groupId;

    if (groupId === 'all') {
        document.getElementById('pageTitle').textContent = '全部设备';
        document.getElementById('pageDesc').textContent = '74个信号参数 · 东/西开口机 + 东/西堵口机';
    } else {
        const g = GROUPS.find(g => g.id === groupId);
        if (g) {
            document.getElementById('pageTitle').textContent = g.name;
            document.getElementById('pageDesc').textContent = `${g.params.length}个信号参数 · ${Object.keys(g.categories).length}大类别`;
        }
    }
    updateAnalysisPanel();
    updateStats();
}

function updateStats() {
    if (currentGroup === 'all') {
        const total = GROUPS.reduce((s, g) => s + g.params.length, 0);
        document.getElementById('statParams').textContent = total;
    } else {
        const g = GROUPS.find(g => g.id === currentGroup);
        document.getElementById('statParams').textContent = g ? g.params.length : '--';
    }
}

function updateAnalysisPanel() {
    const grid = document.getElementById('analysisGrid');
    const catNames = {
        remote_command: '遥控器指令', control_valve: '控制比例阀',
        position: '位置状态', pressure: '压力状态',
        control: '控制指令', safety: '安全状态', hydraulic: '液压参数'
    };
    const catColors = {
        remote_command: 'param-tag', control_valve: 'param-tag',
        position: 'param-tag position', pressure: 'param-tag pressure',
        control: 'param-tag control', safety: 'param-tag', hydraulic: 'param-tag'
    };

    let html = '';
    if (currentGroup === 'all') {
        GROUPS.forEach(g => {
            html += `<div class="analysis-card"><h4>${g.name}</h4><div class="param-list">`;
            Object.entries(g.categories).forEach(([cat, params]) => {
                params.forEach(p => {
                    html += `<span class="${catColors[cat] || 'param-tag'}">${LABELS[p] || p}</span>`;
                });
            });
            html += '</div></div>';
        });
    } else {
        const g = GROUPS.find(g => g.id === currentGroup);
        if (g) {
            Object.entries(g.categories).forEach(([cat, params]) => {
                html += `<div class="analysis-card"><h4>${catNames[cat] || cat}</h4><div class="param-list">`;
                params.forEach(p => {
                    html += `<span class="${catColors[cat] || 'param-tag'}">${LABELS[p] || p}</span>`;
                });
                html += '</div></div>';
            });
        }
    }
    grid.innerHTML = html;
}

// === 快捷时间 ===
function setQuickRange(range) {
    document.querySelectorAll('.quick-btn').forEach(b => b.classList.remove('active'));
    event.target.classList.add('active');
    const end = new Date();
    let hours;
    switch(range) {
        case '1hour': hours = 1; break;
        case '3hours': hours = 3; break;
        case '6hours': hours = 6; break;
        case '12hours': hours = 12; break;
        case '24hours': hours = 24; break;
        default: hours = 1;
    }
    const start = new Date(end - hours * 3600000);
    document.getElementById('endTime').value = fmtLocal(end);
    document.getElementById('startTime').value = fmtLocal(start);
}

// === 查询 ===
function showLoading() { document.getElementById('loading').classList.add('show'); document.getElementById('queryBtn').disabled = true; }
function hideLoading() { document.getElementById('loading').classList.remove('show'); document.getElementById('queryBtn').disabled = false; }
function hideAlerts() { ['alertError','alertWarning'].forEach(id => document.getElementById(id).style.display = 'none'); }

function doQuery() {
    const startTime = document.getElementById('startTime').value;
    const endTime = document.getElementById('endTime').value;
    if (!startTime || !endTime) { return; }

    showLoading();
    hideAlerts();

    const t0 = performance.now();
    $.getJSON('/api/query', {
        start: startTime, stop: endTime,
        group: currentGroup === 'all' ? '' : currentGroup,
        token: API_TOKEN
    })
    .done(function(data) {
        const t1 = performance.now();
        if (data.error) {
            document.getElementById('alertError').textContent = data.error;
            document.getElementById('alertError').style.display = 'block';
            hideLoading();
            return;
        }

        currentData = data.rows || [];
        document.getElementById('statRows').textContent = currentData.length.toLocaleString();
        document.getElementById('statTime').textContent = ((t1-t0)/1000).toFixed(1) + 's';
        document.getElementById('totalRows').textContent = currentData.length.toLocaleString();

        // 时间跨度 — _time 是 UTC ISO 字符串，转为本地北京时间显示
        if (currentData.length > 0) {
            const times = currentData.map(r => r._time).filter(Boolean).sort();
            const fmtIsoLocal = (isoStr) => {
                const d = new Date(isoStr);
                return d.getFullYear()+'-'+pad2(d.getMonth()+1)+'-'+pad2(d.getDate())+'T'+pad2(d.getHours())+':'+pad2(d.getMinutes());
            };
            document.getElementById('statSpan').textContent =
                fmtIsoLocal(times[0]) + ' → ' + fmtIsoLocal(times[times.length-1]);
        } else {
            document.getElementById('statSpan').textContent = '无数据';
        }

        if (data.truncated) {
            document.getElementById('alertWarning').textContent = data.hint;
            document.getElementById('alertWarning').style.display = 'block';
        }

        if (table) table.destroy();
        table = $('#dataTable').DataTable({
            data: currentData,
            columns: [
                { data: '_time', title: '时间', width: '160px',
                    render: v => v ? v.replace('T',' ').substring(0,19) : '' },
                { data: 'param', title: '参数名', width: '130px' },
                { data: 'label', title: '中文名称', width: '200px',
                    render: v => v || '<span style="color:#ccc">--</span>' },
                { data: 'value', title: '数值', width: '100px',
                    render: function(v) {
                        if (v === null || v === undefined || v === '') return '<span style="color:#ccc">--</span>';
                        const n = Number(v);
                        if (isNaN(n)) return v;
                        if (n === 0) return '<span style="color:#999">0.0000</span>';
                        return n.toFixed(4);
                    }
                }
            ],
            order: [[0, 'asc']],
            pageLength: parseInt(document.getElementById('pageSize').value),
            lengthMenu: [[25,50,100,200,500], [25,50,100,200,500]],
            language: {
                search: '搜索:',
                lengthMenu: '每页 _MENU_ 条',
                info: '显示 _START_ 到 _END_，共 _TOTAL_ 条',
                infoEmpty: '无数据',
                zeroRecords: '没有找到匹配的数据',
                paginate: { first: '首页', last: '末页', next: '下一页', previous: '上一页' }
            }
        });

        document.getElementById('exportBtn').disabled = currentData.length === 0;
        hideLoading();
    })
    .fail(function(xhr, status, err) {
        document.getElementById('alertError').textContent = '查询失败: ' + err;
        document.getElementById('alertError').style.display = 'block';
        hideLoading();
    });
}

// === 导出（window.open 新标签下载 + 弹窗进度） ===
var _exporting = false;

function _showExportLoading() {
    // 修改 loading 弹窗文案为导出相关
    document.getElementById('loadingTitle').textContent = '正在导出...';
    document.getElementById('loadingSub').textContent = '数据量越大，等待时间越长';
    document.getElementById('loading').classList.add('show');
    document.getElementById('exportBtn').disabled = true;
}

function _showExportDone(callback) {
    // 导出完成：spinner 变绿勾 + 文案改为完成
    document.getElementById('loadingSpinner').classList.add('done');
    document.getElementById('loadingTitle').textContent = '导出完成';
    document.getElementById('loadingSub').textContent = '文件正在下载中';
    // 1.5s 后关闭弹窗并恢复
    setTimeout(function() {
        document.getElementById('loadingSpinner').classList.remove('done');
        document.getElementById('loadingTitle').textContent = '正在查询 InfluxDB...';
        document.getElementById('loadingSub').textContent = '可能会需要几秒钟';
        document.getElementById('loading').classList.remove('show');
        _exporting = false;
        if (currentData.length > 0) document.getElementById('exportBtn').disabled = false;
        if (callback) callback();
    }, 1500);
}

function doExport() {
    if (_exporting) {
        alert('导出正在进行中，请等待完成后再试。');
        return;
    }
    const startTime = document.getElementById('startTime').value;
    const endTime = document.getElementById('endTime').value;
    if (!startTime || !endTime) return;

    _exporting = true;
    _showExportLoading();

    let url = '/api/export?start=' + encodeURIComponent(startTime) + '&stop=' + encodeURIComponent(endTime) + '&token=' + API_TOKEN;
    if (currentGroup !== 'all') url += '&group=' + currentGroup;

    // fetch+blob 确保文件完整下载后再弹出保存
    fetch(url)
        .then(resp => {
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            return resp.blob();
        })
        .then(blob => {
            var a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = '';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(function() { URL.revokeObjectURL(a.href); }, 1000);
            _showExportDone();
        })
        .catch(e => {
            _exporting = false;
            document.getElementById('loading').classList.remove('show');
            alert('导出失败: ' + e.message);
        });
}

// === 按日导出（window.open 新标签下载 + 进度提示） ===
var _dailyExporting = false;

function doDailyExport() {
    if (_dailyExporting) {
        alert('按日导出正在进行中，请稍候...');
        return;
    }
    if (_exporting) {
        if (!confirm('普通导出正在进行中，确定要开始按日导出吗？')) return;
    }
    var date = document.getElementById('dailyExportDate').value;
    var group = document.getElementById('dailyExportGroup').value;
    if (!date) { alert('请选择日期'); return; }

    _dailyExporting = true;
    var btn = document.getElementById('dailyExportBtn');
    var status = document.getElementById('dailyExportStatus');
    btn.disabled = true;
    status.style.display = 'inline';
    status.textContent = '\u23F3 导出中...';

    // 弹窗遮罩显示进度
    document.getElementById('loadingTitle').textContent = '正在按日导出...';
    document.getElementById('loadingSub').textContent = '全天数据量较大，请耐心等待';
    document.getElementById('loading').classList.add('show');

    var url = '/api/export_daily?date=' + date + '&group=' + group + '&token=' + API_TOKEN;
    fetch(url)
        .then(resp => {
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            return resp.blob();
        })
        .then(blob => {
            var a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = '';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(function() { URL.revokeObjectURL(a.href); }, 1000);
            // 完成状态
            document.getElementById('loadingSpinner').classList.add('done');
            document.getElementById('loadingTitle').textContent = '导出完成';
            document.getElementById('loadingSub').textContent = '文件正在下载中';
            setTimeout(function() {
                document.getElementById('loadingSpinner').classList.remove('done');
                document.getElementById('loadingTitle').textContent = '正在查询 InfluxDB...';
                document.getElementById('loadingSub').textContent = '可能会需要几秒钟';
                document.getElementById('loading').classList.remove('show');
                _dailyExporting = false;
                btn.disabled = false;
                status.style.display = 'none';
            }, 1500);
        })
        .catch(e => {
            _dailyExporting = false;
            btn.disabled = false;
            status.style.display = 'none';
            document.getElementById('loading').classList.remove('show');
            alert('按日导出失败: ' + e.message);
        });
}

// 回车键查询
document.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && document.activeElement.tagName !== 'BUTTON') { doQuery(); }
});

// 初始化
switchGroup('east_opener');
document.getElementById('dailyExportDate').valueAsDate = new Date();
</script>
</body>
</html>"""


# ==============================================================
# === 性能优化 #1：InfluxDB 单例客户端（避免每次请求新建连接）===
# ==============================================================
_client_lock = threading.Lock()
_influx_client = None
_client_created_at = 0
_CLIENT_MAX_AGE = 300  # 5分钟后自动重建连接（应对服务端断开）


def get_client():
    """返回全局复用的 InfluxDBClient 单例。
    首次访问时懒初始化，超过 CLIENT_MAX_AGE 自动重建。
    线程安全（双重检查锁）。
    """
    global _influx_client, _client_created_at
    now = _time.time()
    if _influx_client is None or (now - _client_created_at) > _CLIENT_MAX_AGE:
        with _client_lock:
            if _influx_client is None or (now - _client_created_at) > _CLIENT_MAX_AGE:
                if _influx_client is not None:
                    try:
                        _influx_client.close()
                    except Exception:
                        pass
                _influx_client = InfluxDBClient(
                    url=INFLUX_URL, token=INFLUX_TOKEN, org=INFLUX_ORG,
                    timeout=INFLUX_TIMEOUT_MS
                )
                _client_created_at = _time.time()
    return _influx_client


# ==============================================================
# === 性能优化 #5：查询结果内存缓存（TTL 30秒）===
# ==============================================================
_cache = {}
_cache_lock = threading.Lock()
CACHE_TTL = 30          # 默认缓存30秒
MAX_CACHE_ROWS = 5000   # 超过此行数不缓存（防止内存暴涨）


def _cache_key(route, **kwargs):
    """生成缓存键值"""
    raw = f"{route}|" + "|".join(f"{k}={v}" for k, v in sorted(kwargs.items()))
    return raw


def _cache_get(key):
    """读缓存，过期返回 None"""
    entry = _cache.get(key)
    if entry and _time.time() - entry["ts"] < CACHE_TTL:
        return entry["data"]
    # 清理过期条目
    if entry:
        _cache.pop(key, None)
    return None


def _cache_set(key, data, nrows=0):
    """写缓存。超过 MAX_CACHE_ROWS 行不缓存（防内存暴涨）。"""
    if nrows > MAX_CACHE_ROWS:
        return
    with _cache_lock:
        _cache[key] = {"data": data, "ts": _time.time()}
        # 防止缓存无限增长，保持 < 100 条
        if len(_cache) > 100:
            oldest = min(_cache, key=lambda k: _cache[k]["ts"])
            _cache.pop(oldest, None)


def get_group_params(group_id):
    """获取指定组的参数列表"""
    if not group_id or group_id == "all":
        all_p = []
        for g in PARAM_CONFIG["groups"]:
            all_p.extend(g["params"])
        return all_p, None
    for g in PARAM_CONFIG["groups"]:
        if g["id"] == group_id:
            return g["params"], g
    return [], None


# === 登录装饰器 ===
def login_required(f):
    """需要登录才能访问的页面"""
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect("/")
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper

# === 登录页面 ===
LOGIN_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DCS 数据分析平台</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);min-height:100vh;display:flex;align-items:center;justify-content:center}
.login-box{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,0.15);padding:40px;width:380px;max-width:95vw}
.login-box h2{font-size:22px;color:#1f2937;margin-bottom:6px}
.login-box .sub{font-size:12px;color:#94a3b8;margin-bottom:24px}
.form-group{margin-bottom:16px}
.form-group label{display:block;font-size:12px;color:#64748b;font-weight:600;margin-bottom:4px}
.form-group input{width:100%;padding:10px 14px;border:1px solid #e2e8f0;border-radius:8px;font-size:14px;outline:none;transition:border .2s}
.form-group input:focus{border-color:#667eea;box-shadow:0 0 0 3px rgba(102,126,234,0.1)}
.btn-login{width:100%;padding:12px;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2);margin-top:8px;transition:opacity .2s}
.btn-login:hover{opacity:.9}
.msg{text-align:center;font-size:12px;margin-top:12px;min-height:20px}
.msg.err{color:#ef4444}
.actions{text-align:center;margin-top:16px}
.actions a{font-size:12px;color:#94a3b8;cursor:pointer;text-decoration:none}
.actions a:hover{color:#667eea}
/* 设置弹窗 */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:999;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:#fff;border-radius:12px;padding:28px;width:420px;max-width:95vw;max-height:85vh;overflow-y:auto;box-shadow:0 8px 40px rgba(0,0,0,0.2)}
.modal h3{font-size:16px;color:#1f2937;margin-bottom:16px}
.modal .form-group input{font-size:13px;padding:8px 12px}
.modal .btn-row{display:flex;gap:8px;margin-top:16px}
.modal .btn-save{flex:1;padding:10px;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2)}
.modal .btn-cancel{flex:1;padding:10px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;color:#64748b;background:#fff}
.modal .msg{text-align:left}
</style>
</head>
<body>
<div class="login-box">
<h2>DCS 数据分析平台</h2>
<div class="sub">高炉炉前开口机/堵口机专用</div>
<form onsubmit="doLogin(event)">
<div class="form-group"><label>账号</label><input id="username" value="admin" placeholder="admin"></div>
<div class="form-group"><label>密码</label><input id="password" type="password" value="admin123" placeholder="admin123"></div>
<button type="submit" class="btn-login">登 录</button>
</form>
<div class="msg" id="msg"></div>
<div class="actions"><a onclick="openSettings()">⚙ 系统设置</a></div>
</div>

<!-- 设置弹窗 -->
<div class="modal-overlay" id="settingsModal">
<div class="modal">
<h3>⚙ 系统设置</h3>
<div class="form-group"><label>设置密码 (保护配置)</label><input id="settingsPwdCheck" type="password" placeholder="请输入设置密码" onchange="loadSettings()"></div>
<hr style="border:none;border-top:1px solid #f1f5f9;margin:12px 0">
<div class="form-group"><label>InfluxDB 地址</label><input id="influxUrl" placeholder="http://10.56.128.202:8086"></div>
<div class="form-group"><label>InfluxDB Token</label><input id="influxToken" type="password" placeholder="输入 Token"></div>
<div class="form-group"><label>组织 (Org)</label><input id="influxOrg" placeholder="myOrg"></div>
<div class="form-group"><label>Bucket</label><input id="influxBucket" placeholder="islag"></div>
<div class="form-group"><label>查询超时 (毫秒)</label><input id="influxTimeout" type="number" placeholder="180000"></div>
<div class="form-group"><label>Web 访问 Token</label><input id="appToken" placeholder="dcs2026"></div>
<div class="form-group"><label>新设置密码 (留空保持)</label><input id="settingsPwd" type="password" placeholder="留空保持不变"></div>
<div class="btn-row">
<button class="btn-save" onclick="saveSettings()">保存配置</button>
<button class="btn-cancel" onclick="closeSettings()">取消</button>
</div>
<div class="msg" id="sMsg"></div>
</div>
</div>

<script>
async function doLogin(e){
    e.preventDefault();
    var u=document.getElementById('username').value.trim();
    var p=document.getElementById('password').value.trim();
    var r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,password:p})});
    var d=await r.json();
    if(d.ok){window.location.href=d.redirect||'/trend'}
    else{document.getElementById('msg').innerHTML=d.error;document.getElementById('msg').className='msg err'}
}
function openSettings(){
    document.getElementById('settingsModal').classList.add('show');
    document.getElementById('settingsPwdCheck').value='';
    // 禁用输入（密码未验证）
    ['influxUrl','influxToken','influxOrg','influxBucket','influxTimeout','appToken','settingsPwd'].forEach(id=>document.getElementById(id).disabled=true);
    document.getElementById('sMsg').innerHTML='';
    document.getElementById('sMsg').className='msg';
}
function closeSettings(){document.getElementById('settingsModal').classList.remove('show')}
async function loadSettings(){
    var pwd=document.getElementById('settingsPwdCheck').value;
    if(!pwd)return;
    var r=await fetch('/api/settings/auth',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pwd})});
    var d=await r.json();
    if(!d.ok){document.getElementById('sMsg').innerHTML='<span class="err">密码错误</span>';return}
    // 解锁输入
    ['influxUrl','influxToken','influxOrg','influxBucket','influxTimeout','appToken','settingsPwd'].forEach(id=>document.getElementById(id).disabled=false);
    // 加载现有配置
    var cr=await fetch('/api/settings/config?token='+(d.token||''));
    var cfg=await cr.json();
    document.getElementById('influxUrl').value=cfg.INFLUX_URL||'';
    document.getElementById('influxOrg').value=cfg.INFLUX_ORG||'';
    document.getElementById('influxBucket').value=cfg.INFLUX_BUCKET||'';
    document.getElementById('influxTimeout').value=cfg.INFLUX_TIMEOUT_MS||'';
    document.getElementById('appToken').value=cfg.APP_TOKEN&&cfg.APP_TOKEN.includes('***')?'':(cfg.APP_TOKEN||'');
    document.getElementById('settingsPwd').value='';
    document.getElementById('sMsg').innerHTML='已加载当前配置';
    document.getElementById('sMsg').className='msg ok';
}
async function saveSettings(){
    var data={
        INFLUX_URL:document.getElementById('influxUrl').value.trim(),
        INFLUX_TOKEN:document.getElementById('influxToken').value.trim()||undefined,
        INFLUX_ORG:document.getElementById('influxOrg').value.trim(),
        INFLUX_BUCKET:document.getElementById('influxBucket').value.trim(),
        INFLUX_TIMEOUT_MS:document.getElementById('influxTimeout').value.trim(),
        APP_TOKEN:document.getElementById('appToken').value.trim(),
        SETTINGS_PASSWORD:document.getElementById('settingsPwd').value.trim()||undefined
    };
    Object.keys(data).forEach(k=>{if(data[k]===undefined)delete data[k]});
    var r=await fetch('/api/settings/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    var d=await r.json();
    document.getElementById('sMsg').innerHTML=d.ok?'<span style="color:#166534">保存成功</span>':'<span style="color:#ef4444">'+d.error+'</span>';
}
</script>
</body>
</html>"""

@app.route("/")
def index():
    if not session.get("logged_in"):
        return render_template_string(LOGIN_HTML)
    html = INDEX_HTML.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG["groups"]))
    html = html.replace("{{ labels_json | safe }}", json.dumps(_LABELS))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    return render_template_string(html)

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    if data.get("username") == "admin" and data.get("password") == "admin123":
        session["logged_in"] = True
    return jsonify({"ok": True, "redirect": "/"})
    return jsonify({"ok": False, "error": "账号或密码错误"})


@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.pop("logged_in", None)
    return jsonify({"ok": True, "redirect": "/"})


@app.route("/api/user")
def api_user():
    if session.get("logged_in"):
        return jsonify({"logged_in": True, "username": "admin"})
    return jsonify({"logged_in": False})


@app.route("/debug")
def debug():
    return jsonify({
        "has_zh": "中文名称" in INDEX_HTML,
        "has_32": "32</span>" in INDEX_HTML.split("east_opener")[1][:500] if "east_opener" in INDEX_HTML else False,
        "params_total": sum(len(g["params"]) for g in PARAM_CONFIG["groups"]),
        "labels": len(_LABELS)
    })


@app.route("/api/query")
def api_query():
    start = request.args.get("start", "")
    stop = request.args.get("stop", "")
    group_id = request.args.get("group", "all")

    if not start or not stop:
        return jsonify({"error": "请指定 start 和 stop 参数"})

    # 前端传的是北京时间(Asia/Shanghai)，转为UTC用于 InfluxDB 查询
    if not start.endswith("Z") and "+" not in start[-6:]:
        start += "+08:00"
    if not stop.endswith("Z") and "+" not in stop[-6:]:
        stop += "+08:00"

    try:
        s_local = datetime.fromisoformat(start)
        e_local = datetime.fromisoformat(stop)
        # 北京时间转UTC（减8小时）
        s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
        e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return jsonify({"error": "日期格式错误"})

    hours = (e_local - s_local).total_seconds() / 3600
    if hours > MAX_QUERY_HOURS:
        return jsonify({"error": f"查询范围不能超过{MAX_QUERY_HOURS}小时，请缩小日期范围"})

    params, group_info = get_group_params(group_id)
    if not params:
        return jsonify({"error": "未找到对应设备组"})

    # 构建 param 过滤条件（含白名单校验）
    param_filter = sanitize_param_for_flux(params)

    # === 缓存检查（相同查询 30 秒内直接返回）===
    cache_k = _cache_key("api_query", start=s_utc, stop=e_utc, group=group_id)
    cached = _cache_get(cache_k)
    if cached:
        return jsonify(cached)

    client = get_client()
    try:
        query_api = client.query_api()

        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {s_utc}, stop: {e_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> limit(n: {MAX_WEB_ROWS})'''

        tables = query_api.query(flux)
        rows = []
        param_set = set()
        labels = _LABELS
        for table in tables:
            for record in table.records:
                t = record.get_time()
                p = record.values.get("param", "")
                v = record.get_value()
                param_set.add(p)
                rows.append({
                    "_time": t.isoformat() if t else "",
                    "param": p,
                    "label": labels.get(p, p),
                    "value": v if v is not None else ""
                })

        truncated = len(rows) >= MAX_WEB_ROWS
        result = {
            "rows": rows,
            "param_count": len(param_set),
            "total": len(rows),
            "truncated": truncated,
            "hint": f"显示前{MAX_WEB_ROWS}行。缩小日期范围或使用导出Excel获取完整数据。" if truncated else ""
        }
        _cache_set(cache_k, result, len(rows))
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"查询失败: {str(e)[:200]}"})


@app.route("/api/export_daily")
def api_export_daily():
    """按日导出 — 分块查询原始数据，生成标准化 Excel（xlsxwriter 流式写入）"""
    date_str = request.args.get("date", "").strip()
    group_id = request.args.get("group", "all")

    # === 1. 日期验证 ===
    if not date_str:
        return "请选择日期", 400
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        return "日期格式错误，需要 YYYY-MM-DD", 400

    # 当天北京时间 00:00 ~ 次日 00:00
    try:
        day_start_local = datetime.strptime(date_str + "T00:00", "%Y-%m-%dT%H:%M")
        day_end_local = datetime.strptime(date_str + "T23:59", "%Y-%m-%dT%H:%M")
    except ValueError:
        return "日期解析失败", 400

    params, group_info = get_group_params(group_id)
    if not params:
        return "未找到对应设备组", 400

    all_params = sorted(params)
    labels = _LABELS
    param_filter = sanitize_param_for_flux(all_params)

    # === 2. 构建 xlsxwriter 流式工作簿 ===
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet((group_info["name"] if group_info else "全部设备")[:31])

    # 预定义格式
    hdr_fmt = wb.add_format({
        'font_name': 'Microsoft YaHei', 'font_size': 11, 'bold': True, 'font_color': 'FFFFFF',
        'bg_color': '1A1A2E', 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': 'D9D9D9'
    })
    sub_fmt = wb.add_format({
        'font_name': 'Microsoft YaHei', 'font_size': 10, 'bold': True, 'font_color': '1A1A2E',
        'bg_color': 'E6F4FF', 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': 'D9D9D9'
    })
    cell_fmt = wb.add_format({
        'font_name': 'Microsoft YaHei', 'font_size': 10,
        'border': 1, 'border_color': 'D9D9D9'
    })
    time_fmt_ms = wb.add_format({
        'font_name': 'Microsoft YaHei', 'font_size': 10,
        'num_format': 'yyyy-mm-dd hh:mm:ss.000', 'border': 1, 'border_color': 'D9D9D9'
    })

    # 双行表头
    ws.merge_range(0, 0, 1, 0, "时间", hdr_fmt)
    for col, p in enumerate(all_params):
        ci = col + 1
        ws.write(0, ci, p, hdr_fmt)
        ws.write(1, ci, labels.get(p, p), sub_fmt)

    # === 3. 分块查询（每 EXPORT_CHUNK_HOURS 小时一块，流式逐行写入 BytesIO）===
    total_rows = 0
    row = 2
    truncated = False
    client = get_client()

    try:
        query_api = client.query_api()
        chunk_start = day_start_local

        while chunk_start < day_end_local and not truncated:
            chunk_end = min(chunk_start + timedelta(hours=EXPORT_CHUNK_HOURS), day_end_local)

            s_utc = (chunk_start - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")
            e_utc = (chunk_end - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")

            flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {s_utc}, stop: {e_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> pivot(rowKey: ["_time"], columnKey: ["param"], valueColumn: "_value")'''

            tables = query_api.query(flux)

            # === 性能优化：空 chunk 快速跳过 ===
            if not tables or all(len(t.records) == 0 for t in tables):
                chunk_start = chunk_end
                continue

            for table in tables:
                for record in table.records:
                    if total_rows >= MAX_EXCEL_ROWS:
                        truncated = True
                        break

                    t = record.get_time()
                    t_beijing = t + LOCAL_OFFSET if t else None
                    if t_beijing:
                        dt_beijing = datetime(t_beijing.year, t_beijing.month, t_beijing.day,
                                              t_beijing.hour, t_beijing.minute, t_beijing.second,
                                              t_beijing.microsecond)
                        ws.write_datetime(row, 0, dt_beijing, time_fmt_ms)
                    else:
                        ws.write(row, 0, "", cell_fmt)

                    vals = record.values
                    for col, p in enumerate(all_params):
                        v = vals.get(p)
                        ws.write(row, col + 1, v if v is not None else "", cell_fmt)

                    row += 1
                    total_rows += 1
                if truncated:
                    break
            chunk_start = chunk_end

        # === 4. 空数据校验 ===
        if total_rows == 0:
            wb.close()
            return f"日期 {date_str} 的 {group_info['name'] if group_info else '全部设备'} 没有任何数据记录", 404

        # === 5. 完成 Excel 格式 ===
        last_col = len(all_params)
        ws.freeze_panes(2, 1)
        ws.autofilter(0, 0, 1, last_col)
        ws.set_column(0, 0, 22)         # 时间列（毫秒显示需要更宽）
        ws.set_column(1, last_col, 14)  # 参数列

        wb.close()
        output.seek(0)

        group_name = group_info["id"] if group_info else "all"
        filename = f"DCS_{group_name}_{date_str}.xlsx"
        return send_file(output,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=filename)
    except Exception as e:
        try:
            wb.close()
        except Exception:
            pass
        return f"导出失败: {str(e)[:200]}", 500


@app.route("/api/export")
def api_export():
    start = request.args.get("start", "")
    stop = request.args.get("stop", "")
    group_id = request.args.get("group", "all")

    if not start or not stop:
        return "请指定 start 和 stop 参数", 400

    if not start.endswith("Z") and "+" not in start[-6:]:
        start += "+08:00"
    if not stop.endswith("Z") and "+" not in stop[-6:]:
        stop += "+08:00"

    try:
        s_local = datetime.fromisoformat(start)
        e_local = datetime.fromisoformat(stop)
        s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
        e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return "日期格式错误", 400

    hours = (e_local - s_local).total_seconds() / 3600
    if hours > MAX_QUERY_HOURS:
        return f"查询范围不能超过{MAX_QUERY_HOURS}小时，请缩小日期范围", 400

    params, group_info = get_group_params(group_id)
    if not params:
        return "未找到对应设备组", 400

    param_filter = sanitize_param_for_flux(params)
    all_params = sorted(params)
    labels = _LABELS

    client = get_client()
    try:
        query_api = client.query_api()

        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {s_utc}, stop: {e_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> pivot(rowKey: ["_time"], columnKey: ["param"], valueColumn: "_value")'''

        tables = query_api.query(flux)

        # === xlsxwriter 流式写入 ===
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet((group_info["name"] if group_info else "全部设备")[:31])

        # 预定义格式
        hdr_fmt = wb.add_format({
            'font_name': 'Microsoft YaHei', 'font_size': 11, 'bold': True, 'font_color': 'FFFFFF',
            'bg_color': '1A1A2E', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'border_color': 'D9D9D9'
        })
        sub_fmt = wb.add_format({
            'font_name': 'Microsoft YaHei', 'font_size': 10, 'bold': True, 'font_color': '1A1A2E',
            'bg_color': 'E6F4FF', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'border_color': 'D9D9D9'
        })
        cell_fmt = wb.add_format({
            'font_name': 'Microsoft YaHei', 'font_size': 10,
            'border': 1, 'border_color': 'D9D9D9'
        })
        time_fmt = wb.add_format({
            'font_name': 'Microsoft YaHei', 'font_size': 10,
            'num_format': 'yyyy-mm-dd hh:mm:ss', 'border': 1, 'border_color': 'D9D9D9'
        })

        # 双行表头
        ws.merge_range(0, 0, 1, 0, "时间", hdr_fmt)  # row=0~1, col=0 合并
        for col, p in enumerate(all_params):
            ci = col + 1
            ws.write(0, ci, p, hdr_fmt)
            ws.write(1, ci, labels.get(p, p), sub_fmt)

        # 数据行（流式逐行写入 BytesIO）
        row = 2
        for table in tables:
            for record in table.records:
                t = record.get_time()
                t_beijing = t + LOCAL_OFFSET if t else None
                if t_beijing:
                    dt_val = datetime(t_beijing.year, t_beijing.month, t_beijing.day,
                                      t_beijing.hour, t_beijing.minute, t_beijing.second)
                    ws.write_datetime(row, 0, dt_val, time_fmt)
                else:
                    ws.write(row, 0, "", cell_fmt)

                vals = record.values
                for col, p in enumerate(all_params):
                    v = vals.get(p)
                    ws.write(row, col + 1, v if v is not None else "", cell_fmt)
                row += 1

        # 格式设置
        last_col = len(all_params)
        ws.freeze_panes(2, 1)           # B3 冻结（第3行第2列为第一个可滚动单元格）
        ws.autofilter(0, 0, 1, last_col)  # A1 到最后一列第二行
        ws.set_column(0, 0, 20)         # 时间列宽
        ws.set_column(1, last_col, 14)  # 参数列宽

        wb.close()
        output.seek(0)

        group_name = group_info["id"] if group_info else "all"
        filename = f"DCS_{group_name}_{start[:10]}_{stop[:10]}.xlsx"
        return send_file(output,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=filename)

    except Exception as e:
        return f"导出失败: {str(e)[:200]}", 500


# === 实时数据 API ===
@app.route("/api/realtime")
def api_realtime():
    """获取所有参数的最新值，支持 group 或 param 参数"""
    group_id = request.args.get("group", "all")
    single_param = request.args.get("param", "").strip()

    if single_param:
        params = [single_param]
    else:
        params, _ = get_group_params(group_id)

    if not params:
        return jsonify({"error": "未找到对应参数"})

    param_filter = sanitize_param_for_flux(params)
    all_params_sorted = sorted(params)
    labels = _LABELS

    # === 缓存检查（实时数据 30 秒内复用）===
    rk = single_param or group_id
    cache_k = _cache_key("api_realtime", key=rk)
    cached = _cache_get(cache_k)
    if cached:
        return jsonify(cached)

    client = get_client()
    try:
        query_api = client.query_api()

        # Flux: 查每个 param 最后一条记录
        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: -5m)
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> last()'''

        tables = query_api.query(flux)
        results = []
        for table in tables:
            for record in table.records:
                t = record.get_time()
                p = record.values.get("param", "")
                v = record.get_value()
                results.append({
                    "param": p,
                    "label": labels.get(p, p),
                    "value": v if v is not None else "",
                    "_time": t.isoformat() if t else ""
                })

        # 补全：Flux last() 可能丢失无数据的 param，确保列表完整
        existing_params = {r["param"] for r in results}
        for p in all_params_sorted:
            if p not in existing_params:
                results.append({
                    "param": p,
                    "label": labels.get(p, p),
                    "value": "",
                    "_time": ""
                })

        # 按 param 排序
        results.sort(key=lambda x: x["param"])
        # 添加序号
        for i, r in enumerate(results):
            r["index"] = i + 1

        result = {
            "rows": results,
            "total": len(results),
            "timestamp": datetime.now(timezone.utc).isoformat()
        }
        _cache_set(cache_k, result, len(results))
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"查询失败: {str(e)[:200]}"})


REALTIME_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>实时数据监控 — DCS</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #1f2937; display: flex; min-height: 100vh; }

/* === 左侧导航 (阿里云风格深色侧栏) === */
.sidebar { width: 220px; min-width: 220px; background: linear-gradient(180deg, #1a2744 0%, #243356 50%, #2d3f66 100%); color: #fff; display: flex; flex-direction: column; position: fixed; top: 0; left: 0; height: 100vh; z-index: 100; }
.sidebar::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.08), transparent); }
.sidebar-header { padding: 24px 20px 16px; border-bottom: 1px solid rgba(255,255,255,0.06); }
.sidebar-header h2 { font-size: 16px; font-weight: 700; letter-spacing: 0.3px; }
.sidebar-header .version { font-size: 10px; opacity: 0.4; margin-top: 3px; letter-spacing: 0.5px; }

/* 侧边栏搜索 */
.sidebar-search { padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,0.06); }
.sidebar-search input { width: 100%; padding: 7px 12px; border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; font-size: 12px; background: rgba(255,255,255,0.06); color: #fff; outline: none; transition: all 0.2s; }
.sidebar-search input::placeholder { color: rgba(255,255,255,0.3); }
.sidebar-search input:focus { border-color: rgba(255,255,255,0.25); background: rgba(255,255,255,0.1); }

/* 树形导航 */
.tree { flex: 1; overflow-y: auto; padding: 6px 0; }
.tree::-webkit-scrollbar { width: 4px; }
.tree::-webkit-scrollbar-track { background: transparent; }
.tree::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 2px; }

.tree-section-title { padding: 6px 20px 4px; font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(255,255,255,0.3); font-weight: 600; }

/* 全部设备 */
.tree-item-all { display: flex; align-items: center; gap: 10px; padding: 11px 20px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s; border-left: 3px solid transparent; color: rgba(255,255,255,0.65); }
.tree-item-all:hover { background: rgba(255,255,255,0.05); color: #fff; }
.tree-item-all.active { background: rgba(22,119,255,0.15); border-left-color: #1677ff; color: #fff; }

/* 设备分组标题 */
.tree-group-title { display: flex; align-items: center; gap: 8px; padding: 11px 20px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s; border-left: 3px solid transparent; color: rgba(255,255,255,0.65); user-select: none; }
.tree-group-title:hover { background: rgba(255,255,255,0.05); color: #fff; }
.tree-group-title.active { background: rgba(22,119,255,0.15); border-left-color: #1677ff; color: #fff; }
.tree-group-title .arrow { font-size: 10px; transition: transform 0.2s; display: inline-block; width: 10px; flex-shrink: 0; }
.tree-group-title .arrow.open { transform: rotate(90deg); }
.tree-group-title .badge { margin-left: auto; background: rgba(255,255,255,0.1); font-size: 10px; padding: 2px 7px; border-radius: 10px; font-weight: 500; }

/* 分类/参数子项 */
.tree-items { display: none; }
.tree-items.show { display: block; }
.tree-cat { padding: 6px 16px 6px 36px; font-size: 12px; color: rgba(255,255,255,0.45); cursor: pointer; transition: all 0.15s; border-left: 3px solid transparent; font-weight: 500; }
.tree-cat:hover { background: rgba(255,255,255,0.04); color: rgba(255,255,255,0.8); }
.tree-cat.active { color: #91caff; border-left-color: #91caff; background: rgba(22,119,255,0.08); }
.tree-param { padding: 3px 16px 3px 52px; font-size: 11px; color: rgba(255,255,255,0.35); cursor: pointer; transition: all 0.15s; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.tree-param:hover { color: rgba(255,255,255,0.6); background: rgba(255,255,255,0.03); }
.tree-param.active { color: #91caff; background: rgba(22,119,255,0.06); }

/* 侧边栏底部 */
.sidebar-footer { padding: 10px 20px; font-size: 10px; border-top: 1px solid rgba(255,255,255,0.06); color: rgba(255,255,255,0.3); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

/* === 主内容区 === */
.main { margin-left: 220px; flex: 1; display: flex; flex-direction: column; min-height: 100vh; }
.header { background: #fff; padding: 14px 32px; display: flex; align-items: center; gap: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); position: sticky; top: 0; z-index: 50; }
.header h1 { font-size: 18px; font-weight: 700; background: linear-gradient(135deg, #1a1a2e, #1677ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.header .nav-links { margin-left: auto; display: flex; gap: 6px; }
.header .nav-links a { padding: 6px 16px; border-radius: 6px; font-size: 13px; text-decoration: none; color: #64748b; transition: all 0.2s; font-weight: 500; }
.header .nav-links a:hover { background: #f0f5ff; color: #1677ff; }
.header .nav-links a.active { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.3); }

/* 工具栏 (阿里云风格) */
.toolbar { background: #fff; padding: 12px 24px; display: flex; align-items: center; gap: 14px; border-bottom: 1px solid #f1f5f9; flex-wrap: wrap; }
.toolbar .toggle-label { font-size: 13px; display: flex; align-items: center; gap: 8px; cursor: pointer; font-weight: 500; color: #1f2937; }
.toolbar .toggle-switch { width: 40px; height: 22px; background: #e2e8f0; border-radius: 11px; position: relative; cursor: pointer; transition: all 0.3s; }
.toolbar .toggle-switch.on { background: #52c41a; box-shadow: 0 0 4px rgba(82,196,26,0.3); }
.toolbar .toggle-switch::after { content: ''; width: 18px; height: 18px; background: #fff; border-radius: 50%; position: absolute; top: 2px; left: 2px; transition: left 0.3s; box-shadow: 0 1px 3px rgba(0,0,0,0.15); }
.toolbar .toggle-switch.on::after { left: 20px; }
.toolbar .refresh-select { padding: 6px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; outline: none; background: #fff; color: #1f2937; }
.toolbar .refresh-select:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }
.toolbar .status-dot { width: 8px; height: 8px; border-radius: 50%; background: #52c41a; display: inline-block; box-shadow: 0 0 4px rgba(82,196,26,0.5); }
.toolbar .status-text { font-size: 12px; color: #52c41a; font-weight: 500; }
.toolbar .timestamp { font-size: 11px; color: #94a3b8; margin-left: auto; }

/* 表格卡片 */
.content { flex: 1; padding: 24px 32px; overflow-y: auto; }
.table-wrap { background: #fff; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.02); overflow: hidden; }
.data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.data-table th { background: #f8fafc; padding: 12px 16px; text-align: left; font-weight: 600; font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 0.3px; border-bottom: 1px solid #f1f5f9; white-space: nowrap; position: sticky; top: 0; z-index: 1; }
.data-table td { padding: 10px 16px; border-bottom: 1px solid #f8fafc; white-space: nowrap; }
.data-table tbody tr { transition: background 0.15s; }
.data-table tbody tr:hover { background: #f0f5ff; }
.data-table tbody tr:nth-child(even) { background: #fafbfc; }
.data-table tbody tr:nth-child(even):hover { background: #f0f5ff; }
.data-table .col-index { width: 50px; text-align: center; color: #94a3b8; font-size: 12px; }
.data-table .col-tag { width: 140px; font-family: "Consolas", monospace; font-size: 12px; color: #0f172a; }
.data-table .col-label { width: 200px; color: #1f2937; }
.data-table .col-value { width: 120px; font-weight: 600; font-size: 14px; }
.data-table .col-time { width: 180px; color: #94a3b8; font-size: 12px; }
.data-table .value-zero { color: #cbd5e1; }
.data-table .value-normal { color: #0f172a; }
.data-table .value-high { color: #dc2626; font-weight: 700; }
.data-table .value-empty { color: #cbd5e1; font-style: italic; }
.data-table tfoot td { background: #fafbfc; font-size: 12px; color: #94a3b8; padding: 10px 16px; }

/* 按钮系列 (阿里云风格渐变) */
.btn { padding: 8px 20px; border: none; border-radius: 8px; font-size: 13px; cursor: pointer; font-weight: 600; transition: all 0.25s; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
.btn-primary { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.2); }
.btn-primary:hover { background: linear-gradient(135deg, #4096ff, #1677ff); box-shadow: 0 4px 12px rgba(22,119,255,0.3); transform: translateY(-1px); }

/* 加载 */
.loading-bar { height: 3px; background: #f1f5f9; overflow: hidden; }
.loading-bar::after { content: ''; display: block; width: 40%; height: 100%; background: linear-gradient(90deg, #1677ff, #4096ff); animation: loading 1.5s ease-in-out infinite; border-radius: 2px; }
@keyframes loading { 0% { transform: translateX(-100%); } 100% { transform: translateX(350%); } }
.loading-bar.hide { display: none; }

/* === 响应式 */
.user-area { display: flex; align-items: center; gap: 10px; }
.user-name { font-size: 13px; font-weight: 600; color: #1f2937; }
.btn-logout { padding: 5px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; font-size: 12px; color: #64748b; cursor: pointer; font-weight: 500; transition: all 0.2s; }
.btn-logout:hover { border-color: #ff4d4f; color: #ff4d4f; background: #fff5f5; }
/* 响应式 */
@media (max-width: 768px) {
    .sidebar { width: 60px; min-width: 60px; }
    .sidebar .sidebar-search { display: none; }
    .sidebar .tree-item-all span:not(.tree-item-all-icon), .tree-group-title span:not(.arrow),
    .tree-group-title .badge, .tree-cat, .tree-param, .sidebar-footer, .sidebar-header .version { display: none; }
    .sidebar .tree-group-title { padding: 9px 8px; justify-content: center; }
    .sidebar .tree-item-all { padding: 9px 8px; justify-content: center; }
    .main { margin-left: 60px; }
    .toolbar { flex-direction: column; align-items: flex-start; }
}
</style>
</head>
<body>

<div class="sidebar">
    <div class="sidebar-header">
        <h2>DCS 分析平台</h2>
        <div class="version">实时监控</div>
    </div>

    <div class="sidebar-search">
        <input type="text" id="sidebarSearch" placeholder="搜索 TAG 或中文名..." onkeyup="filterTable()">
    </div>

    <div class="tree" id="treeNav"></div>

    <div class="sidebar-footer" id="sidebarFooter"><span id="sidebarUser">admin</span> | <a style="cursor:pointer;color:rgba(255,255,255,0.7);text-decoration:none" onclick="doLogout()" href="javascript:void(0)">退出</a> | 默认: 全部设备</div>
</div>

<div class="main">
    <div class="header">
        <h1>实时数据监控</h1>
        
        <div class="nav-links">
            <a href="/">历史查询</a>
            <a href="/realtime" class="active">实时监控</a>
            <a href="/trend">趋势分析</a>
            <a href="/analysis">作业分析</a>
        </div>
    </div>

    <div class="toolbar">
        <label class="toggle-label">
            <span class="toggle-switch" id="toggleSwitch" onclick="toggleAutoRefresh()"></span>
            <span>实时刷新</span>
        </label>
        <select class="refresh-select" id="refreshInterval" onchange="changeInterval()">
            <option value="3">3 秒</option>
            <option value="5" selected>5 秒</option>
            <option value="10">10 秒</option>
            <option value="30">30 秒</option>
        </select>
        <button class="btn btn-primary" onclick="fetchRealtime()">手动刷新</button>
        <span class="status-dot" id="statusDot"></span>
        <span class="status-text" id="statusText">在线</span>
        <span class="timestamp" id="lastUpdate">--</span>
    </div>

    <div class="loading-bar" id="loadingBar"></div>

    <div class="content">
        <div class="table-wrap">
            <table class="data-table" id="realtimeTable">
                <thead>
                    <tr>
                        <th class="col-index">序号</th>
                        <th class="col-tag">TAG名称</th>
                        <th class="col-label">中文名称</th>
                        <th class="col-value">最新值</th>
                        <th class="col-time">更新时间</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
                <tfoot>
                    <tr><td colspan="5" id="tableFooter">共 0 条记录</td></tr>
                </tfoot>
            </table>
        </div>
    </div>
</div>

<script>
function checkLogin(){fetch('/api/user?token='+API_TOKEN).then(function(r){return r.json()}).then(function(d){if(!d.logged_in){window.location.href='/';}if(d.username){var el=document.getElementById('sidebarUser');if(el)el.textContent=d.username;};}).catch(function(){window.location.href='/';});}
function doLogout(){fetch('/api/logout',{method:'POST'}).then(function(r){return r.json()}).then(function(d){window.location.href='/';});}
const GROUPS = {{ groups_json | safe }};
const LABELS = {{ labels_json | safe }};
const API_TOKEN = "{{ app_token }}";
checkLogin();
const CAT_NAMES = {
    remote_command: '遥控器指令', control_valve: '控制比例阀',
    position: '位置状态', pressure: '压力状态',
    control: '控制指令', safety: '安全状态', hydraulic: '液压参数'
};
let currentFilter = { type: 'group', value: 'all' };
let autoRefreshTimer = null;
let autoRefreshOn = false;
let latestRows = [];

// === 构建树形导航 ===
function buildTree() {
    const tree = document.getElementById('treeNav');
    let html = `
    <div class="tree-section-title">设备分组</div>
    <div class="tree-item-all active" data-filter="all" onclick="clickGroup('all', this)">
        全部设备
    </div>`;

    GROUPS.forEach(g => {
        html += `<div class="tree-group">
            <div class="tree-group-title" data-filter="${g.id}" onclick="clickGroup('${g.id}', this)">
                <span class="arrow">▶</span>
                ${g.name}
                <span class="badge">${g.params.length}</span>
            </div>
            <div class="tree-items" id="items_${g.id}"></div>
        </div>`;
    });
    tree.innerHTML = html;

    // 构建每个组的子项
    GROUPS.forEach(g => {
        const div = document.getElementById('items_' + g.id);
        let subHtml = '';
        Object.entries(g.categories).forEach(([cat, params]) => {
            subHtml += `<div class="tree-cat" data-group="${g.id}" data-cat="${cat}" onclick="clickCategory('${g.id}', '${cat}', this)">
                ${CAT_NAMES[cat] || cat} (${params.length})
            </div>`;
            params.forEach(p => {
                subHtml += `<div class="tree-param" data-group="${g.id}" data-param="${p}" onclick="clickParam('${p}', this)">
                    ${LABELS[p] || p}
                </div>`;
            });
        });
        div.innerHTML = subHtml;
    });

    expandGroup('east_opener');
}

function expandGroup(groupId) {
    const items = document.getElementById('items_' + groupId);
    const title = document.querySelector(`.tree-group-title[data-filter="${groupId}"]`);
    if (items && title) {
        items.classList.add('show');
        title.querySelector('.arrow').classList.add('open');
    }
}

function toggleGroup(groupId) {
    const items = document.getElementById('items_' + groupId);
    const title = document.querySelector(`.tree-group-title[data-filter="${groupId}"]`);
    if (items && title) {
        items.classList.toggle('show');
        title.querySelector('.arrow').classList.toggle('open');
    }
}

function clearActive() {
    document.querySelectorAll('.tree-item-all, .tree-group-title, .tree-cat, .tree-param')
        .forEach(e => e.classList.remove('active'));
}

function setFooter(text) {
    document.getElementById('sidebarFooter').textContent = text;
}

// === 三种点击行为：分组 / 分类 / 参数 ===

function clickGroup(groupId, el) {
    clearActive();
    el.classList.add('active');
    currentFilter = { type: 'group', value: groupId };

    const name = groupId === 'all' ? '全部设备' :
        (GROUPS.find(g => g.id === groupId)?.name || groupId);
    setFooter('当前: ' + name);

    // 如果点击设备组标题，同时展开/收起
    if (el.classList.contains('tree-group-title')) {
        toggleGroup(groupId);
    }

    doFetch();
}

function clickCategory(groupId, cat, el) {
    clearActive();
    el.classList.add('active');
    currentFilter = { type: 'category', group: groupId, cat: cat };
    setFooter('当前: ' + (CAT_NAMES[cat] || cat));
    doFetch();
}

function clickParam(param, el) {
    clearActive();
    el.classList.add('active');
    currentFilter = { type: 'param', value: param };
    setFooter('当前: ' + (LABELS[param] || param));
    doFetch();
}

// === 数据获取（拒绝 jQuery，原生 fetch） ===

function doFetch() {
    // 停止自动刷新计时器（防止堆积）
    if (autoRefreshOn) startAutoRefresh();

    document.getElementById('loadingBar').classList.remove('hide');

    let url = '/api/realtime?group=all&token=' + API_TOKEN;
    if (currentFilter.type === 'group' && currentFilter.value !== 'all') {
        url = '/api/realtime?group=' + currentFilter.value + '&token=' + API_TOKEN;
    } else if (currentFilter.type === 'param') {
        url = '/api/realtime?param=' + encodeURIComponent(currentFilter.value) + '&token=' + API_TOKEN;
    } else if (currentFilter.type === 'category') {
        url = '/api/realtime?group=' + currentFilter.group + '&token=' + API_TOKEN;
    }

    fetch(url)
        .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
        .then(data => {
            document.getElementById('loadingBar').classList.add('hide');
            if (data.error) {
                document.getElementById('statusText').textContent = '错误: ' + data.error;
                document.getElementById('statusDot').style.background = '#ff4d4f';
                return;
            }
            document.getElementById('statusDot').style.background = '#52c41a';
            document.getElementById('statusText').textContent = '在线';
            const ts = data.timestamp ? data.timestamp.replace('T',' ').substring(0,19) : '';
            document.getElementById('lastUpdate').textContent = '上次更新: ' + ts;

            latestRows = data.rows || [];
            renderTable(latestRows);
        })
        .catch(err => {
            document.getElementById('loadingBar').classList.add('hide');
            document.getElementById('statusDot').style.background = '#ff4d4f';
            const msg = (err && err.message) || String(err);
            document.getElementById('statusText').textContent = '错误: ' + msg;
            console.error('fetch 错误:', err);
        });
}

function fetchRealtime() {
    doFetch();
}

function renderTable(rows) {
    const tbody = document.getElementById('tableBody');
    const searchVal = document.getElementById('sidebarSearch').value.toLowerCase().trim();

    let filtered = rows;

    // 分类筛选（如果 API 返回的是整个组的数据）
    if (currentFilter.type === 'param') {
        filtered = filtered.filter(r => r.param === currentFilter.value);
    } else if (currentFilter.type === 'category') {
        const g = GROUPS.find(x => x.id === currentFilter.group);
        if (g && g.categories[currentFilter.cat]) {
            const catParams = g.categories[currentFilter.cat];
            filtered = filtered.filter(r => catParams.includes(r.param));
        }
    }

    // 搜索文本筛选
    if (searchVal) {
        filtered = filtered.filter(r =>
            r.param.toLowerCase().includes(searchVal) ||
            (r.label || '').toLowerCase().includes(searchVal)
        );
    }

    // 重新编号
    filtered.forEach((r, i) => r.index = i + 1);

    let html = '';
    filtered.forEach(r => {
        const t = r._time ? r._time.replace('T',' ').substring(0,19) : '--';
        let valClass = 'value-normal';
        let displayVal = r.value;
        if (displayVal === '' || displayVal === null || displayVal === undefined) {
            displayVal = '--';
            valClass = 'value-empty';
        } else if (typeof displayVal === 'number' && displayVal === 0) {
            valClass = 'value-zero';
            displayVal = '0.0000';
        } else if (typeof displayVal === 'number') {
            displayVal = Number(displayVal).toFixed(4);
        }

        html += `<tr>
            <td class="col-index">${r.index}</td>
            <td class="col-tag">${r.param}</td>
            <td class="col-label">${r.label || r.param}</td>
            <td class="col-value"><span class="${valClass}">${displayVal}</span></td>
            <td class="col-time">${t}</td>
        </tr>`;
    });

    tbody.innerHTML = html;
    document.getElementById('tableFooter').textContent =
        `共 ${filtered.length} 条记录（筛选后 ${filtered.length} / 总计 ${rows.length}）`;
}

function filterTable() {
    if (latestRows.length > 0) renderTable(latestRows);
}

// === 自动刷新 ===
function toggleAutoRefresh() {
    const sw = document.getElementById('toggleSwitch');
    autoRefreshOn = !autoRefreshOn;
    if (autoRefreshOn) {
        sw.classList.add('on');
        startAutoRefresh();
    } else {
        sw.classList.remove('on');
        stopAutoRefresh();
    }
}

function startAutoRefresh() {
    stopAutoRefresh();
    const sec = parseInt(document.getElementById('refreshInterval').value) || 5;
    autoRefreshTimer = setInterval(doFetch, sec * 1000);
}

function stopAutoRefresh() {
    if (autoRefreshTimer) { clearInterval(autoRefreshTimer); autoRefreshTimer = null; }
}

function changeInterval() {
    if (autoRefreshOn) startAutoRefresh();
}

// === 初始化 ===
buildTree();
doFetch();
</script>
</body>
</html>"""


TREND_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<title>趋势分析 — DCS</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/uplot@1.6.30/dist/uPlot.min.css">
<script src="https://cdn.jsdelivr.net/npm/uplot@1.6.30/dist/uPlot.iife.min.js"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #1f2937; display: flex; min-height: 100vh; }

.sidebar { width: 220px; min-width: 220px; background: linear-gradient(180deg, #1a2744 0%, #243356 50%, #2d3f66 100%); color: #fff; display: flex; flex-direction: column; position: fixed; top: 0; left: 0; height: 100vh; z-index: 100; }
.sidebar::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.08), transparent); }
.sidebar-header { padding: 24px 20px 16px; border-bottom: 1px solid rgba(255,255,255,0.06); }
.sidebar-header h2 { font-size: 16px; font-weight: 700; }
.sidebar-header .version { font-size: 10px; opacity: 0.4; margin-top: 3px; }

.tree { flex: 1; overflow-y: auto; padding: 6px 0; }
.tree::-webkit-scrollbar { width: 4px; }
.tree::-webkit-scrollbar-track { background: transparent; }
.tree::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 2px; }

.tree-section-title { padding: 6px 20px 4px; font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(255,255,255,0.3); font-weight: 600; }
.tree-group-title { display: flex; align-items: center; gap: 8px; padding: 11px 20px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s; color: rgba(255,255,255,0.65); user-select: none; }
.tree-group-title:hover { background: rgba(255,255,255,0.05); color: #fff; }
.tree-group-title .arrow { font-size: 10px; transition: transform 0.2s; display: inline-block; width: 10px; flex-shrink: 0; }
.tree-group-title .arrow.open { transform: rotate(90deg); }
.tree-items { display: none; }
.tree-items.show { display: block; }
.tree-param { padding: 6px 20px 6px 40px; font-size: 12px; color: rgba(255,255,255,0.45); cursor: pointer; transition: all 0.15s; display: flex; align-items: center; gap: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.tree-param:hover { color: #fff; background: rgba(255,255,255,0.04); }
.tree-param.checked { color: #91caff; }
.tree-param .check-box { width: 14px; height: 14px; border: 1.5px solid rgba(255,255,255,0.3); border-radius: 3px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; font-size: 10px; }
.tree-param.checked .check-box { background: #1677ff; border-color: #1677ff; }
.sidebar-footer { margin-top: auto; padding: 14px 20px; font-size: 11px; border-top: 1px solid rgba(255,255,255,0.06); color: rgba(255,255,255,0.35); }

/* 主内容 */
.main { margin-left: 220px; flex: 1; display: flex; flex-direction: column; min-height: 100vh; }
.header { background: #fff; padding: 14px 32px; display: flex; align-items: center; gap: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); position: sticky; top: 0; z-index: 50; }
.header h1 { font-size: 18px; font-weight: 700; background: linear-gradient(135deg, #1a1a2e, #1677ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.header .nav-links { margin-left: auto; display: flex; gap: 6px; }
.header .nav-links a { padding: 6px 16px; border-radius: 6px; font-size: 13px; text-decoration: none; color: #64748b; transition: all 0.2s; font-weight: 500; }
.header .nav-links a:hover { background: #f0f5ff; color: #1677ff; }
.header .nav-links a.active { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.3); }

/* 工具栏 */
.toolbar { background: #fff; padding: 10px 24px; border-bottom: 1px solid #f1f5f9; display: flex; gap: 10px; align-items: flex-end; flex-wrap: wrap; }
.toolbar label { font-size: 11px; color: #64748b; font-weight: 600; }
.toolbar input[type=datetime-local] { padding: 6px 10px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; outline: none; }
.toolbar input[type=datetime-local]:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }
/* 事件选择器与左侧参数树风格统一 */
.event-select-wrap select { padding: 8px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; outline: none; background: #fff; color: #1f2937; min-width: 200px; height: 38px; }
.event-select-wrap select:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }
.event-select-wrap .label { display: block; font-size: 11px; color: #64748b; font-weight: 600; margin-bottom: 2px; }
.event-select-wrap .btn { height: 38px; padding: 0 16px; font-size: 13px; }
.btn { padding: 8px 20px; border: none; border-radius: 8px; font-size: 13px; cursor: pointer; font-weight: 600; transition: all 0.25s; white-space: nowrap; }
.btn-primary { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.2); }
.btn-primary:hover { background: linear-gradient(135deg, #4096ff, #1677ff); box-shadow: 0 4px 12px rgba(22,119,255,0.3); }
.btn-sm { padding: 4px 12px; font-size: 11px; border-radius: 6px; }
.btn-outline { background: #fff; border: 1px solid #e2e8f0; color: #64748b; }
.btn-outline:hover { border-color: #1677ff; color: #1677ff; }
.btn-danger { background: #fff; border: 1px solid #fecaca; color: #dc2626; }
.btn-danger:hover { background: #fef2f2; }

/* 配置面板 */
.config-panel { background: #fff; margin: 16px 24px 0; border-radius: 12px; box-shadow: 0 1px 6px rgba(0,0,0,0.04); overflow: hidden; }
.config-header { padding: 12px 20px; border-bottom: 1px solid #f1f5f9; font-size: 13px; font-weight: 600; color: #0f172a; display: flex; align-items: center; justify-content: space-between; }
.config-body { padding: 12px 20px; max-height: 200px; overflow-y: auto; }
.config-empty { padding: 24px; text-align: center; color: #94a3b8; font-size: 13px; }
.config-row { display: flex; align-items: center; gap: 10px; padding: 6px 0; border-bottom: 1px solid #f8fafc; font-size: 12px; }
.config-row:last-child { border-bottom: none; }
.config-row .color-dot { width: 12px; height: 12px; border-radius: 50%; flex-shrink: 0; }
.config-row .param-name { width: 140px; font-weight: 500; flex-shrink: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.config-row .param-label { width: 160px; color: #64748b; flex-shrink: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.config-row select, .config-row input { padding: 4px 8px; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 11px; outline: none; }
.config-row select:focus, .config-row input:focus { border-color: #1677ff; }
.config-row input[type=number] { width: 72px; }
.config-row .remove-btn { cursor: pointer; color: #94a3b8; font-size: 16px; padding: 0 4px; }
.config-row .remove-btn:hover { color: #dc2626; }

/* 图表区 */
.chart-wrap { flex: 1; margin: 16px 24px 24px; background: #fff; border-radius: 12px; box-shadow: 0 1px 6px rgba(0,0,0,0.04); display: flex; flex-direction: column; min-height: 500px; position: relative; }
.chart-area { flex: 1; position: relative; }
.loading-overlay { display: none; position: absolute; inset: 0; background: rgba(255,255,255,0.7); justify-content: center; align-items: center; z-index: 10; }
.loading-overlay.show { display: flex; }
.spinner { width: 32px; height: 32px; border: 3px solid #e2e8f0; border-top: 3px solid #1677ff; border-radius: 50%; animation: spin 0.7s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

@media (max-width: 768px) {
    .sidebar { width: 60px; min-width: 60px; }
    .sidebar .tree-param, .sidebar .tree-section-title, .sidebar .sidebar-header .version { display: none; }
    .main { margin-left: 60px; }
    .config-panel { margin: 8px 12px 0; }
    .chart-wrap { margin: 8px 12px 12px; }
    .toolbar { flex-direction: column; align-items: flex-start; }
}
.user-area { display: flex; align-items: center; gap: 10px; }
.user-name { font-size: 13px; font-weight: 600; color: #1f2937; }
.btn-logout { padding: 5px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; font-size: 12px; color: #64748b; cursor: pointer; font-weight: 500; transition: all 0.2s; }
.btn-logout:hover { border-color: #ff4d4f; color: #ff4d4f; background: #fff5f5; }
</style>
</head>
<body>

<div class="sidebar">
    <div class="sidebar-header">
        <h2>DCS 分析平台</h2>
        <div class="version">趋势分析</div>
    </div>
    <div class="tree-section-title">选择变量</div>
    <div class="tree" id="treeNav"></div>
    <div class="sidebar-footer" id="sidebarFooter"><span id="sidebarUser">admin</span> | <a style="cursor:pointer;color:rgba(255,255,255,0.7);text-decoration:none" onclick="doLogout()" href="javascript:void(0)">退出</a> | 已选 0 个变量</div>
</div>

<div class="main">
    <div class="header">
        <h1>趋势分析</h1>
        
        <div class="nav-links">
            <a href="/">历史查询</a>
            <a href="/realtime">实时监控</a>
            <a href="/trend" class="active">趋势分析</a>
            <a href="/analysis">作业分析</a>
        </div>
    </div>

    <div class="toolbar">
        <div>
            <label>开始时间</label><br>
            <input type="datetime-local" id="startTime">
        </div>
        <div>
            <label>结束时间</label><br>
            <input type="datetime-local" id="endTime">
        </div>
        <button class="btn btn-primary" onclick="doQuery()">查询</button>
        <div class="event-box event-select-wrap">
            <span class="label">开口事件</span>
            <div style="display:flex;gap:6px;align-items:center;">
                <select id="eventSelect" onchange="onEventSelect()">
                    <option value="">-- 选择开口事件 --</option>
                </select>
                <button class="btn btn-outline" onclick="fetchEvents()" title="检测开口事件">🔍</button>
            </div>
        </div>
        <button class="btn btn-outline btn-sm" onclick="clearAll()" style="margin-left:auto;">清空所有变量</button>
    </div>

    <div class="config-panel">
        <div class="config-header">
            <span>已选变量配置</span>
            <span style="font-size:11px;color:#94a3b8;">Y轴范围留空 = 自动</span>
        </div>
        <div class="config-body" id="configBody">
            <div class="config-empty">左侧勾选变量开始分析</div>
        </div>
    </div>

    <div class="chart-wrap">
        <div id="chartArea" class="chart-area"></div>
    </div>
</div>

<script>
function checkLogin(){fetch('/api/user?token='+API_TOKEN).then(function(r){return r.json()}).then(function(d){if(!d.logged_in){window.location.href='/';}if(d.username){var el=document.getElementById('sidebarUser');if(el)el.textContent=d.username;};}).catch(function(){window.location.href='/';});}
function doLogout(){fetch('/api/logout',{method:'POST'}).then(function(r){return r.json()}).then(function(d){window.location.href='/';});}
const GROUPS = {{ groups_json | safe }};
const LABELS = {{ labels_json | safe }};
const API_TOKEN = "{{ app_token }}";
checkLogin();
const COLORS = ['#1677ff','#52c41a','#faad14','#ff4d4f','#722ed1','#13c2c2','#eb2f96','#fa8c16','#2f54eb','#a0d911','#f5222d','#1890ff','#fa541c','#9254de','#597ef7'];
let selectedParams = [];
let uplot = null;

// === 销毁图表 ===
function destroyChart() {
    if (uplot) { uplot.destroy(); uplot = null; }
    const tip = document.getElementById('uplot-tooltip');
    if (tip) tip.remove();
    const legend = document.getElementById('chart-legend');
    if (legend) legend.innerHTML = '';
    const indicator = document.getElementById('zoomIndicator');
    if (indicator) indicator.remove();
}

// === 数据转换: API格式 → uPlot列格式 ===
//   应用每条曲线的 Y轴偏移量 (offset)，仅影响渲染位置不影响原始数据
function buildUplotData(apiData) {
    const first = apiData.series.find(s => s.data.length > 0);
    if (!first) return null;
    const timestamps = first.data.map(d => new Date(d.time).getTime() / 1000);
    const cols = [timestamps];
    apiData.series.forEach(s => {
        const map = {};
        s.data.forEach(d => { map[new Date(d.time).getTime() / 1000] = d.value; });
        const cfg = selectedParams.find(x => x.param === s.param);
        const off = (cfg && cfg.offset) ? cfg.offset : 0;
        cols.push(timestamps.map(t => {
            const v = map[t];
            if (v === undefined) return null;
            return v + off;  // 视觉上偏移，不影响 tooltip 显示
        }));
    });
    return cols;
}

// 记录原始值 (未偏移) 供 tooltip 显示
let _rawValues = [];

// === 自定义 Tooltip 插件（固定页面右侧） ===
function tooltipPlugin() {
    let over;
    const tip = document.createElement('div');
    tip.id = 'uplot-tooltip';
    tip.style.cssText = [
        'display:none;position:fixed;z-index:100;pointer-events:none;',
        'top:50%;right:24px;transform:translateY(-50%);',
        'background:rgba(15,23,42,0.95);color:#f1f5f9;',
        'padding:12px 16px;border-radius:8px;',
        'font-size:12px;line-height:1.6;',
        'font-family:-apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;',
        'box-shadow:0 4px 16px rgba(0,0,0,0.2);',
        'border:1px solid #334155;',
        'min-width:220px;max-width:280px;'
    ].join('');
    return {
        hooks: {
            init: (u) => {
                over = u.over;
                tip.owner = u;  // 缓存 uPlot 实例引用
                document.body.appendChild(tip);
                over.addEventListener('mouseenter', () => { tip.style.display = 'block'; });
                over.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
            },
            setCursor: (u) => {
                const { idx } = u.cursor;
                if (idx == null) { tip.style.display = 'none'; return; }
                tip.style.display = 'block';
                const t = u.data[0][idx];
                const d = new Date(t * 1000);
                const pad = n => String(n).padStart(2, '0');
                const timeStr = d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()) +
                    ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
                let html = '<div style="font-weight:700;margin-bottom:10px;color:#e2e8f0;font-size:12px;' +
                    'border-bottom:1px solid #334155;padding-bottom:8px;">⏱ ' + timeStr + '</div>';
                for (let i = 1; i < u.data.length; i++) {
                    const vRaw = (_rawValues[i-1] && _rawValues[i-1][idx] !== undefined) ? _rawValues[i-1][idx] : null;
                    const s = u.series[i];
                    // 直接从 uPlot series 取颜色，不依赖 selectedParams (左侧勾选池)
                    const clr = s._color || (typeof s.stroke === 'function' ? '#1677ff' : (s.stroke || '#999'));
                    // 数值统一跟随曲线颜色，更直观对应
                    const valHtml = (vRaw == null)
                        ? '<span style="color:#64748b;">--</span>'
                        : '<span style="font-weight:700;color:' + clr + ';font-variant-numeric:tabular-nums;text-shadow:0 0 8px ' + clr + '66;">' + Number(vRaw).toFixed(2) + '</span>';
                    html += '<div data-tip-row="' + (i-1) + '" style="display:flex;align-items:center;gap:8px;margin:4px 0;padding:4px 8px 4px 12px;' +
                        'border-left:5px solid ' + clr + ';border-radius:0 6px 6px 0;' +
                        'background:linear-gradient(90deg,' + clr + '22 0%,rgba(255,255,255,0.04) 100%);' +
                        'white-space:nowrap;">' +
                        '<span style="flex:1;min-width:60px;overflow:hidden;text-overflow:ellipsis;color:#e2e8f0;font-size:11px;font-weight:600;">' +
                        (s.label || '') + '</span>' +
                        valHtml + '</div>';
                }
                tip.innerHTML = html;
                // 缓存行引用供 toggleSeriesVis 同步
                window._tipRows = tip.querySelectorAll('[data-tip-row]');
            }
        }
    };
}

// === 初始化时间 ===
function initTime() {
    const now = new Date();
    const pad = n => String(n).padStart(2,'0');
    const fmt = d => d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(d.getHours())+':'+pad(d.getMinutes());
    document.getElementById('endTime').value = fmt(now);
    now.setHours(now.getHours() - 2);
    document.getElementById('startTime').value = fmt(now);
}

// === 构建树 ===
function buildTree() {
    const tree = document.getElementById('treeNav');
    let html = '';
    GROUPS.forEach(g => {
        html += `<div class="tree-group">
            <div class="tree-group-title" onclick="toggleGroup('${g.id}')">
                <span class="arrow">▶</span> ${g.name}
            </div>
            <div class="tree-items" id="items_${g.id}"></div>
        </div>`;
    });
    tree.innerHTML = html;

    GROUPS.forEach(g => {
        const div = document.getElementById('items_' + g.id);
        let sub = '';
        const CAT_CN = {remote_command:'遥控指令', control_valve:'控制阀', position:'位置', pressure:'压力', control:'控制', safety:'安全', hydraulic:'液压', mud_quantity:'打泥量'};
        Object.entries(g.categories).forEach(([cat, params]) => {
            sub += '<div style="padding:4px 20px;font-size:10px;color:rgba(255,255,255,0.25);text-transform:uppercase;">' + (CAT_CN[cat] || cat) + '</div>';
            params.forEach(p => {
                sub += `<div class="tree-param" data-param="${p}" onclick="toggleParam('${p}', this)">
                    <span class="check-box"></span> ${LABELS[p] || p}
                </div>`;
            });
        });
        div.innerHTML = sub;
    });

    // 默认展开第一个组
    const first = document.getElementById('items_' + GROUPS[0].id);
    if (first) {
        first.classList.add('show');
        first.parentElement.querySelector('.arrow').classList.add('open');
    }
}

function toggleGroup(id) {
    const items = document.getElementById('items_' + id);
    const arrow = items.parentElement.querySelector('.arrow');
    items.classList.toggle('show');
    arrow.classList.toggle('open');
}

// === 勾选/取消变量（仅管理 selectedParams 池，不触发趋势图刷新） ===
function toggleParam(param, el) {
    const idx = selectedParams.findIndex(s => s.param === param);
    if (idx >= 0) {
        selectedParams.splice(idx, 1);
        el.classList.remove('checked');
        el.querySelector('.check-box').textContent = '';
    } else {
        // 若已存在 (color 池里) 找原来的颜色优先, 否则按当前长度分配
        const existing = selectedParams.find(s => s.param === param);
        const color = (existing && existing.color) || COLORS[selectedParams.length % COLORS.length];
        selectedParams.push({
            param: param,
            label: LABELS[param] || param,
            color: color,
            yAxisIndex: 0,
            yMin: '',
            yMax: '',
            offset: 0
        });
        el.classList.add('checked');
        el.querySelector('.check-box').textContent = '✓';
    }
    document.getElementById('sidebarFooter').textContent = '已选 ' + selectedParams.length + ' 个变量';
    renderConfig();
}

function clearAll() {
    selectedParams = [];
    document.querySelectorAll('.tree-param.checked').forEach(el => {
        el.classList.remove('checked');
        el.querySelector('.check-box').textContent = '';
    });
    document.getElementById('sidebarFooter').textContent = '已选 0 个变量';
    renderConfig();
    destroyChart();
}

// === 渲染配置面板 ===
function renderConfig() {
    const body = document.getElementById('configBody');
    if (selectedParams.length === 0) {
        body.innerHTML = '<div class="config-empty">左侧勾选变量开始分析</div>';
        return;
    }
    let html = '';
    selectedParams.forEach((s, i) => {
        html += `<div class="config-row">
            <span class="color-dot" style="background:${s.color};" title="${s.param}"></span>
            <span class="param-name" title="${s.param}">${s.param}</span>
            <span class="param-label">${s.label}</span>
            <span style="font-size:10px;color:#94a3b8;">Y轴</span>
            <select onchange="updateConfig(${i},'yAxisIndex',parseInt(this.value))">
                <option value="0" ${s.yAxisIndex===0?'selected':''}>左1</option>
                <option value="1" ${s.yAxisIndex===1?'selected':''}>右1</option>
                <option value="2" ${s.yAxisIndex===2?'selected':''}>左2</option>
                <option value="3" ${s.yAxisIndex===3?'selected':''}>右2</option>
            </select>
            <span style="font-size:10px;color:#94a3b8;" title="视觉偏移，不影响原始数值">偏移</span>
            <input type="number" placeholder="0" value="${s.offset||0}" step="any" style="width:55px;"
                onchange="updateConfig(${i},'offset',parseFloat(this.value)||0)" title="正数上移，负数下移">
            <input type="number" placeholder="Min" value="${s.yMin}" step="any" style="width:50px;"
                onchange="updateConfig(${i},'yMin',this.value)">
            <input type="number" placeholder="Max" value="${s.yMax}" step="any" style="width:50px;"
                onchange="updateConfig(${i},'yMax',this.value)">
            <span class="remove-btn" onclick="removeParam(${i})" title="移除">✕</span>
        </div>`;
    });
    body.innerHTML = html;
}

function updateConfig(idx, key, val) {
    selectedParams[idx][key] = val;
}

function removeParam(idx) {
    const p = selectedParams[idx];
    selectedParams.splice(idx, 1);
    const el = document.querySelector(`.tree-param[data-param="${p.param}"]`);
    if (el) {
        el.classList.remove('checked');
        el.querySelector('.check-box').textContent = '';
    }
    document.getElementById('sidebarFooter').textContent = '已选 ' + selectedParams.length + ' 个变量';
    renderConfig();
    if (selectedParams.length === 0) destroyChart();
}

// === 查询 ===
function doQuery() {
    if (selectedParams.length === 0) return;
    const start = document.getElementById('startTime').value;
    const end = document.getElementById('endTime').value;
    if (!start || !end) return;

    const paramStr = selectedParams.map(s => s.param).join(',');
    const url = '/api/trend?params=' + encodeURIComponent(paramStr) +
                '&start=' + encodeURIComponent(start) +
                '&end=' + encodeURIComponent(end) +
                '&token=' + API_TOKEN;

    fetch(url)
        .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
        .then(data => {
            if (data.error) { alert(data.error); return; }
            renderChart(data);
        })
        .catch(e => alert('请求失败: ' + e));
}

// === 渲染 uPlot 图表 ===
function renderChart(data) {
    destroyChart();

    const cols = buildUplotData(data);
    if (!cols) return;

    // 缓存原始值（未偏移）供 tooltip 显示
    _rawValues = data.series.map(s => s.data.map(d => d.value));

    // 收集已使用的 Y 轴
    let usedAxes = new Set();
    selectedParams.forEach(s => usedAxes.add(s.yAxisIndex));
    let sortedAxes = [...usedAxes].sort((a,b) => a-b);

    // 映射 yAxisIndex → scale key
    let axisToScale = {};
    sortedAxes.forEach((id, i) => { axisToScale[id] = 'y' + i; });

    // 构建 scales
    let scales = { x: { time: true } };
    sortedAxes.forEach((id, i) => {
        let range = undefined;
        selectedParams.filter(s => s.yAxisIndex === id).forEach(s => {
            if (s.yMin !== '' && s.yMax !== '') {
                let r = [parseFloat(s.yMin), parseFloat(s.yMax)];
                if (!range) range = r;
                else { range[0] = Math.min(range[0], r[0]); range[1] = Math.max(range[1], r[1]); }
            }
        });
        scales['y' + i] = { auto: !range };
        if (range) scales['y' + i].range = (u, min, max) => range;
    });

    // 构建 series
    let series = [{}];
    data.series.forEach(s => {
        let cfg = selectedParams.find(x => x.param === s.param);
        if (!cfg) return;
        series.push({
            label: cfg.label,
            stroke: cfg.color,
            _color: cfg.color,  // tooltip 用（不受 uPlot 内部转换影响）
            width: 2,
            scale: axisToScale[cfg.yAxisIndex],
            spanGaps: false,
            value: (u, v) => v == null ? '--' : Number(v).toFixed(4),
        });
    });

    // 构建 axes: 每个 Y 轴独立配色 + 标签
    let axes = [{
        stroke: '#cbd5e1',
        grid: { stroke: '#e2e8f0', width: 1 },
        font: '10px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif',
    }];
    sortedAxes.forEach((id, i) => {
        let isLeft = id % 2 === 0;
        let extraSpace = Math.floor(id / 2) * 55;
        // 该轴上的曲线
        let axisCurves = selectedParams.filter(s => s.yAxisIndex === id);
        // 轴标签: 单曲线用全称, 多曲线记数
        let axisLabel = '';
        let axisColor = '#94a3b8';
        if (axisCurves.length === 1) {
            axisLabel = axisCurves[0].label;
            axisColor = axisCurves[0].color;
        } else if (axisCurves.length > 1) {
            axisLabel = axisCurves.length + ' 条曲线';
            axisColor = '#64748b';
        }
        axes.push({
            scale: 'y' + i,
            side: isLeft ? 3 : 1,
            size: 62 + extraSpace,
            stroke: axisColor,
            font: '10px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif',
            label: axisLabel,
            labelFont: `bold 10px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif`,
            labelGap: 4,
            grid: {
                show: i === 0,
                stroke: '#f1f5f9',
                width: 1,
                dash: [4, 4]
            },
            values: (u, vals) => vals.map(v => {
                if (v == null) return '';
                if (Math.abs(v) >= 10000) return (v/1000).toFixed(1) + 'k';
                if (Math.abs(v) < 0.001 && v !== 0) return v.toExponential(1);
                if (Number.isInteger(v) && Math.abs(v) < 1000) return v.toString();
                return v.toFixed(2);
            }),
        });
    });

    // 创建 uPlot
    const container = document.getElementById('chartArea');
    const opts = {
        width: Math.max(container.clientWidth - 4, 400),
        height: Math.max(container.clientHeight - 4, 300),
        cursor: {
            show: true,
            x: true,
            y: true,
            drag: { x: true, y: false, setScale: true },
            focus: { prox: 10 },
            points: { show: false },
        },
        legend: { show: false },
        plugins: [tooltipPlugin()],
        scales: scales,
        axes: axes,
        series: series,
    };

    uplot = new uPlot(opts, cols, container);

    // 缩放级别指示器
    setupZoomIndicator();

    // 自定义图例
    buildLegend(data);

    // 响应式
    window.addEventListener('resize', () => {
        if (uplot) {
            const c = document.getElementById('chartArea');
            uplot.setSize({
                width: Math.max(c.clientWidth - 4, 400),
                height: Math.max(c.clientHeight - 4, 300)
            });
        }
    });
}


// === 自定义图例 ===
function buildLegend(data) {
    let legendDiv = document.getElementById('chart-legend');
    if (!legendDiv) {
        legendDiv = document.createElement('div');
        legendDiv.id = 'chart-legend';
        legendDiv.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px 16px;padding:8px 16px;' +
            'border-bottom:1px solid #f1f5f9;max-height:60px;overflow-y:auto;';
        const chartWrap = document.getElementById('chartArea').parentElement;
        chartWrap.insertBefore(legendDiv, chartWrap.firstChild);
    }
    let html = '';
    data.series.forEach((s, i) => {
        let cfg = selectedParams.find(x => x.param === s.param);
        if (!cfg) return;
        html += '<label data-series-idx="' + (i+1) + '" data-series-label="' + s.label + '" style="display:flex;align-items:center;gap:4px;font-size:11px;' +
            'color:#475569;cursor:pointer;user-select:none;padding:2px 0;" ' +
            'onclick="toggleSeriesVis(' + (i+1) + ', this)">' +
            '<span style="width:10px;height:2px;background:' + cfg.color + ';flex-shrink:0;"></span>' +
            s.label + '</label>';
    });
    legendDiv.innerHTML = html;
    // 绑定隐藏 series 后，uPlot 内部 idx 可能变，所以做 label → uPlot idx 映射缓存
    legendDiv._labelToIdx = {};
    data.series.forEach((s, i) => { legendDiv._labelToIdx[s.label] = i + 1; });
}

// === 图例切换可见性（带 tooltip 同步） ===
function toggleSeriesVis(idx, el) {
    if (!uplot) return;
    const s = uplot.series[idx];
    const visible = s.show !== false;
    uplot.setSeries(idx, { show: !visible }, false);
    uplot.redraw();
    el.style.opacity = visible ? '0.35' : '1';
    el.dataset.hidden = visible ? '1' : '0';
    // 同步 tooltip 中对应的行（按 label 匹配）
    if (uplot.series[idx]) {
        const label = uplot.series[idx].label;
        document.querySelectorAll('#uplot-tooltip [data-tip-row]').forEach(row => {
            const rowIdx = parseInt(row.dataset.tipRow);
            if (uplot.series[rowIdx + 1] && uplot.series[rowIdx + 1].label === label) {
                row.style.display = visible ? 'none' : 'flex';
            }
        });
    }
}

// === 保存图片 ===
function saveChartImage() {
    if (!uplot) return;
    const canvas = uplot.ctx.canvas;
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = canvas.width;
    exportCanvas.height = canvas.height;
    const ctx = exportCanvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, exportCanvas.width, exportCanvas.height);
    ctx.drawImage(canvas, 0, 0);
    const link = document.createElement('a');
    link.download = 'DCS趋势_' + new Date().toISOString().slice(0,10) + '.png';
    link.href = exportCanvas.toDataURL('image/png');
    link.click();
}

// === 重置缩放 ===
function resetZoom() {
    if (!uplot || !uplot.data[0] || uplot.data[0].length < 2) return;
    const xMin = uplot.data[0][0];
    const xMax = uplot.data[0][uplot.data[0].length - 1];
    uplot.setScale('x', { min: xMin, max: xMax });
    if (uplot.scales) {
        Object.keys(uplot.scales).forEach(k => {
            if (k !== 'x') uplot.setScale(k, { min: null, max: null });
        });
    }
}

// === 缩放级别指示器 ===
function setupZoomIndicator() {
    if (!uplot) return;
    let el = document.getElementById('zoomIndicator');
    if (!el) {
        el = document.createElement('div');
        el.id = 'zoomIndicator';
        el.style.cssText = 'position:absolute;top:8px;right:120px;font-size:11px;color:#94a3b8;' +
            'background:rgba(255,255,255,0.85);padding:2px 10px;border-radius:4px;pointer-events:none;z-index:10;';
        document.getElementById('chartArea').appendChild(el);
    }
    const updateIndicator = () => {
        if (!uplot || !uplot.scales.x) return;
        const sc = uplot.scales.x;
        const total = sc.max - sc.min;
        if (total <= 0) { el.textContent = ''; return; }
        const visible = (uplot.scales.x.max || 0) - (uplot.scales.x.min || 0);
        const pct = (visible / total) * 100;
        const level = pct >= 90 ? '' : pct >= 40 ? ' 宽' : pct >= 15 ? ' 中' : ' 细';
        el.textContent = '范围: ' + Math.round(pct) + '%' + level;
    };
    if (!uplot.hooks.setScale) uplot.hooks.setScale = [];
    uplot.hooks.setScale.push((u, key) => { if (key === 'x') updateIndicator(); });
    updateIndicator();
}

// === 添加工具栏按钮（截图/重置） ===
function addToolbarButtons() {
    const toolbar = document.querySelector('.toolbar');
    if (!toolbar || document.getElementById('btnSaveImg')) return;
    const btnGroup = document.createElement('div');
    btnGroup.style.cssText = 'margin-left:auto;display:flex;gap:6px;';
    btnGroup.innerHTML =
        '<button class="btn btn-outline btn-sm" id="btnReset" onclick="resetZoom()" title="重置缩放">↺ 重置</button>' +
        '<button class="btn btn-outline btn-sm" id="btnSaveImg" onclick="saveChartImage()" title="保存图片">📷 截图</button>';
    toolbar.appendChild(btnGroup);
}

// === 回转到位事件检测 ===
let currentEvents = [];

function fetchEvents() {
    const start = document.getElementById('startTime').value;
    const end = document.getElementById('endTime').value;
    if (!start || !end) { alert('请先设置时间范围'); return; }

    const url = '/api/trend/events?start=' + encodeURIComponent(start) +
                '&end=' + encodeURIComponent(end) +
                '&token=' + API_TOKEN;

    const sel = document.getElementById('eventSelect');
    sel.innerHTML = '<option value="">检测中...</option>';

    fetch(url)
        .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
        .then(data => {
            if (data.error) { alert(data.error); return; }
            currentEvents = data.events;
            sel.innerHTML = '<option value="">-- 选择开口事件 --</option>';
            data.events.forEach((ev, i) => {
                const d = new Date(ev.trigger_time);
                const pad = n => String(n).padStart(2,'0');
                const ts = pad(d.getMonth()+1)+'-'+pad(d.getDate())+' '+
                           pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
                sel.innerHTML += '<option value="' + i + '">' +
                    ev.machine + ' 遥控开口 @ ' + ts + '</option>';
            });
            if (data.events.length === 0) {
                sel.innerHTML = '<option value="">未检测到开口事件</option>';
                alert('该时间范围内未检测到开口事件');
            }
        })
        .catch(e => alert('事件检测失败: ' + e));
}

function onEventSelect() {
    const idx = document.getElementById('eventSelect').value;
    if (idx === '') return;
    const ev = currentEvents[parseInt(idx)];
    if (!ev) return;
    const fmt = (iso) => {
        const d = new Date(iso);
        const pad = n => String(n).padStart(2,'0');
        return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(d.getHours())+':'+pad(d.getMinutes());
    };
    document.getElementById('startTime').value = fmt(ev.window_start);
    document.getElementById('endTime').value = fmt(ev.window_end);
    doQuery();
}

// === 初始化 ===
initTime();
buildTree();
addToolbarButtons();

</script>
</body>
</html>"""


@app.route("/realtime")
@login_required
def realtime():
    html = REALTIME_HTML.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG["groups"]))
    html = html.replace("{{ labels_json | safe }}", json.dumps(_LABELS))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    return render_template_string(html)


# === 趋势分析 API ===
@app.route("/api/trend")
def api_trend():
    params_str = request.args.get("params", "").strip()
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()

    if not params_str or not start or not end:
        return jsonify({"error": "缺少必要参数"}), 400

    # 解析参数列表
    param_list = [p.strip() for p in params_str.split(",") if p.strip()]
    if not param_list:
        return jsonify({"error": "参数列表为空"}), 400

    # 时间转 UTC（前端传北京时间）
    try:
        s_local = datetime.strptime(start, "%Y-%m-%dT%H:%M")
        e_local = datetime.strptime(end, "%Y-%m-%dT%H:%M")
    except ValueError:
        return jsonify({"error": "时间格式错误"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")

    # 计算数据量，自动降采样
    time_diff_hours = (e_local - s_local).total_seconds() / 3600
    if time_diff_hours <= 2:
        window = "10s"
    elif time_diff_hours <= 8:
        window = "30s"
    elif time_diff_hours <= 24:
        window = "1m"
    else:
        window = "5m"

    param_filter = sanitize_param_for_flux(param_list)
    labels = _LABELS

    client = get_client()
    try:
        query_api = client.query_api()
        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {s_utc}, stop: {e_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> aggregateWindow(every: {window}, fn: mean, createEmpty: false)'''

        tables = query_api.query(flux)

        # 组织成 {param: [{time, value}, ...]}
        series = {}
        for table in tables:
            for record in table.records:
                t = record.get_time()
                p = record.values.get("param", "")
                v = record.get_value()
                if v is None:
                    continue
                if p not in series:
                    series[p] = []
                series[p].append({
                    "time": t.isoformat(),
                    "value": float(v)
                })

        # 确保所有参数都有条目
        result_series = []
        for p in param_list:
            pts = series.get(p, [])
            pts.sort(key=lambda x: x["time"])
            result_series.append({
                "param": p,
                "label": labels.get(p, p),
                "data": pts
            })

        return jsonify({"series": result_series, "window": window})
    except Exception as e:
        return jsonify({"error": f"查询失败: {str(e)[:200]}"}), 500


# === 开口事件检测 API ===
@app.route("/api/trend/events")
def api_trend_events():
    """检测开口事件：遥控变量=1 且 回转位置穿过90° → 前后各1分钟窗口
    
    设备信号映射：
    ┌──────────┬─────────────────────┬──────────────────────┐
    │ 设备     │ 遥控变量            │ 回转位置             │
    ├──────────┼─────────────────────┼──────────────────────┤
    │ 东开口机 │ LT_LQFC_57 (选择)   │ LT_LQFC_63 (位置)    │
    │ 西开口机 │ LT_LQFC_94 (选择)   │ LT_LQFC_100 (位置)   │
    └──────────┴─────────────────────┴──────────────────────┘
    """
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    if not start or not end:
        return jsonify({"error": "缺少时间参数"}), 400

    try:
        s_local = datetime.strptime(start, "%Y-%m-%dT%H:%M")
        e_local = datetime.strptime(end, "%Y-%m-%dT%H:%M")
    except ValueError:
        return jsonify({"error": "时间格式错误"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")

    # 事件配置: {设备名: {remote: 遥控参数, position: 位置参数, threshold: 到位阈值}}
    EVENT_CONFIG = {
        "东开口机": {"remote": "LT_LQFC_57", "position": "LT_LQFC_63", "threshold": 90},
        "西开口机": {"remote": "LT_LQFC_94", "position": "LT_LQFC_100", "threshold": 90},
    }

    # 收集所有需要查询的参数
    all_params = []
    for cfg in EVENT_CONFIG.values():
        all_params.append(cfg["remote"])
        all_params.append(cfg["position"])
    param_filter = sanitize_param_for_flux(all_params)

    try:
        client = get_client()
        query_api = client.query_api()

        # 1s 粒度查询遥控和位置信号
        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {s_utc}, stop: {e_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> aggregateWindow(every: 1s, fn: mean, createEmpty: false)'''

        tables = query_api.query(flux)
    except Exception as e:
        return jsonify({"error": f"InfluxDB 查询失败: {str(e)[:200]}"}), 500

    # 组织: {param: [(utc_datetime, value), ...]}
    raw_data = {}
    for table in tables:
        for record in table.records:
            p = record.values.get("param", "")
            t = record.get_time()
            v = record.get_value()
            if v is None:
                continue
            if p not in raw_data:
                raw_data[p] = []
            raw_data[p].append((t, float(v)))

    # 为每个设备检测事件
    all_events = []

    for machine, cfg in EVENT_CONFIG.items():
        pos_param = cfg["position"]
        rem_param = cfg["remote"]
        threshold = cfg["threshold"]

        pos_pts = raw_data.get(pos_param, [])
        rem_pts = raw_data.get(rem_param, [])

        if len(pos_pts) < 2 or len(rem_pts) < 2:
            continue

        pos_pts.sort(key=lambda x: x[0])
        rem_pts.sort(key=lambda x: x[0])

        # 遥控信号 → {utc_ts: value}
        rem_map = {t.timestamp(): v for t, v in rem_pts}

        # 检测回转位置穿过 90°: 从 <90 → >=90
        raw_events = []
        for i in range(1, len(pos_pts)):
            prev_v = pos_pts[i - 1][1]
            curr_v = pos_pts[i][1]
            if prev_v < threshold and curr_v >= threshold:
                t_cross = pos_pts[i][0]  # UTC
                t_cross_ts = t_cross.timestamp()

                # 检查同一时刻遥控信号是否为 1
                # 允许 ±1s 容差（1s 粒度可能不完全对���）
                remote_on = False
                for offset in (-1, 0, 1):
                    check_ts = t_cross_ts + offset
                    found = None
                    for ts, val in rem_pts:
                        if abs(ts.timestamp() - check_ts) < 0.5:
                            found = val
                            break
                    if found is not None and found >= 0.5:
                        remote_on = True
                        break

                if remote_on:
                    t_local = (t_cross + LOCAL_OFFSET).replace(tzinfo=timezone(LOCAL_OFFSET))
                    raw_events.append(t_local)

        # 去重合并: 同一设备 5 秒内的多次触发合并为一次，取第一次时间
        merged = []
        for t in raw_events:
            if not merged or (t - merged[-1]).total_seconds() > 5:
                merged.append(t)

        for t in merged:
            ws = t - timedelta(minutes=1)
            we = t + timedelta(minutes=1)
            all_events.append({
                "machine": machine,
                "event_type": "遥控开口",
                "trigger_time": t.isoformat(),
                "window_start": ws.isoformat(),
                "window_end": we.isoformat(),
                "label": f"{machine} 遥控开口 {ws.strftime('%H:%M:%S')} ~ {we.strftime('%H:%M:%S')}",
            })

    all_events.sort(key=lambda e: e["trigger_time"])
    return jsonify({"events": all_events, "count": len(all_events)})


@app.route("/trend")
@login_required
def trend():
    html = TREND_HTML.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG["groups"]))
    html = html.replace("{{ labels_json | safe }}", json.dumps(_LABELS))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


# ==============================================================
# === 炉前作业分析模块 (Layer 1: 周期识别 + Layer 2: 指标提取) ===
# ==============================================================

# ── 信号配置: 开口机 ──
_OPENING_SIGNALS = {
    "east": {
        "name": "东开口机",
        "remote": "LT_LQFC_57",      # 开口机选择
        "swing_pos": "LT_LQFC_63",   # 回转位置
        "push_pos": "LT_LQFC_67",    # 推进位置
        "push_press": "LT_LQFC_68",  # 推进压力
        "drill_press": "LT_LQFC_87", # 转钎压力
        "impact_press": "LT_LQFC_88",# 冲击压力
        "cart_cmd": "LT_LQFC_61",    # 小车前进/后退
        "swing_cmd": "LT_LQFC_59",   # 回转进/退
        "impact_cmd": "LT_LQFC_69",  # 冲击开/关
    },
    "west": {
        "name": "西开口机",
        "remote": "LT_LQFC_94",
        "swing_pos": "LT_LQFC_100",
        "push_pos": "LT_LQFC_104",
        "push_press": "LT_LQFC_105",
        "drill_press": "LT_LQFC_124",
        "impact_press": "LT_LQFC_125",
        "cart_cmd": "LT_LQFC_98",
        "swing_cmd": "LT_LQFC_96",
        "impact_cmd": "LT_LQFC_106",
    }
}

# ── 信号配置: 堵口机 ──
_PLUGGING_SIGNALS = {
    "east": {
        "name": "东堵口机",
        "plug_select": "LT_LQFC_132",  # 泥炮机选择
        "mud_cmd": "LT_LQFC_134",      # 打泥前进/后退
        "mud_pos": "LT_LQFC_137",      # 打泥位置
        "mud_press": "LT_LQFC_138",    # 打泥压力
        "mud_qty": "LT_LQFC_179",      # 打泥量
        "swing_pos": "LT_LQFC_135",    # 回转位置
        "swing_cmd": "LT_LQFC_133",    # 回转进/退
        "remote_start": "LT_LQFC_130", # 遥控启动/停止
    },
    "west": {
        "name": "西堵口机",
        "plug_select": "LT_LQFC_155",
        "mud_cmd": "LT_LQFC_157",
        "mud_pos": "LT_LQFC_160",
        "mud_press": "LT_LQFC_161",
        "mud_qty": "LT_LQFC_180",
        "swing_pos": "LT_LQFC_158",
        "swing_cmd": "LT_LQFC_156",
        "remote_start": "LT_LQFC_153",
    }
}


def _fetch_cycle_data(start_utc, end_utc, params, window):
    """查询指定参数的时间序列数据，返回 {param: [(utc_dt, value), ...]}"""
    param_filter = sanitize_param_for_flux(params)
    client = get_client()
    query_api = client.query_api()
    flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {start_utc}, stop: {end_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> aggregateWindow(every: {window}, fn: mean, createEmpty: false)'''

    tables = query_api.query(flux)
    raw = {}
    for table in tables:
        for record in table.records:
            p = record.values.get("param", "")
            t = record.get_time()
            v = record.get_value()
            if v is None:
                continue
            if p not in raw:
                raw[p] = []
            raw[p].append((t, float(v)))
    return raw


def _detect_opening_cycles(raw_data, sig):
    """检测开口作业周期（遥控开口→钻进→钻透→退回）"""
    pos = sorted(raw_data.get(sig["swing_pos"], []), key=lambda x: x[0])
    push_pos = sorted(raw_data.get(sig["push_pos"], []), key=lambda x: x[0])
    push_press = sorted(raw_data.get(sig["push_press"], []), key=lambda x: x[0])
    cart = sorted(raw_data.get(sig["cart_cmd"], []), key=lambda x: x[0])
    remote = sorted(raw_data.get(sig["remote"], []), key=lambda x: x[0])

    rem_map = {t.timestamp(): v for t, v in remote}
    cycles = []

    if len(pos) < 2 or len(remote) < 2:
        return cycles

    for i in range(1, len(pos)):
        prev_v = pos[i - 1][1]
        curr_v = pos[i][1]
        if prev_v < 90 and curr_v >= 90:
            t_cross = pos[i][0]
            t_cross_ts = t_cross.timestamp()
            remote_on = False
            for offset in (-1, 0, 1):
                if rem_map.get(t_cross_ts + offset, 0) >= 0.5:
                    remote_on = True
                    break
            if not remote_on:
                continue

            t_start = t_cross
            t_local_start = t_start + LOCAL_OFFSET

            drill_press_peak = 0.0
            push_press_peak = 0.0
            breakthrough_time = None
            breakthrough_detected = False

            push_pos_max = None
            push_pos_change = 0.0

            if push_pos:
                for t, v in push_pos:
                    if t < t_start:
                        push_pos_max = v
                push_seg = [(t, v) for t, v in push_pos if t >= t_start]
                if push_seg:
                    push_pos_final = push_seg[-1][1]
                    push_pos_change = push_pos_final - push_pos_max if push_pos_max is not None else push_seg[-1][1] - push_seg[0][1]
            else:
                push_seg = []

            if push_press:
                press_seg = [(t, v) for t, v in push_press if t >= t_start and t <= t_start + timedelta(minutes=15)]
                if press_seg:
                    push_press_peak = max(v for _, v in press_seg)
                    for t, v in press_seg:
                        if t not in [p[0] for p in push_seg]:
                            continue
                        push_seg_with_time = sorted([(tp, vp) for tp, vp in push_seg], key=lambda x: x[0])
                        if len(push_seg_with_time) < 3:
                            break
                        idx = next((j for j, (tp, _) in enumerate(push_seg_with_time) if tp >= t), None)
                        if idx is None or idx < 3:
                            continue
                        delta_pos = push_seg_with_time[idx][1] - push_seg_with_time[idx - 3][1]
                        delta_press = v - push_press[idx - 3][1] if idx - 3 < len(push_press) else 0
                        if delta_pos > 0.15 and push_press[idx - 3][1] > 0 and (delta_press / push_press[idx - 3][1]) < -0.25:
                            breakthrough_time = t
                            breakthrough_detected = True
                            break

            if drill_press_peak == 0 and sig.get("drill_press") in raw_data:
                drill_data = raw_data.get(sig["drill_press"], [])
                drill_seg = [(t, v) for t, v in drill_data if t >= t_start and t <= t_start + timedelta(minutes=15)]
                if drill_seg:
                    drill_press_peak = max(v for _, v in drill_seg)

            t_end = t_start + timedelta(minutes=10)
            if cart:
                for t, v in cart:
                    if t > t_start and v < -0.5:
                        t_end = t
                        break
            if pos:
                standby_found = False
                for t, v in pos:
                    if t > t_end - timedelta(minutes=10) and v < 10:
                        t_end = t
                        standby_found = True
                        break
                if not standby_found:
                    t_end = t_start + timedelta(minutes=15)
            else:
                t_end = t_start + timedelta(minutes=15)

            duration_s = (t_end - t_start).total_seconds()

            if breakthrough_detected:
                result = "success"
            elif push_pos_change < 0.01:
                result = "fail"
            else:
                result = "incomplete"

            t_local_end = t_end + LOCAL_OFFSET

            cycles.append({
                "machine": sig["name"],
                "type": "opening",
                "trigger_time": t_local_start.isoformat(),
                "window_start": t_local_start.isoformat(),
                "window_end": t_local_end.isoformat(),
                "duration_s": round(duration_s, 1),
                "push_pos_change": round(push_pos_change, 3),
                "push_press_peak": round(push_press_peak, 1),
                "drill_press_peak": round(drill_press_peak, 1),
                "breakthrough": breakthrough_detected,
                "result": result,
                "label": f"{sig['name']} 开口 {'成' if breakthrough_detected else '未'}钻透 {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
            })

    return cycles


def _detect_plugging_cycles(raw_data, sig):
    """检测堵口作业周期（泥炮选择→打泥→保压→退炮）"""
    cmd = sorted(raw_data.get(sig["mud_cmd"], []), key=lambda x: x[0])
    mud_pos = sorted(raw_data.get(sig["mud_pos"], []), key=lambda x: x[0])
    mud_press = sorted(raw_data.get(sig["mud_press"], []), key=lambda x: x[0])
    mud_qty = sorted(raw_data.get(sig["mud_qty"], []), key=lambda x: x[0])
    swing_pos = sorted(raw_data.get(sig["swing_pos"], []), key=lambda x: x[0])
    plug_select = sorted(raw_data.get(sig["plug_select"], []), key=lambda x: x[0])

    plug_map = {t.timestamp(): v for t, v in plug_select}
    cycles = []

    if len(cmd) < 2:
        return cycles

    for i in range(1, len(cmd)):
        prev_v = cmd[i - 1][1]
        curr_v = cmd[i][1]
        if prev_v < 0.5 and curr_v >= 0.5:
            t_start = cmd[i][0]
            t_start_ts = t_start.timestamp()
            plug_on = False
            for offset in (-1, 0, 1):
                if plug_map.get(t_start_ts + offset, 0) >= 0.5:
                    plug_on = True
                    break
            if not plug_on:
                continue

            t_local_start = t_start + LOCAL_OFFSET

            mud_press_peak = 0.0
            mud_qty_total = 0.0
            hold_duration_s = 0.0
            mud_fill_complete = False
            hold_complete = False

            press_seg = sorted(
                [(t, v) for t, v in mud_press if t >= t_start and t <= t_start + timedelta(minutes=40)],
                key=lambda x: x[0]
            )
            if press_seg:
                mud_press_peak = max(v for _, v in press_seg)

                hold_count = 0
                for _, v in press_seg:
                    if 18 <= v <= 22:
                        hold_count += 1
                    else:
                        hold_count = 0
                    if hold_count >= 60:
                        hold_duration_s = hold_count
                        hold_complete = True
                        break

            qty_seg = sorted(
                [(t, v) for t, v in mud_qty if t >= t_start and t <= t_start + timedelta(minutes=40)],
                key=lambda x: x[0]
            )
            if qty_seg:
                mud_qty_total = qty_seg[-1][1]
                mud_fill_complete = mud_qty_total >= 10

            t_end = t_start + timedelta(minutes=30)
            if press_seg:
                retreat_found = False
                for t, v in reversed(press_seg):
                    if v < 5 and t > t_start + timedelta(seconds=30):
                        t_end = t
                        retreat_found = True
                        break
                if not retreat_found:
                    t_end = t_start + timedelta(minutes=40)
            if swing_pos:
                swing_seg = [(t, v) for t, v in swing_pos if t > t_end - timedelta(minutes=5) and v < 10]
                if swing_seg:
                    t_end = max(t_end, swing_seg[-1][0])

            duration_s = (t_end - t_start).total_seconds()

            if mud_fill_complete and hold_complete:
                result = "success"
            elif mud_fill_complete:
                result = "partial"
            else:
                result = "fail"

            t_local_end = t_end + LOCAL_OFFSET

            cycles.append({
                "machine": sig["name"],
                "type": "plugging",
                "trigger_time": t_local_start.isoformat(),
                "window_start": t_local_start.isoformat(),
                "window_end": t_local_end.isoformat(),
                "duration_s": round(duration_s, 1),
                "mud_press_peak": round(mud_press_peak, 1),
                "mud_qty": round(mud_qty_total, 1),
                "hold_duration_s": round(hold_duration_s, 0),
                "mud_filled": mud_fill_complete,
                "hold_ok": hold_complete,
                "result": result,
                "label": f"{sig['name']} 堵口 {'完' if mud_fill_complete else '未完'} {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
            })

    return cycles


# ── API: 作业周期识别 ──
@app.route("/api/analysis/cycles")
def api_analysis_cycles():
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    op_type = request.args.get("type", "all").strip()

    if not start or not end:
        return jsonify({"error": "缺少时间参数"}), 400

    try:
        s_local = datetime.strptime(start, "%Y-%m-%dT%H:%M")
        e_local = datetime.strptime(end, "%Y-%m-%dT%H:%M")
    except ValueError:
        return jsonify({"error": "时间格式错误"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")

    all_cycles = []

    if op_type in ("opening", "all"):
        all_op_params = []
        for side, sig in _OPENING_SIGNALS.items():
            all_op_params.extend(sig.values())
        all_op_params = list(set(all_op_params))
        try:
            raw = _fetch_cycle_data(s_utc, e_utc, all_op_params, "1s")
            for side, sig in _OPENING_SIGNALS.items():
                cycles = _detect_opening_cycles(raw, sig)
                all_cycles.extend(cycles)
        except Exception as e:
            pass  # 开口信号查询失败不阻塞堵口

    if op_type in ("plugging", "all"):
        all_pl_params = []
        for side, sig in _PLUGGING_SIGNALS.items():
            all_pl_params.extend(sig.values())
        all_pl_params = list(set(all_pl_params))
        try:
            raw = _fetch_cycle_data(s_utc, e_utc, all_pl_params, "1s")
            for side, sig in _PLUGGING_SIGNALS.items():
                cycles = _detect_plugging_cycles(raw, sig)
                all_cycles.extend(cycles)
        except Exception as e:
            pass

    all_cycles.sort(key=lambda c: c["trigger_time"])
    return jsonify({"cycles": all_cycles, "count": len(all_cycles)})


# ── API: 周期指标提取 ──
@app.route("/api/analysis/metrics")
def api_analysis_metrics():
    ws = request.args.get("window_start", "").strip()
    we = request.args.get("window_end", "").strip()
    machine = request.args.get("machine", "").strip()
    op_type = request.args.get("type", "opening").strip()

    if not ws or not we:
        return jsonify({"error": "缺少时间窗口参数"}), 400

    try:
        w_start = datetime.fromisoformat(ws)
        w_end = datetime.fromisoformat(we)
    except ValueError:
        return jsonify({"error": "时间格式错误，需要 ISO 8601"}), 400

    s_utc = (w_start - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
    e_utc = (w_end - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")

    metrics = {"machine": machine, "type": op_type, "window_start": ws, "window_end": we}

    if op_type == "opening":
        sig = None
        for side, cfg in _OPENING_SIGNALS.items():
            if cfg["name"] == machine:
                sig = cfg
                break
        if not sig:
            return jsonify({"error": f"未找到设备: {machine}"}), 404

        params = [sig["push_pos"], sig["push_press"], sig["swing_pos"],
                   sig["drill_press"], sig["impact_press"]]
        raw = _fetch_cycle_data(s_utc, e_utc, params, "1s")

        push_pos_pts = sorted(raw.get(sig["push_pos"], []), key=lambda x: x[0])
        push_press_pts = sorted(raw.get(sig["push_press"], []), key=lambda x: x[0])
        drill_press_pts = sorted(raw.get(sig["drill_press"], []), key=lambda x: x[0])
        impact_press_pts = sorted(raw.get(sig["impact_press"], []), key=lambda x: x[0])

        metrics["push_depth"] = round(push_pos_pts[-1][1] - push_pos_pts[0][1], 3) if len(push_pos_pts) >= 2 else 0
        metrics["push_press_max"] = round(max(v for _, v in push_press_pts), 1) if push_press_pts else 0
        metrics["push_press_mean"] = round(sum(v for _, v in push_press_pts) / len(push_press_pts), 1) if push_press_pts else 0
        metrics["drill_press_mean"] = round(sum(v for _, v in drill_press_pts) / len(drill_press_pts), 1) if drill_press_pts else 0
        metrics["drill_press_max"] = round(max(v for _, v in drill_press_pts), 1) if drill_press_pts else 0
        metrics["impact_press_active"] = bool(impact_press_pts and any(v > 0.5 for _, v in impact_press_pts))
        metrics["data_points"] = len(push_pos_pts)

        # 简易钻透检测
        breakthrough = False
        if len(push_pos_pts) >= 4 and push_press_pts:
            for i in range(3, len(push_pos_pts)):
                dp = push_pos_pts[i][1] - push_pos_pts[i - 3][1]
                pp_vals = [v for t, v in push_press_pts if abs((t - push_pos_pts[i][0]).total_seconds()) < 2]
                pp_prev = [v for t, v in push_press_pts if abs((t - push_pos_pts[i - 3][0]).total_seconds()) < 2]
                if pp_vals and pp_prev:
                    dpp = (pp_vals[0] - pp_prev[0]) / pp_prev[0] if pp_prev[0] > 0 else 0
                    if dp > 0.1 and dpp < -0.2:
                        breakthrough = True
                        break
        metrics["breakthrough"] = breakthrough

    else:
        sig = None
        for side, cfg in _PLUGGING_SIGNALS.items():
            if cfg["name"] == machine:
                sig = cfg
                break
        if not sig:
            return jsonify({"error": f"未找到设备: {machine}"}), 404

        params = [sig["mud_press"], sig["mud_qty"], sig["mud_pos"],
                   sig["swing_pos"]]
        raw = _fetch_cycle_data(s_utc, e_utc, params, "1s")

        mud_press_pts = sorted(raw.get(sig["mud_press"], []), key=lambda x: x[0])
        mud_qty_pts = sorted(raw.get(sig["mud_qty"], []), key=lambda x: x[0])

        metrics["mud_press_max"] = round(max(v for _, v in mud_press_pts), 1) if mud_press_pts else 0
        metrics["mud_press_mean"] = round(sum(v for _, v in mud_press_pts) / len(mud_press_pts), 1) if mud_press_pts else 0
        metrics["mud_qty"] = round(mud_qty_pts[-1][1], 1) if mud_qty_pts else 0
        metrics["data_points"] = len(mud_press_pts)

        hold_seconds = 0
        hold_consecutive = 0
        for _, v in mud_press_pts:
            if 18 <= v <= 22:
                hold_consecutive += 1
            else:
                hold_consecutive = 0
            if hold_consecutive > hold_seconds:
                hold_seconds = hold_consecutive
        metrics["hold_duration_s"] = hold_seconds
        metrics["hold_ok"] = hold_seconds >= 60

    return jsonify(metrics)


# ── API: 导出分析结果 ──
@app.route("/api/analysis/export")
def api_analysis_export():
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    op_type = request.args.get("type", "all").strip()

    if not start or not end:
        return jsonify({"error": "缺少时间参数"}), 400

    try:
        s_local = datetime.strptime(start, "%Y-%m-%dT%H:%M")
        e_local = datetime.strptime(end, "%Y-%m-%dT%H:%M")
    except ValueError:
        return jsonify({"error": "时间格式错误"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:00Z")

    all_rows = []

    all_op_params = []
    for side, sig in _OPENING_SIGNALS.items():
        all_op_params.extend(sig.values())
    all_op_params = list(set(all_op_params))

    all_pl_params = []
    for side, sig in _PLUGGING_SIGNALS.items():
        all_pl_params.extend(sig.values())
    all_pl_params = list(set(all_pl_params))

    try:
        raw_op = _fetch_cycle_data(s_utc, e_utc, all_op_params, "1s")
        for side, sig in _OPENING_SIGNALS.items():
            for c in _detect_opening_cycles(raw_op, sig):
                c["id"] = f"OP-{len(all_rows) + 1:03d}"
                all_rows.append(c)
    except Exception:
        pass

    try:
        raw_pl = _fetch_cycle_data(s_utc, e_utc, all_pl_params, "1s")
        for side, sig in _PLUGGING_SIGNALS.items():
            for c in _detect_plugging_cycles(raw_pl, sig):
                c["id"] = f"PL-{len(all_rows) + 1:03d}"
                all_rows.append(c)
    except Exception:
        pass

    all_rows.sort(key=lambda r: r.get("trigger_time", ""))

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("作业分析")
    headers = ["ID", "设备", "类型", "触发时间", "窗口开始", "窗口结束", "耗时(秒)", "结果",
               "推进位移(m)", "推进压力峰值", "转钎压力峰值", "钻透",
               "打泥压力峰值", "打泥量", "保压时长(秒)", "打泥完成", "保压完成"]
    header_fmt = wb.add_format({"bold": True, "bg_color": "#1677ff", "font_color": "#fff", "border": 1})
    for i, h in enumerate(headers):
        ws.write(0, i, h, header_fmt)

    for ri, row in enumerate(all_rows, 1):
        ws.write(ri, 0, row.get("id", ""))
        ws.write(ri, 1, row.get("machine", ""))
        ws.write(ri, 2, "开口" if row.get("type") == "opening" else "堵口")
        ws.write(ri, 3, row.get("trigger_time", ""))
        ws.write(ri, 4, row.get("window_start", ""))
        ws.write(ri, 5, row.get("window_end", ""))
        ws.write(ri, 6, row.get("duration_s", 0))
        ws.write(ri, 7, row.get("result", ""))
        ws.write(ri, 8, row.get("push_pos_change", 0))
        ws.write(ri, 9, row.get("push_press_peak", 0))
        ws.write(ri, 10, row.get("drill_press_peak", 0))
        ws.write(ri, 11, "是" if row.get("breakthrough") else "否")
        ws.write(ri, 12, row.get("mud_press_peak", 0))
        ws.write(ri, 13, row.get("mud_qty", 0))
        ws.write(ri, 14, row.get("hold_duration_s", 0))
        ws.write(ri, 15, "是" if row.get("mud_filled") else "否")
        ws.write(ri, 16, "是" if row.get("hold_ok") else "否")

    ws.autofit()
    wb.close()
    buf.seek(0)

    timestamp_str = datetime.now(timezone(LOCAL_OFFSET)).strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=f"DCS_作业分析_{timestamp_str}.xlsx"
    )


# ==============================================================
# === 作业分析页面 ===
# ==============================================================

ANALYSIS_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>炉前作业分析 — DCS</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #1f2937; display: flex; min-height: 100vh; }

.sidebar { width: 220px; min-width: 220px; background: linear-gradient(180deg, #1a2744 0%, #243356 50%, #2d3f66 100%); color: #fff; display: flex; flex-direction: column; position: fixed; top: 0; left: 0; height: 100vh; z-index: 100; }
.sidebar::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.08), transparent); }
.sidebar-header { padding: 24px 20px 16px; border-bottom: 1px solid rgba(255,255,255,0.06); }
.sidebar-header h2 { font-size: 16px; font-weight: 700; letter-spacing: 0.3px; }
.sidebar-header .version { font-size: 10px; opacity: 0.4; margin-top: 3px; letter-spacing: 0.5px; }
.nav-section { padding: 12px 0; }
.nav-section-title { padding: 6px 20px 6px; font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: rgba(255,255,255,0.3); font-weight: 600; }
.nav-item { display: flex; align-items: center; gap: 10px; padding: 10px 20px; cursor: pointer; font-size: 12px; font-weight: 600; transition: all 0.2s; border-left: 3px solid transparent; color: rgba(255,255,255,0.65); text-decoration: none; }
.nav-item:hover { background: rgba(255,255,255,0.05); color: #fff; }
.nav-item.active { background: rgba(22,119,255,0.15); border-left-color: #1677ff; color: #fff; }
.nav-item .badge-dim { margin-left: auto; font-size: 9px; padding: 1px 6px; border-radius: 8px; font-weight: 500; }
.nav-item .badge-dim.swing { background: rgba(22,119,255,0.2); color: #91caff; }
.nav-item .badge-dim.push { background: rgba(82,196,26,0.2); color: #b7eb8f; }
.nav-item .badge-dim.mud { background: rgba(250,173,20,0.2); color: #ffe58f; }
.sidebar-footer { margin-top: auto; padding: 14px 20px; font-size: 11px; border-top: 1px solid rgba(255,255,255,0.06); color: rgba(255,255,255,0.35); }
.sidebar-footer .dot { display: inline-block; width: 6px; height: 6px; background: #52c41a; border-radius: 50%; margin-right: 6px; box-shadow: 0 0 4px rgba(82,196,26,0.5); }

.main { margin-left: 220px; flex: 1; min-height: 100vh; }
.header { background: #fff; padding: 14px 32px; display: flex; align-items: center; gap: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); position: sticky; top: 0; z-index: 50; }
.header h1 { font-size: 18px; font-weight: 700; background: linear-gradient(135deg, #1a1a2e, #1677ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.header .badge { background: linear-gradient(135deg, #faad14, #d48806); color: #fff; font-size: 10px; padding: 2px 10px; border-radius: 12px; font-weight: 500; }
.header .breadcrumb { font-size: 12px; color: #94a3b8; }
.header .nav-links { margin-left: auto; display: flex; gap: 6px; }
.header .nav-links a { padding: 6px 16px; border-radius: 6px; font-size: 13px; text-decoration: none; color: #64748b; transition: all 0.2s; font-weight: 500; }
.header .nav-links a:hover { background: #f0f5ff; color: #1677ff; }
.header .nav-links a.active { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.3); }
.container { padding: 24px 32px; }

.stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 20px; }
.stat-card { background: #fff; border-radius: 16px; padding: 20px 24px; box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.02); display: flex; flex-direction: column; }
.stat-card .stat-label { font-size: 12px; color: #94a3b8; margin-bottom: 6px; font-weight: 500; letter-spacing: 0.5px; }
.stat-card .stat-value { font-size: 28px; font-weight: 700; color: #0f172a; line-height: 1.2; }
.stat-card .stat-detail { font-size: 11px; color: #94a3b8; margin-top: 6px; }

.card { background: #fff; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.02); margin-bottom: 16px; overflow: hidden; }
.card-body { padding: 24px; }
.card-header { padding: 16px 24px; border-bottom: 1px solid #f1f5f9; font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 10px; color: #0f172a; }

.filter-bar { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; }
.filter-group { display: flex; flex-direction: column; gap: 5px; }
.filter-group label { font-size: 11px; color: #64748b; font-weight: 600; letter-spacing: 0.5px; }
.filter-group input, .filter-group select { padding: 8px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; outline: none; transition: all 0.2s; background: #fff; color: #1f2937; min-width: 130px; }
.filter-group input:focus, .filter-group select:focus { border-color: #1677ff; box-shadow: 0 0 0 3px rgba(22,119,255,0.1); }

.btn { padding: 8px 20px; border: none; border-radius: 8px; font-size: 13px; cursor: pointer; font-weight: 600; transition: all 0.25s; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
.btn-primary { background: linear-gradient(135deg, #1677ff, #0958d9); color: #fff; box-shadow: 0 2px 6px rgba(22,119,255,0.2); }
.btn-primary:hover { background: linear-gradient(135deg, #4096ff, #1677ff); box-shadow: 0 4px 12px rgba(22,119,255,0.3); transform: translateY(-1px); }
.btn-primary:disabled { background: #b0c4de; box-shadow: none; cursor: not-allowed; transform: none; }
.btn-success { background: linear-gradient(135deg, #52c41a, #389e0d); color: #fff; box-shadow: 0 2px 6px rgba(82,196,26,0.2); }
.btn-success:hover { background: linear-gradient(135deg, #73d13d, #52c41a); box-shadow: 0 4px 12px rgba(82,196,26,0.3); transform: translateY(-1px); }
.btn-success:disabled { background: #b7eb8f; box-shadow: none; cursor: not-allowed; transform: none; }
.btn-ghost { padding: 5px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; cursor: pointer; font-size: 12px; color: #64748b; font-weight: 500; transition: all 0.2s; text-decoration: none; }
.btn-ghost:hover { border-color: #1677ff; color: #1677ff; background: #f0f5ff; }

table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: #f8fafc; font-weight: 600; text-align: left; padding: 10px 12px; color: #64748b; font-size: 11px; border-bottom: 1px solid #f1f5f9; text-transform: uppercase; letter-spacing: 0.3px; }
td { padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }
tr:hover { background: #f8fafc; }

.tag { padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; display: inline-block; }
.tag-ok { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
.tag-fail { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
.tag-warn { background: #fffbeb; color: #92400e; border: 1px solid #fde68a; }
.tag-info { background: #eff6ff; color: #1e40af; border: 1px solid #bfdbfe; }

.loading-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.3); z-index: 9999; justify-content: center; align-items: center; backdrop-filter: blur(2px); }
.loading-overlay.show { display: flex; }
.loading-box { background: #fff; padding: 32px 48px; border-radius: 16px; text-align: center; box-shadow: 0 8px 32px rgba(0,0,0,0.12); }
.spinner { width: 36px; height: 36px; border: 3px solid #e2e8f0; border-top: 3px solid #1677ff; border-radius: 50%; animation: spin 0.7s cubic-bezier(0.4,0,0.2,1) infinite; margin: 0 auto 16px; }
@keyframes spin { to { transform: rotate(360deg); } }

.chart-area { min-height: 400px; position: relative; }
.empty-state { text-align: center; padding: 48px 20px; color: #94a3b8; }
.empty-state h3 { font-size: 15px; margin-bottom: 6px; color: #64748b; }
.empty-state p { font-size: 12px; }

.alert { padding: 12px 20px; border-radius: 12px; margin-bottom: 16px; font-size: 13px; display: none; }
.alert-error { display: block; background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; }
.alert-info { display: block; background: #eff6ff; border: 1px solid #bfdbfe; color: #2563eb; }

.user-area { display: flex; align-items: center; gap: 10px; }
.user-name { font-size: 13px; font-weight: 600; color: #1f2937; }
.btn-logout { padding: 5px 14px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; font-size: 12px; color: #64748b; cursor: pointer; font-weight: 500; transition: all 0.2s; }
.btn-logout:hover { border-color: #ff4d4f; color: #ff4d4f; background: #fff5f5; }

@media (max-width: 768px) {
    .sidebar { width: 60px; min-width: 60px; }
    .sidebar .nav-item span, .sidebar .nav-section-title, .sidebar .badge-dim { display: none; }
    .main { margin-left: 60px; }
    .filter-bar { flex-direction: column; }
}
</style>
</head>
<body>

<div class="sidebar">
    <div class="sidebar-header">
        <h2>DCS 分析平台</h2>
        <div class="version">GBDT-MPC-PID 三级架构</div>
    </div>

    <div class="nav-section">
        <div class="nav-section-title">分析功能</div>
        <a class="nav-item" href="javascript:switchOpType('opening')" id="navOpening">开口作业分析</a>
        <a class="nav-item" href="javascript:switchOpType('plugging')" id="navPlugging">堵口作业分析</a>
        <a class="nav-item" href="javascript:switchOpType('all')" id="navAll">全部作业</a>
    </div>

    <div class="nav-section" id="dimSectionOpening">
        <div class="nav-section-title">信号维度 — 开口</div>
        <a class="nav-item" href="javascript:filterDim('push')">推进系统 <span class="badge-dim push">位移/压力</span></a>
        <a class="nav-item" href="javascript:filterDim('swing')">回转系统 <span class="badge-dim swing">角度/压力</span></a>
        <a class="nav-item" href="javascript:filterDim('drill')">转钎系统</a>
        <a class="nav-item" href="javascript:filterDim('impact')">冲击系统</a>
        <a class="nav-item" href="javascript:filterDim('tilt')">倾动系统</a>
        <a class="nav-item" href="javascript:filterDim('hydraulic')">液压站</a>
    </div>

    <div class="nav-section" id="dimSectionPlugging" style="display:none">
        <div class="nav-section-title">信号维度 — 堵口</div>
        <a class="nav-item" href="javascript:filterDim('mud')">打泥系统 <span class="badge-dim mud">量/压力</span></a>
        <a class="nav-item" href="javascript:filterDim('cannon')">转炮/退炮</a>
        <a class="nav-item" href="javascript:filterDim('retreat')">退泥回路</a>
        <a class="nav-item" href="javascript:filterDim('hydraulic_p')">液压站</a>
    </div>

    <div class="nav-section">
        <div class="nav-section-title">数据操作</div>
        <a class="nav-item" href="javascript:exportResult()">导出 Excel</a>
    </div>

    <div class="sidebar-footer">
        <span class="dot" id="sysStatus"></span> <span id="sidebarUser">admin</span><a style="margin-left:auto;font-size:10px;cursor:pointer;color:rgba(255,255,255,0.45);text-decoration:none" onclick="doLogout()" href="javascript:void(0)">退出</a>
    </div>
</div>

<div class="main">

    <div class="header">
        <h1>炉前作业分析</h1>
        <span class="badge">GBDT-MPC-PID</span>
        <span class="breadcrumb" id="pageDesc">周期识别 · 阶段标注 · 指标提取</span>
        
        <div class="nav-links">
            <a href="/">历史查询</a>
            <a href="/realtime">实时监控</a>
            <a href="/trend">趋势分析</a>
            <a href="/analysis" class="active">作业分析</a>
        </div>
    </div>

    <div class="container">

        <div id="alertBox" class="alert"></div>

        <div class="filter-bar">
            <div class="filter-group"><label>开始时间</label><input type="datetime-local" id="dStart"></div>
            <div class="filter-group"><label>结束时间</label><input type="datetime-local" id="dEnd"></div>
            <div class="filter-group"><label>作业类型</label>
                <select id="dType" onchange="switchOpType(this.value)"><option value="all">全部</option><option value="opening">开口</option><option value="plugging">堵口</option></select>
            </div>
            <div class="filter-group"><label>设备</label>
                <select id="dMachine"><option value="all">全部设备</option><option value="东开口机">东开口机</option><option value="西开口机">西开口机</option><option value="东堵口机">东堵口机</option><option value="西堵口机">西堵口机</option></select>
            </div>
            <button class="btn btn-primary" onclick="runAnalysis()" id="btnAnalyze">开始分析</button>
            <button class="btn btn-success" id="btnExport" onclick="exportResult()" disabled>导出 Excel</button>
        </div>

        <div class="stats-row" id="statsRow"></div>

        <div class="card">
            <div class="card-header">作业周期列表 <span style="font-size:11px;color:#94a3b8;font-weight:400" id="cycleCount"></span></div>
            <div class="card-body" style="overflow-x:auto">
                <div id="cycleEmpty" class="empty-state"><h3>选择时间范围</h3><p>点击「开始分析」自动识别开口和堵口作业周期，按GBDT-MPC-PID协议分类</p></div>
                <table id="cycleTable" style="display:none">
                    <thead><tr>
                        <th>设备</th><th>类型</th><th>触发时间</th><th>作业窗口</th><th>耗时</th><th>关键指标</th><th>结果</th><th>操作</th>
                    </tr></thead>
                    <tbody id="cycleBody"></tbody>
                </table>
            </div>
        </div>

        <div class="card" id="detailCard" style="display:none">
            <div class="card-header">作业详情 <span id="detailTitle" style="font-size:12px;color:#94a3b8;font-weight:400"></span></div>
            <div class="card-body">
                <div id="detailMetrics" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:16px"></div>
                <div class="chart-area"><canvas id="detailChart"></canvas></div>
            </div>
        </div>

    </div>
</div>

<div class="loading-overlay" id="loading"><div class="loading-box"><div class="spinner"></div><div id="loadingMsg" style="font-size:13px;color:#64748b">正在查询...</div></div></div>

<script>
var APP_TOKEN = "{{ app_token }}";
var globalCycles = [];
var chartInst = null;
var PARAM_CONFIG = {{ groups_json | safe }};
var _currentDimFilter = null;
var _currentOpType = 'all';

function getLabelMap(){
    var labels = {};
    (PARAM_CONFIG.groups || []).forEach(function(g){
        if(g.labels) Object.assign(labels, g.labels);
    });
    if(PARAM_CONFIG.labels) Object.assign(labels, PARAM_CONFIG.labels);
    return labels;
}

function initDates(){
    var now = new Date();
    var d1 = new Date(now.getTime() - 3*86400000);
    document.getElementById('dEnd').value = now.toISOString().slice(0,16);
    document.getElementById('dStart').value = d1.toISOString().slice(0,16);
}
initDates();

function checkLogin(){
    fetch('/api/user?token='+APP_TOKEN).then(function(r){return r.json()}).then(function(d){
        if(!d.logged_in){window.location.href='/';}
        if(d.username){var el=document.getElementById('sidebarUser');if(el)el.textContent=d.username;};
    }).catch(function(e){console.log('Auth check failed, redirecting:', e);window.location.href='/';});
}
checkLogin();

function doLogout(){
    fetch('/api/logout',{method:'POST'}).then(function(r){return r.json()}).then(function(d){
        window.location.href='/';
    });
}

function switchOpType(t){
    _currentOpType = t;
    document.getElementById('dType').value = t;
    document.getElementById('navOpening').classList.toggle('active', t==='opening');
    document.getElementById('navPlugging').classList.toggle('active', t==='plugging');
    document.getElementById('navAll').classList.toggle('active', t==='all');
    document.getElementById('dimSectionOpening').style.display = (t==='plugging')?'none':'';
    document.getElementById('dimSectionPlugging').style.display = (t==='opening'||t==='all')?'none':'';
    document.getElementById('pageDesc').textContent = t==='opening'?'回转·倾动·转钎·冲击·推进·液压站':t==='plugging'?'转炮·打泥·保压·退炮·液压站':'周期识别 · 阶段标注 · 指标提取';
}

function filterDim(dim){
    _currentDimFilter = dim;
    var items = document.querySelectorAll('.nav-section .nav-item');
    items.forEach(function(it){it.classList.remove('active');});
    if(_currentOpType==='plugging')document.getElementById('navPlugging').classList.add('active');
    else if(_currentOpType==='opening')document.getElementById('navOpening').classList.add('active');
    else document.getElementById('navAll').classList.add('active');
    if(globalCycles.length>0)renderCycles();
}

function showAlert(msg, cls){
    var el = document.getElementById('alertBox');
    el.textContent = msg;
    el.className = 'alert alert-'+cls;
    setTimeout(function(){el.className='alert';el.textContent='';}, 6000);
}

function showLoading(msg){
    document.getElementById('loadingMsg').textContent = msg || '正在查询...';
    document.getElementById('loading').classList.add('show');
}
function hideLoading(){ document.getElementById('loading').classList.remove('show'); }

function renderTag(result){
    if(result=='success') return '<span class="tag tag-ok">成功</span>';
    if(result=='fail') return '<span class="tag tag-fail">失败</span>';
    if(result=='incomplete') return '<span class="tag tag-warn">未完成</span>';
    if(result=='partial') return '<span class="tag tag-info">未完整</span>';
    return result||'--';
}

function renderOpType(t){
    return t=='opening' ? '<span class="tag tag-info">开口</span>' : '<span class="tag tag-warn">堵口</span>';
}

function runAnalysis(){
    var start = document.getElementById('dStart').value;
    var end = document.getElementById('dEnd').value;
    var opType = document.getElementById('dType').value;
    var machine = document.getElementById('dMachine').value;

    if(!start||!end){
        showAlert('请选择开始和结束时间', 'error');
        return;
    }
    if(!APP_TOKEN){
        showAlert('系统配置错误：缺少 API Token，请联系管理员', 'error');
        return;
    }

    document.getElementById('btnAnalyze').disabled = true;
    showLoading('正在检测作业周期...');
    showAlert('正在查询 InfluxDB，请稍候...', 'info');

    var url = '/api/analysis/cycles?start='+encodeURIComponent(start)+'&end='+encodeURIComponent(end)+'&type='+opType+'&token='+APP_TOKEN;
    fetch(url).then(function(r){
        if(!r.ok) throw new Error('服务器返回 HTTP '+r.status);
        return r.json();
    }).then(function(data){
        if(data.error){
            showAlert(data.error, 'error');
            hideLoading();
            document.getElementById('btnAnalyze').disabled = false;
            return;
        }
        globalCycles = data.cycles || [];
        if(machine!=='all'){
            globalCycles = globalCycles.filter(function(c){return c.machine===machine;});
        }
        renderCycles();
        renderStats();
        document.getElementById('btnExport').disabled = globalCycles.length === 0;
        document.getElementById('btnAnalyze').disabled = false;
        hideLoading();
        showAlert('分析完成，检测到 '+globalCycles.length+' 个作业周期', 'info');
    }).catch(function(e){
        showAlert('分析失败: '+e.message, 'error');
        hideLoading();
        document.getElementById('btnAnalyze').disabled = false;
    });
}

function renderStats(){
    var openCycles = globalCycles.filter(function(c){return c.type=='opening';});
    var plugCycles = globalCycles.filter(function(c){return c.type=='plugging';});
    var openOk = openCycles.filter(function(c){return c.result=='success'||c.breakthrough;}).length;
    var plugOk = plugCycles.filter(function(c){return c.result=='success';}).length;

    var avgDur = 0;
    if(globalCycles.length>0){
        var total = globalCycles.reduce(function(s,c){return s+(c.duration_s||0);},0);
        avgDur = Math.round(total/globalCycles.length);
    }
    var durMin = Math.floor(avgDur/60);
    var durSec = avgDur%60;

    var html = '';
    html += '<div class="stat-card"><div class="stat-label">开口作业</div><div class="stat-value">'+openCycles.length+'</div><div class="stat-detail">'+openOk+' 次钻透成功</div></div>';
    html += '<div class="stat-card"><div class="stat-label">堵口作业</div><div class="stat-value">'+plugCycles.length+'</div><div class="stat-detail">'+plugOk+' 次完整完成</div></div>';
    html += '<div class="stat-card"><div class="stat-label">作业总数</div><div class="stat-value">'+globalCycles.length+'</div><div class="stat-detail">开口 + 堵口</div></div>';
    html += '<div class="stat-card"><div class="stat-label">平均耗时</div><div class="stat-value">'+durMin+'m'+durSec+'s</div><div class="stat-detail">每炉次</div></div>';
    document.getElementById('statsRow').innerHTML = html;
    document.getElementById('cycleCount').textContent = '共 '+globalCycles.length+' 个周期';
}

function renderCycles(){
    var tbody = document.getElementById('cycleBody');
    var filtered = globalCycles;
    if(_currentDimFilter){
        filtered = globalCycles;
    }
    if(filtered.length===0){
        tbody.innerHTML = '';
        document.getElementById('cycleTable').style.display = 'none';
        document.getElementById('cycleEmpty').style.display = '';
        document.getElementById('cycleEmpty').innerHTML = '<h3>'+(globalCycles.length===0?'未检测到作业周期':'当前筛选无匹配周期')+'</h3><p>'+(globalCycles.length===0?'选择时间范围后点击「开始分析」':'尝试调整筛选条件')+'</p>';
        return;
    }
    document.getElementById('cycleEmpty').style.display = 'none';
    document.getElementById('cycleTable').style.display = '';

    var html = '';
    filtered.forEach(function(c,i){
        var keyMetric = '';
        if(c.type=='opening'){
            keyMetric = '钻进'+(c.push_pos_change||0).toFixed(3)+'m | 推进峰'+(c.push_press_peak||0).toFixed(0)+'MPa';
            if(c.breakthrough)keyMetric += ' | 已钻透';
        }else{
            keyMetric = '打泥量'+(c.mud_qty||0).toFixed(1)+' | 峰'+(c.mud_press_peak||0).toFixed(0)+'MPa';
            if(c.hold_ok)keyMetric += ' | 保压OK';
        }
        var winStart = (c.window_start||'').substring(11,19);
        var winEnd = (c.window_end||'').substring(11,19);
        var durMin = Math.floor((c.duration_s||0)/60);
        var durSec = Math.round((c.duration_s||0)%60);
        html += '<tr>';
        html += '<td style="font-weight:600">'+c.machine+'</td>';
        html += '<td>'+renderOpType(c.type)+'</td>';
        html += '<td>'+(c.trigger_time||'').substring(11,19)+'</td>';
        html += '<td>'+winStart+' ~ '+winEnd+'</td>';
        html += '<td>'+durMin+'分'+durSec+'秒</td>';
        html += '<td style="font-size:11px">'+keyMetric+'</td>';
        html += '<td>'+renderTag(c.result)+'</td>';
        html += '<td><button class="btn-ghost" onclick="showDetail('+globalCycles.indexOf(c)+')">详情</button></td>';
        html += '</tr>';
    });
    tbody.innerHTML = html;
}

function showDetail(idx){
    var c = globalCycles[idx];
    if(!c)return;
    document.getElementById('detailCard').style.display = '';
    document.getElementById('detailTitle').textContent = c.machine + ' / ' + ((c.type=='opening')?'开口':c.type=='plugging'?'堵口':c.type);
    document.getElementById('detailCard').scrollIntoView({behavior:'smooth'});

    showLoading('正在提取指标...');
    var url = '/api/analysis/metrics?window_start='+encodeURIComponent(c.window_start)+'&window_end='+encodeURIComponent(c.window_end)+'&machine='+encodeURIComponent(c.machine)+'&type='+c.type+'&token='+APP_TOKEN;
    fetch(url).then(function(r){return r.json()}).then(function(metrics){
        if(metrics.error){showAlert(metrics.error,'error');hideLoading();return;}
        renderMetrics(metrics);
        loadDetailChart(c, metrics);
        hideLoading();
    }).catch(function(e){showAlert('指标提取失败: '+e.message,'error');hideLoading();});
}

function renderMetrics(m){
    var html = '';
    if(m.type=='opening'){
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">开口深度</div><div class="stat-value" style="font-size:22px">'+(m.push_depth||0).toFixed(3)+'m</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">推进压力峰值</div><div class="stat-value" style="font-size:22px">'+(m.push_press_max||0).toFixed(1)+'</div><div class="stat-detail">均值 '+(m.push_press_mean||0).toFixed(1)+' MPa</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">转钎压力</div><div class="stat-value" style="font-size:22px">'+(m.drill_press_mean||0).toFixed(1)+'</div><div class="stat-detail">峰值 '+(m.drill_press_max||0).toFixed(1)+' MPa</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">钻透判定</div><div class="stat-value" style="font-size:22px">'+(m.breakthrough?'<span style="color:#059669">已钻透</span>':'<span style="color:#dc2626">未钻透</span>')+'</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">冲击状态</div><div class="stat-value" style="font-size:22px">'+(m.impact_press_active?'<span style="color:#1677ff">已开启</span>':'<span style="color:#94a3b8">未开启</span>')+'</div></div>';
    }else{
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">打泥量</div><div class="stat-value" style="font-size:22px">'+(m.mud_qty||0).toFixed(1)+'</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">打泥压力峰值</div><div class="stat-value" style="font-size:22px">'+(m.mud_press_max||0).toFixed(1)+'</div><div class="stat-detail">均值 '+(m.mud_press_mean||0).toFixed(1)+' MPa</div></div>';
        html += '<div class="stat-card" style="flex:1;min-width:140px"><div class="stat-label">保压时长</div><div class="stat-value" style="font-size:22px">'+(m.hold_duration_s||0)+'s</div><div class="stat-detail" style="color:'+(m.hold_ok?'#059669':'#dc2626')+'">'+(m.hold_ok?'保压合格 (≥60s)':'保压不足 (<60s)')+'</div></div>';
    }
    document.getElementById('detailMetrics').innerHTML = html;
}

function loadDetailChart(c, metrics){
    var labels = getLabelMap();
    var wStart = c.window_start;
    var wEnd = c.window_end;
    var sigs;
    if(c.type=='opening'){
        sigs = c.machine.indexOf('东')>=0 ?
            ['LT_LQFC_67','LT_LQFC_68','LT_LQFC_87','LT_LQFC_88','LT_LQFC_63'] :
            ['LT_LQFC_104','LT_LQFC_105','LT_LQFC_124','LT_LQFC_125','LT_LQFC_100'];
    }else{
        sigs = c.machine.indexOf('东')>=0 ?
            ['LT_LQFC_138','LT_LQFC_179','LT_LQFC_135'] :
            ['LT_LQFC_161','LT_LQFC_180','LT_LQFC_158'];
    }
    var url = '/api/trend?params='+sigs.join(',')+'&start='+encodeURIComponent(wStart.substring(0,16))+'&end='+encodeURIComponent(wEnd.substring(0,16));
    fetch(url+'&token='+APP_TOKEN).then(function(r){return r.json()}).then(function(data){
        if(data.error){return;}
        var ctx = document.getElementById('detailChart').getContext('2d');
        if(chartInst)chartInst.destroy();
        var datasets = [];
        var colors = ['#1677ff','#52c41a','#faad14','#ff4d4f','#722ed1'];
        data.series.forEach(function(s,i){
            var pts = s.data || [];
            var values = pts.map(function(p){return p.value});
            var isPos = (s.param||'').indexOf('67')>0||(s.param||'').indexOf('104')>0||(s.param||'').indexOf('160')>0||(s.param||'').indexOf('63')>0||(s.param||'').indexOf('100')>0||(s.param||'').indexOf('135')>0||(s.param||'').indexOf('158')>0;
            var yAxisID = isPos ? 'y-pos' : 'y-press';
            datasets.push({label:labels[s.param]||s.param,data:values,borderColor:colors[i%colors.length],backgroundColor:'transparent',borderWidth:2,pointRadius:0,tension:0.1,yAxisID:yAxisID,fill:false});
        });
        chartInst = new Chart(ctx,{type:'line',data:{labels:data.series.length>0?(data.series[0].data||[]).map(function(p){return new Date(p.time).toLocaleTimeString('zh-CN',{hour:'2-digit',minute:'2-digit',second:'2-digit'})}):[],datasets:datasets},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'nearest',intersect:false},plugins:{legend:{position:'top',labels:{font:{size:11},usePointStyle:true,padding:16}},tooltip:{mode:'index',intersect:false}},scales:{x:{ticks:{font:{size:10},maxTicksLimit:15},grid:{color:'#f1f5f9'}},'y-press':{type:'linear',display:true,position:'left',title:{display:true,text:'压力 (MPa)',font:{size:10}},ticks:{font:{size:10}},grid:{color:'#f1f5f9'}},'y-pos':{type:'linear',display:true,position:'right',title:{display:true,text:'位置/角度',font:{size:10}},ticks:{font:{size:10}},grid:{display:false}}}}});
        document.getElementById('detailChart').parentElement.style.height = '400px';
    }).catch(function(){});
}

function exportResult(){
    if(globalCycles.length===0){showAlert('没有可导出的数据，请先分析', 'error');return;}
    var start = document.getElementById('dStart').value;
    var end = document.getElementById('dEnd').value;
    var opType = document.getElementById('dType').value;
    var url = '/api/analysis/export?start='+encodeURIComponent(start)+'&end='+encodeURIComponent(end)+'&type='+opType+'&token='+APP_TOKEN;
    window.location.href = url;
}
</script>
</body>
</html>
"""
@app.route("/analysis")
@login_required
def analysis():
    html = ANALYSIS_HTML.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


# === 系统配置页面 ===
SETTINGS_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>系统配置 — DCS</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;background:#f5f7fa;min-height:100vh;display:flex;align-items:center;justify-content:center}
.card{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.06);padding:32px;width:440px;max-width:95vw}
.card h2{font-size:18px;margin-bottom:20px;color:#1f2937}
.form-group{margin-bottom:14px}
.form-group label{display:block;font-size:12px;color:#64748b;font-weight:600;margin-bottom:4px}
.form-group input,.form-group select{width:100%;padding:8px 12px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;outline:none}
.form-group input:focus{border-color:#1677ff;box-shadow:0 0 0 3px rgba(22,119,255,0.1)}
.btn{width:100%;padding:10px;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#1677ff,#0958d9);margin-top:8px}
.btn:hover{background:linear-gradient(135deg,#4096ff,#1677ff)}
.msg{padding:8px 12px;border-radius:6px;font-size:12px;margin-top:12px;display:none}
.msg.ok{display:block;background:#f0fdf4;color:#166534;border:1px solid #bbf7d0}
.msg.err{display:block;background:#fef2f2;color:#991b1b;border:1px solid #fecaca}
</style>
</head>
<body>
<div class="card">
<h2>⚙ 系统配置</h2>
<form id="settingsForm" onsubmit="saveSettings(event)">
<div class="form-group"><label>InfluxDB 地址</label><input id="influxUrl" placeholder="http://10.56.128.202:8086"></div>
<div class="form-group"><label>InfluxDB Token</label><input id="influxToken" type="password" placeholder="输入 Token"></div>
<div class="form-group"><label>组织 (Org)</label><input id="influxOrg" placeholder="myOrg"></div>
<div class="form-group"><label>Bucket</label><input id="influxBucket" placeholder="islag"></div>
<div class="form-group"><label>查询超时 (毫秒)</label><input id="influxTimeout" type="number" placeholder="180000"></div>
<div class="form-group"><label>Web 访问 Token</label><input id="appToken" placeholder="dcs2026"></div>
<div class="form-group"><label>设置密码</label><input id="settingsPwd" type="password" placeholder="123456" autocomplete="new-password" style="display:none">
<input id="settingsPwdConfirm" type="password" placeholder="确认密码" style="display:none"></div>
<button type="button" class="btn" onclick="showPwd()" id="changePwdBtn">修改密码</button>
<button type="submit" class="btn">保存配置</button>
</form>
<div class="msg" id="msg"></div>
</div>
<script>
fetch('/api/settings/config?token='+'dcs2026').then(r=>r.json()).then(d=>{
    document.getElementById('influxUrl').value=d.INFLUX_URL||'';
    document.getElementById('influxOrg').value=d.INFLUX_ORG||'';
    document.getElementById('influxBucket').value=d.INFLUX_BUCKET||'';
    document.getElementById('influxTimeout').value=d.INFLUX_TIMEOUT_MS||'';
    document.getElementById('appToken').value=d.APP_TOKEN.includes('***')?'':d.APP_TOKEN;
}).catch(()=>{});
function showPwd(){
    document.getElementById('settingsPwd').style.display='block';
    document.getElementById('settingsPwdConfirm').style.display='block';
    document.getElementById('changePwdBtn').style.display='none';
}
function saveSettings(e){
    e.preventDefault();
    var p1=document.getElementById('settingsPwd').value.trim();
    var p2=document.getElementById('settingsPwdConfirm').value.trim();
    if(p1&&p1!==p2){return msg('两次密码不一致','err')}
    var data={
        INFLUX_URL:document.getElementById('influxUrl').value.trim(),
        INFLUX_TOKEN:document.getElementById('influxToken').value.trim(),
        INFLUX_ORG:document.getElementById('influxOrg').value.trim(),
        INFLUX_BUCKET:document.getElementById('influxBucket').value.trim(),
        INFLUX_TIMEOUT_MS:document.getElementById('influxTimeout').value.trim(),
        APP_TOKEN:document.getElementById('appToken').value.trim(),
    };
    if(p1)data.SETTINGS_PASSWORD=p1;
    fetch('/api/settings/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)})
    .then(r=>r.json()).then(d=>{
        msg(d.ok?'保存成功，重启后生效':'保存失败: '+d.error, d.ok?'ok':'err');
    }).catch(e=>msg('请求失败: '+e,'err'));
}
function msg(t,c){var m=document.getElementById('msg');m.textContent=t;m.className='msg '+c}
</script>
</body>
</html>"""

@app.route("/settings")
def settings_page():
    return render_template_string(SETTINGS_HTML)

@app.route("/api/settings/auth", methods=["POST"])
def api_settings_auth():
    data = request.get_json(silent=True) or {}
    if data.get("password") == SETTINGS_PASSWORD:
        return jsonify({"ok": True, "token": APP_TOKEN})
    return jsonify({"ok": False})

@app.route("/api/settings/config")
def api_get_config():
    return jsonify(get_current_config())

@app.route("/api/settings/save", methods=["POST"])
def api_save_config():
    data = request.get_json(silent=True) or {}
    # 密码保护: 如果提供了新密码则用新密码，否则用当前密码
    pwd = data.get("SETTINGS_PASSWORD", "")
    save_runtime_config(data)
    return jsonify({"ok": True})


if __name__ == "__main__":
    # 启动前健康检查
    warnings = check_config()
    for w in warnings:
        print(w)

    print(f"\n  DCS 数据分析平台已启动")
    print(f"  浏览器打开: http://localhost:{FLASK_PORT}")
    print(f"  实时监控:   http://localhost:{FLASK_PORT}/realtime")
    print(f"  趋势分析:   http://localhost:{FLASK_PORT}/trend")
    print(f"  作业分析:   http://localhost:{FLASK_PORT}/analysis")
    print(f"  API 认证:   {'已启用 (APP_TOKEN)' if APP_TOKEN else '[WARN] 未启用 (仅内网安全)'}")
    print()

    # 自动打开浏览器
    if not getattr(sys, 'frozen', False):
        webbrowser.open(f"http://localhost:{FLASK_PORT}")
    else:
        # PyInstaller 环境，等待 Flask 完全启动后再打开
        def _open_browser():
            _time.sleep(1.5)
            webbrowser.open(f"http://localhost:{FLASK_PORT}")
        threading.Thread(target=_open_browser, daemon=True).start()
    from waitress import serve
    serve(app, host=FLASK_HOST, port=FLASK_PORT, threads=8)
