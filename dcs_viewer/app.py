"""
DCS 数据分析平台 — 开口机/堵口机专用查询工具
支持: 四组设备分组查询、时序分析、Excel导出
"""
import functools
import io
import json
import logging
import os
import sys
import threading
import time as _time
import webbrowser

logger = logging.getLogger(__name__)
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
    return _Path(__file__).resolve().parent.parent  # 项目根目录 (非 dcs_viewer/)

_EXE_DIR_CFG = _get_exe_dir()
_CONFIG_FILE = _EXE_DIR_CFG / "dcs_config.json"

_DEFAULTS_CFG = {
    "INFLUX_URL": "http://10.56.128.202:8086",
    "INFLUX_TOKEN": "",
    "INFLUX_ORG": "myOrg",
    "INFLUX_BUCKET": "islag",
    "INFLUX_TIMEOUT_MS": 180000,
    "FLASK_HOST": "0.0.0.0",
    "FLASK_PORT": 5000,
    "APP_TOKEN": "",
    "SETTINGS_PASSWORD": "admin123",
}

def _load_cfg():
    cfg = dict(_DEFAULTS_CFG)
    for k in _DEFAULTS_CFG:
        env_val = os.environ.get(k)
        if env_val is not None:
            cfg[k] = env_val
    # 首次运行 PyInstaller 打包版：自动从 _MEIPASS 释放 dcs_config.json 到 EXE 同目录
    if getattr(sys, 'frozen', False) and not _CONFIG_FILE.exists():
        import shutil as _shutil
        _bundle_cfg = _Path(sys._MEIPASS) / "dcs_config.json"
        if _bundle_cfg.exists():
            _shutil.copy(_bundle_cfg, _CONFIG_FILE)
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = _json.load(f)
            for k, v in saved.items():
                if k in cfg:
                    cfg[k] = v
        except Exception as e:
            logger.warning("配置文件解析失败，使用默认值: %s", e)
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
    # 白名单：仅允许写入已知的配置 key，防止注入攻击
    _ALLOWED_KEYS = set(_DEFAULTS_CFG.keys())
    current = {}
    if _CONFIG_FILE.exists():
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                current = _json.load(f)
        except Exception as e:
            logger.warning("配置文件读取失败，使用空配置: %s", e)
    current.update({k: v for k, v in data.items() if k in _ALLOWED_KEYS and v != "" and v is not None})
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

# === 从 dcs_platform 导入已验证的工具函数（避免重复定义） ===
try:
    from dcs_platform.core.config import (
        validate_param_name, validate_params,
        sanitize_param_for_flux, check_config,
    )
except ImportError:
    # PyInstaller 打包时 dcs_platform 可能未在路径中，保留 fallback
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

from dcs_platform.routes.rules_api import rules_bp
from dcs_platform.routes.knowledge_api import knowledge_bp
from dcs_platform.routes.training_api import training_bp
from dcs_platform.routes.analysis_api import analysis_bp
from dcs_platform.routes.labeling_api import labeling_bp
app.register_blueprint(rules_bp)
app.register_blueprint(knowledge_bp)
app.register_blueprint(training_bp)
app.register_blueprint(analysis_bp)
app.register_blueprint(labeling_bp)

# === PyInstaller 静态文件服务：优先从 _MEIPASS 加载，fallback 到文件系统 ===
from flask import send_from_directory as _send_from_directory

@app.route("/static/<path:filename>")
def _serve_static(filename):
    # 1) PyInstaller 打包：_MEIPASS 中查找
    if getattr(sys, 'frozen', False):
        _bundle_static = _Path(sys._MEIPASS) / "dcs_viewer" / "static"
        _fp = _bundle_static / filename
        if _fp.exists():
            return _send_from_directory(str(_bundle_static), filename)
    # 2) 开发模式：文件系统中的 static 目录
    for _root in [_EXE_DIR / "dcs_viewer" / "static", _BASE_DIR / "dcs_viewer" / "static",
                  _BASE_DIR / "static"]:
        _fp = _root / filename
        if _fp.exists():
            return _send_from_directory(str(_root), filename)
    return "File not found", 404


# === 调试端点（检查配置值，发布前可删除）===
@app.route("/debug-config")
def _debug_config():
    return jsonify({
        "SETTINGS_PASSWORD": repr(SETTINGS_PASSWORD),
        "match_admin123": SETTINGS_PASSWORD == "admin123",
        "APP_TOKEN": repr(APP_TOKEN),
    })


def _auto_seed_rules():
    try:
        from dcs_platform.core.db import _get_conn, upsert_rule_group, insert_rule
        conn = _get_conn()
        if conn.execute("SELECT COUNT(*) as n FROM rule_groups").fetchone()["n"] > 0: return
        PLUG=[("打泥超压报警","打泥压力>25MPa持续3s","AND",10,[("LT_LQFC_138","gt",25.0,3.0,">25MPa@3s")]),("打泥量过低","打泥量<0.5L","AND",7,[("LT_LQFC_179","lt",0.5,0,"<0.5L")]),("退炮压力异常","退炮压力>20MPa","AND",6,[("LT_LQFC_140","gt",20.0,1.0,">20MPa")]),("液压站温度告警","液压温度>55C","AND",2,[("LT_LQFC_150","gt",55.0,0,">55C")]),("液压站温度严重告警","液压温度>60C","AND",3,[("LT_LQFC_150","gt",60.0,0,">60C")]),("液压站压力异常","压力[70,130]bar外","OR",2,[("LT_LQFC_151","lt",70.0,0,"<70bar"),("LT_LQFC_151","gt",130.0,0,">130bar")])]
        OPEN=[("推进压力变化率异常","变化率>50MPa/s","AND",7,[("LT_LQFC_68","change_rate_gt",50.0,0,">50MPa/s")]),("冲击压力过高","冲击压力>18MPa","AND",6,[("LT_LQFC_88","gt",18.0,1.0,">18MPa")]),("转钎压力过高","转钎压力>15MPa","AND",6,[("LT_LQFC_87","gt",15.0,1.0,">15MPa")]),("出铁间隔过短","间隔<20min","AND",6,[("_prev.tapping_interval_min","lt",20.0,0,"<20min")]),("上次堵口打泥量偏高","打泥量>90L","AND",5,[("_prev.mud_quantity","gt",90.0,0,">90L")]),("液压站温度告警","液压温度>55C","AND",2,[("LT_LQFC_150","gt",55.0,0,">55C")]),("液压站温度严重告警","液压温度>60C","AND",3,[("LT_LQFC_150","gt",60.0,0,">60C")]),("液压站压力异常","压力[70,130]bar外","OR",2,[("LT_LQFC_151","lt",70.0,0,"<70bar"),("LT_LQFC_151","gt",130.0,0,">130bar")])]
        for nm,desc,logic,pri,rules in PLUG:
            gid=upsert_rule_group(None,"plugging",nm,desc,logic,pri)
            for pn,op,tv,dur,rn in rules: insert_rule(gid,pn,op,tv,rn,None,dur)
        for nm,desc,logic,pri,rules in OPEN:
            gid=upsert_rule_group(None,"opening",nm,desc,logic,pri)
            for pn,op,tv,dur,rn in rules: insert_rule(gid,pn,op,tv,rn,None,dur)
        conn.commit()
        print("[INFO] Auto-seeded 6 plugging + 8 opening rules")
    except Exception as e: print(f"[WARN] Auto-seed: {e}")
_auto_seed_rules()

# === 登录验证装饰器 ===
Compress(app)
app.config["COMPRESS_ALGORITHM"] = "gzip"
app.config["COMPRESS_MIN_SIZE"] = 500       # 小于500字节不压缩
app.config["COMPRESS_LEVEL"] = 6            # 压缩级别（1-9，6为平衡点）

# 北京时间 (UTC+8)
LOCAL_OFFSET = timedelta(hours=8)

def _parse_time(time_str):
    """弹性解析 datetime-local 输入，支持:
    - "2026-07-03T00:00"          (无秒数, 北京时间)
    - "2026-07-03T00:00:00"       (带秒数, 北京时间)
    - "2026-07-03T00:00:00.000"   (带毫秒, 北京时间)
    - "2026-07-03T00:00:00.000Z"  (带毫秒 + Z, UTC时间)
    - "2026-07-03T00:00:00Z"      (带秒数 + Z, UTC时间)
    返回 datetime (naive, 北京时间)
    """
    time_str = time_str.strip()
    is_utc = time_str.endswith("Z")
    if is_utc:
        time_str = time_str[:-1]  # 剥掉 Z
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            dt = datetime.strptime(time_str, fmt)
            if is_utc:
                dt = dt + LOCAL_OFFSET  # UTC → 北京时间
            return dt
        except ValueError:
            continue
    raise ValueError(f"无法解析时间: {time_str}")

# === API 访问认证 ===
@app.before_request
def _check_api_auth():
    """对所有 /api/ 路由做 Token 认证（登录/设置认证绕过）"""
    if not request.path.startswith("/api/"):
        return None
    if request.path in ("/api/login", "/api/settings/auth", "/api/debug/config"):
        return None  # 登录/设置认证不需要 API Token
    if not APP_TOKEN:
        return None
    token = request.args.get("token", "") or request.headers.get("X-API-Token", "")
    if token != APP_TOKEN:
        return jsonify({"error": "未授权访问 — 请提供有效的 API Token"}), 401
    return None

# === 加载参数分组 ===
# PyInstaller: --add-data "dcs_viewer/param_groups.json;dcs_viewer" → {_MEIPASS}/dcs_viewer/param_groups.json
GROUPS_FILE = None
for _candidate in [
    _BASE_DIR / "dcs_viewer" / "param_groups.json",
    _BASE_DIR / "param_groups.json",
    _EXE_DIR / "param_groups.json",
]:
    if _candidate.exists():
        GROUPS_FILE = _candidate
        break
if getattr(sys, 'frozen', False) and GROUPS_FILE is None:
    _bundle = _Path(sys._MEIPASS) / "dcs_viewer" / "param_groups.json"
    if _bundle.exists():
        GROUPS_FILE = _bundle
if GROUPS_FILE is None:
    raise FileNotFoundError("param_groups.json not found")
with open(GROUPS_FILE, "r", encoding="utf-8") as f:
    PARAM_CONFIG = json.load(f)

# === 性能优化：预映射标签，避免每个记录都查字典 ===
_LABELS = PARAM_CONFIG.get("labels", {})

def _load_html_template(filename):
    candidates = [
        _EXE_DIR / "dcs_viewer" / "templates" / filename,
        _EXE_DIR / "dcs_viewer" / filename,
        _BASE_DIR / "dcs_viewer" / "templates" / filename,
        _BASE_DIR / "dcs_viewer" / filename,
        _EXE_DIR / filename,
        _BASE_DIR / filename,
        Path(__file__).resolve().parent / filename,
    ]
    if getattr(sys, "frozen", False):
        candidates.insert(0, Path(sys._MEIPASS) / "dcs_viewer" / "templates" / filename)
        candidates.insert(1, Path(sys._MEIPASS) / "dcs_viewer" / filename)
    for p in candidates:
        if p.exists():
            with open(p, "r", encoding="utf-8") as fh: return fh.read()
    return None
RULES_OPENING_HTML = _load_html_template("rules_opening.html") or ""
RULES_PLUGGING_HTML = _load_html_template("rules_plugging.html") or ""
KNOWLEDGE_BASE_HTML = _load_html_template("knowledge_base.html") or ""

INDEX_HTML = _load_html_template("index.html") or ""


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
                    except Exception as e:
                        logger.warning("关闭旧 InfluxDB 客户端失败: %s", e)
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

# === 登录页面 (内嵌，避免 PyInstaller 文件加载缓存问题) ===
LOGIN_INLINE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>炉前开堵口数据分析平台</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;display:flex;align-items:center;justify-content:center}
.login-box{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,0.15);padding:40px;width:380px;max-width:95vw}
.login-box h2{font-size:22px;color:#1f2937;margin-bottom:6px}
.login-box .sub{font-size:12px;color:#94a3b8;margin-bottom:24px}
.form-group{margin-bottom:16px}
.form-group label{display:block;font-size:12px;color:#64748b;font-weight:600;margin-bottom:4px}
.form-group input{width:100%;padding:10px 14px;border:1px solid #e2e8f0;border-radius:8px;font-size:14px;outline:none;transition:border .2s}
.form-group input:focus{border-color:#667eea}
.btn-login{width:100%;padding:12px;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2);margin-top:8px}
.btn-login:hover{opacity:.9}
.msg{text-align:center;font-size:12px;margin-top:12px;min-height:20px}
.msg.err{color:#ef4444}
.actions{text-align:center;margin-top:16px}
.actions a{font-size:12px;color:#94a3b8;cursor:pointer;text-decoration:none}
.actions a:hover{color:#667eea}
</style>
</head>
<body>
<div class="login-box">
<h2>炉前开堵口数据分析平台</h2>
<div class="sub">高炉炉前开口机/堵口机专用</div>
<form onsubmit="doLogin(event)">
<div class="form-group"><label>账号</label><input id="username" value="admin" placeholder="admin"></div>
<div class="form-group"><label>密码</label><input id="password" type="password" placeholder="请输入密码"></div>
<button type="submit" class="btn-login">登 录</button>
</form>
<div class="msg" id="msg"></div>
</div>
<script>
async function doLogin(e){
    e.preventDefault();
    var u=document.getElementById('username').value.trim();
    var p=document.getElementById('password').value.trim();
    var r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,password:p})});
    var d=await r.json();
    if(d.ok){setTimeout(function(){window.location.href='/?token=""" + (APP_TOKEN or "") + r"""'},300)}
    else{console.log('LOGIN FAILED:',d);document.getElementById('msg').innerHTML=d.error;document.getElementById('msg').className='msg err'}
}
</script>
</body>
</html>"""

LOGIN_HTML = LOGIN_INLINE  # 始终使用内嵌版，避免 PyInstaller 缓存旧文件

@app.route("/")
def index():
    if not session.get("logged_in"):
        if request.args.get("token") == APP_TOKEN:
            session["logged_in"] = True
        else:
            return render_template_string(LOGIN_HTML.replace("{{ app_token }}", APP_TOKEN or ""))
    html = INDEX_HTML.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG["groups"]))
    html = html.replace("{{ labels_json | safe }}", json.dumps(_LABELS))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    return render_template_string(html)

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    username = data.get("username", "")
    password = data.get("password", "")
    pwd_correct = password == SETTINGS_PASSWORD
    user_correct = username == "admin"
    if user_correct and pwd_correct:
        logger.info("登录成功: user=%s", username)
        session["logged_in"] = True
        return jsonify({"ok": True, "redirect": "/"})
    # 登录失败
    logger.warning("登录失败: user=%s pwd_len=%d user_ok=%s pwd_ok=%s",
                   username, len(password), user_correct, pwd_correct)
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
        except Exception as e2:
            logger.warning("关闭 Excel workbook 失败: %s", e2)
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


REALTIME_HTML = _load_html_template("realtime.html") or ""


TREND_HTML = _load_html_template("trend.html") or ""


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
        s_local = _parse_time(start)
        e_local = _parse_time(end)
    except ValueError as e:
        return jsonify({"error": f"时间格式错误: {e}"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")

    # 计算数据量，自动降采样
    time_diff_hours = (e_local - s_local).total_seconds() / 3600
    if time_diff_hours <= 0.5:
        window = "1s"
    elif time_diff_hours <= 2:
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
  |> aggregateWindow(every: {window}, fn: last, createEmpty: false)'''

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

        return jsonify({
            "series": result_series,
            "window": window,
            "query_range": {"start": s_utc, "end": e_utc}
        })
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
        s_local = _parse_time(start)
        e_local = _parse_time(end)
    except ValueError as e:
        return jsonify({"error": f"时间格式错误: {e}"}), 400

    s_utc = (s_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")
    e_utc = (e_local - LOCAL_OFFSET).strftime("%Y-%m-%dT%H:%M:%SZ")

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




# === 信号配置: 设备状态识别 ===
_STATE_SIGNALS = {
    "east_opener": {
        "name": "东开口机",
        "remote_select": "LT_LQFC_57",
        "swing_cmd": "LT_LQFC_59",
        "cart_cmd": "LT_LQFC_61",
        "impact_cmd": "LT_LQFC_69",
        "swing_pos": "LT_LQFC_63",
        "push_pos": "LT_LQFC_67",
    },
    "west_opener": {
        "name": "西开口机",
        "remote_select": "LT_LQFC_94",
        "swing_cmd": "LT_LQFC_96",
        "cart_cmd": "LT_LQFC_98",
        "impact_cmd": "LT_LQFC_106",
        "swing_pos": "LT_LQFC_100",
        "push_pos": "LT_LQFC_104",
    },
    "east_plugger": {
        "name": "东堵口机",
        "remote_power": "LT_LQFC_129",
        "remote_start": "LT_LQFC_130",
        "emergency_stop": "LT_LQFC_142",
        "set_mode": "LT_LQFC_141",
        "wait_pos": "LT_LQFC_143",
        "work_pos": "LT_LQFC_144",
        "hydraulic_temp": "LT_LQFC_150",
        "hydraulic_press": "LT_LQFC_151",
    },
    "west_plugger": {
        "name": "西堵口机",
        "remote_power": "LT_LQFC_152",
        "remote_start": "LT_LQFC_153",
        "emergency_stop": "LT_LQFC_165",
        "set_mode": "LT_LQFC_164",
        "wait_pos": "LT_LQFC_166",
        "work_pos": "LT_LQFC_167",
        "hydraulic_temp": "LT_LQFC_173",
        "hydraulic_press": "LT_LQFC_174",
    },
}


def _detect_states(raw_data, sig, machine_id):
    states = []

    if 'remote_start' in sig:
        # Plugger state detection
        remote_power = sorted(raw_data.get(sig['remote_power'], []), key=lambda x: x[0])
        remote_start = sorted(raw_data.get(sig['remote_start'], []), key=lambda x: x[0])
        emergency = sorted(raw_data.get(sig['emergency_stop'], []), key=lambda x: x[0])
        set_mode = sorted(raw_data.get(sig['set_mode'], []), key=lambda x: x[0])
        hydraulic_temp = sorted(raw_data.get(sig['hydraulic_temp'], []), key=lambda x: x[0])

        if not remote_power:
            return states

        rp_map = {}
        for t, v in remote_power: rp_map[int(t.timestamp())] = v
        rs_map = {}
        for t, v in remote_start: rs_map[int(t.timestamp())] = v
        em_map = {}
        for t, v in emergency: em_map[int(t.timestamp())] = v
        sm_map = {}
        for t, v in set_mode: sm_map[int(t.timestamp())] = v
        temp_map = {}
        for t, v in hydraulic_temp: temp_map[int(t.timestamp())] = v

        all_ts = sorted(set(list(rp_map.keys()) + list(rs_map.keys()) + list(em_map.keys())))

        prev_state = None
        for ts in all_ts:
            rp = rp_map.get(ts, rp_map.get(ts - 1, rp_map.get(ts + 1, 0)))
            rs = rs_map.get(ts, rs_map.get(ts - 1, rs_map.get(ts + 1, 0)))
            em = em_map.get(ts, em_map.get(ts - 1, em_map.get(ts + 1, 0)))
            sm = sm_map.get(ts, sm_map.get(ts - 1, sm_map.get(ts + 1, 0)))
            ht = temp_map.get(ts, temp_map.get(ts - 1, temp_map.get(ts + 1, 50)))

            if sm >= 0.5 and rs < 0.5:
                state = 'maintenance'
            elif em >= 0.5 or ht > 60:
                state = 'fault'
            elif rs >= 0.5 and em < 0.5:
                state = 'working'
            elif rp >= 0.5 and rs < 0.5 and em < 0.5:
                state = 'standby'
            else:
                state = 'offline'

            if state != prev_state:
                t_local = datetime.fromtimestamp(ts, tz=timezone.utc) + LOCAL_OFFSET
                states.append({
                    'machine': sig['name'],
                    'machine_id': machine_id,
                    'time': t_local.isoformat(),
                    'state': state,
                    'label': {'standby': u'\u5f85\u673a', 'working': u'\u5de5\u4f5c\u4e2d', 'fault': u'\u6545\u969c', 'maintenance': u'\u7ef4\u62a4', 'offline': u'\u79bb\u7ebf'}[state],
                })
                prev_state = state
    else:
        # Opener state detection (simplified)
        remote_select = sorted(raw_data.get(sig['remote_select'], []), key=lambda x: x[0])
        swing_cmd = sorted(raw_data.get(sig['swing_cmd'], []), key=lambda x: x[0])
        cart_cmd = sorted(raw_data.get(sig['cart_cmd'], []), key=lambda x: x[0])
        impact_cmd = sorted(raw_data.get(sig['impact_cmd'], []), key=lambda x: x[0])

        if not remote_select:
            return states

        rs_map = {}
        for t, v in remote_select: rs_map[int(t.timestamp())] = v
        sc_map = {}
        for t, v in swing_cmd: sc_map[int(t.timestamp())] = v
        cc_map = {}
        for t, v in cart_cmd: cc_map[int(t.timestamp())] = v
        ic_map = {}
        for t, v in impact_cmd: ic_map[int(t.timestamp())] = v

        all_ts = sorted(set(list(rs_map.keys()) + list(sc_map.keys())))

        prev_state = None
        for ts in all_ts:
            rs = rs_map.get(ts, rs_map.get(ts - 1, rs_map.get(ts + 1, 0)))
            sc = sc_map.get(ts, sc_map.get(ts - 1, sc_map.get(ts + 1, 0)))
            cc = cc_map.get(ts, cc_map.get(ts - 1, cc_map.get(ts + 1, 0)))
            ic = ic_map.get(ts, ic_map.get(ts - 1, ic_map.get(ts + 1, 0)))

            if rs >= 0.5 and sc < 0.5 and cc < 0.5 and ic < 0.5:
                state = 'standby'
            elif rs >= 0.5 and (sc >= 0.5 or cc >= 0.5 or ic >= 0.5):
                state = 'working'
            else:
                state = 'offline'

            if state != prev_state:
                t_local = datetime.fromtimestamp(ts, tz=timezone.utc) + LOCAL_OFFSET
                states.append({
                    'machine': sig['name'],
                    'machine_id': machine_id,
                    'time': t_local.isoformat(),
                    'state': state,
                    'label': {'standby': u'\u5f85\u673a', 'working': u'\u5de5\u4f5c\u4e2d', 'offline': u'\u79bb\u7ebf'}[state],
                })
                prev_state = state

    return states

# ============================================================
# === 智能化模型: GBDT-MPC-PID 三级协同架构 (数据分析层) ===
# ============================================================

def _classify_opening_phases(push_pos_data, push_press_data, drill_press_data,
                               impact_press_data, t_start, t_end):
    """开口工艺阶段分类: approach/initial_drill/steady_drill/breakthrough/retract"""
    phases = []

    seg_push_pos = sorted([(t, v) for t, v in push_pos_data if t_start <= t <= t_end], key=lambda x: x[0])
    seg_push_press = sorted([(t, v) for t, v in push_press_data if t_start <= t <= t_end], key=lambda x: x[0])
    seg_drill = sorted([(t, v) for t, v in drill_press_data if t_start <= t <= t_end], key=lambda x: x[0])
    seg_impact = sorted([(t, v) for t, v in impact_press_data if t_start <= t <= t_end], key=lambda x: x[0])

    if not seg_push_pos:
        return [{"phase": "unknown", "start": t_start.isoformat(), "end": t_end.isoformat(),
                 "duration_s": (t_end - t_start).total_seconds()}]

    # Phase 1: Approach (push position < 0.01 change)
    approach_end = t_start
    for i in range(1, len(seg_push_pos)):
        if seg_push_pos[i][1] - seg_push_pos[0][1] > 0.01:
            approach_end = seg_push_pos[i][0]
            break
    if approach_end > t_start:
        phases.append({"phase": "approach", "label": "接近",
                        "start": (t_start + LOCAL_OFFSET).isoformat(),
                        "end": (approach_end + LOCAL_OFFSET).isoformat(),
                        "duration_s": round((approach_end - t_start).total_seconds(), 1)})

    # Phase 2: Initial drilling (push advancing, drill pressure < 2)
    init_start = approach_end
    init_end = init_start
    for t, v in seg_drill:
        if t > init_start and v > 2:
            init_end = t
            break
    if init_start < (seg_push_pos[-1][0] if seg_push_pos else t_end):
        if init_end <= init_start:
            init_end = (seg_push_pos[-1][0] if seg_push_pos else t_end)
        phases.append({"phase": "initial_drill", "label": "初钻",
                        "start": (init_start + LOCAL_OFFSET).isoformat(),
                        "end": (init_end + LOCAL_OFFSET).isoformat(),
                        "duration_s": round((init_end - init_start).total_seconds(), 1)})

    # Phase 3: Steady drilling (drill pressure sustained, push advancing)
    steady_start = init_end
    steady_end = t_end

    # Detect breakthrough: push position increase + pressure drop
    pos_dict = {int(t.timestamp()): v for t, v in seg_push_pos}
    for j in range(3, min(len(seg_push_press), len(seg_push_pos))):
        t_curr_ts = int(seg_push_press[j][0].timestamp())
        dp_curr = pos_dict.get(t_curr_ts, 0)
        dp_old = pos_dict.get(int(seg_push_press[j - 3][0].timestamp()), 0)
        delta_pos = dp_curr - dp_old
        p_ratio = (seg_push_press[j][1] - seg_push_press[j - 3][1]) / seg_push_press[j - 3][1] if seg_push_press[j - 3][1] > 0 else 0
        if delta_pos > 0.1 and p_ratio < -0.2:
            steady_end = seg_push_press[j][0]
            # Breakthrough phase
            if steady_end > steady_start:
                phases.append({"phase": "steady_drill", "label": "稳态钻进",
                                "start": (steady_start + LOCAL_OFFSET).isoformat(),
                                "end": (steady_end + LOCAL_OFFSET).isoformat(),
                                "duration_s": round((steady_end - steady_start).total_seconds(), 1)})
                phases.append({"phase": "breakthrough", "label": "钻透突破",
                                "start": (steady_end + LOCAL_OFFSET).isoformat(),
                                "end": (t_end + LOCAL_OFFSET).isoformat(),
                                "duration_s": round((t_end - steady_end).total_seconds(), 1)})
                break
    else:
        if steady_start < t_end:
            phases.append({"phase": "steady_drill", "label": "稳态钻进",
                            "start": (steady_start + LOCAL_OFFSET).isoformat(),
                            "end": (t_end + LOCAL_OFFSET).isoformat(),
                            "duration_s": round((t_end - steady_start).total_seconds(), 1)})

    return phases


def _classify_plugging_phases(mud_pos_data, mud_press_data, mud_qty_data, swing_data, t_start, t_end):
    """堵口工艺阶段分类: swing_to_position/mud_fill/pressure_hold/swing_back"""
    phases = []

    seg_mud_press = sorted([(t, v) for t, v in mud_press_data if t_start <= t <= t_end], key=lambda x: x[0])
    seg_mud_qty = sorted([(t, v) for t, v in mud_qty_data if t_start <= t <= t_end], key=lambda x: x[0])

    if not seg_mud_press:
        return [{"phase": "unknown", "start": t_start.isoformat(), "end": t_end.isoformat(),
                 "duration_s": (t_end - t_start).total_seconds()}]

    # Phase 1: Mud filling (pressure building, quantity increasing)
    fill_end = t_start
    for t, v in seg_mud_press:
        if v > 15:
            fill_end = t
            break
    if fill_end > t_start:
        mud_start_qty = seg_mud_qty[0][1] if seg_mud_qty else 0
        mud_end_qty = next((v for t, v in seg_mud_qty if t >= fill_end), mud_start_qty)
        phases.append({"phase": "mud_fill", "label": "打泥填充",
                        "start": (t_start + LOCAL_OFFSET).isoformat(),
                        "end": (fill_end + LOCAL_OFFSET).isoformat(),
                        "duration_s": round((fill_end - t_start).total_seconds(), 1),
                        "mud_volume": round(mud_end_qty - mud_start_qty, 1)})

    # Phase 2: Pressure hold (pressure 18-22 range)
    hold_start = fill_end
    hold_end = hold_start
    hold_count = 0
    for t, v in seg_mud_press:
        if t <= hold_start:
            continue
        if 18 <= v <= 22:
            hold_count += 1
            hold_end = t
        else:
            if hold_count >= 60:
                break
            hold_count = 0
    if hold_end > hold_start:
        phases.append({"phase": "pressure_hold", "label": "保压",
                        "start": (hold_start + LOCAL_OFFSET).isoformat(),
                        "end": (hold_end + LOCAL_OFFSET).isoformat(),
                        "duration_s": round((hold_end - hold_start).total_seconds(), 1),
                        "hold_count_s": hold_count})

    # Phase 3: Retraction
    retract_start = hold_end
    if retract_start < t_end:
        phases.append({"phase": "retract", "label": "回退",
                        "start": (retract_start + LOCAL_OFFSET).isoformat(),
                        "end": (t_end + LOCAL_OFFSET).isoformat(),
                        "duration_s": round((t_end - retract_start).total_seconds(), 1)})

    return phases


def _score_opening_quality(cycle, phases):
    """开口质量评分 (0-100)"""
    score = 100.0
    details = []

    # Duration check (normal range: 5-20 min)
    dur_min = cycle.get("duration_s", 0) / 60
    if dur_min < 3:
        score -= 20
        details.append("耗时过短(异常)")
    elif dur_min > 25:
        score -= 20
        details.append("耗时过长")
    elif dur_min > 18:
        score -= 10
        details.append("耗时偏长")

    # Push position change (should be > 0.05 for meaningful drilling)
    push_change = cycle.get("push_pos_change", 0)
    if push_change < 0.01:
        score -= 25
        details.append("推进量不足")
    elif push_change < 0.03:
        score -= 10
        details.append("推进量偏小")

    # Breakthrough detection
    if cycle.get("breakthrough"):
        details.append("已钻透")
    else:
        score -= 15
        details.append("未检测到钻透")

    # Push pressure peak (abnormal if > 25 MPa)
    push_peak = cycle.get("push_press_peak", 0)
    if push_peak > 25:
        score -= 10
        details.append(f"推进压力过高({push_peak:.0f}MPa)")

    # Drill pressure peak (abnormal if > 30 MPa)
    drill_peak = cycle.get("drill_press_peak", 0)
    if drill_peak > 30:
        score -= 10
        details.append(f"钻压过高({drill_peak:.0f}MPa)")

    # Phase completeness
    phase_names = [p["phase"] for p in phases]
    expected = ["approach", "initial_drill", "steady_drill"]
    for p in expected:
        if p not in phase_names:
            score -= 8
            details.append(f"缺少{p}阶段")

    score = max(0, score)
    return {"score": round(score, 1), "grade": "优" if score >= 85 else ("良" if score >= 70 else ("中" if score >= 50 else "差")),
            "details": details}


def _score_plugging_quality(cycle, phases):
    """堵口质量评分 (0-100)"""
    score = 100.0
    details = []

    # Duration check
    dur_min = cycle.get("duration_s", 0) / 60
    if dur_min < 2:
        score -= 20
        details.append("耗时过短(异常)")
    elif dur_min > 20:
        score -= 15
        details.append("耗时过长")

    # Mud quantity (should be >= 10)
    mud_qty = cycle.get("mud_qty", 0)
    if mud_qty < 5:
        score -= 25
        details.append("打泥量严重不足")
    elif mud_qty < 10:
        score -= 10
        details.append("打泥量偏少")

    # Pressure hold
    if cycle.get("hold_ok"):
        details.append(f"保压OK({cycle.get('hold_duration_s', 0):.0f}s)")
    else:
        score -= 20
        details.append("保压不充分")

    # Mud pressure peak
    mud_peak = cycle.get("mud_press_peak", 0)
    if mud_peak < 10:
        score -= 10
        details.append("打泥压力过低")
    elif mud_peak > 30:
        score -= 10
        details.append("打泥压力过高")

    # Phase completeness
    phase_names = [p["phase"] for p in phases]
    if "mud_fill" not in phase_names:
        score -= 10
        details.append("缺少打泥阶段")
    if "pressure_hold" not in phase_names:
        score -= 10
        details.append("缺少保压阶段")

    score = max(0, score)
    return {"score": round(score, 1), "grade": "优" if score >= 85 else ("良" if score >= 70 else ("中" if score >= 50 else "差")),
            "details": details}


def _detect_anomalies(cycle, phases, raw_press_data, raw_temp_data):
    """异常检测: 压力/温度/时序异常"""
    anomalies = []

    # Pressure spike detection
    press_vals = [v for _, v in raw_press_data]
    if press_vals:
        avg = sum(press_vals) / len(press_vals)
        std = (sum((v - avg) ** 2 for v in press_vals) / len(press_vals)) ** 0.5
        for _, v in raw_press_data:
            if std > 0 and abs(v - avg) > 4 * std:
                anomalies.append("压力剧烈波动")
                break

    # Temperature trend
    if raw_temp_data and len(raw_temp_data) >= 2:
        temps = [v for _, v in raw_temp_data]
        temp_rise = temps[-1] - temps[0]
        if temp_rise > 5:
            anomalies.append(f"温度快速上升({temp_rise:.1f}°C)")

    # Duration anomaly
    if cycle["type"] == "opening":
        if cycle.get("duration_s", 0) > 1800:  # > 30 min
            anomalies.append("开口超时(>30分钟)")
    else:
        if cycle.get("duration_s", 0) > 2400:  # > 40 min
            anomalies.append("堵口超时(>40分钟)")

    # Result check
    if cycle.get("result") == "fail":
        anomalies.append("作业失败")
    elif cycle.get("result") == "incomplete":
        anomalies.append("作业未完成")

    return anomalies


def _recommend_parameters(cycle, all_cycles_history, phases):
    """基于历史数据启发式推荐最优参数 (模拟GBDT决策)"""
    recs = []

    hist_same_type = [c for c in all_cycles_history if c["type"] == cycle["type"] and
                      c["machine"] == cycle.get("machine", "") and c.get("result") == "success"]
    if not hist_same_type:
        hist_same_type = [c for c in all_cycles_history if c["type"] == cycle["type"] and c.get("result") == "success"]

    if cycle["type"] == "opening":
        # Recommend push speed based on successful cycles
        if hist_same_type:
            push_changes = [c.get("push_pos_change", 0) for c in hist_same_type if c.get("push_pos_change", 0) > 0.01]
            durations = [c.get("duration_s", 600) for c in hist_same_type if c.get("duration_s", 0) > 60]
            if push_changes:
                avg_change = sum(push_changes) / len(push_changes)
                recs.append({"param": "target_push_depth", "value": round(avg_change, 3),
                             "unit": "m", "reason": f"历史成功均值({len(push_changes)}次)"})
            if durations:
                avg_dur = sum(durations) / len(durations)
                recs.append({"param": "expected_duration", "value": round(avg_dur / 60, 1),
                             "unit": "min", "reason": f"历史成功均值({len(durations)}次)"})
        if not recs:
            recs.append({"param": "target_push_depth", "value": 0.08, "unit": "m", "reason": "默认推荐值"})

        # Drill pressure recommendation
        if hist_same_type:
            drill_peaks = [c.get("drill_press_peak", 0) for c in hist_same_type if c.get("drill_press_peak", 0) > 0]
            if drill_peaks:
                recs.append({"param": "drill_press_limit", "value": round(sum(drill_peaks) / len(drill_peaks) * 1.2, 1),
                             "unit": "MPa", "reason": "成功峰值1.2x安全边际"})

    else:  # plugging
        if hist_same_type:
            mud_qtys = [c.get("mud_qty", 0) for c in hist_same_type if c.get("mud_qty", 0) > 5]
            hold_durs = [c.get("hold_duration_s", 0) for c in hist_same_type if c.get("hold_duration_s", 0) > 30]
            if mud_qtys:
                avg_qty = sum(mud_qtys) / len(mud_qtys)
                recs.append({"param": "target_mud_qty", "value": round(avg_qty, 1),
                             "unit": "", "reason": f"历史成功均值({len(mud_qtys)}次)"})
            if hold_durs:
                avg_hold = sum(hold_durs) / len(hold_durs)
                recs.append({"param": "hold_duration", "value": round(avg_hold, 0),
                             "unit": "s", "reason": f"历史成功均值({len(hold_durs)}次)"})
        if not recs:
            recs.append({"param": "target_mud_qty", "value": 12.0, "unit": "", "reason": "默认推荐值"})
            recs.append({"param": "hold_duration", "value": 60, "unit": "s", "reason": "默认60秒保压"})

    return recs



def _fetch_cycle_data(start_utc, end_utc, params, window, timeout_ms=30000):
    """查询指定参数的时间序列数据，返回 {param: [(utc_dt, value), ...]}

    Args:
        timeout_ms: InfluxDB 查询超时，默认 30s（分析专用短超时）
    """
    param_filter = sanitize_param_for_flux(params)
    client = InfluxDBClient(
        url=INFLUX_URL, token=INFLUX_TOKEN, org=INFLUX_ORG,
        timeout=timeout_ms
    )
    try:
        query_api = client.query_api()
        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: {start_utc}, stop: {end_utc})
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> filter(fn: (r) => r._field == "value")
  |> filter(fn: (r) => {param_filter})
  |> aggregateWindow(every: {window}, fn: last, createEmpty: false)'''

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
    except Exception as e:
        print(f"[WARN] _fetch_cycle_data failed (timeout={timeout_ms}ms): {e}", flush=True)
        return {}
    finally:
        client.close()


def _detect_opening_cycles(raw_data, sig):
    """检测开口作业周期 — O(n) 高效版本"""
    pos = sorted(raw_data.get(sig["swing_pos"], []), key=lambda x: x[0])
    push_pos = sorted(raw_data.get(sig["push_pos"], []), key=lambda x: x[0])
    push_press = sorted(raw_data.get(sig["push_press"], []), key=lambda x: x[0])
    remote = sorted(raw_data.get(sig["remote"], []), key=lambda x: x[0])
    drill = sorted(raw_data.get(sig["drill_press"], []), key=lambda x: x[0])

    if len(pos) < 2 or len(remote) < 2:
        return []

    rem_map = {int(t.timestamp()): v for t, v in remote}
    cycles = []

    for i in range(1, len(pos)):
        prev_v, curr_v = pos[i - 1][1], pos[i][1]
        if not (prev_v < 90 and curr_v >= 90):
            continue
        t_cross = pos[i][0]
        t_cross_ts = int(t_cross.timestamp())

        remote_on = rem_map.get(t_cross_ts, 0) >= 0.5
        if not remote_on:
            for offset in (-1, 1):
                if rem_map.get(t_cross_ts + offset, 0) >= 0.5:
                    remote_on = True
                    break
        if not remote_on:
            continue

        t_start = t_cross
        # 先用较大窗口扫描数据，再检测实际结束时间
        t_scan = t_start + timedelta(minutes=30)

        push_pos_change = 0.0
        push_press_peak = 0.0
        drill_press_peak = 0.0
        breakthrough_detected = False

        # Calculate push position change (扫描 30min 窗口找实际结束)
        pos_before = [v for t, v in push_pos if t < t_start]
        pos_after = [(t, v) for t, v in push_pos if t_start <= t <= t_scan]
        if pos_after:
            ref = pos_before[-1] if pos_before else pos_after[0][1]
            push_pos_change = pos_after[-1][1] - ref

        # 检测实际结束时间：推进位移最大值之后开始回退
        t_end = t_start + timedelta(minutes=15)  # 默认15分钟
        if len(pos_after) >= 2:
            max_idx = max(range(len(pos_after)), key=lambda i: pos_after[i][1])
            t_end = pos_after[max_idx][0]
            # 从最大值之后继续扫描，找位移开始下降的点
            for j in range(max_idx + 1, len(pos_after)):
                if pos_after[j][1] < pos_after[j-1][1] - 0.01:
                    t_end = pos_after[j][0]
                    break
                t_end = pos_after[j][0]

        # Calculate pressure peaks and breakthrough detection
        if push_press:
            press_seg = [(t, v) for t, v in push_press if t_start <= t <= t_end]
            if press_seg:
                push_press_peak = max(v for _, v in press_seg)
                # Efficient breakthrough: sliding window of 3 consecutive samples
                pos_dict = {int(t.timestamp()): v for t, v in push_pos if t_start <= t <= t_end}
                for j in range(3, len(press_seg)):
                    p_curr = press_seg[j][1]
                    p_old = press_seg[j - 3][1]
                    t_curr_ts = int(press_seg[j][0].timestamp())
                    t_old_ts = int(press_seg[j - 3][0].timestamp())
                    dp_curr = pos_dict.get(t_curr_ts, pos_dict.get(t_curr_ts + 1, pos_dict.get(t_curr_ts - 1, 0)))
                    dp_old = pos_dict.get(t_old_ts, pos_dict.get(t_old_ts + 1, pos_dict.get(t_old_ts - 1, 0)))
                    delta_pos = dp_curr - dp_old
                    delta_press_ratio = (p_curr - p_old) / p_old if p_old > 0 else 0
                    if delta_pos > 0.1 and delta_press_ratio < -0.2:
                        breakthrough_detected = True
                        break

        if drill:
            drill_seg = [v for t, v in drill if t_start <= t <= t_end]
            if drill_seg:
                drill_press_peak = max(drill_seg)

        result = "success" if breakthrough_detected else ("incomplete" if push_pos_change > 0.01 else "fail")

        t_local_start = t_start + LOCAL_OFFSET
        t_local_end = t_end + LOCAL_OFFSET

        cycles.append({
            "machine": sig["name"],
            "type": "opening",
            "trigger_time": t_local_start.isoformat(),
            "window_start": t_local_start.isoformat(),
            "window_end": t_local_end.isoformat(),
            "duration_s": round((t_end - t_start).total_seconds(), 1),
            "push_pos_change": round(push_pos_change, 3),
            "push_press_peak": round(push_press_peak, 1),
            "drill_press_peak": round(drill_press_peak, 1),
            "breakthrough": breakthrough_detected,
            "result": result,
            "label": f"{sig['name']} 开口 {'成' if breakthrough_detected else '未'}钻透 {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
        })

    return cycles

def _detect_plugging_cycles(raw_data, sig):
    """检测堵口作业周期 — 基于泥炮选择信号触发"""
    plug_cmd = sorted(raw_data.get(sig["plug_select"], []), key=lambda x: x[0])
    mud_press = sorted(raw_data.get(sig["mud_press"], []), key=lambda x: x[0])
    mud_qty = sorted(raw_data.get(sig["mud_qty"], []), key=lambda x: x[0])

    if len(plug_cmd) < 2:
        return []

    cycles = []

    for i in range(1, len(plug_cmd)):
        prev_v, curr_v = plug_cmd[i - 1][1], plug_cmd[i][1]
        if not (prev_v < 0.5 and curr_v >= 0.5):
            continue
        t_start = plug_cmd[i][0]

        t_end = t_start + timedelta(minutes=40)
        mud_press_peak = 0.0
        mud_qty_total = 0.0
        hold_duration_s = 0.0
        mud_fill_complete = False
        hold_complete = False

        press_seg = sorted([(t, v) for t, v in mud_press if t_start <= t <= t_end], key=lambda x: x[0])
        if press_seg:
            mud_press_peak = max(v for _, v in press_seg)
            hold_count = 0
            for _, v in press_seg:
                if 18 <= v <= 22:
                    hold_count += 1
                    if hold_count >= 60:
                        hold_duration_s = float(hold_count)
                        hold_complete = True
                        break
                else:
                    hold_count = 0

        qty_seg = sorted([(t, v) for t, v in mud_qty if t_start <= t <= t_end], key=lambda x: x[0])
        if qty_seg:
            mud_qty_total = qty_seg[-1][1]
            mud_fill_complete = mud_qty_total >= 10

        # Find retreat end
        for t, v in reversed(press_seg):
            if v < 5 and t > t_start + timedelta(seconds=30):
                t_end = t
                break

        result = "success" if (mud_fill_complete and hold_complete) else ("partial" if mud_fill_complete else "fail")

        t_local_start = t_start + LOCAL_OFFSET
        t_local_end = t_end + LOCAL_OFFSET

        cycles.append({
            "machine": sig["name"],
            "type": "plugging",
            "trigger_time": t_local_start.isoformat(),
            "window_start": t_local_start.isoformat(),
            "window_end": t_local_end.isoformat(),
            "duration_s": round((t_end - t_start).total_seconds(), 1),
            "mud_press_peak": round(mud_press_peak, 1),
            "mud_qty": round(mud_qty_total, 1),
            "hold_duration_s": round(hold_duration_s, 0),
            "mud_filled": mud_fill_complete,
            "hold_ok": hold_complete,
            "result": result,
            "label": f"{sig['name']} 堵口 {'完' if mud_fill_complete else '未完'} {t_local_start.strftime('%H:%M:%S')} ~ {t_local_end.strftime('%H:%M:%S')}",
        })

    return cycles


# === 作业分析页面 ===
# ==============================================================

ANALYSIS_HTML = _load_html_template("analysis.html") or ""
@app.route("/rules/opening")
@login_required
def rules_opening():
    html = (RULES_OPENING_HTML or "").replace("{{ app_token }}", APP_TOKEN or "").replace("{{app_token}}", APP_TOKEN or "")
    if not html: return "Page not found", 404
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/rules/plugging")
@login_required
def rules_plugging():
    html = (RULES_PLUGGING_HTML or "").replace("{{ app_token }}", APP_TOKEN or "").replace("{{app_token}}", APP_TOKEN or "")
    if not html: return "Page not found", 404
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/knowledge-base")
@login_required
def knowledge_base():
    html = (KNOWLEDGE_BASE_HTML or "").replace("{{ app_token }}", APP_TOKEN or "").replace("{{app_token}}", APP_TOKEN or "")
    if not html: return "Page not found", 404
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/labeling")
@login_required
def labeling():
    html = _load_html_template("labeling.html") or ""
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/report")
@login_required
def report():
    html = _load_html_template("report.html") or ""
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/detect-config")
@login_required
def detect_config():
    html = _load_html_template("detect_config.html") or ""
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/variable-config")
@login_required
def variable_config():
    html = _load_html_template("variable_config.html") or ""
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

@app.route("/analysis")
@login_required
def analysis():
    html = _load_html_template("analysis.html") or _load_html_template("analysis_new.html") or ""
    html = html.replace("{{ groups_json | safe }}", json.dumps(PARAM_CONFIG))
    html = html.replace("{{ app_token }}", APP_TOKEN or "")
    resp = make_response(render_template_string(html))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


# === 系统配置页面 ===
SETTINGS_HTML = _load_html_template("settings.html") or ""

@app.route("/settings")
def settings_page():
    return render_template_string(SETTINGS_HTML.replace("{{ app_token }}", APP_TOKEN or ""))

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
    print(f"  [DEBUG] SETTINGS_PASSWORD={repr(SETTINGS_PASSWORD)} APP_TOKEN={repr(APP_TOKEN)}")
    print(f"  趋势分析:   http://localhost:{FLASK_PORT}/trend")
    print(f"  作业分析:   http://localhost:{FLASK_PORT}/analysis")
    print(f"  循环标注:   http://localhost:{FLASK_PORT}/labeling")
    print(f"  检测规则:   http://localhost:{FLASK_PORT}/detect-config")
    print(f"  API 认证:   {'已启用 (APP_TOKEN)' if APP_TOKEN else '[WARN] 未启用 (仅内网安全)'}")
    print()

    # 启动自整定调度器
    try:
        from dcs_platform.services.self_tuning import start_scheduler
        start_scheduler()
        print("  [Tuning] 自整定调度器已启动")
    except Exception as e:
        print(f"  [Tuning] 调度器启动失败: {e}")

    # 自动打开浏览器
    if not getattr(sys, 'frozen', False):
        webbrowser.open(f"http://localhost:{FLASK_PORT}")
    else:
        # PyInstaller 环境，等待 Flask 完全启动后再打开
        def _open_browser():
            _time.sleep(1.5)
            webbrowser.open(f"http://localhost:{FLASK_PORT}")
        threading.Thread(target=_open_browser, daemon=True).start()
    try:
        from waitress import serve
        serve(app, host=FLASK_HOST, port=FLASK_PORT, threads=8)
    except ImportError:
        print("  提示: 未安装 waitress，使用 Flask 内置开发服务器")
        app.run(host=FLASK_HOST, port=FLASK_PORT, debug=False, threaded=True)
