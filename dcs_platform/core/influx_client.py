"""
DCS 平台核心模块 — InfluxDB 客户端封装

提供线程安全的 InfluxDBClient 单例、查询缓存、通用查询函数。
"""
import logging
import threading
import time as _time
from typing import Optional

from influxdb_client import InfluxDBClient

from dcs_platform.core.config import (
    INFLUX_URL, INFLUX_TOKEN, INFLUX_ORG, INFLUX_BUCKET,
    INFLUX_MEASUREMENT, INFLUX_TIMEOUT_MS, sanitize_param_for_flux,
)

logger = logging.getLogger(__name__)

# ==================== InfluxDBClient 单例 ====================

_client_lock = threading.Lock()
_influx_client: Optional[InfluxDBClient] = None
_client_created_at: float = 0
_CLIENT_MAX_AGE: float = 300  # 5分钟后自动重建连接


def get_client() -> InfluxDBClient:
    """返回全局复用的 InfluxDBClient 单例。

    首次访问时懒初始化，超过 CLIENT_MAX_AGE 自动重建。
    线程安全（双重检查锁）。
    """
    global _influx_client, _client_created_at
    now = _time.time()
    if _influx_client is None or (now - _client_created_at) > _CLIENT_MAX_AGE:
        with _client_lock:
            now = _time.time()  # 重新获取时间，防止锁等待期间过期
            if _influx_client is None or (now - _client_created_at) > _CLIENT_MAX_AGE:
                if _influx_client is not None:
                    try:
                        _influx_client.close()
                    except Exception as e:
                        logger.debug("关闭旧 InfluxDB 客户端失败: %s", e)
                _influx_client = InfluxDBClient(
                    url=INFLUX_URL, token=INFLUX_TOKEN, org=INFLUX_ORG,
                    timeout=INFLUX_TIMEOUT_MS
                )
                _client_created_at = _time.time()
    return _influx_client


def reset_client():
    """强制关闭并重建 InfluxDBClient（配置变更后调用）"""
    global _influx_client, _client_created_at
    with _client_lock:
        if _influx_client is not None:
            try:
                _influx_client.close()
            except Exception as e:
                logger.debug("关闭旧 InfluxDB 客户端失败: %s", e)
        _influx_client = None
        _client_created_at = 0


# ==================== 查询结果缓存 ====================

_cache: dict = {}
_cache_lock = threading.Lock()
CACHE_TTL: float = 30       # 默认缓存30秒
MAX_CACHE_ROWS: int = 5000  # 超过此行数不缓存


def _cache_key(route: str, **kwargs) -> str:
    raw = f"{route}|" + "|".join(f"{k}={v}" for k, v in sorted(kwargs.items()))
    return raw


def _cache_get(key: str):
    """读缓存，过期返回 None"""
    with _cache_lock:
        entry = _cache.get(key)
        if entry and _time.time() - entry["ts"] < CACHE_TTL:
            return entry["data"]
        if entry:
            _cache.pop(key, None)
        return None


def _cache_set(key: str, data, nrows: int = 0):
    """写缓存。超过 MAX_CACHE_ROWS 不缓存"""
    if nrows > MAX_CACHE_ROWS:
        return
    with _cache_lock:
        _cache[key] = {"data": data, "ts": _time.time()}
        if len(_cache) > 100:
            oldest = min(_cache, key=lambda k: _cache[k]["ts"])
            _cache.pop(oldest, None)


def clear_cache():
    """清空所有缓存"""
    with _cache_lock:
        _cache.clear()


# ==================== 通用查询函数 ====================

def fetch_timeseries(
    start_utc: str,
    end_utc: str,
    params: list[str],
    window: str = "auto",
    timeout_ms: int = 30000,
) -> dict:
    """查询指定参数的时序数据。

    Args:
        start_utc: 开始时间（ISO 8601 UTC 格式）
        end_utc: 结束时间
        params: 参数名列表
        window: 聚合窗口，如 "1s", "500ms"，默认 "auto" 根据时间范围自动选择
        timeout_ms: 查询超时毫秒

    Returns:
        {param: [(datetime, float), ...]}
    """
    # 自动选择聚合窗口
    if window == "auto":
        try:
            from datetime import datetime
            s = datetime.fromisoformat(start_utc.replace("Z", "+00:00"))
            e = datetime.fromisoformat(end_utc.replace("Z", "+00:00"))
            hours = max(0.5, (e - s).total_seconds() / 3600)
            if hours <= 2:
                window = "10s"
            elif hours <= 8:
                window = "30s"
            elif hours <= 24:
                window = "1m"
            else:
                window = "5m"
        except Exception:
            window = "30s"  # fallback

    param_filter = sanitize_param_for_flux(params)
    client = get_client()
    try:
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
    except Exception as e:
        logger.warning("fetch_timeseries failed (timeout=%sms): %s", timeout_ms, e)
        return {}


def query_raw_flux(flux_query: str) -> dict:
    """执行原始 Flux 查询并返回结构化结果。

    Returns:
        {param: [(datetime, float), ...]}
    """
    client = get_client()
    try:
        query_api = client.query_api()
        tables = query_api.query(flux_query)
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
        logger.warning("query_raw_flux failed: %s", e)
        return {}


def ping() -> bool:
    """测试 InfluxDB 连通性"""
    try:
        client = get_client()
        query_api = client.query_api()
        flux = f'''from(bucket: "{INFLUX_BUCKET}")
  |> range(start: -1m)
  |> filter(fn: (r) => r._measurement == "{INFLUX_MEASUREMENT}")
  |> limit(n: 1)'''
        tables = query_api.query(flux)
        return len(tables) > 0
    except Exception:
        return False
