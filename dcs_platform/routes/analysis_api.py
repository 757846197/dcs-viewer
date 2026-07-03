"""DCS 平台 — 分析 API 路由

提供 Phase 3-6 全部分析功能的 HTTP API 端点。
挂载在 Flask app 下，复用现有 @app.before_request 的 Token 认证。
"""
import io
import json
import logging
from datetime import datetime, timedelta, timezone

import xlsxwriter
from flask import Blueprint, request, jsonify, send_file

from dcs_platform.core.influx_client import fetch_timeseries, ping
from dcs_platform.core.config import sanitize_param_for_flux
from dcs_platform.core.db import get_cycles, get_label_stats, export_labels, insert_cycle
from dcs_platform.services.group_service import (
    get_group_params, get_param_label, get_all_groups, get_group_by_id,
)
from dcs_platform.services.analysis.state_identifier import (
    compute_state_timeline, compute_state_stats, map_east_plugger_params,
    map_west_plugger_params, map_east_opener_params, STATE_LABELS,
)
from dcs_platform.services.analysis.process_analyzer import (
    analyze_equipment, generate_summary_report,
)

logger = logging.getLogger(__name__)

analysis_bp = Blueprint("analysis", __name__, url_prefix="/api/analysis")


@analysis_bp.route("/ping")
def api_ping():
    ok = ping()
    return jsonify({
        "status": "ok" if ok else "error",
        "influxdb": ok,
        "reachable": ok,
        "message": "InfluxDB 连接正常" if ok else "InfluxDB 无响应，请检查网络或 InfluxDB 服务状态"
    })


@analysis_bp.route("/equipment")
def api_equipment_list():
    groups = get_all_groups()
    return jsonify([{
        "id": g["id"], "name": g["name"],
        "param_count": len(g["params"]),
        "categories": list(g.get("categories", {}).keys()),
    } for g in groups])


@analysis_bp.route("/equipment/<equipment_id>/overview")
def api_equipment_overview(equipment_id):
    params, group = get_group_params(equipment_id)
    if not params:
        return jsonify({"error": f"Unknown equipment: {equipment_id}"}), 404

    # Get recent data snapshot (last hour)
    now = datetime.now(timezone.utc)
    start = (now - timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    data = fetch_timeseries(start, end, params[:10], "30s", timeout_ms=15000)

    signals = {}
    for p in params[:10]:
        label = get_param_label(p)
        vals = data.get(p, [])
        if vals:
            latest = vals[-1]
            signals[p] = {
                "label": label, "latest_value": round(latest[1], 3),
                "latest_time": latest[0].isoformat(),
                "sample_count": len(vals),
            }
        else:
            signals[p] = {"label": label, "status": "no_data"}

    return jsonify({
        "equipment_id": equipment_id,
        "name": group["name"] if group else equipment_id,
        "total_params": len(params),
        "signals": signals,
    })


@analysis_bp.route("/state-timeline")
def api_state_timeline():
    equipment_id = request.args.get("equipment_id", "east_plugger")
    hours = int(request.args.get("hours", 24))

    now = datetime.now(timezone.utc)
    start = (now - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    mapping = map_east_plugger_params()
    param_names = [v for v in mapping.values() if v]

    data = fetch_timeseries(start, end, param_names, "30s", timeout_ms=30000)

    timeline = compute_state_timeline(data, mapping)
    stats = compute_state_stats(timeline)

    return jsonify({
        "equipment_id": equipment_id,
        "hours": hours,
        "timeline": timeline,
        "stats": stats,
        "state_labels": STATE_LABELS,
    })


@analysis_bp.route("/cycles")
def api_cycles():
    """GET /api/analysis/cycles?start=...&end=...&type=...&token=...
    
    Detect operation cycles from InfluxDB data in the given time range.
    Uses remote command + position crossing threshold for opening,
    remote_start + mud_cmd edge for plugging.
    Supports type=opening|plugging|all.
    """
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    cycle_type = request.args.get("type", "all")
    limit = int(request.args.get("limit", 200))
    
    # Normalize time format: frontend sends "2026-07-03T00:00" (no seconds/Z)
    # Flux requires RFC3339: "2026-07-03T00:00:00Z"
    start = _normalize_time(start)
    end = _normalize_time(end)
    
    # ── 正确信号配置 ──
    # 开口机: remote(binary 0/1) + swing_pos(穿越90°阈值)
    # 堵口机: remote_start(binary 0/1) + mud_cmd(边沿)
    OPENING_CONFIG = {
        "东开口机": {"remote": "LT_LQFC_57", "swing_pos": "LT_LQFC_63",
                      "push_pos": "LT_LQFC_67", "push_press": "LT_LQFC_68",
                      "drill_press": "LT_LQFC_69", "impact_press": "LT_LQFC_59"},
        "西开口机": {"remote": "LT_LQFC_94", "swing_pos": "LT_LQFC_100",
                      "push_pos": "LT_LQFC_104", "push_press": "LT_LQFC_105",
                      "drill_press": "LT_LQFC_124", "impact_press": "LT_LQFC_125"},
    }
    PLUGGING_CONFIG = {
        "东堵口机": {"remote_start": "LT_LQFC_130", "mud_cmd": "LT_LQFC_134",
                      "mud_pos": "LT_LQFC_137", "mud_press": "LT_LQFC_138",
                      "mud_qty": "LT_LQFC_179", "swing_pos": "LT_LQFC_135"},
        "西堵口机": {"remote_start": "LT_LQFC_153", "mud_cmd": "LT_LQFC_157",
                      "mud_pos": "LT_LQFC_160", "mud_press": "LT_LQFC_161",
                      "mud_qty": "LT_LQFC_180", "swing_pos": "LT_LQFC_158"},
    }
    
    cycles = []
    
    # ━━━ 开口检测: remote==1 AND swing_pos 穿越 90° ━━━
    if cycle_type in ("opening", "all"):
        for machine, sig in OPENING_CONFIG.items():
            # 只查检测必需的两个信号，加快速度
            detect_params = [sig["remote"], sig["swing_pos"]]
            try:
                data = fetch_timeseries(start, end, detect_params, timeout_ms=60000)
            except Exception:
                continue
            
            remote = _build_time_map(data.get(sig["remote"], []))
            swing = sorted(data.get(sig["swing_pos"], []), key=lambda x: x[0])
            
            if len(swing) < 2 or len(remote) < 2:
                continue
            
            in_cycle = False
            cycle_start = None
            for i in range(1, len(swing)):
                prev_v, curr_v = swing[i-1][1], swing[i][1]
                t = swing[i][0]
                ts = t.timestamp()
                
                if not in_cycle and _crossed(prev_v, curr_v, 90):
                    # 检查遥控信号 (±2s 容差)
                    if _remote_nearby(remote, ts, tolerance=2):
                        in_cycle = True
                        cycle_start = t
                    
                elif in_cycle and _crossed(curr_v, prev_v, 90):
                    # 回转退出 90° = 周期结束
                    duration = (t - cycle_start).total_seconds()
                    if 30 <= duration <= 3600:
                        cycles.append({
                            "machine": machine, "type": "opening",
                            "trigger_time": cycle_start.isoformat(),
                            "window_start": cycle_start.isoformat(),
                            "window_end": t.isoformat(),
                            "duration_s": round(duration, 1),
                            "result": "unknown", "breakthrough": False,
                        })
                    in_cycle = False
                    if len(cycles) >= limit: break
            
            # 未闭合的周期（仍然在作业中）
            if in_cycle and swing:
                last_t = swing[-1][0]
                duration = (last_t - cycle_start).total_seconds()
                if 30 <= duration <= 3600:
                    cycles.append({
                        "machine": machine, "type": "opening",
                        "trigger_time": cycle_start.isoformat(),
                        "window_start": cycle_start.isoformat(),
                        "window_end": last_t.isoformat(),
                        "duration_s": round(duration, 1),
                        "result": "unknown", "breakthrough": False,
                    })
    
    # ━━━ 堵口检测: remote_start==1 AND mud_cmd 边沿 ━━━
    if cycle_type in ("plugging", "all"):
        for machine, sig in PLUGGING_CONFIG.items():
            detect_params = [sig["remote_start"], sig["mud_cmd"]]
            try:
                data = fetch_timeseries(start, end, detect_params, timeout_ms=60000)
            except Exception:
                continue
            
            remote_start = _build_time_map(data.get(sig["remote_start"], []))
            mud_cmd = sorted(data.get(sig["mud_cmd"], []), key=lambda x: x[0])
            
            if len(mud_cmd) < 2 or len(remote_start) < 2:
                continue
            
            in_cycle = False
            cycle_start = None
            for i in range(1, len(mud_cmd)):
                prev_v, curr_v = mud_cmd[i-1][1], mud_cmd[i][1]
                t = mud_cmd[i][0]
                ts = t.timestamp()
                
                # Rising edge on mud_cmd (0→1)
                if not in_cycle and prev_v < 0.5 and curr_v >= 0.5:
                    if _remote_nearby(remote_start, ts, tolerance=2):
                        in_cycle = True
                        cycle_start = t
                
                # Falling edge (1→0) = cycle end
                elif in_cycle and prev_v >= 0.5 and curr_v < 0.5:
                    duration = (t - cycle_start).total_seconds()
                    if 30 <= duration <= 3600:
                        cycles.append({
                            "machine": machine, "type": "plugging",
                            "trigger_time": cycle_start.isoformat(),
                            "window_start": cycle_start.isoformat(),
                            "window_end": t.isoformat(),
                            "duration_s": round(duration, 1),
                            "result": "unknown", "hold_ok": False,
                        })
                    in_cycle = False
                    if len(cycles) >= limit: break
            
            if in_cycle and mud_cmd:
                last_t = mud_cmd[-1][0]
                duration = (last_t - cycle_start).total_seconds()
                if 30 <= duration <= 3600:
                    cycles.append({
                        "machine": machine, "type": "plugging",
                        "trigger_time": cycle_start.isoformat(),
                        "window_start": cycle_start.isoformat(),
                        "window_end": last_t.isoformat(),
                        "duration_s": round(duration, 1),
                        "result": "unknown", "hold_ok": False,
                    })
    
    cycles.sort(key=lambda c: c["trigger_time"])
    
    # ── 去重并存入数据库 ──
    saved_count = 0
    if cycles:
        _saved = set()
        try:
            existing = get_cycles(limit=50000)
            _saved = {(c["equipment_id"], c["cycle_type"], c["start_time"]) for c in existing}
        except Exception:
            pass
        for c in cycles:
            key = (c["machine"], c["type"], c["window_start"])
            if key not in _saved:
                try:
                    insert_cycle(c["machine"], c["type"], c["window_start"],
                                 c["window_end"], c["duration_s"], 0.8)
                    _saved.add(key)
                    saved_count += 1
                except Exception:
                    pass
    
    return jsonify({
        "cycles": cycles[:limit],
        "count": len(cycles),
        "saved_to_db": saved_count
    })


def _normalize_time(time_str):
    """Normalize time to RFC3339 UTC: '2026-07-03T00:00' -> '2026-07-03T00:00:00Z'"""
    if not time_str:
        return time_str
    # Already has seconds + timezone
    if "Z" in time_str or "+" in time_str[-6:]:
        return time_str
    # Has seconds but no timezone: add Z
    if len(time_str) >= 19 and time_str[10] == 'T':
        return time_str + "Z"
    # No seconds: "2026-07-03T00:00" -> "2026-07-03T00:00:00Z"
    if "T" in time_str:
        return time_str + ":00Z"
    return time_str


def _build_time_map(pairs):
    """将 [(datetime, float), ...] 转为 {timestamp: value} 字典"""
    return {t.timestamp(): v for t, v in pairs}

def _crossed(prev_val, curr_val, threshold):
    """检查是否穿越阈值"""
    return (prev_val < threshold and curr_val >= threshold) or \
           (prev_val > threshold and curr_val <= threshold)

def _remote_nearby(rem_map, target_ts, tolerance=2):
    """检查 target_ts ± tolerance 秒内遥控信号是否为 1"""
    for offset in range(-tolerance, tolerance + 1):
        if rem_map.get(target_ts + offset, 0) >= 0.5:
            return True
    return False


@analysis_bp.route("/pressure-position")
def api_pressure_position():
    equipment_id = request.args.get("equipment_id", "east_plugger")
    hours = int(request.args.get("hours", 24))

    now = datetime.now(timezone.utc)
    start = (now - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    if "plugger" in equipment_id:
        pressure_param = "LT_LQFC_138" if "east" in equipment_id else "LT_LQFC_161"
        position_param = "LT_LQFC_137" if "east" in equipment_id else "LT_LQFC_160"
    else:
        pressure_param = "LT_LQFC_68" if "east" in equipment_id else "LT_LQFC_105"
        position_param = "LT_LQFC_67" if "east" in equipment_id else "LT_LQFC_104"

    data = fetch_timeseries(start, end, [pressure_param, position_param], "1s", timeout_ms=30000)

    pressure_vals = data.get(pressure_param, [])
    position_vals = data.get(position_param, [])

    # Align by time and build scatter data
    import bisect
    pos_times = [t for t, _ in position_vals]
    scatter = []
    for t, p in pressure_vals:
        idx = bisect.bisect_right(pos_times, t) - 1
        if idx >= 0:
            dt = (t - pos_times[idx]).total_seconds()
            if abs(dt) < 2:
                scatter.append({
                    "time": t.isoformat(),
                    "position": round(position_vals[idx][1], 3),
                    "pressure": round(p, 3),
                })

    # Compute envelope
    if scatter:
        positions = [s["position"] for s in scatter]
        pressures = [s["pressure"] for s in scatter]
        mean_p = sum(pressures) / len(pressures)
        std_p = (sum((v - mean_p) ** 2 for v in pressures) / len(pressures)) ** 0.5

        return jsonify({
            "equipment_id": equipment_id,
            "pressure_param": pressure_param,
            "position_param": position_param,
            "position_label": get_param_label(position_param),
            "pressure_label": get_param_label(pressure_param),
            "scatter": scatter[:2000],
            "envelope": {"mean": round(mean_p, 3), "upper": round(mean_p + 2 * std_p, 3),
                         "lower": round(mean_p - 2 * std_p, 3)},
            "count": len(scatter),
        })

    return jsonify({"equipment_id": equipment_id, "scatter": [], "count": 0,
                    "message": "No data for pressure-position analysis in this period"})


@analysis_bp.route("/health-metrics")
def api_health_metrics():
    equipment_id = request.args.get("equipment_id", "east_plugger")
    hours = int(request.args.get("hours", 168))  # default 7 days

    now = datetime.now(timezone.utc)
    start = (now - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Get temperature and pressure params
    if "east" in equipment_id:
        temp_p = "LT_LQFC_150"
        press_p = "LT_LQFC_151"
    else:
        temp_p = "LT_LQFC_173"
        press_p = "LT_LQFC_174"

    data = fetch_timeseries(start, end, [temp_p, press_p], "60s", timeout_ms=30000)

    temp_vals = [v for _, v in data.get(temp_p, [])]
    press_vals = [v for _, v in data.get(press_p, [])]

    result = {"equipment_id": equipment_id, "hours": hours}

    if temp_vals:
        mean_t = sum(temp_vals) / len(temp_vals)
        max_t = max(temp_vals)
        temp_slope = 0
        if len(temp_vals) > 1:
            n = len(temp_vals)
            x_mean = (n - 1) / 2
            y_mean = mean_t
            num = sum((i - x_mean) * (v - y_mean) for i, v in enumerate(temp_vals))
            den = sum((i - x_mean) ** 2 for i in range(n))
            temp_slope = num / den if den > 0 else 0

        result["temperature"] = {
            "mean": round(mean_t, 1), "max": round(max_t, 1),
            "slope_per_sample": round(temp_slope, 6),
            "alerts": {
                "warn": max_t > 55, "critical": max_t > 60,
                "trending_up": temp_slope > 0.001,
            },
        }

    if press_vals:
        mean_p = sum(press_vals) / len(press_vals)
        std_p = (sum((v - mean_p) ** 2 for v in press_vals) / len(press_vals)) ** 0.5
        volatility = std_p / mean_p if mean_p > 0 else 0
        result["pressure"] = {
            "mean": round(mean_p, 3), "std": round(std_p, 3),
            "volatility": round(volatility, 4),
            "min": round(min(press_vals), 3), "max": round(max(press_vals), 3),
        }

    # Health score
    score = 100
    if temp_vals:
        if any(v > 60 for v in temp_vals):
            score -= 30
        elif any(v > 55 for v in temp_vals):
            score -= 15
        if abs(temp_slope) > 0.01:
            score -= 10
    if press_vals and volatility > 0.5:
        score -= 20

    result["health_score"] = max(0, min(100, round(score)))
    result["grade"] = "A" if score >= 90 else "B" if score >= 75 else "C" if score >= 60 else "D"

    return jsonify(result)


@analysis_bp.route("/report")
def api_report():
    equipment_id = request.args.get("equipment_id")
    eids = [equipment_id] if equipment_id else None
    report = generate_summary_report(eids)
    report["labeling_progress"] = get_label_stats()
    return jsonify(report)


@analysis_bp.route("/labeling-stats")
def api_labeling_stats():
    return jsonify(get_label_stats())


@analysis_bp.route("/labeling-export")
def api_labeling_export():
    cycle_type = request.args.get("type")
    labels = export_labels(cycle_type)
    return jsonify({"labels": labels, "count": len(labels)})


# === Helper: extract metrics from a cycle data window ===

def _extract_cycle_metrics(
    data: dict, start_time, end_time, is_plugging: bool, equipment_id: str
) -> dict:
    """Extract key metrics from raw signal data within a cycle window."""
    metrics = {}
    
    if is_plugging:
        press_key = "LT_LQFC_138" if "east" in equipment_id else "LT_LQFC_161"
        mud_key = "LT_LQFC_179" if "east" in equipment_id else "LT_LQFC_180"
        hold_key = "LT_LQFC_134" if "east" in equipment_id else "LT_LQFC_157"
        
        press_vals = [v for t, v in data.get(press_key, []) if start_time <= t <= end_time]
        mud_vals = [v for t, v in data.get(mud_key, []) if start_time <= t <= end_time]
        hold_vals = [v for t, v in data.get(hold_key, []) if start_time <= t <= end_time]
        
        if mud_vals:
            metrics["mud_qty"] = round(max(mud_vals) - min(mud_vals), 2)
        else:
            metrics["mud_qty"] = 0
        if press_vals:
            metrics["mud_press_peak"] = round(max(press_vals), 1)
            metrics["mud_press_mean"] = round(sum(press_vals) / len(press_vals), 1)
        else:
            metrics["mud_press_peak"] = 0
            metrics["mud_press_mean"] = 0
        
        # Hold duration: count high signal after peak
        hold_ok = False
        if hold_vals:
            hold_count = sum(1 for v in hold_vals if v >= 0.5)
            hold_duration = hold_count  # 1 sample ≈ 1 second
            metrics["hold_duration_s"] = hold_duration
            metrics["hold_ok"] = hold_duration >= 60
            hold_ok = hold_duration >= 60
        else:
            metrics["hold_duration_s"] = 0
            metrics["hold_ok"] = False
        
        # Result determination
        if metrics["mud_qty"] > 0.1 and hold_ok:
            metrics["result"] = "success"
        elif metrics["mud_qty"] > 0.1:
            metrics["result"] = "partial"
        else:
            metrics["result"] = "fail"
        metrics["breakthrough"] = hold_ok
    else:
        pos_key = "LT_LQFC_67" if "east" in equipment_id else "LT_LQFC_104"
        press_key = "LT_LQFC_68" if "east" in equipment_id else "LT_LQFC_105"
        drill_key = "LT_LQFC_87" if "east" in equipment_id else "LT_LQFC_124"
        impact_key = "LT_LQFC_69" if "east" in equipment_id else "LT_LQFC_106"
        
        pos_vals = [v for t, v in data.get(pos_key, []) if start_time <= t <= end_time]
        press_vals = [v for t, v in data.get(press_key, []) if start_time <= t <= end_time]
        drill_vals = [v for t, v in data.get(drill_key, []) if start_time <= t <= end_time]
        impact_vals = [v for t, v in data.get(impact_key, []) if start_time <= t <= end_time]
        
        if pos_vals:
            metrics["push_depth"] = round(max(pos_vals) - min(pos_vals), 3)
            metrics["push_pos_change"] = round(max(pos_vals) - min(pos_vals), 3)
        else:
            metrics["push_depth"] = 0
            metrics["push_pos_change"] = 0
        if press_vals:
            metrics["push_press_max"] = round(max(press_vals), 1)
            metrics["push_press_mean"] = round(sum(press_vals) / len(press_vals), 1)
            metrics["push_press_peak"] = round(max(press_vals), 1)
        else:
            metrics["push_press_max"] = 0
            metrics["push_press_mean"] = 0
            metrics["push_press_peak"] = 0
        if drill_vals:
            metrics["drill_press_max"] = round(max(drill_vals), 1)
            metrics["drill_press_mean"] = round(sum(drill_vals) / len(drill_vals), 1)
        else:
            metrics["drill_press_max"] = 0
            metrics["drill_press_mean"] = 0
        
        impact_active = any(v >= 0.5 for v in impact_vals) if impact_vals else False
        metrics["impact_press_active"] = impact_active
        
        # Breakthrough: depth > 2m AND pressure dropped > 20%
        breakthrough = False
        if len(pos_vals) > 5 and len(press_vals) > 5:
            depth_change = max(pos_vals) - min(pos_vals)
            early_press = sum(press_vals[:len(press_vals)//3]) / max(1, len(press_vals)//3)
            late_press = sum(press_vals[2*len(press_vals)//3:]) / max(1, len(press_vals) - 2*len(press_vals)//3)
            if depth_change > 2.0 and early_press > 0 and (early_press - late_press) / early_press > 0.2:
                breakthrough = True
        
        metrics["breakthrough"] = breakthrough
        if breakthrough:
            metrics["result"] = "success"
        elif metrics["push_depth"] > 0.1:
            metrics["result"] = "incomplete"
        else:
            metrics["result"] = "fail"
    
    return metrics


# === Endpoint 1: Multi-machine state timeline ===

@analysis_bp.route("/states")
def api_states():
    """GET /api/analysis/states?start=...&end=...&machine=all&token=xxx

    Return per-timestamp state snapshots for one or all machines.
    Uses the existing compute_state_timeline and equipment mapping functions.
    """
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    machine = request.args.get("machine", "all")

    # Mapping: (equipment_id, machine_label, mapping_function_or_dict)
    machine_configs = [
        ("east_plugger", "东堵口机", map_east_plugger_params()),
        ("west_plugger", "西堵口机", map_west_plugger_params()),
        ("east_opener", "东开口机", map_east_opener_params()),
        # West opener uses an inline mapping (no dedicated function yet)
        ("west_opener", "西开口机", {
            "remote_power": "LT_LQFC_94",
            "remote_start": "LT_LQFC_97",
            "emergency_stop": "LT_LQFC_128",
            "set_spare": None,
            "hydraulic_temp": "LT_LQFC_127",
        }),
    ]

    if machine != "all":
        machine_configs = [c for c in machine_configs if c[1] == machine]
        if not machine_configs:
            machine_configs = [c for c in machine_configs if c[0] == machine]
        if not machine_configs:
            return jsonify({"error": f"Unknown machine: {machine}"}), 400

    all_states = []

    for eq_id, label, mapping in machine_configs:
        # Collect all param names from mapping
        param_names = [v for v in mapping.values() if v]
        if not param_names:
            continue

        try:
            data = fetch_timeseries(start, end, param_names, "30s", timeout_ms=30000)
        except Exception:
            continue

        timeline = compute_state_timeline(data, mapping)

        # Convert timeline segments into per-30s-interval snapshots
        # For each segment, emit one snapshot at segment start (simplified format)
        for seg in timeline:
            all_states.append({
                "machine": label,
                "time": seg["start_time"],
                "state": seg["state"],
                "label": seg["label"],
            })
        # Also emit a final snapshot for the last segment end
        if timeline:
            last = timeline[-1]
            all_states.append({
                "machine": label,
                "time": last["end_time"],
                "state": last["state"],
                "label": last["label"],
            })

    all_states.sort(key=lambda s: s["time"])
    return jsonify({"states": all_states})


# === Endpoint 2: Windowed metrics ===

@analysis_bp.route("/metrics")
def api_metrics():
    """GET /api/analysis/metrics?window_start=...&window_end=...&machine=...&type=...&token=xxx

    Query InfluxDB for the given time window and extract key operational
    indicators using _extract_cycle_metrics.
    """
    window_start = request.args.get("window_start", "")
    window_end = request.args.get("window_end", "")
    machine = request.args.get("machine", "东堵口机")
    metrics_type = request.args.get("type", "opening")

    # Resolve machine label to equipment_id
    machine_map = {
        "东堵口机": "east_plugger",
        "西堵口机": "west_plugger",
        "东开口机": "east_opener",
        "西开口机": "west_opener",
    }
    equipment_id = machine_map.get(machine, "east_plugger")
    is_plugging = metrics_type == "plugging"

    # Build param list for the equipment + type
    if is_plugging:
        if "east" in equipment_id:
            params = ["LT_LQFC_138", "LT_LQFC_179", "LT_LQFC_134"]
        else:
            params = ["LT_LQFC_161", "LT_LQFC_180", "LT_LQFC_157"]
    else:
        if "east" in equipment_id:
            params = ["LT_LQFC_67", "LT_LQFC_68", "LT_LQFC_87", "LT_LQFC_69"]
        else:
            params = ["LT_LQFC_104", "LT_LQFC_105", "LT_LQFC_124", "LT_LQFC_106"]

    try:
        data = fetch_timeseries(window_start, window_end, params, "1s", timeout_ms=30000)
    except Exception as e:
        return jsonify({"error": f"InfluxDB query failed: {e}"}), 500

    # Parse the window bounds
    try:
        st = datetime.fromisoformat(window_start.replace("Z", "+00:00"))
        et = datetime.fromisoformat(window_end.replace("Z", "+00:00"))
    except Exception:
        # Fallback: use the data range
        all_vals = [t for p in params for t, _ in data.get(p, [])]
        st = min(all_vals) if all_vals else datetime.now(timezone.utc)
        et = max(all_vals) if all_vals else datetime.now(timezone.utc)

    metrics = _extract_cycle_metrics(data, st, et, is_plugging, equipment_id)

    # Build the response in the shape the frontend expects
    result = {"type": metrics_type}
    if is_plugging:
        result.update({
            "push_depth": None,
            "push_press_max": None,
            "push_press_mean": None,
            "drill_press_max": None,
            "drill_press_mean": None,
            "breakthrough": metrics.get("breakthrough", False),
            "impact_press_active": None,
            "mud_qty": metrics.get("mud_qty", 0),
            "mud_press_max": metrics.get("mud_press_peak", 0),
            "hold_duration_s": metrics.get("hold_duration_s", 0),
        })
    else:
        result.update({
            "push_depth": metrics.get("push_depth", 0),
            "push_press_max": metrics.get("push_press_max", 0),
            "push_press_mean": metrics.get("push_press_mean", 0),
            "drill_press_max": metrics.get("drill_press_max", 0),
            "drill_press_mean": metrics.get("drill_press_mean", 0),
            "breakthrough": metrics.get("breakthrough", False),
            "impact_press_active": metrics.get("impact_press_active", False),
            "mud_qty": None,
            "mud_press_max": None,
            "hold_duration_s": None,
        })

    return jsonify(result)


# === Endpoint 3: Excel export ===

@analysis_bp.route("/export")
def api_export():
    """GET /api/analysis/export?start=...&end=...&type=...&token=xxx

    Generate an xlsx report for all cycles in the time range and return
    it as an attachment.
    """
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    cycle_type = request.args.get("type", "all")

    # Reuse cycle detection logic (same as /cycles endpoint)
    equipment_configs = [
        ("east_opener", "东开口机", "LT_LQFC_61", False),
        ("west_opener", "西开口机", "LT_LQFC_98", False),
        ("east_plugger", "东堵口机", "LT_LQFC_134", True),
        ("west_plugger", "西堵口机", "LT_LQFC_157", True),
    ]

    cycles = []
    for eq_id, machine_label, trigger_signal, is_plugging in equipment_configs:
        if cycle_type != "all" and (
            (cycle_type == "opening" and is_plugging) or
            (cycle_type == "plugging" and not is_plugging)
        ):
            continue

        if is_plugging:
            params = ["LT_LQFC_134", "LT_LQFC_133", "LT_LQFC_135", "LT_LQFC_137",
                       "LT_LQFC_138", "LT_LQFC_179"] if "east" in eq_id else \
                      ["LT_LQFC_157", "LT_LQFC_156", "LT_LQFC_158", "LT_LQFC_160",
                       "LT_LQFC_161", "LT_LQFC_180"]
        else:
            params = ["LT_LQFC_61", "LT_LQFC_63", "LT_LQFC_67", "LT_LQFC_68",
                       "LT_LQFC_69", "LT_LQFC_59"] if "east" in eq_id else \
                      ["LT_LQFC_98", "LT_LQFC_100", "LT_LQFC_104", "LT_LQFC_105",
                       "LT_LQFC_106", "LT_LQFC_96"]

        try:
            data = fetch_timeseries(start, end, params, "1s", timeout_ms=60000)
        except Exception:
            continue

        trigger_data = data.get(trigger_signal, [])
        if len(trigger_data) < 10:
            continue

        in_cycle = False
        cycle_start_time = None
        min_duration_s = 30
        max_duration_s = 3600

        for i, (t, v) in enumerate(trigger_data):
            if i == 0:
                continue
            prev_v = trigger_data[i - 1][1]

            if not in_cycle and prev_v < 0.5 and v >= 0.5:
                in_cycle = True
                cycle_start_time = t
            elif in_cycle and prev_v >= 0.5 and v < 0.5:
                cycle_end_time = t
                duration = (cycle_end_time - cycle_start_time).total_seconds()

                if min_duration_s <= duration <= max_duration_s:
                    metrics = _extract_cycle_metrics(
                        data, cycle_start_time, cycle_end_time, is_plugging, eq_id
                    )
                    cycles.append({
                        "machine": machine_label,
                        "type": "plugging" if is_plugging else "opening",
                        "trigger_time": cycle_start_time.isoformat(),
                        "duration_s": round(duration, 1),
                        "result": metrics.get("result", "unknown"),
                        **metrics,
                    })

                in_cycle = False
                if len(cycles) >= 2000:
                    break

        if len(cycles) >= 2000:
            break

    cycles.sort(key=lambda c: c["trigger_time"])

    # Build xlsx in memory
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    sheet = workbook.add_worksheet("循环分析结果")

    # Styles
    header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#1d4ed8", "font_color": "#ffffff",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    cell_fmt = workbook.add_format({"border": 1, "align": "center"})
    success_fmt = workbook.add_format({
        "border": 1, "align": "center", "bg_color": "#dcfce7",
    })
    fail_fmt = workbook.add_format({
        "border": 1, "align": "center", "bg_color": "#fee2e2",
    })

    # Headers
    headers = [
        "设备", "类型", "触发时间", "耗时(s)", "结果",
        "推进深度(m)", "推进压力峰值", "钻孔压力峰值",
        "打通", "冲击压力激活", "泥浆量", "泥浆压力峰值", "保压时长(s)",
    ]
    for col, h in enumerate(headers):
        sheet.write(0, col, h, header_fmt)

    # Data rows
    for row_idx, c in enumerate(cycles, start=1):
        is_success = c.get("result") == "success"
        fmt = success_fmt if is_success else fail_fmt if c.get("result") == "fail" else cell_fmt

        sheet.write(row_idx, 0, c.get("machine", ""), fmt)
        sheet.write(row_idx, 1, c.get("type", ""), fmt)
        sheet.write(row_idx, 2, c.get("trigger_time", ""), fmt)
        sheet.write(row_idx, 3, c.get("duration_s", ""), fmt)
        sheet.write(row_idx, 4, c.get("result", ""), fmt)
        sheet.write(row_idx, 5, c.get("push_depth") or c.get("push_pos_change", ""), fmt)
        sheet.write(row_idx, 6, c.get("push_press_peak") or c.get("push_press_max", ""), fmt)
        sheet.write(row_idx, 7, c.get("drill_press_max", ""), fmt)
        sheet.write(row_idx, 8, "是" if c.get("breakthrough") else "否", fmt)
        sheet.write(row_idx, 9, "是" if c.get("impact_press_active") else "-", fmt)
        sheet.write(row_idx, 10, c.get("mud_qty", ""), fmt)
        sheet.write(row_idx, 11, c.get("mud_press_peak", ""), fmt)
        sheet.write(row_idx, 12, c.get("hold_duration_s", ""), fmt)

    # Column widths
    sheet.set_column(0, 0, 12)   # 设备
    sheet.set_column(1, 1, 8)    # 类型
    sheet.set_column(2, 2, 22)   # 触发时间
    sheet.set_column(3, 3, 10)   # 耗时
    sheet.set_column(4, 4, 10)   # 结果
    sheet.set_column(5, 12, 14)  # 指标列

    workbook.close()
    output.seek(0)

    filename = f"analysis_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def register_routes(app):
    """将分析路由注册到 Flask app"""
    app.register_blueprint(analysis_bp)
    logger.info("Analysis API routes registered")
