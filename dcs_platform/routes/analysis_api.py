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
from dcs_platform.core.db import get_cycles, get_label_stats, export_labels, insert_cycle, \
    get_detect_configs, get_detect_config, get_default_detect_config, \
    upsert_detect_config, toggle_detect_config, delete_detect_config, \
    get_tuning_config, update_tuning_config, get_tuning_runs, get_tuning_run, \
    get_tuning_history, insert_tuning_run, insert_tuning_history, \
    get_result_judge_configs, get_result_judge_config, get_default_result_config, \
    upsert_result_judge_config, toggle_result_judge_config, delete_result_judge_config, \
    get_encoder_calibration, upsert_encoder_calibration
from dcs_platform.services.dynamic_detector import DynamicBreakthroughDetector, run_dynamic_analysis
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

# === 信号映射: 东设备信号 → 西设备信号 ===
# 判定规则默认使用东设备信号名，西设备自动映射
_EAST_TO_WEST_MAP = {
    # 开口机
    "LT_LQFC_57": "LT_LQFC_94",   # 开口机选择
    "LT_LQFC_59": "LT_LQFC_96",   # 回转进/退
    "LT_LQFC_61": "LT_LQFC_98",   # 小车前进/后退
    "LT_LQFC_63": "LT_LQFC_100",  # 回转位置
    "LT_LQFC_64": "LT_LQFC_101",  # 回转压力
    "LT_LQFC_66": "LT_LQFC_103",  # 倾动压力
    "LT_LQFC_67": "LT_LQFC_104",  # 推进位置
    "LT_LQFC_68": "LT_LQFC_105",  # 推进压力
    "LT_LQFC_69": "LT_LQFC_106",  # 冲击开/关
    "LT_LQFC_74": "LT_LQFC_111",  # 回转回油
    "LT_LQFC_75": "LT_LQFC_112",  # 倾动回油
    "LT_LQFC_87": "LT_LQFC_124",  # 转钎压力
    "LT_LQFC_88": "LT_LQFC_125",  # 冲击压力
    # 堵口机
    "LT_LQFC_130": "LT_LQFC_153", # 遥控启动
    "LT_LQFC_133": "LT_LQFC_156", # 回转进/退
    "LT_LQFC_134": "LT_LQFC_157", # 打泥前进/后退
    "LT_LQFC_135": "LT_LQFC_158", # 回转位置
    "LT_LQFC_137": "LT_LQFC_160", # 打泥位置
    "LT_LQFC_138": "LT_LQFC_161", # 打泥压力
    "LT_LQFC_139": "LT_LQFC_162", # 退泥压力
    "LT_LQFC_179": "LT_LQFC_180", # 打泥量
}


def _resolve_signal(rule_signal, equipment_id):
    """将规则中的信号名映射到实际设备信号名。
    
    规则以东设备信号名为基准，如果当前设备是西设备，自动映射。
    """
    if isinstance(equipment_id, str) and "west" in equipment_id:
        return _EAST_TO_WEST_MAP.get(rule_signal, rule_signal)
    return rule_signal


@analysis_bp.route("/ping")
def api_ping():
    ok = ping()
    return jsonify({
        "status": "ok" if ok else "error",
        "influxdb": ok,
        "reachable": ok,
        "message": "InfluxDB 连接正常" if ok else "InfluxDB 无响应，请检查网络或 InfluxDB 服务状态"
    })


# ===== 周期检测配置 CRUD =====

@analysis_bp.route("/detect-configs")
def api_detect_configs():
    """GET /api/analysis/detect-configs?type=opening  含绑定统计"""
    from dcs_platform.core.db import _get_conn
    cycle_type = request.args.get("type", "").strip() or None
    configs = get_detect_configs(cycle_type=cycle_type)
    for c in configs:
        did = c.get("id")
        c["bound_judge_count"] = _get_conn().execute(
            "SELECT COUNT(*) FROM result_judge_configs WHERE detect_config_id=? AND enabled=1", (did,)
        ).fetchone()[0]
        c["bound_group_count"] = _get_conn().execute(
            "SELECT COUNT(*) FROM rule_groups WHERE detect_config_id=? AND enabled=1", (did,)
        ).fetchone()[0]
    return jsonify({"configs": configs, "count": len(configs)})


@analysis_bp.route("/detect-configs/<int:config_id>")
def api_get_detect_config(config_id):
    config = get_detect_config(config_id)
    if not config:
        return jsonify({"error": "配置不存在"}), 404
    return jsonify({"config": config})


@analysis_bp.route("/detect-configs", methods=["POST"])
def api_create_detect_config():
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip()
    cycle_type = data.get("cycle_type", "opening")
    config_data = data.get("config", {})
    description = data.get("description", "")
    is_default = data.get("is_default", 0)
    if not name:
        return jsonify({"error": "名称不能为空"}), 400
    config_id = upsert_detect_config(
        None, name, cycle_type,
        json.dumps(config_data, ensure_ascii=False),
        description, is_default
    )
    return jsonify({"ok": True, "id": config_id})


@analysis_bp.route("/detect-configs/<int:config_id>", methods=["PUT"])
def api_update_detect_config(config_id):
    data = request.get_json(silent=True) or {}
    existing = get_detect_config(config_id)
    if not existing:
        return jsonify({"error": "配置不存在"}), 404
    upsert_detect_config(
        config_id,
        data.get("name", existing["name"]),
        data.get("cycle_type", existing["cycle_type"]),
        json.dumps(data.get("config", existing["config"]), ensure_ascii=False),
        data.get("description", existing.get("description", "")),
        data.get("is_default", existing.get("is_default", 0))
    )
    return jsonify({"ok": True})


@analysis_bp.route("/detect-configs/<int:config_id>/toggle", methods=["POST"])
def api_toggle_detect_config(config_id):
    data = request.get_json(silent=True) or {}
    enabled = data.get("enabled", 1)
    toggle_detect_config(config_id, enabled)
    return jsonify({"ok": True})


@analysis_bp.route("/detect-configs/<int:config_id>", methods=["DELETE"])
def api_delete_detect_config(config_id):
    delete_detect_config(config_id)
    return jsonify({"ok": True})


@analysis_bp.route("/detect-configs/default")
def api_default_detect_config():
    """GET /api/analysis/detect-configs/default?type=opening"""
    cycle_type = request.args.get("type", "opening")
    config = get_default_detect_config(cycle_type)
    return jsonify({"config": config})


# ===== 自整定 API =====

@analysis_bp.route("/tuning/config", methods=["GET", "POST"])
def api_tuning_config():
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        update_tuning_config(**{k: v for k, v in data.items()
                              if k in ("auto_mode", "schedule_hour", "eval_min_samples",
                                       "min_accuracy", "max_false_rate")})
    return jsonify({"config": get_tuning_config()})


@analysis_bp.route("/tuning/trigger", methods=["POST"])
def api_tuning_trigger():
    """手动触发一次自整定"""
    data = request.get_json(silent=True) or {}
    cycle_type = data.get("cycle_type", "opening")
    from dcs_platform.services.self_tuning import run_self_tuning
    result = run_self_tuning(cycle_type, "manual")
    return jsonify(result)


@analysis_bp.route("/tuning/runs")
def api_tuning_runs():
    cycle_type = request.args.get("type", "").strip() or None
    runs = get_tuning_runs(cycle_type=cycle_type)
    return jsonify({"runs": runs, "count": len(runs)})


@analysis_bp.route("/tuning/runs/<int:run_id>")
def api_tuning_run_detail(run_id):
    run = get_tuning_run(run_id)
    if not run:
        return jsonify({"error": "not found"}), 404
    return jsonify({"run": run})


@analysis_bp.route("/tuning/history")
def api_tuning_history():
    config_id = request.args.get("config_id", "").strip() or None
    if config_id:
        config_id = int(config_id)
    history = get_tuning_history(config_id=config_id)
    return jsonify({"history": history, "count": len(history)})


# ===== 判定规则配置 CRUD =====

@analysis_bp.route("/result-configs")
def api_result_configs():
    """GET /api/analysis/result-configs?type=opening&category=success"""
    cycle_type = request.args.get("type", "").strip() or None
    category = request.args.get("category", "").strip() or None
    configs = get_result_judge_configs(cycle_type=cycle_type, category=category)
    return jsonify({"configs": configs, "count": len(configs)})


@analysis_bp.route("/result-configs/<int:config_id>")
def api_get_result_config(config_id):
    c = get_result_judge_config(config_id)
    if not c: return jsonify({"error": "不存在"}), 404
    return jsonify({"config": c})


@analysis_bp.route("/result-configs", methods=["POST"])
def api_create_result_config():
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip()
    if not name: return jsonify({"error": "名称不能为空"}), 400
    config_id = upsert_result_judge_config(
        None, name,
        data.get("cycle_type", "opening"),
        data.get("category", "success"),
        json.dumps(data.get("params", []), ensure_ascii=False),
        data.get("logic_op", "AND"),
        data.get("is_default", 0),
        data.get("description", ""),
        data.get("priority", 0),
        data.get("is_static", 0),
        data.get("detect_config_id", 0)
    )
    return jsonify({"ok": True, "id": config_id})


@analysis_bp.route("/result-configs/<int:config_id>", methods=["PUT"])
def api_update_result_config(config_id):
    existing = get_result_judge_config(config_id)
    if not existing: return jsonify({"error": "不存在"}), 404
    data = request.get_json(silent=True) or {}
    upsert_result_judge_config(
        config_id,
        data.get("name", existing["name"]),
        existing["cycle_type"],
        existing["category"],
        json.dumps(data.get("params", existing["params"]), ensure_ascii=False),
        data.get("logic_op", existing.get("logic_op", "AND")),
        data.get("is_default", existing.get("is_default", 0)),
        data.get("description", existing.get("description", "")),
        data.get("priority", existing.get("priority", 0)),
        data.get("is_static", existing.get("is_static", 0)),
        data.get("detect_config_id", existing.get("detect_config_id", 0))
    )
    return jsonify({"ok": True})


@analysis_bp.route("/result-configs/<int:config_id>/toggle", methods=["POST"])
def api_toggle_result_config(config_id):
    data = request.get_json(silent=True) or {}
    toggle_result_judge_config(config_id, data.get("enabled", 1))
    return jsonify({"ok": True})


@analysis_bp.route("/result-configs/<int:config_id>", methods=["DELETE"])
def api_delete_result_config(config_id):
    delete_result_judge_config(config_id)
    return jsonify({"ok": True})


# ========== 变量采集配置 API ==========

@analysis_bp.route("/variable-configs")
def api_variable_configs():
    """获取变量配置列表，支持按 cycle_type/dimension/filter 过滤"""
    from dcs_platform.core.db import get_variable_configs, get_variable_config_options
    cycle_type = request.args.get("type", "")
    equipment = request.args.get("equipment", "")
    dimension = request.args.get("dimension", "")
    mode = request.args.get("mode", "")

    if mode == "options":
        # 仅返回下拉选项（精简字段）
        options = get_variable_config_options(cycle_type or None)
        return jsonify({"options": options})

    configs = get_variable_configs(
        cycle_type=cycle_type or None,
        equipment=equipment or None,
        dimension=dimension or None,
    )
    return jsonify({"configs": configs, "total": len(configs)})


@analysis_bp.route("/variable-configs/<int:config_id>")
def api_get_variable_config(config_id):
    from dcs_platform.core.db import get_variable_config
    cfg = get_variable_config(config_id)
    if not cfg:
        return jsonify({"error": "not found"}), 404
    return jsonify(cfg)


@analysis_bp.route("/variable-configs", methods=["POST"])
def api_create_variable_config():
    from dcs_platform.core.db import upsert_variable_config
    data = request.get_json(force=True, silent=True) or {}
    tag_name = (data.get("tag_name") or "").strip()
    chinese_name = (data.get("chinese_name") or "").strip()
    if not tag_name or not chinese_name:
        return jsonify({"error": "tag_name 和 chinese_name 不能为空"}), 400
    kwargs = {
        "unit": data.get("unit", ""),
        "cycle_type": data.get("cycle_type", "common"),
        "equipment": data.get("equipment", ""),
        "dimension": data.get("dimension", ""),
        "data_type": data.get("data_type", "float"),
        "description": data.get("description", ""),
        "is_active": data.get("is_active", 1),
    }
    new_id = upsert_variable_config(None, tag_name, chinese_name, **kwargs)
    return jsonify({"ok": True, "id": new_id})


@analysis_bp.route("/variable-configs/<int:config_id>", methods=["PUT"])
def api_update_variable_config(config_id):
    from dcs_platform.core.db import upsert_variable_config, get_variable_config
    cfg = get_variable_config(config_id)
    if not cfg:
        return jsonify({"error": "not found"}), 404
    data = request.get_json(force=True, silent=True) or {}
    tag_name = (data.get("tag_name") or cfg["tag_name"]).strip()
    chinese_name = (data.get("chinese_name") or cfg["chinese_name"]).strip()
    kwargs = {}
    for k in ("unit", "cycle_type", "equipment", "dimension", "data_type", "description", "is_active"):
        if k in data:
            kwargs[k] = data[k]
    upsert_variable_config(config_id, tag_name, chinese_name, **kwargs)
    return jsonify({"ok": True})


@analysis_bp.route("/variable-configs/<int:config_id>", methods=["DELETE"])
def api_delete_variable_config(config_id):
    from dcs_platform.core.db import delete_variable_config
    delete_variable_config(config_id)
    return jsonify({"ok": True})


@analysis_bp.route("/variable-configs/seed", methods=["POST"])
def api_reseed_variable_configs():
    """重新从 param_groups.json 种子变量配置（开发调试用）"""
    from dcs_platform.core.db import _get_conn, seed_variable_configs
    _get_conn().execute("DELETE FROM variable_configs")
    _get_conn().commit()
    seed_variable_configs()
    return jsonify({"ok": True, "message": "变量配置已重新种子"})


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
    config_id = request.args.get("config_id", "")
    
    # Normalize time format: frontend sends "2026-07-03T00:00" (no seconds/Z)
    # Flux requires RFC3339: "2026-07-03T00:00:00Z"
    start = _normalize_time(start)
    end = _normalize_time(end)
    
    # 自适应超时: 根据时间窗口自动调整
    try:
        _rs = datetime.fromisoformat(start.replace("Z","+00:00"))
        _re = datetime.fromisoformat(end.replace("Z","+00:00"))
        _range_secs = (_re - _rs).total_seconds()
        _timeout_ms = max(60000, min(180000, int(_range_secs * 1000)))
    except Exception:
        _timeout_ms = 60000
    
    # ── 正确信号配置 ──
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
    
    # ━━━ 开口检测 ━━━
    if cycle_type in ("opening", "all"):
        # 加载检测配置
        opening_rules, opening_filter = _load_detect_rules("opening", config_id)
        
        # 若配置包含阈值规则，使用阈值联合检测（分别检测东西设备）
        if _has_threshold_rule(opening_rules):
            for machine, sig in OPENING_CONFIG.items():
                is_west = "西" in machine
                mapped_rules = _map_rules_for_equipment(opening_rules, is_west)
                machine_cycles = _detect_threshold_cycles(
                    start, end, mapped_rules, opening_filter, "opening", machine
                )
                cycles.extend(machine_cycles)
        else:
            #  legacy: remote==1 AND swing_pos 穿越 90°
            for machine, sig in OPENING_CONFIG.items():
                # 用配置中的信号替代硬编码
                remote_sig = sig["remote"]
                swing_sig = sig["swing_pos"]
                threshold = 90
                tolerance = 2
            
                if opening_rules:
                    for rule in opening_rules:
                        role = rule.get("role", "")
                        val = rule.get("threshold", 90)
                        tol = rule.get("tolerance_s", 2)
                        sig_name = rule.get("signal", "")
                        if role == "remote":
                            remote_sig = sig_name
                        elif role == "crossing":
                            swing_sig = sig_name
                            threshold = val
                            tolerance = tol
            
                detect_params = [remote_sig, swing_sig]
                try:
                    data = fetch_timeseries(start, end, detect_params, timeout_ms=_timeout_ms)
                except Exception as e:
                    logger.warning("Opening detection InfluxDB query failed for %s: %s", machine, e)
                    continue
            
                remote = _build_time_map(data.get(remote_sig, []))
                swing = sorted(data.get(swing_sig, []), key=lambda x: x[0])
            
                if len(swing) < 2 or len(remote) < 2:
                    continue
            
                in_cycle = False
                cycle_start = None
                filter_min = opening_filter.get("filter_min_s", 30)
                filter_max = opening_filter.get("filter_max_s", 3600)
            
                for i in range(1, len(swing)):
                    prev_v, curr_v = swing[i-1][1], swing[i][1]
                    t = swing[i][0]
                    ts = t.timestamp()
                
                    if not in_cycle and _crossed(prev_v, curr_v, threshold):
                        if _remote_nearby(remote, ts, tolerance=tolerance):
                            in_cycle = True
                            cycle_start = t
                    
                    elif in_cycle and _crossed(curr_v, prev_v, threshold):
                        duration = (t - cycle_start).total_seconds()
                        if filter_min <= duration <= filter_max:
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
            
                if in_cycle and swing:
                    last_t = swing[-1][0]
                    duration = (last_t - cycle_start).total_seconds()
                    if filter_min <= duration <= filter_max:
                        cycles.append({
                            "machine": machine, "type": "opening",
                            "trigger_time": cycle_start.isoformat(),
                            "window_start": cycle_start.isoformat(),
                            "window_end": last_t.isoformat(),
                            "duration_s": round(duration, 1),
                            "result": "unknown", "breakthrough": False,
                        })
    
    # ━━━ 堵口检测 ━━━
    if cycle_type in ("plugging", "all"):
        plugging_rules, plugging_filter = _load_detect_rules("plugging", config_id)
        
        if _has_threshold_rule(plugging_rules):
            for machine, sig in PLUGGING_CONFIG.items():
                is_west = "西" in machine
                mapped_rules = _map_rules_for_equipment(plugging_rules, is_west)
                machine_cycles = _detect_threshold_cycles(
                    start, end, mapped_rules, plugging_filter, "plugging", machine
                )
                cycles.extend(machine_cycles)
        else:
            # legacy: remote_start==1 AND mud_cmd 边沿
            # 从规则中读取 mud_cmd 边沿阈值（非二进制信号需要>thr 检测）
            mud_edge = 0.5  # 默认二进制边沿
            for rule in plugging_rules:
                if rule.get("role") == "mud_cmd":
                    mud_edge = rule.get("threshold", 0.5)
                    break
            
            for machine, sig in PLUGGING_CONFIG.items():
                is_west = "西" in machine
                # 用规则中的信号覆盖硬编码（支持西设备映射）
                remote_sig = sig["remote_start"]
                mud_sig = sig["mud_cmd"]
                for rule in plugging_rules:
                    r_sig = rule.get("signal", "")
                    if is_west and r_sig in _EAST_TO_WEST_MAP:
                        r_sig = _EAST_TO_WEST_MAP[r_sig]
                    if rule.get("role") == "remote":
                        remote_sig = r_sig
                    elif rule.get("role") == "mud_cmd":
                        mud_sig = r_sig
                
                detect_params = [remote_sig, mud_sig]
                try:
                    data = fetch_timeseries(start, end, detect_params, timeout_ms=_timeout_ms)
                except Exception as e:
                    logger.warning("Plugging detection InfluxDB query failed for %s: %s", machine, e)
                    continue
            
                remote_start = _build_time_map(data.get(remote_sig, []))
                mud_cmd = sorted(data.get(mud_sig, []), key=lambda x: x[0])
            
                if len(mud_cmd) < 2 or len(remote_start) < 2:
                    continue
            
                in_cycle = False
                cycle_start = None
                for i in range(1, len(mud_cmd)):
                    prev_v, curr_v = mud_cmd[i-1][1], mud_cmd[i][1]
                    t = mud_cmd[i][0]
                    ts = t.timestamp()
                
                    # Rising edge on mud_cmd (crossing mud_edge threshold)
                    if not in_cycle and prev_v < mud_edge and curr_v >= mud_edge:
                        if _remote_nearby(remote_start, ts, tolerance=2):
                            in_cycle = True
                            cycle_start = t
                
                    # Falling edge = cycle end
                    elif in_cycle and prev_v >= mud_edge and curr_v < mud_edge:
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
    
    # ── 批量提取指标和结果判定（基于判定规则配置）──
    # 大时间范围跳过批量 enrich（InfluxDB 查询量过大），cycles 保留 'unknown' 结果
    # 前端点「详情」时再通过 /api/analysis/metrics 逐个评估
    try:
        range_start = datetime.fromisoformat(start.replace("Z", "+00:00"))
        range_end = datetime.fromisoformat(end.replace("Z", "+00:00"))
        range_days = (range_end - range_start).total_seconds() / 86400
    except Exception:
        range_days = 0
    
    if range_days <= 1.5 and cycles:
        try: dc_id = int(config_id) if config_id else 0
        except: dc_id = 0
        _enrich_cycles_with_metrics(cycles, start, end, detect_config_id=dc_id)
    
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
        "saved_to_db": saved_count,
        "hint": "未检测到作业周期，请检查日期范围是否有DCS数据" if len(cycles) == 0 else ""
    })


def _enrich_cycles_with_metrics(cycles, time_start, time_end, detect_config_id=0):
    """批量提取所有周期的判定指标（一次 InfluxDB 查询 + 逐个窗口评估）"""
    if not cycles:
        return
    
    # 按 (type, east/west) 分组，因为需要不同的信号集
    groups = {}
    for c in cycles:
        is_plug = c["type"] == "plugging"
        is_west = "西" in c.get("machine", "")
        key = ("plugging" if is_plug else "opening", "west" if is_west else "east")
        groups.setdefault(key, []).append(c)
    
    for (cyc_type, direction), group_cycles in groups.items():
        equip_id = f"{direction}_{'plugger' if cyc_type == 'plugging' else 'opener'}"
        is_plugging = cyc_type == "plugging"
        
        # 收集本组需要的所有信号
        signals = _get_relevant_signals(cyc_type, equip_id)
        resolved_signals = [_resolve_signal(s, equip_id) for s in signals]
        
        try:
            data = fetch_timeseries(time_start, time_end, resolved_signals, timeout_ms=60000)
        except Exception:
            # 逐个窗口重试
            for c in group_cycles:
                try:
                    _enrich_single_cycle(c, is_plugging, equip_id)
                except Exception:
                    pass
            continue
        
        # 逐个窗口评估
        for c in group_cycles:
            try:
                st_str = c["window_start"]
                et_str = c["window_end"]
                st = datetime.fromisoformat(st_str.replace("Z", "+00:00"))
                et = datetime.fromisoformat(et_str.replace("Z", "+00:00"))
            except Exception:
                continue
            
            metrics = _extract_cycle_metrics(data, st, et, is_plugging, equip_id, detect_config_id)
            c["result"] = metrics.get("result", "unknown")
            c["breakthrough"] = metrics.get("breakthrough", False)
            c["metrics"] = metrics  # 完整指标供前端使用


def _enrich_single_cycle(cycle, is_plugging, equipment_id):
    """单个周期的信号查询 + 指标提取（回退方案）"""
    try:
        signals = _get_relevant_signals("plugging" if is_plugging else "opening", equipment_id)
        resolved = [_resolve_signal(s, equipment_id) for s in signals]
        data = fetch_timeseries(
            cycle["window_start"], cycle["window_end"], resolved, timeout_ms=30000
        )
        st = datetime.fromisoformat(cycle["window_start"].replace("Z", "+00:00"))
        et = datetime.fromisoformat(cycle["window_end"].replace("Z", "+00:00"))
        metrics = _extract_cycle_metrics(data, st, et, is_plugging, equipment_id)
        cycle["result"] = metrics.get("result", "unknown")
        cycle["breakthrough"] = metrics.get("breakthrough", False)
        cycle["metrics"] = metrics
    except Exception:
        pass


def _load_detect_rules(cycle_type, config_id=""):
    """从数据库加载检测规则配置。
    Returns: (rules_list, filter_config)
    """
    rules = []
    filter_cfg = {"filter_min_s": 30, "filter_max_s": 3600}
    
    if config_id:
        try:
            config = get_detect_config(int(config_id))
        except (ValueError, TypeError):
            config = None
    else:
        config = get_default_detect_config(cycle_type)
    
    if config and config.get("enabled"):
        cfg = config.get("config", {})
        rules = cfg.get("rules", [])
        filter_cfg["filter_min_s"] = cfg.get("filter_min_s", 30)
        filter_cfg["filter_max_s"] = cfg.get("filter_max_s", 3600)
    
    return rules, filter_cfg


def _has_threshold_rule(rules):
    """检查规则列表是否包含阈值条件规则"""
    return any(r.get("role") == "threshold" for r in rules)


def _map_rules_for_equipment(rules, is_west):
    """将规则中的信号名映射到对应设备。
    
    规则默认使用东设备信号名，西设备需要映射。
    返回新规则列表（浅拷贝，修改 signal 字段）。
    """
    if not is_west:
        return rules
    mapped = []
    for r in rules:
        new_r = dict(r)
        sig = new_r.get("signal", "")
        if sig in _EAST_TO_WEST_MAP:
            new_r["signal"] = _EAST_TO_WEST_MAP[sig]
        mapped.append(new_r)
    return mapped


def _detect_threshold_cycles(start, end, rules, filter_cfg, cycle_type, machine_label=""):
    """基于阈值规则的周期检测：所有规则条件同时满足的持续区间"""
    signals = list(set(r["signal"] for r in rules if r.get("signal")))
    if not signals:
        return []

    try:
        data = fetch_timeseries(start, end, signals, timeout_ms=60000)
    except Exception:
        return []

    # 构建每条规则的时间序列
    rule_series = []
    for rule in rules:
        pts = data.get(rule.get("signal", ""), [])
        if not pts:
            continue
        rule_series.append((rule, sorted(pts, key=lambda x: x[0])))

    if not rule_series:
        return []

    all_times = sorted(set(t for _, pts in rule_series for t, _ in pts))
    if not all_times:
        return []

    filter_min = filter_cfg.get("filter_min_s", 30)
    filter_max = filter_cfg.get("filter_max_s", 3600)

    cycles = []
    in_cycle = False
    cycle_start = None

    for t in all_times:
        all_true = True
        for rule, pts in rule_series:
            # 取 t 时刻或之前最近的值
            val = None
            for pt_t, pt_v in pts:
                if pt_t <= t:
                    val = pt_v
                else:
                    break
            if val is None:
                all_true = False
                break

            role = rule.get("role", "")
            if role == "remote":
                if not (val == 1 or val > 0.5):
                    all_true = False
                    break
            elif role == "threshold":
                threshold = rule.get("threshold", 0)
                operator = rule.get("operator", "gt")
                if operator == "gt" and not (val > threshold):
                    all_true = False
                    break
                if operator == "lt" and not (val < threshold):
                    all_true = False
                    break

        if all_true and not in_cycle:
            in_cycle = True
            cycle_start = t
        elif not all_true and in_cycle:
            duration = (t - cycle_start).total_seconds()
            if filter_min <= duration <= filter_max:
                cycles.append({
                    "machine": machine_label or ("东开口机" if cycle_type == "opening" else "东堵口机"),
                    "type": cycle_type,
                    "trigger_time": cycle_start.isoformat(),
                    "window_start": cycle_start.isoformat(),
                    "window_end": t.isoformat(),
                    "duration_s": round(duration, 1),
                    "result": "unknown", "breakthrough": False,
                })
            in_cycle = False

    if in_cycle and cycle_start:
        duration = (all_times[-1] - cycle_start).total_seconds()
        if filter_min <= duration <= filter_max:
            cycles.append({
                "machine": machine_label or ("东开口机" if cycle_type == "opening" else "东堵口机"),
                "type": cycle_type,
                "trigger_time": cycle_start.isoformat(),
                "window_start": cycle_start.isoformat(),
                "window_end": all_times[-1].isoformat(),
                "duration_s": round(duration, 1),
                "result": "unknown", "breakthrough": False,
            })

    return cycles


def _normalize_time(time_str):
    """Normalize time to RFC3339 UTC.
    
    前端发送本地时间(北京时间 UTC+8, 无时区标记), 需转为 UTC 再查 InfluxDB.
    '2026-07-03T00:00' -> '2026-07-02T16:00:00Z'
    已有 Z/+xx:xx 后缀的保持不变.
    """
    if not time_str:
        return time_str
    # Already has timezone marker
    if "Z" in time_str or "+" in time_str[-6:]:
        return time_str
    # Has seconds but no timezone: add Z
    if len(time_str) >= 19 and time_str[10] == 'T':
        return time_str + "Z"
    # No seconds: "2026-07-03T00:00" — assume Beijing local, convert to UTC
    if "T" in time_str:
        try:
            from datetime import timedelta
            dt = datetime.fromisoformat(time_str)
            dt = dt - timedelta(hours=8)  # Beijing UTC+8 → UTC
            return dt.strftime("%Y-%m-%dT%H:%M:%S") + "Z"
        except Exception:
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

def _compute_signal_metrics(data, start_time, end_time, signal_name):
    """计算单个信号在时间窗口内的基本统计量"""
    vals = [v for t, v in data.get(signal_name, []) if start_time <= t <= end_time]
    if not vals:
        return {"min": 0, "max": 0, "mean": 0, "range": 0, "count": 0}
    return {
        "min": min(vals), "max": max(vals),
        "mean": sum(vals) / len(vals),
        "range": max(vals) - min(vals),
        "count": len(vals),
        "late_early_ratio": _late_early_ratio(vals),
        "values": vals,
    }

def _late_early_ratio(vals):
    """计算后期/前期压力比值（用于钻透判定）
    
    取前1/3为"前期"，后1/3为"后期"，比值越小说明后期压力降得越多。
    """
    n = len(vals)
    if n < 6:
        return 1.0
    third = max(1, n // 3)
    early = sum(vals[:third]) / third
    late = sum(vals[-third:]) / third
    if early <= 0:
        return 1.0
    return late / early

def _eval_param(param, signal_metrics):
    """根据判定规则的参数定义，评估信号是否满足条件
    
    Args:
        param: 判定参数 {param_name, value, operator, label, unit}
        signal_metrics: 信号统计量 {min, max, mean, range, late_early_ratio, values}
    
    Returns: True/False
    """
    op = param.get("operator", "gte")
    threshold = float(param.get("value", 0))
    param_name = param.get("param_name", "")
    
    # 根据 param_name 语义决定使用哪个统计量
    # LT_LQFC_67 在开口中使用 range (行程位移)，在钻透/深度中也可能用 ratio
    # LT_LQFC_68 通常使用 late_early_ratio (压力骤降比)
    # LT_LQFC_69 使用 max (是否激活)
    # LT_LQFC_179 使用 range (打泥量)
    # LT_LQFC_137 使用 range 或 max
    # LT_LQFC_138 使用 max (峰值)
    # LT_LQFC_134 使用 count (保压时长)
    
    if op in ("gt", "gte", "lt", "lte"):
        # 根据信号类型选择合适的统计量
        val = _pick_value_for_param(param_name, param, signal_metrics)
        if op == "gt": return val > threshold
        if op == "gte": return val >= threshold
        if op == "lt": return val < threshold
        if op == "lte": return val <= threshold
    
    if op == "eq":
        # bool 类型：信号是否激活
        val = signal_metrics.get("max", 0)
        if threshold == 1:
            return val >= 0.5  # 二进制信号 >= 0.5 视为 1
        return val == threshold
    
    return True  # 未知操作符，默认通过

def _pick_value_for_param(param_name, param, sm):
    """根据参数名和语义选择合适的信号统计量"""
    param_name = param_name or ""
    label = (param.get("label", "") or "").lower()
    
    # 压力骤降比: 使用 late_early_ratio
    if "骤降" in label or "降比" in label or "ratio" in label:
        return 1.0 - sm.get("late_early_ratio", 1.0)  # 转换为降幅
    if param_name == "LT_LQFC_68":
        return 1.0 - sm.get("late_early_ratio", 1.0)  # 默认: 压力降幅
    
    # 行程占比/到位比: 使用 range 相对于期望行程的比值
    if "占比" in label or "到位比" in label or "ratio" in param.get("unit", ""):
        return sm.get("range", 0)
    
    # 位移/行程: 使用 range
    if param_name in ("LT_LQFC_67", "LT_LQFC_137"):
        if sm.get("range", 0) > 0:
            return sm["range"]
        return sm.get("max", 0)
    
    # 压力峰值: 使用 max
    if param_name in ("LT_LQFC_68", "LT_LQFC_138", "LT_LQFC_66", "LT_LQFC_87"):
        return sm.get("max", 0)
    
    # 时长/保压: 使用 count
    if param_name in ("LT_LQFC_134", "LT_LQFC_69"):
        return sm.get("count", 0)
    
    # 打泥量: 使用 range
    if param_name in ("LT_LQFC_179", "LT_LQFC_180"):
        return sm.get("range", 0)
    
    # 大臂到位角度: 使用 min（min越小说明已回位）
    if param_name == "LT_LQFC_63":
        return sm.get("min", 0)
    
    # 编码器校正: 总是 true（仅标记）
    if "encoder_offset" in param_name or "calib" in param_name:
        return 1
    
    # 默认: 使用 range
    return sm.get("range", sm.get("max", 0))


def _extract_cycle_metrics(
    data: dict, start_time, end_time, is_plugging: bool, equipment_id: str,
    detect_config_id: int = 0
) -> dict:
    """从配置驱动的判定规则提取周期指标和判定结果。
    
    优先级: result_judge_configs 配置表 > 硬编码兜底逻辑
    """
    metrics = {}
    cycle_type = "plugging" if is_plugging else "opening"
    
    # 1. 预先计算所有相关信号的基本统计量    
    signal_stats = {}
    relevant_signals = _get_relevant_signals(cycle_type, equipment_id)
    for sig in relevant_signals:
        resolved = _resolve_signal(sig, equipment_id)
        signal_stats[sig] = _compute_signal_metrics(data, start_time, end_time, resolved)
    
    # 2. 提取基础指标（用于前端展示）
    if is_plugging:
        mud_sig = _resolve_signal("LT_LQFC_179", equipment_id)
        mud_press_sig = _resolve_signal("LT_LQFC_138", equipment_id)
        hold_sig = _resolve_signal("LT_LQFC_134", equipment_id)
        
        mud_stats = signal_stats.get("LT_LQFC_179", {})
        press_stats = signal_stats.get("LT_LQFC_138", {})
        hold_stats = signal_stats.get("LT_LQFC_134", {})
        
        metrics["mud_qty"] = round(mud_stats.get("range", 0), 2)
        metrics["mud_press_peak"] = round(press_stats.get("max", 0), 1)
        metrics["mud_press_mean"] = round(press_stats.get("mean", 0), 1)
        metrics["hold_duration_s"] = hold_stats.get("count", 0)
        metrics["hold_ok"] = metrics["hold_duration_s"] >= 60
        metrics["breakthrough"] = metrics["mud_press_peak"] >= 15 and metrics["mud_qty"] >= 80
    else:
        pos_sig = _resolve_signal("LT_LQFC_67", equipment_id)
        press_sig = _resolve_signal("LT_LQFC_68", equipment_id)
        drill_sig = _resolve_signal("LT_LQFC_87", equipment_id)
        impact_sig = _resolve_signal("LT_LQFC_69", equipment_id)
        
        pos_stats = signal_stats.get("LT_LQFC_67", {})
        press_stats = signal_stats.get("LT_LQFC_68", {})
        drill_stats = signal_stats.get("LT_LQFC_87", {})
        impact_stats = signal_stats.get("LT_LQFC_69", {})
        
        metrics["push_depth"] = round(pos_stats.get("range", 0), 3)
        metrics["push_pos_change"] = round(pos_stats.get("range", 0), 3)
        metrics["push_press_max"] = round(press_stats.get("max", 0), 1)
        metrics["push_press_mean"] = round(press_stats.get("mean", 0), 1)
        metrics["push_press_peak"] = round(press_stats.get("max", 0), 1)
        metrics["drill_press_max"] = round(drill_stats.get("max", 0), 1)
        metrics["drill_press_mean"] = round(drill_stats.get("mean", 0), 1)
        metrics["impact_press_active"] = impact_stats.get("max", 0) >= 0.5
        
        # 硬编码钻透兜底
        breakthrough = False
        pos_vals = signal_stats.get("LT_LQFC_67", {}).get("values", [])
        press_vals = signal_stats.get("LT_LQFC_68", {}).get("values", [])
        if len(pos_vals) > 5 and len(press_vals) > 5:
            depth_change = max(pos_vals) - min(pos_vals)
            if depth_change > 2.0:
                ratio = _late_early_ratio(press_vals)
                if ratio < 0.8:
                    breakthrough = True
        metrics["breakthrough"] = breakthrough
    
    # 3. 从配置表加载判定规则并评估（按检测规则绑定 + 优先级排序）
    judge_configs = _load_judge_configs(cycle_type, detect_config_id)
    triggered_alerts = []  # 命中的告警规则
    
    # 也加载 rule_groups 告警规则并评估
    rule_groups = _load_rule_groups(cycle_type, detect_config_id)
    if rule_groups:
        triggered_alerts = _evaluate_rule_groups(rule_groups, signal_stats, cycle_type)
        metrics["triggered_alerts"] = triggered_alerts
    if judge_configs:
        verdicts = {}
        # 按 category 分组，组内按 priority DESC 排序
        by_cat = {}
        for cfg in judge_configs:
            cat = cfg.get("category", "")
            by_cat.setdefault(cat, []).append(cfg)
        for cat in by_cat:
            by_cat[cat].sort(key=lambda c: c.get("priority", 0), reverse=True)
        
        for cat, cfgs in by_cat.items():
            for cfg in cfgs:
                logic_op = cfg.get("logic_op", "AND")
                params = cfg.get("params", [])
                is_static = cfg.get("is_static", 0)
                
                # is_static=1: 无变量, 永远匹配 (如兜底规则)
                if is_static:
                    if cat == "fallback":
                        # fallback 类规则不直接设定结果, 而是让后续硬编码兜底处理
                        pass
                    else:
                        verdicts[cat] = True
                    break
                
                if not params:
                    continue
                
                results = []
                for p in params:
                    param_name = p.get("param_name", "")
                    if not param_name or "encoder_offset" in param_name or "calib" in param_name:
                        results.append(True)
                        continue
                    sm = signal_stats.get(param_name)
                    if sm is None:
                        results.append(False)
                        continue
                    results.append(_eval_param(p, sm))
                
                match = all(results) if logic_op == "AND" else any(results)
                if match:
                    verdicts[cat] = True
                    break  # 同 category 首条匹配即停止
        
        # 4. 综合判定结果 + 冲突解决策略:
        #    优先级顺序: success > fail > incomplete/unfinished > hardcoded fallback
        metrics["judge_details"] = verdicts
        
        if is_plugging:
            if verdicts.get("success"):
                metrics["result"] = "success"
            elif verdicts.get("fail"):
                metrics["result"] = "fail"
            elif verdicts.get("unfinished"):
                metrics["result"] = "partial"
            else:
                # 兜底
                if metrics["mud_qty"] > 0.1 and metrics["hold_ok"]:
                    metrics["result"] = "success"
                elif metrics["mud_qty"] > 0.1:
                    metrics["result"] = "partial"
                else:
                    metrics["result"] = "fail"
            
            # 堵口钻透 = 泥炮到位
            if verdicts.get("breakthrough"):
                metrics["breakthrough"] = True
        else:
            if verdicts.get("success"):
                metrics["result"] = "success"
            elif verdicts.get("fail"):
                metrics["result"] = "fail"
            elif verdicts.get("incomplete"):
                metrics["result"] = "incomplete"
            else:
                if metrics["breakthrough"]:
                    metrics["result"] = "success"
                elif metrics["push_depth"] > 0.1:
                    metrics["result"] = "incomplete"
                else:
                    metrics["result"] = "fail"
            
            if verdicts.get("breakthrough"):
                metrics["breakthrough"] = True
        
        # 深度计算
        if verdicts.get("depth"):
            depth_sig = "LT_LQFC_67" if not is_plugging else "LT_LQFC_137"
            ds = signal_stats.get(depth_sig, {})
            metrics["depth_effective"] = ds.get("range", 0) >= 1.0
    else:
        # 无配置表规则，使用硬编码兜底
        if is_plugging:
            if metrics["mud_qty"] > 0.1 and metrics["hold_ok"]:
                metrics["result"] = "success"
            elif metrics["mud_qty"] > 0.1:
                metrics["result"] = "partial"
            else:
                metrics["result"] = "fail"
        else:
            if metrics["breakthrough"]:
                metrics["result"] = "success"
            elif metrics["push_depth"] > 0.1:
                metrics["result"] = "incomplete"
            else:
                metrics["result"] = "fail"
    
    return metrics


def _get_relevant_signals(cycle_type, equipment_id):
    """获取当前作业类型和设备的全部相关信号名（以东设备为基准）"""
    if cycle_type == "opening":
        return ["LT_LQFC_67", "LT_LQFC_68", "LT_LQFC_69", "LT_LQFC_63",
                "LT_LQFC_87", "LT_LQFC_66", "LT_LQFC_64", "LT_LQFC_88"]
    else:
        return ["LT_LQFC_179", "LT_LQFC_138", "LT_LQFC_137",
                "LT_LQFC_134", "LT_LQFC_135", "LT_LQFC_136"]


def _load_judge_configs(cycle_type, detect_config_id=0):
    """从数据库加载判定规则配置（按检测规则绑定过滤）"""
    try:
        configs = get_result_judge_configs(cycle_type=cycle_type)
        result = [c for c in configs if c.get("enabled")]
        if detect_config_id:
            result = [c for c in result if c.get("detect_config_id") == detect_config_id]
        return result
    except Exception:
        return []


def _load_rule_groups(cycle_type, detect_config_id=0):
    """从数据库加载告警/评分规则组（按检测规则绑定过滤, 含子规则）"""
    try:
        from dcs_platform.core.db import _get_conn
        q = "SELECT * FROM rule_groups WHERE enabled=1 AND cycle_type=?"
        args = [cycle_type]
        if detect_config_id:
            q += " AND detect_config_id=?"
            args.append(detect_config_id)
        q += " ORDER BY priority DESC"
        rows = _get_conn().execute(q, tuple(args)).fetchall()
        groups = []
        for r in rows:
            g = dict(r)
            g["rules"] = [dict(r2) for r2 in _get_conn().execute(
                "SELECT * FROM rules WHERE group_id=? AND enabled=1 ORDER BY priority DESC",
                (r["id"],)
            ).fetchall()]
            groups.append(g)
        return groups
    except Exception:
        return []


def _evaluate_rule_groups(rule_groups, signal_stats, cycle_type):
    """评估 rule_groups 规则，返回命中的告警列表"""
    triggered = []
    for grp in rule_groups:
        logic_op = grp.get("logic_op", "AND")
        rules = grp.get("rules", [])
        if not rules:
            continue
        results = []
        for rule in rules:
            param = rule.get("param_name", "")
            if not param or param.startswith("_prev."):
                results.append(True)  # 历史引用参数跳过
                continue
            sm = signal_stats.get(param)
            if sm is None:
                results.append(False)
                continue
            op = rule.get("operator", "gt")
            thr = rule.get("threshold_value", 0)
            val = _extract_signal_value(sm, op)
            met = False
            if op in ("gt",): met = val > thr
            elif op in ("lt",): met = val < thr
            elif op in ("gte",): met = val >= thr
            elif op in ("lte",): met = val <= thr
            elif op in ("eq",): met = val == thr
            elif op in ("between",):
                thr2 = rule.get("threshold_value2", 0)
                met = thr <= val <= thr2
            results.append(met)
        match = all(results) if logic_op == "AND" else any(results)
        if match:
            triggered.append({
                "group_name": grp.get("name", ""),
                "group_id": grp.get("id"),
                "priority": grp.get("priority", 0),
            })
    return triggered


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

    # Build param list using variable config signals
    relevant_signals = _get_relevant_signals(
        "plugging" if is_plugging else "opening", equipment_id
    )
    params = [_resolve_signal(s, equipment_id) for s in relevant_signals]
    params = list(set(params))  # 去重

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


# ===================================================================
#  编码器校准 & 动态钻透/深度计算 API
# ===================================================================

@analysis_bp.route("/encoder-calibration")
def api_get_encoder_calib():
    """GET /api/analysis/encoder-calibration?machine=&cycle_type="""
    machine = request.args.get("machine", "")
    cycle_type = request.args.get("cycle_type", "")
    calibrations = get_encoder_calibration(machine or None, cycle_type or None)
    return jsonify({"calibrations": calibrations, "count": len(calibrations)})


@analysis_bp.route("/encoder-calibration", methods=["POST"])
def api_upsert_encoder_calib():
    """POST /api/analysis/encoder-calibration"""
    data = request.get_json(silent=True) or {}
    machine = data.get("machine", "").strip()
    if not machine:
        return jsonify({"error": "设备名不能为空"}), 400
    upsert_encoder_calibration(
        machine=machine,
        cycle_type=data.get("cycle_type", "opening"),
        position_signal=data.get("position_signal", ""),
        offset_baseline=float(data.get("offset_baseline", 0)),
        travel_range_min=float(data.get("travel_range_min", 0.1)),
        travel_range_max=float(data.get("travel_range_max", 3.0)),
        slope_correction=float(data.get("slope_correction", 1.0)),
        description=data.get("description", ""),
    )
    return jsonify({"ok": True})


@analysis_bp.route("/dynamic-analysis")
def api_dynamic_analysis():
    """GET /api/analysis/dynamic-analysis?start=...&end=...&type=...&machine=...
    
    动态钻透判定 + 铁口深度计算:
    - 自动建立位置基线
    - 编码器偏移校正
    - 多维钻透判定（行程占比 + 速度拐点 + 位移稳定）
    - 输出修正后铁口深度
    """
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    cycle_type = request.args.get("type", "opening")
    machine = request.args.get("machine", "东开口机")

    if not start or not end:
        return jsonify({"error": "缺少时间参数"}), 400

    start = _normalize_time(start)
    end = _normalize_time(end)

    # 加载编码器校准
    calibs = get_encoder_calibration(machine=machine)
    calib = calibs[0] if calibs else {}

    # 确定位置信号
    MACHINE_SIGNALS = {
        ("opening", "东开口机"): "LT_LQFC_67",
        ("opening", "西开口机"): "LT_LQFC_104",
        ("plugging", "东堵口机"): "LT_LQFC_137",
        ("plugging", "西堵口机"): "LT_LQFC_160",
    }
    pos_signal = calib.get("position_signal") or MACHINE_SIGNALS.get(
        (cycle_type, machine), ""
    )

    if not pos_signal:
        return jsonify({"error": "未找到位置信号配置"}), 400

    result = run_dynamic_analysis(start, end, cycle_type, pos_signal, calib, machine)
    return jsonify(result)


def register_routes(app):
    """将分析路由注册到 Flask app"""
    app.register_blueprint(analysis_bp)
    logger.info("Analysis API routes registered")
