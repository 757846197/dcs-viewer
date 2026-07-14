"""
DCS 平台 — 规则 REST API

提供开口/堵口规则组的 CRUD、规则评估、阈值建议端点。
挂载为 Flask Blueprint，复用 Token 认证。
"""
import json
import logging
from flask import Blueprint, request, jsonify

from dcs_platform.core.db import (
    get_rule_groups, get_rule_group,
    upsert_rule_group, delete_rule_group, toggle_rule_group,
    clear_rules_in_group, insert_rule,
)
from dcs_platform.services.rule_engine import (
    load_rules, evaluate_rule, evaluate_rules, get_rule_suggestions, get_signal_dimensions,
)

logger = logging.getLogger(__name__)

rules_bp = Blueprint("rules", __name__, url_prefix="/api/rules")


def _check_token():
    """Token 认证"""
    token = request.args.get("token", "") or request.headers.get("X-API-Token", "")
    try:
        from dcs_platform.core.config import APP_TOKEN
    except ImportError:
        APP_TOKEN = ""  # config 未加载时拒绝所有请求
    if APP_TOKEN and token != APP_TOKEN:
        return False
    return True


def _unauthorized():
    return jsonify({"error": "Unauthorized"}), 401


# ===== 规则组 CRUD =====

@rules_bp.route("")
def api_list_rules():
    """GET /api/rules?cycle_type=plugging"""
    if not _check_token():
        return _unauthorized()
    cycle_type = request.args.get("cycle_type", "").strip() or None
    groups = get_rule_groups(cycle_type=cycle_type)
    return jsonify({"groups": groups, "count": len(groups)})


@rules_bp.route("/<int:group_id>")
def api_get_rule_group(group_id):
    """GET /api/rules/{id}"""
    if not _check_token():
        return _unauthorized()
    g = get_rule_group(group_id)
    if not g:
        return jsonify({"error": "Rule group not found"}), 404
    return jsonify({"group": g})


@rules_bp.route("", methods=["POST"])
def api_create_rule_group():
    """POST /api/rules
    Body: {cycle_type, name, description, logic_op, priority, rules: [{param_name, operator, threshold_value, ...}]}
    """
    if not _check_token():
        return _unauthorized()

    data = request.get_json(silent=True) or {}
    cycle_type = data.get("cycle_type")
    name = data.get("name", "").strip()

    if not cycle_type or cycle_type not in ("opening", "plugging"):
        return jsonify({"error": "cycle_type must be 'opening' or 'plugging'"}), 400
    if not name:
        return jsonify({"error": "name is required"}), 400

    group_id = upsert_rule_group(
        group_id=None,
        cycle_type=cycle_type,
        name=name,
        description=data.get("description", ""),
        logic_op=data.get("logic_op", "AND"),
        priority=data.get("priority", 0),
        enabled=data.get("enabled", 1),
        detect_config_id=data.get("detect_config_id", 0),
    )

    # 插入子规则
    for rule_data in data.get("rules", []):
        insert_rule(
            group_id=group_id,
            param_name=rule_data.get("param_name", ""),
            operator=rule_data.get("operator", "gt"),
            threshold_value=rule_data.get("threshold_value", 0),
            name=rule_data.get("name", ""),
            threshold_value2=rule_data.get("threshold_value2"),
            duration_s=rule_data.get("duration_s", 0),
            enabled=rule_data.get("enabled", 1),
            priority=rule_data.get("priority", 0),
        )

    g = get_rule_group(group_id)
    return jsonify({"ok": True, "group": g}), 201


@rules_bp.route("/<int:group_id>", methods=["PUT"])
def api_update_rule_group(group_id):
    """PUT /api/rules/{id} — 全量替换"""
    if not _check_token():
        return _unauthorized()

    existing = get_rule_group(group_id)
    if not existing:
        return jsonify({"error": "Rule group not found"}), 404

    data = request.get_json(silent=True) or {}
    group_id = upsert_rule_group(
        group_id=group_id,
        cycle_type=data.get("cycle_type", existing["cycle_type"]),
        name=data.get("name", existing["name"]),
        description=data.get("description", existing.get("description", "")),
        logic_op=data.get("logic_op", existing.get("logic_op", "AND")),
        priority=data.get("priority", existing.get("priority", 0)),
        enabled=data.get("enabled", existing.get("enabled", 1)),
        detect_config_id=data.get("detect_config_id", existing.get("detect_config_id", 0)),
    )

    if "rules" in data:
        clear_rules_in_group(group_id)
        for rule_data in data["rules"]:
            insert_rule(
                group_id=group_id,
                param_name=rule_data.get("param_name", ""),
                operator=rule_data.get("operator", "gt"),
                threshold_value=rule_data.get("threshold_value", 0),
                name=rule_data.get("name", ""),
                threshold_value2=rule_data.get("threshold_value2"),
                duration_s=rule_data.get("duration_s", 0),
                enabled=rule_data.get("enabled", 1),
                priority=rule_data.get("priority", 0),
            )

    g = get_rule_group(group_id)
    return jsonify({"ok": True, "group": g})


@rules_bp.route("/<int:group_id>", methods=["DELETE"])
def api_delete_rule_group(group_id):
    """DELETE /api/rules/{id}"""
    if not _check_token():
        return _unauthorized()
    delete_rule_group(group_id)
    return jsonify({"ok": True})


@rules_bp.route("/<int:group_id>/toggle", methods=["PUT"])
def api_toggle_rule_group(group_id):
    """PUT /api/rules/{id}/toggle  Body: {enabled: 0|1}"""
    if not _check_token():
        return _unauthorized()
    data = request.get_json(silent=True) or {}
    enabled = int(data.get("enabled", 1))
    toggle_rule_group(group_id, enabled)
    return jsonify({"ok": True, "enabled": enabled})


# ===== 规则评估 =====

@rules_bp.route("/evaluate", methods=["POST"])
def api_evaluate_rules():
    """POST /api/rules/evaluate  Body: {cycle_id, cycle_type}"""
    if not _check_token():
        return _unauthorized()

    data = request.get_json(silent=True) or {}
    cycle_type = data.get("cycle_type")
    signal_data = data.get("signal_data", {})

    if not cycle_type:
        return jsonify({"error": "cycle_type is required"}), 400

    try:
        result = evaluate_rules(cycle_type, signal_data)
        return jsonify(result)
    except Exception as e:
        logger.exception("Rule evaluation failed")
        return jsonify({"error": str(e)}), 500


# ===== 综合判定（加权打分） =====

@rules_bp.route("/verdict", methods=["POST"])
def api_rule_verdict():
    """POST /api/rules/verdict  Body: {cycle_type, signal_data, cycle_id?} → 好/需关注/不好"""
    if not _check_token():
        return _unauthorized()

    data = request.get_json(silent=True) or {}
    cycle_type = data.get("cycle_type")
    signal_data = data.get("signal_data", {})
    cycle_id = data.get("cycle_id")

    if not cycle_type:
        return jsonify({"error": "cycle_type is required"}), 400

    try:
        from dcs_platform.services.rule_engine import evaluate_cycle_verdict
        result = evaluate_cycle_verdict(cycle_type, signal_data, cycle_id)
        return jsonify(result)
    except Exception as e:
        logger.exception("Verdict evaluation failed")
        return jsonify({"error": str(e)}), 500


# ===== 规则建议 =====

@rules_bp.route("/suggestions")
def api_rule_suggestions():
    """GET /api/rules/suggestions?cycle_type=plugging&days=30"""
    if not _check_token():
        return _unauthorized()

    cycle_type = request.args.get("cycle_type", "").strip() or None
    days = int(request.args.get("days", 30))

    suggestions = get_rule_suggestions(cycle_type, days)
    return jsonify({"suggestions": suggestions})


# ===== 信号列表（供前端条件编辑器使用） =====

OPENING_SIGNALS_LIST = [
    {"param": "LT_LQFC_68", "label": "推进进油压力", "unit": "MPa", "dim": "push"},
    {"param": "LT_LQFC_85", "label": "推进回油压力", "unit": "MPa", "dim": "push"},
    {"param": "LT_LQFC_67", "label": "开口小车位移", "unit": "mm", "dim": "push"},
    {"param": "LT_LQFC_64", "label": "回转进油压力", "unit": "MPa", "dim": "swing"},
    {"param": "LT_LQFC_74", "label": "回转回油压力", "unit": "MPa", "dim": "swing"},
    {"param": "LT_LQFC_63", "label": "大臂旋转角", "unit": "deg", "dim": "swing"},
    {"param": "LT_LQFC_87", "label": "转钎进油压力", "unit": "MPa", "dim": "drill"},
    {"param": "LT_LQFC_86", "label": "转钎回油压力", "unit": "MPa", "dim": "drill"},
    {"param": "LT_LQFC_88", "label": "冲击进油压力", "unit": "MPa", "dim": "impact"},
    {"param": "LT_LQFC_89", "label": "冲击回油压力", "unit": "MPa", "dim": "impact"},
    {"param": "LT_LQFC_66", "label": "倾动进油压力", "unit": "MPa", "dim": "tilt"},
    {"param": "LT_LQFC_75", "label": "倾动回油压力", "unit": "MPa", "dim": "tilt"},
    {"param": "LT_LQFC_151", "label": "液压站压力", "unit": "MPa", "dim": "hydraulic"},
    {"param": "LT_LQFC_150", "label": "液压站温度", "unit": "°C", "dim": "hydraulic"},
]

PLUGGING_SIGNALS_LIST = [
    {"param": "LT_LQFC_136", "label": "转炮压力", "unit": "MPa", "dim": "cannon"},
    {"param": "LT_LQFC_140", "label": "退炮压力", "unit": "MPa", "dim": "cannon"},
    {"param": "LT_LQFC_138", "label": "打泥压力", "unit": "MPa", "dim": "mud"},
    {"param": "LT_LQFC_139", "label": "退泥压力", "unit": "MPa", "dim": "retreat"},
    {"param": "LT_LQFC_179", "label": "打泥量", "unit": "L", "dim": "mud"},
    {"param": "LT_LQFC_137", "label": "打泥位置", "unit": "mm", "dim": "mud"},
    {"param": "LT_LQFC_151", "label": "液压站压力", "unit": "MPa", "dim": "hydraulic"},
    {"param": "LT_LQFC_150", "label": "液压站温度", "unit": "°C", "dim": "hydraulic"},
]


@rules_bp.route("/signals")
def api_signal_list():
    """GET /api/rules/signals?cycle_type=plugging"""
    if not _check_token():
        return _unauthorized()
    cycle_type = request.args.get("cycle_type", "").strip()
    signals = PLUGGING_SIGNALS_LIST if cycle_type == "plugging" else OPENING_SIGNALS_LIST
    return jsonify({"signals": signals, "cycle_type": cycle_type or "opening"})
