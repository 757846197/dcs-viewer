"""DCS 平台核心模块 — SQLite 数据库"""
import json, logging, sqlite3, threading
from datetime import datetime
from pathlib import Path
from contextlib import contextmanager

logger = logging.getLogger(__name__)

DB_PATH = Path(__file__).resolve().parent.parent.parent / "dcs_analysis.db"
_local = threading.local()


def _get_conn() -> sqlite3.Connection:
    """获取当前线程的 SQLite 连接（惰性创建）。

    注意：该连接由 _local 持有，在线程结束时需显式调用 close_connection() 清理。
    """
    if not hasattr(_local, "conn") or _local.conn is None:
        _local.conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        _local.conn.row_factory = sqlite3.Row
        _local.conn.execute("PRAGMA journal_mode=WAL")
        _local.conn.execute("PRAGMA foreign_keys=ON")
    return _local.conn


def close_connection():
    """关闭当前线程的 SQLite 连接并释放资源。

    线程结束前应调用此函数，避免连接泄漏。重复调用安全（幂等）。
    """
    conn = getattr(_local, "conn", None)
    if conn is not None:
        try:
            conn.close()
        except Exception as e:
            logger.warning("关闭 SQLite 连接失败: %s", e)
        _local.conn = None


@contextmanager
def get_connection():
    """上下文管理器：获取数据库连接，退出时不关闭（复用线程本地连接）。

    用法::

        with get_connection() as conn:
            conn.execute("SELECT ...")

    该连接由线程本地存储管理，线程结束时调用 close_connection() 清理。
    """
    conn = _get_conn()
    try:
        yield conn
    finally:
        # 不在此处关闭连接，因为它被线程本地存储复用；
        # 仅回滚未提交的事务以避免锁定。
        try:
            conn.rollback()
        except Exception as e:
            logger.warning("SQLite 回滚失败: %s", e)

def init_db():
    conn = _get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS cycles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id TEXT NOT NULL,
            cycle_type TEXT NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            duration_s REAL,
            confidence REAL DEFAULT 0.0,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS cycle_stages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL REFERENCES cycles(id),
            stage_name TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            avg_pressure REAL,
            peak_pressure REAL,
            avg_speed REAL,
            metadata TEXT
        );
        CREATE TABLE IF NOT EXISTS labels (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL REFERENCES cycles(id),
            status TEXT,
            success INTEGER,
            labeler TEXT,
            phases TEXT,
            anomaly_tags TEXT,
            had_intervention INTEGER DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS feature_cache (
            cycle_id INTEGER NOT NULL REFERENCES cycles(id),
            feature_name TEXT NOT NULL,
            value REAL,
            computed_at TEXT DEFAULT (datetime('now')),
            PRIMARY KEY (cycle_id, feature_name)
        );
        CREATE INDEX IF NOT EXISTS idx_cycles_equipment ON cycles(equipment_id, start_time);
        CREATE INDEX IF NOT EXISTS idx_labels_cycle ON labels(cycle_id);

        -- 规则引擎
        CREATE TABLE IF NOT EXISTS rule_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_type TEXT NOT NULL CHECK(cycle_type IN ('opening','plugging')),
            name TEXT NOT NULL,
            description TEXT DEFAULT '',
            logic_op TEXT DEFAULT 'AND' CHECK(logic_op IN ('AND','OR')),
            priority INTEGER DEFAULT 0,
            enabled INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL REFERENCES rule_groups(id) ON DELETE CASCADE,
            name TEXT DEFAULT '',
            param_name TEXT NOT NULL,
            operator TEXT NOT NULL CHECK(operator IN ('gt','lt','gte','lte','eq','between','change_rate_gt','change_rate_lt')),
            threshold_value REAL NOT NULL,
            threshold_value2 REAL,
            duration_s REAL DEFAULT 0,
            enabled INTEGER DEFAULT 1,
            priority INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS rule_eval_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rule_id INTEGER REFERENCES rules(id) ON DELETE SET NULL,
            group_id INTEGER REFERENCES rule_groups(id) ON DELETE SET NULL,
            cycle_id INTEGER REFERENCES cycles(id) ON DELETE SET NULL,
            cycle_type TEXT,
            eval_result TEXT CHECK(eval_result IN ('pass','fail')),
            detail TEXT,
            eval_time TEXT DEFAULT (datetime('now'))
        );
        CREATE INDEX IF NOT EXISTS idx_rules_group ON rules(group_id);
        CREATE INDEX IF NOT EXISTS idx_rule_groups_type ON rule_groups(cycle_type);
        CREATE INDEX IF NOT EXISTS idx_rule_eval_cycle ON rule_eval_log(cycle_id);

        -- 知识库
        CREATE TABLE IF NOT EXISTS knowledge_base (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL UNIQUE REFERENCES cycles(id),
            cycle_type TEXT NOT NULL,
            title TEXT,
            tags TEXT DEFAULT '[]',
            status TEXT DEFAULT 'typical' CHECK(status IN ('exemplary','typical','edge_case','anomaly','archived')),
            operator_notes TEXT DEFAULT '',
            quality_score INTEGER DEFAULT 3 CHECK(quality_score BETWEEN 1 AND 5),
            source TEXT DEFAULT 'manual',
            added_by TEXT DEFAULT 'admin',
            added_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE INDEX IF NOT EXISTS idx_kb_cycle_type ON knowledge_base(cycle_type);
        CREATE INDEX IF NOT EXISTS idx_kb_status ON knowledge_base(status);

        -- 周期检测配置
        CREATE TABLE IF NOT EXISTS detect_configs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            cycle_type TEXT NOT NULL CHECK(cycle_type IN ('opening','plugging','all')),
            description TEXT DEFAULT '',
            config_json TEXT NOT NULL,       -- JSON: 检测规则定义
            enabled INTEGER DEFAULT 1,
            is_default INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );

        -- 判定规则配置（结果分类: 成功/失败/未完成/未完整/钻透/铁口深度）
        CREATE TABLE IF NOT EXISTS result_judge_configs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,                    -- 配置名称
            cycle_type TEXT NOT NULL CHECK(cycle_type IN ('opening','plugging')),
            category TEXT NOT NULL CHECK(category IN ('success','fail','incomplete','unfinished','breakthrough','depth')),
            description TEXT DEFAULT '',            -- 判定条件描述
            params_json TEXT NOT NULL,              -- [{param_name, value, unit, data_type, range_min, range_max}]
            logic_op TEXT DEFAULT 'AND' CHECK(logic_op IN ('AND','OR')),
            enabled INTEGER DEFAULT 1,
            is_default INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );

        -- 自整定运行记录
        CREATE TABLE IF NOT EXISTS tuning_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id INTEGER,                    -- result_judge_configs.id (判定规则)
            cycle_type TEXT NOT NULL,
            status TEXT DEFAULT 'running' CHECK(status IN ('running','completed','failed')),
            run_mode TEXT DEFAULT 'auto' CHECK(run_mode IN ('auto','manual')),
            started_at TEXT DEFAULT (datetime('now')),
            finished_at TEXT,
            -- 各阶段结果 (JSON)
            collect_stats TEXT,    -- {signal_count, time_range, ...}
            analysis_result TEXT,  -- {features: [...], patterns: [...]}
            tuned_params TEXT,     -- {rules: [...], filter_min_s, filter_max_s}
            eval_result TEXT,      -- {accuracy, false_positive_rate, samples, ...}
            error_message TEXT
        );

        CREATE TABLE IF NOT EXISTS tuning_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            config_id INTEGER REFERENCES detect_configs(id),
            run_id INTEGER REFERENCES tuning_runs(id),
            param_name TEXT NOT NULL,
            old_value TEXT,
            new_value TEXT NOT NULL,
            change_reason TEXT,
            changed_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS tuning_config (
            id INTEGER PRIMARY KEY CHECK(id=1),
            auto_mode INTEGER DEFAULT 1,        -- 0=手动, 1=自动
            schedule_hour INTEGER DEFAULT 2,     -- 每日触发时间(UTC)
            eval_min_samples INTEGER DEFAULT 10, -- 评估最少样本数
            min_accuracy REAL DEFAULT 0.7,       -- 最低准确率要求
            max_false_rate REAL DEFAULT 0.15,    -- 最大误报率
            updated_at TEXT DEFAULT (datetime('now'))
        );

        INSERT OR IGNORE INTO tuning_config(id) VALUES(1);

        -- 编码器校准表（动态基线/偏移校正）
        CREATE TABLE IF NOT EXISTS variable_configs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tag_name TEXT NOT NULL UNIQUE,          -- InfluxDB 信号名 (LT_LQFC_XX)
            chinese_name TEXT NOT NULL,              -- 中文名称
            unit TEXT DEFAULT '',                    -- 单位 (mm, MPa, deg, etc.)
            cycle_type TEXT NOT NULL DEFAULT 'common' CHECK(cycle_type IN ('opening','plugging','common')),
            equipment TEXT DEFAULT '',               -- 所属设备: east_opener/west_opener/east_plugger/west_plugger
            dimension TEXT DEFAULT '',               -- 维度分类: remote_command/control_valve/position/pressure/drill/impact/mud/hydraulic
            data_type TEXT DEFAULT 'float' CHECK(data_type IN ('float','int','bool')),
            description TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE INDEX IF NOT EXISTS idx_varconfig_cycle ON variable_configs(cycle_type);
        CREATE INDEX IF NOT EXISTS idx_varconfig_tag ON variable_configs(tag_name);

        CREATE TABLE IF NOT EXISTS encoder_calibration (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL UNIQUE,          -- 设备名: 东开口机/西开口机/东堵口机/西堵口机
            cycle_type TEXT NOT NULL,              -- opening/plugging
            position_signal TEXT NOT NULL,          -- 小车位置信号名
            offset_baseline REAL DEFAULT 0.0,      -- 编码器偏移基线（m）
            travel_range_min REAL DEFAULT 0.0,     -- 最小有效行程（m）
            travel_range_max REAL DEFAULT 3.0,     -- 最大有效行程（m）
            slope_correction REAL DEFAULT 1.0,     -- 斜率修正系数（编码器比例误差）
            last_calibrated_at TEXT,
            description TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );

        -- 默认编码器校准
        INSERT OR IGNORE INTO encoder_calibration(machine, cycle_type, position_signal, offset_baseline, travel_range_min, travel_range_max)
        VALUES
            ('东开口机', 'opening', 'LT_LQFC_67', 0.0, 0.1, 2.5),
            ('西开口机', 'opening', 'LT_LQFC_104', 0.0, 0.1, 2.5),
            ('东堵口机', 'plugging', 'LT_LQFC_137', 0.0, 0.1, 2.0),
            ('西堵口机', 'plugging', 'LT_LQFC_160', 0.0, 0.1, 2.0);

        -- 模型训练记录
        CREATE TABLE IF NOT EXISTS model_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_type TEXT NOT NULL CHECK(cycle_type IN ('opening','plugging')),
            trained_at TEXT DEFAULT (datetime('now')),
            n_samples INTEGER,
            n_features INTEGER,
            best_score REAL,
            cv_folds INTEGER DEFAULT 5,
            model_path TEXT,
            meta_json TEXT,
            knowledge_base_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'running' CHECK(status IN ('running','completed','failed'))
        );
    """)
    conn.commit()

def insert_cycle(equipment_id, cycle_type, start_time, end_time, duration_s, confidence=0.0):
    c = _get_conn().execute(
        "INSERT INTO cycles(equipment_id,cycle_type,start_time,end_time,duration_s,confidence) VALUES(?,?,?,?,?,?)",
        (equipment_id, cycle_type, str(start_time), str(end_time), duration_s, confidence))
    _get_conn().commit()
    return c.lastrowid

def insert_stage(cycle_id, stage_name, start_time, end_time, avg_pressure=None, peak_pressure=None, avg_speed=None, metadata=None):
    _get_conn().execute(
        "INSERT INTO cycle_stages(cycle_id,stage_name,start_time,end_time,avg_pressure,peak_pressure,avg_speed,metadata) VALUES(?,?,?,?,?,?,?,?)",
        (cycle_id, stage_name, str(start_time) if start_time else None, str(end_time) if end_time else None,
         avg_pressure, peak_pressure, avg_speed, json.dumps(metadata) if metadata else None))
    _get_conn().commit()

def upsert_label(cycle_id, status=None, success=None, labeler=None, phases=None, anomaly_tags=None, had_intervention=0, notes=None):
    conn = _get_conn()
    existing = conn.execute("SELECT id FROM labels WHERE cycle_id=?", (cycle_id,)).fetchone()
    if existing:
        conn.execute(
            "UPDATE labels SET status=?,success=?,labeler=?,phases=?,anomaly_tags=?,had_intervention=?,notes=? WHERE cycle_id=?",
            (status, success, labeler, json.dumps(phases) if phases else None,
             json.dumps(anomaly_tags) if anomaly_tags else None, had_intervention, notes, cycle_id))
    else:
        conn.execute(
            "INSERT INTO labels(cycle_id,status,success,labeler,phases,anomaly_tags,had_intervention,notes) VALUES(?,?,?,?,?,?,?,?)",
            (cycle_id, status, success, labeler, json.dumps(phases) if phases else None,
             json.dumps(anomaly_tags) if anomaly_tags else None, had_intervention, notes))
    conn.commit()

def upsert_feature(cycle_id, feature_name, value):
    _get_conn().execute(
        "INSERT OR REPLACE INTO feature_cache(cycle_id,feature_name,value) VALUES(?,?,?)",
        (cycle_id, feature_name, value))
    _get_conn().commit()

def get_cycles(equipment_id=None, cycle_type=None, limit=100, offset=0):
    """查询 cycles，支持分页（LIMIT + OFFSET）。

    参数:
        equipment_id: 可选设备筛选
        cycle_type: 可选循环类型筛选
        limit: 返回条数上限（默认 100）
        offset: 跳过的条数（默认 0，用于分页）
    返回:
        list[dict]
    """
    q = "SELECT * FROM cycles WHERE 1=1"
    args = []
    if equipment_id:
        q += " AND equipment_id=?"; args.append(equipment_id)
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    q += " ORDER BY start_time DESC LIMIT ? OFFSET ?"
    args.append(limit)
    args.append(offset)
    return [dict(r) for r in _get_conn().execute(q, args).fetchall()]


def count_cycles(equipment_id=None, cycle_type=None):
    """返回匹配条件的 cycles 总数（用于分页）。"""
    q = "SELECT COUNT(*) as n FROM cycles WHERE 1=1"
    args = []
    if equipment_id:
        q += " AND equipment_id=?"; args.append(equipment_id)
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    return _get_conn().execute(q, args).fetchone()["n"]


def get_label_stats():
    conn = _get_conn()
    total = conn.execute("SELECT COUNT(*) as n FROM cycles").fetchone()["n"]
    labeled = conn.execute("SELECT COUNT(DISTINCT cycle_id) as n FROM labels WHERE status IS NOT NULL").fetchone()["n"]
    return {"total": total, "labeled": labeled, "unlabeled": max(0, total - labeled)}

def get_features(cycle_id):
    return {r["feature_name"]: r["value"] for r in
            _get_conn().execute("SELECT feature_name,value FROM feature_cache WHERE cycle_id=?", (cycle_id,)).fetchall()}

def export_labels(cycle_type=None):
    q = """SELECT c.*, l.status, l.success, l.phases, l.anomaly_tags, l.had_intervention, l.notes
           FROM cycles c LEFT JOIN labels l ON c.id=l.cycle_id WHERE l.status IS NOT NULL"""
    args = []
    if cycle_type: q += " AND c.cycle_type=?"; args.append(cycle_type)
    q += " ORDER BY c.start_time"
    return [dict(r) for r in _get_conn().execute(q, args).fetchall()]


# ===================================================================
#  规则引擎 CRUD
# ===================================================================

def upsert_rule_group(group_id, cycle_type, name, description="", logic_op="AND", priority=0, enabled=1):
    conn = _get_conn()
    if group_id:
        conn.execute(
            "UPDATE rule_groups SET cycle_type=?,name=?,description=?,logic_op=?,priority=?,enabled=?,updated_at=datetime('now') WHERE id=?",
            (cycle_type, name, description, logic_op, priority, enabled, group_id))
        conn.commit()
        return group_id
    c = conn.execute(
        "INSERT INTO rule_groups(cycle_type,name,description,logic_op,priority,enabled) VALUES(?,?,?,?,?,?)",
        (cycle_type, name, description, logic_op, priority, enabled))
    conn.commit()
    return c.lastrowid


def delete_rule_group(group_id):
    _get_conn().execute("DELETE FROM rules WHERE group_id=?", (group_id,))
    _get_conn().execute("DELETE FROM rule_groups WHERE id=?", (group_id,))
    _get_conn().commit()


def toggle_rule_group(group_id, enabled):
    _get_conn().execute("UPDATE rule_groups SET enabled=?,updated_at=datetime('now') WHERE id=?", (enabled, group_id))
    _get_conn().commit()


def insert_rule(group_id, param_name, operator, threshold_value, name="", threshold_value2=None, duration_s=0, enabled=1, priority=0):
    c = _get_conn().execute(
        "INSERT INTO rules(group_id,name,param_name,operator,threshold_value,threshold_value2,duration_s,enabled,priority) VALUES(?,?,?,?,?,?,?,?,?)",
        (group_id, name, param_name, operator, threshold_value, threshold_value2, duration_s, enabled, priority))
    _get_conn().commit()
    return c.lastrowid


def clear_rules_in_group(group_id):
    _get_conn().execute("DELETE FROM rules WHERE group_id=?", (group_id,))
    _get_conn().commit()


def get_rule_groups(cycle_type=None):
    q = "SELECT * FROM rule_groups WHERE 1=1"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"
        args.append(cycle_type)
    q += " ORDER BY priority DESC"
    groups = [dict(r) for r in _get_conn().execute(q, args).fetchall()]
    for g in groups:
        g["rules"] = [dict(r) for r in _get_conn().execute(
            "SELECT * FROM rules WHERE group_id=? ORDER BY priority DESC", (g["id"],)).fetchall()]
    return groups


def get_rule_group(group_id):
    g = _get_conn().execute("SELECT * FROM rule_groups WHERE id=?", (group_id,)).fetchone()
    if not g:
        return None
    g = dict(g)
    g["rules"] = [dict(r) for r in _get_conn().execute(
        "SELECT * FROM rules WHERE group_id=? ORDER BY priority DESC", (g["id"],)).fetchall()]
    return g


def insert_rule_eval_log(rule_id, group_id, cycle_id, cycle_type, eval_result, detail=""):
    _get_conn().execute(
        "INSERT INTO rule_eval_log(rule_id,group_id,cycle_id,cycle_type,eval_result,detail) VALUES(?,?,?,?,?,?)",
        (rule_id, group_id, cycle_id, cycle_type, eval_result, detail))
    _get_conn().commit()


# ===================================================================
#  知识库 CRUD
# ===================================================================

def insert_knowledge_entry(cycle_id, cycle_type, title=None, tags=None, status="typical", operator_notes="", quality_score=3, added_by="admin"):
    conn = _get_conn()
    existing = conn.execute("SELECT id FROM knowledge_base WHERE cycle_id=?", (cycle_id,)).fetchone()
    if existing:
        raise ValueError(f"cycle_id {cycle_id} already exists in knowledge_base")
    tags_json = json.dumps(tags) if tags else "[]"
    auto_title = title or f"{cycle_type}_{cycle_id}"
    c = conn.execute(
        "INSERT INTO knowledge_base(cycle_id,cycle_type,title,tags,status,operator_notes,quality_score,added_by) VALUES(?,?,?,?,?,?,?,?)",
        (cycle_id, cycle_type, auto_title, tags_json, status, operator_notes, quality_score, added_by))
    conn.commit()
    return c.lastrowid


def update_knowledge_entry(entry_id, title=None, tags=None, status=None, operator_notes=None, quality_score=None):
    conn = _get_conn()
    updates = []
    args = []
    if title is not None:
        updates.append("title=?"); args.append(title)
    if tags is not None:
        updates.append("tags=?"); args.append(json.dumps(tags))
    if status is not None:
        updates.append("status=?"); args.append(status)
    if operator_notes is not None:
        updates.append("operator_notes=?"); args.append(operator_notes)
    if quality_score is not None:
        updates.append("quality_score=?"); args.append(quality_score)
    if updates:
        updates.append("updated_at=datetime('now')")
        args.append(entry_id)
        conn.execute(f"UPDATE knowledge_base SET {','.join(updates)} WHERE id=?", args)
        conn.commit()


def archive_knowledge_entry(entry_id):
    _get_conn().execute("UPDATE knowledge_base SET status='archived',updated_at=datetime('now') WHERE id=?", (entry_id,))
    _get_conn().commit()


def get_knowledge_entries(cycle_type=None, status=None, tags=None, page=1, per_page=50):
    q = """SELECT kb.*, c.equipment_id, c.start_time, c.end_time, c.duration_s
           FROM knowledge_base kb LEFT JOIN cycles c ON kb.cycle_id=c.id WHERE 1=1"""
    args = []
    if cycle_type:
        q += " AND kb.cycle_type=?"; args.append(cycle_type)
    if status:
        q += " AND kb.status=?"; args.append(status)
    if tags:
        for t in tags:
            q += " AND kb.tags LIKE ?"; args.append(f"%{t}%")
    q += " ORDER BY kb.added_at DESC"
    if per_page:
        q += " LIMIT ? OFFSET ?"
        args.append(per_page)
        args.append((page - 1) * per_page)
    rows = [dict(r) for r in _get_conn().execute(q, args).fetchall()]
    for r in rows:
        try:
            r["tags"] = json.loads(r.get("tags", "[]"))
        except (json.JSONDecodeError, TypeError):
            r["tags"] = []
    return rows


def get_knowledge_entry(entry_id):
    r = _get_conn().execute(
        "SELECT kb.*, c.equipment_id, c.start_time, c.end_time, c.duration_s FROM knowledge_base kb LEFT JOIN cycles c ON kb.cycle_id=c.id WHERE kb.id=?",
        (entry_id,)).fetchone()
    if r:
        r = dict(r)
        try:
            r["tags"] = json.loads(r.get("tags", "[]"))
        except (json.JSONDecodeError, TypeError):
            r["tags"] = []
        return r
    return None


def get_knowledge_stats(cycle_type=None):
    conn = _get_conn()
    q = "SELECT cycle_type, status, COUNT(*) as cnt FROM knowledge_base WHERE status != 'archived'"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    q += " GROUP BY cycle_type, status"
    rows = conn.execute(q, args).fetchall()
    total = conn.execute("SELECT COUNT(*) FROM knowledge_base WHERE status != 'archived'" + (" AND cycle_type=?" if cycle_type else ""),
                         (cycle_type,) if cycle_type else ()).fetchone()[0]
    by_status = {}
    for r in rows:
        ct = r["cycle_type"]
        if ct not in by_status:
            by_status[ct] = {}
        by_status[ct][r["status"]] = r["cnt"]
    return {"total": total, "by_type_status": by_status}


# ===================================================================
#  模型训练记录
# ===================================================================

def insert_model_run(cycle_type, n_samples=None, n_features=None, best_score=None, cv_folds=5,
                     model_path=None, meta_json=None, knowledge_base_count=0, status="running"):
    c = _get_conn().execute(
        "INSERT INTO model_runs(cycle_type,n_samples,n_features,best_score,cv_folds,model_path,meta_json,knowledge_base_count,status) VALUES(?,?,?,?,?,?,?,?,?)",
        (cycle_type, n_samples, n_features, best_score, cv_folds, model_path, meta_json, knowledge_base_count, status))
    _get_conn().commit()
    return c.lastrowid


def update_model_run(run_id, **kwargs):
    allowed = ["n_samples", "n_features", "best_score", "model_path", "meta_json", "knowledge_base_count", "status"]
    updates = []
    args = []
    for k in allowed:
        if k in kwargs:
            updates.append(f"{k}=?")
            args.append(kwargs[k])
    if updates:
        args.append(run_id)
        _get_conn().execute(f"UPDATE model_runs SET {','.join(updates)} WHERE id=?", args)
        _get_conn().commit()


def get_model_runs(cycle_type=None, limit=20):
    q = "SELECT * FROM model_runs WHERE 1=1"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    q += " ORDER BY trained_at DESC LIMIT ?"; args.append(limit)
    return [dict(r) for r in _get_conn().execute(q, args).fetchall()]


def get_model_run(run_id):
    r = _get_conn().execute("SELECT * FROM model_runs WHERE id=?", (run_id,)).fetchone()
    return dict(r) if r else None


# ===================================================================
#  周期检测配置 CRUD
# ===================================================================

_DEFAULT_OPENING_CONFIG = {
    "type": "opening",
    "rules": [
        {"signal": "LT_LQFC_57", "role": "remote",     "label": "遥控选择"},
        {"signal": "LT_LQFC_67", "role": "threshold",  "label": "推进位置前进(>-0.5m)", "threshold": -0.5,  "operator": "gt"},
        {"signal": "LT_LQFC_68", "role": "threshold",  "label": "推进压力(>3MPa)",       "threshold": 3.0,   "operator": "gt"},
    ],
    "filter_min_s": 30,
    "filter_max_s": 3600
}

_DEFAULT_PLUGGING_CONFIG = {
    "type": "plugging",
    "rules": [
        {"signal": "LT_LQFC_130", "role": "remote",    "label": "遥控启动"},
        {"signal": "LT_LQFC_135", "role": "threshold", "label": "回转到位(<50度)",    "threshold": 50.0, "operator": "lt"},
        {"signal": "LT_LQFC_138", "role": "threshold", "label": "打泥压力(>5MPa)",     "threshold": 5.0,  "operator": "gt"},
        {"signal": "LT_LQFC_134", "role": "threshold", "label": "打泥指令激活",        "threshold": 0.5,  "operator": "gt"},
    ],
    "filter_min_s": 30,
    "filter_max_s": 3600
}

def get_detect_configs(cycle_type=None):
    """获取所有检测配置"""
    q = "SELECT * FROM detect_configs WHERE 1=1"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    q += " ORDER BY is_default DESC, name ASC"
    rows = _get_conn().execute(q, args).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        try: d["config"] = json.loads(d["config_json"])
        except: d["config"] = {}
        del d["config_json"]
        result.append(d)
    return result

def get_detect_config(config_id):
    r = _get_conn().execute("SELECT * FROM detect_configs WHERE id=?", (config_id,)).fetchone()
    if not r: return None
    d = dict(r)
    try: d["config"] = json.loads(d["config_json"])
    except: d["config"] = {}
    del d["config_json"]
    return d

def get_default_detect_config(cycle_type):
    """获取指定类型的默认检测配置。无默认时自动创建。"""
    r = _get_conn().execute(
        "SELECT * FROM detect_configs WHERE cycle_type=? AND is_default=1 AND enabled=1 LIMIT 1",
        (cycle_type,)
    ).fetchone()
    if r:
        d = dict(r)
        try: d["config"] = json.loads(d["config_json"])
        except: d["config"] = {}
        del d["config_json"]
        return d
    # 自动创建默认配置
    if cycle_type == "opening":
        cfg = _DEFAULT_OPENING_CONFIG
        name = "默认开口检测"
    else:
        cfg = _DEFAULT_PLUGGING_CONFIG
        name = "默认堵口检测"
    config_id = upsert_detect_config(None, name, cycle_type, json.dumps(cfg, ensure_ascii=False), "", is_default=1)
    return get_detect_config(config_id)

def upsert_detect_config(config_id, name, cycle_type, config_json, description="", is_default=0):
    conn = _get_conn()
    if config_id:
        conn.execute(
            "UPDATE detect_configs SET name=?,cycle_type=?,config_json=?,description=?,is_default=?,updated_at=datetime('now') WHERE id=?",
            (name, cycle_type, config_json, description, is_default, config_id))
        conn.commit()
        return config_id
    # 如果设为首选，清除其他同类型的首选
    if is_default:
        conn.execute("UPDATE detect_configs SET is_default=0 WHERE cycle_type=?", (cycle_type,))
    c = conn.execute(
        "INSERT INTO detect_configs(name,cycle_type,config_json,description,is_default) VALUES(?,?,?,?,?)",
        (name, cycle_type, config_json, description, is_default))
    conn.commit()
    return c.lastrowid

def toggle_detect_config(config_id, enabled):
    _get_conn().execute("UPDATE detect_configs SET enabled=?,updated_at=datetime('now') WHERE id=?", (enabled, config_id))
    _get_conn().commit()

def delete_detect_config(config_id):
    _get_conn().execute("DELETE FROM detect_configs WHERE id=?", (config_id,))
    _get_conn().commit()


# ===================================================================
#  自整定 CRUD
# ===================================================================

# ===================================================================
#  判定规则配置 CRUD
# ===================================================================

_DEFAULT_RESULT_PARAMS = {
    # === 开口作业判定规则 ===
    # 信号源: 东开口机 LT_LQFC_57~89, 西开口机 LT_LQFC_94~125
    # 核心信号: 67=推进位置(行程), 68=推进压力, 69=冲击指令, 63=回转位置
    "opening": {
        "success": {"name": "开口成功判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_67", "value": 0.5, "unit": "m", "data_type": "float",
             "range_min": 0.1, "range_max": 2.0, "label": "推进位置位移量 (max-min)", "operator": "gte"},
            {"param_name": "LT_LQFC_68", "value": 0.15, "unit": "ratio", "data_type": "float",
             "range_min": 0.05, "range_max": 0.5, "label": "推进压力骤降比 (晚期/早期)", "operator": "lt"},
        ]},
        "fail": {"name": "开口失败判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_67", "value": 0.1, "unit": "m", "data_type": "float",
             "range_min": 0.01, "range_max": 0.3, "label": "推进位置位移上限 (无有效钻进)", "operator": "lt"},
        ]},
        "incomplete": {"name": "开口未完成判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_67", "value": 0.1, "unit": "m", "data_type": "float",
             "range_min": 0.05, "range_max": 1.0, "label": "推进位置位移下限 (有钻进)", "operator": "gte"},
            {"param_name": "LT_LQFC_68", "value": 0.15, "unit": "ratio", "data_type": "float",
             "range_min": 0.05, "range_max": 0.5, "label": "推进压力骤降比 (未达标,未钻透)", "operator": "gte"},
        ]},
        "breakthrough": {"name": "钻透判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_67", "value": 0.6, "unit": "ratio", "data_type": "float",
             "range_min": 0.3, "range_max": 0.95, "label": "推进行程占比 (实际/全量程)", "operator": "gte"},
            {"param_name": "LT_LQFC_68", "value": 0.2, "unit": "ratio", "data_type": "float",
             "range_min": 0.1, "range_max": 0.5, "label": "推进压力骤降比 (晚期/早期≤20%)", "operator": "lt"},
            {"param_name": "LT_LQFC_69", "value": 1, "unit": "bool", "data_type": "int",
             "range_min": 0, "range_max": 1, "label": "冲击指令已激活", "operator": "eq"},
        ]},
        "depth": {"name": "铁口深度计算", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_67", "value": 1.0, "unit": "m", "data_type": "float",
             "range_min": 0.5, "range_max": 3.0, "label": "推进有效行程下限 (max-基线)", "operator": "gte"},
            {"param_name": "LT_LQFC_63", "value": 30, "unit": "deg", "data_type": "float",
             "range_min": 10, "range_max": 90, "label": "大臂到位角度上限", "operator": "lt"},
            {"param_name": "encoder_offset_calib", "value": 1, "unit": "bool", "data_type": "int",
             "range_min": 0, "range_max": 1, "label": "编码器偏移自动校正", "operator": "eq"},
        ]},
    },
    # === 堵口作业判定规则 ===
    # 信号源: 东堵口机 LT_LQFC_129~151,179, 西堵口机 LT_LQFC_152~167,180
    # 核心信号: 179/180=打泥量, 138/161=打泥压力, 137/160=打泥位置(行程), 134/157=打泥指令
    "plugging": {
        "success": {"name": "堵口成功判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_179", "value": 100, "unit": "L", "data_type": "float",
             "range_min": 50, "range_max": 500, "label": "打泥总量 (max-min)", "operator": "gte"},
            {"param_name": "LT_LQFC_134", "value": 60, "unit": "s", "data_type": "int",
             "range_min": 30, "range_max": 300, "label": "保压时长下限", "operator": "gte"},
        ]},
        "fail": {"name": "堵口失败判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_179", "value": 50, "unit": "L", "data_type": "float",
             "range_min": 10, "range_max": 200, "label": "打泥量下限 (严重不足)", "operator": "lt"},
        ]},
        "unfinished": {"name": "堵口未完整判定", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_179", "value": 50, "unit": "L", "data_type": "float",
             "range_min": 10, "range_max": 200, "label": "打泥量下限 (部分完成)", "operator": "gte"},
            {"param_name": "LT_LQFC_134", "value": 60, "unit": "s", "data_type": "int",
             "range_min": 10, "range_max": 300, "label": "保压时长上限 (不足)", "operator": "lt"},
        ]},
        "breakthrough": {"name": "钻透判定 (泥炮到位)", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_137", "value": 0.8, "unit": "ratio", "data_type": "float",
             "range_min": 0.5, "range_max": 1.0, "label": "打泥位置到位比 (行程/工作位)", "operator": "gte"},
            {"param_name": "LT_LQFC_138", "value": 15.0, "unit": "MPa", "data_type": "float",
             "range_min": 5.0, "range_max": 30.0, "label": "打泥压力峰值", "operator": "gte"},
        ]},
        "depth": {"name": "铁口深度计算 (堵口)", "logic": "AND", "params": [
            {"param_name": "LT_LQFC_137", "value": 100, "unit": "mm", "data_type": "float",
             "range_min": 50, "range_max": 500, "label": "打泥行程下限 (max-基线)", "operator": "gte"},
            {"param_name": "LT_LQFC_138", "value": 10.0, "unit": "MPa", "data_type": "float",
             "range_min": 5.0, "range_max": 25.0, "label": "打泥平均压力下限", "operator": "gte"},
        ]},
    }
}

def get_result_judge_configs(cycle_type=None, category=None):
    q = "SELECT * FROM result_judge_configs WHERE 1=1"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    if category:
        q += " AND category=?"; args.append(category)
    q += " ORDER BY cycle_type, category"
    rows = _get_conn().execute(q, args).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        try: d["params"] = json.loads(d["params_json"])
        except: d["params"] = []
        del d["params_json"]
        result.append(d)
    return result

def get_result_judge_config(config_id):
    r = _get_conn().execute("SELECT * FROM result_judge_configs WHERE id=?", (config_id,)).fetchone()
    if not r: return None
    d = dict(r)
    try: d["params"] = json.loads(d["params_json"])
    except: d["params"] = []
    del d["params_json"]
    return d

def get_default_result_config(cycle_type, category):
    r = _get_conn().execute(
        "SELECT * FROM result_judge_configs WHERE cycle_type=? AND category=? AND is_default=1 AND enabled=1 LIMIT 1",
        (cycle_type, category)).fetchone()
    if r:
        d = dict(r)
        try: d["params"] = json.loads(d["params_json"])
        except: d["params"] = []
        del d["params_json"]
        return d
    # 首次访问自动创建默认配置
    defaults = _DEFAULT_RESULT_PARAMS.get(cycle_type, {}).get(category)
    if defaults:
        config_id = upsert_result_judge_config(
            None, defaults["name"], cycle_type, category,
            json.dumps(defaults["params"], ensure_ascii=False),
            defaults.get("logic", "AND"), is_default=1
        )
        return get_result_judge_config(config_id)
    return None

def upsert_result_judge_config(config_id, name, cycle_type, category, params_json, logic_op="AND", is_default=0, description=""):
    conn = _get_conn()
    if config_id:
        conn.execute(
            "UPDATE result_judge_configs SET name=?,params_json=?,logic_op=?,is_default=?,description=?,updated_at=datetime('now') WHERE id=?",
            (name, params_json, logic_op, is_default, description, config_id))
        conn.commit()
        return config_id
    c = conn.execute(
        "INSERT INTO result_judge_configs(name,cycle_type,category,params_json,logic_op,is_default,description) VALUES(?,?,?,?,?,?,?)",
        (name, cycle_type, category, params_json, logic_op, is_default, description))
    conn.commit()
    return c.lastrowid

def toggle_result_judge_config(config_id, enabled):
    _get_conn().execute("UPDATE result_judge_configs SET enabled=?,updated_at=datetime('now') WHERE id=?", (enabled, config_id))
    _get_conn().commit()

def delete_result_judge_config(config_id):
    _get_conn().execute("DELETE FROM result_judge_configs WHERE id=?", (config_id,))
    _get_conn().commit()

def seed_default_result_configs():
    """初始化默认判定规则（已存在则跳过）"""
    existing = get_result_judge_configs()
    if existing:
        return
    for cycle_type in ("opening", "plugging"):
        for category in _DEFAULT_RESULT_PARAMS[cycle_type]:
            get_default_result_config(cycle_type, category)

def get_tuning_config():
    r = _get_conn().execute("SELECT * FROM tuning_config WHERE id=1").fetchone()
    return dict(r) if r else {"auto_mode": 1, "schedule_hour": 2, "eval_min_samples": 10, "min_accuracy": 0.7, "max_false_rate": 0.15}

def update_tuning_config(**kwargs):
    allowed = ["auto_mode", "schedule_hour", "eval_min_samples", "min_accuracy", "max_false_rate"]
    updates = []
    args = []
    for k in allowed:
        if k in kwargs:
            updates.append(f"{k}=?")
            args.append(kwargs[k])
    if updates:
        updates.append("updated_at=datetime('now')")
        _get_conn().execute(f"UPDATE tuning_config SET {','.join(updates)} WHERE id=1", args)
        _get_conn().commit()

def insert_tuning_run(config_id, cycle_type, run_mode="auto"):
    c = _get_conn().execute(
        "INSERT INTO tuning_runs(config_id,cycle_type,status,run_mode) VALUES(?,?,?,?)",
        (config_id, cycle_type, "running", run_mode))
    _get_conn().commit()
    return c.lastrowid

def update_tuning_run(run_id, **kwargs):
    allowed = ["status", "finished_at", "collect_stats", "analysis_result", "tuned_params", "eval_result", "error_message"]
    updates = []
    args = []
    for k in allowed:
        if k in kwargs and kwargs[k] is not None:
            updates.append(f"{k}=?")
            args.append(json.dumps(kwargs[k]) if isinstance(kwargs[k], (dict, list)) else kwargs[k])
    if updates:
        args.append(run_id)
        _get_conn().execute(f"UPDATE tuning_runs SET {','.join(updates)} WHERE id=?", args)
        _get_conn().commit()

def get_tuning_runs(cycle_type=None, limit=20):
    q = "SELECT * FROM tuning_runs WHERE 1=1"
    args = []
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    q += " ORDER BY started_at DESC LIMIT ?"; args.append(limit)
    rows = _get_conn().execute(q, args).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        for f in ("collect_stats", "analysis_result", "tuned_params", "eval_result"):
            try: d[f] = json.loads(d[f]) if d.get(f) else None
            except: pass
        result.append(d)
    return result

def get_tuning_run(run_id):
    r = _get_conn().execute("SELECT * FROM tuning_runs WHERE id=?", (run_id,)).fetchone()
    if not r: return None
    d = dict(r)
    for f in ("collect_stats", "analysis_result", "tuned_params", "eval_result"):
        try: d[f] = json.loads(d[f]) if d.get(f) else None
        except: pass
    return d

def insert_tuning_history(config_id, run_id, param_name, old_value, new_value, change_reason=""):
    _get_conn().execute(
        "INSERT INTO tuning_history(config_id,run_id,param_name,old_value,new_value,change_reason) VALUES(?,?,?,?,?,?)",
        (config_id, run_id, param_name, str(old_value) if old_value else "", str(new_value), change_reason))
    _get_conn().commit()

def get_tuning_history(config_id=None, limit=50):
    q = "SELECT * FROM tuning_history WHERE 1=1"
    args = []
    if config_id:
        q += " AND config_id=?"; args.append(config_id)
    q += " ORDER BY changed_at DESC LIMIT ?"; args.append(limit)
    return [dict(r) for r in _get_conn().execute(q, args).fetchall()]

def _migrate_result_judge_configs():
    """迁移 result_judge_configs 表: 添加 breakthrough/depth 分类"""
    conn = _get_conn()
    # 检查是否已有新分类
    try:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(result_judge_configs)").fetchall()]
    except Exception:
        return  # 表不存在, CREATE TABLE IF NOT EXISTS 会处理
    if 'category' not in cols:
        return
    
    # 检查现有约束
    info = conn.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='result_judge_configs'").fetchone()
    if not info:
        return
    sql_def = info[0]
    if 'breakthrough' in sql_def and 'depth' in sql_def:
        return  # 已迁移
    
    # 执行迁移: 备份旧表 → 创建新表 → 恢复数据
    conn.execute("ALTER TABLE result_judge_configs RENAME TO result_judge_configs_old")
    conn.execute("""CREATE TABLE result_judge_configs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        cycle_type TEXT NOT NULL CHECK(cycle_type IN ('opening','plugging')),
        category TEXT NOT NULL CHECK(category IN ('success','fail','incomplete','unfinished','breakthrough','depth')),
        description TEXT DEFAULT '',
        params_json TEXT NOT NULL,
        logic_op TEXT DEFAULT 'AND' CHECK(logic_op IN ('AND','OR')),
        enabled INTEGER DEFAULT 1,
        is_default INTEGER DEFAULT 0,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
    )""")
    conn.execute("INSERT INTO result_judge_configs SELECT * FROM result_judge_configs_old")
    conn.execute("DROP TABLE result_judge_configs_old")
    conn.commit()

# ===================================================================
#  编码器校准 CRUD
# ===================================================================

def get_encoder_calibration(machine=None, cycle_type=None):
    q = "SELECT * FROM encoder_calibration WHERE 1=1"
    args = []
    if machine:
        q += " AND machine=?"; args.append(machine)
    if cycle_type:
        q += " AND cycle_type=?"; args.append(cycle_type)
    rows = _get_conn().execute(q, args).fetchall()
    return [dict(r) for r in rows]


def upsert_encoder_calibration(machine, cycle_type, position_signal, offset_baseline,
                                travel_range_min=0.0, travel_range_max=3.0,
                                slope_correction=1.0, description=""):
    conn = _get_conn()
    existing = conn.execute(
        "SELECT id FROM encoder_calibration WHERE machine=?", (machine,)
    ).fetchone()
    if existing:
        conn.execute("""UPDATE encoder_calibration SET
            cycle_type=?, position_signal=?, offset_baseline=?,
            travel_range_min=?, travel_range_max=?, slope_correction=?,
            description=?, last_calibrated_at=datetime('now'),
            updated_at=datetime('now')
            WHERE machine=?""",
            (cycle_type, position_signal, offset_baseline, travel_range_min,
             travel_range_max, slope_correction, description, machine))
    else:
        conn.execute("""INSERT INTO encoder_calibration
            (machine, cycle_type, position_signal, offset_baseline,
             travel_range_min, travel_range_max, slope_correction, description)
            VALUES(?,?,?,?,?,?,?,?)""",
            (machine, cycle_type, position_signal, offset_baseline,
             travel_range_min, travel_range_max, slope_correction, description))
    conn.commit()


# ========== 变量采集配置 ==========

def get_variable_configs(cycle_type=None, equipment=None, dimension=None):
    """获取变量配置列表，支持按类型/设备/维度过滤"""
    q = "SELECT * FROM variable_configs WHERE is_active=1"
    args = []
    if cycle_type and cycle_type != "all":
        q += " AND (cycle_type=? OR cycle_type='common')"
        args.append(cycle_type)
    if equipment:
        q += " AND (equipment=? OR equipment='')"
        args.append(equipment)
    if dimension:
        q += " AND dimension=?"
        args.append(dimension)
    q += " ORDER BY cycle_type, dimension, id"
    rows = _get_conn().execute(q, args).fetchall()
    return [dict(r) for r in rows]


def get_variable_config(config_id):
    r = _get_conn().execute("SELECT * FROM variable_configs WHERE id=?", (config_id,)).fetchone()
    return dict(r) if r else None


def get_variable_config_by_tag(tag_name):
    r = _get_conn().execute("SELECT * FROM variable_configs WHERE tag_name=?", (tag_name,)).fetchone()
    return dict(r) if r else None


def upsert_variable_config(config_id, tag_name, chinese_name, **kwargs):
    """创建或更新变量配置"""
    conn = _get_conn()
    if config_id:
        updates = ["tag_name=?", "chinese_name=?", "updated_at=datetime('now')"]
        args = [tag_name, chinese_name]
        for k in ("unit", "cycle_type", "equipment", "dimension", "data_type", "description", "is_active"):
            if k in kwargs:
                updates.append(f"{k}=?")
                args.append(kwargs[k])
        args.append(config_id)
        conn.execute(f"UPDATE variable_configs SET {','.join(updates)} WHERE id=?", args)
        conn.commit()
        return config_id
    else:
        c = conn.execute(
            """INSERT INTO variable_configs(tag_name,chinese_name,unit,cycle_type,equipment,dimension,data_type,description,is_active)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (tag_name, chinese_name,
             kwargs.get("unit", ""), kwargs.get("cycle_type", "common"),
             kwargs.get("equipment", ""), kwargs.get("dimension", ""),
             kwargs.get("data_type", "float"), kwargs.get("description", ""),
             kwargs.get("is_active", 1)))
        conn.commit()
        return c.lastrowid


def delete_variable_config(config_id):
    _get_conn().execute("DELETE FROM variable_configs WHERE id=?", (config_id,))
    _get_conn().commit()


def get_variable_config_options(cycle_type=None):
    """获取下拉选择器选项: [{tag_name, chinese_name, unit}]"""
    q = "SELECT tag_name, chinese_name, unit, dimension FROM variable_configs WHERE is_active=1"
    args = []
    if cycle_type and cycle_type != "all":
        q += " AND (cycle_type=? OR cycle_type='common')"
        args.append(cycle_type)
    q += " ORDER BY dimension, chinese_name"
    rows = _get_conn().execute(q, args).fetchall()
    return [{"tag_name": r[0], "chinese_name": r[1], "unit": r[2], "dimension": r[3]} for r in rows]


_VARIABLE_SEED = [
    # === 东开口机 (opening / east_opener) ===
    ("LT_LQFC_57", "东开口机选择", "", "opening", "east_opener", "remote_command", "bool"),
    ("LT_LQFC_59", "回转进/退阀", "", "opening", "east_opener", "control_valve", "bool"),
    ("LT_LQFC_60", "转钎正传/反转", "", "opening", "east_opener", "remote_command", "bool"),
    ("LT_LQFC_61", "小车前进/后退阀", "", "opening", "east_opener", "control_valve", "bool"),
    ("LT_LQFC_62", "挂钩/脱钩", "", "opening", "east_opener", "remote_command", "bool"),
    ("LT_LQFC_63", "回转位置", "deg", "opening", "east_opener", "position", "float"),
    ("LT_LQFC_64", "回转压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_65", "挂钩位置", "mm", "opening", "east_opener", "position", "float"),
    ("LT_LQFC_66", "倾动压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_67", "推进位置", "mm", "opening", "east_opener", "position", "float"),
    ("LT_LQFC_68", "推进压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_69", "冲击开/关", "", "opening", "east_opener", "remote_command", "bool"),
    ("LT_LQFC_74", "回转回油压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_75", "倾动回油压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_85", "送给回油压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_86", "转钎回油压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_87", "转钎压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_88", "冲击压力", "MPa", "opening", "east_opener", "pressure", "float"),
    ("LT_LQFC_89", "冲击回油压力", "MPa", "opening", "east_opener", "pressure", "float"),

    # === 西开口机 (opening / west_opener) ===
    ("LT_LQFC_94", "西开口机选择", "", "opening", "west_opener", "remote_command", "bool"),
    ("LT_LQFC_96", "回转进/退阀", "", "opening", "west_opener", "control_valve", "bool"),
    ("LT_LQFC_98", "小车前进/后退阀", "", "opening", "west_opener", "control_valve", "bool"),
    ("LT_LQFC_100", "回转位置", "deg", "opening", "west_opener", "position", "float"),
    ("LT_LQFC_101", "回转压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_103", "倾动压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_104", "推进位置", "mm", "opening", "west_opener", "position", "float"),
    ("LT_LQFC_105", "推进压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_106", "冲击开/关", "", "opening", "west_opener", "remote_command", "bool"),
    ("LT_LQFC_111", "回转回油压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_112", "倾动回油压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_122", "送给回油压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_124", "转钎压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_125", "冲击压力", "MPa", "opening", "west_opener", "pressure", "float"),
    ("LT_LQFC_126", "冲击回油压力", "MPa", "opening", "west_opener", "pressure", "float"),

    # === 东堵口机 (plugging / east_plugger) ===
    ("LT_LQFC_129", "遥控电源", "", "plugging", "east_plugger", "remote_command", "bool"),
    ("LT_LQFC_130", "遥控启动/停止", "", "plugging", "east_plugger", "remote_command", "bool"),
    ("LT_LQFC_133", "回转进/退阀", "", "plugging", "east_plugger", "control_valve", "bool"),
    ("LT_LQFC_134", "打泥前进/后退", "", "plugging", "east_plugger", "control_valve", "bool"),
    ("LT_LQFC_135", "回转位置", "deg", "plugging", "east_plugger", "position", "float"),
    ("LT_LQFC_136", "转炮压力", "MPa", "plugging", "east_plugger", "pressure", "float"),
    ("LT_LQFC_137", "打泥位置", "mm", "plugging", "east_plugger", "position", "float"),
    ("LT_LQFC_138", "打泥压力", "MPa", "plugging", "east_plugger", "pressure", "float"),
    ("LT_LQFC_139", "退泥压力", "MPa", "plugging", "east_plugger", "pressure", "float"),
    ("LT_LQFC_140", "退炮压力", "MPa", "plugging", "east_plugger", "pressure", "float"),
    ("LT_LQFC_141", "SET/备用", "", "plugging", "east_plugger", "safety", "bool"),
    ("LT_LQFC_142", "急停", "", "plugging", "east_plugger", "safety", "bool"),
    ("LT_LQFC_143", "等待位置", "deg", "plugging", "east_plugger", "position", "float"),
    ("LT_LQFC_144", "工作位置", "deg", "plugging", "east_plugger", "position", "float"),
    ("LT_LQFC_179", "打泥量", "L", "plugging", "east_plugger", "mud", "float"),

    # === 西堵口机 (plugging / west_plugger) ===
    ("LT_LQFC_152", "遥控电源", "", "plugging", "west_plugger", "remote_command", "bool"),
    ("LT_LQFC_153", "遥控启动/停止", "", "plugging", "west_plugger", "remote_command", "bool"),
    ("LT_LQFC_156", "回转进/退阀", "", "plugging", "west_plugger", "control_valve", "bool"),
    ("LT_LQFC_157", "打泥前进/后退", "", "plugging", "west_plugger", "control_valve", "bool"),
    ("LT_LQFC_158", "回转位置", "deg", "plugging", "west_plugger", "position", "float"),
    ("LT_LQFC_159", "转炮压力", "MPa", "plugging", "west_plugger", "pressure", "float"),
    ("LT_LQFC_160", "打泥位置", "mm", "plugging", "west_plugger", "position", "float"),
    ("LT_LQFC_161", "打泥压力", "MPa", "plugging", "west_plugger", "pressure", "float"),
    ("LT_LQFC_162", "退泥压力", "MPa", "plugging", "west_plugger", "pressure", "float"),
    ("LT_LQFC_163", "退炮压力", "MPa", "plugging", "west_plugger", "pressure", "float"),
    ("LT_LQFC_164", "SET/备用", "", "plugging", "west_plugger", "safety", "bool"),
    ("LT_LQFC_165", "急停", "", "plugging", "west_plugger", "safety", "bool"),
    ("LT_LQFC_166", "等待位置", "deg", "plugging", "west_plugger", "position", "float"),
    ("LT_LQFC_167", "工作位置", "deg", "plugging", "west_plugger", "position", "float"),
    ("LT_LQFC_180", "打泥量", "L", "plugging", "west_plugger", "mud", "float"),

    # === 公共信号 ===
    ("LT_LQFC_150", "东液压站温度", "°C", "common", "", "hydraulic", "float"),
    ("LT_LQFC_151", "东液压站压力", "MPa", "common", "", "hydraulic", "float"),
    ("LT_LQFC_173", "西液压站温度", "°C", "common", "", "hydraulic", "float"),
    ("LT_LQFC_174", "西液压站压力", "MPa", "common", "", "hydraulic", "float"),
]


def seed_variable_configs():
    """初始化变量采集配置（已存在则跳过）"""
    conn = _get_conn()
    existing = conn.execute("SELECT COUNT(*) FROM variable_configs").fetchone()[0]
    if existing > 0:
        logger.info("variable_configs 已存在 %d 条记录，跳过种子", existing)
        return
    for tag, cname, unit, ctype, equip, dim, dtype in _VARIABLE_SEED:
        try:
            conn.execute(
                """INSERT INTO variable_configs(tag_name,chinese_name,unit,cycle_type,equipment,dimension,data_type)
                   VALUES(?,?,?,?,?,?,?)""",
                (tag, cname, unit, ctype, equip, dim, dtype))
        except sqlite3.IntegrityError:
            pass  # UNIQUE constraint
    conn.commit()
    logger.info("variable_configs 种子完成，共 %d 条", len(_VARIABLE_SEED))


init_db()
_migrate_result_judge_configs()
seed_default_result_configs()
seed_variable_configs()
