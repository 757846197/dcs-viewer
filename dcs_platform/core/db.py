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
        {"signal": "LT_LQFC_59", "role": "threshold",  "label": "倾动压力",          "threshold": 10.0, "operator": "gt"},
        {"signal": "LT_LQFC_69", "role": "threshold",  "label": "回转压力",          "threshold": 10.0, "operator": "gt"},
        {"signal": "LT_LQFC_68", "role": "threshold",  "label": "推进压力",          "threshold": 10.0, "operator": "gt"},
        {"signal": "LT_LQFC_67", "role": "threshold",  "label": "推进比例阀给定",     "threshold": 12.0, "operator": "gt"}
    ],
    "filter_min_s": 30,
    "filter_max_s": 3600
}

_DEFAULT_PLUGGING_CONFIG = {
    "type": "plugging",
    "rules": [
        {"signal": "LT_LQFC_130", "role": "remote",    "label": "遥控启动"},
        {"signal": "LT_LQFC_135", "role": "threshold", "label": "回转压力",          "threshold": 10.0, "operator": "gt"},
        {"signal": "LT_LQFC_138", "role": "threshold", "label": "打泥压力",          "threshold": 10.0, "operator": "gt"},
        {"signal": "LT_LQFC_134", "role": "threshold", "label": "打泥比例阀给定",     "threshold": 12.0, "operator": "gt"}
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
    "opening": {
        "success": {"name": "开口成功判定", "logic": "AND", "params": [
            {"param_name": "push_pos_change", "value": 0.1, "unit": "m", "data_type": "float",
             "range_min": 0.01, "range_max": 0.5, "label": "推进位移骤增阈值", "operator": "gt"},
            {"param_name": "push_press_drop_ratio", "value": 0.2, "unit": "%", "data_type": "float",
             "range_min": 0.05, "range_max": 0.5, "label": "压力骤降阈值", "operator": "gt"},
        ]},
        "fail": {"name": "开口失败判定", "logic": "AND", "params": [
            {"param_name": "no_effective_drill", "value": 1, "unit": "", "data_type": "bool",
             "range_min": 0, "range_max": 1, "label": "无有效钻进位移", "operator": "eq"},
        ]},
        "incomplete": {"name": "开口未完成判定", "logic": "AND", "params": [
            {"param_name": "has_drill_no_breakthrough", "value": 1, "unit": "", "data_type": "bool",
             "range_min": 0, "range_max": 1, "label": "有钻进但未钻透", "operator": "eq"},
        ]},
        "breakthrough": {"name": "钻透判定", "logic": "AND", "params": [
            {"param_name": "drill_pos_reach", "value": 1.5, "unit": "m", "data_type": "float",
             "range_min": 0.5, "range_max": 3.0, "label": "钻头到位深度", "operator": "gte"},
            {"param_name": "impact_press_drop", "value": 5.0, "unit": "MPa", "data_type": "float",
             "range_min": 2.0, "range_max": 20.0, "label": "冲击压力骤降", "operator": "lt"},
        ]},
        "depth": {"name": "铁口深度计算", "logic": "AND", "params": [
            {"param_name": "opening_duration_min", "value": 60, "unit": "s", "data_type": "int",
             "range_min": 30, "range_max": 300, "label": "最小开口时长", "operator": "gte"},
            {"param_name": "push_total_distance", "value": 1.0, "unit": "m", "data_type": "float",
             "range_min": 0.5, "range_max": 2.5, "label": "推进总行程", "operator": "gte"},
            {"param_name": "depth_ratio", "value": 0.8, "unit": "", "data_type": "float",
             "range_min": 0.5, "range_max": 1.0, "label": "深度达标比例", "operator": "gte"},
        ]},
    },
    "plugging": {
        "success": {"name": "堵口成功判定", "logic": "AND", "params": [
            {"param_name": "mud_qty_min", "value": 100, "unit": "kg", "data_type": "float",
             "range_min": 50, "range_max": 500, "label": "打泥量达标值", "operator": "gte"},
            {"param_name": "hold_duration_min", "value": 60, "unit": "s", "data_type": "int",
             "range_min": 30, "range_max": 300, "label": "保压时间下限", "operator": "gte"},
        ]},
        "fail": {"name": "堵口失败判定", "logic": "AND", "params": [
            {"param_name": "mud_qty_below_min", "value": 1, "unit": "", "data_type": "bool",
             "range_min": 0, "range_max": 1, "label": "打泥量未达标", "operator": "eq"},
        ]},
        "unfinished": {"name": "堵口未完整判定", "logic": "AND", "params": [
            {"param_name": "mud_done_hold_short", "value": 1, "unit": "", "data_type": "bool",
             "range_min": 0, "range_max": 1, "label": "打泥完成但保压不足", "operator": "eq"},
        ]},
        "breakthrough": {"name": "钻透判定", "logic": "AND", "params": [
            {"param_name": "mud_press_peak", "value": 15.0, "unit": "MPa", "data_type": "float",
             "range_min": 5.0, "range_max": 30.0, "label": "打泥压力峰值", "operator": "gte"},
            {"param_name": "mud_qty_total", "value": 80, "unit": "kg", "data_type": "float",
             "range_min": 30, "range_max": 300, "label": "堵口打泥总量", "operator": "gte"},
        ]},
        "depth": {"name": "铁口深度计算", "logic": "AND", "params": [
            {"param_name": "plugging_duration_min", "value": 45, "unit": "s", "data_type": "int",
             "range_min": 20, "range_max": 180, "label": "最小堵口时长", "operator": "gte"},
            {"param_name": "mud_flow_total", "value": 60, "unit": "kg", "data_type": "float",
             "range_min": 20, "range_max": 200, "label": "总打泥流量", "operator": "gte"},
            {"param_name": "depth_ratio", "value": 0.75, "unit": "", "data_type": "float",
             "range_min": 0.5, "range_max": 1.0, "label": "深度达标比例", "operator": "gte"},
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

init_db()
_migrate_result_judge_configs()
seed_default_result_configs()
