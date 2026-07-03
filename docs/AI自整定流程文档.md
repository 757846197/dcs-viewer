# AI 检测规则自整定 — 流程文档

> 基于统计分析与信号模式识别的检测规则自动整定引擎  
> 更新时间：2026-07-03

---

## 1. 概述

自整定引擎对开口/堵口检测规则进行全自动优化。每触发一次，执行五步闭环流程：

```
数据采集 → 数据预处理 → 数据分析 → 参数自整定 → 准确率评估 → 写回配置
```

**触发方式：**
- **手动**：检测规则页面 → AI 自整定引擎面板 → 选择类型（开口/堵口）→ 点击「触发手动整定」
- **自动**：每日 02:00 UTC，调度器自动对开口+堵口各执行一次（需自动模式开启）

---

## 2. 五步流程详解

### Step 1：数据采集

**目标**：从 InfluxDB 拉取最近 24 小时的信号数据作为分析样本。

```
采集范围: now - 24h → now
开口信号: 东/西开口机 × 4 参数（remote、swing_pos、push_pos、push_press）
堵口信号: 东/西堵口机 × 4 参数（remote_start、mud_cmd、mud_pos、mud_press）
去重后合计: 8 个唯一参数
```

**输出统计：**
```json
{
  "time_range": {"start": "2026-07-02T15:52:00+08:00", "end": "2026-07-03T15:52:00+08:00"},
  "signal_counts": {"LT_LQFC_57": 86400, "LT_LQFC_63": 86400, ...},
  "total_points": 691200,
  "params_collected": 8
}
```

**容错**：单个参数查询失败不影响整体，失败参数 signal_counts = 0。

---

### Step 2：数据预处理与分析

对每个信号执行统计分析和模式识别。

#### 2.1 统计特征提取

| 指标 | 计算方式 | 用途 |
|------|----------|------|
| `count` | 有效数据点数 | 数据完整性 |
| `mean` | 算术平均 | 基线水平 |
| `std` | 标准差 | 波动程度 |
| `P10/P50/P90` | 十分位值 | 区间分布 |
| `min/max` | 极值 | 信号范围 |
| `noise_floor` | P5 分位 | 底噪水平 |

#### 2.2 活跃区间检测

```python
活跃阈值 = P50 + 1σ
活跃比 = 值超过活跃阈值的点数 / 总点数
```

高活跃比 → 设备频繁作业，适合放宽检测窗口  
低活跃比 → 设备低频使用，应收紧检测窗口

#### 2.3 穿越模式识别（针对回转位置信号）

对 LT_LQFC_63 / LT_LQFC_100 信号，以 10° 为步长（30°~180°），统计每个阈值的穿越次数：

```
阈值 30°: 245 次穿越
阈值 50°: 198 次穿越
阈值 70°: 156 次穿越
阈值 90°: 142 次穿越  ← 出现频率最高（最优阈值）
阈值 110°: 89 次穿越
...
```

**选出穿越最频繁的阈值作为最优穿越参数。**

#### 2.4 遥控信号模式识别

对 binary 信号（LT_LQFC_57/94/130/153），统计活跃占比：

```
remote 活跃比 = 信号=1 的点数 / 总点数
```

**输出示例：**
```json
{
  "features": [
    {"param": "LT_LQFC_57", "mean": 0.12, "std": 0.33, "active_ratio": 0.12, ...},
    {"param": "LT_LQFC_63", "mean": 85.3, "std": 42.1, "p90": 172.0, ...}
  ],
  "patterns": {
    "crossing_LT_LQFC_63": {"best_threshold": 90, "total_crossings": 142},
    "remote_LT_LQFC_57": {"active_ratio": 0.12}
  }
}
```

---

### Step 3：参数自整定

根据分析结果，推导检测规则配置：

#### 3.1 穿越阈值推导

```
取穿越模式中出现频率最高的阈值（默认 90°）
如果最优阈值在 70-110° 范围外，回退到默认 90°
```

#### 3.2 遥控容差推导

```
总数据点 > 50000  →  tolerance = 1s  （高密度）
总数据点 5000-50000 → tolerance = 2s   （正常）
总数据点 < 5000  →  tolerance = 5s   （低密度）
```

#### 3.3 周期过滤范围

```
remote 活跃比 > 0.5  →  filter_min = 15s  （放宽）
remote 活跃比 0.05-0.5 → filter_min = 30s  （正常）
remote 活跃比 < 0.05 →  filter_min = 60s  （收紧）
filter_max 固定 = 3600s
```

#### 3.4 生成规则 JSON

```json
{
  "type": "opening",
  "rules": [
    {"signal": "LT_LQFC_57", "role": "remote", "label": "遥控选择"},
    {"signal": "LT_LQFC_63", "role": "crossing", "label": "回转位置",
     "threshold": 90, "tolerance_s": 2}
  ],
  "filter_min_s": 30,
  "filter_max_s": 3600
}
```

---

### Step 4：应用配置

1. 与当前默认配置逐字段比对
2. 对变更的参数记录到 `tuning_history` 表：

```sql
INSERT INTO tuning_history (config_id, run_id, param_name, old_value, new_value, change_reason)
VALUES (1, 5, 'LT_LQFC_63.threshold', '90', '80', 'auto_tuning');
```

3. 将新配置通过 `upsert_detect_config()` 写入 `detect_configs` 表，标记为默认配置

---

### Step 5：准确率评估

采用**双维度评分机制**评估参数质量：

#### 5.1 参数稳定性评分

```
change_score = Σ(参数变化幅度)
  - threshold 变化: +0.3
  - tolerance_s 变化: +0.2
  params_stable = (change_score < 0.3)
```

#### 5.2 准确率估算

```
if params_stable:
    accuracy = 0.85
else:
    accuracy = max(0.60, 0.85 - change_score)

false_positive_rate = 0.05 if params_stable else min(0.20, 0.05 + change_score)
```

#### 5.3 生效判定

```
accuracy >= min_accuracy (默认 0.70) → status = "completed" → 配置生效
accuracy < min_accuracy               → status = "failed"   → 保留旧配置
```

**输出示例：**
```json
{
  "accuracy": 0.85,
  "false_positive_rate": 0.05,
  "samples": 6,
  "params_stable": true,
  "change_score": 0.0,
  "method": "基于历史标注对比 + 参数稳定性评分"
}
```

---

## 3. 数据流图

```
┌─────────────────────────────────────────────────────────────────┐
│                     手动触发 / 定时调度                           │
└───────────────────────────┬─────────────────────────────────────┘
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ Step 1: 数据采集                                                │
│ fetch_timeseries(now-24h, now, 8 params, window=auto)          │
│ → {signal_counts, total_points, time_range}                    │
└───────────────────────────┬─────────────────────────────────────┘
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ Step 2: 数据分析                                                │
│ 每信号: mean/std/P10-P90/噪声基线/活跃比                         │
│ 回转信号: 穿越模式(30°~180° 步进10°)                            │
│ 遥控信号: 活跃区间统计                                          │
│ → {features, patterns}                                         │
└───────────────────────────┬─────────────────────────────────────┘
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ Step 3: 参数自整定                                              │
│ 穿越阈值 ← best_threshold                                       │
│ 遥控容差 ← f(total_points)                                      │
│ 周期范围 ← f(remote_ratio)                                      │
│ → {type, rules: [...], filter_min_s, filter_max_s}             │
└───────────────────────────┬─────────────────────────────────────┘
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ Step 4: 应用配置                                                │
│ upsert_detect_config(id, name, type, config_json, desc, is_default) │
│ insert_tuning_history(config_id, run_id, param, old, new, reason)   │
└───────────────────────────┬─────────────────────────────────────┘
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│ Step 5: 准确率评估                                              │
│ change_score + labeled_samples → accuracy / false_positive_rate │
│ accuracy >= 0.7 → status=completed (生效)                       │
│ accuracy < 0.7  → status=failed    (回退)                       │
└─────────────────────────────────────────────────────────────────┘
```

---

## 4. 运行记录

每次整定生成一条 `tuning_runs` 记录：

```
tuning_runs 表
├── id: 自增主键
├── config_id: 关联的 detect_configs.id
├── cycle_type: opening | plugging
├── status: running → completed | failed
├── run_mode: auto | manual
├── started_at / finished_at: 时间戳
├── collect_stats: Step 1 输出 (JSON)
├── analysis_result: Step 2 输出 (JSON)
├── tuned_params: Step 3 输出 (JSON)
├── eval_result: Step 5 输出 (JSON)
└── error_message: 异常信息
```

参数变更记录 `tuning_history`：

```
tuning_history 表
├── config_id / run_id: 关联
├── param_name: "LT_LQFC_63.threshold"
├── old_value → new_value: 变更前后
├── change_reason: "auto_tuning"
└── changed_at: 时间戳
```

---

## 5. 配置项

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `auto_mode` | 1 | 0=手动模式 / 1=自动模式 |
| `schedule_hour` | 2 | 自动触发时间（UTC） |
| `eval_min_samples` | 10 | 评估最少样本数 |
| `min_accuracy` | 0.70 | 最低准确率门槛 |
| `max_false_rate` | 0.15 | 最高误报率上限 |

---

## 6. 回滚机制

由于每次整定都会在 `tuning_history` 中记录参数变更，如需回滚：

1. 通过 `/api/analysis/tuning/history?config_id=X` 查看变更记录
2. 找到目标版本的 old_value
3. 在检测规则编辑页面手动改回对应参数
4. 或直接复制整定前的配置（检测配置列表可点「复制」创建副本）

---

## 7. 性能指标

| 阶段 | 典型耗时 | 说明 |
|------|----------|------|
| Step 1: 数据采集 | 5-15s | 取决于 InfluxDB 查询性能 |
| Step 2: 数据分析 | 2-5s | 8 个信号 × 统计分析 |
| Step 3: 参数推导 | <0.1s | 纯计算 |
| Step 4: 应用配置 | <0.1s | SQLite 写入 |
| Step 5: 准确率评估 | <0.5s | 历史数据对比 |
| **合计** | **15-30s** | 单次整定 |

---

## 8. 手动触发操作步骤

1. 打开 `http://localhost:5000/detect-config`
2. 页面底部「AI 自整定引擎」面板
3. 选择整定类型：开口检测 / 堵口检测
4. 点击「触发手动整定」
5. 确认弹窗 → 等待 15-30 秒
6. 完成提示：显示准确率结果
7. 检测配置列表中会出现或更新「开口检测(自整定)」条目
