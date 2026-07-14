# 作业分析流程说明书

> **从点击「开始分析」到结果产出的完整业务逻辑**  
> 文件: `analysis.html` + `analysis_api.py /cycles` + `_extract_cycle_metrics`

---

## 1. 整体流程概览

```
用户点击「开始分析」
       │
       ▼
① 前端校验 (时间范围≤24h, 非空)
       │
       ▼
② GET /api/analysis/cycles?start=...&end=...&type=...
       │
       ▼
③ 后端: 时间标准化 (北京时间→UTC, _normalize_time)
       │
       ├──► ④ 开口检测 (remote=1 + 回转位置穿越90°)
       │        ├─ 东开口机: LT_LQFC_57 + LT_LQFC_63
       │        └─ 西开口机: LT_LQFC_94 + LT_LQFC_100
       │
       └──► ⑤ 堵口检测 (遥控启动=1 + 打泥前进边沿)
                ├─ 东堵口机: LT_LQFC_130 + LT_LQFC_134
                └─ 西堵口机: LT_LQFC_153 + LT_LQFC_157
       │
       ▼
⑥ 去重 + 周期合并 (同设备5秒内合并)
       │
       ▼
⑦ 批量指标提取 (_enrich_cycles_with_metrics, ≤1.5天)
       │
       ├──► ⑧ 信号统计计算 (max/min/mean/range/count)
       │
       └──► ⑨ 判定规则匹配 (result_judge_configs, 优先级排序)
       │
       ▼
⑩ 去重入库 (SQLite cycles表, 防重复)
       │
       ▼
⑪ JSON 返回前端
       │
       ▼
⑫ 前端渲染: 统计卡片 + 周期列表 + 详情面板
```

---

## 2. 前端入口 — `runAnalysis()`

### 2.1 参数校验 (analysis.html L540-562)

```
用户选择参数:
  ┌─────────────┬──────────────────────────────────┐
  │ 开始时间     │ 日期选择器, 格式: 2026-07-14T08:00│
  │ 结束时间     │ 日期选择器                        │
  │ 作业类型     │ dropdown: 全部/开口/堵口           │
  │ 设备筛选     │ dropdown: 全部/东开/西开/东堵/西堵  │
  └─────────────┴──────────────────────────────────┘

校验规则:
  - 时间必填
  - 跨度 ≤ 24 小时
  - 结束时间 > 开始时间
  - APP_TOKEN 存在
```

### 2.2 API 调用

```javascript
GET /api/analysis/cycles?start=2026-07-14T08:00&end=2026-07-14T16:00&type=all&token=dcs2026
```

### 2.3 结果处理

```
API 返回 → data.cycles[]
    │
    ├─ 设备筛选: globalCycles.filter(c.machine === selectedMachine)
    │
    ├─ 加载标签映射: loadLabelMap() → 从 variable_configs 拉取 signal→中文名
    │
    ├─ renderStats()    → 顶部4个统计卡片
    ├─ renderCycles()   → 周期表格 (机器/类型/时间/指标/结果/详情按钮)
    └─ 导出Excel: btnExport → /api/analysis/export
```

---

## 3. 后端周期检测 — `api_cycles()`

### 3.1 信号配置

| 设备 | 开口检测 | 堵口检测 |
|------|---------|---------|
| 东开口机 | remote=57, swing=63 | — |
| 西开口机 | remote=94, swing=100 | — |
| 东堵口机 | — | remote=130, mud_cmd=134 |
| 西堵口机 | — | remote=153, mud_cmd=157 |

### 3.2 开口检测算法

```
检测配置: 从 detect_configs 表加载 → has_threshold_rule ?
    ├─ 有阈值规则 → _detect_threshold_cycles() (多信号联合)
    └─ 无阈值规则 → legacy 穿越算法 (当前默认)

Legacy穿越算法:
┌─────────────────────────────────────────────────┐
│ ① InfluxDB查询: remote + swing_pos 两个信号      │
│                                                  │
│ ② remote信号: 构建时间戳映射 {ts: value}          │
│    swing_pos信号: 按时间排序                      │
│                                                  │
│ ③ 遍历swing_pos:                                 │
│    if prev < 90° and curr >= 90°:                │
│      → 回转位置穿越90° (Rising edge)             │
│      if remote==1 nearby (±2s):                  │
│        → cycle_start = curr_time                 │
│                                                  │
│    if prev >= 90° and curr < 90°:                │
│      → 回转离开90° (Falling edge)                │
│      if in_cycle:                                │
│        duration = curr_time - cycle_start        │
│        if 30s ≤ duration ≤ 3600s:                │
│          → 记录为1个完整开口周期                  │
└─────────────────────────────────────────────────┘
```

### 3.3 堵口检测算法

```
Legacy边沿检测:
┌─────────────────────────────────────────────────┐
│ ① InfluxDB查询: remote_start + mud_cmd           │
│                                                  │
│ ② mud_cmd信号: 非二进制三态 (4=退,12=中位,20=进) │
│    边沿阈值: 15 (来自detect_configs mud_cmd rule) │
│                                                  │
│ ③ 遍历mud_cmd:                                   │
│    if prev < 15 and curr >= 15:                  │
│      → 打泥前进上升沿                             │
│      if remote_start==1 nearby (±2s):            │
│        → cycle_start = curr_time                 │
│                                                  │
│    if prev >= 15 and curr < 15:                  │
│      → 打泥后退下降沿                             │
│      if in_cycle and 30s ≤ dur ≤ 3600s:          │
│        → 记录为堵口周期                           │
└─────────────────────────────────────────────────┘
```

### 3.4 去重策略

```
同设备 + 同类型 + window_start 完全一致 → 跳过
同设备 + 5秒内多次触发 → 只记录最长的那次
```

---

## 4. 指标提取 — `_extract_cycle_metrics()`

### 4.1 信号统计计算

对每个相关信号, 在 `[window_start, window_end]` 窗口内计算:

```python
{
    "max":  窗口内最大值,
    "min":  窗口内最小值,
    "mean": 窗口内均值,
    "range": max - min,        # 变化量
    "count": 采样点数,          # 持续时间估算
    "values": [v1, v2, ...],  # 原始值(用于变化率计算)
}
```

### 4.2 基础指标 (硬编码, 用于前端展示)

**开口作业**:
| 指标 | 来源信号 | 含义 |
|------|---------|------|
| push_depth | LT_LQFC_67 range | 推进位移量 (m) |
| push_press_peak | LT_LQFC_68 max | 推进压力峰值 (MPa) |
| drill_press_mean | LT_LQFC_87 mean | 转钎压力均值 (MPa) |
| impact_active | LT_LQFC_69 max ≥ 0.5 | 冲击指令已激活 |
| push_pos_change | LT_LQFC_67 range | 等价于 push_depth |

**堵口作业**:
| 指标 | 来源信号 | 含义 |
|------|---------|------|
| mud_qty | LT_LQFC_179 range | 打泥量 (L) |
| mud_press_peak | LT_LQFC_138 max | 打泥压力峰值 (MPa) |
| mud_press_mean | LT_LQFC_138 mean | 打泥压力均值 (MPa) |
| hold_duration_s | LT_LQFC_134 count | 保压采样点数 |

### 4.3 硬编码兜底判定 (in case no config rules)

```
开口: push_depth>0.1 + press_ratio<0.8 → breakthrough → "success"
       push_depth>0.1 → "incomplete"
       else → "fail"

堵口: mud_qty>0.1 + hold_ok → "success"
       mud_qty>0.1 → "partial"
       else → "fail"
```

### 4.4 配置驱动判定 (result_judge_configs 优先级排序)

```
加载数据库配置 → 按 priority DESC 排序 → 分组评估:

if is_static==1:      params=[] → 无条件匹配
if logic_op==AND:     all(conditions)
if logic_op==OR:      any(conditions)

同category内首条匹配即停止 (priority高的优先)

综合判定 (cross-category冲突):
  verdicts["success"] 优先
  → else verdicts["fail"]
  → else verdicts["incomplete"]/["unfinished"]  
  → else 硬编码兜底
```

---

## 5. 前端渲染链路

### 5.1 统计卡片 — `renderStats()`

```
┌──────────┬──────────┬──────────┬──────────┐
│ 开口次数  │ 堵口次数  │ 成功率    │ 成功率    │
│ 30次      │ 8次      │ 73%      │ 38%      │
│ 22次钻透  │ 3次完成  │ 钻透/总数  │ 完成/总数  │
└──────────┴──────────┴──────────┴──────────┘

统计逻辑:
  - 开口 → 钻透 = result=="success" OR breakthrough==True
  - 堵口 → 完成 = result=="success"
```

### 5.2 周期列表 — `renderCycles()`

```
┌────────┬──────┬──────────────────┬──────────┬────────┬──────────┬──────┬──────┐
│ 设备    │ 类型  │ 触发时间          │ 窗口        │ 持续   │ 关键指标  │ 结果  │ 详情 │
├────────┼──────┼──────────────────┼──────────┼────────┼──────────┼──────┼──────┤
│东开口机 │ 开口  │ 07-14 08:53:13  │08:53~08:55│ 2分0秒  │钻进0.5m   │成功  │ 详情 │
│西堵口机 │ 堵口  │ 07-14 09:12:05  │09:12~09:13│ 1分0秒  │打泥量120L │失败  │ 详情 │
└────────┴──────┴──────────────────┴──────────┴────────┴──────────┴──────┴──────┘

关键指标显示:
  开口: getLabel(LT_LQFC_67) + '{深度}m | ' + getLabel(LT_LQFC_68) + '峰{压力}MPa'
  堵口: getLabel(LT_LQFC_179) + '{打泥量} | ' + getLabel(LT_LQFC_138) + '峰{压力}MPa'

  标签来源: variable_configs 表 (动态, 修改即生效)
```

### 5.3 详情面板 — `showDetail()`

```
点击「详情」→ GET /api/analysis/metrics → 弹窗展示:

┌──────┬──────┬──────┬──────┬──────┐
│开口深度│ 推进压力│ 转钎压力│钻透判定│冲击状态│
│0.523m │ 13.5MPa│ 8.2MPa │已钻透  │已开启  │
└──────┴──────┴──────┴──────┴──────┘

下方面板: Chart.js 曲线图 (从 variable_configs 动态选信号)
  - 双Y轴: 左轴压力(MPa), 右轴位置/角度
  - 标签: getLabel(signal_name) → variable_configs 动态读取
```

---

## 6. 关键数据流总结

```
InfluxDB (DCS raw data)
    │  fetch_timeseries()
    ▼
周期检测 (remote + crossing/edge)
    │  cycles[] = [{machine, type, trigger_time, window_start, window_end}]
    ▼
指标提取 (_extract_cycle_metrics)
    │  + push_depth, mud_qty, hold_ok, breakthrough...
    │  + result (success/fail/incomplete/partial)
    ▼
数据库入库 (cycles表, 去重)
    │
    ▼
JSON → 前端
    │
    ├─ renderStats()      ← globalCycles 聚合统计
    ├─ renderCycles()     ← 逐行渲染, 用 loadLabelMap() 获取中文名
    └─ showDetail()       ← /api/analysis/metrics 逐个评估, 用 variable_configs 选信号
```

---

> **版本**: v1.0 | **作者**: @senior-developer | **日期**: 2026-07-14
