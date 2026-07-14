# Agent Handoff Protocol — DCS 开堵口分析模型

> **定位**: 多 AI Agent (WorkBuddy / Codex) 协作交接文档  
> **更新规则**: 每次 Agent 完成实质性工作后必须更新本文件  
> **解析约定**: `[ ]` 未完成, `[x]` 已完成, `[~]` 进行中, `@agent` 负责人

---

## 1. 项目概要

| 字段 | 值 |
|------|-----|
| 项目名称 | 炉前开/堵口作业智能化辅助模型 (DCS Analysis Platform) |
| Git 仓库 | https://github.com/757846197/dcs-viewer.git |
| 技术栈 | Python 3.13 (Flask + waitress), SQLite, InfluxDB 2.7 |
| 部署方式 | PyInstaller 单文件 EXE (`dist/DCSViewer.exe`) |
| 环境配置 | 需 `.env` 文件 (InfluxDB Token 等), 不在 Git 中 |
| 数据库 | SQLite: `dcs_analysis.db` (已纳入 Git, 含种子数据) |

---

## 2. 当前进度 & 已完成任务

### 已完成 (最近 7 天)

- [x] 变量采集配置表 `variable_configs` — DB + API + 前端 CRUD 完成 (`@senior-developer`)
- [x] 作业分析页面标签配置驱动 (`loadLabelMap` 替代硬编码) (`@senior-developer`)
- [x] 规则表重构: `priority` + `is_static` + `fallback` 分类 (`@senior-developer`)
- [x] 68 条变量中文名按标准映射表统一修正 (`@senior-developer`)
- [x] 开口检测切换为 legacy 穿越算法 (修复 LT_LQFC_67 阈值 100→-0.5m) (`@senior-developer`)
- [x] 堵口检测 mud_cmd 边缘阈值修正 (0.5→15, 三态信号 4/12/20) (`@senior-developer`)
- [x] EXE 打包: 静态文件 `send_from_directory`→`send_file` 修复 404 (`@senior-developer`)
- [x] 所有页面侧边栏添加 `变量配置` 导航入口 (`@senior-developer`)
- [x] 趋势页 uPlot 缓存问题 (`?v=1.6.30` + no-cache 响应头) (`@senior-developer`)
- [x] 时区修正: `_normalize_time` 北京时间→UTC (`@senior-developer`)

### 较早完成

- [x] 周期检测 (开口/堵口) — legacy + threshold 双路径
- [x] 判定规则配置 CRUD (`result_judge_configs`)
- [x] 检测规则配置 (`detect_configs`)
- [x] 实时监控页面
- [x] 趋势分析页面
- [x] 历史查询页面
- [x] EXE 单文件打包 (PyInstaller 6.21)

---

## 3. 待办任务 & 优先级

### P0 — 阻断

- [ ] EXE 需要重新打包 (包含最新 HTML + DB 修改)
- [ ] `.env.example` 需同步到 `dist/` 目录

### P1 — 高优先级

- [ ] self_tuning 调度器: APScheduler 未打包 (EXE 启动日志: "APScheduler 未安装")
- [ ] 规则页面 (`rules_opening/plugging.html`) 添加 priority 和 is_static 编辑功能
- [ ] 检测规则页面支持中文编辑维护
- [ ] 作业分析周期列表添加"刷新"按钮 (避免重复加载)

### P2 — 中优先级

- [ ] 趋势页导出 PNG/CSV 功能
- [ ] 知识库页面内容填充
- [ ] 循环标注页面完善标注流程
- [ ] 分析报告页面生成 PDF

### P3 — 低优先级 / 优化

- [ ] 前端组件统一化 (减少模板重复 CSS)
- [ ] API 响应压缩 (flask-compress 已安装，验证配置)
- [ ] 数据库查询优化 (大时间范围分批查询)
- [ ] 添加单元测试

---

## 4. 技术架构 & 关键决策

### 目录结构

```
dcs_viewer/           ← Flask 应用 + 前端模板
  app.py              ← 主入口 (路由 + 启动)
  templates/          ← 页面模板 (Jinja2)
  static/lib/         ← 前端库 (uPlot, Chart.js, jQuery, DataTables)
  variable_config.html  ← 变量采集配置页
  detect_config.html    ← 检测规则配置页
  rules_opening.html    ← 开口判定规则
  rules_plugging.html   ← 堵口判定规则

dcs_platform/         ← 后端业务逻辑
  core/
    db.py             ← SQLite 操作 + 种子数据
    influx_client.py  ← InfluxDB 查询封装
    config.py         ← 集中配置 (读 .env)
  routes/
    analysis_api.py   ← 作业分析 + 变量配置 + 判定规则 API
    rules_api.py      ← 规则 CRUD + 评估
  services/
    analysis/         ← 分析子模块
    self_tuning.py    ← 自整定调度器
```

### 关键决策记录

| 决策 | 日期 | 理由 |
|------|------|------|
| 数据库纳入 Git | 2026-07-14 | 种子数据与代码同步, EXE 打包需要 |
| 变量配置表 `is_active` 软删除 | 2026-07 | 已有规则的引用不级联删除 |
| 判定规则用 `priority` 排序 | 2026-07-14 | 解决同 category 多规则冲突 |
| `is_static=1` 空参数规则 | 2026-07-14 | 兜底 fallback, 不依赖信号 |
| `send_file` 替代 `send_from_directory` | 2026-07-14 | PyInstaller `safe_join` 拒绝 _MEIPASS 路径 |
| 开口检测用 legacy 穿越算法 | 2026-07-04 | 阈值路径 `>100mm` 与实际数据 `-1.8~2.6m` 不匹配 |

---

## 5. 已知问题 & 风险

| 问题 | 严重度 | 状态 |
|------|--------|------|
| EXE 打包后需重新测试所有页面 | 中 | [~] 最近一次打包 07-14, 含最新更改 |
| GitHub Push 间歇性连接失败 | 低 | commit 本地保存, 网络恢复后补 push |
| InfluxDB Token 仍存在于 `config.py` 环境变量默认值 | 高 | 应改为空白, 仅从 .env 读取 |
| APScheduler 未打包到 EXE | 中 | 自整定功能在 EXE 中不可用 |
| 两个进程抢 5000 端口时前端连到旧进程 | 中 | DEV+EXE 同时运行时需手动 kill |

---

## 6. 环境配置 & 依赖

### 运行时依赖 (EXE 已打包)

```
Flask 3.x, waitress 3.x, influxdb_client, flask_compress,
xlsxwriter, sqlite3 (_sqlite3), numpy, charset_normalizer
```

### 开发依赖

```
pyinstaller 6.21 (仅打包用)
```

### 环境变量 (`.env` 文件, 不在 Git 中)

```ini
INFLUXDB_URL=http://10.56.128.202:8086
INFLUXDB_TOKEN=<secret>
INFLUXDB_ORG=myOrg
INFLUXDB_BUCKET=islag
FLASK_HOST=0.0.0.0
FLASK_PORT=5000
APP_TOKEN=dcs2026
SETTINGS_PASSWORD=admin123
```

### 运行命令

```bash
# DEV 模式
cd dcs_viewer && python app.py

# 打包
pyinstaller DCSViewer.spec --noconfirm

# EXE 运行
dist/DCSViewer.exe
```

---

## 7. Agent 分工 & 职责边界

| Agent | 职责 | 文件范围 |
|-------|------|---------|
| **senior-developer** (WorkBuddy) | 全栈开发: Flask 后端 + 前端模板 + EXE 打包 + DB 设计 | `dcs_platform/`, `dcs_viewer/`, `DCSViewer.spec` |
| **Codex** (待分配) | 代码审查、测试用例、文档维护、前端优化 | 全项目 |

### 交接约定

1. **每次完成工作 → 更新本文件的 Section 2 (已完成) 和 Section 3 (待办)**
2. **提交前 → `git add AGENT_HANDOFF.md` 与功能代码一起 commit**
3. **接手项目 → 先读本文件, 再执行 `git pull`, 然后开始工作**
4. **遇到问题 → 更新 Section 5 (已知问题)**

---

## 8. 上下文依赖 & 注意事项

### 操作习惯

- 彬哥 (项目 owner) 用简洁指令式反馈: "可以""不行""重做"
- 需用真实数据截图确认效果
- 每次改完必须 commit + push

### 数据特点

- InfluxDB 数据为 UTC 时间, 前端传北京时间
- 控制阀信号 (回转阀/打泥阀) 是三态 (4/12/20), 非二进制
- 东西设备信号有 23+n 的偏移映射 (如 57→94, 63→100, 67→104)

### 打包注意事项

- `DCSViewer.spec` 的 `datas` 需包含所有 HTML/JS/CSS/DB
- `dcs_platform` 使用 `('dcs_platform','dcs_platform')` 树型复制
- 每次新增文件/页面必须同步更新 `.spec`

---

> **最后更新**: 2026-07-14 12:47 CST  
> **更新者**: @senior-developer  
> **下一接手 Agent**: @Codex — 请从 Section 3 P0 任务开始
