# 🛡️ 中控检测数据实时监控报警器 Watchdog V3.1

**寒锐钴业 — 过程质量 SPC 异常监控系统**

自动扫描中控检测 Excel 日报/月报，提取 QC 规格标准，实时检测数据超限并推送飞书报警通知。

---

## 📋 功能特性

| 特性 | 说明 |
|------|------|
| **规格自动提取** | 从 Excel 标准行自动解析 ≤10、≥80、80~110、(参考) 等规格格式 |
| **白名单过滤** | 只监控指定样品（正则匹配），跳过无关数据 |
| **报警去重** | 同一样品同一检测值不重复通知 |
| **规格覆盖** | 不修改 Excel，通过配置单独调整报警阈值 |
| **Monitor 模式** | 带 `(参考)` 标记的规格仅记录不报警 |
| **多文件支持** | 同时监控多个 Excel 文件 |
| **并行扫描** | 多工作表并行处理，提升扫描速度 |
| **飞书通知** | 通过飞书机器人 Webhook 推送卡片消息，按样品分组展示 |
| **数据库存储** | SQLite 数据库（可插拔架构，预留 MySQL/PG/飞书多维表格） |
| **日志滚动** | 日志文件自动滚动（≤50MB/个，保留 3 个备份） |
| **自动重连** | 数据库异常自动重连 |
| **优雅退出** | Ctrl+C 信号处理后完成当前周期再退出 |
| **锁文件** | 防止多实例同时运行 |
| **启动校验** | 启动时自动检查配置完整性、文件存在性、白名单正则合法性 |

---

## 🏗️ 系统架构

```
config.json ───────┬── Excel文件路径/密码
                   ├── 白名单正则模式
                   ├── 飞书 Webhook URL
                   ├── 阈值覆盖
                   └── 扫描间隔/重试/日志配置
                        │
┌─────────────────────┘
▼
watchdog.py V3.1 ───┬── ① 读取配置 + 启动校验
                    ├── ② 预编译白名单正则
                    ├── ③ 逐文件打开 Excel（支持密码）
                    │       └── 并行扫描各工作表
                    │            ├── 提取标准行规格 → 写入 qc_specs 表
                    │            ├── 白名单过滤
                    │            ├── 行指纹 MD5 去重
                    │            └── 逐行检查超限 → 写入 alarm_log 表
                    ├── ④ 报警去重（同值不重复通知）
                    ├── ⑤ 飞书卡片消息推送（分段防截断）
                    └── ⑥ 写入 CSV 备份
```

### 数据库架构（可插拔）

```
DatabaseBackend（抽象接口）
    ├── SQLiteBackend  ✅ 当前实现（零配置，即开即用）
    ├── MySQLBackend     ── 预留
    ├── PostgreSQLBackend ── 预留
    └── FeishuBitableBackend ── 预留
```

---

## 🚀 快速部署

### 环境要求

- **Python 3.8+**
- **Excel 文件**：中控检测日报/月报（支持密码加密）
- **网络**：可访问飞书 Webhook URL

### Windows 部署

**方式一：一键部署（推荐）**

```batch
1. 将 watchdog_v2 文件夹复制到目标机器
2. 编辑 config.json 配置 Excel 路径和 Webhook
3. 双击 deploy.bat
```

自动完成：检查 Python → 安装依赖 → 初始化数据库 → 测试扫描

**方式二：手动部署**

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 编辑配置
#    打开 config.json，修改 Excel 路径和飞书 Webhook

# 3. 初始化数据库
python watchdog.py --init-db

# 4. 测试扫描
python watchdog.py --once

# 5. 持续运行（前台）
python watchdog.py

# 6. 或 Windows 快捷方式
双击 run_watchdog.bat
```

### Linux 部署（服务器）

```bash
# 安装依赖
pip install -r requirements.txt

# 编辑配置
vi config.json

# 初始化数据库
python watchdog.py --init-db

# 测试
python watchdog.py --once

# 后台运行
python watchdog.py --daemon

# 或使用 systemd 服务
sudo cp watchdog.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable watchdog
sudo systemctl start watchdog
```

---

## ⚙️ 配置说明

### config.json 结构

| 配置段 | 字段 | 说明 |
|--------|------|------|
| `excel.files` | `path` | Excel 文件路径（支持 Z: 网络盘） |
| | `password` | 文件打开密码（无密码不用填） |
| | `exclude_sheets` | 跳过的 Sheet 名称列表 |
| `filter` | `whitelist_enabled` | 是否启用白名单 |
| | `whitelist_patterns` | 正则表达式列表，如 `"一浸液[GHIJKL]槽"` |
| `feishu` | `webhook_url` | 飞书机器人 Webhook 地址 |
| `spec_overrides` | `"样品/项目"` | 阈值覆盖，如 `"精滤液/Al":{"alarm_upper": 0.6}` |
| `database` | `type` | 数据库类型（当前仅 `sqlite`） |
| `watch` | `interval_seconds` | 扫描间隔（秒） |
| `log` | `level` | 日志级别（DEBUG/INFO/WARNING/ERROR） |

### 白名单示例

```json
{
  "whitelist_enabled": true,
  "whitelist_patterns": [
    "一浸液[GHIJKL]槽",
    "除铁铝液[IJK]槽",
    "二浸液[EFGH]槽",
    "二浸渣[EFGH]槽",
    "^精滤液"
  ]
}
```

### 规格覆盖示例（不修改 Excel）

```json
{
  "spec_overrides": {
    "精滤液/Al": {
      "alarm_upper": 0.6
    },
    "精滤液C槽/Fe2+": {
      "alarm_upper": 0.1
    }
  }
}
```

---

## 📖 使用方式

```bash
# 持续监控（前台循环）
python watchdog.py

# 只跑一次（测试用）
python watchdog.py --once

# 初始化/查看数据库
python watchdog.py --init-db

# 后台运行（Linux/Unix）
python watchdog.py --daemon

# 跳过启动校验（调试用）
python watchdog.py --once --skip-validation
```

---

## 🗄️ 数据库表结构

### samples — 检测数据

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| row_hash | TEXT UNIQUE | 行指纹 MD5（去重用） |
| sample_name | TEXT | 样品名称 |
| test_date | TEXT | 检测日期 |
| workshop | TEXT | 车间/工作表名 |
| source_file | TEXT | 来源 Excel 文件名 |
| data_json | TEXT | 全部检测结果 JSON |
| synced_at | DATETIME | 同步时间 |

### alarm_log — 报警记录

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| sample_name | TEXT | 样品名称 |
| item | TEXT | 超限检测项目 |
| value | REAL | 检测值 |
| spec | TEXT | 规格描述（如 ≤10） |
| deviation | TEXT | 偏差百分比 |
| mode | TEXT | alarm（报警）/ monitor（仅记录） |
| level | TEXT | 三级/二级/一级 |
| status | TEXT | 待处置/已处置/已闭环 |
| alarmed_at | DATETIME | 报警时间 |

### qc_specs — QC 规格

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER PK | 自增主键 |
| sample_pattern | TEXT | 样品名称正则 |
| item | TEXT | 检测项目 |
| upper_limit | REAL | 规格上限 |
| lower_limit | REAL | 规格下限 |
| mode | TEXT | alarm/monitor |
| source | TEXT | auto（自动提取）/ manual（手动录入） |

---

## 🔧 常见问题

**Q: Excel 文件被占用无法读取？**
→ 自动重试机制：指数退避最多 3 次，每次间隔 5~60 秒。

**Q: 飞书没收到通知？**
→ 检查 config.json 中的 `feishu.webhook_url` 是否正确
→ 检查是否所有报警都是 `monitor` 模式（带 `(参考)` 标记）
→ 检查控制台日志是否有 `✅ 飞书通知已发送`

**Q: 只想监控某些样品？**
→ 编辑 `filter.whitelist_patterns`，正则不匹配的样品自动跳过。

**Q: Excel 标准行写的规格太严格，想调宽？**
→ 使用 `spec_overrides` 覆盖，不修改 Excel 原文件。

**Q: 如何查看运行状态？**
→ 在 `watchdog.log` 查看详细运行日志
→ `python watchdog.py --init-db` 查看数据库统计

---

## 📦 文件清单

| 文件 | 说明 |
|------|------|
| `watchdog.py` | 主程序（V3.1 增强版） |
| `config.json` | ⚠️ 配置文件（含隐私信息，需现场配置） |
| `config.template.json` | 配置模板（占位符版本，安全提交用） |
| `requirements.txt` | Python 依赖清单 |
| `deploy.bat` | Windows 首次部署脚本 |
| `run_watchdog.bat` | Windows 持续运行脚本 |
| `watchdog.service` | Linux systemd 服务定义文件 |
| `设计文档.md` | 详细设计文档 |
| `README.md` | 本文件 |
| `检测数据库.sqlite` | SQLite 数据库（自动生成） |
| `watchdog.log` | 运行日志（自动滚动） |
| `alarm_history.csv` | 报警记录备份（自动生成） |

---

## ⚠️ 安全提示

- `config.json` 包含 Webhook URL 和 Excel 密码，**请勿提交到 Git 仓库**
- `.gitignore` 已配置忽略 `config.json`
- 使用 `config.template.json` 作为安全模板提交版本库

---

*寒锐钴业 — 过程质量部 · 三级风险管控体系*
