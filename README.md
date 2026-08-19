# DMHY 追番助手

本地 Web 界面，包装 [`dmhy_scraper.py`](dmhy_scraper.py) 浏览、更新、整理 [DMHY（動漫花園）](https://share.dmhy.org) 的发布资源。支持增量更新、全站搜索、追番跟踪和 Excel 导出，开箱即用，无需注册。

## 功能

- **近期发布**：3 / 7 / 14 / 31 天预设 + 自定义 1–365 天窗口；按番剧名称 / 集数 / 字幕组搜索已加载数据；编码（AV1/X265/X264/VP9）、语言（简体/繁体/日文/无字幕）、字幕组筛选；每页 30 条分页。
- **全站搜索**：一键切换为实时搜索 DMHY 全站资源（不受已加载数据限制）；DMHY 搜索服务故障（HTTP 500）时自动回退到站点 RSS（仅覆盖最近约 500 条，界面会标注）。结果同样支持复制磁链 / 追番 / 已使用标记。
- **复制与状态**：勾选多行一键复制磁链（换行连接），单行复制；复制过的资源标记为「已使用」，状态保存在浏览器 localStorage，刷新不丢。
- **追番**：每行「追番」按钮收藏作品（localStorage）；「追番更新中」面板按 3/7/14/31 日窗口展示有更新的追番作品（动态「N 日内完成」列）；「已追番」列展示全部追番作品；可一键复制窗口内所有磁链。
- **更新**：按分组增量抓取（天数由 `updates.json` 中该分组上次成功时间推算，失败分组不会拖累其他分组），按 topic ID 合并进持久化 CSV；抓取失败自动走该分组的 RSS 兜底，并在界面标注「RSS 兜底」。
- **表格**：按需生成 `DMHY更新表_YYYY-MM-DD.xlsx`（Asia/Shanghai 时区日期），标题行可展开查看全部发布明细，新增发布与父标题行浅绿色高亮。
- **运行记录 / 更新日志**：可折叠的运行日志面板与版本变更弹窗。

## 快速开始

依赖：**Python 3** + `curl`；Excel 导出额外需要 `openpyxl`：

```bash
pip install -r requirements.txt
```

### 方式一：直接运行

```bash
python3 app.py --open        # 默认端口 8765，自动打开浏览器
python3 app.py --port 8766   # 指定端口
```

### 方式二：macOS 常驻服务

`启动DMHY追番助手.command` 会通过 launchd（`com.dmhy.zhufan.plist`）在后台启动服务（端口 8766，`KeepAlive` 保活，重启后自恢复），然后打开浏览器。

```bash
./启动DMHY追番助手.command
```

### 离线演示

```bash
python3 app.py --demo        # 不访问网络，用现有 CSV 演示界面
```

## 数据文件

应用启动时会预加载合并后的 CSV（后台线程，界面立即可用，缓存完成后自动刷新）。数据默认存放在**应用目录的上一级**：

| 文件 | 说明 |
| --- | --- |
| `lolihouse_topics.csv` / `7acg/7acg_topics.csv` | 各分组合并后的发布数据（增量更新的持久化层） |
| `lolihouse_by_series.html` / `7acg/7acg_by_series.html` | 按片名整理的离线 HTML（由脚本生成） |
| `updates.json` | 各分组最后成功更新时间等运行状态 |
| `DMHY更新表_YYYY-MM-DD.xlsx` | 按需生成的 Excel 工作簿 |

## 支持的分组

| 分组 | 来源 | 说明 |
| --- | --- | --- |
| LoliHouse | `team_id/657` | 官方发布页 + 过滤后的 RSS |
| 7³ACG | `user_id/759200` | 发布者页需 `--title-contains '[7³ACG]'` 过滤（账号含无关发布） |

## 网络与故障处理

- 本地 DNS 被污染导致直连挂起时，脚本优先走已核实的 `DMHY_RESOLVE_IP=104.25.61.106` 路由（保留 TLS 校验）。
- 连接失败（6/7/28）重试一次直连；HTTP 500（56）属 DMHY 服务端故障，界面明确提示并自动启用 RSS 兜底，而不是无意义地切换路由。

## 单独使用 scraper CLI

界面之外，`dmhy_scraper.py` 也可独立使用（CSV/JSON/HTML 仅需 Python 3 + curl）：

```bash
# 抓取 LoliHouse 全部发布
python3 dmhy_scraper.py "https://share.dmhy.org/topics/list/team_id/657" \
  --output-dir "/absolute/output/path"

# 增量刷新 7³ACG 并合并进工作簿
python3 dmhy_scraper.py "https://share.dmhy.org/topics/list/user_id/759200" \
  --title-contains "[7³ACG]" --group-name "7³ACG" \
  --since-days 31 --merge-xlsx \
  --output-dir "/absolute/output/7acg" \
  --xlsx "/absolute/output/DMHY更新表_YYYY-MM-DD.xlsx" --sheet-name "7³ACG"

# 从既有 JSON 重新整理（无需联网）
python3 dmhy_scraper.py --from-json "/absolute/output/team_topics.json" \
  --output-dir "/absolute/output/path"
```

常用参数：`--since-days`（增量窗口）、`--merge-xlsx`（合并进现有工作簿并高亮新增）、`--title-contains` / `--group-name`（发布者页过滤与命名）、`--raw-only`（只输出原始 CSV/JSON）、`--workers 1..8`（并发数，默认 4）。

## HTTP API

| 路由 | 说明 |
| --- | --- |
| `GET /` `GET /styles.css` `GET /app.js` | 界面静态资源 |
| `GET /api/status` | 当前更新状态 / 进度 |
| `GET /api/recent?days=&groups=&q=&all=` | 近期发布 / 全部已加载数据 |
| `GET /api/search?q=&groups=` | 全站实时搜索（2 分钟缓存） |
| `GET /api/workbook` | 下载当日 `DMHY更新表_YYYY-MM-DD.xlsx`（按需生成） |
| `GET /api/version` | 更新日志条目 |
| `GET /api/health` | 健康检查 |

## 目录结构

```
DMHY追番助手/
├── app.py                 # HTTP 服务 + 更新工作线程
├── dmhy_scraper.py        # 抓取 / 整理 CLI
├── requirements.txt       # openpyxl
├── static/
│   ├── index.html         # 单页界面
│   ├── app.js             # 前端逻辑
│   └── styles.css         # 样式
├── 启动DMHY追番助手.command # macOS launchd 启动器
└── com.dmhy.zhufan.plist  # launchd 服务定义
```

## 说明

- 数据来源为 DMHY 公开页面，仅用于个人学习与整理，请遵守站点条款。
- 追番、已使用、全站开关等偏好保存在浏览器 localStorage。
- 界面为单页全宽布局，桌面与移动端自适应。
