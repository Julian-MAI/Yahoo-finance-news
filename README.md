# 雅虎财经新闻推送系统

基于 Yahoo Finance RSS 的新闻抓取与中文报告生成工具。

功能概览：
- 多 RSS 源抓取（综合 + 个股）
- 正文多策略解析（含 RSS 摘要兜底）
- 英文自动翻译为中文
- 自动分类（政策/宏观经济、行业动态、公司新闻）
- 生成高质量 Word 报告
- 可选推送到企业微信、钉钉、Slack、Telegram

## 1. 环境要求

- Python 3.9+
- 可访问外网（用于 RSS 抓取与翻译）

安装依赖：

```bash
pip install requests feedparser beautifulsoup4 deep-translator python-docx
```

## 2. 文件说明

- `news_bot_full.py`：主程序
- `push_config.json`：推送配置
- `output/`：最新输出目录
- `history/`：历史报告目录

## 3. 推送配置（可选）

编辑 `push_config.json`：

```json
{
  "wechat_work_webhook": "",
  "dingtalk_webhook": "",
  "slack_webhook": "",
  "telegram_bot_token": "",
  "telegram_chat_id": ""
}
```

说明：
- 不填则对应渠道不会推送。
- 建议至少配置 1 个渠道用于接收摘要通知。

## 4. 运行方式

基础运行：

```bash
python news_bot_full.py
```

常用参数：

```bash
python news_bot_full.py --max-per-category 5 --max-total 20 --report-format word --no-push
```

参数说明：
- `--max-per-category`：每个分类最多处理条数，默认 `5`
- `--max-total`：总处理条数上限，默认 `20`
- `--output-dir`：输出目录，默认 `output/`
- `--report-format`：
  - `word`：仅生成 Word（默认，推荐）
  - `all`：生成 Word + TXT + JSON
- `--no-push`：只生成报告，不执行推送

## 5. 输出结果

默认模式（`--report-format word`）：
- `output/latest_report.docx`
- `output/history/news_report_YYYYMMDD.docx`

完整模式（`--report-format all`）额外生成：
- `output/latest_summary.json`
- `output/latest_detail.txt`
- `output/latest_summary.txt`
- `output/history/news_summary_YYYYMMDD.json`
- `output/history/news_detail_YYYYMMDD.txt`
- `output/history/news_summary_YYYYMMDD.txt`

## 6. 运行流程

程序执行顺序：
1. 拉取 RSS 新闻
2. 按关键词分类
3. 抓取正文
4. 翻译为中文
5. 生成报告
6. 发送摘要推送（可选）

## 7. 常见问题

1. 抓取失败或正文很短
- 可能是目标页面结构变化或网络问题，程序会自动回退到 RSS 摘要。

2. 翻译失败
- 程序内置重试机制，连续失败时会保留原文。

3. 没有收到推送
- 检查 `push_config.json` 是否填写正确。
- 使用 `--no-push` 时不会推送。

## 8. 推荐命令

日常只要 Word 报告：

```bash
python news_bot_full.py --report-format word --no-push
```

需要完整归档数据：

```bash
python news_bot_full.py --report-format all
```
