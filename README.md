# 雅虎财经新闻定时推送系统

## 设置定时任务 (Cron)

### 方法1: 使用系统Cron (Linux/Mac)

编辑crontab文件:
```bash
crontab -e
```

添加以下行（每天早上9点执行）:
```
0 9 * * * cd /mnt/okcomputer/output/finance_news_bot && /usr/bin/python3 news_bot.py >> /mnt/okcomputer/output/finance_news_bot/cron.log 2>&1
```

### 方法2: 使用Python schedule库

安装schedule:
```bash
pip install schedule
```

运行调度脚本:
```bash
python3 scheduler.py
```

### 方法3: 使用系统服务 (systemd)

创建服务文件 `/etc/systemd/system/finance-news.service`:
```ini
[Unit]
Description=Finance News Daily Push
After=network.target

[Service]
Type=oneshot
ExecStart=/usr/bin/python3 /mnt/okcomputer/output/finance_news_bot/news_bot.py
User=your_username

[Install]
WantedBy=multi-user.target
```

创建定时器文件 `/etc/systemd/system/finance-news.timer`:
```ini
[Unit]
Description=Run Finance News Bot daily at 9:00 AM

[Timer]
OnCalendar=*-*-* 09:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

启用并启动定时器:
```bash
sudo systemctl daemon-reload
sudo systemctl enable finance-news.timer
sudo systemctl start finance-news.timer
```

## 查看定时任务状态

```bash
# 查看cron任务
crontab -l

# 查看systemd定时器
systemctl list-timers --all

# 查看日志
journalctl -u finance-news.service
```

## 推送方式扩展

当前系统生成报告文件，可通过以下方式推送:

1. **邮件推送**: 配置SMTP服务器
2. **企业微信**: 使用Webhook
3. **钉钉**: 使用Webhook
4. **Slack**: 使用Incoming Webhook
5. **Telegram**: 使用Bot API
6. **短信**: 使用云服务商SMS API
