"""
OHTC 專案通知系統
=================
支援：
- Email 通知
- Microsoft Teams Webhook
- Slack Webhook
- Line Notify
- 自訂 Webhook
"""

import smtplib
import requests
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pandas as pd
from typing import List, Dict, Optional
import os


class NotificationConfig:
    """通知設定"""
    def __init__(self):
        # Email 設定
        self.email_enabled = False
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', 587))
        self.smtp_user = os.getenv('SMTP_USER', '')
        self.smtp_password = os.getenv('SMTP_PASSWORD', '')
        self.email_recipients = os.getenv('EMAIL_RECIPIENTS', '').split(',')
        
        # Teams 設定
        self.teams_enabled = False
        self.teams_webhook_url = os.getenv('TEAMS_WEBHOOK_URL', '')
        
        # Slack 設定
        self.slack_enabled = False
        self.slack_webhook_url = os.getenv('SLACK_WEBHOOK_URL', '')
        
        # Line Notify 設定
        self.line_enabled = False
        self.line_token = os.getenv('LINE_NOTIFY_TOKEN', '')


class ProjectNotifier:
    """專案通知器"""
    
    def __init__(self, config: NotificationConfig):
        self.config = config
    
    def send_delay_alert(self, delay_tasks: List[Dict], project_name: str = 'OHTC 專案'):
        """發送延遲警報"""
        if not delay_tasks:
            return
        
        title = f"⚠️ {project_name} - 延遲警報"
        
        message = f"""
**{project_name} 延遲警報**

共有 {len(delay_tasks)} 個任務延遲，需要立即關注：

"""
        for task in delay_tasks[:10]:
            message += f"- **{task['task']}** ({task['owner']})\n"
            message += f"  誤差: {task.get('variance_days', 'N/A')} 天\n"
        
        if len(delay_tasks) > 10:
            message += f"\n... 還有 {len(delay_tasks) - 10} 個延遲項目"
        
        message += f"\n\n發送時間: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        self._send_all(title, message)
    
    def send_daily_summary(self, summary: Dict, project_name: str = 'OHTC 專案'):
        """發送每日摘要"""
        title = f"📊 {project_name} - 每日摘要"
        
        total = summary['total']
        done = summary['done']
        going = summary['going']
        delay = summary['delay']
        
        message = f"""
**{project_name} 每日摘要**

📊 **進度概況**
- 總任務: {total} 項
- ✅ 已完成: {done} 項 ({done/total*100:.1f}%)
- 🔄 進行中: {going} 項
- ⚠️ 延遲中: {delay} 項

📅 **即將到期** (7天內)
"""
        for task in summary.get('upcoming', [])[:5]:
            message += f"- {task['task']} ({task['owner']})\n"
        
        message += f"\n發送時間: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        self._send_all(title, message)
    
    def send_weekly_report(self, report_content: str, project_name: str = 'OHTC 專案'):
        """發送週報"""
        title = f"📋 {project_name} - 週報"
        self._send_all(title, report_content)
    
    def send_milestone_complete(self, milestone: str, project_name: str = 'OHTC 專案'):
        """發送里程碑完成通知"""
        title = f"🎉 {project_name} - 里程碑完成"
        message = f"""
**恭喜！里程碑已完成**

🎯 **{milestone}**

完成時間: {datetime.now().strftime('%Y-%m-%d %H:%M')}
"""
        self._send_all(title, message)
    
    def _send_all(self, title: str, message: str):
        """發送到所有啟用的通道"""
        if self.config.email_enabled:
            self._send_email(title, message)
        
        if self.config.teams_enabled:
            self._send_teams(title, message)
        
        if self.config.slack_enabled:
            self._send_slack(title, message)
        
        if self.config.line_enabled:
            self._send_line(f"{title}\n\n{message}")
    
    def _send_email(self, subject: str, body: str):
        """發送 Email"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.config.smtp_user
            msg['To'] = ', '.join(self.config.email_recipients)
            msg['Subject'] = subject
            
            # 轉換 Markdown 為 HTML
            html_body = self._markdown_to_html(body)
            msg.attach(MIMEText(html_body, 'html'))
            
            with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
                server.starttls()
                server.login(self.config.smtp_user, self.config.smtp_password)
                server.send_message(msg)
            
            print(f"✅ Email 已發送至 {len(self.config.email_recipients)} 位收件人")
        except Exception as e:
            print(f"❌ Email 發送失敗: {e}")
    
    def _send_teams(self, title: str, message: str):
        """發送 Microsoft Teams 訊息"""
        try:
            payload = {
                "@type": "MessageCard",
                "@context": "http://schema.org/extensions",
                "themeColor": "0076D7",
                "summary": title,
                "sections": [{
                    "activityTitle": title,
                    "text": message.replace('\n', '<br>'),
                    "markdown": True
                }]
            }
            
            response = requests.post(
                self.config.teams_webhook_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.status_code == 200:
                print("✅ Teams 訊息已發送")
            else:
                print(f"❌ Teams 發送失敗: {response.status_code}")
        except Exception as e:
            print(f"❌ Teams 發送失敗: {e}")
    
    def _send_slack(self, title: str, message: str):
        """發送 Slack 訊息"""
        try:
            payload = {
                "blocks": [
                    {
                        "type": "header",
                        "text": {
                            "type": "plain_text",
                            "text": title
                        }
                    },
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": message
                        }
                    }
                ]
            }
            
            response = requests.post(
                self.config.slack_webhook_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.status_code == 200:
                print("✅ Slack 訊息已發送")
            else:
                print(f"❌ Slack 發送失敗: {response.status_code}")
        except Exception as e:
            print(f"❌ Slack 發送失敗: {e}")
    
    def _send_line(self, message: str):
        """發送 Line Notify 訊息"""
        try:
            headers = {
                'Authorization': f'Bearer {self.config.line_token}',
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = requests.post(
                'https://notify-api.line.me/api/notify',
                headers=headers,
                data={'message': message}
            )
            
            if response.status_code == 200:
                print("✅ Line 訊息已發送")
            else:
                print(f"❌ Line 發送失敗: {response.status_code}")
        except Exception as e:
            print(f"❌ Line 發送失敗: {e}")
    
    def _markdown_to_html(self, md: str) -> str:
        """簡易 Markdown 轉 HTML"""
        html = md
        html = html.replace('**', '<strong>').replace('**', '</strong>')
        html = html.replace('\n', '<br>')
        html = html.replace('- ', '• ')
        return f"<html><body style='font-family: Arial, sans-serif;'>{html}</body></html>"


class ScheduledNotifier:
    """排程通知器"""
    
    def __init__(self, notifier: ProjectNotifier, excel_path: str):
        self.notifier = notifier
        self.excel_path = excel_path
    
    def load_data(self) -> pd.DataFrame:
        """載入資料"""
        df = pd.read_excel(self.excel_path, sheet_name='軟體時程', header=None)
        
        tasks = []
        for i in range(6, len(df)):
            row = df.iloc[i]
            task_name = row[0]
            
            if pd.notna(task_name) and str(task_name).strip():
                task = {
                    'task': str(task_name).strip(),
                    'owner': str(row[2]) if pd.notna(row[2]) else '',
                    'status': str(row[7]) if pd.notna(row[7]) else '',
                    'plan_end': pd.to_datetime(row[9]) if pd.notna(row[9]) else None,
                    'variance_days': int(row[14]) if pd.notna(row[14]) else 0,
                }
                tasks.append(task)
        
        return pd.DataFrame(tasks)
    
    def check_and_notify(self):
        """檢查並發送通知"""
        df = self.load_data()
        today = datetime.now()
        
        # 檢查延遲項目
        delay_tasks = df[df['status'] == 'Delay'].to_dict('records')
        if delay_tasks:
            self.notifier.send_delay_alert(delay_tasks)
        
        # 檢查即將到期
        upcoming = df[
            (df['status'] == 'Going') & 
            (df['plan_end'].notna()) &
            (df['plan_end'] <= today + timedelta(days=3))
        ]
        
        if not upcoming.empty:
            self.notifier.send_delay_alert(
                upcoming.to_dict('records'),
                project_name='OHTC 專案 - 即將到期提醒'
            )
    
    def send_summary(self):
        """發送摘要"""
        df = self.load_data()
        today = datetime.now()
        
        summary = {
            'total': len(df),
            'done': len(df[df['status'] == 'Done']),
            'going': len(df[df['status'] == 'Going']),
            'delay': len(df[df['status'] == 'Delay']),
            'upcoming': df[
                (df['status'] == 'Going') & 
                (df['plan_end'].notna()) &
                (df['plan_end'] <= today + timedelta(days=7))
            ].to_dict('records'),
        }
        
        self.notifier.send_daily_summary(summary)


# 使用範例
if __name__ == '__main__':
    # 設定通知配置
    config = NotificationConfig()
    
    # 啟用 Teams 通知（範例）
    # config.teams_enabled = True
    # config.teams_webhook_url = 'YOUR_TEAMS_WEBHOOK_URL'
    
    # 啟用 Email 通知（範例）
    # config.email_enabled = True
    # config.smtp_user = 'your_email@company.com'
    # config.smtp_password = 'your_password'
    # config.email_recipients = ['team@company.com']
    
    notifier = ProjectNotifier(config)
    
    # 測試發送
    test_tasks = [
        {'task': 'HA測試', 'owner': 'OHTC：銅鑼測試', 'variance_days': -5},
        {'task': 'Golden驗證', 'owner': 'OHTL', 'variance_days': -10},
    ]
    
    # notifier.send_delay_alert(test_tasks, 'SPIL EL P2 AMHS')
    
    print("通知系統已初始化。請設定 webhook URL 或 SMTP 後使用。")
