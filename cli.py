#!/usr/bin/env python3
"""
OHTC 專案管理 CLI 工具
=======================
用於快速查詢和更新專案狀態

使用方式:
    python cli.py status              # 查看專案狀態摘要
    python cli.py delay               # 列出延遲項目
    python cli.py upcoming            # 列出即將到期項目
    python cli.py search <keyword>    # 搜尋任務
    python cli.py report              # 生成週報
"""

import argparse
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import sys

# 顏色輸出
class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    END = '\033[0m'


def load_data(file_path):
    """載入 Excel 資料"""
    try:
        df = pd.read_excel(file_path, sheet_name='軟體時程', header=None)
        
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
    except Exception as e:
        print(f"{Colors.RED}錯誤: 無法載入檔案 - {e}{Colors.END}")
        sys.exit(1)


def cmd_status(df):
    """顯示專案狀態摘要"""
    total = len(df)
    done = len(df[df['status'] == 'Done'])
    going = len(df[df['status'] == 'Going'])
    delay = len(df[df['status'] == 'Delay'])
    
    print(f"\n{Colors.BOLD}═══════════════════════════════════════{Colors.END}")
    print(f"{Colors.BOLD}  📊 OHTC 專案狀態摘要{Colors.END}")
    print(f"{Colors.BOLD}═══════════════════════════════════════{Colors.END}\n")
    
    print(f"  📋 總任務數:   {Colors.BOLD}{total}{Colors.END}")
    print(f"  {Colors.GREEN}✅ 已完成:     {done} ({done/total*100:.1f}%){Colors.END}")
    print(f"  {Colors.YELLOW}🔄 進行中:     {going} ({going/total*100:.1f}%){Colors.END}")
    print(f"  {Colors.RED}⚠️  延遲中:     {delay} ({delay/total*100:.1f}%){Colors.END}")
    
    # 進度條
    progress = done / total
    bar_length = 30
    filled = int(bar_length * progress)
    bar = '█' * filled + '░' * (bar_length - filled)
    print(f"\n  進度: [{Colors.GREEN}{bar}{Colors.END}] {progress*100:.1f}%")
    
    print(f"\n{Colors.BOLD}═══════════════════════════════════════{Colors.END}\n")


def cmd_delay(df):
    """列出延遲項目"""
    delay_df = df[df['status'] == 'Delay']
    
    print(f"\n{Colors.RED}{Colors.BOLD}⚠️  延遲項目清單 ({len(delay_df)} 項){Colors.END}\n")
    
    if delay_df.empty:
        print(f"  {Colors.GREEN}🎉 太棒了！沒有延遲項目！{Colors.END}\n")
        return
    
    for idx, task in delay_df.iterrows():
        variance = task['variance_days']
        risk = "🔴 高" if abs(variance) > 7 else "🟡 中" if abs(variance) > 3 else "🟢 低"
        
        print(f"  {Colors.RED}●{Colors.END} {task['task'][:40]}")
        print(f"    負責: {task['owner']:<15} 誤差: {variance:+d} 天  風險: {risk}")
        print()


def cmd_upcoming(df, days=7):
    """列出即將到期項目"""
    today = datetime.now()
    upcoming = df[
        (df['status'] == 'Going') & 
        (df['plan_end'].notna()) &
        (df['plan_end'] <= today + timedelta(days=days)) &
        (df['plan_end'] >= today)
    ]
    
    print(f"\n{Colors.YELLOW}{Colors.BOLD}⏰ 即將到期項目 ({days} 天內, {len(upcoming)} 項){Colors.END}\n")
    
    if upcoming.empty:
        print(f"  {Colors.GREEN}✓ 近期沒有到期項目{Colors.END}\n")
        return
    
    for idx, task in upcoming.iterrows():
        days_left = (task['plan_end'] - today).days
        urgency = "🔴" if days_left <= 2 else "🟡" if days_left <= 5 else "🟢"
        
        print(f"  {urgency} {task['task'][:40]}")
        print(f"    負責: {task['owner']:<15} 剩餘: {days_left} 天  截止: {task['plan_end'].strftime('%m/%d')}")
        print()


def cmd_search(df, keyword):
    """搜尋任務"""
    results = df[df['task'].str.contains(keyword, case=False, na=False)]
    
    print(f"\n{Colors.CYAN}{Colors.BOLD}🔍 搜尋結果: '{keyword}' ({len(results)} 項){Colors.END}\n")
    
    if results.empty:
        print(f"  找不到包含 '{keyword}' 的任務\n")
        return
    
    status_colors = {
        'Done': Colors.GREEN,
        'Going': Colors.YELLOW,
        'Delay': Colors.RED,
    }
    
    for idx, task in results.iterrows():
        color = status_colors.get(task['status'], Colors.WHITE)
        status_icon = {'Done': '✅', 'Going': '🔄', 'Delay': '⚠️'}.get(task['status'], '❓')
        
        print(f"  {status_icon} {color}{task['task'][:50]}{Colors.END}")
        print(f"    負責: {task['owner']:<15} 狀態: {task['status']}")
        print()


def cmd_report(df):
    """生成簡易週報"""
    today = datetime.now()
    week_start = today - timedelta(days=today.weekday())
    
    total = len(df)
    done = len(df[df['status'] == 'Done'])
    delay = len(df[df['status'] == 'Delay'])
    
    print(f"\n{Colors.BOLD}{'═' * 50}{Colors.END}")
    print(f"{Colors.BOLD}  📋 OHTC 專案週報{Colors.END}")
    print(f"{Colors.BOLD}{'═' * 50}{Colors.END}")
    print(f"\n  報告日期: {today.strftime('%Y-%m-%d')}")
    print(f"  報告週期: {week_start.strftime('%Y-%m-%d')} ~ {(week_start + timedelta(days=6)).strftime('%Y-%m-%d')}")
    
    print(f"\n{Colors.BOLD}  【進度概況】{Colors.END}")
    print(f"  - 總任務: {total} 項")
    print(f"  - 已完成: {done} 項 ({done/total*100:.1f}%)")
    print(f"  - 延遲中: {delay} 項")
    
    print(f"\n{Colors.BOLD}  【延遲項目】{Colors.END}")
    delay_df = df[df['status'] == 'Delay']
    if delay_df.empty:
        print("  - 無延遲項目 ✅")
    else:
        for _, task in delay_df.head(5).iterrows():
            print(f"  - {task['task'][:35]} ({task['owner']})")
        if len(delay_df) > 5:
            print(f"  - ... 還有 {len(delay_df) - 5} 項")
    
    print(f"\n{Colors.BOLD}{'═' * 50}{Colors.END}\n")


def cmd_owner(df, owner_name=None):
    """按負責單位統計"""
    owner_stats = df.groupby('owner').agg({
        'task': 'count',
        'status': lambda x: (x == 'Done').sum()
    }).reset_index()
    owner_stats.columns = ['owner', 'total', 'done']
    owner_stats['pending'] = owner_stats['total'] - owner_stats['done']
    owner_stats = owner_stats[owner_stats['owner'] != ''].sort_values('total', ascending=False)
    
    print(f"\n{Colors.PURPLE}{Colors.BOLD}👥 負責單位工作量統計{Colors.END}\n")
    
    print(f"  {'負責單位':<20} {'總數':>6} {'完成':>6} {'待辦':>6} {'完成率':>8}")
    print(f"  {'-' * 50}")
    
    for _, row in owner_stats.iterrows():
        rate = row['done'] / row['total'] * 100 if row['total'] > 0 else 0
        color = Colors.GREEN if rate >= 70 else Colors.YELLOW if rate >= 30 else Colors.RED
        print(f"  {row['owner']:<20} {row['total']:>6} {row['done']:>6} {row['pending']:>6} {color}{rate:>7.1f}%{Colors.END}")
    
    print()


def main():
    parser = argparse.ArgumentParser(
        description='OHTC 專案管理 CLI 工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
    python cli.py status
    python cli.py delay
    python cli.py upcoming --days 14
    python cli.py search "OHTC"
    python cli.py owner
    python cli.py report
        """
    )
    
    parser.add_argument('command', choices=['status', 'delay', 'upcoming', 'search', 'report', 'owner'],
                       help='要執行的命令')
    parser.add_argument('keyword', nargs='?', default='', help='搜尋關鍵字 (用於 search 命令)')
    parser.add_argument('-f', '--file', default='schedule.xlsx', help='Excel 檔案路徑')
    parser.add_argument('-d', '--days', type=int, default=7, help='天數 (用於 upcoming 命令)')
    
    args = parser.parse_args()
    
    # 尋找 Excel 檔案
    file_path = Path(args.file)
    if not file_path.exists():
        # 嘗試在當前目錄找 xlsx 檔案
        xlsx_files = list(Path('.').glob('*.xlsx'))
        if xlsx_files:
            file_path = xlsx_files[0]
            print(f"{Colors.CYAN}使用檔案: {file_path}{Colors.END}")
        else:
            print(f"{Colors.RED}錯誤: 找不到 Excel 檔案{Colors.END}")
            print(f"請使用 -f 參數指定檔案路徑")
            sys.exit(1)
    
    df = load_data(file_path)
    
    if args.command == 'status':
        cmd_status(df)
    elif args.command == 'delay':
        cmd_delay(df)
    elif args.command == 'upcoming':
        cmd_upcoming(df, args.days)
    elif args.command == 'search':
        if not args.keyword:
            print(f"{Colors.RED}錯誤: search 命令需要提供關鍵字{Colors.END}")
            sys.exit(1)
        cmd_search(df, args.keyword)
    elif args.command == 'report':
        cmd_report(df)
    elif args.command == 'owner':
        cmd_owner(df)


if __name__ == '__main__':
    main()
