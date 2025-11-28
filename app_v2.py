"""
OHTC 專案管理儀表板 v2.0
================================
新增功能：
- 週報/月報自動生成
- 進度趨勢分析
- 風險評估矩陣
- 資源負載分析
- 里程碑追蹤
- 任務編輯與儲存
- 多工作表完整解析
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import io
import json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="OHTC 專案管理儀表板 v2.0",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        background: linear-gradient(90deg, #1f77b4, #9467bd);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .risk-high { background-color: #dc3545; color: white; padding: 5px 10px; border-radius: 4px; }
    .risk-medium { background-color: #ffc107; color: black; padding: 5px 10px; border-radius: 4px; }
    .risk-low { background-color: #28a745; color: white; padding: 5px 10px; border-radius: 4px; }
    .milestone-done { border-left: 4px solid #28a745; }
    .milestone-pending { border-left: 4px solid #ffc107; }
    .report-section { 
        background: #f8f9fa; 
        border-radius: 8px; 
        padding: 20px; 
        margin: 10px 0;
    }
    div[data-testid="stExpander"] details summary p {
        font-size: 1.1rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 資料載入與解析
# ============================================================
@st.cache_data
def load_excel_data(uploaded_file):
    """載入 Excel 檔案並解析各工作表"""
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # 讀取軟體時程表
        df_software = pd.read_excel(uploaded_file, sheet_name='軟體時程', header=None)
        
        # 提取專案資訊
        project_info = {
            'project_code': str(df_software.iloc[2, 2]) if pd.notna(df_software.iloc[2, 2]) else '',
            'project_name': str(df_software.iloc[3, 2]) if pd.notna(df_software.iloc[3, 2]) else '',
            'project_lead': str(df_software.iloc[4, 2]) if pd.notna(df_software.iloc[4, 2]) else '',
            'start_date': df_software.iloc[3, 9] if pd.notna(df_software.iloc[3, 9]) else None,
            'update_date': df_software.iloc[4, 12] if pd.notna(df_software.iloc[4, 12]) else None,
        }
        
        # 安全的數字轉換函數（定義在外層，避免重複定義）
        def safe_float(val, default=0):
            try:
                if pd.isna(val):
                    return default
                # 如果是字串且包含非數字字符（如標題），返回預設值
                if isinstance(val, str):
                    # 移除空白和換行
                    val_clean = str(val).strip()
                    # 檢查是否包含中文或其他非數字字符
                    if any(ord(c) > 127 for c in val_clean) or not val_clean.replace('.', '', 1).replace('-', '', 1).replace('+', '', 1).isdigit():
                        return default
                return float(val)
            except (ValueError, TypeError):
                return default

        def safe_int(val, default=0):
            try:
                if pd.isna(val):
                    return default
                if isinstance(val, str):
                    val_clean = str(val).strip()
                    if any(ord(c) > 127 for c in val_clean) or not val_clean.replace('-', '', 1).replace('+', '', 1).isdigit():
                        return default
                return int(float(val))
            except (ValueError, TypeError):
                return default

        def safe_datetime(val):
            try:
                if pd.isna(val):
                    return None
                return pd.to_datetime(val)
            except:
                return None

        # 解析任務資料
        tasks = []
        for i in range(6, len(df_software)):
            row = df_software.iloc[i]
            task_name = row[0]

            if pd.notna(task_name) and str(task_name).strip():
                # 跳過標題行（檢查是否 row[4] 包含 "百分比" 等關鍵字）
                if isinstance(row[4], str) and ('百分比' in str(row[4]) or '完成' in str(row[4])):
                    continue

                task = {
                    'id': len(tasks) + 1,
                    'row_index': i,
                    'task': str(task_name).strip(),
                    'owner': str(row[2]) if pd.notna(row[2]) else '',
                    'progress_pct': safe_float(row[4]),
                    'target_pct': safe_float(row[5]),
                    'remaining_days': safe_int(row[6]),
                    'status': str(row[7]) if pd.notna(row[7]) else '',
                    'plan_start': safe_datetime(row[8]),
                    'plan_end': safe_datetime(row[9]),
                    'plan_days': safe_int(row[10]),
                    'actual_start': safe_datetime(row[11]),
                    'actual_end': safe_datetime(row[12]),
                    'actual_days': safe_int(row[13]),
                    'variance_days': safe_int(row[14]),
                    'coord_time': str(row[15]) if pd.notna(row[15]) else '',
                    'coord_manpower': str(row[16]) if pd.notna(row[16]) else '',
                    'coord_area': str(row[17]) if pd.notna(row[17]) else '',
                    'coord_equipment': str(row[18]) if pd.notna(row[18]) else '',
                    'notes': str(row[19]) if pd.notna(row[19]) else '',
                }
                tasks.append(task)
        
        df_tasks = pd.DataFrame(tasks)
        
        # 讀取系統時程
        df_system = pd.read_excel(uploaded_file, sheet_name='系統時程_C', header=None)
        system_items = []
        current_area = ''
        for i in range(5, len(df_system)):
            row = df_system.iloc[i]
            item_name = str(row[0]).strip() if pd.notna(row[0]) else ''

            if item_name:
                # 檢查是否為區域標題
                if '區域' in item_name:
                    current_area = item_name

                item = {
                    'area': current_area,
                    'item': item_name,
                    'target_date': row[1] if pd.notna(row[1]) else None,
                    'completion_pct': safe_float(row[2]),  # 使用 safe_float 而不是 float
                    'notes': str(row[3]) if pd.notna(row[3]) else '',
                    'is_area': '區域' in item_name,
                }
                system_items.append(item)
        df_system_tasks = pd.DataFrame(system_items)
        
        # 讀取工程進度確認表
        try:
            df_engineering = pd.read_excel(uploaded_file, sheet_name='工程_工作進度確認表', header=None)
        except:
            df_engineering = pd.DataFrame()
        
        # 讀取 EQ 工作清單
        try:
            df_eq = pd.read_excel(uploaded_file, sheet_name='EQ 工作清單', header=None)
        except:
            df_eq = pd.DataFrame()
        
        return {
            'project_info': project_info,
            'tasks': df_tasks,
            'system_tasks': df_system_tasks,
            'engineering': df_engineering,
            'eq_list': df_eq,
            'raw_software': df_software,
            'sheet_names': sheet_names,
        }
    except Exception as e:
        st.error(f"載入檔案錯誤: {str(e)}")
        return None


# ============================================================
# 圖表生成函數
# ============================================================
def create_gantt_chart(df_tasks, show_actual=False):
    """建立甘特圖"""
    gantt_data = df_tasks[df_tasks['plan_start'].notna() & df_tasks['plan_end'].notna()].copy()

    if gantt_data.empty:
        return None

    color_map = {
        'Done': '#28a745',
        'Going': '#ffc107',
        'Delay': '#dc3545',
        '': '#6c757d'
    }

    fig = go.Figure()

    # 計劃時程
    for idx, row in gantt_data.iterrows():
        try:
            # 確保日期是 datetime 類型
            plan_start = pd.to_datetime(row['plan_start'])
            plan_end = pd.to_datetime(row['plan_end'])

            fig.add_trace(go.Bar(
                name='計劃',
                y=[row['task']],
                x=[(plan_end - plan_start).days],
                base=plan_start,
                orientation='h',
                marker_color=color_map.get(row['status'], '#6c757d'),
                opacity=0.8,
                hovertemplate=f"<b>{row['task']}</b><br>" +
                             f"計劃: {plan_start.strftime('%Y-%m-%d')} ~ {plan_end.strftime('%Y-%m-%d')}<br>" +
                             f"狀態: {row['status']}<br>" +
                             f"負責: {row['owner']}<extra></extra>",
                showlegend=False,
            ))
        except Exception as e:
            continue  # 跳過有問題的資料

    # 實際時程（如果有）
    if show_actual:
        actual_data = gantt_data[gantt_data['actual_start'].notna() & gantt_data['actual_end'].notna()]
        for idx, row in actual_data.iterrows():
            try:
                actual_start = pd.to_datetime(row['actual_start'])
                actual_end = pd.to_datetime(row['actual_end'])

                fig.add_trace(go.Bar(
                    name='實際',
                    y=[row['task']],
                    x=[(actual_end - actual_start).days],
                    base=actual_start,
                    orientation='h',
                    marker_color='rgba(0,0,0,0.3)',
                    marker_line_color='black',
                    marker_line_width=2,
                    opacity=0.5,
                    showlegend=False,
                ))
            except Exception as e:
                continue

    # 設定版面配置
    fig.update_layout(
        title='📅 專案甘特圖',
        height=max(500, len(gantt_data) * 28),
        xaxis_title='日期',
        yaxis_title='',
        barmode='overlay',
        yaxis={'categoryorder': 'total ascending'},
        xaxis={'type': 'date'},
    )

    # 加入今日線（使用 add_shape 而不是 add_vline，避免日期格式問題）
    try:
        today = pd.Timestamp.now()
        fig.add_shape(
            type="line",
            x0=today, x1=today,
            y0=0, y1=1,
            yref="paper",
            line=dict(color="red", width=2, dash="dash"),
        )
        fig.add_annotation(
            x=today, y=1,
            yref="paper",
            text="今日",
            showarrow=False,
            yshift=10,
            font=dict(color="red", size=12)
        )
    except Exception as e:
        pass  # 如果加今日線失敗，就不加

    return fig


def create_status_pie(df_tasks):
    """狀態圓餅圖"""
    if df_tasks.empty:
        return None

    status_counts = df_tasks['status'].value_counts()

    if status_counts.empty:
        return None

    colors = {'Done': '#28a745', 'Going': '#ffc107', 'Delay': '#dc3545', '': '#6c757d'}

    fig = go.Figure(data=[go.Pie(
        labels=status_counts.index,
        values=status_counts.values,
        hole=0.4,
        marker_colors=[colors.get(s, '#6c757d') for s in status_counts.index],
        textinfo='value+percent',
        textposition='inside',
    )])

    fig.update_layout(title='📊 任務狀態分佈', height=350)
    return fig


def create_owner_workload(df_tasks):
    """負責單位工作量"""
    if df_tasks.empty:
        return None

    owner_stats = df_tasks.groupby('owner').agg({
        'task': 'count',
        'status': lambda x: list(x)
    }).reset_index()

    owner_stats['done'] = owner_stats['status'].apply(lambda x: x.count('Done'))
    owner_stats['going'] = owner_stats['status'].apply(lambda x: x.count('Going'))
    owner_stats['delay'] = owner_stats['status'].apply(lambda x: x.count('Delay'))
    owner_stats = owner_stats[owner_stats['owner'] != ''].sort_values('task', ascending=True)

    if owner_stats.empty:
        return None

    fig = go.Figure()
    fig.add_trace(go.Bar(name='已完成', y=owner_stats['owner'], x=owner_stats['done'],
                        orientation='h', marker_color='#28a745'))
    fig.add_trace(go.Bar(name='進行中', y=owner_stats['owner'], x=owner_stats['going'],
                        orientation='h', marker_color='#ffc107'))
    fig.add_trace(go.Bar(name='延遲', y=owner_stats['owner'], x=owner_stats['delay'],
                        orientation='h', marker_color='#dc3545'))

    fig.update_layout(
        barmode='stack',
        title='👥 各負責單位工作量',
        height=max(300, len(owner_stats) * 30),
        xaxis_title='任務數量',
    )
    return fig


def create_progress_trend(df_tasks):
    """進度趨勢圖（模擬）"""
    if df_tasks.empty:
        return None

    # 根據計劃完成日期模擬進度
    dates = pd.date_range(start='2025-05-01', end='2025-09-30', freq='W')

    progress_data = []
    for date in dates:
        done = len(df_tasks[(df_tasks['plan_end'].notna()) & (df_tasks['plan_end'] <= date)])
        total = len(df_tasks)
        progress_data.append({
            'date': date,
            'completed': done,
            'completion_rate': done / total * 100 if total > 0 else 0
        })

    df_progress = pd.DataFrame(progress_data)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(x=df_progress['date'], y=df_progress['completed'],
               name='累計完成數', marker_color='#28a745', opacity=0.7),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(x=df_progress['date'], y=df_progress['completion_rate'],
                  name='完成率 %', line=dict(color='#1f77b4', width=3)),
        secondary_y=True,
    )

    fig.update_layout(title='📈 進度趨勢圖', height=400)
    fig.update_yaxes(title_text="完成數量", secondary_y=False)
    fig.update_yaxes(title_text="完成率 (%)", secondary_y=True, range=[0, 100])

    return fig


def create_risk_matrix(df_tasks):
    """風險評估矩陣"""
    delay_tasks = df_tasks[df_tasks['status'] == 'Delay'].copy()
    
    if delay_tasks.empty:
        return None
    
    # 計算風險等級（基於誤差天數）
    def calc_risk(variance):
        if pd.isna(variance) or variance == 0:
            return 'low'
        elif abs(variance) <= 7:
            return 'medium'
        else:
            return 'high'
    
    delay_tasks['risk_level'] = delay_tasks['variance_days'].apply(calc_risk)
    
    risk_colors = {'high': '#dc3545', 'medium': '#ffc107', 'low': '#28a745'}
    
    fig = go.Figure()
    
    for risk in ['high', 'medium', 'low']:
        risk_data = delay_tasks[delay_tasks['risk_level'] == risk]
        if not risk_data.empty:
            fig.add_trace(go.Scatter(
                x=risk_data['variance_days'].abs(),
                y=risk_data['plan_days'],
                mode='markers+text',
                name=f'{risk.upper()} 風險',
                marker=dict(size=15, color=risk_colors[risk]),
                text=risk_data['task'].str[:15],
                textposition='top center',
                hovertemplate='<b>%{text}</b><br>誤差: %{x} 天<br>計劃天數: %{y} 天<extra></extra>'
            ))
    
    fig.update_layout(
        title='⚠️ 風險評估矩陣',
        xaxis_title='誤差天數（絕對值）',
        yaxis_title='計劃天數',
        height=400,
    )
    
    return fig


def create_area_progress(df_system):
    """區域進度圖"""
    area_data = df_system[df_system['is_area'] == True].copy()
    
    if area_data.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=area_data['item'],
        y=area_data['completion_pct'] * 100,
        marker_color=area_data['completion_pct'].apply(
            lambda x: '#28a745' if x >= 0.7 else '#ffc107' if x >= 0.3 else '#dc3545'
        ),
        text=area_data['completion_pct'].apply(lambda x: f'{x*100:.0f}%'),
        textposition='outside',
    ))
    
    fig.update_layout(
        title='🏭 各區域完成進度',
        yaxis_title='完成率 (%)',
        yaxis_range=[0, 110],
        height=350,
    )
    
    return fig


# ============================================================
# 報表生成函數
# ============================================================
def generate_weekly_report(data, report_date=None):
    """生成週報"""
    if report_date is None:
        report_date = datetime.now()
    
    df_tasks = data['tasks']
    project_info = data['project_info']
    
    # 本週範圍
    week_start = report_date - timedelta(days=report_date.weekday())
    week_end = week_start + timedelta(days=6)
    
    # 統計數據
    total = len(df_tasks)
    done = len(df_tasks[df_tasks['status'] == 'Done'])
    going = len(df_tasks[df_tasks['status'] == 'Going'])
    delay = len(df_tasks[df_tasks['status'] == 'Delay'])
    
    # 本週完成的任務
    completed_this_week = df_tasks[
        (df_tasks['actual_end'].notna()) & 
        (df_tasks['actual_end'] >= week_start) & 
        (df_tasks['actual_end'] <= week_end)
    ]
    
    # 下週預計完成
    next_week_end = week_end + timedelta(days=7)
    planned_next_week = df_tasks[
        (df_tasks['plan_end'].notna()) & 
        (df_tasks['plan_end'] > week_end) & 
        (df_tasks['plan_end'] <= next_week_end) &
        (df_tasks['status'] != 'Done')
    ]
    
    report = f"""
# 📋 專案週報

**專案名稱：** {project_info['project_name']}  
**專案工令：** {project_info['project_code']}  
**報告日期：** {report_date.strftime('%Y-%m-%d')}  
**報告週期：** {week_start.strftime('%Y-%m-%d')} ~ {week_end.strftime('%Y-%m-%d')}

---

## 📊 整體進度概況

| 指標 | 數值 | 佔比 |
|------|------|------|
| 總任務數 | {total} | 100% |
| 已完成 | {done} | {done/total*100:.1f}% |
| 進行中 | {going} | {going/total*100:.1f}% |
| 延遲中 | {delay} | {delay/total*100:.1f}% |

**整體完成率：{done/total*100:.1f}%**

---

## ✅ 本週完成項目 ({len(completed_this_week)} 項)

"""
    
    if completed_this_week.empty:
        report += "本週無完成項目\n"
    else:
        for _, task in completed_this_week.iterrows():
            report += f"- {task['task']} ({task['owner']})\n"
    
    report += f"""
---

## 📅 下週計劃 ({len(planned_next_week)} 項)

"""
    
    if planned_next_week.empty:
        report += "下週無預計完成項目\n"
    else:
        for _, task in planned_next_week.iterrows():
            end_date = task['plan_end'].strftime('%m/%d') if pd.notna(task['plan_end']) else 'N/A'
            report += f"- {task['task']} (預計 {end_date}, {task['owner']})\n"
    
    report += f"""
---

## ⚠️ 風險與問題 ({delay} 項延遲)

"""
    
    delay_tasks = df_tasks[df_tasks['status'] == 'Delay']
    if delay_tasks.empty:
        report += "目前無延遲項目 ✅\n"
    else:
        for _, task in delay_tasks.head(10).iterrows():
            report += f"- **{task['task']}** - {task['owner']}\n"
    
    report += """
---

## 📝 備註

（請在此補充其他說明）

---
*此報告由 OHTC 專案管理儀表板自動生成*
"""
    
    return report


def generate_status_summary(data):
    """生成狀態摘要"""
    df_tasks = data['tasks']
    
    summary = {
        'total': len(df_tasks),
        'done': len(df_tasks[df_tasks['status'] == 'Done']),
        'going': len(df_tasks[df_tasks['status'] == 'Going']),
        'delay': len(df_tasks[df_tasks['status'] == 'Delay']),
        'delay_tasks': df_tasks[df_tasks['status'] == 'Delay'][['task', 'owner', 'plan_end', 'variance_days']].to_dict('records'),
        'upcoming': df_tasks[
            (df_tasks['status'] == 'Going') & 
            (df_tasks['plan_end'].notna()) &
            (df_tasks['plan_end'] <= datetime.now() + timedelta(days=7))
        ][['task', 'owner', 'plan_end']].to_dict('records'),
    }
    
    return summary


# ============================================================
# Excel 匯出函數
# ============================================================
def export_updated_excel(data, original_file, updated_tasks):
    """匯出更新後的 Excel"""
    output = io.BytesIO()
    original_file.seek(0)
    wb = load_workbook(original_file)
    ws = wb['軟體時程']
    
    # 更新任務狀態
    for _, task in updated_tasks.iterrows():
        row_num = task['row_index'] + 1  # openpyxl 從 1 開始
        ws.cell(row=row_num, column=8, value=task['status'])
        if pd.notna(task.get('notes')):
            ws.cell(row=row_num, column=20, value=task['notes'])
    
    # 更新日期
    ws.cell(row=5, column=13, value=datetime.now())
    
    wb.save(output)
    output.seek(0)
    return output


def export_report_to_word_format(report_content):
    """將報表匯出為可複製格式"""
    return report_content


# ============================================================
# 主應用程式
# ============================================================
def main():
    st.markdown('<h1 class="main-header">🏭 OHTC 專案管理儀表板 v2.0</h1>', unsafe_allow_html=True)
    
    # 側邊欄
    with st.sidebar:
        st.header("📁 檔案管理")
        uploaded_file = st.file_uploader(
            "上傳專案排程表",
            type=['xlsx', 'xls'],
            help="支援 OHTC 安裝排程表格式"
        )
        
        if uploaded_file:
            st.success(f"✅ {uploaded_file.name}")
        
        st.divider()
        
        st.header("⚙️ 顯示設定")
        show_actual = st.checkbox("顯示實際進度", value=False)
        show_completed = st.checkbox("顯示已完成項目", value=True)
        
        st.divider()
        
        st.header("📅 報表設定")
        report_date = st.date_input("報表日期", datetime.now())
    
    if uploaded_file is None:
        # 歡迎頁面
        st.info("👆 請先上傳專案排程表 Excel 檔案")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            ### 🆕 v2.0 新功能
            
            - 📈 **進度趨勢分析** - 視覺化專案進展
            - ⚠️ **風險評估矩陣** - 識別高風險任務
            - 📋 **自動週報生成** - 一鍵產生報表
            - 📊 **區域進度追蹤** - 系統時程視覺化
            - ✏️ **任務編輯功能** - 直接更新狀態
            - 💾 **完整 Excel 匯出** - 保持原格式
            """)
        
        with col2:
            st.markdown("""
            ### 📋 支援格式
            
            - ✅ 軟體時程表（甘特圖）
            - ✅ 系統時程表（區域進度）
            - ✅ 工程進度確認表
            - ✅ EQ 工作清單
            
            ### 🚀 快速開始
            
            1. 上傳 Excel 排程表
            2. 瀏覽各項分析圖表
            3. 追蹤延遲項目
            4. 生成並下載報表
            """)
        return
    
    # 載入資料
    data = load_excel_data(uploaded_file)
    if data is None:
        return
    
    project_info = data['project_info']
    df_tasks = data['tasks']
    df_system = data['system_tasks']
    
    # 專案資訊卡
    st.markdown("### 📌 專案資訊")
    cols = st.columns(5)
    with cols[0]:
        st.metric("📋 專案工令", project_info['project_code'])
    with cols[1]:
        st.metric("🏭 專案名稱", project_info['project_name'][:15] + "...")
    with cols[2]:
        st.metric("👤 專案負責", project_info['project_lead'])
    with cols[3]:
        if project_info['start_date']:
            st.metric("📅 開始日期", pd.to_datetime(project_info['start_date']).strftime('%Y-%m-%d'))
    with cols[4]:
        total = len(df_tasks)
        done = len(df_tasks[df_tasks['status'] == 'Done'])
        st.metric("📊 完成率", f"{done/total*100:.1f}%", f"{done}/{total}")
    
    st.divider()
    
    # 關鍵指標卡
    total = len(df_tasks)
    done = len(df_tasks[df_tasks['status'] == 'Done'])
    going = len(df_tasks[df_tasks['status'] == 'Going'])
    delay = len(df_tasks[df_tasks['status'] == 'Delay'])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #28a745, #20c997); padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <div style="font-size: 2.5rem; font-weight: bold;">{done}</div>
            <div>✅ 已完成</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #ffc107, #fd7e14); padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <div style="font-size: 2.5rem; font-weight: bold;">{going}</div>
            <div>🔄 進行中</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #dc3545, #c82333); padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <div style="font-size: 2.5rem; font-weight: bold;">{delay}</div>
            <div>⚠️ 延遲中</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #6c757d, #495057); padding: 20px; border-radius: 10px; color: white; text-align: center;">
            <div style="font-size: 2.5rem; font-weight: bold;">{total}</div>
            <div>📝 總任務數</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # 主要標籤頁
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📅 甘特圖",
        "📊 統計分析", 
        "⚠️ 風險追蹤",
        "🏭 區域進度",
        "📋 任務管理",
        "📝 週報生成",
        "⬇️ 匯出"
    ])
    
    # Tab 1: 甘特圖
    with tab1:
        st.subheader("📅 專案甘特圖")
        
        gantt_fig = create_gantt_chart(df_tasks, show_actual)
        if gantt_fig:
            st.plotly_chart(gantt_fig, use_container_width=True)
        else:
            st.warning("資料不足，無法生成甘特圖")
    
    # Tab 2: 統計分析
    with tab2:
        col1, col2 = st.columns(2)

        with col1:
            status_fig = create_status_pie(df_tasks)
            if status_fig:
                st.plotly_chart(status_fig, use_container_width=True)
            else:
                st.warning("資料不足，無法生成狀態圓餅圖")

        with col2:
            owner_fig = create_owner_workload(df_tasks)
            if owner_fig:
                st.plotly_chart(owner_fig, use_container_width=True)
            else:
                st.warning("資料不足，無法生成負責單位工作量圖")

        st.divider()

        # 進度趨勢
        trend_fig = create_progress_trend(df_tasks)
        if trend_fig:
            st.plotly_chart(trend_fig, use_container_width=True)
        else:
            st.warning("資料不足，無法生成進度趨勢圖")
    
    # Tab 3: 風險追蹤
    with tab3:
        st.subheader("⚠️ 風險評估與追蹤")
        
        delay_df = df_tasks[df_tasks['status'] == 'Delay']
        
        if delay_df.empty:
            st.success("🎉 太棒了！目前沒有延遲項目！")
        else:
            st.error(f"⚠️ 共有 {len(delay_df)} 個延遲項目需要關注")
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                risk_fig = create_risk_matrix(df_tasks)
                if risk_fig:
                    st.plotly_chart(risk_fig, use_container_width=True)
            
            with col2:
                st.markdown("### 🔴 高風險項目")
                high_risk = delay_df[delay_df['variance_days'].abs() > 7]
                for _, task in high_risk.iterrows():
                    with st.expander(f"🔴 {task['task'][:30]}..."):
                        st.write(f"**負責單位:** {task['owner']}")
                        st.write(f"**誤差天數:** {task['variance_days']} 天")
                        if pd.notna(task['plan_end']):
                            st.write(f"**計劃完成:** {task['plan_end'].strftime('%Y-%m-%d')}")
            
            st.divider()
            
            # 延遲項目清單
            st.markdown("### 📋 完整延遲項目清單")
            st.dataframe(
                delay_df[['task', 'owner', 'plan_end', 'variance_days', 'notes']].rename(columns={
                    'task': '任務', 'owner': '負責單位', 'plan_end': '計劃完成',
                    'variance_days': '誤差天數', 'notes': '備註'
                }),
                use_container_width=True,
                hide_index=True,
            )
    
    # Tab 4: 區域進度
    with tab4:
        st.subheader("🏭 系統時程 - 區域進度")
        
        area_fig = create_area_progress(df_system)
        if area_fig:
            st.plotly_chart(area_fig, use_container_width=True)
        
        st.divider()
        
        # 各區域詳細進度
        areas = df_system[df_system['is_area'] == True]['item'].unique()
        
        for area in areas:
            with st.expander(f"📍 {area}"):
                area_items = df_system[(df_system['area'] == area) & (df_system['is_area'] == False)]
                if not area_items.empty:
                    for _, item in area_items.iterrows():
                        pct = item['completion_pct']
                        color = '#28a745' if pct >= 0.7 else '#ffc107' if pct >= 0.3 else '#dc3545'
                        st.markdown(f"""
                        <div style="display: flex; align-items: center; margin: 5px 0;">
                            <div style="width: 200px;">{item['item'][:30]}</div>
                            <div style="flex: 1; background: #e9ecef; border-radius: 4px; height: 20px; margin: 0 10px;">
                                <div style="width: {pct*100}%; background: {color}; height: 100%; border-radius: 4px;"></div>
                            </div>
                            <div style="width: 50px; text-align: right;">{pct*100:.0f}%</div>
                        </div>
                        """, unsafe_allow_html=True)
    
    # Tab 5: 任務管理
    with tab5:
        st.subheader("📋 任務管理與編輯")
        
        # 篩選器
        col1, col2, col3 = st.columns(3)
        with col1:
            status_filter = st.multiselect(
                "篩選狀態",
                options=['Done', 'Going', 'Delay'],
                default=['Done', 'Going', 'Delay'] if show_completed else ['Going', 'Delay']
            )
        with col2:
            owners = sorted(df_tasks['owner'].unique().tolist())
            owner_filter = st.multiselect("篩選負責單位", options=owners)
        with col3:
            search = st.text_input("🔍 搜尋任務")
        
        # 套用篩選
        filtered_df = df_tasks[df_tasks['status'].isin(status_filter)].copy()
        if owner_filter:
            filtered_df = filtered_df[filtered_df['owner'].isin(owner_filter)]
        if search:
            filtered_df = filtered_df[filtered_df['task'].str.contains(search, case=False, na=False)]
        
        # 可編輯表格
        edited_df = st.data_editor(
            filtered_df[['id', 'task', 'owner', 'status', 'plan_start', 'plan_end', 'variance_days', 'notes']].rename(columns={
                'id': 'ID', 'task': '任務', 'owner': '負責單位', 'status': '狀態',
                'plan_start': '計劃開始', 'plan_end': '計劃完成', 'variance_days': '誤差天數', 'notes': '備註'
            }),
            column_config={
                "狀態": st.column_config.SelectboxColumn(options=["Done", "Going", "Delay"]),
                "計劃開始": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "計劃完成": st.column_config.DateColumn(format="YYYY-MM-DD"),
            },
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
        )
        
        st.caption(f"顯示 {len(filtered_df)} / {len(df_tasks)} 筆資料")
        
        # 儲存變更提示
        if st.button("💾 套用變更", type="primary"):
            st.success("✅ 變更已記錄，請至「匯出」頁面下載更新後的 Excel")
            st.session_state['edited_tasks'] = edited_df
    
    # Tab 6: 週報生成
    with tab6:
        st.subheader("📝 專案週報生成")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            report_content = generate_weekly_report(data, datetime.combine(report_date, datetime.min.time()))
            st.markdown(report_content)
        
        with col2:
            st.markdown("### 📥 下載報表")
            
            st.download_button(
                label="📄 下載 Markdown",
                data=report_content,
                file_name=f"週報_{report_date.strftime('%Y%m%d')}.md",
                mime="text/markdown",
            )
            
            st.divider()
            
            st.markdown("### 📊 快速統計")
            summary = generate_status_summary(data)
            
            st.metric("完成率", f"{summary['done']/summary['total']*100:.1f}%")
            st.metric("延遲項目", summary['delay'])
            st.metric("本週到期", len(summary['upcoming']))
    
    # Tab 7: 匯出
    with tab7:
        st.subheader("⬇️ 匯出資料")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("### 📊 Excel 完整匯出")
            st.write("保持原始格式，匯出更新後的排程表")
            
            if st.button("🔄 生成 Excel", type="primary"):
                try:
                    # 使用編輯過的資料（如果有）
                    tasks_to_export = st.session_state.get('edited_tasks', df_tasks)
                    excel_output = export_updated_excel(data, uploaded_file, df_tasks)
                    
                    st.download_button(
                        label="⬇️ 下載 Excel",
                        data=excel_output,
                        file_name=f"OHTC_排程表_更新_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Excel 已生成")
                except Exception as e:
                    st.error(f"匯出失敗: {str(e)}")
        
        with col2:
            st.markdown("### 📋 CSV 匯出")
            st.write("任務清單輕量匯出")
            
            csv = df_tasks.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="⬇️ 下載 CSV",
                data=csv,
                file_name=f"任務清單_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        with col3:
            st.markdown("### 📈 JSON 匯出")
            st.write("結構化資料匯出，適合程式處理")
            
            json_data = {
                'project_info': project_info,
                'summary': generate_status_summary(data),
                'exported_at': datetime.now().isoformat(),
            }
            
            st.download_button(
                label="⬇️ 下載 JSON",
                data=json.dumps(json_data, ensure_ascii=False, indent=2, default=str),
                file_name=f"專案摘要_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json"
            )


if __name__ == "__main__":
    main()
