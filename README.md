# 🏭 OHTC 專案管理儀表板 v2.0

一個專為 OHTC/AMHS 安裝排程設計的專案管理視覺化工具。

## 🔥 快速部署（30 秒上線）

**不需要安裝 Python！完全在雲端執行！**

詳細部署步驟請參閱 [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md)

**推薦方案：Streamlit Community Cloud（免費）**
1. 上傳程式碼到 GitHub
2. 連結 https://share.streamlit.io
3. 一鍵部署，獲得永久網址
4. 團隊成員直接用瀏覽器存取

---

## 📦 包含工具

| 工具 | 說明 | 使用方式 |
|------|------|----------|
| `app.py` | 基礎版儀表板 | `streamlit run app.py` |
| `app_v2.py` | 增強版儀表板 (推薦) | `streamlit run app_v2.py` |
| `cli.py` | 命令列工具 | `python cli.py status` |
| `notifications.py` | 通知系統 | 整合 Teams/Slack/Email |
| `template_generator.py` | 模板生成器 | `python template_generator.py -n "專案名"` |

## ✨ 功能特色

### 📊 視覺化呈現
- **甘特圖** - 直觀顯示專案時程，自動標示今日進度線
- **狀態圓餅圖** - 一眼看出 Done/Going/Delay 比例
- **負責單位工作量** - 堆疊長條圖顯示各單位任務分配
- **區域進度** - 系統時程各區域完成率
- **進度趨勢圖** - 追蹤專案進展變化
- **風險評估矩陣** - 識別高風險延遲任務

### ⚠️ 智慧追蹤
- **延遲項目警示** - 紅色醒目標示所有 Delay 狀態任務
- **即將到期提醒** - 7 天內到期的任務自動提醒
- **誤差天數分析** - 計算實際與計劃的差異

### 📋 任務管理
- **多條件篩選** - 依狀態、負責單位、關鍵字搜尋
- **即時編輯** - 直接在表格中修改狀態
- **排序功能** - 點擊欄位標題排序

### ⬇️ 匯出功能
- **Excel 匯出** - 保持原始格式，更新後匯出
- **CSV 匯出** - 任務清單輕量匯出

---

## 🚀 快速開始

### 方式一：本機安裝

```bash
# 1. 安裝相依套件
pip install -r requirements.txt

# 2. 啟動應用程式
streamlit run app.py

# 3. 瀏覽器自動開啟 http://localhost:8501
```

### 方式二：Docker 部署 (推薦團隊使用)

```bash
# 建立並啟動容器
docker build -t ohtc-dashboard .
docker run -p 8501:8501 ohtc-dashboard
```

---

## 📁 支援的 Excel 格式

此工具專為以下格式的排程表設計：

### 軟體時程表 (工作表名稱: `軟體時程`)
| 欄位 | 說明 |
|------|------|
| 項目 | 任務名稱 |
| 負責單位 | 負責團隊或人員 |
| 實際完成進度 | 完成百分比 |
| 進度 | 狀態 (Done/Going/Delay) |
| 計劃開始日期 | 預計開始日期 |
| 計劃完成日期 | 預計完成日期 |
| 實際開始日期 | 實際開始日期 |
| 實際完成日期 | 實際完成日期 |
| 誤差天數 | 實際與計劃差異 |

### 系統時程表 (工作表名稱: `系統時程_C`)
包含各區域 (A-H) 的系統驗證進度。

---

## 🎯 使用流程

1. **上傳檔案** - 在左側邊欄上傳 Excel 排程表
2. **瀏覽總覽** - 查看關鍵指標和整體完成率
3. **分析圖表** - 透過甘特圖和統計圖表了解進度
4. **追蹤延遲** - 關注需要處理的延遲項目
5. **匯出報表** - 下載更新後的 Excel 或 CSV

---

## 🔧 客製化

### 修改狀態顏色
在 `app.py` 中修改 `color_map`:
```python
color_map = {
    'Done': '#28a745',   # 綠色
    'Going': '#ffc107',  # 黃色
    'Delay': '#dc3545',  # 紅色
}
```

### 調整甘特圖高度
修改 `create_gantt_chart` 函數中的 height 計算邏輯。

### 新增欄位
在 `load_excel_data` 函數中擴充 task 字典。

---

## 📝 更新日誌

### v2.0.1 (2025-11-28)
- 🐛 修正 Excel 讀取錯誤（`could not convert string to float` 問題）
- ✨ 增強資料驗證，自動跳過標題行
- 🛡️ 新增安全的類型轉換函數，避免程式崩潰
- 📚 新增完整部署指南（DEPLOYMENT_GUIDE.md）
- 🚀 支援 Streamlit Cloud 一鍵部署

### v1.0.0 (2025-11-26)
- 初始版本
- 支援軟體時程和系統時程解析
- 甘特圖、狀態圖、負責單位圖表
- 延遲追蹤和到期提醒
- Excel/CSV 匯出功能

---

## 🤝 團隊部署建議

### 內網部署
```bash
# 允許區域網路存取
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

### 雲端部署
- **Streamlit Cloud** - 免費，適合小團隊
- **Azure App Service** - 企業級，整合 AD 認證
- **AWS ECS** - 彈性擴展

---

## 📞 支援

如有問題或功能建議，請聯繫 TIM SMA 軟體整合團隊。
