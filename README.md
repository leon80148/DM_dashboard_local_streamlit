# 糖尿病品質管理 Dashboard
目前僅提供展望門診系統使用，用於展示和分析糖尿病患者的相關數據。

## 功能特點

- HbA1c 分布統計
- 月度收案統計
- 糖尿病統計指標
- 病患資料列表（可搜尋和分頁）
- 資料匯出功能

## 系統需求

- Python 3.9 或更新版本
- Windows 作業系統（建議 Windows 10 或更新版本）

## 安裝和執行方式

### 方式一：使用批次檔案（推薦）

1. **執行批次檔案**
   - 直接雙擊執行 `run_dashboard.bat`
   - 批次檔案會自動：
     - 檢查 Python 安裝狀態
     - 安裝必要套件
     - 啟動應用程式

### 方式二：手動執行

1. **安裝 Python**
   - 從 [Python 官網](https://www.python.org/downloads/) 下載並安裝 Python 3.9 或更新版本
   - 安裝時請勾選「Add Python to PATH」選項

2. **安裝必要套件**
   - 開啟命令提示字元（CMD）
   - 切換到程式目錄：
     ```bash
     cd C:\DM_quality\streamlit
     ```
   - 安裝必要套件：
     ```bash
     pip install -r requirements.txt
     ```

3. **啟動應用程式**
   - 在命令提示字元中執行：
     ```bash
     streamlit run app.py
     ```
   - 程式會自動在預設瀏覽器中開啟（通常是 http://localhost:8501）

## 目錄結構

```
streamlit/
├── app.py              # 主程式
├── data_processor.py   # 資料處理模組
├── database_utils.py   # 資料庫工具模組
├── requirements.txt    # 套件需求檔案
├── settings.ini        # 設定檔
├── run_dashboard.bat   # 啟動批次檔
└── README.md          # 說明文件
```

## 技術支援
由Cursor/claude-3.5-sonnet提供技術支援
版本更新日期：2024-12-08