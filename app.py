import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import configparser
import os
from database_utils import update_database, ensure_directories
from data_processor import DataProcessor
from io import BytesIO
import subprocess

# 設定頁面（必須是第一個 Streamlit 命令）
st.set_page_config(
    page_title="糖尿病品質管理 Dashboard",
    page_icon="🏥",
    layout="wide"
)

# 確保目錄存在
ensure_directories()

# 讀取設定檔
config = configparser.ConfigParser()
config.read('settings.ini')

# 檢查資料庫是否存在
database_path = config.get('Paths', 'database_path')
if not os.path.exists(database_path):
    st.title("糖尿病品質管理 Dashboard")
    st.warning("尚無資料，請選擇資料來源")
    
    # 側邊欄
    with st.sidebar:
        st.title("控制面板")
        
        # 檔案選擇功能
        st.subheader("資料來源設定")
        st.info("請選擇選擇展望資料夾 \n\n (內有以下檔案：'CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF')。 \n\n 注意：程式會將來源資料夾的檔案複製到工作目錄進行處理，不會修改原始檔案。")
        
        # 顯示當前路徑
        current_path = config.get('Paths', 'source_folder')
        st.text_input("目前資料來源路徑", value=current_path, disabled=True)
        
        # 選擇資料夾按鈕
        if st.button("選擇資料來源"):
            try:
                folder_path = None
                if os.name == 'nt':  # Windows
                    # 修改 PowerShell 命令
                    cmd = '''
                    Add-Type -AssemblyName System.Windows.Forms
                    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
                    $dialog.Description = "選擇包含 DBF 檔案的資料夾"
                    $dialog.ShowDialog() | Out-Null
                    $dialog.SelectedPath
                    '''
                    result = subprocess.run(['powershell', '-command', cmd], capture_output=True, text=True)
                    folder_path = result.stdout.strip()
                else:  # Linux/Mac
                    result = subprocess.run(['zenity', '--file-selection', '--directory'], capture_output=True, text=True)
                    folder_path = result.stdout.strip()

                # 檢查是否選擇了有效的資料夾
                if not folder_path or not os.path.exists(folder_path):
                    st.error("未選擇有效的資料夾")
                else:
                    # 檢查必要的檔案是否存在
                    required_files = ['CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF']
                    
                    # 獲取資料夾中的所有檔案並轉換為大寫
                    existing_files = [f.upper() for f in os.listdir(folder_path)]
                    
                    # 檢查每個必要的檔案（不區分大小寫）
                    missing_files = [file for file in required_files if file.upper() not in existing_files]
                    
                    if missing_files:
                        st.error(f"選擇的資料夾缺少以下檔案：\n{', '.join(missing_files)}")
                    else:
                        # 更新設定檔
                        config.set('Paths', 'source_folder', folder_path)
                        with open('settings.ini', 'w') as f:
                            config.write(f)
                        st.success(f"已設定資料來源：{folder_path}")
                        
                        # 自動觸發更新
                        with st.spinner("更新資料中..."):
                            success, message = update_database(
                                folder_path,
                                config.get('Paths', 'dbf_folder'),
                                config.get('Paths', 'database_path')
                            )
                            if success:
                                st.success("資料更新成功！")
                                st.rerun()
                            else:
                                st.error(f"更新失敗：{message}")
                        
            except Exception as e:
                st.error(f"選擇資料夾時發生錯誤：{str(e)}")
    
    # 如果沒有資料庫，不顯示其他內容
    st.stop()

# 初始化資料處理器
data_processor = DataProcessor(database_path)

# 側邊欄
with st.sidebar:
    st.title("控制面板")
    
    # 檔案選擇功能
    st.subheader("資料來源設定")
    st.info("請選擇選擇展望資料夾 \n\n (內有以下檔案：'CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF')。 \n\n 注意：程式會將來源資料夾的檔案複製到工作目錄進行處理，不會修改原始檔案。")
    
    # 顯示當前路徑
    current_path = config.get('Paths', 'source_folder')
    st.text_input("目前資料來源路徑", value=current_path, disabled=True)
    
    # 選擇資料夾按鈕
    if st.button("選擇資料來源"):
        try:
            folder_path = None
            if os.name == 'nt':  # Windows
                # 修改 PowerShell 命令
                cmd = '''
                Add-Type -AssemblyName System.Windows.Forms
                $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $dialog.Description = "選擇包含 DBF 檔案的資料夾"
                $dialog.ShowDialog() | Out-Null
                $dialog.SelectedPath
                '''
                result = subprocess.run(['powershell', '-command', cmd], capture_output=True, text=True)
                folder_path = result.stdout.strip()
            else:  # Linux/Mac
                result = subprocess.run(['zenity', '--file-selection', '--directory'], capture_output=True, text=True)
                folder_path = result.stdout.strip()

            # 檢查是否選擇了有效的資料夾
            if not folder_path or not os.path.exists(folder_path):
                st.error("未選擇有效的資料夾")
            else:
                # 檢查必要的檔案是否存在
                required_files = ['CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF']
                
                # 獲取資料夾中的所有檔案並轉換為大寫
                existing_files = [f.upper() for f in os.listdir(folder_path)]
                
                # 檢查每個必要的檔案（不區分大小寫）
                missing_files = [file for file in required_files if file.upper() not in existing_files]
                
                if missing_files:
                    st.error(f"選擇的資料夾缺少以下檔案：\n{', '.join(missing_files)}")
                else:
                    # 更新設定檔
                    config.set('Paths', 'source_folder', folder_path)
                    with open('settings.ini', 'w') as f:
                        config.write(f)
                    st.success(f"已設定資料來源：{folder_path}")
                    
                    # 自動觸發更新
                    with st.spinner("更新資料中..."):
                        success, message = update_database(
                            folder_path,
                            config.get('Paths', 'dbf_folder'),
                            config.get('Paths', 'database_path')
                        )
                        if success:
                            st.success("資料更新成功！")
                            if 'full_data' in st.session_state:
                                del st.session_state['full_data']
                            st.rerun()
                        else:
                            st.error(f"更新失敗：{message}")
                    
        except Exception as e:
            st.error(f"選擇資料夾時發生錯���：{str(e)}")

# HbA1c 分布
col1, col2 = st.columns(2)

with col1:
    st.subheader("HbA1c 分布")
    hba1c_data = data_processor.get_hba1c_distribution()
    fig = px.pie(
        hba1c_data,
        values='count',
        names='range',
        title="HbA1c 分布"
    )
    st.plotly_chart(fig)

# 月度統計
with col2:
    st.subheader("月度收案統計")
    monthly_data = data_processor.get_monthly_statistics()
    fig = px.bar(
        monthly_data,
        x='month',
        y='diabetes_cases',
        title="月度收案人數"
    )
    st.plotly_chart(fig)

# 糖尿病統計
st.subheader("糖尿病統計")
diabetes_stats = data_processor.get_diabetes_statistics()

# 創建統計卡片的列表
stats_cards = [
    {"title": "糖尿病人收案率", "data": diabetes_stats.get('diabetes_case_rate')},
    {"title": "眼底檢查率-全部", "data": diabetes_stats.get('diabetes_eye_exam_rate')},
    {"title": "眼底檢查率-收案", "data": diabetes_stats.get('cased_eye_exam_rate')},
    {"title": "血糖控制率（A1c＜ 7%）-全部", "data": diabetes_stats.get('diabetes_dm_control_rate')},
    {"title": "血糖控制率（A1c＜ 7%）-收案", "data": diabetes_stats.get('diabetes_dm_control_rate_cased')},
    {"title": "血壓控制率（BP＜130/80）-全部", "data": diabetes_stats.get('diabetes_bp_control_rate')},
    {"title": "血壓控制率（BP＜130/80）-收案", "data": diabetes_stats.get('diabetes_bp_control_rate_cased')},
    {"title": "血脂控制率（LDL＜100）-全部", "data": diabetes_stats.get('diabetes_ldl_control_rate')},
    {"title": "血脂控制率（LDL＜100）-收案", "data": diabetes_stats.get('diabetes_ldl_control_rate_cased')},
    {"title": "ABC達標率（A1c＜ 7%且BP＜140/90且LDL＜100）-全部", "data": diabetes_stats.get('diabetes_all_control_rate')},
    {"title": "ABC達標率（A1c＜ 7%且BP＜140/90且LDL＜100）-收案", "data": diabetes_stats.get('diabetes_all_control_rate_cased')}
]

# 使用 columns 創建網格布局
cols = st.columns(3)
for i, stat in enumerate(stats_cards):
    with cols[i % 3]:
        st.markdown(
            f"""
            <div style="
                padding: 1rem;
                border-radius: 0.5rem;
                border: 1px solid #ddd;
                margin-bottom: 1rem;
                background-color: white;
            ">
                <h3 style="
                    font-size: 1rem;
                    margin-bottom: 0.5rem;
                    color: #333;
                ">{stat['title']}</h3>
                <div style="
                    font-size: 1.5rem;
                    color: #2196F3;
                    font-weight: bold;
                ">{stat['data']['percentage']}%</div>
                <div style="
                    font-size: 0.9rem;
                    color: #666;
                ">({stat['data']['numerator']} / {stat['data']['denominator']})</div>
            </div>
            """,
            unsafe_allow_html=True
        )

# 病患資料表格
st.subheader("病患資料列表")

# 如果 session state 中沒有資料，則獲取資料
if 'full_data' not in st.session_state:
    st.session_state.full_data = data_processor.get_patient_list()

# 搜尋功能
search_term = st.text_input("搜尋病患 (病歷號或姓名)")
if search_term:
    df = st.session_state.full_data[
        (st.session_state.full_data['病歷號'].astype(str).str.contains(search_term)) | 
        (st.session_state.full_data['姓名'].str.contains(search_term))
    ]
else:
    df = st.session_state.full_data

# 分頁功能
col1, col2, col3 = st.columns([1, 2, 1])
with col1:
    rows_per_page = st.selectbox(
        "每頁顯示筆數", 
        options=[50, 100, 500],
        index=0  # 預設選擇第一個選項 (50筆)
    )

# 計算總頁數
total_records = len(df)
total_pages = (total_records + rows_per_page - 1) // rows_per_page

# 初始化或重置頁碼
if 'current_page' not in st.session_state:
    st.session_state.current_page = 1

# 頁面切換函數
def change_page(new_page):
    new_page = max(1, min(new_page, total_pages))
    if new_page != st.session_state.current_page:
        st.session_state.current_page = new_page
        st.rerun()

with col2:
    # 分頁導航按鈕
    cols = st.columns(5)
    
    # 首頁按
    if cols[0].button("⏮️", key='first_page'):
        change_page(1)
    
    # 上一頁按鈕
    if cols[1].button("◀️", key='prev_page'):
        change_page(st.session_state.current_page - 1)
    
    # 當前頁碼/總頁數
    page_text = f"第 {st.session_state.current_page} / {total_pages} 頁"
    cols[2].markdown(f"<div style='text-align: center'>{page_text}</div>", unsafe_allow_html=True)
    
    # 下一頁按鈕
    if cols[3].button("▶️", key='next_page'):
        change_page(st.session_state.current_page + 1)
    
    # 末頁按鈕
    if cols[4].button("⏭️", key='last_page'):
        change_page(total_pages)

with col3:
    st.write(f"共 {total_records} 筆資料")

# 計算當前頁的資料範圍
start_idx = (st.session_state.current_page - 1) * rows_per_page
end_idx = min(start_idx + rows_per_page, total_records)

# 顯示表格
st.dataframe(df.iloc[start_idx:end_idx], height=400)

# 匯出功能
if st.button("匯出全部病患資料資料"):
    with st.spinner("準備匯出資料..."):
        export_df = data_processor.export_patient_data()
        
        # 創建 BytesIO 對象
        output = BytesIO()
        
        # 將數據寫入 Excel
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False, sheet_name='病患資料')
            
            # 獲取 xlsxwriter 工作簿和工作表對象
            workbook = writer.book
            worksheet = writer.sheets['病患資料']
            
            # 設置標題格式
            header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': '#D9E1F2',
                'border': 1
            })
            
            # 設置數據格式
            data_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'border': 1
            })
            
            # 設置列寬和格式
            for idx, col in enumerate(export_df.columns):
                column_width = max(
                    len(str(col)) * 2,
                    export_df[col].astype(str).str.len().max() * 2
                )
                worksheet.set_column(idx, idx, column_width, data_format)
                worksheet.write(0, idx, col, header_format)
 
        output.seek(0)

        # 下載按鈕
        st.download_button(
            label="點擊下載 Excel 檔案",
            data=output.getvalue(),
            file_name=f"diabetes_patient_data_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )