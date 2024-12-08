import os
import shutil
from dbfread import DBF
import sqlite3
import pandas as pd
import logging
from datetime import datetime
import pytz
import configparser

# 設置時區
taipei_tz = pytz.timezone('Asia/Taipei')

def ensure_directories():
    """確保所有必要的目錄都存在"""
    config = configparser.ConfigParser()
    config.read('settings.ini')
    
    # 設定預設路徑
    default_paths = {
        'source_folder': './source_file',
        'dbf_folder': './dbf_data',
        'database_path': './database/database.db'
    }
    
    # 如果設定檔中沒有 Paths 區段，則創建它
    if 'Paths' not in config:
        config['Paths'] = {}
    
    # 檢查並設定預設值
    for key, default_value in default_paths.items():
        if not config.get('Paths', key, fallback=''):
            config.set('Paths', key, default_value)
    
    # 保存更新後的設定
    with open('settings.ini', 'w') as f:
        config.write(f)
    
    # 創建必要的目錄
    paths = [
        os.path.dirname(config.get('Paths', 'database_path')),
        config.get('Paths', 'dbf_folder'),
        config.get('Paths', 'source_folder')
    ]
    
    for path in paths:
        if path:  # 只在路徑非空時創建目錄
            os.makedirs(path, exist_ok=True)
            print(f"確保目錄存在: {path}")

def copy_required_files(source_folder, dest_folder):
    """複製必要的 DBF 檔案到指定目錄，不修改原始檔案"""
    required_files = ['CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF']
    
    try:
        # 確保目標目錄存在
        os.makedirs(dest_folder, exist_ok=True)
        
        if not os.path.exists(source_folder):
            return False, f"錯誤：源目錄 {source_folder} 不存在"
        
        # 獲取資料夾中的所有檔案並轉換為大寫
        existing_files = {f.upper(): f for f in os.listdir(source_folder)}
        
        # 檢查必要的檔案
        missing_files = []
        for file in required_files:
            if file.upper() not in existing_files:
                missing_files.append(file)
        
        if missing_files:
            return False, f"錯誤：在源目錄中缺少以下檔案：{missing_files}"
        
        # 複製檔案到工作目錄（使用實際的檔案名稱）
        for file in required_files:
            source_path = os.path.join(source_folder, existing_files[file.upper()])
            dest_path = os.path.join(dest_folder, file)
            try:
                # 使用 copy2 保留檔案的元數據，但不修改原始檔案
                shutil.copy2(source_path, dest_path)
                print(f"成功複製 {file} 到工作目錄")
            except Exception as e:
                return False, f"複製 {file} 時發生錯誤：{str(e)}"
        
        return True, "檔案複製成功"
    except Exception as e:
        return False, f"複製檔案時發生錯誤：{str(e)}"

def convert_dbf_to_sqlite(dbf_folder, sqlite_db_path):
    """轉換工作目錄中的 DBF 檔案到 SQLite 資料庫，不接觸原始檔案"""
    try:
        os.makedirs(os.path.dirname(sqlite_db_path), exist_ok=True)
        
        required_files = ['CO01M.DBF', 'CO02M.DBF', 'CO03M.DBF', 'CO18H.DBF']
        
        # 檢查工作目錄中的檔案
        for file in required_files:
            if not os.path.exists(os.path.join(dbf_folder, file)):
                return False, f"工作目錄中缺少檔案：{file}"
        
        file_encodings = {
            'CO01M.DBF': 'latin-1',
            'CO02M.DBF': 'cp950',
            'CO03M.DBF': 'cp950',
            'CO18H.DBF': 'cp950'
        }
        
        conn = sqlite3.connect(sqlite_db_path)
        
        for dbf_file in required_files:
            try:
                dbf_path = os.path.join(dbf_folder, dbf_file)
                table_name = os.path.splitext(dbf_file)[0]
                
                encoding = file_encodings.get(dbf_file.upper(), 'cp950')
                
                dbf = DBF(dbf_path, encoding=encoding, ignore_missing_memofile=True)
                df = pd.DataFrame(list(dbf))
                
                if dbf_file.upper() == 'CO01M.DBF' and 'MNAME' in df.columns:
                    df['MNAME'] = df['MNAME'].apply(
                        lambda x: x.encode('latin-1').decode('big5', errors='ignore') 
                        if isinstance(x, str) else x
                    )
                
                df.to_sql(table_name, conn, if_exists='replace', index=False)
                
            except Exception as e:
                print(f"處理 {dbf_file} 時發生錯誤: {str(e)}")
                continue
        
        conn.close()
        return True, "資料庫轉換成功"
    except Exception as e:
        return False, f"資料庫轉換失敗：{str(e)}"

def update_database(source_folder, dbf_folder, sqlite_db_path):
    """更新資料庫的主要函數"""
    success, message = copy_required_files(source_folder, dbf_folder)
    if not success:
        return False, message
        
    success, message = convert_dbf_to_sqlite(dbf_folder, sqlite_db_path)
    if not success:
        return False, message
        
    return True, "資料庫更新成功" 