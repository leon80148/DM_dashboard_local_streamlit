import pandas as pd
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text
import pytz
import streamlit as st

taipei_tz = pytz.timezone('Asia/Taipei')

class DataProcessor:
    def __init__(self, database_path):
        self.engine = create_engine(f'sqlite:///{database_path}')
        
    def get_hba1c_distribution(self):
        try:
            current_date = datetime.now()
            one_year_ago = current_date - timedelta(days=365)
            one_year_ago_roc = f"{one_year_ago.year - 1911:03d}{one_year_ago.month:02d}{one_year_ago.day:02d}"
            
            query = text("""
                SELECT 
                    CASE 
                        WHEN CAST(REPLACE(CO18H.hval, ' ', '') AS DECIMAL) < 7 THEN 'HbA1c < 7%'
                        WHEN CAST(REPLACE(CO18H.hval, ' ', '') AS DECIMAL) >= 8.5 THEN 'HbA1c ≥ 8.5%'
                        ELSE '7% ≤ HbA1c < 8.5%'
                    END AS hba1c_range,
                    COUNT(DISTINCT CO18H.kcstmr) AS count
                FROM CO18H
                INNER JOIN CO03M ON CO18H.kcstmr = CO03M.kcstmr
                WHERE CO18H.hitem = 'Z0SHbA1c'
                  AND CO18H.hdate >= :one_year_ago
                  AND (substr(CO03M.labno, 1, 3) BETWEEN 'E08' AND 'E13'
                    OR substr(CO03M.lacd01, 1, 3) BETWEEN 'E08' AND 'E13'
                    OR substr(CO03M.lacd02, 1, 3) BETWEEN 'E08' AND 'E13')
                GROUP BY 
                    CASE 
                        WHEN CAST(REPLACE(CO18H.hval, ' ', '') AS DECIMAL) < 7 THEN 'HbA1c < 7%'
                        WHEN CAST(REPLACE(CO18H.hval, ' ', '') AS DECIMAL) >= 8.5 THEN 'HbA1c ≥ 8.5%'
                        ELSE '7% ≤ HbA1c < 8.5%'
                    END
            """)
            
            with self.engine.connect() as conn:
                result = conn.execute(query, {'one_year_ago': one_year_ago_roc})
                data = [{"range": row.hba1c_range, "count": row.count} for row in result]
                
                # 添加日誌來幫助調試
                print(f"查詢參數 one_year_ago: {one_year_ago_roc}")
                print(f"查詢結果: {data}")
                
                # 如果沒有數據，返回預設值
                if not data:
                    print("警告：查詢沒有返回任何數據")
                    data = [
                        {"range": "HbA1c < 7%", "count": 0},
                        {"range": "7% ≤ HbA1c < 8.5%", "count": 0},
                        {"range": "HbA1c ≥ 8.5%", "count": 0}
                    ]
                
                return data
                
        except Exception as e:
            st.error(f"獲取 HbA1c 分布資料時發生錯誤：{str(e)}")
            print(f"詳細錯誤：{str(e)}")  # 添加詳細錯誤日誌
            return []

    def get_monthly_statistics(self):
        """獲取月度統計數據"""
        current_date = datetime.now()
        six_months_ago = current_date - timedelta(days=180)
        six_months_ago_roc = f"{six_months_ago.year - 1911:03d}{six_months_ago.month:02d}{six_months_ago.day:02d}"
        
        query = """
            WITH RECURSIVE 
            Months(month_date) AS (
                SELECT :start_date
                UNION ALL
                SELECT 
                    CASE 
                        WHEN substr(month_date, 4, 2) = '12' 
                        THEN (substr(month_date, 1, 3) + 1) || '01'
                        ELSE substr(month_date, 1, 3) || 
                             printf('%02d', CAST(substr(month_date, 4, 2) AS INTEGER) + 1)
                    END
                FROM Months
                WHERE month_date < substr(:end_date, 1, 5)
            ),
            MonthlyData AS (
                SELECT 
                    substr(idate, 1, 5) AS month,
                    COUNT(DISTINCT kcstmr) AS diabetes_cases
                FROM CO02M
                WHERE idate >= :six_months_ago
                    AND dno IN ('P1407C', 'P1408C', 'P1409C', 'P7001C', 'P7002C')
                GROUP BY substr(idate, 1, 5)
            )
            SELECT 
                m.month_date as month,
                COALESCE(md.diabetes_cases, 0) as diabetes_cases
            FROM Months m
            LEFT JOIN MonthlyData md ON m.month_date = md.month
            ORDER BY m.month_date
        """
        
        with self.engine.connect() as conn:
            result = conn.execute(text(query), {
                'start_date': six_months_ago_roc[:5],
                'end_date': f"{current_date.year - 1911:03d}{current_date.month:02d}",
                'six_months_ago': six_months_ago_roc
            })
            data = [{"month": f"{row.month[:3]}/{row.month[3:]}", 
                    "diabetes_cases": row.diabetes_cases} for row in result]
        return data

    def get_patient_list(self):
        """獲取病患列表"""
        current_year = datetime.now().year - 1911
        one_year_ago = f"{current_year - 1:03d}0101"
        
        query = """
            WITH DiabetesPatients AS (
                SELECT DISTINCT 
                    CO01M.kcstmr as '病歷號',
                    CO01M.mname as '姓名'
                FROM CO01M
                INNER JOIN CO03M ON CO01M.kcstmr = CO03M.kcstmr
                WHERE (substr(CO03M.labno, 1, 3) BETWEEN 'E08' AND 'E13'
                   OR substr(CO03M.lacd01, 1, 3) BETWEEN 'E08' AND 'E13'
                   OR substr(CO03M.lacd02, 1, 3) BETWEEN 'E08' AND 'E13')
            )
            SELECT 
                dp.'病歷號',
                dp.'姓名',
                h.latest_hba1c as '最近HbA1c值',
                h.latest_hba1c_date as 'HbA1c檢查日期',
                c.case_type as '最近收案狀態',
                c.latest_case_status_date as '收案日期',
                e.eye_exam_date as '眼底檢查日期'
            FROM DiabetesPatients dp
            LEFT JOIN (
                SELECT 
                    kcstmr,
                    hval as latest_hba1c,
                    hdate as latest_hba1c_date,
                    ROW_NUMBER() OVER (PARTITION BY kcstmr ORDER BY hdate DESC) as rn
                FROM CO18H
                WHERE hitem = 'Z0SHbA1c'
                AND hdate >= :one_year_ago
            ) h ON dp.'病歷號' = h.kcstmr AND h.rn = 1
            LEFT JOIN (
                SELECT 
                    kcstmr,
                    CASE 
                        WHEN dno = 'P1407C' THEN 'DM初診'
                        WHEN dno = 'P1408C' THEN 'DM複診'
                        WHEN dno = 'P1409C' THEN 'DM年度'
                        WHEN dno = 'P7001C' THEN 'DKD複診'
                        WHEN dno = 'P7002C' THEN 'DKD年度'
                    END as case_type,
                    idate as latest_case_status_date,
                    ROW_NUMBER() OVER (PARTITION BY kcstmr ORDER BY idate DESC) as rn
                FROM CO02M
                WHERE dno IN ('P1407C', 'P1408C', 'P1409C', 'P7001C', 'P7002C')
            ) c ON dp.'病歷號' = c.kcstmr AND c.rn = 1
            LEFT JOIN (
                SELECT 
                    kcstmr,
                    MAX(idate) as eye_exam_date
                FROM CO02M
                WHERE dno = '23502C'
                GROUP BY kcstmr
            ) e ON dp.'病歷號' = e.kcstmr
            ORDER BY dp.'病歷號'
        """
        
        with self.engine.connect() as conn:
            result = pd.read_sql(query, conn, params={"one_year_ago": one_year_ago})
        return result

    def export_patient_data(self):
        """匯出病患資料"""
        df = self.get_patient_list()
        return df  # 直接返回，不需要重新命名

    def get_diabetes_statistics(self):
        """獲取糖尿病統計數據"""
        try:
            current_date = datetime.now()
            one_year_ago = current_date - timedelta(days=365)
            one_year_ago_roc = f"{one_year_ago.year - 1911:03d}{one_year_ago.month:02d}{one_year_ago.day:02d}"
            current_year_roc = f"{current_date.year - 1911:03d}0101"

            query = text("""
                WITH DiabetesPatients AS (
                    SELECT DISTINCT kcstmr
                    FROM CO03M
                    WHERE (SUBSTR(labno, 1, 3) BETWEEN 'E08' AND 'E13'
                        OR SUBSTR(lacd01, 1, 3) BETWEEN 'E08' AND 'E13'
                        OR SUBSTR(lacd02, 1, 3) BETWEEN 'E08' AND 'E13')
                    AND idate >= :one_year_ago
                ),
                CasedPatients AS (
                    SELECT DISTINCT kcstmr
                    FROM CO02M
                    WHERE dno IN ('P1407C', 'P1408C', 'P1409C', 'P7001C', 'P7002C')
                    AND idate >= :one_year_ago
                ),
                EyeExamPatients AS (
                    SELECT DISTINCT kcstmr
                    FROM CO02M
                    WHERE dno = '23502C'
                    AND idate >= :current_year
                ),
                BP_Controlled AS (
                    SELECT DISTINCT kcstmr
                    FROM CO18H
                    WHERE hitem = 'BP'
                    AND CAST(SUBSTR(hval, 1, INSTR(hval, '/') - 1) AS INTEGER) < 130
                    AND CAST(SUBSTR(hval, INSTR(hval, '/') + 1) AS INTEGER) < 80
                    AND hdate >= :current_year
                ),
                LDL_Controlled AS (
                    SELECT DISTINCT kcstmr
                    FROM CO18H
                    WHERE hitem = 'Z0SLDL'
                    AND CAST(hval AS FLOAT) < 100
                    AND hdate >= :current_year
                ),
                DM_Controlled AS (
                    SELECT DISTINCT kcstmr
                    FROM CO18H
                    WHERE hitem = 'Z0SHbA1c'
                    AND CAST(hval AS FLOAT) < 7
                    AND hdate >= :current_year
                ),
                All_Controlled AS (
                    SELECT DISTINCT d.kcstmr
                    FROM DiabetesPatients d
                    JOIN CO18H h1 ON d.kcstmr = h1.kcstmr
                    JOIN CO18H h2 ON d.kcstmr = h2.kcstmr
                    JOIN CO18H h3 ON d.kcstmr = h3.kcstmr
                    WHERE h1.hitem = 'Z0SHbA1c' 
                    AND CAST(h1.hval AS FLOAT) < 7
                    AND h1.hdate >= :current_year
                    AND h2.hitem = 'BP'
                    AND CAST(SUBSTR(h2.hval, 1, INSTR(h2.hval, '/') - 1) AS INTEGER) < 140
                    AND CAST(SUBSTR(h2.hval, INSTR(h2.hval, '/') + 1) AS INTEGER) < 90
                    AND h2.hdate >= :current_year
                    AND h3.hitem = 'Z0SLDL'
                    AND CAST(h3.hval AS FLOAT) < 100
                    AND h3.hdate >= :current_year
                )
                SELECT
                    (SELECT COUNT(*) FROM DiabetesPatients) AS total_diabetes_patients,
                    (SELECT COUNT(*) FROM CasedPatients) AS total_cased_patients,
                    (SELECT COUNT(*) FROM DiabetesPatients WHERE kcstmr IN (SELECT kcstmr FROM EyeExamPatients)) AS diabetes_eye_exam_patients,
                    (SELECT COUNT(*) FROM CasedPatients WHERE kcstmr IN (SELECT kcstmr FROM EyeExamPatients)) AS cased_eye_exam_patients,
                    (SELECT COUNT(*) FROM DiabetesPatients WHERE kcstmr IN (SELECT kcstmr FROM BP_Controlled)) AS bp_controlled_patients,
                    (SELECT COUNT(*) FROM DiabetesPatients WHERE kcstmr IN (SELECT kcstmr FROM LDL_Controlled)) AS ldl_controlled_patients,
                    (SELECT COUNT(*) FROM DiabetesPatients WHERE kcstmr IN (SELECT kcstmr FROM DM_Controlled)) AS dm_controlled_patients,
                    (SELECT COUNT(*) FROM DiabetesPatients WHERE kcstmr IN (SELECT kcstmr FROM All_Controlled)) AS all_controlled_patients,
                    (SELECT COUNT(*) FROM CasedPatients WHERE kcstmr IN (SELECT kcstmr FROM BP_Controlled)) AS bp_controlled_patients_cased,
                    (SELECT COUNT(*) FROM CasedPatients WHERE kcstmr IN (SELECT kcstmr FROM LDL_Controlled)) AS ldl_controlled_patients_cased,
                    (SELECT COUNT(*) FROM CasedPatients WHERE kcstmr IN (SELECT kcstmr FROM DM_Controlled)) AS dm_controlled_patients_cased,
                    (SELECT COUNT(*) FROM CasedPatients WHERE kcstmr IN (SELECT kcstmr FROM All_Controlled)) AS all_controlled_patients_cased
            """)

            with self.engine.connect() as conn:
                result = conn.execute(query, {
                    'current_year': current_year_roc,
                    'one_year_ago': one_year_ago_roc
                }).fetchone()

                total_diabetes = result.total_diabetes_patients or 0
                total_cased = result.total_cased_patients or 0

                return {
                    'diabetes_case_rate': {
                        'numerator': total_cased,
                        'denominator': total_diabetes,
                        'percentage': round(total_cased / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'diabetes_eye_exam_rate': {
                        'numerator': result.diabetes_eye_exam_patients,
                        'denominator': total_diabetes,
                        'percentage': round(result.diabetes_eye_exam_patients / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'cased_eye_exam_rate': {
                        'numerator': result.cased_eye_exam_patients,
                        'denominator': total_cased,
                        'percentage': round(result.cased_eye_exam_patients / total_cased * 100, 2) if total_cased else 0
                    },
                    'diabetes_dm_control_rate': {
                        'numerator': result.dm_controlled_patients,
                        'denominator': total_diabetes,
                        'percentage': round(result.dm_controlled_patients / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'diabetes_bp_control_rate': {
                        'numerator': result.bp_controlled_patients,
                        'denominator': total_diabetes,
                        'percentage': round(result.bp_controlled_patients / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'diabetes_ldl_control_rate': {
                        'numerator': result.ldl_controlled_patients,
                        'denominator': total_diabetes,
                        'percentage': round(result.ldl_controlled_patients / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'diabetes_all_control_rate': {
                        'numerator': result.all_controlled_patients,
                        'denominator': total_diabetes,
                        'percentage': round(result.all_controlled_patients / total_diabetes * 100, 2) if total_diabetes else 0
                    },
                    'diabetes_dm_control_rate_cased': {
                        'numerator': result.dm_controlled_patients_cased,
                        'denominator': total_cased,
                        'percentage': round(result.dm_controlled_patients_cased / total_cased * 100, 2) if total_cased else 0
                    },
                    'diabetes_bp_control_rate_cased': {
                        'numerator': result.bp_controlled_patients_cased,
                        'denominator': total_cased,
                        'percentage': round(result.bp_controlled_patients_cased / total_cased * 100, 2) if total_cased else 0
                    },
                    'diabetes_ldl_control_rate_cased': {
                        'numerator': result.ldl_controlled_patients_cased,
                        'denominator': total_cased,
                        'percentage': round(result.ldl_controlled_patients_cased / total_cased * 100, 2) if total_cased else 0
                    },
                    'diabetes_all_control_rate_cased': {
                        'numerator': result.all_controlled_patients_cased,
                        'denominator': total_cased,
                        'percentage': round(result.all_controlled_patients_cased / total_cased * 100, 2) if total_cased else 0
                    }
                }
        except Exception as e:
            st.error(f"獲取糖尿病統計資料時發生錯誤：{str(e)}")
            return {}

    # ... 其他資料處理方法 ... 