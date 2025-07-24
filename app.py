import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------
# UI Strings
# --------------------------
lang = st.sidebar.selectbox("言語 / Language", ["日本語", "한국어"])

if lang == "日本語":
    TITLE = "現金流動表 チェックツール"
    UPLOAD_LABEL = "現月のファイル (現在)"
    PREV1_LABEL = "前月のファイル"
    PREV2_LABEL = "前々月のファイル"
    MODE_LABEL = "比較モード"
    MODE_OPTIONS = ["單月比較", "総計比較"]
    THRESHOLD_LABEL = "異常感知基準 (%)"
    RUN_ANALYSIS = "分析開始"
    SUMMARY_LABEL = "サマリー結果"
    DETAIL_LABEL = "詳細異常リスト"
    EXPAND_LABEL = "詳細を見る"
    REPORT_LABEL = "数値比較レポート"
    ERROR_LABEL = "数式/入力/参照 チェック"
else:
    TITLE = "현금흐름표 체크툴"
    UPLOAD_LABEL = "이번 달 파일 (현재)"
    PREV1_LABEL = "전월 파일"
    PREV2_LABEL = "전전월 파일"
    MODE_LABEL = "비교 모드"
    MODE_OPTIONS = ["단월 비교", "누계 비교"]
    THRESHOLD_LABEL = "이상 감지 기준 (%)"
    RUN_ANALYSIS = "분석 시작"
    SUMMARY_LABEL = "요약 결과"
    DETAIL_LABEL = "상세 이상 리스트"
    EXPAND_LABEL = "상세 보기"
    REPORT_LABEL = "수치 비교 리포트"
    ERROR_LABEL = "수식/입력/참조 체크"

# --------------------------
# UI Layout
# --------------------------
st.title(TITLE)

uploaded_now = st.file_uploader(UPLOAD_LABEL, type=["xlsx"])
uploaded_prev1 = st.file_uploader(PREV1_LABEL, type=["xlsx"])
uploaded_prev2 = st.file_uploader(PREV2_LABEL, type=["xlsx"])

mode = st.radio(MODE_LABEL, MODE_OPTIONS, horizontal=True)
threshold = st.slider(THRESHOLD_LABEL, min_value=1, max_value=100, value=10, step=1)

# --------------------------
# Utility
# --------------------------
def get_target_sheets(meta_df):
    watch_mask = meta_df.iloc[:, 1].astype(str).str.contains("中間計算", na=False)
    return meta_df[watch_mask].iloc[:, 0].tolist()

def format_number(n):
    try:
        return f"{int(n):,}"
    except:
        return n

# --------------------------
# Main Analysis
# --------------------------
if uploaded_now:
    if st.button(RUN_ANALYSIS):
        try:
            xls_now = pd.ExcelFile(uploaded_now, engine="openpyxl")
            meta_df = xls_now.parse("meta", header=None) if "meta" in xls_now.sheet_names else None

            if meta_df is None:
                st.error("metaシートがありません")
            else:
                target_sheets = get_target_sheets(meta_df)
                report_dict = {}

                for sheet in target_sheets:
                    if sheet not in xls_now.sheet_names:
                        continue
                    df_now = xls_now.parse(sheet, header=0)

                    xls_prev1 = pd.ExcelFile(uploaded_prev1, engine="openpyxl") if uploaded_prev1 else None
                    xls_prev2 = pd.ExcelFile(uploaded_prev2, engine="openpyxl") if uploaded_prev2 else None

                    df_prev1 = xls_prev1.parse(sheet, header=0) if xls_prev1 and sheet in xls_prev1.sheet_names else None
                    df_prev2 = xls_prev2.parse(sheet, header=0) if xls_prev2 and sheet in xls_prev2.sheet_names else None

                    report_data = []
                    for i in range(df_now.shape[0]):
                        item_name = df_now.iloc[i, 0] if i < len(df_now) else f"項目{i}"
                        now_val = df_now.iloc[i, 1:].sum(numeric_only=True)
