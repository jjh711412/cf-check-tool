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
    PREV3_LABEL = "前々々月のファイル"
    MODE_LABEL = "比較モード"
    MODE_OPTIONS = ["單月比較", "総計比較"]
    THRESHOLD_LABEL = "異常感知基準 (%)"
    RUN_ANALYSIS = "分析開始"
    SUMMARY_LABEL = "サマリー結果"
    DETAIL_LABEL = "詳細異常リスト"
else:
    TITLE = "현금흐름표 체크툴"
    UPLOAD_LABEL = "이번 달 파일 (현재)"
    PREV1_LABEL = "전월 파일"
    PREV2_LABEL = "전전월 파일"
    PREV3_LABEL = "전전전월 파일"
    MODE_LABEL = "비교 모드"
    MODE_OPTIONS = ["단월 비교", "누계 비교"]
    THRESHOLD_LABEL = "이상 감지 기준 (%)"
    RUN_ANALYSIS = "분석 시작"
    SUMMARY_LABEL = "요약 결과"
    DETAIL_LABEL = "상세 이상 리스트"

# --------------------------
# UI Layout
# --------------------------
st.title(TITLE)

uploaded_now = st.file_uploader(UPLOAD_LABEL, type=["xlsx"])
uploaded_prev1 = st.file_uploader(PREV1_LABEL, type=["xlsx"])
uploaded_prev2 = st.file_uploader(PREV2_LABEL, type=["xlsx"])
uploaded_prev3 = st.file_uploader(PREV3_LABEL, type=["xlsx"])

mode = st.radio(MODE_LABEL, MODE_OPTIONS, horizontal=True)
threshold = st.slider(THRESHOLD_LABEL, min_value=1, max_value=100, value=10, step=1)

compare_files = {
    "前月": uploaded_prev1,
    "前々月": uploaded_prev2,
    "前々々月": uploaded_prev3,
}

# --------------------------
# Utility
# --------------------------
def get_target_sheets(meta_df):
    watch_mask = meta_df.iloc[:, 1].astype(str).str.contains("中間計算", na=False)
    return meta_df[watch_mask].iloc[:, 0].tolist()

def calculate_diff(df_now, df_prev):
    try:
        df1 = df_now.fillna(0).select_dtypes(include='number')
        df2 = df_prev.fillna(0).select_dtypes(include='number')
        diff = (df1 - df2).abs()
        max_change = diff.max().max()
        return max_change
    except:
        return None

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
                report_rows = []

                for sheet in target_sheets:
                    if sheet not in xls_now.sheet_names:
                        continue
                    df_now = xls_now.parse(sheet, header=None)

                    for label, file in compare_files.items():
                        if file is not None:
                            xls_prev = pd.ExcelFile(file, engine="openpyxl")
                            if sheet in xls_prev.sheet_names:
                                df_prev = xls_prev.parse(sheet, header=None)
                                try:
                                    df1 = df_now.fillna(0).select_dtypes(include='number')
                                    df2 = df_prev.fillna(0).select_dtypes(include='number')
                                    diff = ((df1 - df2) / (df2.replace(0, 1))).abs() * 100
                                    over_threshold = diff > threshold
                                    count = over_threshold.sum().sum()
                                    if count > 0:
                                        report_rows.append([sheet, label, int(count), f"{threshold}%超過"])
                                except:
                                    continue

                if report_rows:
                    result_df = pd.DataFrame(report_rows, columns=["シート名", "比較対象", "異常セル数", "条件"])
                    st.subheader(SUMMARY_LABEL)
                    st.dataframe(result_df)
                else:
                    st.success("異常は検出されませんでした。")

        except Exception as e:
            st.error(f"処理中エラー: {str(e)}")
else:
    st.info("現月ファイルをアップロードしてください。")
