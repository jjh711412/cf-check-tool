
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# 設定 - 言語選択
lang = st.sidebar.selectbox("言語 / Language", ["日本語", "한국어"])

# 文字列 設定
if lang == "日本語":
    TITLE = "現金流動表 チェックツール"
    UPLOAD_LABEL = "今月のファイル (現在)"
    PREV_UPLOAD_LABEL = "前月のファイル (比較用)"
    MODE_LABEL = "比較モード"
    MODE_OPTIONS = ["単月比較", "総計比較"]
    THRESHOLD_LABEL = "異常感知基準 (%)"
    RUN_ANALYSIS = "分析開始"
    TARGET_SHEET_LABEL = "監視対象シート"
    SUMMARY_LABEL = "サマリー結果"
    DETAIL_LABEL = "詳細異常リスト"
    RISK_LABEL = "リスクレベル評価"
else:
    TITLE = "현금흐름표 체크툴"
    UPLOAD_LABEL = "이번 달 파일 (현재)"
    PREV_UPLOAD_LABEL = "이전 달 파일 (비교용)"
    MODE_LABEL = "비교 모드"
    MODE_OPTIONS = ["단월 비교", "누계 비교"]
    THRESHOLD_LABEL = "이상 감지 기준 (%)"
    RUN_ANALYSIS = "분석 시작"
    TARGET_SHEET_LABEL = "감시 대상 시트"
    SUMMARY_LABEL = "요약 결과"
    DETAIL_LABEL = "상세 이상 리스트"
    RISK_LABEL = "위험도 평가"

# メインUI
st.title(TITLE)

uploaded_file = st.file_uploader(UPLOAD_LABEL, type=["xlsx"])
prev_file = st.file_uploader(PREV_UPLOAD_LABEL, type=["xlsx"])

mode = st.radio(MODE_LABEL, MODE_OPTIONS, horizontal=True)
threshold = st.slider(THRESHOLD_LABEL, min_value=1, max_value=100, value=10, step=1)

if uploaded_file and prev_file:
    if st.button(RUN_ANALYSIS):
        try:
            xls_now = pd.ExcelFile(uploaded_file, engine="openpyxl")
            xls_prev = pd.ExcelFile(prev_file, engine="openpyxl")

            if "meta" in xls_now.sheet_names:
                meta_df = xls_now.parse("meta", header=None)
                watch_mask = meta_df.iloc[:, 1].astype(str).str.contains("中間計算", na=False)
                target_sheets = meta_df[watch_mask].iloc[:, 0].tolist()

                all_results = []

                for sheet in target_sheets:
                    if sheet not in xls_now.sheet_names:
                        continue

                    df_now = xls_now.parse(sheet, header=None)
                    df_prev = xls_prev.parse(sheet, header=None) if sheet in xls_prev.sheet_names else None

                    errors = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!"]
                    formula_errors = df_now.applymap(lambda x: any(e in str(x) for e in errors if isinstance(x, str))).sum().sum()
                    if formula_errors > 0:
                        all_results.append([sheet, "数式エラーの検出", formula_errors, "重大"])

                    ref_lost = df_now.isin([0, "0", "=0", "", None]).sum().sum()
                    if ref_lost > 0:
                        risk = "要注意" if ref_lost > 10 else "軽微"
                        all_results.append([sheet, "参照の欠落", ref_lost, risk])

                    input_missing = df_now.isnull().sum().sum()
                    if input_missing > 0:
                        risk = "要注意" if input_missing > 20 else "軽微"
                        all_results.append([sheet, "入力漏れ", input_missing, risk])

                    if df_prev is not None:
                        row_diff = df_now.shape[0] - df_prev.shape[0]
                        col_diff = df_now.shape[1] - df_prev.shape[1]
                        if row_diff != 0 or col_diff != 0:
                            all_results.append([sheet, "構造の相違（行・列）", f"行:{row_diff}, 列:{col_diff}", "要注意"])

                if all_results:
                    result_df = pd.DataFrame(all_results, columns=["シート名", "チェック項目", "異常数", "リスクレベル"])
                    st.subheader(SUMMARY_LABEL)
                    st.dataframe(result_df)
                else:
                    st.success("異常は検出されませんでした。")
            else:
                st.error("metaシートがありません")
        except Exception as e:
            st.error(f"処理中エラー: {str(e)}")
else:
    st.info("2つのExcelファイルをアップロードしてください。")
