import streamlit as st
import pandas as pd

# 🎌 다국어 설정
lang = st.sidebar.selectbox("言語 / Language", ["日本語", "한국어"])
if lang == "日本語":
    TITLE = "現金流動表チェックツール"
    UPLOAD_LABEL = "今月のファイル（現在）"
    PREV_LABEL = "先月のファイル（比較対象）"
    MODE_LABEL = "比較モード"
    MODE_OPTIONS = ["單月比較", "累計比較"]
    THRESHOLD_LABEL = "異常感知基準 (%)"
    RUN_BUTTON = "分析開始"
    SUMMARY = "📋 異常サマリー"
    DETAILS = "📌 詳細異常リスト"
    MGMT_QUESTIONS = "📣 経営者が考慮すべき質問"
else:
    TITLE = "현금흐름표 체크툴"
    UPLOAD_LABEL = "이번 달 파일 (현재)"
    PREV_LABEL = "이전 달 파일 (비교 대상)"
    MODE_LABEL = "비교 모드"
    MODE_OPTIONS = ["단월 비교", "누계 비교"]
    THRESHOLD_LABEL = "이상 감지 기준 (%)"
    RUN_BUTTON = "분석 시작"
    SUMMARY = "📋 이상 요약"
    DETAILS = "📌 상세 이상 항목"
    MGMT_QUESTIONS = "📣 경영진 확인 포인트"

# 🖥️ UI 기본
st.set_page_config(page_title=TITLE, layout="wide")
st.title(TITLE)

file_now = st.file_uploader(UPLOAD_LABEL, type=["xlsx"])
file_prev = st.file_uploader(PREV_LABEL, type=["xlsx"])
mode = st.radio(MODE_LABEL, MODE_OPTIONS, horizontal=True)
threshold = st.slider(THRESHOLD_LABEL, 5, 100, 10)

# 📊 분석 버튼 클릭
if file_now and file_prev and st.button(RUN_BUTTON):
    try:
        # 📂 시트 읽기
        sheet_now = pd.read_excel(file_now, sheet_name="(報告用)CVCフォーマット", header=None)
        sheet_prev = pd.read_excel(file_prev, sheet_name="(報告用)CVCフォーマット", header=None)

        # ⛳️ header 자동 탐색
        header_idx = sheet_now.apply(lambda row: row.astype(str).str.contains("code", na=False)).any(axis=1)
        header_row = sheet_now[header_idx].index[0] if header_idx.any() else 0
        sheet_now.columns = sheet_now.iloc[header_row]
        sheet_prev.columns = sheet_prev.iloc[header_row]
        df_now = sheet_now.iloc[header_row+1:].reset_index(drop=True)
        df_prev = sheet_prev.iloc[header_row+1:].reset_index(drop=True)

        # 📉 비교 컬럼
        value_col = "単月額" if "單月" in mode else "累計額"
        code_col = "code"

        results = []
        questions = []

        for _, row in df_now.iterrows():
            code = str(row.get(code_col, "")).strip()
            name = str(row.get("日本語", "")).strip()
            try:
                val_now = float(row.get(value_col))
                val_prev = float(df_prev[df_prev[code_col] == code][value_col].values[0])
            except:
                continue

            if pd.isna(val_now) or pd.isna(val_prev): continue

            diff = val_now - val_prev
            rate = (diff / val_prev * 100) if val_prev != 0 else 0

            if abs(rate) >= threshold:
                results.append({
                    "コード": code,
                    "項目": name,
                    "現月": f"{int(val_now):,}",
                    "前月": f"{int(val_prev):,}",
                    "増減額": f"{int(diff):,}",
                    "増減率": f"{rate:.1f}%"
                })

                # 📣 경영진 질문 예시
                if "投資" in name or "設備" in name:
                    questions.append(f"・『{name}』において{rate:.1f}%の変動が確認されました。これは一時的なものでしょうか？")

        # 🧾 결과 출력
        st.subheader(SUMMARY)
        if results:
            st.dataframe(pd.DataFrame(results), use_container_width=True)
        else:
            st.success("異常な増減は検出されませんでした。")

        if questions:
            st.subheader(MGMT_QUESTIONS)
            for q in questions:
                st.markdown(q)

        # 🔎 수식 에러 등 감시 시트 분석
        xls_now = pd.ExcelFile(file_now)
        if "meta" in xls_now.sheet_names:
            meta_df = xls_now.parse("meta", header=None)
            watch_mask = meta_df.iloc[:, 1].astype(str).str.contains("中間計算", na=False)
            target_sheets = meta_df[watch_mask].iloc[:, 0].tolist()

            detail_logs = []
            for sheet in target_sheets:
                if sheet not in xls_now.sheet_names: continue
                df = xls_now.parse(sheet, header=None)

                errors = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!"]
                err_count = df.applymap(lambda x: any(e in str(x) for e in errors if isinstance(x, str))).sum().sum()
                if err_count > 0:
                    detail_logs.append([sheet, "数式エラー", err_count, "重大"])

                blanks = df.isin(["", None, 0, "0"]).sum().sum()
                if blanks > 10:
                    detail_logs.append([sheet, "入力/参照の欠落", blanks, "要注意"])

                # 構造比較
                df_prev = pd.read_excel(file_prev, sheet_name=sheet, header=None) if sheet in pd.ExcelFile(file_prev).sheet_names else None
                if df_prev is not None:
                    row_diff = df.shape[0] - df_prev.shape[0]
                    col_diff = df.shape[1] - df_prev.shape[1]
                    if row_diff != 0 or col_diff != 0:
                        detail_logs.append([sheet, "構造変化", f"行: {row_diff}, 列: {col_diff}", "注意"])

            if detail_logs:
                st.subheader(DETAILS)
                df_detail = pd.DataFrame(detail_logs, columns=["シート名", "異常タイプ", "内容", "レベル"])
                st.dataframe(df_detail, use_container_width=True)

    except Exception as e:
        st.error(f"処理中エラー: {e}")
