import streamlit as st
import pandas as pd

# 언어 설정
lang = st.sidebar.selectbox("言語 / Language", ["日本語", "한국어"])
if lang == "日本語":
    TITLE = "現金流動表チェックツール (デバッグモード)"
    UPLOAD_LABEL = "今月のファイル（現在）"
    PREV_LABEL = "先月のファイル（比較対象）"
    MODE_LABEL = "比較モード"
    MODE_OPTIONS = ["單月比較", "累計比較"]
    THRESHOLD_LABEL = "異常感知基準 (%)"
    RUN_BUTTON = "分析開始"
else:
    TITLE = "현금흐름표 체크툴 (디버그모드)"
    UPLOAD_LABEL = "이번 달 파일 (현재)"
    PREV_LABEL = "이전 달 파일 (비교 대상)"
    MODE_LABEL = "비교 모드"
    MODE_OPTIONS = ["단월 비교", "누계 비교"]
    THRESHOLD_LABEL = "이상 감지 기준 (%)"
    RUN_BUTTON = "분석 시작"

st.set_page_config(page_title=TITLE, layout="wide")
st.title(TITLE)

file_now = st.file_uploader(UPLOAD_LABEL, type=["xlsx"])
file_prev = st.file_uploader(PREV_LABEL, type=["xlsx"])
mode = st.radio(MODE_LABEL, MODE_OPTIONS, horizontal=True)
threshold = st.slider(THRESHOLD_LABEL, 5, 100, 10)

if file_now and file_prev and st.button(RUN_BUTTON):
    try:
        st.info("📤 ファイルを読み込み中です...")

        sheet_name = "(報告用)CVCフォーマット"
        sheet_now = pd.read_excel(file_now, sheet_name=sheet_name, header=None)
        sheet_prev = pd.read_excel(file_prev, sheet_name=sheet_name, header=None)

        header_idx = sheet_now.apply(lambda row: row.astype(str).str.contains("code", na=False)).any(axis=1)
        header_row = sheet_now[header_idx].index[0] if header_idx.any() else 0

        st.info(f"🧭 ヘッダー検出行: {header_row}")
        st.write("📋 カラム一覧:", sheet_now.iloc[header_row].tolist())

        sheet_now.columns = sheet_now.iloc[header_row]
        sheet_prev.columns = sheet_prev.iloc[header_row]
        df_now = sheet_now.iloc[header_row+1:].reset_index(drop=True)
        df_prev = sheet_prev.iloc[header_row+1:].reset_index(drop=True)

        # 🔍 주요 열 확인
        code_col_candidates = [col for col in df_now.columns if "code" in str(col).lower()]
        value_col_candidates = [col for col in df_now.columns if "単月" in str(col) or "累計" in str(col)]

        if not code_col_candidates:
            st.error("❌ 'code' 라는 이름의 열이 없음. '項目コード', 'コード' 같은 이름일 수 있어.")
        if not value_col_candidates:
            st.error("❌ '単月額' 또는 '累計額' 같은 금액 열이 없음.")

        code_col = code_col_candidates[0] if code_col_candidates else None
        value_col = value_col_candidates[0] if value_col_candidates else None

        if not code_col or not value_col:
            st.stop()

        st.success(f"✅ 항목코드 열: {code_col}, 금액 열: {value_col}")

        results = []
        missing_prev_count = 0
        conversion_fail_count = 0

        for _, row in df_now.iterrows():
            code = str(row.get(code_col)).strip()
            try:
                val_now = float(row.get(value_col))
            except:
                conversion_fail_count += 1
                continue

            try:
                val_prev = float(df_prev[df_prev[code_col] == code][value_col].values[0])
            except:
                missing_prev_count += 1
                continue

            if pd.isna(val_now) or pd.isna(val_prev): continue

            diff = val_now - val_prev
            rate = (diff / val_prev * 100) if val_prev != 0 else 0

            if abs(rate) >= threshold:
                results.append({
                    "コード": code,
                    "現月": f"{int(val_now):,}",
                    "前月": f"{int(val_prev):,}",
                    "増減額": f"{int(diff):,}",
                    "増減率": f"{rate:.1f}%"
                })

        st.write("📊 이상 증감 결과")
        if results:
            st.dataframe(pd.DataFrame(results))
        else:
            st.warning("❗ 이상 수치 없음. 혹은 추출 실패.")

        st.info(f"🔎 비교 불가능한 항목 개수 (전월에 없음): {missing_prev_count}")
        st.info(f"⚠️ 금액 추출 실패한 항목 개수: {conversion_fail_count}")

    except Exception as e:
        st.error(f"処理中エラー: {e}")
