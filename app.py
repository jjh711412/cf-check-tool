import streamlit as st
import pandas as pd

# 항목 매핑 예시 (추후 meta 시트 기반으로 자동 생성 가능)
item_mapping = {
    "1": "調整後EBITDA",
    "2": "運転資本の増減",
    "3": "割賦未払金及びリース債務の支払い",
    "4": "法人税支払",
    "5": "純利益"
}

def generate_monthly_report_by_colname(df_now, df_prev, item_mapping, code_col="code", value_col="累計額"):
    report = []
    for idx in range(len(df_now)):
        code = str(df_now.loc[idx, code_col]).strip()
        name = item_mapping.get(code, "不明な項目")

        try:
            val_now = float(df_now.loc[idx, value_col])
            val_prev = float(df_prev.loc[idx, value_col])
            if pd.isna(val_now) or pd.isna(val_prev):
                continue
        except (ValueError, TypeError):
            continue

        diff = val_now - val_prev
        rate = (diff / val_prev * 100) if val_prev != 0 else 0

        report.append({
            "コード": code,
            "項目名": name,
            "現月": f"{int(round(val_now)):,}",
            "前月": f"{int(round(val_prev)):,}",
            "増減額": f"{int(round(diff)):,}",
            "増減率": f"{round(rate, 1)}%"
        })

    return pd.DataFrame(report)

# UI 시작
st.set_page_config(page_title="月次CFチェックツール", layout="wide")
st.title("📊 現金流動表：月次比較レポート")

uploaded_file = st.file_uploader("📤 今月ファイルをアップロード", type=["xlsx"])
prev_file = st.file_uploader("📥 先月ファイルをアップロード", type=["xlsx"])

if uploaded_file and prev_file:
    try:
        df_now = pd.read_excel(uploaded_file, sheet_name="(報告用)CVCフォーマット", header=None)
        df_prev = pd.read_excel(prev_file, sheet_name="(報告用)CVCフォーマット", header=None)

        header_idx = df_now.apply(lambda row: row.astype(str).str.contains("code", na=False)).any(axis=1)
        header_row_index = df_now[header_idx].index[0] if header_idx.any() else 0

        df_now.columns = df_now.iloc[header_row_index]
        df_now = df_now.iloc[header_row_index + 1:].reset_index(drop=True)
        df_prev.columns = df_prev.iloc[header_row_index]
        df_prev = df_prev.iloc[header_row_index + 1:].reset_index(drop=True)

        result = generate_monthly_report_by_colname(df_now, df_prev, item_mapping)

        if not result.empty:
            st.success("📈 増減比較が完了しました")
            st.dataframe(result, use_container_width=True)
        else:
            st.warning("比較可能な数値データが見つかりませんでした。")

    except Exception as e:
        st.error(f"処理中エラー: {e}")
else:
    st.info("🔁 今月と前月の2つのExcelファイルをアップロードしてください。")
