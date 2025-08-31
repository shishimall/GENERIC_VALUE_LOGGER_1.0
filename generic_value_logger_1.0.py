# generic_value_logger

import streamlit as st
import pandas as pd
import datetime
import os
import gspread
from io import BytesIO
from openpyxl import Workbook

# ---------- Google Sheets認証設定 ----------
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

base_path = os.path.dirname(__file__)
json_path = os.path.join(base_path, "gspread_service_account.json")

try:
    gc = gspread.service_account(filename=json_path)
    SPREADSHEET_ID = "1n-jQhBD5u2jsv_cQskF81xy9p6lM5ZLcgmix22mQpho"
    worksheet = gc.open_by_key(SPREADSHEET_ID).sheet1
    sheet_data = worksheet.get_all_records()
    sheet_df = pd.DataFrame(sheet_data)
except Exception as e:
    sheet_df = pd.DataFrame()
    st.warning(f"Google Sheetsの読み込みに失敗しました: {e}")

# ---------- Streamlitアプリ ----------
st.set_page_config(page_title="汎用値記録", layout="wide")
st.title("📒 汎用値記録")

# ---------- 入力フォーム ----------
st.sidebar.subheader(":pencil: 新規記録")
category = st.sidebar.text_input("カテゴリ")
value = st.sidebar.text_input("値")
note = st.sidebar.text_area("メモ")

if st.sidebar.button("記録"):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    new_data = pd.DataFrame([[now, category, value, note]], columns=["日時", "カテゴリ", "値", "メモ"])

    sheet_df = pd.concat([new_data, sheet_df], ignore_index=True)

    # Sheetsに追記
    try:
        worksheet.insert_row([now, category, value, note], index=2)
    except Exception as e:
        st.error(f"Google Sheetsへの書き込みに失敗しました: {e}")

# ---------- 表示 ----------
st.dataframe(sheet_df, use_container_width=True)

# ---------- Excelダウンロード ----------
def convert_df_to_excel(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "記録"
    ws.append(df.columns.tolist())
    for row in df.values:
        ws.append(row.tolist())
    wb.save(output)
    return output.getvalue()

if not sheet_df.empty:
    excel_data = convert_df_to_excel(sheet_df)
    st.download_button(
        label="📄 Excelダウンロード",
        data=excel_data,
        file_name="汎用スケール記録.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
