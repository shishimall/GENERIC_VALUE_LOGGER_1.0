# generic_value_logger (cloud/local 両対応)

import streamlit as st
import pandas as pd
import datetime
import os
import gspread
from io import BytesIO
from openpyxl import Workbook
from google.oauth2.service_account import Credentials

# ---------- Google Sheets認証設定 ----------
SCOPE = ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]

def get_gspread_client():
    """
    優先順:
      1) st.secrets["gspread_service_account"] (Streamlit Cloud)
      2) ローカルの gspread_service_account.json
    """
    # 1) Cloud: st.secrets から読み取り
    try:
        svc_info = st.secrets.get("gspread_service_account", None)
        if svc_info:
            creds = Credentials.from_service_account_info(svc_info, scopes=SCOPE)
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"st.secretsからの認証に失敗: {e}")

    # 2) Local: JSONファイルから読み取り
    try:
        base_path = os.path.dirname(__file__)
        json_path = os.path.join(base_path, "gspread_service_account.json")
        creds = Credentials.from_service_account_file(json_path, scopes=SCOPE)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"ローカルJSONからの認証に失敗: {e}")
        return None

def get_spreadsheet_id():
    """
    Cloud では st.secrets['app']['SPREADSHEET_ID'] を優先。
    無ければハードコード値を使用（必要に応じて書き換え可）。
    """
    try:
        sid = st.secrets.get("app", {}).get("SPREADSHEET_ID", "").strip()
        if sid:
            return sid
    except Exception:
        pass
    # フォールバック（従来のID）
    return "1n-jQhBD5u2jsv_cQskF81xy9p6lM5ZLcgmix22mQpho"

# --- クライアントとデータ読込
gc = get_gspread_client()
worksheet = None
sheet_df = pd.DataFrame()

if gc:
    try:
        SPREADSHEET_ID = get_spreadsheet_id()
        worksheet = gc.open_by_key(SPREADSHEET_ID).sheet1
        sheet_data = worksheet.get_all_records()
        sheet_df = pd.DataFrame(sheet_data)
    except Exception as e:
        st.warning(f"Google Sheetsの読み込みに失敗しました: {e}")
else:
    st.warning("Google Sheets の認証クライアントが取得できませんでした。")

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

    # 画面側の即時反映
    sheet_df = pd.concat([new_data, sheet_df], ignore_index=True)

    # Sheetsに追記（Cloud/Localいずれも同じ）
    if worksheet:
        try:
            worksheet.insert_row([now, category, value, note], index=2)
            st.success("書き込みに成功しました。")
        except Exception as e:
            st.error(f"Google Sheetsへの書き込みに失敗しました: {e}")
    else:
        st.error("ワークシートが未接続のため書き込みできません。")

# ---------- 表示 ----------
st.dataframe(sheet_df, use_container_width=True)

# ---------- Excelダウンロード ----------
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "記録"
    if not df.empty:
        ws.append(df.columns.tolist())
        for row in df.itertuples(index=False, name=None):
            ws.append(list(row))
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
