# generic_value_logger (Cloud/Local 自動切替・強化版)

import streamlit as st
import pandas as pd
import datetime
import os, json
import gspread
from io import BytesIO
from openpyxl import Workbook
from google.oauth2.service_account import Credentials

SCOPE = ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]

REQUIRED_KEYS = {
    "type","project_id","private_key_id","private_key","client_email","client_id","token_uri"
}

def _normalize_private_key(k: str) -> str:
    # TOML では "\n" として保存されることが多いので実改行に戻す
    return k.replace("\\n", "\n") if "\\n" in k else k

def _maybe_credentials_from_info(info: dict):
    # private_key 正規化
    if "private_key" in info and isinstance(info["private_key"], str):
        info = dict(info)
        info["private_key"] = _normalize_private_key(info["private_key"])
    return Credentials.from_service_account_info(info, scopes=SCOPE)

def get_gspread_client():
    """
    認証の探索順:
      1) st.secrets["gspread_service_account"] という “テーブル” がある
      2) st.secrets にサービスアカウントのキーが “直置き” されている
      3) st.secrets["GSPREAD_SERVICE_ACCOUNT_JSON"] に JSON 文字列が入っている
      4) ローカル gspread_service_account.json
    """
    # 1) [gspread_service_account] テーブル
    try:
        svc_tbl = st.secrets.get("gspread_service_account", None)
        if svc_tbl and REQUIRED_KEYS.issubset(svc_tbl.keys()):
            creds = _maybe_credentials_from_info(dict(svc_tbl))
            st.info("認証: st.secrets[gspread_service_account] を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secretsテーブルからの認証に失敗: {e}")

    # 2) トップレベル直置き (type / client_email などが st.secrets 直下にある)
    try:
        if REQUIRED_KEYS.issubset(set(st.secrets.keys())):
            info = {k: st.secrets[k] for k in REQUIRED_KEYS}
            # 他の任意キーもあれば足す
            for k in ("auth_uri","token_uri","auth_provider_x509_cert_url","client_x509_cert_url","universe_domain"):
                if k in st.secrets: info[k] = st.secrets[k]
            creds = _maybe_credentials_from_info(info)
            st.info("認証: st.secrets (トップレベル直置き) を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secrets直置きからの認証に失敗: {e}")

    # 3) JSON文字列で保存しているパターン
    try:
        if "GSPREAD_SERVICE_ACCOUNT_JSON" in st.secrets:
            info = json.loads(st.secrets["GSPREAD_SERVICE_ACCOUNT_JSON"])
            creds = _maybe_credentials_from_info(info)
            st.info("認証: st.secrets['GSPREAD_SERVICE_ACCOUNT_JSON'] を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secretsのJSON文字列からの認証に失敗: {e}")

    # 4) ローカル
    try:
        base_path = os.path.dirname(__file__)
        json_path = os.path.join(base_path, "gspread_service_account.json")
        creds = Credentials.from_service_account_file(json_path, scopes=SCOPE)
        st.info("認証: ローカル gspread_service_account.json を使用", icon="🖥️")
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"ローカルJSONからの認証に失敗: {e}")
        return None

def get_spreadsheet_id():
    # 優先順: secrets["app"]["SPREADSHEET_ID"] → secrets["SPREADSHEET_ID"] → フォールバック値
    try:
        sid = st.secrets.get("app", {}).get("SPREADSHEET_ID", "").strip()
        if sid: return sid
    except Exception:
        pass
    sid2 = str(st.secrets.get("SPREADSHEET_ID", "")).strip() if hasattr(st, "secrets") else ""
    return sid2 or "1n-jQhBD5u2jsv_cQskF81xy9p6lM5ZLcgmix22mQpho"

# ---- 起動処理
st.set_page_config(page_title="汎用値記録", layout="wide")
st.title("📒 汎用値記録")

gc = get_gspread_client()
worksheet = None
sheet_df = pd.DataFrame(columns=["日時","カテゴリ","値","メモ"])

if gc:
    try:
        SPREADSHEET_ID = get_spreadsheet_id()
        worksheet = gc.open_by_key(SPREADSHEET_ID).sheet1
        data = worksheet.get_all_records()
        if data:
            sheet_df = pd.DataFrame(data)
        else:
            sheet_df = pd.DataFrame(columns=["日時","カテゴリ","値","メモ"])
    except Exception as e:
        st.warning(f"Google Sheetsの読み込みに失敗しました: {e}")
else:
    st.warning("Google Sheets の認証クライアントが取得できませんでした。")

# ---- 入力フォーム
st.sidebar.subheader(":pencil: 新規記録")
category = st.sidebar.text_input("カテゴリ")
value = st.sidebar.text_input("値")
note = st.sidebar.text_area("メモ")

if st.sidebar.button("記録"):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    new_row = {"日時": now, "カテゴリ": category, "値": value, "メモ": note}
    sheet_df = pd.concat([pd.DataFrame([new_row]), sheet_df], ignore_index=True)

    if worksheet:
        try:
            worksheet.insert_row([now, category, value, note], index=2)
            st.success("書き込みに成功しました。")
        except Exception as e:
            st.error(f"Google Sheetsへの書き込みに失敗しました: {e}")
    else:
        st.error("ワークシート未接続のため書き込みできません。")

st.dataframe(sheet_df if not sheet_df.empty else pd.DataFrame([{"empty": ""}]), use_container_width=True)

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
    st.download_button(
        label="📄 Excelダウンロード",
        data=convert_df_to_excel(sheet_df),
        file_name="汎用スケール記録.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
