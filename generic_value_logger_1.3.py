# generic_value_logger (Cloud/Local 自動切替・入力/編集フォーム・並べ替え対応 “全文”)
# ------------------------------------------------------------
# 仕様：
# - 認証は Cloud: st.secrets / Local: gspread_service_account.json を自動判定（どちらも対応）
# - 新規記録（サイドバー）：ボタン押下までリロードなし
# - メイン表（data_editor）：編集 → 「編集を保存」までリロード/書き込みなし
# - 並べ替え（第1/第2キー・昇降順・保存時に順序反映するか選択）をフォーム適用時のみ反映
# - Excel ダウンロード（現在表示の内容）
# ------------------------------------------------------------

import streamlit as st
import pandas as pd
import datetime
import os, json
import gspread
from io import BytesIO
from openpyxl import Workbook
from google.oauth2.service_account import Credentials

# ---------------- Google 認証まわり（安定版） ----------------
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
REQUIRED_KEYS = {"type","project_id","private_key_id","private_key","client_email","client_id","token_uri"}
RECOMMENDED_KEYS = {"auth_uri","auth_provider_x509_cert_url","client_x509_cert_url","universe_domain"}

def _normalize_private_key(k: str) -> str:
    return k.replace("\\n","\n") if isinstance(k,str) and "\\n" in k else k

def _creds_from_info(info: dict) -> Credentials:
    info = dict(info)
    if "private_key" in info:
        info["private_key"] = _normalize_private_key(info["private_key"])
    return Credentials.from_service_account_info(info, scopes=SCOPE)

def get_gspread_client():
    # 1) [gspread_service_account]
    try:
        tbl = st.secrets.get("gspread_service_account", None)
        if tbl and REQUIRED_KEYS.issubset(tbl.keys()):
            creds = _creds_from_info(tbl)
            st.info("認証: st.secrets[gspread_service_account] を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secrets[gspread_service_account] 失敗: {e}")

    # 2) [gcp_service_account]
    try:
        tbl2 = st.secrets.get("gcp_service_account", None)
        if tbl2 and REQUIRED_KEYS.issubset(tbl2.keys()):
            creds = _creds_from_info(tbl2)
            st.info("認証: st.secrets[gcp_service_account] を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secrets[gcp_service_account] 失敗: {e}")

    # 3) 直置き
    try:
        root_keys = set(getattr(st, "secrets", {}).keys())
        if REQUIRED_KEYS.issubset(root_keys):
            info = {k: st.secrets[k] for k in REQUIRED_KEYS}
            for k in RECOMMENDED_KEYS:
                if k in st.secrets: info[k] = st.secrets[k]
            creds = _creds_from_info(info)
            st.info("認証: st.secrets(直置き) を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secrets直置き 失敗: {e}")

    # 4) JSON文字列
    try:
        if "GSPREAD_SERVICE_ACCOUNT_JSON" in st.secrets:
            info = json.loads(st.secrets["GSPREAD_SERVICE_ACCOUNT_JSON"])
            creds = _creds_from_info(info)
            st.info("認証: st.secrets['GSPREAD_SERVICE_ACCOUNT_JSON'] を使用", icon="🔐")
            return gspread.authorize(creds)
    except Exception as e:
        st.warning(f"secrets JSON文字列 失敗: {e}")

    # 5) ローカル
    try:
        base = os.path.dirname(__file__)
        json_path = os.path.join(base, "gspread_service_account.json")
        creds = Credentials.from_service_account_file(json_path, scopes=SCOPE)
        st.info("認証: ローカル gspread_service_account.json を使用", icon="🖥️")
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"ローカルJSON 認証失敗: {e}")
        return None

def get_spreadsheet_id():
    try:
        sid = st.secrets.get("app", {}).get("SPREADSHEET_ID", "").strip()
        if sid: return sid
    except Exception:
        pass
    sid2 = str(st.secrets.get("SPREADSHEET_ID", "")).strip() if hasattr(st,"secrets") else ""
    return sid2 or "1n-jQhBD5u2jsv_cQskF81xy9p6lM5ZLcgmix22mQpho"

# ---------------- アプリ設定 ----------------
st.set_page_config(page_title="汎用値記録", layout="wide")
st.title("📒 汎用値記録")

# 管理カラム（横持ちヘッダー）
COLUMNS = ["日時", "カテゴリ", "項目", "値", "単位", "補足"]

gc = get_gspread_client()
worksheet = None

def fetch_sheet_df():
    """Google Sheets → DataFrame（ヘッダー補正込み）"""
    if not gc:
        return None, pd.DataFrame(columns=COLUMNS)
    try:
        sid = get_spreadsheet_id()
        ws = gc.open_by_key(sid).sheet1
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            df = pd.DataFrame(columns=COLUMNS)
        else:
            # 必須カラムが足りない場合に補完・順序を合わせる
            for c in COLUMNS:
                if c not in df.columns:
                    df[c] = ""
            df = df[COLUMNS]
        return ws, df
    except Exception as e:
        st.warning(f"Google Sheetsの読み込みに失敗: {e}")
        return None, pd.DataFrame(columns=COLUMNS)

if gc:
    worksheet, _df = fetch_sheet_df()
else:
    _df = pd.DataFrame(columns=COLUMNS)

# --------- セッション状態（編集用の作業コピー） ---------
if "df" not in st.session_state:
    st.session_state.df = _df.copy()
if "last_saved_df" not in st.session_state:
    st.session_state.last_saved_df = _df.copy()

# ===================== サイドバー：新規記録（記録押下までリロードなし） =====================
st.sidebar.subheader(":pencil: 新規記録（記録ボタンまでリロードしません）")
with st.sidebar.form(key="create_form", clear_on_submit=True):
    in_category = st.text_input("カテゴリ", key="in_category")
    in_item     = st.text_input("項目", key="in_item")
    in_value    = st.text_input("値", key="in_value")
    in_unit     = st.text_input("単位", key="in_unit")
    in_note     = st.text_input("補足", key="in_note")
    submitted_create = st.form_submit_button("記録")

if submitted_create:
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    new_row = {"日時": now, "カテゴリ": in_category, "項目": in_item,
               "値": in_value, "単位": in_unit, "補足": in_note}
    # 画面の作業コピーに先頭挿入
    st.session_state.df = pd.concat([pd.DataFrame([new_row]), st.session_state.df], ignore_index=True)

    # シートにも追記（上から2行目）
    if worksheet:
        try:
            worksheet.insert_row([now, in_category, in_item, in_value, in_unit, in_note], index=2)
            st.success("新規記録を保存しました。")
            st.session_state.last_saved_df = st.session_state.df.copy()
        except Exception as e:
            st.error(f"Sheetsへの書き込みに失敗: {e}")
    else:
        st.error("ワークシート未接続のため書き込みできません。")

# ===================== 並べ替え（適用ボタンまでリロードなし） =====================
def _mk_sort_columns(df: pd.DataFrame, colname: str) -> pd.Series:
    """列ごとに“並べ替え用”の型へ変換（日時/数値/文字列）"""
    s = df[colname].fillna("")
    if colname == "日時":
        return pd.to_datetime(s, errors="coerce")
    if colname == "値":
        return pd.to_numeric(s, errors="coerce")
    return s.astype(str)

def _apply_sort(df: pd.DataFrame, sort1: str|None, asc1: bool,
                sort2: str|None, asc2: bool) -> pd.DataFrame:
    out = df.copy()
    keys = []; asc = []
    if sort1 and sort1 in out.columns:
        out["_sort1"] = _mk_sort_columns(out, sort1)
        keys.append("_sort1"); asc.append(asc1)
    if sort2 and sort2 in out.columns and sort2 != sort1:
        out["_sort2"] = _mk_sort_columns(out, sort2)
        keys.append("_sort2"); asc.append(asc2)
    if keys:
        out = out.sort_values(by=keys, ascending=asc, kind="mergesort")  # 安定ソート
        out = out.drop(columns=[c for c in ["_sort1","_sort2"] if c in out.columns])
    return out

if "sort_settings" not in st.session_state:
    st.session_state.sort_settings = {
        "sort1": "日時",
        "asc1": False,   # 日付は新しい順がデフォルト
        "sort2": None,
        "asc2": True,
        "save_in_sorted_order": True,  # 保存時に表示順を反映
    }

with st.form(key="sort_form"):
    st.markdown("#### 並べ替え（適用を押すまで反映しません）")
    c1, c2, c3 = st.columns([2,1,2])
    with c1:
        sort1 = st.selectbox(
            "第1キー",
            options=[None]+COLUMNS,
            index=(COLUMNS.index(st.session_state.sort_settings["sort1"])+1) if st.session_state.sort_settings["sort1"] in COLUMNS else 0
        )
    with c2:
        asc1 = st.radio("順序(第1)", options=["昇順","降順"], horizontal=True,
                        index=0 if st.session_state.sort_settings["asc1"] else 1) == "昇順"
    with c3:
        sort2 = st.selectbox(
            "第2キー（任意）",
            options=[None]+COLUMNS,
            index=([None]+COLUMNS).index(st.session_state.sort_settings["sort2"])
        )
    asc2 = st.radio("順序(第2)", options=["昇順","降順"], horizontal=True,
                    index=0 if st.session_state.sort_settings["asc2"] else 1) == "昇順"

    save_sorted = st.checkbox(
        "保存時に表示中の並び順を反映する",
        value=st.session_state.sort_settings["save_in_sorted_order"]
    )
    colA, colB = st.columns([1,1])
    apply_sort = colA.form_submit_button("適用")
    clear_sort = colB.form_submit_button("ソート解除")

if clear_sort:
    st.session_state.sort_settings = {"sort1": None, "asc1": True, "sort2": None, "asc2": True, "save_in_sorted_order": True}
if apply_sort:
    st.session_state.sort_settings = {
        "sort1": sort1, "asc1": asc1, "sort2": sort2, "asc2": asc2,
        "save_in_sorted_order": save_sorted
    }

# 表示用データ（指定があれば並べ替えを適用）
ss = st.session_state.sort_settings
display_df = _apply_sort(st.session_state.df, ss["sort1"], ss["asc1"], ss["sort2"], ss["asc2"])

# ===================== メイン：テーブル編集（保存ボタンまで書き込み/リロードなし） =====================
st.markdown("### 表示・編集")
st.caption("※ 表は編集可能です。**下の「編集を保存」ボタンを押すまでリロードも書き込みも行いません。**")

with st.form(key="edit_form"):
    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "値": st.column_config.NumberColumn("値", help="数値で並べ替えやすくなります"),
            "日時": st.column_config.TextColumn("日時", disabled=True, help="自動入力（編集不可）"),
        },
        key="data_editor",
    )
    save_edits = st.form_submit_button("編集を保存")

if save_edits:
    # 現在表示の順序で保存するかどうか
    if ss["save_in_sorted_order"]:
        st.session_state.df = edited_df.copy()
    else:
        # 表示順を採用しない場合（インデックス順へ）
        st.session_state.df = edited_df.sort_index().copy()

    if worksheet:
        try:
            worksheet.clear()
            values = [COLUMNS] + st.session_state.df[COLUMNS].fillna("").astype(str).values.tolist()
            worksheet.update(values)
            st.success("編集内容を保存しました。")
            st.session_state.last_saved_df = st.session_state.df.copy()
        except Exception as e:
            st.error(f"Sheetsへの編集保存に失敗: {e}")
    else:
        st.error("ワークシート未接続のため保存できません。")

# ===================== Excel ダウンロード（現在の表示内容） =====================
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "記録"
    ws.append(COLUMNS)
    for row in df[COLUMNS].fillna("").itertuples(index=False, name=None):
        ws.append(list(row))
    wb.save(output)
    return output.getvalue()

if not st.session_state.df.empty:
    st.download_button(
        label="📄 Excelダウンロード",
        data=to_excel_bytes(st.session_state.df),
        file_name="汎用スケール記録.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
