# -*- coding: utf-8 -*-
"""
app.py
=======================================================================
Web App Streamlit: Hệ thống kiểm tra lỗi BHYT (XML3 / Excel)
Bệnh viện Đa khoa Anh Quất.
=======================================================================
"""

import os
import sys
import json
import base64
import pickle
import unicodedata
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode, JsCode
    HAS_AGGRID = True
except ImportError:
    HAS_AGGRID = False


# --- Đảm bảo import được checker.py dù chạy từ đâu ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from checker import (  # noqa: E402
    DK_COLUMNS,
    DK_TITLES,
    build_template_excel_bytes,
    build_output_excel_bytes,
    make_output_filename,
    load_reference_data,
    read_input_dataframe,
    run_check,
)
from bed_manager import (  # noqa: E402
    page_bed_management,
    extract_bed_records,
    save_bed_dataset,
)

# ======================= CẤU HÌNH CHUNG =======================
DATA_REF_DIR = os.path.join(BASE_DIR, 'data_ref')
os.makedirs(DATA_REF_DIR, exist_ok=True)
ACTIVE_REF_FILE = os.path.join(DATA_REF_DIR, '_active_ref.json')
LATEST_RESULT_FILE = os.path.join(DATA_REF_DIR, '_latest_result.pkl')
TH_NOTES_FILE = os.path.join(DATA_REF_DIR, '_th_notes.json')
LOGO_PATH = os.path.join(BASE_DIR, 'Logo.jpg')

COLOR_MAIN = "#1565C0"
DK_COLORS = {
    1: "#D32F2F", 2: "#E64A19", 3: "#F57C00",
    4: "#7B1FA2", 5: "#0288D1", 6: "#00796B",
}

PAGE_SCAN = "🔍 Soi lỗi XML3 (6 ĐK)"
PAGE_BED = "🛏️ Sơ đồ & Kiểm tra Giường"
PAGE_CATALOG = "📂 Quản lý Danh mục Cổng"
MAX_ALL_ROWS = 2000
OPT_ALL_ROWS = f"Tất cả (≤{MAX_ALL_ROWS:,})"

# Tỷ lệ độ rộng 5 ô lọc + 1 nút Đồng bộ (anh có thể tự chỉnh to/nhỏ ở đây)
FILTER_COL_WIDTHS = [2.2, 1.6, 1.6, 1.1, 1.2, 0.45]

st.set_page_config(
    layout="wide",
    page_title="Kiểm Tra Lỗi BHYT - BVĐK Anh Quất",
    page_icon="🩺",
    initial_sidebar_state="expanded",
)


# ======================= CUSTOM CSS =======================
def inject_css():
    st.markdown(
        """
        <style>
        html, body, .stApp {
            font-family: 'Segoe UI','Roboto',sans-serif;
        }
        .stApp { background-color: #F8FAFC; }
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header[data-testid="stHeader"] { background: transparent; }
        .block-container {
            padding-top: 0.55rem;
            padding-bottom: 1.0rem;
            padding-left: 1.1rem !important;
            padding-right: 1.1rem !important;
            max-width: 100% !important;
        }

        /* ---------- Header mảnh ---------- */
        .app-header {
            display: flex; align-items: center; gap: 14px;
            background: linear-gradient(100deg,#1565C0 0%,#1E88E5 55%,#42A5F5 100%);
            border-radius: 10px; padding: 7px 16px; margin-bottom: 6px;
            box-shadow: 0 4px 12px rgba(21,101,192,0.18);
        }
        .app-logo {
            height: 40px; width: 40px; border-radius: 8px; object-fit: cover;
            background: #FFFFFF; padding: 2px;
        }
        .app-title {
            color: #FFFFFF; font-size: 16.5px; font-weight: 700;
            letter-spacing: .3px; line-height: 1.2;
        }
        .app-sub {
            color: #E3F2FD; font-size: 11px; font-weight: 600;
            letter-spacing: .5px; margin-top: 1px;
        }

        /* ---------- Section ---------- */
        .section-title {
            font-size: 14.5px; font-weight: 700; color: #0F172A;
            margin: 4px 0 8px 0;
        }
        .soft-note { color: #64748B; font-size: 12px; white-space: nowrap; }

        /* ---------- Card ---------- */
        .card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 14px 16px; box-shadow: 0 1px 3px rgba(15,23,42,.04);
        }

        /* ---------- KPI mảnh ---------- */
        .kpi-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 9px;
            padding: 6px 6px; text-align: center; margin-bottom: 4px;
            box-shadow: 0 1px 2px rgba(15,23,42,.04);
        }
        .kpi-value { font-size: 17px; font-weight: 700; line-height: 1.1; }
        .kpi-label {
            font-size: 10px; color: #64748B; font-weight: 600; margin-top: 2px;
            text-transform: uppercase; letter-spacing: .2px;
        }

        /* ---------- Sidebar ---------- */
        section[data-testid="stSidebar"] {
            background: #FFFFFF; border-right: 1px solid #E2E8F0;
        }

        /* ---------- Buttons ---------- */
        .stButton > button, .stDownloadButton > button {
            border-radius: 7px; font-weight: 600; padding: 0.25rem 0.5rem;
            min-height: 34px; font-size: 12.5px;
        }
        .stButton > button[kind="primary"] { background: #1565C0; border-color: #1565C0; }

        div[data-testid="stVerticalBlock"] > div:has(iframe) { margin-bottom: -14px !important; }
        hr { margin: 6px 0 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_close_dropdown_on_outside_click():
    """Tự đóng menu xổ xuống của multiselect/selectbox khi click ra vùng ngoài."""
    components.html(
        """
        <script>
        (function() {
            const doc = window.parent.document;
            if (doc.__closeDropdownBound) return;
            doc.__closeDropdownBound = true;

            doc.addEventListener('mousedown', function(e) {
                const popover = doc.querySelector('[data-baseweb="popover"]');
                if (!popover) return;
                const insidePopover = popover.contains(e.target);
                const insideControl = e.target.closest('[data-baseweb="select"]');
                if (!insidePopover && !insideControl) {
                    const activeInput = doc.activeElement;
                    if (activeInput) activeInput.blur();
                    doc.dispatchEvent(new KeyboardEvent('keydown', {
                        key: 'Escape',
                        code: 'Escape',
                        keyCode: 27,
                        which: 27,
                        bubbles: true
                    }));
                }
            }, true);
        })();
        </script>
        """,
        height=0,
        width=0,
    )


def render_header():
    logo_html = ""
    if os.path.exists(LOGO_PATH):
        try:
            with open(LOGO_PATH, 'rb') as fh:
                b64 = base64.b64encode(fh.read()).decode()
            logo_html = f'<img src="data:image/jpeg;base64,{b64}" class="app-logo"/>'
        except Exception:
            logo_html = ""

    # Lấy thông tin thời điểm nạp & khoảng ngày dữ liệu
    updated_at = st.session_state.get('updated_at', '')
    data_range = st.session_state.get('data_range', '')
    if not updated_at or not data_range:
        saved = load_latest_result()
        if saved:
            updated_at = updated_at or saved.get('updated_at', '')
            data_range = data_range or saved.get('data_range', '')

    meta_html = ""
    if updated_at or data_range:
        parts = []
        if updated_at:
            parts.append(f"🕒 Nạp lúc: <b>{updated_at}</b>")
        if data_range:
            parts.append(f"📅 Dữ liệu: <b>{data_range}</b>")
        meta_html = (
            "<div style='margin-left:auto;background:rgba(255,255,255,0.16);border:1px solid rgba(255,255,255,0.28);"
            "border-radius:8px;padding:6px 12px;color:#FFFFFF;font-size:12.5px;display:flex;gap:14px;align-items:center;'>"
            + " &nbsp;|&nbsp; ".join(parts)
            + "</div>"
        )

    st.markdown(
        f"""
        <div class="app-header">
            {logo_html}
            <div class="app-header-text">
                <div class="app-title">BỆNH VIỆN ĐA KHOA ANH QUẤT</div>
                <div class="app-sub">HỆ THỐNG KIỂM TRA LỖI BHYT (XML3) &amp; QUẢN LÝ GIƯỜNG BỆNH</div>
            </div>
            {meta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_top_channel_switcher():
    """2 nút chuyển kênh chính ngay dưới banner Bệnh viện."""
    cur = st.session_state.get("active_page", PAGE_SCAN)
    c1, c2 = st.columns(2, gap="small")
    with c1:
        if st.button(
            PAGE_SCAN,
            type="primary" if cur == PAGE_SCAN else "secondary",
            use_container_width=True,
            key="nav_btn_scan",
        ):
            st.session_state["active_page"] = PAGE_SCAN
            st.rerun()
    with c2:
        if st.button(
            PAGE_BED,
            type="primary" if cur == PAGE_BED else "secondary",
            use_container_width=True,
            key="nav_btn_bed",
        ):
            st.session_state["active_page"] = PAGE_BED
            st.rerun()


def render_kpi_row(items):
    cols = st.columns(len(items))
    for col, (label, value, color) in zip(cols, items):
        col.markdown(
            f"""<div class="kpi-card" style="border-top:3px solid {color};">
                    <div class="kpi-value" style="color:{color};">{value}</div>
                    <div class="kpi-label">{label}</div>
                </div>""",
            unsafe_allow_html=True,
        )


def render_excel_table(df, key, height=820, editable_cols=None):
    editable_cols = set(editable_cols or [])
    if not HAS_AGGRID:
        st.dataframe(df, use_container_width=True, hide_index=True, height=height)
        return df

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        filter=True,
        floatingFilter=True,
        sortable=True,
        resizable=True,
        minWidth=95,
        filterParams={"defaultOption": "contains"},
    )

    for narrow_col, w in [
        ("STT", 60), ("Dòng", 70), ("Id", 70), ("MA_NHOM", 90),
        ("Tên bệnh nhân", 160), ("Khoa", 150), ("Trạng thái", 115),
    ]:
        if narrow_col in df.columns:
            gb.configure_column(narrow_col, width=w, minWidth=w)
    for hidden_col in ["Dòng", "Mã DV", "Mã DV_check"]:
        if hidden_col in df.columns:
            gb.configure_column(hidden_col, hide=True)
    if "Đã sửa" in editable_cols and "Đã sửa" in df.columns:
        gb.configure_column("Đã sửa", editable=True, width=82, minWidth=82, cellStyle={"textAlign": "center"})
    if "Ghi chú" in editable_cols and "Ghi chú" in df.columns:
        gb.configure_column("Ghi chú", editable=True, width=180, minWidth=130)
    if "Chi tiết lỗi" in df.columns:
        max_len = int(df["Chi tiết lỗi"].astype(str).str.len().max() or 30)
        w_err = min(max(max_len * 8 + 40, 380), 900)
        gb.configure_column("Chi tiết lỗi", width=w_err, minWidth=w_err, tooltipField="Chi tiết lỗi")
    gb.configure_column("::auto_unique_id::", hide=True)
    red_check_style = JsCode("function(params) { return {'color': '#d92d20', 'fontWeight': '500'}; }")

    for col in df.columns:
        if "check" in str(col).lower():
            gb.configure_column(col, cellStyle=red_check_style)
    gb.configure_grid_options(
        headerHeight=32,
        floatingFiltersHeight=30,
        rowHeight=36,
        singleClickEdit=True,
        stopEditingWhenCellsLoseFocus=True,
        enableCellTextSelection=True,
        ensureDomOrder=True,
    )

    custom_css = {
        ".ag-header": {"background-color": "#F8FAFC !important", "border-bottom": "1px solid #CBD5E1 !important"},
        ".ag-header-cell": {
            "border-right": "1px solid #E2E8F0 !important", "font-size": "12.5px !important",
            "font-weight": "600 !important", "color": "#334155 !important",
            "padding-left": "8px !important", "padding-right": "8px !important",
        },
        ".ag-cell": {
            "border-right": "1px solid #E2E8F0 !important", "font-size": "13px !important",
            "padding-left": "8px !important", "padding-right": "8px !important", "line-height": "34px !important",
        },
        ".ag-row": {"border-bottom": "1px solid #E2E8F0 !important"},
        ".ag-floating-filter-input input": {
            "height": "22px !important", "min-height": "22px !important",
            "font-size": "12px !important", "padding": "0 6px !important", "border-radius": "4px !important",
        },
        ".ag-floating-filter-button": {"margin-left": "2px !important"},
        ".ag-floating-filter-button-button": {"width": "18px !important", "height": "18px !important"},
    }

    upd_mode = GridUpdateMode.FILTERING_CHANGED
    if editable_cols:
        upd_mode = GridUpdateMode.VALUE_CHANGED | GridUpdateMode.FILTERING_CHANGED

    resp = AgGrid(
        df,
        gridOptions=gb.build(),
        custom_css=custom_css,
        height=height,
        theme="alpine",
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=upd_mode,
        key=key,
        allow_unsafe_jscode=True,
    )
    filtered = resp.get("data")
    if filtered is not None:
        out_df = pd.DataFrame(filtered)
        return out_df.drop(columns=["::auto_unique_id::"], errors="ignore")
    return df


# ======================= TIỆN ÍCH FILE / STATE / PHÂN QUYỀN =======================
def load_active_ref():
    if os.path.exists(ACTIVE_REF_FILE):
        try:
            with open(ACTIVE_REF_FILE, 'r', encoding='utf-8') as fh:
                name = json.load(fh).get('active_ref', '')
            path = os.path.join(DATA_REF_DIR, name)
            if name and os.path.exists(path):
                return path
        except Exception:
            return None
    return None


def set_active_ref(filename):
    try:
        with open(ACTIVE_REF_FILE, 'w', encoding='utf-8') as fh:
            json.dump({'active_ref': filename}, fh, ensure_ascii=False)
    except Exception:
        pass


def is_admin():
    try:
        host = st.context.headers.get("Host", "").lower()
        return host.startswith("localhost") or host.startswith("127.0.0.1")
    except Exception:
        return False


def load_th_notes():
    if os.path.exists(TH_NOTES_FILE):
        try:
            with open(TH_NOTES_FILE, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def save_th_notes(notes):
    try:
        tmp_path = TH_NOTES_FILE + ".tmp"
        with open(tmp_path, 'w', encoding='utf-8') as fh:
            json.dump(notes, fh, ensure_ascii=False, indent=2)
        os.replace(tmp_path, TH_NOTES_FILE)
    except Exception:
        pass


def reset_th_notes():
    save_th_notes({})


def clear_cached_exports():
    st.session_state.pop('_excel_full_bytes', None)
    st.session_state.pop('_excel_filtered_bytes', None)
    st.session_state.pop('_excel_filtered_sig', None)


def attach_notes_to_th(df_th):
    if df_th.empty:
        df = df_th.copy()
        df["Đã sửa"] = pd.Series(dtype=bool)
        df["Ghi chú"] = pd.Series(dtype=str)
        return df

    notes = load_th_notes()
    df = df_th.copy()
    da_sua_list, ghi_chu_list = [], []
    for stt in df["STT"]:
        rec = notes.get(str(stt), {})
        da_sua_list.append(bool(rec.get("da_sua", False)))
        ghi_chu_list.append(str(rec.get("ghi_chu", "")))

    df["Đã sửa"] = da_sua_list
    df["Ghi chú"] = ghi_chu_list
    other_cols = [c for c in df.columns if c not in ("STT", "Đã sửa", "Ghi chú")]
    return df[["STT", "Đã sửa", "Ghi chú"] + other_cols]


def sync_edited_notes(page_before_df, page_after_df):
    if page_before_df.empty or page_after_df.empty:
        return False
    if "STT" not in page_after_df.columns or "Đã sửa" not in page_after_df.columns:
        return False

    before_map = {}
    for _, r in page_before_df.iterrows():
        before_map[str(r["STT"])] = (
            bool(r.get("Đã sửa", False)),
            str(r.get("Ghi chú", "") or "").strip(),
        )

    changes = {}
    for _, r in page_after_df.iterrows():
        k = str(r["STT"])
        new_val = (
            bool(r.get("Đã sửa", False)),
            str(r.get("Ghi chú", "") or "").strip(),
        )
        if k in before_map and before_map[k] != new_val:
            changes[k] = {"da_sua": new_val[0], "ghi_chu": new_val[1]}

    if not changes:
        return False

    notes = load_th_notes()
    for k, val in changes.items():
        if val["da_sua"] or val["ghi_chu"]:
            notes[k] = val
        else:
            notes.pop(k, None)
    save_th_notes(notes)
    clear_cached_exports()
    return True


def format_th_for_export(df_th):
    df = df_th.copy()
    if "Đã sửa" in df.columns:
        df["Đã sửa"] = df["Đã sửa"].apply(lambda x: "Đã sửa" if bool(x) else "")
    return df


def save_latest_result(result, src_filename, used_ref, updated_at="", data_range=""):
    try:
        with open(LATEST_RESULT_FILE, 'wb') as fh:
            pickle.dump({
                'result': result,
                'src_filename': src_filename,
                'used_ref': used_ref,
                'updated_at': updated_at,
                'data_range': data_range,
            }, fh)
    except Exception:
        pass


def load_latest_result():
    if os.path.exists(LATEST_RESULT_FILE):
        try:
            with open(LATEST_RESULT_FILE, 'rb') as fh:
                return pickle.load(fh)
        except Exception:
            return None
    return None


@st.cache_data(show_spinner=False)
def _load_catalogs_cached(ref_path, ref_mtime):
    return load_reference_data(ref_path)


def get_catalogs(ref_path):
    return _load_catalogs_cached(ref_path, os.path.getmtime(ref_path))


@st.cache_data(show_spinner=False)
def _load_visible_sheets_cached(ref_path, ref_mtime):
    """Đọc toàn bộ các sheet ĐANG HIỆN (visible) trong file Excel tham chiếu."""
    xl = pd.ExcelFile(ref_path)
    visible_names = list(xl.sheet_names)
    if ref_path.lower().endswith(('.xlsx', '.xlsm')):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(ref_path, read_only=True, data_only=True)
            visible_names = [name for name in wb.sheetnames if wb[name].sheet_state == 'visible']
            wb.close()
        except Exception:
            pass
    sheets = {}
    for name in visible_names:
        try:
            sheets[name] = xl.parse(name, dtype=str).fillna("")
        except Exception:
            continue
    return sheets


def get_visible_sheets(ref_path):
    return _load_visible_sheets_cached(ref_path, os.path.getmtime(ref_path))


def _strip_accents(text):
    s = unicodedata.normalize('NFD', str(text).lower())
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.replace('đ', 'd')


def match_query(df, query, columns):
    if not query or df.empty:
        return df
    tokens = [t for t in _strip_accents(query).split() if t]
    if not tokens:
        return df
    cols = [c for c in columns if c in df.columns]
    if not cols:
        return df
    combined = df[cols].fillna("").astype(str).agg(" ".join, axis=1).map(_strip_accents)
    mask = pd.Series(True, index=df.index)
    for tok in tokens:
        mask = mask & combined.str.contains(tok, na=False, regex=False)
    return df[mask]


# ======================= SIDEBAR (NẠP DỮ LIỆU CHUNG + DANH MỤC CỔNG) =======================
def build_sidebar(admin=True):
    if "active_page" not in st.session_state:
        st.session_state["active_page"] = PAGE_SCAN

    with st.sidebar:
        st.markdown("### 🩺 BHYT Anh Quất")
        st.markdown("<div class='soft-note'>Kiểm tra XML3 &amp; Giường bệnh</div>", unsafe_allow_html=True)
        st.markdown("---")

        active_ref = load_active_ref()
        ref_path = active_ref

        if admin:
            st.markdown("**📤 1. Nạp hồ sơ kiểm tra (Excel / XML)**")
            input_file = st.file_uploader(
                "Chọn file Excel (.xlsx, .xls) hoặc XML (.xml)",
                type=["xlsx", "xls", "xml"],
                key="sidebar_data_uploader",
                label_visibility="collapsed",
            )
            run_btn = st.button("🚀 Quét lỗi & Nạp Giường", type="primary", use_container_width=True)

            if run_btn:
                if input_file is None:
                    st.warning("Vui lòng chọn file dữ liệu trước khi quét.")
                else:
                    progress = st.progress(0.0, text="Đang đọc file...")

                    def cb(pct, msg):
                        progress.progress(min(max(int(pct * 100), 0), 100), text=msg)

                    try:
                        df_raw = read_input_dataframe(input_file.getvalue(), input_file.name)
                        catalogs = {}
                        if ref_path and os.path.exists(ref_path):
                            catalogs, _ = get_catalogs(ref_path)

                        result = run_check(df_raw, catalogs=catalogs, ref_path=ref_path or "", progress_cb=cb)
                        st.session_state['result'] = result
                        st.session_state['src_filename'] = input_file.name
                        st.session_state['used_ref'] = bool(ref_path and os.path.exists(ref_path))
                        st.session_state['th_page'] = 0
                        st.session_state['ct_page'] = 0
                        clear_cached_exports()
                        df_bed_new = extract_bed_records(df_raw)
                        now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
                        if not df_bed_new.empty:
                            d_min = df_bed_new["START_DT"].min().strftime("%d/%m/%Y")
                            d_max = df_bed_new["END_DT"].max().strftime("%d/%m/%Y")
                            range_str = f"{d_min} → {d_max}"
                            save_bed_dataset(df_bed_new, input_file.name)
                            st.session_state.pop("bed_applied_filter", None)
                        else:
                            range_str = ""

                        st.session_state['updated_at'] = now_str
                        st.session_state['data_range'] = range_str
                        save_latest_result(
                            result, input_file.name,
                            bool(ref_path and os.path.exists(ref_path)),
                            updated_at=now_str,
                            data_range=range_str,
                        )
                        reset_th_notes()
                        progress.empty()
                        st.toast(
                            f"✅ Xong {input_file.name}: {result['stats']['total_errors']:,} lỗi XML3 • {len(df_bed_new):,} dòng giường",
                            icon="🎉",
                        )
                        st.rerun()
                    except Exception as ex:
                        progress.empty()
                        st.error(f"❌ Lỗi xử lý: {ex}")

            cur_src = st.session_state.get('src_filename', '')
            if not cur_src:
                saved = load_latest_result()
                cur_src = saved.get('src_filename', '') if saved else ''
            if cur_src:
                st.caption(f"📄 Đang dùng: **{cur_src}**")

            st.markdown("---")

            st.markdown("**⚙️ 2. Danh mục đối chiếu (Cổng BHYT)**")
            uploaded_ref = st.file_uploader(
                "File tham chiếu (.xlsx, .xls, .ods)",
                type=["xlsx", "xls", "ods"],
                key="ref_uploader",
                label_visibility="collapsed",
            )
            if uploaded_ref is not None:
                safe_name = os.path.basename(uploaded_ref.name)
                save_path = os.path.join(DATA_REF_DIR, safe_name)
                try:
                    with open(save_path, 'wb') as fh:
                        fh.write(uploaded_ref.getbuffer())
                    set_active_ref(safe_name)
                    ref_path = save_path
                    st.session_state.pop('_catalog_cache_key', None)
                except Exception as ex:
                    st.error(f"Không lưu được file: {ex}")

            if ref_path and os.path.exists(ref_path):
                st.success(f"✅ {os.path.basename(ref_path)}")
                c_dl, c_rm = st.columns(2, gap="small")
                with c_dl:
                    st.download_button(
                        "📥 File mẫu",
                        data=build_template_excel_bytes(),
                        file_name="MAU_THAM_CHIEU.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with c_rm:
                    if st.button("🔄 Đổi file", use_container_width=True):
                        set_active_ref("")
                        st.rerun()
            else:
                st.warning("⚠️ Chưa nạp file tham chiếu.")
                st.download_button(
                    "📥 Tải file mẫu (MAU_THAM_CHIEU.xlsx)",
                    data=build_template_excel_bytes(),
                    file_name="MAU_THAM_CHIEU.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.info(
                "👨‍⚕️ **Chế độ Bác sĩ / Khoa phòng**\n\n"
                "Lọc khoa/tên mình, tích **Đã sửa** hoặc gõ **Ghi chú** — hệ thống tự lưu về máy chủ."
            )

        st.markdown("---")
        is_cat = st.session_state.get("active_page") == PAGE_CATALOG
        if st.button(
            PAGE_CATALOG,
            type="primary" if is_cat else "secondary",
            use_container_width=True,
            key="sidebar_btn_catalog",
        ):
            st.session_state["active_page"] = PAGE_CATALOG
            st.rerun()

        st.markdown("---")
        role_label = "Quản trị" if admin else "Xem & Đánh dấu"
        st.markdown(
            f"<div class='soft-note'>Phiên bản Web • {role_label} • Engine: checker.py</div>",
            unsafe_allow_html=True,
        )

    return st.session_state["active_page"], (ref_path if (ref_path and os.path.exists(ref_path)) else None)


# ======================= TRANG 1: SOI LỖI XML3 =======================
def page_scan(ref_path, admin=True):
    result = st.session_state.get('result')
    src_filename = st.session_state.get('src_filename', '')
    used_ref = st.session_state.get('used_ref', False)

    if not result:
        saved = load_latest_result()
        if saved:
            result = saved['result']
            src_filename = saved.get('src_filename', '')
            used_ref = saved.get('used_ref', False)

    if not result:
        msg = (
            "Chưa có dữ liệu kiểm tra. Vui lòng chờ quản trị viên cập nhật hồ sơ."
            if not admin else
            "Chưa có kết quả. Hãy nạp file Excel/XML ở cột bên trái và bấm <b>🚀 Quét lỗi &amp; Nạp Giường</b>."
        )
        st.markdown(f"<div class='card'><span class='soft-note'>{msg}</span></div>", unsafe_allow_html=True)
        return

    render_results(result, src_filename, used_ref)


# ======================= HIỂN THỊ & LỌC KẾT QUẢ =======================
SEARCH_COLS_TH = [
    'Dòng', 'Tên bệnh nhân', 'Khoa', 'Trạng thái',
    'Mã DV', 'Tên dịch vụ', 'Chi tiết lỗi', 'Ghi chú',
    'Mã DV_check', 'Tên dịch vụ_check',
]
SEARCH_COLS_CT = [
    'MA_LK', 'Tên bệnh nhân', 'Khoa', 'Trạng thái',
    'MA_DICH_VU', 'TEN_DICH_VU', 'NGUOI_THUC_HIEN', 'MA_MAY',
] + DK_COLUMNS


def filter_th(df_th, selected_dk, selected_khoa, selected_status, selected_fix, query):
    df = df_th
    if selected_dk:
        df = df[df['Loại lỗi'].isin(selected_dk)]
    if selected_khoa and 'Khoa' in df.columns:
        df = df[df['Khoa'].isin(selected_khoa)]
    if selected_status != "Tất cả trạng thái" and 'Trạng thái' in df.columns:
        df = df[df['Trạng thái'] == selected_status]
    if selected_fix == "Chưa sửa" and 'Đã sửa' in df.columns:
        df = df[~df['Đã sửa'].astype(bool)]
    elif selected_fix == "Đã sửa" and 'Đã sửa' in df.columns:
        df = df[df['Đã sửa'].astype(bool)]
    return match_query(df, query, SEARCH_COLS_TH)


def filter_chitiet(df_out, selected_dk, selected_khoa, selected_status, query, only_errors):
    df = df_out
    if only_errors:
        mask_err = pd.Series(False, index=df.index)
        for col in DK_COLUMNS:
            if col in df.columns:
                mask_err = mask_err | (df[col].astype(str).str.strip() != "")
        df = df[mask_err]
    if selected_dk:
        mask = pd.Series(False, index=df.index)
        for col in selected_dk:
            if col in df.columns:
                mask = mask | (df[col].astype(str).str.strip() != "")
        df = df[mask]
    if selected_khoa and 'Khoa' in df.columns:
        df = df[df['Khoa'].isin(selected_khoa)]
    if selected_status != "Tất cả trạng thái" and 'Trạng thái' in df.columns:
        df = df[df['Trạng thái'] == selected_status]
    return match_query(df, query, SEARCH_COLS_CT)


def render_results(result, src_filename, used_ref):
    df_out = result['df_out']
    df_th_with_notes = attach_notes_to_th(result['df_th'])
    stats = result['stats']

    # ----- Dải KPI mảnh -----
    kpi_items = [
        ("Tổng dòng quét", f"{stats['total_rows']:,}", "#1565C0"),
        ("Tổng lỗi", f"{stats['total_errors']:,}", "#C62828"),
    ]
    for i, col in enumerate(DK_COLUMNS, start=1):
        kpi_items.append((f"ĐK {i}", f"{stats['by_dk'][col]:,}", DK_COLORS[i]))
    render_kpi_row(kpi_items)

    # ----- 1 HÀNG DUY NHẤT: 5 Ô LỌC + 1 NÚT ĐỒNG BỘ GỌN -----
    khoa_options = sorted(
        [k for k in df_out['Khoa'].dropna().astype(str).str.strip().unique() if k]
    ) if 'Khoa' in df_out.columns else []

    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns(FILTER_COL_WIDTHS, gap="small", vertical_alignment="center")
    with fc1:
        query = st.text_input(
            "Tìm kiếm",
            placeholder="🔎 Tìm Tên BN, BS, Mã DV, Máy...",
            key="search_query",
            label_visibility="collapsed",
        )
    with fc2:
        selected_khoa = st.multiselect(
            "Khoa",
            options=khoa_options,
            default=[],
            placeholder="🏥 Lọc Khoa...",
            key="khoa_filter",
            label_visibility="collapsed",
        )
    with fc3:
        selected_dk = st.multiselect(
            "Điều kiện",
            options=DK_COLUMNS,
            default=[],
            placeholder="🎯 Lọc ĐK 1→6...",
            key="dk_filter",
            label_visibility="collapsed",
        )
    with fc4:
        selected_status = st.selectbox(
            "Trạng thái",
            options=["Tất cả trạng thái", "Đang điều trị", "Ra viện"],
            key="status_filter",
            label_visibility="collapsed",
        )
    with fc5:
        selected_fix = st.selectbox(
            "Tình trạng sửa",
            options=["Tất cả (Đã/Chưa sửa)", "Chưa sửa", "Đã sửa"],
            key="fix_filter",
            label_visibility="collapsed",
        )
    with fc6:
        if st.button("🔄", use_container_width=True, help="Đồng bộ trạng thái Đã sửa / Ghi chú mới nhất"):
            clear_cached_exports()
            st.rerun()

    # ----- Xác định tab đang chọn & tính dữ liệu lọc -----
    if "scan_subtab" not in st.session_state:
        st.session_state["scan_subtab"] = "th"
    cur_sub = st.session_state["scan_subtab"]

    df_th_view = filter_th(df_th_with_notes, selected_dk, selected_khoa, selected_status, selected_fix, query)
    only_errors = st.session_state.get("only_errors", True)
    df_ct_view = filter_chitiet(df_out, selected_dk, selected_khoa, selected_status, query, only_errors)

    # Xóa cache Excel lọc nếu bộ lọc thay đổi
    filter_sig = (tuple(selected_dk), tuple(selected_khoa), selected_status, selected_fix, query, only_errors)
    if st.session_state.get('_excel_filtered_sig') != filter_sig:
        st.session_state.pop('_excel_filtered_bytes', None)
        st.session_state['_excel_filtered_sig'] = filter_sig

    # Tính phân trang cho tab đang chọn
    active_df = df_th_view if cur_sub == "th" else df_ct_view
    page_key = "th_page" if cur_sub == "th" else "ct_page"
    size_key = "th_pgsz" if cur_sub == "th" else "ct_pgsz"
    size_options = [100, 200, 500, OPT_ALL_ROWS]
    if size_key not in st.session_state or st.session_state[size_key] not in size_options:
        st.session_state[size_key] = 100
    if page_key not in st.session_state:
        st.session_state[page_key] = 0

    total_rows = len(active_df)
    selected_sz = st.session_state[size_key]
    if selected_sz == OPT_ALL_ROWS:
        page_size = max(1, min(total_rows, MAX_ALL_ROWS))
    else:
        page_size = int(selected_sz)

    n_pages = max(1, (total_rows + page_size - 1) // page_size)
    page = min(max(0, int(st.session_state[page_key])), n_pages - 1)
    st.session_state[page_key] = page
    start = page * page_size
    end = min(start + page_size, total_rows)
    range_txt = f"{start + 1:,}–{end:,}" if total_rows > 0 else "0"

    done_count = int(df_th_with_notes["Đã sửa"].sum()) if not df_th_with_notes.empty else 0
    full_name = make_output_filename(src_filename)

    # ----- THANH CÔNG CỤ TẤT CẢ TRONG 1 HÀNG: [TH | Chi tiết] + [Hiển thị...] + [Phân trang] + [Excel Full | Lọc] -----
    tb_th, tb_ct, tb_info, tb_sz, tb_nav, tb_ex_full, tb_ex_flt = st.columns(
        [1.05, 0.95, 2.65, 0.80, 1.50, 0.70, 0.70],
        gap="small",
        vertical_alignment="center",
    )

    with tb_th:
        if st.button(
            "📋 Tổng hợp (TH)",
            type="primary" if cur_sub == "th" else "secondary",
            use_container_width=True,
            key="subtab_btn_th",
        ):
            st.session_state["scan_subtab"] = "th"
            st.rerun()

    with tb_ct:
        if st.button(
            "📑 Chi tiết",
            type="primary" if cur_sub == "ct" else "secondary",
            use_container_width=True,
            key="subtab_btn_ct",
        ):
            st.session_state["scan_subtab"] = "ct"
            st.rerun()

    with tb_info:
        if cur_sub == "th":
            st.markdown(
                f"<div class='soft-note'>Hiển thị <b>{range_txt}</b>/{total_rows:,} • "
                f"✅ Đã sửa: <b>{done_count:,}</b>/{len(df_th_with_notes):,}</div>",
                unsafe_allow_html=True,
            )
        else:
            c_i1, c_i2 = st.columns([1.2, 1.1], vertical_alignment="center")
            with c_i1:
                st.markdown(
                    f"<div class='soft-note'><b>{range_txt}</b>/{total_rows:,} (gốc {len(df_out):,})</div>",
                    unsafe_allow_html=True,
                )
            with c_i2:
                st.checkbox("Chỉ dòng lỗi", value=only_errors, key="only_errors")

    with tb_sz:
        st.selectbox("Dòng/trang", options=size_options, key=size_key, label_visibility="collapsed")

    with tb_nav:
        n1, n2, n3, n4, n5 = st.columns([0.55, 0.55, 1.1, 0.55, 0.55], vertical_alignment="center")
        if n1.button("«", key=f"{page_key}_first", disabled=(page <= 0), use_container_width=True):
            st.session_state[page_key] = 0
            st.rerun()
        if n2.button("‹", key=f"{page_key}_prev", disabled=(page <= 0), use_container_width=True):
            st.session_state[page_key] = page - 1
            st.rerun()
        n3.markdown(
            f"<div style='text-align:center;font-size:12px;font-weight:600;color:#334155;'>{page + 1}/{n_pages}</div>",
            unsafe_allow_html=True,
        )
        if n4.button("›", key=f"{page_key}_next", disabled=(page >= n_pages - 1), use_container_width=True):
            st.session_state[page_key] = page + 1
            st.rerun()
        if n5.button("»", key=f"{page_key}_last", disabled=(page >= n_pages - 1), use_container_width=True):
            st.session_state[page_key] = n_pages - 1
            st.rerun()

    with tb_ex_full:
        if st.session_state.get('_excel_full_bytes'):
            st.download_button(
                "⬇️ Full",
                data=st.session_state['_excel_full_bytes'],
                file_name=full_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Bấm để tải Excel đầy đủ (TH + chitiet)",
            )
        else:
            if st.button("📥 Full", use_container_width=True, help="Tạo file Excel đầy đủ (TH + chitiet)"):
                with st.spinner("Đang tạo Excel Full..."):
                    latest_th = attach_notes_to_th(result['df_th'])
                    st.session_state['_excel_full_bytes'] = build_output_excel_bytes(
                        format_th_for_export(latest_th), df_out
                    )
                st.rerun()

    with tb_ex_flt:
        if st.session_state.get('_excel_filtered_bytes'):
            st.download_button(
                "⬇️ Lọc",
                data=st.session_state['_excel_filtered_bytes'],
                file_name=full_name.replace(".xlsx", "_Loc.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Bấm để tải Excel theo bộ lọc hiện tại",
            )
        else:
            if st.button("📥 Lọc", use_container_width=True, help="Tạo file Excel theo bộ lọc hiện tại"):
                with st.spinner("Đang tạo Excel Lọc..."):
                    latest_th = attach_notes_to_th(result['df_th'])
                    th_flt = filter_th(latest_th, selected_dk, selected_khoa, selected_status, selected_fix, query)
                    ct_flt = filter_chitiet(df_out, selected_dk, selected_khoa, selected_status, query, only_errors)
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        format_th_for_export(th_flt).to_excel(writer, sheet_name='TH', index=False)
                        ct_flt.to_excel(writer, sheet_name='chitiet', index=False)
                    buf.seek(0)
                    st.session_state['_excel_filtered_bytes'] = buf.getvalue()
                st.rerun()

    # ----- BẢNG DỮ LIỆU CHÍNH -----
    if cur_sub == "th":
        if df_th_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
        else:
            page_before = df_th_view.iloc[start:end].copy()
            page_after = render_excel_table(
                page_before,
                key=f"grid_th_{page}_{page_size}",
                height=820,
                editable_cols=["Đã sửa", "Ghi chú"],
            )
            if sync_edited_notes(page_before, page_after):
                st.toast("✅ Đã lưu trạng thái & ghi chú lên máy chủ!", icon="💾")
    else:
        if df_ct_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
        else:
            st.dataframe(df_ct_view.iloc[start:end], use_container_width=True, hide_index=True, height=660)

    if not used_ref:
        st.info("ℹ️ Lần quét này chưa dùng file tham chiếu — ĐK 5 & ĐK 6 (phần cấu hình) chưa được kiểm tra đầy đủ.")


# ======================= TRANG QUẢN LÝ DANH MỤC CỔNG =======================
def page_catalog(ref_path, admin=True):
    st.markdown("<div class='section-title'>📂 Quản lý Danh mục Cổng BHYT</div>", unsafe_allow_html=True)

    if not ref_path:
        msg = (
            "⚠️ Chưa có file tham chiếu. Hãy tải lên ở cột bên trái (mục **2. Danh mục đối chiếu**)."
            if admin else
            "⚠️ Chưa có dữ liệu danh mục tham chiếu. Vui lòng chờ quản trị viên cập nhật."
        )
        st.warning(msg)
        return

    st.success(f"✅ File đang dùng: **{os.path.basename(ref_path)}**")
    visible_sheets = get_visible_sheets(ref_path)

    if not visible_sheets:
        st.info("Không tìm thấy sheet nào đang hiển thị trong file tham chiếu.")
        return

    sheet_names = list(visible_sheets.keys())
    kpi_items = [(name, f"{len(df_s):,}", COLOR_MAIN) for name, df_s in visible_sheets.items()]
    render_kpi_row(kpi_items[:6])

    tabs = st.tabs([f"📄 {name}" for name in sheet_names])
    for tab, name in zip(tabs, sheet_names):
        with tab:
            df_sheet = visible_sheets[name]
            if df_sheet is None or df_sheet.empty:
                st.info(f"Sheet **{name}** chưa có dữ liệu.")
                continue
            kw = st.text_input(f"🔎 Tìm trong {name}", key=f"cat_search_{name}")
            view = match_query(df_sheet, kw, list(df_sheet.columns)) if kw else df_sheet
            st.caption(f"Hiển thị **{len(view):,}** / **{len(df_sheet):,}** dòng")
            st.dataframe(view, use_container_width=True, hide_index=True, height=520)


# ======================= MAIN =======================
def main():
    inject_css()
    inject_close_dropdown_on_outside_click()
    admin = is_admin()
    page, ref_path = build_sidebar(admin=admin)
    render_header()
    render_top_channel_switcher()

    if page == PAGE_SCAN:
        page_scan(ref_path, admin=admin)
    elif page == PAGE_BED:
        page_bed_management(ref_path, admin=admin, render_kpi_fn=render_kpi_row, show_uploader=False)
    else:
        page_catalog(ref_path, admin=admin)


if __name__ == "__main__":
    main()
