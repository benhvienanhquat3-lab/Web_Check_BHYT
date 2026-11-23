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
from bed_manager import page_bed_management

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
PAGE_CATALOG = "📂 Quản lý Danh mục Cổng"
PAGE_BED = "🛏️ Sơ đồ & Kiểm tra Giường"
MAX_ALL_ROWS = 2000
OPT_ALL_ROWS = f"Tất cả (≤{MAX_ALL_ROWS:,})"

st.set_page_config(
    layout="wide",
    page_title="Kiểm Tra Lỗi BHYT - BVĐK Anh Quất",
    page_icon="🩺",
    initial_sidebar_state="expanded",
)


# ======================= CUSTOM CSS (GỌN GÀNG - TỐI ƯU BẢNG) =======================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, .stApp {
            font-family: 'Inter','Segoe UI','Roboto',sans-serif;
        }
        .stApp { background-color: #F8FAFC; }
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header[data-testid="stHeader"] { background: transparent; }
        .block-container {
            padding-top: 0.65rem;
            padding-bottom: 1.5rem;
            padding-left: 1.2rem !important;
            padding-right: 1.2rem !important;
            max-width: 100% !important;
        }

        /* ---------- Header mảnh ---------- */
        .app-header {
            display: flex; align-items: center; gap: 14px;
            background: linear-gradient(100deg,#1565C0 0%,#1E88E5 55%,#42A5F5 100%);
            border-radius: 10px; padding: 8px 16px; margin-bottom: 10px;
            box-shadow: 0 4px 12px rgba(21,101,192,0.18);
        }
        .app-logo {
            height: 42px; width: 42px; border-radius: 8px; object-fit: cover;
            background: #FFFFFF; padding: 2px;
        }
        .app-title {
            color: #FFFFFF; font-size: 17px; font-weight: 700;
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
        .soft-note { color: #64748B; font-size: 12.5px; }

        /* ---------- Card ---------- */
        .card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 14px 16px; box-shadow: 0 1px 3px rgba(15,23,42,.04);
        }

        /* ---------- KPI mảnh ---------- */
        .kpi-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 9px;
            padding: 7px 6px; text-align: center; margin-bottom: 4px;
            box-shadow: 0 1px 2px rgba(15,23,42,.04);
        }
        .kpi-value { font-size: 18px; font-weight: 700; line-height: 1.1; }
        .kpi-label {
            font-size: 10px; color: #64748B; font-weight: 600; margin-top: 2px;
            text-transform: uppercase; letter-spacing: .2px;
        }

        /* ---------- Sidebar ---------- */
        section[data-testid="stSidebar"] {
            background: #FFFFFF; border-right: 1px solid #E2E8F0;
        }

        /* ---------- Buttons ---------- */
        .stButton > button, .stDownloadButton > button { border-radius: 8px; font-weight: 600; }
        .stButton > button[kind="primary"] { background: #1565C0; border-color: #1565C0; }

        /* Thu gọn khoảng trắng giữa bảng và nút tải Excel */
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
    st.markdown(
        f"""
        <div class="app-header">
            {logo_html}
            <div class="app-header-text">
                <div class="app-title">BỆNH VIỆN ĐA KHOA ANH QUẤT</div>
                <div class="app-sub">HỆ THỐNG KIỂM TRA LỖI BHYT (XML3)</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_row(items):
    """items: list of (label, value, color)."""
    cols = st.columns(len(items))
    for col, (label, value, color) in zip(cols, items):
        col.markdown(
            f"""<div class="kpi-card" style="border-top:3px solid {color};">
                    <div class="kpi-value" style="color:{color};">{value}</div>
                    <div class="kpi-label">{label}</div>
                </div>""",
            unsafe_allow_html=True,
        )


def render_excel_table(df, key, height=620, editable_cols=None):
    """Hiển thị 1 trang dữ liệu dạng bảng kẻ ô kiểu Excel (AgGrid)."""
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
        gb.configure_column(
            "Đã sửa",
            editable=True,
            width=82,
            minWidth=82,
            cellStyle={"textAlign": "center"},
        )
    if "Ghi chú" in editable_cols and "Ghi chú" in df.columns:
        gb.configure_column(
            "Ghi chú",
            editable=True,
            width=180,
            minWidth=130,
        )

    gb.configure_column("::auto_unique_id::", hide=True)
    red_check_style = JsCode("""
    function(params) {
        return {
            'color': '#d92d20',
            'fontWeight': '500'
        };
    }
    """)

    for col in df.columns:
        if "check" in str(col).lower():
            gb.configure_column(col, cellStyle=red_check_style)
    gb.configure_grid_options(
        headerHeight=32,
        floatingFiltersHeight=30,
        rowHeight=30,
        singleClickEdit=True,
        stopEditingWhenCellsLoseFocus=True,
        enableCellTextSelection=True,
        ensureDomOrder=True,
    )

    custom_css = {
        ".ag-header": {
            "background-color": "#F8FAFC !important",
            "border-bottom": "1px solid #CBD5E1 !important",
        },
        ".ag-header-cell": {
            "border-right": "1px solid #E2E8F0 !important",
            "font-size": "12.5px !important",
            "font-weight": "600 !important",
            "color": "#334155 !important",
            "padding-left": "8px !important",
            "padding-right": "8px !important",
        },
        ".ag-cell": {
            "border-right": "1px solid #E2E8F0 !important",
            "font-size": "13px !important",
            "padding-left": "8px !important",
            "padding-right": "8px !important",
            "line-height": "28px !important",
        },
        ".ag-row": {
            "border-bottom": "1px solid #E2E8F0 !important",
        },
        ".ag-floating-filter-input input": {
            "height": "22px !important",
            "min-height": "22px !important",
            "font-size": "12px !important",
            "padding": "0 6px !important",
            "border-radius": "4px !important",
        },
        ".ag-floating-filter-button": {
            "margin-left": "2px !important",
        },
        ".ag-floating-filter-button-button": {
            "width": "18px !important",
            "height": "18px !important",
        },
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


# ======================= TIỆN ÍCH FILE / STATE / PHÂN QUYỀN / ĐỒNG BỘ =======================
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


def attach_notes_to_th(df_th):
    if df_th.empty:
        df = df_th.copy()
        df["Đã sửa"] = pd.Series(dtype=bool)
        df["Ghi chú"] = pd.Series(dtype=str)
        return df

    notes = load_th_notes()
    df = df_th.copy()
    da_sua_list = []
    ghi_chu_list = []
    for stt in df["STT"]:
        rec = notes.get(str(stt), {})
        da_sua_list.append(bool(rec.get("da_sua", False)))
        ghi_chu_list.append(str(rec.get("ghi_chu", "")))

    df["Đã sửa"] = da_sua_list
    df["Ghi chú"] = ghi_chu_list
    other_cols = [c for c in df.columns if c not in ("STT", "Đã sửa", "Ghi chú")]
    return df[["STT", "Đã sửa", "Ghi chú"] + other_cols]


def sync_edited_notes(page_before_df, page_after_df):
    """So sánh theo STT trong slice trang hiện tại và gộp vào notes máy chủ."""
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


def save_latest_result(result, src_filename, used_ref):
    try:
        with open(LATEST_RESULT_FILE, 'wb') as fh:
            pickle.dump({
                'result': result,
                'src_filename': src_filename,
                'used_ref': used_ref,
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


def _strip_accents(text):
    s = unicodedata.normalize('NFD', str(text).lower())
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.replace('đ', 'd')


def match_query(df, query, columns):
    """Lọc thông minh: bỏ dấu tiếng Việt, gõ nhiều từ rời nhau vẫn khớp."""
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


def render_pagination_bar(total_rows, page_key, size_key, info_html=""):
    """Thanh phân trang gọn gàng, có tùy chọn Tất cả (≤2.000 dòng) để lọc full trong bảng."""
    size_options = [100, 200, 500, OPT_ALL_ROWS]
    if size_key not in st.session_state or st.session_state[size_key] not in size_options:
        st.session_state[size_key] = 100
    if page_key not in st.session_state:
        st.session_state[page_key] = 0

    col_info, col_sz, col_nav = st.columns([4.0, 1.5, 2.5], vertical_alignment="center")

    with col_sz:
        selected_sz = st.selectbox(
            "Dòng/trang",
            options=size_options,
            key=size_key,
            label_visibility="collapsed",
        )

    cap_note = ""
    if selected_sz == OPT_ALL_ROWS:
        if total_rows <= MAX_ALL_ROWS:
            page_size = max(1, total_rows)
        else:
            page_size = MAX_ALL_ROWS
            cap_note = f" <i>(giới hạn {MAX_ALL_ROWS:,} dòng/trang để tránh đơ)</i>"
    else:
        page_size = int(selected_sz)

    n_pages = max(1, (total_rows + page_size - 1) // page_size)
    page = min(max(0, int(st.session_state[page_key])), n_pages - 1)
    st.session_state[page_key] = page

    start = page * page_size
    end = min(start + page_size, total_rows)

    with col_info:
        range_txt = f"{start + 1:,}–{end:,}" if total_rows > 0 else "0"
        st.markdown(
            f"<div class='soft-note'>Hiển thị <b>{range_txt}</b> / <b>{total_rows:,}</b> dòng "
            f"{info_html}{cap_note}</div>",
            unsafe_allow_html=True,
        )

    with col_nav:
        b1, b2, b3, b4, b5 = st.columns([0.7, 0.7, 1.6, 0.7, 0.7], vertical_alignment="center")
        if b1.button("«", key=f"{page_key}_first", disabled=(page <= 0), use_container_width=True):
            st.session_state[page_key] = 0
            st.rerun()
        if b2.button("‹", key=f"{page_key}_prev", disabled=(page <= 0), use_container_width=True):
            st.session_state[page_key] = page - 1
            st.rerun()
        b3.markdown(
            f"<div style='text-align:center;font-size:12.5px;font-weight:600;color:#334155;'>"
            f"{page + 1} / {n_pages}</div>",
            unsafe_allow_html=True,
        )
        if b4.button("›", key=f"{page_key}_next", disabled=(page >= n_pages - 1), use_container_width=True):
            st.session_state[page_key] = page + 1
            st.rerun()
        if b5.button("»", key=f"{page_key}_last", disabled=(page >= n_pages - 1), use_container_width=True):
            st.session_state[page_key] = n_pages - 1
            st.rerun()

    return start, end, page, page_size


# ======================= SIDEBAR =======================
def build_sidebar(admin=True):
    with st.sidebar:
        st.markdown("### 🩺 BHYT Anh Quất")
        st.markdown("<div class='soft-note'>Hệ thống kiểm tra lỗi hồ sơ XML3</div>", unsafe_allow_html=True)
        st.markdown("---")

        st.markdown("**🧭 Chức năng**")
        page = st.radio(
            "Điều hướng",
            options=[PAGE_SCAN, PAGE_BED, PAGE_CATALOG],
            label_visibility="collapsed",
        )
        st.markdown("---")

        active_ref = load_active_ref()
        ref_path = active_ref

        if admin:
            st.markdown("**⚙️ Cấu hình dữ liệu Cổng BHYT**")

            st.download_button(
                "📥 Tải file mẫu (MAU_THAM_CHIEU.xlsx)",
                data=build_template_excel_bytes(),
                file_name="MAU_THAM_CHIEU.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.markdown(
                "<div class='soft-note'>Tải lên file danh mục (NVYT / TTB / DVKT). "
                "File sẽ được lưu vào thư mục <b>data_ref/</b> và tự động ghi nhớ.</div>",
                unsafe_allow_html=True,
            )

            uploaded_ref = st.file_uploader(
                "File tham chiếu (.xlsx, .xls, .ods)",
                type=["xlsx", "xls", "ods"],
                key="ref_uploader",
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
                st.success(f"✅ Đang dùng: {os.path.basename(ref_path)}")
                if st.button("🔄 Xóa/Đổi file", use_container_width=True):
                    set_active_ref("")
                    st.session_state.pop('_catalog_cache_key', None)
                    st.rerun()
            else:
                st.warning("⚠️ Chưa nạp file tham chiếu (bắt buộc cho ĐK 5 & 6).")
        else:
            st.info(
                "👨‍⚕️ **Chế độ Bác sĩ / Khoa phòng**\n\n"
                "Lọc khoa/tên mình, tích **Đã sửa** hoặc gõ **Ghi chú** — hệ thống tự động lưu về máy chủ."
            )

        st.markdown("---")
        role_label = "Quản trị" if admin else "Xem & Đánh dấu"
        st.markdown(
            f"<div class='soft-note'>Phiên bản Web • {role_label} • Engine: checker.py</div>",
            unsafe_allow_html=True,
        )

    return page, (ref_path if (ref_path and os.path.exists(ref_path)) else None)


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

    has_result = result is not None

    if admin:
        exp_label = (
            f"📤 Nạp hồ sơ mới / Quét lại (đang xem: {src_filename})"
            if has_result and src_filename else
            "📤 Nạp hồ sơ cần kiểm tra"
        )
        with st.expander(exp_label, expanded=not has_result):
            if not ref_path:
                st.info("💡 Chưa có file tham chiếu ở Sidebar — hệ thống sẽ chỉ chạy **ĐK 1 → ĐK 4** "
                        "(bỏ qua ĐK 5 & ĐK 6).")

            col_upload, col_action = st.columns([3.5, 1], gap="medium", vertical_alignment="center")
            with col_upload:
                input_file = st.file_uploader(
                    "Kéo thả file Excel (.xlsx, .xls) hoặc XML (.xml)",
                    type=["xlsx", "xls", "xml"],
                    key="data_uploader",
                    label_visibility="collapsed",
                )
            with col_action:
                run_btn = st.button("🚀 Bắt đầu quét lỗi", type="primary", use_container_width=True)

            if run_btn:
                if input_file is None:
                    st.warning("Vui lòng tải lên file dữ liệu trước khi quét.")
                    st.stop()

                progress = st.progress(0.0, text="Đang khởi động...")

                def cb(pct, msg):
                    progress.progress(min(max(int(pct * 100), 0), 100), text=msg)

                try:
                    df_raw = read_input_dataframe(input_file.getvalue(), input_file.name)
                    catalogs = {}
                    if ref_path:
                        catalogs, _ = get_catalogs(ref_path)
                    result = run_check(df_raw, catalogs=catalogs, ref_path=ref_path or "", progress_cb=cb)
                except Exception as ex:
                    progress.empty()
                    st.error(f"❌ Lỗi xử lý: {ex}")
                    st.stop()

                progress.progress(100, text="Hoàn tất!")
                progress.empty()

                st.session_state['result'] = result
                st.session_state['src_filename'] = input_file.name
                st.session_state['used_ref'] = bool(ref_path)
                st.session_state['th_page'] = 0
                st.session_state['ct_page'] = 0
                clear_cached_exports()
                save_latest_result(result, input_file.name, bool(ref_path))
                reset_th_notes()
                st.toast(f"✅ Quét xong {input_file.name} ({result['stats']['total_errors']:,} lỗi)", icon="🎉")
                st.rerun()

    if not result:
        msg = (
            "Chưa có dữ liệu kiểm tra. Vui lòng chờ quản trị viên cập nhật hồ sơ."
            if not admin else
            "Chưa có kết quả. Hãy tải file ở trên và bấm <b>🚀 Bắt đầu quét lỗi</b>."
        )
        st.markdown(f"<div class='card'><span class='soft-note'>{msg}</span></div>", unsafe_allow_html=True)
        return

    if not admin and src_filename:
        st.caption(f"📄 Đang xem kết quả mới nhất của file: **{src_filename}**")

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
    """Lọc nhanh theo cột chính xác trước, rồi mới chạy tìm kiếm không dấu."""
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

    # ----- Thanh lọc full dữ liệu (2 hàng mảnh) -----
    khoa_options = sorted(
        [k for k in df_out['Khoa'].dropna().astype(str).str.strip().unique() if k]
    ) if 'Khoa' in df_out.columns else []

    r1_c1, r1_c2, r1_c3 = st.columns([2.8, 2.1, 2.1], gap="small")
    with r1_c1:
        query = st.text_input(
            "🔎 Tìm kiếm nhanh",
            placeholder="🔎 Tìm không dấu: Tên BN, Bác sĩ, Mã DV, Mã máy, Ghi chú...",
            key="search_query",
            label_visibility="collapsed",
        )
    with r1_c2:
        selected_khoa = st.multiselect(
            "🏥 Lọc theo Khoa",
            options=khoa_options,
            default=[],
            placeholder="🏥 Lọc theo Khoa / Phòng khám...",
            key="khoa_filter",
            label_visibility="collapsed",
        )
    with r1_c3:
        selected_dk = st.multiselect(
            "🎯 Lọc điều kiện",
            options=DK_COLUMNS,
            default=[],
            placeholder="🎯 Lọc điều kiện (ĐK 1 → ĐK 6)...",
            key="dk_filter",
            label_visibility="collapsed",
        )

    r2_c1, r2_c2, r2_c3 = st.columns([2.2, 2.2, 1.0], gap="small")
    with r2_c1:
        selected_status = st.selectbox(
            "📋 Trạng thái BN",
            options=["Tất cả trạng thái", "Đang điều trị", "Ra viện"],
            key="status_filter",
            label_visibility="collapsed",
        )
    with r2_c2:
        selected_fix = st.selectbox(
            "✅ Tình trạng sửa",
            options=["Tất cả (Đã sửa & Chưa sửa)", "Chưa sửa", "Đã sửa"],
            key="fix_filter",
            label_visibility="collapsed",
        )
    with r2_c3:
        if st.button("🔄 Đồng bộ", use_container_width=True, help="Tải lại trạng thái Đã sửa / Ghi chú mới nhất"):
            clear_cached_exports()
            st.rerun()

    tab_th, tab_ct = st.tabs(["📋 Tổng hợp lỗi (TH)", "📑 Dữ liệu chi tiết (chitiet)"])

    with tab_th:
        df_th_view = filter_th(df_th_with_notes, selected_dk, selected_khoa, selected_status, selected_fix, query)
        done_count = int(df_th_with_notes["Đã sửa"].sum()) if not df_th_with_notes.empty else 0

        extra_info = (
            f"• ✅ Đã sửa: <b>{done_count:,}</b>/{len(df_th_with_notes):,} "
            f"<i>(Tích Đã sửa / gõ Ghi chú tự lưu máy chủ)</i>"
        )
        start, end, th_page, th_pgsz = render_pagination_bar(
            len(df_th_view), "th_page", "th_pgsz", info_html=extra_info
        )

        if df_th_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
        else:
            page_before = df_th_view.iloc[start:end].copy()
            page_after = render_excel_table(
                page_before,
                key=f"grid_th_{th_page}_{th_pgsz}",
                height=620,
                editable_cols=["Đã sửa", "Ghi chú"],
            )
            if sync_edited_notes(page_before, page_after):
                st.toast("✅ Đã lưu trạng thái & ghi chú lên máy chủ!", icon="💾")
                df_th_with_notes = attach_notes_to_th(result['df_th'])
                df_th_view = filter_th(df_th_with_notes, selected_dk, selected_khoa, selected_status, selected_fix, query)

    with tab_ct:
        only_errors = st.checkbox("Chỉ hiện các dòng có lỗi", value=True, key="only_errors")
        df_ct_view = filter_chitiet(df_out, selected_dk, selected_khoa, selected_status, query, only_errors)

        ct_start, ct_end, _, _ = render_pagination_bar(
            len(df_ct_view), "ct_page", "ct_pgsz",
            info_html=f"(tổng gốc: {len(df_out):,})"
        )

        if df_ct_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
        else:
            st.dataframe(
                df_ct_view.iloc[ct_start:ct_end],
                use_container_width=True,
                hide_index=True,
                height=620,
            )

    # ----- Tải kết quả (Chỉ sinh file Excel khi bấm nút — không lag 40k dòng) -----
    st.markdown("---")
    d1, d2 = st.columns(2, gap="large")
    full_name = make_output_filename(src_filename)

    with d1:
        if st.button("📦 Tạo file Excel đầy đủ (TH + chitiet)", type="primary", use_container_width=True):
            with st.spinner("Đang tạo file Excel đầy đủ..."):
                latest_th = attach_notes_to_th(result['df_th'])
                st.session_state['_excel_full_bytes'] = build_output_excel_bytes(
                    format_th_for_export(latest_th), df_out
                )
        if st.session_state.get('_excel_full_bytes'):
            st.download_button(
                "⬇️ Tải xuống Excel đầy đủ",
                data=st.session_state['_excel_full_bytes'],
                file_name=full_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with d2:
        if st.button("📦 Tạo file Excel danh sách đang lọc", use_container_width=True):
            with st.spinner("Đang tạo file Excel theo bộ lọc hiện tại..."):
                latest_th = attach_notes_to_th(result['df_th'])
                th_filtered_full = filter_th(
                    latest_th, selected_dk, selected_khoa, selected_status, selected_fix, query
                )
                ct_filtered_full = filter_chitiet(
                    df_out, selected_dk, selected_khoa, selected_status, query,
                    st.session_state.get("only_errors", True)
                )
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    format_th_for_export(th_filtered_full).to_excel(writer, sheet_name='TH', index=False)
                    ct_filtered_full.to_excel(writer, sheet_name='chitiet', index=False)
                buf.seek(0)
                st.session_state['_excel_filtered_bytes'] = buf.getvalue()
        if st.session_state.get('_excel_filtered_bytes'):
            st.download_button(
                "⬇️ Tải xuống Excel đang lọc",
                data=st.session_state['_excel_filtered_bytes'],
                file_name=full_name.replace(".xlsx", "_Loc.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if not used_ref:
        st.info("ℹ️ Lần quét này chưa dùng file tham chiếu — ĐK 5 & ĐK 6 (phần cấu hình) chưa được kiểm tra đầy đủ.")


# ======================= TRANG 2: QUẢN LÝ DANH MỤC CỔNG =======================
def page_catalog(ref_path, admin=True):
    st.markdown("<div class='section-title'>📂 Quản lý Danh mục Cổng BHYT</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='soft-note'>Xem nhanh dữ liệu 3 sheet danh mục đang dùng: "
        "<b>data_nvyt</b> (nhân viên y tế), <b>data_tbyt</b> (trang thiết bị), <b>data_dvkt</b> (dịch vụ kỹ thuật).</div>",
        unsafe_allow_html=True,
    )

    if not ref_path:
        if admin:
            st.warning("⚠️ Chưa có file tham chiếu. Hãy tải lên ở Sidebar (mục **Cấu hình dữ liệu Cổng BHYT**).")
        else:
            st.warning("⚠️ Chưa có dữ liệu danh mục tham chiếu. Vui lòng chờ quản trị viên cập nhật.")
        st.markdown("---")
        st.markdown("<div class='section-title'>🧩 Module mở rộng (sắp ra mắt)</div>", unsafe_allow_html=True)
        with st.expander("🛏️ Kiểm tra giường bệnh / nằm ghép / giường trống theo khoa"):
            st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")
        with st.expander("🏥 Quản lý danh mục Giường / Khoa phòng / DVKT"):
            st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")
        return

    st.success(f"✅ File đang dùng: **{os.path.basename(ref_path)}**")
    catalogs, frames = get_catalogs(ref_path)

    c1, c2, c3 = st.columns(3)
    c1.metric("Bác sĩ (data_nvyt)", f"{len(catalogs['dict_ten_bs']):,}")
    c2.metric("Thiết bị (data_tbyt)", f"{len(catalogs['dict_may']):,}")
    c3.metric("DVKT (data_dvkt)", f"{len(catalogs['dict_dvkt_hieuluc']):,}")

    st.markdown("---")

    tabs = st.tabs(["👨‍⚕️ data_nvyt", "🩻 data_tbyt", "🧾 data_dvkt"])
    sheet_map = [
        ('data_nvyt', ['MACCHN', 'HO_TEN', 'Mã CK', 'Tên Công khám']),
        ('data_tbyt', ['MA_MAY', 'TEN_TB', 'KY_HIEU', 'CONGTY_SX']),
        ('data_dvkt', ['MA', 'TEN', 'HIEU_LUC', 'LY_DO_TU_CHOI']),
    ]
    for tab, (sheet_name, cols) in zip(tabs, sheet_map):
        with tab:
            df_sheet = frames.get(sheet_name)
            if df_sheet is None or df_sheet.empty:
                st.info(f"Sheet **{sheet_name}** chưa có dữ liệu.")
                continue
            kw = st.text_input(f"🔎 Tìm trong {sheet_name}", key=f"cat_search_{sheet_name}")
            view = df_sheet
            if kw:
                view = match_query(df_sheet, kw, list(df_sheet.columns))
            st.caption(f"{len(view):,} / {len(df_sheet):,} dòng")
            st.dataframe(view, use_container_width=True, hide_index=True, height=460)

    st.markdown("---")
    st.markdown("<div class='section-title'>🧩 Module mở rộng (sắp ra mắt)</div>", unsafe_allow_html=True)
    with st.expander("🛏️ Kiểm tra giường bệnh / nằm ghép / giường trống theo khoa"):
        st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")
    with st.expander("🏥 Quản lý danh mục Giường / Khoa phòng / DVKT"):
        st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")


# ======================= MAIN =======================
def main():
    inject_css()
    inject_close_dropdown_on_outside_click()
    admin = is_admin()
    page, ref_path = build_sidebar(admin=admin)
    render_header()

    if page == PAGE_SCAN:
        page_scan(ref_path, admin=admin)
    elif page == PAGE_BED:
        page_bed_management(ref_path, admin=admin, render_kpi_fn=render_kpi_row)
    else:
        page_catalog(ref_path, admin=admin)


if __name__ == "__main__":
    main()
