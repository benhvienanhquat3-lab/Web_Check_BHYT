# -*- coding: utf-8 -*-
"""
app.py
=======================================================================
Web App Streamlit: Hệ thống kiểm tra lỗi BHYT (XML3 / Excel)
Bệnh viện Đa khoa Anh Quất.

Kiến trúc file gọn gàng:
    * Toàn bộ logic nghiệp vụ -> ``checker.py``
    * Giao diện & điều hướng  -> ``app.py`` (file này)
=======================================================================
"""

import os
import sys
import json
import base64
import pickle
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
    HAS_AGGRID = True
except ImportError:
    HAS_AGGRID = False


def render_excel_table(df, key, height=460, editable_cols=None):
    """Hiển thị bảng kẻ ô kiểu Excel, thanh lọc gọn, hỗ trợ tích Đã sửa & Ghi chú."""
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
    )

    # Căn độ rộng gọn cho các cột ngắn
    for narrow_col, w in [("STT", 65), ("Dòng", 75), ("Id", 75), ("MA_NHOM", 90)]:
        if narrow_col in df.columns:
            gb.configure_column(narrow_col, width=w, minWidth=w)

    # Cấu hình cột cho phép bác sĩ tích chọn & ghi chú trực tiếp
    if "Đã sửa" in editable_cols and "Đã sửa" in df.columns:
        gb.configure_column(
            "Đã sửa",
            editable=True,
            width=85,
            minWidth=85,
            cellStyle={"textAlign": "center"},
        )
    if "Ghi chú" in editable_cols and "Ghi chú" in df.columns:
        gb.configure_column(
            "Ghi chú",
            editable=True,
            width=190,
            minWidth=140,
        )

    # Ẩn cột nội bộ của AgGrid nếu tự sinh ra
    gb.configure_column("::auto_unique_id::", hide=True)

    gb.configure_grid_options(
        headerHeight=32,            # Tiêu đề cột gọn
        floatingFiltersHeight=30,   # Hàng ô lọc nhỏ thanh thoát
        rowHeight=30,               # Chiều cao dòng gọn như bản cũ
        singleClickEdit=True,       # Bấm 1 lần là tích hoặc gõ luôn
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
    )
    filtered = resp.get("data")
    if filtered is not None:
        out_df = pd.DataFrame(filtered)
        return out_df.drop(columns=["::auto_unique_id::"], errors="ignore")
    return df


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

st.set_page_config(
    layout="wide",
    page_title="Kiểm Tra Lỗi BHYT - BVĐK Anh Quất",
    page_icon="🩺",
    initial_sidebar_state="expanded",
)


# ======================= CUSTOM CSS (Medical SaaS) =======================
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
            padding-top: 1.1rem;
            padding-bottom: 2rem;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            max-width: 100% !important;
        }

        /* ---------- Header ---------- */
        .app-header {
            display: flex; align-items: center; gap: 18px;
            background: linear-gradient(100deg,#1565C0 0%,#1E88E5 55%,#42A5F5 100%);
            border-radius: 12px; padding: 14px 22px; margin-bottom: 16px;
            box-shadow: 0 6px 18px rgba(21,101,192,0.22);
        }
        .app-logo {
            height: 60px; width: 60px; border-radius: 10px; object-fit: cover;
            background: #FFFFFF; padding: 3px;
        }
        .app-title {
            color: #FFFFFF; font-size: 21px; font-weight: 700;
            letter-spacing: .4px; line-height: 1.25;
        }
        .app-sub {
            color: #E3F2FD; font-size: 13px; font-weight: 600;
            letter-spacing: .7px; margin-top: 2px;
        }

        /* ---------- Section ---------- */
        .section-title {
            font-size: 16px; font-weight: 700; color: #0F172A;
            margin: 6px 0 12px 0;
        }
        .soft-note { color: #64748B; font-size: 13px; }

        /* ---------- Card ---------- */
        .card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
            padding: 16px 18px; box-shadow: 0 1px 3px rgba(15,23,42,.04);
        }

        /* ---------- KPI ---------- */
        .kpi-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
            padding: 12px 10px; text-align: center; margin-bottom: 8px;
            box-shadow: 0 1px 3px rgba(15,23,42,.05);
        }
        .kpi-value { font-size: 23px; font-weight: 700; line-height: 1.1; }
        .kpi-label {
            font-size: 11px; color: #64748B; font-weight: 600; margin-top: 4px;
            text-transform: uppercase; letter-spacing: .3px;
        }

        /* ---------- Sidebar ---------- */
        section[data-testid="stSidebar"] {
            background: #FFFFFF; border-right: 1px solid #E2E8F0;
        }

        /* ---------- Buttons ---------- */
        .stButton > button, .stDownloadButton > button { border-radius: 9px; font-weight: 600; }
        .stButton > button[kind="primary"] { background: #1565C0; border-color: #1565C0; }
        </style>
        """,
        unsafe_allow_html=True,
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
            f"""<div class="kpi-card" style="border-top:4px solid {color};">
                    <div class="kpi-value" style="color:{color};">{value}</div>
                    <div class="kpi-label">{label}</div>
                </div>""",
            unsafe_allow_html=True,
        )


# ======================= TIỆN ÍCH FILE / STATE / PHÂN QUYỀN / ĐỒNG BỘ =======================
def load_active_ref():
    """Trả về đường dẫn file tham chiếu đang dùng (nhớ từ lần trước)."""
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
    """Chỉ cấp quyền Admin khi truy cập trực tiếp từ máy chủ (localhost / 127.0.0.1)."""
    try:
        host = st.context.headers.get("Host", "").lower()
        return host.startswith("localhost") or host.startswith("127.0.0.1")
    except Exception:
        return False


def load_th_notes():
    """Đọc trạng thái Đã sửa & Ghi chú đang lưu trên máy chủ."""
    if os.path.exists(TH_NOTES_FILE):
        try:
            with open(TH_NOTES_FILE, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def save_th_notes(notes):
    """Ghi trạng thái Đã sửa & Ghi chú lên máy chủ."""
    try:
        tmp_path = TH_NOTES_FILE + ".tmp"
        with open(tmp_path, 'w', encoding='utf-8') as fh:
            json.dump(notes, fh, ensure_ascii=False, indent=2)
        os.replace(tmp_path, TH_NOTES_FILE)
    except Exception:
        pass


def reset_th_notes():
    """Xóa ghi chú cũ khi quét bộ hồ sơ mới."""
    save_th_notes({})


def attach_notes_to_th(df_th):
    """Gắn 2 cột 'Đã sửa' và 'Ghi chú' (đồng bộ từ máy chủ) ngay cạnh cột STT."""
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

    # Đặt ngay cạnh STT để bác sĩ dễ nhìn & tích luôn
    df["Đã sửa"] = da_sua_list
    df["Ghi chú"] = ghi_chu_list
    other_cols = [c for c in df.columns if c not in ("STT", "Đã sửa", "Ghi chú")]
    return df[["STT", "Đã sửa", "Ghi chú"] + other_cols]


def sync_edited_notes(before_df, after_df):
    """So sánh các dòng thay đổi và đồng bộ đúng các STT vừa sửa về máy chủ."""
    if before_df.empty or after_df.empty:
        return False
    if "STT" not in after_df.columns or "Đã sửa" not in after_df.columns:
        return False

    before_map = {}
    for _, r in before_df.iterrows():
        before_map[str(r["STT"])] = (
            bool(r.get("Đã sửa", False)),
            str(r.get("Ghi chú", "") or "").strip(),
        )

    changes = {}
    for _, r in after_df.iterrows():
        k = str(r["STT"])
        new_val = (
            bool(r.get("Đã sửa", False)),
            str(r.get("Ghi chú", "") or "").strip(),
        )
        if k in before_map and before_map[k] != new_val:
            changes[k] = {"da_sua": new_val[0], "ghi_chu": new_val[1]}

    if not changes:
        return False

    # Đọc bản mới nhất trên máy chủ rồi chỉ gộp đúng các dòng vừa sửa (tránh đè BS khác)
    notes = load_th_notes()
    for k, val in changes.items():
        if val["da_sua"] or val["ghi_chu"]:
            notes[k] = val
        else:
            notes.pop(k, None)
    save_th_notes(notes)
    return True


def format_th_for_export(df_th):
    """Định dạng cột Đã sửa thành chữ gọn gàng khi xuất Excel."""
    df = df_th.copy()
    if "Đã sửa" in df.columns:
        df["Đã sửa"] = df["Đã sửa"].apply(lambda x: "Đã sửa" if bool(x) else "")
    return df


def save_latest_result(result, src_filename, used_ref):
    """Lưu kết quả quét mới nhất để các máy khác vào qua IP đều xem được ngay."""
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
    """Đọc kết quả quét mới nhất đã lưu trên máy chủ."""
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
    """Đọc 3 sheet danh mục tham chiếu (cache theo thời gian sửa file)."""
    return _load_catalogs_cached(ref_path, os.path.getmtime(ref_path))


def match_query(df, query, columns):
    """Lọc nhanh theo từ khoá trên các cột chỉ định."""
    if not query or df.empty:
        return df
    q = query.strip().lower()
    mask = pd.Series(False, index=df.index)
    for col in columns:
        if col in df.columns:
            mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False, regex=False)
    return df[mask]


# ======================= SIDEBAR =======================
def build_sidebar(admin=True):
    """Sidebar: menu điều hướng + cấu hình file tham chiếu (ẩn với người xem qua IP)."""
    with st.sidebar:
        st.markdown("### 🩺 BHYT Anh Quất")
        st.markdown("<div class='soft-note'>Hệ thống kiểm tra lỗi hồ sơ XML3</div>", unsafe_allow_html=True)
        st.markdown("---")

        st.markdown("**🧭 Chức năng**")
        page = st.radio(
            "Điều hướng",
            options=[PAGE_SCAN, PAGE_CATALOG],
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
                "Lọc tên mình, tích **Đã sửa** hoặc gõ **Ghi chú** — hệ thống tự động lưu về máy chủ."
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
    if admin:
        st.markdown("<div class='section-title'>📤 Nạp hồ sơ cần kiểm tra</div>", unsafe_allow_html=True)

        if not ref_path:
            st.info("💡 Chưa có file tham chiếu ở Sidebar — hệ thống sẽ chỉ chạy **ĐK 1 → ĐK 4** "
                    "(bỏ qua ĐK 5 & ĐK 6). Bạn có thể nạp danh mục Cổng để kiểm tra đầy đủ 6 điều kiện.")

        col_upload, col_action = st.columns([3, 1], gap="large")

        with col_upload:
            input_file = st.file_uploader(
                "Kéo thả file Excel (.xlsx, .xls) hoặc XML (.xml)",
                type=["xlsx", "xls", "xml"],
                key="data_uploader",
                help="Hỗ trợ file XML3 xuất từ cổng giám định BHYT hoặc file Excel chi tiết.",
            )

        with col_action:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
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
            save_latest_result(result, input_file.name, bool(ref_path))
            reset_th_notes()
            st.success(f"✅ Quét xong **{input_file.name}** — phát hiện **{result['stats']['total_errors']:,}** lỗi.")

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
            "Chưa có kết quả. Hãy tải file và bấm <b>🚀 Bắt đầu quét lỗi</b>."
        )
        st.markdown(f"<div class='card'><span class='soft-note'>{msg}</span></div>", unsafe_allow_html=True)
        return

    if not admin and src_filename:
        st.info(f"📄 Đang xem kết quả kiểm tra mới nhất của file: **{src_filename}**")

    render_results(result, src_filename, used_ref)


# ======================= HIỂN THỊ KẾT QUẢ =======================
SEARCH_COLS_TH = ['Dòng', 'Mã DV', 'Tên dịch vụ', 'Chi tiết lỗi', 'Ghi chú',
                  'Mã DV_check', 'Tên dịch vụ_check']
SEARCH_COLS_CT = ['MA_LK', 'MA_DICH_VU', 'TEN_DICH_VU', 'NGUOI_THUC_HIEN',
                  'MA_MAY'] + DK_COLUMNS


def filter_th(df_th, selected_dk, query):
    df = df_th
    if selected_dk:
        df = df[df['Loại lỗi'].isin(selected_dk)]
    return match_query(df, query, SEARCH_COLS_TH)


def filter_chitiet(df_out, selected_dk, query, only_errors):
    df = df_out
    if selected_dk:
        mask = pd.Series(False, index=df.index)
        for col in selected_dk:
            if col in df.columns:
                mask = mask | (df[col].astype(str).str.strip() != "")
        df = df[mask]
    if only_errors:
        mask_err = pd.Series(False, index=df.index)
        for col in DK_COLUMNS:
            if col in df.columns:
                mask_err = mask_err | (df[col].astype(str).str.strip() != "")
        df = df[mask_err]
    return match_query(df, query, SEARCH_COLS_CT)


def render_results(result, src_filename, used_ref):
    df_out = result['df_out']
    df_th_with_notes = attach_notes_to_th(result['df_th'])
    stats = result['stats']

    st.markdown("<div class='section-title'>📊 Bảng điều khiển kết quả</div>", unsafe_allow_html=True)

    kpi_items = [
        ("Tổng dòng quét", f"{stats['total_rows']:,}", "#1565C0"),
        ("Tổng lỗi phát hiện", f"{stats['total_errors']:,}", "#C62828"),
    ]
    for i, col in enumerate(DK_COLUMNS, start=1):
        kpi_items.append((f"ĐK {i}", f"{stats['by_dk'][col]:,}", DK_COLORS[i]))
    render_kpi_row(kpi_items)

    st.markdown("---")

    # ----- Thanh lọc & tìm kiếm -----
    with st.container():
        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            query = st.text_input(
                "🔎 Tìm kiếm nhanh",
                placeholder="Tìm theo MA_LK, Mã DV, Tên dịch vụ, Bác sĩ, Mã máy, Chi tiết lỗi, Ghi chú...",
                key="search_query",
            )
        with c2:
            selected_dk = st.multiselect(
                "🎯 Lọc theo điều kiện",
                options=DK_COLUMNS,
                default=[],
                format_func=lambda c: c,
                key="dk_filter",
            )

    tab_th, tab_ct = st.tabs(["📋 Tổng hợp lỗi (TH)", "📑 Dữ liệu chi tiết (chitiet)"])

    with tab_th:
        df_th_view = filter_th(df_th_with_notes, selected_dk, query)
        done_count = int(df_th_with_notes["Đã sửa"].sum()) if not df_th_with_notes.empty else 0

        top_l, top_r = st.columns([4, 1])
        with top_l:
            st.markdown(
                f"<div class='soft-note'>Hiển thị <b>{len(df_th_view):,}</b> / {len(df_th_with_notes):,} dòng lỗi "
                f"• ✅ Đã sửa: <b>{done_count:,}</b> / {len(df_th_with_notes):,} "
                f"<i>(Click trực tiếp vào cột <b>Đã sửa</b> hoặc <b>Ghi chú</b> để lưu tự động về máy chủ)</i></div>",
                unsafe_allow_html=True,
            )
        with top_r:
            if st.button("🔄 Đồng bộ mới nhất", use_container_width=True):
                st.rerun()

        if df_th_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
            th_live = df_th_view
        else:
            th_live = render_excel_table(
                df_th_view,
                key="grid_th",
                height=460,
                editable_cols=["Đã sửa", "Ghi chú"],
            )
            if sync_edited_notes(df_th_view, th_live):
                st.toast("✅ Đã lưu trạng thái & ghi chú lên máy chủ!", icon="💾")
                df_th_with_notes = attach_notes_to_th(result['df_th'])

    with tab_ct:
        only_errors = st.checkbox("Chỉ hiện các dòng có lỗi", value=True, key="only_errors")
        df_ct_view = filter_chitiet(df_out, selected_dk, query, only_errors)
        st.markdown(
            f"<div class='soft-note'>Hiển thị <b>{len(df_ct_view):,}</b> / {len(df_out):,} dòng chi tiết</div>",
            unsafe_allow_html=True,
        )
        if df_ct_view.empty:
            st.info("Không có dòng nào khớp bộ lọc hiện tại.")
            ct_live = df_ct_view
        else:
            ct_live = render_excel_table(df_ct_view, key="grid_ct", height=460)

    # ----- Tải kết quả -----
    st.markdown("---")
    st.markdown("<div class='section-title'>⬇️ Tải kết quả</div>", unsafe_allow_html=True)

    d1, d2 = st.columns(2, gap="large")
    full_name = make_output_filename(src_filename)

    with d1:
        st.download_button(
            "⬇️ Tải file Excel đầy đủ (TH + chitiet)",
            data=build_output_excel_bytes(format_th_for_export(df_th_with_notes), df_out),
            file_name=full_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with d2:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            format_th_for_export(th_live).to_excel(writer, sheet_name='TH', index=False)
            ct_live.to_excel(writer, sheet_name='chitiet', index=False)
        buffer.seek(0)
        st.download_button(
            "⬇️ Tải riêng danh sách đang lọc",
            data=buffer.getvalue(),
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
            st.dataframe(view, use_container_width=True, hide_index=True, height=380)

    st.markdown("---")
    st.markdown("<div class='section-title'>🧩 Module mở rộng (sắp ra mắt)</div>", unsafe_allow_html=True)
    with st.expander("🛏️ Kiểm tra giường bệnh / nằm ghép / giường trống theo khoa"):
        st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")
    with st.expander("🏥 Quản lý danh mục Giường / Khoa phòng / DVKT"):
        st.caption("Module đang được phát triển — sẽ dùng chung thư mục `data_ref/`.")


# ======================= MAIN =======================
def main():
    inject_css()
    admin = is_admin()
    page, ref_path = build_sidebar(admin=admin)
    render_header()

    if page == PAGE_SCAN:
        page_scan(ref_path, admin=admin)
    else:
        page_catalog(ref_path, admin=admin)


if __name__ == "__main__":
    main()