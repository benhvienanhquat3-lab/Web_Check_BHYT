# -*- coding: utf-8 -*-
"""
bed_manager.py
=======================================================================
Module độc lập: Sơ đồ & Kiểm tra trùng Giường bệnh nội trú (MA_NHOM = 15)
- Tách độc lập hoàn toàn THEO TỪNG NGÀY MỘT (không gộp đè giữa các ngày).
- Chỉ dùng mã giường thật (MA_GIUONG_CV130 / MA_GIUONG); BN trống mã giường
  tách ra bảng phụ riêng, tuyệt đối không tự gán tạm gây trùng ảo.
- BN trống cả NGAY_KQ và NGAY_RA: rải từng ngày đến 23:59 ngày lớn nhất
  của cột NGAY_KQ trong file, đổ màu tím (#7C3AED) để nhận diện giờ tạm tính.
- BN có NGAY_RA nhưng trống NGAY_KQ: tách cảnh báo bảng phụ riêng theo ngày.
- Cảnh báo đỏ trên thanh ngày khi có khoa vừa còn giường trống vừa nằm ghép/trùng.
- Bắt trùng đầy đủ theo khoảng giờ trong ngày:
  + 2 BN khác nhau trùng giờ trên 1 giường (phân biệt Nam - Nữ nằm chung).
  + 1 BN tự trùng giờ trên cùng 1 giường (nhiều lượt chồng chéo).
  + 1 BN nằm 2 giường khác nhau trùng giờ (kể cả khác khoa).
  + Nhiều BN / nhiều lượt không trùng giờ -> thẻ màu xanh, hiện đủ số lượng & khung giờ.
- Giường sẵn sàng có 2 loại:
  + Có BN nhưng ra trước 23:59 -> hiện "🟢 Sẵn sàng từ: HH:MM" (giờ ra + 1 phút).
  + Không có BN nào trong ngày -> hiện "Trống" và "Sẵn sàng kê".
- Cảnh báo hở khung giờ nằm giường trong ngày (>= 60 phút).
- Cổng chờ cột Số lượng (SO_LUONG / SL): hiện cạnh khung giờ & cảnh báo ngược SL
  (lượt ngắn hơn tính SL > 0 trong khi lượt dài nhất trong ngày tính SL = 0).
=======================================================================
"""

import os
import html
import pickle
from io import BytesIO
from datetime import datetime, date, time as dt_time, timedelta

import pandas as pd
import streamlit as st

from checker import read_input_dataframe

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BED_DATA_DIR = os.path.join(BASE_DIR, "data_ref", "bed_module")
os.makedirs(BED_DATA_DIR, exist_ok=True)
BED_LATEST_FILE = os.path.join(BED_DATA_DIR, "_bed_latest.pkl")

GAP_WARNING_MINUTES = 60  # Ngưỡng cảnh báo hở khung giờ trong ngày (phút)


# ======================= TIỆN ÍCH CHUẨN HÓA =======================
def _clean_str(val):
    s = str(val or "").strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _normalize_gender(val):
    s = _clean_str(val).lower()
    if s in ("1", "nam", "m", "male"):
        return "Nam"
    if s in ("2", "nữ", "nu", "f", "female"):
        return "Nữ"
    return _clean_str(val)


def _extract_birth_year(nam_sinh):
    s = _clean_str(nam_sinh)
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 4:
        return digits[:4]
    return ""


def _calc_age_str(nam_sinh):
    yr_str = _extract_birth_year(nam_sinh)
    if yr_str:
        try:
            age = datetime.now().year - int(yr_str)
            if 0 <= age <= 120:
                return f"{age} tuổi"
        except Exception:
            pass
    return ""


def _format_sl_str(val):
    """Chuẩn hóa hiển thị số lượng (1.0 -> '1', 0.5 -> '0.5', trống -> '')."""
    s = _clean_str(val).replace(",", ".")
    if not s:
        return "", None
    try:
        num = float(s)
        if num.is_integer():
            return str(int(num)), num
        return f"{num:g}", num
    except Exception:
        return s, None


def _format_patient_short_label(rec):
    """Định dạng gọn cho bảng trùng: TÊN BN (Năm sinh - Giới tính), bỏ mã LK dài."""
    yr = _extract_birth_year(rec.get("NAM_SINH", ""))
    gt = rec.get("GIOI_TINH", "")
    meta = " - ".join([x for x in [yr, gt] if x])
    if meta:
        return f"{rec['TENBENHNHAN']} ({meta})"
    return rec["TENBENHNHAN"]


def _get_display_interval(rec, day_start, end_of_day_limit):
    """
    Với BN tạm tính (chưa có ngày kết thúc): cắt rải theo đúng ngày đang xem.
    Với BN đã chốt ngày kết thúc: giữ nguyên START_DT -> END_DT thực tế.
    """
    if rec.get("IS_ESTIMATED_END"):
        d_s = max(rec["START_DT"], day_start)
        d_e = min(rec["END_DT"], end_of_day_limit)
        if d_e < d_s:
            d_e = d_s
        return d_s, d_e
    return rec["START_DT"], rec["END_DT"]


def _format_card_time_html(start_dt, end_dt, is_est_end=False, sl_str=""):
    """
    Định dạng dd/mm HH:MM -> dd/mm HH:MM (+ SL nếu có).
    - Giờ tạm tính: đổ màu tím đậm (#7C3AED) để nhận diện chưa chốt ngày kết thúc.
    - Giờ bắt đầu khác 00:00 -> đỏ đậm (#DC2626).
    - Giờ kết thúc (đã chốt) khác 23:59 -> đỏ đậm (#DC2626).
    """
    s_hm = start_dt.strftime("%H:%M")
    e_hm = end_dt.strftime("%H:%M")
    s_txt = start_dt.strftime("%d/%m %H:%M")
    e_txt = end_dt.strftime("%d/%m %H:%M")

    if is_est_end:
        if s_hm != "00:00":
            s_html = f"<span style='color:#DC2626;font-weight:700;'>{s_txt}</span>"
        else:
            s_html = f"<span style='color:#7C3AED;font-weight:700;'>{s_txt}</span>"
        e_html = f"<span style='color:#7C3AED;font-weight:700;'>{e_txt} (tạm)</span>"
    else:
        if s_hm != "00:00":
            s_html = f"<span style='color:#DC2626;font-weight:700;'>{s_txt}</span>"
        else:
            s_html = f"<span>{s_txt}</span>"

        if e_hm != "23:59":
            e_html = f"<span style='color:#DC2626;font-weight:700;'>{e_txt}</span>"
        else:
            e_html = f"<span>{e_txt}</span>"

    sl_html = f" <span style='color:#0F172A;font-weight:700;'>({html.escape(sl_str)})</span>" if sl_str else ""
    return f"{s_html} → {e_html}{sl_html}"


def _format_plain_interval(rec, day_start, end_of_day_limit):
    d_s, d_e = _get_display_interval(rec, day_start, end_of_day_limit)
    est_tag = " (tạm)" if rec.get("IS_ESTIMATED_END") else ""
    sl_tag = f" (SL:{rec['SO_LUONG_STR']})" if rec.get("SO_LUONG_STR") else ""
    return f"{d_s.strftime('%d/%m %H:%M')} → {d_e.strftime('%d/%m %H:%M')}{est_tag}{sl_tag}"


def _short_dept_name(ma_khoa, ten_khoa):
    mapping = {
        "K282930": "LCK",
        "K1631": "YHCT",
        "K192726": "Ngoại-Sản",
        "K02031813": "Nội-Nhi",
    }
    if ma_khoa in mapping:
        return mapping[ma_khoa]
    s = str(ten_khoa or ma_khoa).replace("Khoa ", "").strip()
    return s[:15] + "…" if len(s) > 15 else s


def _format_overlap_desc(n_mixed, n_multi_bed, n_self, n_time_only):
    parts = []
    if n_mixed > 0:
        parts.append(f"{n_mixed} trùng Nam-Nữ")
    if n_multi_bed > 0:
        parts.append(f"{n_multi_bed} BN nằm 2 giường")
    if n_self > 0:
        parts.append(f"{n_self} tự trùng giờ")
    if n_time_only > 0:
        parts.append(f"{n_time_only} trùng TG")
    return ", ".join(parts)


# ======================= ĐỌC DANH MỤC GIƯỜNG =======================
@st.cache_data(show_spinner=False)
def _load_bed_catalog_cached(ref_path, mtime):
    xl = pd.ExcelFile(ref_path)

    khoa_name_map = {}
    df_giuong_pd = pd.DataFrame()
    if "data_giuong" in xl.sheet_names:
        df_g = xl.parse("data_giuong", dtype=str).fillna("")
        if "MA_KHOA" in df_g.columns:
            for _, r in df_g.iterrows():
                mk = _clean_str(r.get("MA_KHOA"))
                tk = _clean_str(r.get("TEN_KHOA"))
                if mk and tk:
                    khoa_name_map[mk] = tk
        if "GIUONG_PD" in df_g.columns:
            df_g["GIUONG_PD_NUM"] = pd.to_numeric(df_g["GIUONG_PD"], errors="coerce").fillna(0).astype(int)
            df_g = df_g[df_g["GIUONG_PD_NUM"] > 0].copy()
            if "HIEU_LUC" in df_g.columns:
                valid_hl = df_g["HIEU_LUC"].astype(str).str.strip().str.lower()
                df_g = df_g[valid_hl.isin(["có", "co", "1", "true", ""])]
            df_g["MA_KHOA"] = df_g["MA_KHOA"].map(_clean_str)
            df_g["TEN_KHOA"] = df_g["TEN_KHOA"].map(_clean_str)
            df_giuong_pd = df_g.drop_duplicates(subset=["MA_KHOA"], keep="last").reset_index(drop=True)

    if "dm_giuong_chitiet" in xl.sheet_names:
        df_ct = xl.parse("dm_giuong_chitiet", dtype=str).fillna("")
        df_ct.columns = [str(c).strip().upper() for c in df_ct.columns]

        col_ma = next((c for c in ["MÃ", "MA", "MA_GIUONG", "MÃ GIƯỜNG"] if c in df_ct.columns), None)
        col_khoa = next((c for c in ["MÃ KHOA", "MA_KHOA"] if c in df_ct.columns), None)
        col_ten = next((c for c in ["TÊN GIƯỜNG", "TEN_GIUONG"] if c in df_ct.columns), None)
        col_loai = next((c for c in ["LOẠI", "LOAI", "LOAI_GIUONG"] if c in df_ct.columns), None)
        col_phong = next((c for c in ["PHÒNG", "PHONG", "TEN_PHONG"] if c in df_ct.columns), None)

        if col_ma and col_khoa:
            grouped = {}
            for _, r in df_ct.iterrows():
                mk = _clean_str(r.get(col_khoa))
                code = _clean_str(r.get(col_ma)).upper()
                if not mk or not code:
                    continue
                grouped.setdefault(mk, [])
                if not any(b["code"] == code for b in grouped[mk]):
                    grouped[mk].append({
                        "code": code,
                        "name": _clean_str(r.get(col_ten)) if col_ten else "",
                        "loai": _clean_str(r.get(col_loai)) if col_loai else "",
                        "phong": _clean_str(r.get(col_phong)) if col_phong else "",
                    })

            if grouped:
                catalog = []
                ordered_khoas = list(df_giuong_pd["MA_KHOA"]) if not df_giuong_pd.empty else []
                for mk in grouped:
                    if mk not in ordered_khoas:
                        ordered_khoas.append(mk)

                for mk in ordered_khoas:
                    if mk not in grouped:
                        continue
                    beds_sorted = sorted(grouped[mk], key=lambda x: x["code"])
                    catalog.append({
                        "MA_KHOA": mk,
                        "TEN_KHOA": khoa_name_map.get(mk, mk),
                        "BEDS": beds_sorted,
                        "SOURCE": "dm_giuong_chitiet",
                    })
                return catalog

    catalog = []
    for _, r in df_giuong_pd.iterrows():
        n_beds = int(r["GIUONG_PD_NUM"])
        catalog.append({
            "MA_KHOA": r["MA_KHOA"],
            "TEN_KHOA": r["TEN_KHOA"] or r["MA_KHOA"],
            "BEDS": [{"code": f"H{i:03d}", "name": "", "loai": "", "phong": ""} for i in range(1, n_beds + 1)],
            "SOURCE": "data_giuong",
        })
    return catalog


def load_bed_catalog(ref_path):
    if not ref_path or not os.path.exists(ref_path):
        return []
    try:
        return _load_bed_catalog_cached(ref_path, os.path.getmtime(ref_path))
    except Exception:
        return []


# ======================= LƯU / TẢI DỮ LIỆU GIƯỜNG =======================
def save_bed_dataset(df_bed, src_filename):
    payload = {
        "df_bed": df_bed,
        "src_filename": src_filename,
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    tmp = BED_LATEST_FILE + ".tmp"
    with open(tmp, "wb") as fh:
        pickle.dump(payload, fh)
    os.replace(tmp, BED_LATEST_FILE)


def load_bed_dataset():
    if not os.path.exists(BED_LATEST_FILE):
        return None
    try:
        with open(BED_LATEST_FILE, "rb") as fh:
            return pickle.load(fh)
    except Exception:
        return None


# ======================= CHUẨN HÓA NGÀY & LỌC NHÓM 15 =======================
def parse_bhyt_datetime(val, is_end=False):
    s = _clean_str(val)
    if not s:
        return None
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    try:
        if len(digits) >= 12 and s.isdigit():
            return datetime.strptime(digits[:12], "%Y%m%d%H%M")
        if len(digits) == 8 and s.isdigit():
            d = datetime.strptime(digits, "%Y%m%d").date()
            return datetime.combine(d, dt_time(23, 59) if is_end else dt_time(0, 0))
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        return None
    return None


def extract_bed_records(df_raw):
    if df_raw is None or df_raw.empty or "MA_NHOM" not in df_raw.columns:
        return pd.DataFrame()

    df = df_raw.copy().fillna("")
    ma_nhom = df["MA_NHOM"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df = df[ma_nhom == "15"].copy()
    if df.empty:
        return df

    # Nhận diện cột Số lượng nếu file có (cổng chờ: SO_LUONG / SL)
    col_map_upper = {str(c).strip().upper(): c for c in df.columns}
    sl_col = None
    for cand in ["SO_LUONG", "SL", "SOLUONG", "SỐ LƯỢNG"]:
        if cand in col_map_upper:
            sl_col = col_map_upper[cand]
            break

    for col in [
        "MA_LK", "TENBENHNHAN", "NAM_SINH", "GIOI_TINH", "MA_BENH",
        "MA_KHOA", "TEN_KHOA", "MA_DICH_VU", "TEN_DICH_VU", "TYLE_TT",
        "MA_GIUONG_CV130", "MA_GIUONG",
        "NGAY_YL", "NGAY_TH_YL", "NGAY_KQ", "NGAY_VAO", "NGAY_RA",
    ]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(_clean_str)

    if sl_col is not None:
        sl_parsed = df[sl_col].map(_format_sl_str)
        df["SO_LUONG_STR"] = [x[0] for x in sl_parsed]
        df["SO_LUONG_NUM"] = [x[1] for x in sl_parsed]
    else:
        df["SO_LUONG_STR"] = ""
        df["SO_LUONG_NUM"] = None

    df["GIOI_TINH_STD"] = df["GIOI_TINH"].map(_normalize_gender)
    df["MA_GIUONG_STD"] = df["MA_GIUONG_CV130"].where(
        df["MA_GIUONG_CV130"] != "", df["MA_GIUONG"]
    ).str.upper()

    df["TYLE_TT_NUM"] = pd.to_numeric(df["TYLE_TT"], errors="coerce").fillna(100)

    # 1. Tìm mốc ngày kết quả lớn nhất của toàn bộ file (để rải cho BN đang điều trị)
    parsed_kq = [parse_bhyt_datetime(v, is_end=True) for v in df["NGAY_KQ"]]
    valid_kq = [d for d in parsed_kq if d is not None]

    parsed_starts = []
    for _, r in df.iterrows():
        st_dt = (
            parse_bhyt_datetime(r["NGAY_TH_YL"], is_end=False)
            or parse_bhyt_datetime(r["NGAY_YL"], is_end=False)
            or parse_bhyt_datetime(r["NGAY_VAO"], is_end=False)
        )
        parsed_starts.append(st_dt)

    if valid_kq:
        max_kq_date = max(valid_kq).date()
        fallback_end_dt = datetime.combine(max_kq_date, dt_time(23, 59))
    else:
        valid_st = [d for d in parsed_starts if d is not None]
        max_st_date = max(valid_st).date() if valid_st else date.today()
        fallback_end_dt = datetime.combine(max_st_date, dt_time(23, 59))

    # 2. Xác định START_DT, END_DT và các cờ cảnh báo
    ends = []
    missing_kq_has_ra_flags = []
    is_est_end_flags = []

    for idx_i, (_, r) in enumerate(df.iterrows()):
        st_dt = parsed_starts[idx_i]
        kq_dt = parsed_kq[idx_i]
        ra_dt = parse_bhyt_datetime(r["NGAY_RA"], is_end=True)

        if kq_dt is not None:
            en_dt = kq_dt
            missing_kq_has_ra_flags.append(False)
            is_est_end_flags.append(False)
        elif ra_dt is not None:
            en_dt = ra_dt
            missing_kq_has_ra_flags.append(True)
            is_est_end_flags.append(False)
        else:
            en_dt = fallback_end_dt
            missing_kq_has_ra_flags.append(False)
            is_est_end_flags.append(True)

        if st_dt and en_dt and en_dt < st_dt:
            en_dt = st_dt
        ends.append(en_dt)

    df["START_DT"] = parsed_starts
    df["END_DT"] = ends
    df["MISSING_KQ_HAS_RA"] = missing_kq_has_ra_flags
    df["IS_ESTIMATED_END"] = is_est_end_flags
    df = df[df["START_DT"].notna()].copy()
    return df.reset_index(drop=True)


# ======================= KIỂM TRA TRÙNG GIƯỜNG TRONG 1 NGÀY =======================
def _get_overlap_window(s1, e1, s2, e2, win_start=None, win_end=None):
    ov_s = max(s1, s2)
    ov_e = min(e1, e2)
    if win_start is not None:
        ov_s = max(ov_s, win_start)
    if win_end is not None:
        ov_e = min(ov_e, win_end)
    if ov_s < ov_e:
        return ov_s, ov_e
    return None


def _is_mixed_gender(g1, g2):
    return {g1, g2} == {"Nam", "Nữ"}


def _group_bed_patients(occupants, day_start, end_of_day_limit):
    """
    Gom các lượt nằm trên cùng 1 giường theo bệnh nhân (KHÔNG ẩn lượt thứ 2).
    Mỗi bệnh nhân giữ đủ danh sách `slots` sắp theo thời gian tăng dần.
    """
    grouped = []
    by_key = {}
    for o in sorted(occupants, key=lambda x: (x["START_DT"], x["END_DT"])):
        k = o["MA_LK"] or o["TENBENHNHAN"]
        d_s, d_e = _get_display_interval(o, day_start, end_of_day_limit)
        slot_item = {
            "START_DT": o["START_DT"],
            "END_DT": o["END_DT"],
            "DISP_START": d_s,
            "DISP_END": d_e,
            "IS_ESTIMATED_END": o.get("IS_ESTIMATED_END", False),
            "MISSING_KQ_HAS_RA": o.get("MISSING_KQ_HAS_RA", False),
            "TEN_DICH_VU": o.get("TEN_DICH_VU", ""),
            "MA_DICH_VU": o.get("MA_DICH_VU", ""),
            "TYLE_TT": o.get("TYLE_TT", 100),
            "SO_LUONG_STR": o.get("SO_LUONG_STR", ""),
            "SO_LUONG_NUM": o.get("SO_LUONG_NUM", None),
        }
        if k not in by_key:
            p_obj = {
                "KEY": k,
                "MA_LK": o["MA_LK"],
                "TENBENHNHAN": o["TENBENHNHAN"],
                "NAM_SINH": o.get("NAM_SINH", ""),
                "GIOI_TINH": o.get("GIOI_TINH", ""),
                "MA_BENH": o.get("MA_BENH", ""),
                "TRANG_THAI_BN": o.get("TRANG_THAI_BN", ""),
                "slots": [slot_item],
                "other_bed_conflicts": [],
            }
            by_key[k] = p_obj
            grouped.append(p_obj)
        else:
            by_key[k]["slots"].append(slot_item)
            if o.get("TRANG_THAI_BN") == "Ra viện":
                by_key[k]["TRANG_THAI_BN"] = "Ra viện"
    return grouped


def build_bed_map_for_day(bed_catalog, df_day, day_start, day_end, day_str="", df_all=None):
    """
    Tính sơ đồ giường & toàn bộ cảnh báo cho ĐÚNG 1 NGÀY:
    - Trùng giờ 2 BN trên cùng giường (kể cả Nam - Nữ nằm chung)
    - Tự trùng giờ của 1 BN trên cùng 1 giường
    - 1 BN nằm 2 giường khác nhau trùng giờ (kể cả khác khoa)
    - BN chưa gán mã giường
    - BN có ngày ra nhưng trống NGAY_KQ
    - Hở khung giờ nằm giường trong ngày (>= GAP_WARNING_MINUTES)
    - Ngược Số lượng giường (giờ ngắn tính SL > 0, giờ dài nhất trong ngày tính SL = 0)
    """
    dept_results = []
    detail_rows = []
    overlap_rows = []
    unassigned_rows = []
    missing_kq_rows = []
    gap_rows = []
    sl_warning_rows = []

    end_of_day_limit = datetime.combine(day_end.date(), dt_time(23, 59))
    khoa_name_lookup = {d["MA_KHOA"]: d["TEN_KHOA"] for d in bed_catalog}

    # Lưu toàn bộ lượt nằm trong ngày theo từng BN để soi:
    # (1) 1 BN nằm 2 giường trùng giờ, (2) hở khung giờ >= 60 phút, (3) ngược SL
    patient_all_recs = {}

    for dept in bed_catalog:
        ma_khoa = dept["MA_KHOA"]
        ten_khoa = dept["TEN_KHOA"]
        bed_meta = {b["code"]: b for b in dept["BEDS"]}
        bed_codes = [b["code"] for b in dept["BEDS"]]
        approved_beds_count = len(bed_codes)
        beds = {code: [] for code in bed_codes}
        dept_unique_patients = set()

        sub = df_day[df_day["MA_KHOA"] == ma_khoa].copy()
        if not sub.empty:
            sub = sub.sort_values(by=["START_DT", "MA_LK"], ascending=True)

            for _, row in sub.iterrows():
                is_missing_kq_ra = bool(row.get("MISSING_KQ_HAS_RA", False))
                is_est_end = bool(row.get("IS_ESTIMATED_END", False))
                sl_str = _clean_str(row.get("SO_LUONG_STR", ""))
                sl_num = row.get("SO_LUONG_NUM", None)
                if pd.isna(sl_num):
                    sl_num = None

                rec = {
                    "MA_LK": row["MA_LK"],
                    "TENBENHNHAN": row["TENBENHNHAN"] or f"BN {row['MA_LK']}",
                    "NAM_SINH": row.get("NAM_SINH", ""),
                    "GIOI_TINH": row.get("GIOI_TINH_STD", ""),
                    "MA_BENH": row.get("MA_BENH", ""),
                    "MA_KHOA": ma_khoa,
                    "TEN_KHOA": ten_khoa,
                    "MA_DICH_VU": row["MA_DICH_VU"],
                    "TEN_DICH_VU": row["TEN_DICH_VU"],
                    "TYLE_TT": row["TYLE_TT_NUM"],
                    "START_DT": row["START_DT"],
                    "END_DT": row["END_DT"],
                    "TRANG_THAI_BN": "Ra viện" if row["NGAY_RA"] else "Đang điều trị",
                    "MISSING_KQ_HAS_RA": is_missing_kq_ra,
                    "IS_ESTIMATED_END": is_est_end,
                    "SO_LUONG_STR": sl_str,
                    "SO_LUONG_NUM": sl_num,
                }
                p_key = rec["MA_LK"] or rec["TENBENHNHAN"]
                dept_unique_patients.add(p_key)

                real_code = row.get("MA_GIUONG_STD", "").strip()
                rec["BED_CODE"] = real_code
                patient_all_recs.setdefault(p_key, []).append(rec)

                d_s, d_e = _get_display_interval(rec, day_start, end_of_day_limit)
                end_disp_str = d_e.strftime("%d/%m/%Y %H:%M") + (" (tạm tính)" if is_est_end else "")

                if real_code:
                    if real_code not in beds:
                        beds[real_code] = []
                        bed_meta[real_code] = {"code": real_code, "name": "Ngoài danh mục", "loai": "", "phong": ""}
                    beds[real_code].append(rec)
                    assigned_code = real_code
                    source_tag = "MA_GIUONG_CV130"
                else:
                    assigned_code = ""
                    source_tag = "Chưa gán"
                    unassigned_rows.append({
                        "Ngày": day_str,
                        "Mã khoa": ma_khoa,
                        "Tên khoa": ten_khoa,
                        "Tên bệnh nhân": _format_patient_short_label(rec),
                        "Trạng thái BN": rec["TRANG_THAI_BN"],
                        "Tên dịch vụ giường": rec["TEN_DICH_VU"],
                        "Từ thời điểm": d_s.strftime("%d/%m/%Y %H:%M"),
                        "Đến thời điểm": end_disp_str,
                    })

                if is_missing_kq_ra:
                    missing_kq_rows.append({
                        "Ngày": day_str,
                        "Mã khoa": ma_khoa,
                        "Tên khoa": ten_khoa,
                        "Mã giường": assigned_code or "(Chưa gán)",
                        "Tên bệnh nhân": _format_patient_short_label(rec),
                        "Ngày bắt đầu": rec["START_DT"].strftime("%d/%m/%Y %H:%M"),
                        "Ngày ra viện (NGAY_RA)": rec["END_DT"].strftime("%d/%m/%Y %H:%M"),
                        "Cảnh báo": "Đã có Ngày ra nhưng trống NGAY_KQ",
                    })

                meta = bed_meta.get(assigned_code, {})
                det_row = {
                    "Ngày": day_str,
                    "Mã khoa": ma_khoa,
                    "Tên khoa": ten_khoa,
                    "Phòng": meta.get("phong", ""),
                    "Mã giường": assigned_code or "(Chưa gán)",
                    "Nguồn mã": source_tag,
                    "Mã LK": rec["MA_LK"],
                    "Tên bệnh nhân": rec["TENBENHNHAN"],
                    "Năm sinh": _extract_birth_year(rec["NAM_SINH"]),
                    "Giới tính": rec["GIOI_TINH"],
                    "Trạng thái BN": rec["TRANG_THAI_BN"],
                    "Tên dịch vụ giường": rec["TEN_DICH_VU"],
                    "Tỷ lệ TT (%)": int(rec["TYLE_TT"]),
                    "Từ thời điểm": d_s.strftime("%d/%m/%Y %H:%M"),
                    "Đến thời điểm": end_disp_str,
                }
                if sl_str != "":
                    det_row["Số lượng"] = sl_str
                detail_rows.append(det_row)

        bed_cards = []
        cnt_empty = cnt_single = cnt_shared = 0
        for code, occupants in beds.items():
            grouped_patients = _group_bed_patients(occupants, day_start, end_of_day_limit)

            has_low_rate = any(o["TYLE_TT"] < 100 for o in occupants)
            has_time_overlap = False
            has_mixed_gender = False
            has_self_overlap = False
            seen_pairs = set()

            for i in range(len(occupants)):
                for j in range(i + 1, len(occupants)):
                    p1, p2 = occupants[i], occupants[j]
                    k1 = p1["MA_LK"] or p1["TENBENHNHAN"]
                    k2 = p2["MA_LK"] or p2["TENBENHNHAN"]

                    ov = _get_overlap_window(
                        p1["START_DT"], p1["END_DT"],
                        p2["START_DT"], p2["END_DT"],
                        day_start, day_end,
                    )
                    if ov:
                        has_time_overlap = True
                        is_same_bn = (k1 == k2)
                        if is_same_bn:
                            has_self_overlap = True
                            mixed = False
                            warn_label = "⚠️ Tự trùng giờ trên cùng giường"
                        else:
                            mixed = _is_mixed_gender(p1.get("GIOI_TINH", ""), p2.get("GIOI_TINH", ""))
                            if mixed:
                                has_mixed_gender = True
                            warn_label = "🚨 NAM - NỮ NẰM CHUNG" if mixed else "⚠️ Trùng giờ giường"

                        pair_key = (code, tuple(sorted([k1, k2])), ov[0], ov[1], warn_label)
                        if pair_key not in seen_pairs:
                            seen_pairs.add(pair_key)
                            mins = int((ov[1] - ov[0]).total_seconds() // 60)
                            overlap_rows.append({
                                "Ngày": day_str,
                                "Cảnh báo": warn_label,
                                "Mã khoa": ma_khoa,
                                "Tên khoa": ten_khoa,
                                "Phòng": bed_meta.get(code, {}).get("phong", ""),
                                "Mã giường": code,
                                "BN 1": _format_patient_short_label(p1),
                                "Giờ BN 1": _format_plain_interval(p1, day_start, end_of_day_limit),
                                "TT BN 1 (%)": int(p1["TYLE_TT"]),
                                "BN 2": _format_patient_short_label(p2),
                                "Giờ BN 2": _format_plain_interval(p2, day_start, end_of_day_limit),
                                "TT BN 2 (%)": int(p2["TYLE_TT"]),
                                "Khoảng giờ trùng": f"{ov[0].strftime('%d/%m/%Y %H:%M')} → {ov[1].strftime('%d/%m/%Y %H:%M')}",
                                "Số phút trùng": mins,
                            })

            # Tính thời điểm sẵn sàng trong ngày nếu BN ra trước 23:59 (cộng 1 phút)
            ready_after_dt = None
            if not occupants:
                status = "empty"
                cnt_empty += 1
            else:
                max_end_dt = max(o["END_DT"] for o in occupants)
                if max_end_dt < end_of_day_limit:
                    ready_after_dt = max_end_dt + timedelta(minutes=1)
                if has_time_overlap or has_low_rate:
                    status = "shared"
                    cnt_shared += 1
                else:
                    status = "single"
                    cnt_single += 1

            meta = bed_meta.get(code, {})
            bed_cards.append({
                "code": code,
                "name": meta.get("name", ""),
                "loai": meta.get("loai", ""),
                "phong": meta.get("phong", ""),
                "status": status,
                "patients": grouped_patients,
                "n_patients": len(grouped_patients),
                "has_overlap": has_time_overlap,
                "has_mixed_gender": has_mixed_gender,
                "has_self_overlap": has_self_overlap,
                "has_multi_bed_conflict": False,
                "ready_after_dt": ready_after_dt,
            })

        dept_results.append({
            "ma_khoa": ma_khoa,
            "ten_khoa": ten_khoa,
            "total": len(bed_cards),
            "total_approved": approved_beds_count,
            "actual_patients": len(dept_unique_patients),
            "empty": cnt_empty,
            "single": cnt_single,
            "shared": cnt_shared,
            "beds": bed_cards,
        })

    # ----- QUÉT THEO TỪNG BỆNH NHÂN TRONG NGÀY -----
    # (1) 1 BN nằm 2 giường khác nhau trùng giờ
    # (2) Hở khung giờ nằm giường >= GAP_WARNING_MINUTES
    # (3) Ngược Số lượng giường (giờ ngắn tính SL > 0, giờ dài nhất tính 0)
    multi_bed_tags = {}  # (ma_khoa, bed_code, p_key) -> set of conflict labels
    seen_multi_bed_pairs = set()

    for p_key, recs in patient_all_recs.items():
        if len(recs) < 2:
            continue

        # (1) 1 BN nằm 2 giường khác nhau trùng giờ
        assigned_recs = [r for r in recs if r.get("BED_CODE")]
        for i in range(len(assigned_recs)):
            for j in range(i + 1, len(assigned_recs)):
                r1, r2 = assigned_recs[i], assigned_recs[j]
                b1 = (r1["MA_KHOA"], r1["BED_CODE"])
                b2 = (r2["MA_KHOA"], r2["BED_CODE"])
                if b1 == b2:
                    continue
                ov = _get_overlap_window(
                    r1["START_DT"], r1["END_DT"],
                    r2["START_DT"], r2["END_DT"],
                    day_start, day_end,
                )
                if ov:
                    pair_key = (p_key, tuple(sorted([b1, b2])), ov[0], ov[1])
                    if pair_key not in seen_multi_bed_pairs:
                        seen_multi_bed_pairs.add(pair_key)
                        mins = int((ov[1] - ov[0]).total_seconds() // 60)
                        lbl_bed1 = r1["BED_CODE"] if r1["MA_KHOA"] == r2["MA_KHOA"] else f"{r1['BED_CODE']} ({_short_dept_name(r1['MA_KHOA'], r1['TEN_KHOA'])})"
                        lbl_bed2 = r2["BED_CODE"] if r1["MA_KHOA"] == r2["MA_KHOA"] else f"{r2['BED_CODE']} ({_short_dept_name(r2['MA_KHOA'], r2['TEN_KHOA'])})"

                        multi_bed_tags.setdefault((r1["MA_KHOA"], r1["BED_CODE"], p_key), set()).add(lbl_bed2)
                        multi_bed_tags.setdefault((r2["MA_KHOA"], r2["BED_CODE"], p_key), set()).add(lbl_bed1)

                        mk_disp = r1["MA_KHOA"] if r1["MA_KHOA"] == r2["MA_KHOA"] else f"{r1['MA_KHOA']} ↔ {r2['MA_KHOA']}"
                        tk_disp = r1["TEN_KHOA"] if r1["MA_KHOA"] == r2["MA_KHOA"] else f"{r1['TEN_KHOA']} ↔ {r2['TEN_KHOA']}"
                        overlap_rows.append({
                            "Ngày": day_str,
                            "Cảnh báo": "🚨 1 BN NẰM 2 GIƯỜNG TRÙNG GIỜ",
                            "Mã khoa": mk_disp,
                            "Tên khoa": tk_disp,
                            "Phòng": "",
                            "Mã giường": f"{r1['BED_CODE']} ↔ {r2['BED_CODE']}",
                            "BN 1": f"{_format_patient_short_label(r1)} [{r1['BED_CODE']}]",
                            "Giờ BN 1": _format_plain_interval(r1, day_start, end_of_day_limit),
                            "TT BN 1 (%)": int(r1["TYLE_TT"]),
                            "BN 2": f"{_format_patient_short_label(r2)} [{r2['BED_CODE']}]",
                            "Giờ BN 2": _format_plain_interval(r2, day_start, end_of_day_limit),
                            "TT BN 2 (%)": int(r2["TYLE_TT"]),
                            "Khoảng giờ trùng": f"{ov[0].strftime('%d/%m/%Y %H:%M')} → {ov[1].strftime('%d/%m/%Y %H:%M')}",
                            "Số phút trùng": mins,
                        })

        # (2) Hở khung giờ nằm giường trong ngày (>= GAP_WARNING_MINUTES)
        clipped = []
        for r in recs:
            cs = max(r["START_DT"], day_start)
            ce = min(r["END_DT"], end_of_day_limit)
            if ce >= cs:
                clipped.append((cs, ce, r))
        clipped.sort(key=lambda x: (x[0], x[1]))

        if len(clipped) >= 2:
            merged = []
            cur_s, cur_e, cur_r_last = clipped[0]
            for cs, ce, r in clipped[1:]:
                if cs <= cur_e:
                    if ce > cur_e:
                        cur_e = ce
                        cur_r_last = r
                else:
                    merged.append((cur_s, cur_e, cur_r_last))
                    cur_s, cur_e, cur_r_last = cs, ce, r
            merged.append((cur_s, cur_e, cur_r_last))

            for idx_m in range(len(merged) - 1):
                _, prev_e, prev_r = merged[idx_m]
                next_s, _, next_r = merged[idx_m + 1]
                gap_mins = int((next_s - prev_e).total_seconds() // 60)
                if gap_mins >= GAP_WARNING_MINUTES:
                    gap_rows.append({
                        "Ngày": day_str,
                        "Mã khoa": prev_r["MA_KHOA"],
                        "Tên khoa": prev_r["TEN_KHOA"],
                        "Tên bệnh nhân": _format_patient_short_label(prev_r),
                        "Giường trước": prev_r.get("BED_CODE") or "(Chưa gán)",
                        "Giường sau": next_r.get("BED_CODE") or "(Chưa gán)",
                        "Khoảng trống": f"{prev_e.strftime('%H:%M')} → {next_s.strftime('%H:%M')}",
                        "Số phút trống": gap_mins,
                        "Cảnh báo": f"⚠️ Hở khung giờ ({gap_mins} phút ≥ {GAP_WARNING_MINUTES}p)",
                    })

        # (3) Cổng chờ bắt lỗi ngược Số lượng (khi có cột SO_LUONG / SL)
        sl_recs = [r for r in recs if r.get("SO_LUONG_NUM") is not None]
        if len(sl_recs) >= 2:
            dur_list = []
            for r in sl_recs:
                cs = max(r["START_DT"], day_start)
                ce = min(r["END_DT"], end_of_day_limit)
                dur = max(0, int((ce - cs).total_seconds() // 60))
                dur_list.append((dur, r["SO_LUONG_NUM"], r))

            max_dur = max(x[0] for x in dur_list)
            longest_zero = [x for x in dur_list if x[0] == max_dur and x[1] == 0]
            shorter_pos = [x for x in dur_list if x[0] < max_dur and x[1] > 0]
            if longest_zero and shorter_pos:
                ref_r = sl_recs[0]
                detail_parts = []
                for dur, sl_v, r in sorted(dur_list, key=lambda x: x[2]["START_DT"]):
                    cs = max(r["START_DT"], day_start)
                    ce = min(r["END_DT"], end_of_day_limit)
                    b_lbl = r.get("BED_CODE") or "Chưa gán"
                    detail_parts.append(f"{b_lbl} {cs.strftime('%H:%M')}-{ce.strftime('%H:%M')} ({dur}p, SL:{r['SO_LUONG_STR']})")
                sl_warning_rows.append({
                    "Ngày": day_str,
                    "Mã khoa": ref_r["MA_KHOA"],
                    "Tên khoa": ref_r["TEN_KHOA"],
                    "Tên bệnh nhân": _format_patient_short_label(ref_r),
                    "Chi tiết các lượt trong ngày": " | ".join(detail_parts),
                    "Cảnh báo": "⚠️ Ngược SL giường (lượt ngắn SL > 0, lượt dài nhất SL = 0)",
                })

    # Gắn tag cảnh báo 1 BN nằm 2 giường lên từng thẻ giường liên quan
    if multi_bed_tags:
        for d in dept_results:
            mk = d["ma_khoa"]
            for b in d["beds"]:
                bed_has_mb = False
                for p in b["patients"]:
                    others = multi_bed_tags.get((mk, b["code"], p["KEY"]))
                    if others:
                        p["other_bed_conflicts"] = sorted(others)
                        bed_has_mb = True
                if bed_has_mb:
                    b["has_multi_bed_conflict"] = True
                    b["has_overlap"] = True
                    if b["status"] == "single":
                        b["status"] = "shared"
                        d["single"] = max(0, d["single"] - 1)
                        d["shared"] += 1

    df_ov = pd.DataFrame(overlap_rows)
    if not df_ov.empty:
        def _ov_prio(msg):
            s = str(msg)
            if "NAM - NỮ" in s:
                return 0
            if "2 GIƯỜNG" in s:
                return 1
            if "Tự trùng" in s:
                return 2
            return 3
        df_ov["_prio"] = df_ov["Cảnh báo"].apply(_ov_prio)
        df_ov = df_ov.sort_values(by=["_prio", "Mã khoa", "Mã giường"]).drop(columns=["_prio"]).reset_index(drop=True)

    return (
        dept_results,
        pd.DataFrame(detail_rows),
        df_ov,
        pd.DataFrame(unassigned_rows),
        pd.DataFrame(missing_kq_rows),
        pd.DataFrame(gap_rows),
        pd.DataFrame(sl_warning_rows),
    )


def _build_excel_bytes(
    df_summary,
    df_overlap,
    df_detail,
    df_unassigned=None,
    df_missing_kq=None,
    df_gaps=None,
    df_sl_warn=None,
):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="TongHop_TheoNgay", index=False)
        df_overlap.to_excel(writer, sheet_name="Trung_Giuong", index=False)
        df_detail.to_excel(writer, sheet_name="ChiTiet_Giuong", index=False)
        if df_unassigned is not None and not df_unassigned.empty:
            df_unassigned.to_excel(writer, sheet_name="Chua_Gan_Giuong", index=False)
        if df_missing_kq is not None and not df_missing_kq.empty:
            df_missing_kq.to_excel(writer, sheet_name="CoNgayRa_TrongKQ", index=False)
        if df_gaps is not None and not df_gaps.empty:
            df_gaps.to_excel(writer, sheet_name="Ho_Khung_Gio", index=False)
        if df_sl_warn is not None and not df_sl_warn.empty:
            df_sl_warn.to_excel(writer, sheet_name="Nguoc_SL_Giuong", index=False)
    buf.seek(0)
    return buf.getvalue()


# ======================= POPUP HIỆN ĐẠI XEM CHI TIẾT GIƯỜNG =======================
@st.dialog("Chi tiết giường bệnh", width="large")
def _show_bed_modal(dept, bed, day_str=""):
    phong_str = f" — {bed['phong']}" if bed.get("phong") else ""
    name_str = f" ({bed['name']})" if bed.get("name") else ""
    sub_parts = [f"Ngày {day_str}" if day_str else "", f"Hiện có {bed['n_patients']} bệnh nhân nằm", dept["ten_khoa"]]
    sub_parts = [p for p in sub_parts if p]

    st.markdown(
        f"""
        <div style="margin-top:-10px;margin-bottom:12px;padding-bottom:10px;border-bottom:1px solid #E2E8F0;">
            <div style="font-size:17px;font-weight:700;color:#0F172A;">
                {html.escape(bed['code'])}{html.escape(name_str)}{html.escape(phong_str)}
            </div>
            <div style="font-size:12.5px;color:#64748B;margin-top:2px;">
                {html.escape(' • '.join(sub_parts))}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if bed.get("has_mixed_gender"):
        st.error("🚨 **CẢNH BÁO:** Giường này đang có bệnh nhân **NAM và NỮ nằm chung** cùng thời điểm!")
    if bed.get("has_multi_bed_conflict"):
        st.error("🚨 **CẢNH BÁO:** Có bệnh nhân tại giường này đồng thời **nằm trùng giờ ở giường khác**!")
    if bed.get("has_self_overlap"):
        st.warning("⚠️ Có bệnh nhân bị **tự trùng khung giờ** giữa các lượt nằm trên cùng giường này.")
    elif bed.get("has_overlap") and not bed.get("has_mixed_gender") and not bed.get("has_multi_bed_conflict"):
        st.warning("⚠️ Giường này có bệnh nhân trùng khung giờ nằm ghép.")

    for idx, p in enumerate(bed["patients"], start=1):
        age_str = _calc_age_str(p.get("NAM_SINH", ""))
        yr_str = _extract_birth_year(p.get("NAM_SINH", ""))
        g_parts = [x for x in [p.get("GIOI_TINH", ""), f"NS {yr_str}" if yr_str else "", age_str] if x]
        gt_tuoi = ", ".join(g_parts) if g_parts else "Chưa rõ"
        ma_benh_txt = f" • Mã bệnh: <b>{html.escape(p['MA_BENH'])}</b>" if p.get("MA_BENH") else ""
        st_icon = "✅ " if p.get("TRANG_THAI_BN") == "Ra viện" else ""

        other_beds = p.get("other_bed_conflicts", [])
        other_bed_html = ""
        if other_beds:
            other_bed_html = (
                f"<div style='margin-bottom:6px;padding:4px 8px;border-radius:6px;background:#FEF2F2;"
                f"border:1px solid #FCA5A5;color:#B91C1C;font-size:12px;font-weight:700;'>"
                f"🚨 Đang nằm trùng giờ với giường: {html.escape(', '.join(other_beds))}</div>"
            )

        slots_html_parts = []
        for s_idx, sl in enumerate(p.get("slots", []), start=1):
            d_s = sl["DISP_START"].strftime("%d/%m/%Y %H:%M")
            d_e = sl["DISP_END"].strftime("%d/%m/%Y %H:%M")
            if sl.get("IS_ESTIMATED_END"):
                time_disp = f"<span style='color:#7C3AED;font-weight:700;'>{d_s} → {d_e} (tạm tính)</span>"
            elif sl.get("MISSING_KQ_HAS_RA"):
                time_disp = f"<b style='color:#DC2626;'>{d_s} → {d_e} (theo NGAY_RA, trống KQ)</b>"
            else:
                time_disp = f"<b style='color:#0F172A;'>{d_s} → {d_e}</b>"

            sl_badge = f" • SL: <b style='color:#1D4ED8;'>{html.escape(sl['SO_LUONG_STR'])}</b>" if sl.get("SO_LUONG_STR") else ""
            dv_txt = sl.get("TEN_DICH_VU") or "Giường nội trú"
            prefix = f"Lượt {s_idx}: " if len(p.get("slots", [])) > 1 else "Thời gian nằm: "
            slots_html_parts.append(
                f"<div style='background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;"
                f"padding:7px 10px;font-size:12.5px;color:#334155;margin-top:6px;'>"
                f"<div>{prefix}{time_disp}{sl_badge}</div>"
                f"<div style='margin-top:2px;'>Dịch vụ: <b style='color:#0F172A;'>{html.escape(dv_txt)}</b> "
                f"• Tỷ lệ TT: <b style='color:#15803D;'>{int(sl['TYLE_TT'])}%</b></div>"
                f"</div>"
            )

        st.markdown(
            f"""
            <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:12px;
                        padding:14px 16px;margin-bottom:10px;box-shadow:0 1px 2px rgba(15,23,42,.04);">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:6px;">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <span style="display:inline-flex;align-items:center;justify-content:center;
                                     width:24px;height:24px;border-radius:50%;background:#1D4ED8;
                                     color:#FFFFFF;font-size:12px;font-weight:700;">{idx}</span>
                        <span style="font-size:14.5px;font-weight:700;color:#0F172A;">
                            {st_icon}{html.escape(p['TENBENHNHAN'])}
                        </span>
                    </div>
                    <span style="background:#EFF6FF;color:#1D4ED8;font-size:11.5px;font-weight:700;
                                 padding:3px 10px;border-radius:6px;border:1px solid #DBEAFE;">
                        {html.escape(gt_tuoi)}
                    </span>
                </div>
                <div style="font-size:12px;color:#475569;margin-bottom:4px;">
                    Trạng thái: <b>{st_icon}{html.escape(p.get('TRANG_THAI_BN', ''))}</b>{ma_benh_txt}
                    • LK: <span style="color:#64748B;">{html.escape(p['MA_LK'])}</span>
                </div>
                {other_bed_html}
                {''.join(slots_html_parts)}
            </div>
            """,
            unsafe_allow_html=True,
        )

    c_sp, c_close = st.columns([5, 1])
    with c_close:
        if st.button("Đóng", use_container_width=True, key=f"close_modal_{day_str}_{dept['ma_khoa']}_{bed['code']}"):
            st.rerun()


# ======================= RENDER SƠ ĐỒ HTML/CSS =======================
def _inject_bed_css():
    st.markdown(
        """
        <style>
        .bed-dept-header-bar {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 9px 14px; margin: 10px 0 8px 0;
            display: flex; justify-content: space-between; align-items: center;
            flex-wrap: wrap; gap: 8px;
            box-shadow: 0 1px 2px rgba(15,23,42,.04);
        }
        .bed-dept-title { font-size: 14px; font-weight: 700; color: #0F172A; }
        .bed-dept-meta { font-size: 12px; color: #475569; display: flex; gap: 10px; flex-wrap: wrap; }
        .bed-card {
            border-radius: 8px; padding: 8px 10px; border: 1px solid;
            font-size: 12px; line-height: 1.35; min-height: 92px;
        }
        .bed-card-occ {
            border-bottom-left-radius: 0; border-bottom-right-radius: 0; border-bottom: none;
        }
        .bed-empty   { background: #F0FDF4; border-color: #86EFAC; color: #166534; }
        .bed-single  { background: #EFF6FF; border-color: #93C5FD; color: #1E40AF; }
        .bed-shared  { background: #FFF7ED; border-color: #FDBA74; color: #9A3412; }
        .bed-overlap { background: #FEF2F2; border-color: #FCA5A5; color: #991B1B; }
        .bed-mixed   { background: #FEF2F2; border: 2px solid #DC2626; color: #991B1B; }
        .bed-top {
            display: flex; justify-content: space-between; align-items: center;
            font-weight: 700; margin-bottom: 3px;
        }
        .bed-badge {
            font-size: 10px; padding: 1px 6px; border-radius: 999px; font-weight: 600;
        }
        .badge-empty   { background: #DCFCE7; color: #15803D; }
        .badge-single  { background: #DBEAFE; color: #1D4ED8; }
        .badge-shared  { background: #FFEDD5; color: #C2410C; }
        .badge-overlap { background: #FEE2E2; color: #B91C1C; }
        .badge-mixed   { background: #DC2626; color: #FFFFFF; }
        .bed-bn {
            font-size: 11.5px; font-weight: 600; color: #0F172A; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis;
        }
        .bed-sub {
            font-size: 10.5px; color: #64748B; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis;
        }
        .bed-time {
            font-size: 10px; color: #475569; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis; margin-bottom: 2px;
        }
        .bed-ready-tag {
            margin-top: 4px; padding: 2px 6px; border-radius: 5px;
            background: #DCFCE7; color: #15803D; font-size: 11.5px;
            font-weight: 800; display: inline-block;
        }
        .bed-multibed-tag {
            margin-top: 2px; margin-bottom: 2px; padding: 1px 5px; border-radius: 4px;
            background: #FEE2E2; color: #B91C1C; font-size: 10px;
            font-weight: 700; display: inline-block;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_patient_slots_html(p, compact=False):
    """Render tên BN + toàn bộ các khung giờ nằm trong ngày + tag trùng giường khác (nếu có)."""
    st_icon = "✅ " if p.get("TRANG_THAI_BN") == "Ra viện" else ""
    yr = _extract_birth_year(p.get("NAM_SINH", ""))
    gt = p.get("GIOI_TINH", "")

    lines = []
    if not compact:
        sub_parts = [x for x in [f"NS: {yr}" if yr else "", gt] if x]
        sub_txt = " • ".join(sub_parts) if sub_parts else p.get("TRANG_THAI_BN", "")
        lines.append(
            f"<div class='bed-bn' title='{html.escape(p['TENBENHNHAN'])}'>{st_icon}{html.escape(p['TENBENHNHAN'])}</div>"
        )
        lines.append(f"<div class='bed-sub'>{html.escape(sub_txt)}</div>")
    else:
        tag_parts = [x for x in [yr, gt] if x]
        tag_str = f" [{' • '.join(tag_parts)}]" if tag_parts else ""
        lines.append(
            f"<div class='bed-bn' title='{html.escape(p['TENBENHNHAN'])}'>• {st_icon}{html.escape(p['TENBENHNHAN'])}{html.escape(tag_str)}</div>"
        )

    for sl in p.get("slots", []):
        t_html = _format_card_time_html(
            sl["DISP_START"],
            sl["DISP_END"],
            is_est_end=sl.get("IS_ESTIMATED_END", False),
            sl_str=sl.get("SO_LUONG_STR", ""),
        )
        lines.append(f"<div class='bed-time'>{t_html}</div>")

    other_beds = p.get("other_bed_conflicts", [])
    if other_beds:
        lines.append(
            f"<div class='bed-multibed-tag'>⚠️ Trùng giường {html.escape(', '.join(other_beds))}</div>"
        )

    return "".join(lines)


def _render_dept_grid(dept, status_filter, search_kw, key_prefix="d", day_str=""):
    kw = search_kw.strip().lower()
    visible_beds = []

    for b in dept["beds"]:
        if status_filter == "Giường trống" and b["status"] != "empty":
            continue
        if status_filter == "Đang nằm riêng" and b["status"] != "single":
            continue
        if status_filter == "Trùng / Nằm ghép" and b["status"] != "shared":
            continue
        if status_filter == "🚨 Nam - Nữ nằm chung" and not b.get("has_mixed_gender"):
            continue

        if kw:
            haystack = (
                f"{b['code']} {b['name']} {b['loai']} {b['phong']} "
                + " ".join(
                    f"{p['TENBENHNHAN']} {p['MA_LK']} {p.get('GIOI_TINH', '')} {p.get('NAM_SINH', '')}"
                    for p in b["patients"]
                )
            ).lower()
            if kw not in haystack:
                continue
        visible_beds.append(b)

    if not visible_beds:
        return

    bn_thucte = dept.get("actual_patients", 0)
    tot_pd = dept.get("total_approved", dept["total"])
    pct = (bn_thucte / tot_pd * 100) if tot_pd else 0.0
    st.markdown(
        f"""
        <div class="bed-dept-header-bar">
            <div class="bed-dept-title">🏥 {html.escape(dept['ten_khoa'])} ({html.escape(dept['ma_khoa'])})</div>
            <div class="bed-dept-meta">
                <span>BN thực tế / Giường PD: <b>{bn_thucte}/{tot_pd}</b></span>
                <span style="color:#15803D;">Trống: <b>{dept['empty']}</b></span>
                <span style="color:#1D4ED8;">Nằm riêng: <b>{dept['single']}</b></span>
                <span style="color:#B91C1C;">Trùng / Ghép: <b>{dept['shared']}</b></span>
                <span>Công suất: <b>{pct:.1f}%</b></span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    n_cols = 6
    for i in range(0, len(visible_beds), n_cols):
        row_beds = visible_beds[i : i + n_cols]
        cols = st.columns(n_cols, gap="small")
        for col, b in zip(cols, row_beds):
            with col:
                meta_parts = [p for p in [b.get("phong"), b.get("loai")] if p]
                meta_line = f"<div class='bed-sub'>{html.escape(' • '.join(meta_parts))}</div>" if meta_parts else ""

                if b["status"] == "empty":
                    cls, bcls, lbl = "bed-empty", "badge-empty", "Trống"
                    ready_html = (
                        "<div style='margin-top:8px;font-size:13px;font-weight:800;color:#15803D;'>"
                        "Sẵn sàng kê</div>"
                    )
                    body = meta_line + ready_html
                    st.markdown(
                        f"<div class='bed-card {cls}'>"
                        f"  <div class='bed-top'><span>{html.escape(b['code'])}</span>"
                        f"  <span class='bed-badge {bcls}'>{lbl}</span></div>"
                        f"  {body}"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    if b["status"] == "single":
                        cls, bcls, lbl = "bed-single", "badge-single", f"{b['n_patients']} BN"
                        if b["n_patients"] == 1:
                            body = meta_line + _render_patient_slots_html(b["patients"][0], compact=False)
                        else:
                            body = meta_line + "".join(
                                _render_patient_slots_html(p, compact=True) for p in b["patients"][:3]
                            )
                    else:
                        if b.get("has_mixed_gender"):
                            cls, bcls, lbl = "bed-mixed", "badge-mixed", "🚨 NAM - NỮ"
                        elif b.get("has_multi_bed_conflict"):
                            cls, bcls, lbl = "bed-mixed", "badge-mixed", "🚨 2 GIƯỜNG"
                        elif b.get("has_self_overlap"):
                            cls, bcls, lbl = "bed-overlap", "badge-overlap", "TỰ TRÙNG"
                        elif b.get("has_overlap"):
                            cls, bcls, lbl = "bed-overlap", "badge-overlap", f"TRÙNG ({b['n_patients']})"
                        else:
                            cls, bcls, lbl = "bed-shared", "badge-shared", f"Ghép ({b['n_patients']})"

                        compact_flag = (b["n_patients"] > 1)
                        body = meta_line + "".join(
                            _render_patient_slots_html(p, compact=compact_flag) for p in b["patients"][:3]
                        )

                    # Nếu BN cuối cùng ra trước 23:59 -> hiện giờ sẵn sàng (giờ ra + 1 phút)
                    ready_dt = b.get("ready_after_dt")
                    if ready_dt is not None:
                        body += f"<div class='bed-ready-tag'>🟢 Sẵn sàng từ: {ready_dt.strftime('%H:%M')}</div>"

                    st.markdown(
                        f"<div class='bed-card bed-card-occ {cls}'>"
                        f"  <div class='bed-top'><span>{html.escape(b['code'])}</span>"
                        f"  <span class='bed-badge {bcls}'>{lbl}</span></div>"
                        f"  {body}"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        f"👁️ Xem ({b['n_patients']} BN)",
                        key=f"btn_{key_prefix}_{dept['ma_khoa']}_{b['code']}",
                        use_container_width=True,
                    ):
                        _show_bed_modal(dept, b, day_str=day_str)


# ======================= TRANG CHÍNH: SƠ ĐỒ GIƯỜNG =======================
def page_bed_management(ref_path, admin=True, render_kpi_fn=None, show_uploader=True):
    _inject_bed_css()
    st.markdown(
        "<div class='section-title'>🛏️ Sơ đồ & Kiểm tra trùng Giường bệnh theo từng ngày (MA_NHOM = 15 • MA_GIUONG_CV130)</div>",
        unsafe_allow_html=True,
    )

    bed_catalog = load_bed_catalog(ref_path)
    if not bed_catalog:
        st.warning("⚠️ Chưa đọc được danh mục giường (`dm_giuong_chitiet` hoặc `data_giuong`) từ file tham chiếu ở Sidebar.")
        return

    src_cat = bed_catalog[0].get("SOURCE", "data_giuong")
    st.caption(f"📚 Nguồn danh mục giường đang dùng: **{src_cat}**")

    saved = load_bed_dataset()
    df_bed_all = saved["df_bed"] if saved else pd.DataFrame()
    src_name = saved.get("src_filename", "") if saved else ""
    updated_at = saved.get("updated_at", "") if saved else ""

    if not df_bed_all.empty:
        if "MISSING_KQ_HAS_RA" not in df_bed_all.columns:
            df_bed_all["MISSING_KQ_HAS_RA"] = False
        if "IS_ESTIMATED_END" not in df_bed_all.columns:
            df_bed_all["IS_ESTIMATED_END"] = False
        if "SO_LUONG_STR" not in df_bed_all.columns:
            df_bed_all["SO_LUONG_STR"] = ""
        if "SO_LUONG_NUM" not in df_bed_all.columns:
            df_bed_all["SO_LUONG_NUM"] = None

    if admin and show_uploader:
        exp_title = (
            f"📤 Nạp báo cáo giường mới (đang dùng: {src_name} • cập nhật {updated_at})"
            if src_name else
            "📤 Nạp file báo cáo có dữ liệu giường (MA_NHOM = 15 & MA_GIUONG_CV130)"
        )
        with st.expander(exp_title, expanded=df_bed_all.empty):
            c_up, c_btn = st.columns([3.5, 1], vertical_alignment="center")
            with c_up:
                up_file = st.file_uploader(
                    "Chọn file Excel (.xlsx, .xls) hoặc XML (.xml) có MA_NHOM = 15",
                    type=["xlsx", "xls", "xml"],
                    key="bed_file_uploader",
                    label_visibility="collapsed",
                )
            with c_btn:
                load_btn = st.button("📥 Nạp dữ liệu giường", type="primary", use_container_width=True)

            if load_btn:
                if up_file is None:
                    st.warning("Vui lòng chọn file trước khi nạp.")
                else:
                    try:
                        with st.spinner("Đang lọc dữ liệu MA_NHOM = 15..."):
                            df_raw = read_input_dataframe(up_file.getvalue(), up_file.name)
                            df_bed_new = extract_bed_records(df_raw)
                        if df_bed_new.empty:
                            st.error("❌ File vừa tải lên có **0 dòng `MA_NHOM = 15`**.")
                        else:
                            save_bed_dataset(df_bed_new, up_file.name)
                            st.session_state.pop("bed_applied_filter", None)
                            st.toast(f"✅ Đã nạp {len(df_bed_new):,} dòng giường từ {up_file.name}", icon="🛏️")
                            st.rerun()
                    except Exception as ex:
                        st.error(f"❌ Lỗi đọc file: {ex}")

    today = date.today()
    if not df_bed_all.empty:
        min_d = df_bed_all["START_DT"].min().date()
        max_d = df_bed_all["END_DT"].max().date()
    else:
        min_d, max_d = today, today

    khoa_labels = [f"{d['TEN_KHOA']} ({d['MA_KHOA']})" for d in bed_catalog]

    if "bed_applied_filter" not in st.session_state:
        st.session_state["bed_applied_filter"] = {
            "tu_ngay": min_d,
            "tu_gio": dt_time(0, 0),
            "den_ngay": max_d,
            "den_gio": dt_time(23, 59),
            "sel_khoa": [],
            "status_filter": "Tất cả",
            "search_kw": "",
        }

    flt = st.session_state["bed_applied_filter"]

    # ----- Khung lọc có nút Tìm kiếm -----
    with st.form("bed_search_form", clear_on_submit=False):
        t1, t2, t3, t4 = st.columns([1.3, 1.0, 1.3, 1.0], gap="small")
        with t1:
            in_tu_ngay = st.date_input("Từ ngày", value=flt["tu_ngay"], format="DD/MM/YYYY")
        with t2:
            in_tu_gio = st.time_input("Từ giờ", value=flt["tu_gio"], step=300)
        with t3:
            in_den_ngay = st.date_input("Đến ngày", value=flt["den_ngay"], format="DD/MM/YYYY")
        with t4:
            in_den_gio = st.time_input("Đến giờ", value=flt["den_gio"], step=300)

        f1, f2, f3, f4 = st.columns([1.9, 1.3, 1.5, 0.9], gap="small", vertical_alignment="bottom")
        with f1:
            in_khoa = st.multiselect(
                "Khoa nội trú",
                options=khoa_labels,
                default=[k for k in flt["sel_khoa"] if k in khoa_labels],
                placeholder="Tất cả khoa...",
            )
        with f2:
            status_opts = ["Tất cả", "Giường trống", "Đang nằm riêng", "Trùng / Nằm ghép", "🚨 Nam - Nữ nằm chung"]
            in_status = st.selectbox(
                "Trạng thái giường",
                options=status_opts,
                index=status_opts.index(flt["status_filter"]) if flt["status_filter"] in status_opts else 0,
            )
        with f3:
            in_kw = st.text_input("Tìm nhanh", value=flt["search_kw"], placeholder="Mã giường H001, Phòng, Tên BN...")
        with f4:
            submitted = st.form_submit_button("🔍 Tìm kiếm", type="primary", use_container_width=True)

        if submitted:
            st.session_state["bed_applied_filter"] = {
                "tu_ngay": in_tu_ngay,
                "tu_gio": in_tu_gio,
                "den_ngay": in_den_ngay,
                "den_gio": in_den_gio,
                "sel_khoa": in_khoa,
                "status_filter": in_status,
                "search_kw": in_kw,
            }
            flt = st.session_state["bed_applied_filter"]

    tu_ngay, tu_gio = flt["tu_ngay"], flt["tu_gio"]
    den_ngay, den_gio = flt["den_ngay"], flt["den_gio"]
    sel_khoa = flt["sel_khoa"]
    status_filter = flt["status_filter"]
    search_kw = flt["search_kw"]

    range_start = datetime.combine(tu_ngay, tu_gio)
    range_end = datetime.combine(den_ngay, den_gio)
    if range_start > range_end:
        st.error("⚠️ Thời điểm 'Từ ngày/giờ' phải nhỏ hơn hoặc bằng 'Đến ngày/giờ'.")
        return

    total_days = (den_ngay - tu_ngay).days + 1
    if total_days > 62:
        st.warning(f"⚠️ Khoảng thời gian đang chọn dài **{total_days} ngày** (tối đa 62 ngày/lần). Vui lòng thu hẹp khoảng ngày.")
        return

    if not df_bed_all.empty:
        mask_time = (df_bed_all["START_DT"] <= range_end) & (df_bed_all["END_DT"] >= range_start)
        df_filtered = df_bed_all[mask_time].copy()
    else:
        df_filtered = pd.DataFrame(columns=[
            "MA_KHOA", "MA_LK", "TENBENHNHAN", "GIOI_TINH_STD",
            "MA_GIUONG_STD", "START_DT", "END_DT", "MISSING_KQ_HAS_RA",
            "IS_ESTIMATED_END", "SO_LUONG_STR", "SO_LUONG_NUM",
        ])

    if sel_khoa:
        chosen_codes = [s.split("(")[-1].rstrip(")") for s in sel_khoa]
        cat_view = [d for d in bed_catalog if d["MA_KHOA"] in chosen_codes]
    else:
        cat_view = bed_catalog

    # ----- TÍNH ĐỘC LẬP TỪNG NGÀY MỘT -----
    daily_packs = []
    all_ov_frames = []
    all_detail_frames = []
    all_unassigned_frames = []
    all_missing_kq_frames = []
    all_gap_frames = []
    all_sl_warn_frames = []
    all_summary_rows = []

    for offset in range(total_days):
        cur_d = tu_ngay + timedelta(days=offset)
        date_str = cur_d.strftime("%d/%m/%Y")
        d_start = max(range_start, datetime.combine(cur_d, dt_time(0, 0, 0)))
        d_end = min(range_end, datetime.combine(cur_d, dt_time(23, 59, 59)))

        if not df_filtered.empty:
            df_day = df_filtered[(df_filtered["START_DT"] <= d_end) & (df_filtered["END_DT"] >= d_start)]
        else:
            df_day = df_filtered

        (
            day_depts,
            day_detail,
            day_ov,
            day_unassigned,
            day_missing_kq,
            day_gaps,
            day_sl_warn,
        ) = build_bed_map_for_day(cat_view, df_day, d_start, d_end, day_str=date_str, df_all=df_bed_all)

        day_n_ov = len(day_ov)
        if not day_ov.empty:
            day_n_mixed = int(day_ov["Cảnh báo"].str.contains("NAM - NỮ", na=False).sum())
            day_n_multi_bed = int(day_ov["Cảnh báo"].str.contains("2 GIƯỜNG", na=False).sum())
            day_n_self = int(day_ov["Cảnh báo"].str.contains("Tự trùng", na=False).sum())
            day_n_time_only = day_n_ov - day_n_mixed - day_n_multi_bed - day_n_self
        else:
            day_n_mixed = day_n_multi_bed = day_n_self = day_n_time_only = 0

        daily_packs.append({
            "cur_d": cur_d,
            "date_str": date_str,
            "depts": day_depts,
            "detail": day_detail,
            "ov": day_ov,
            "unassigned": day_unassigned,
            "missing_kq": day_missing_kq,
            "gaps": day_gaps,
            "sl_warn": day_sl_warn,
            "n_ov": day_n_ov,
            "n_mixed": day_n_mixed,
            "n_multi_bed": day_n_multi_bed,
            "n_self": day_n_self,
            "n_time_only": day_n_time_only,
        })

        if not day_ov.empty:
            all_ov_frames.append(day_ov)
        if not day_detail.empty:
            all_detail_frames.append(day_detail)
        if not day_unassigned.empty:
            all_unassigned_frames.append(day_unassigned)
        if not day_missing_kq.empty:
            all_missing_kq_frames.append(day_missing_kq)
        if not day_gaps.empty:
            all_gap_frames.append(day_gaps)
        if not day_sl_warn.empty:
            all_sl_warn_frames.append(day_sl_warn)
        for d in day_depts:
            tot_pd = d["total_approved"]
            bn_tt = d["actual_patients"]
            all_summary_rows.append({
                "Ngày": date_str,
                "Mã khoa": d["ma_khoa"],
                "Tên khoa": d["ten_khoa"],
                "BN thực tế": bn_tt,
                "Giường phê duyệt": tot_pd,
                "Giường trống": d["empty"],
                "Nằm riêng": d["single"],
                "Trùng / Ghép": d["shared"],
                "Công suất (%)": round(bn_tt / tot_pd * 100, 1) if tot_pd else 0,
            })

    df_all_ov = pd.concat(all_ov_frames, ignore_index=True) if all_ov_frames else pd.DataFrame()
    df_all_detail = pd.concat(all_detail_frames, ignore_index=True) if all_detail_frames else pd.DataFrame()
    df_all_unassigned = pd.concat(all_unassigned_frames, ignore_index=True) if all_unassigned_frames else pd.DataFrame()
    df_all_missing_kq = pd.concat(all_missing_kq_frames, ignore_index=True) if all_missing_kq_frames else pd.DataFrame()
    df_all_gaps = pd.concat(all_gap_frames, ignore_index=True) if all_gap_frames else pd.DataFrame()
    df_all_sl_warn = pd.concat(all_sl_warn_frames, ignore_index=True) if all_sl_warn_frames else pd.DataFrame()
    df_all_summary = pd.DataFrame(all_summary_rows)

    total_beds_catalog = sum(len(d["BEDS"]) for d in cat_view)
    days_with_ov = sum(1 for p in daily_packs if p["n_ov"] > 0)
    total_ov_pairs = len(df_all_ov)
    total_mixed_pairs = sum(p["n_mixed"] for p in daily_packs)
    total_multi_bed_pairs = sum(p["n_multi_bed"] for p in daily_packs)

    if total_mixed_pairs > 0 or total_multi_bed_pairs > 0:
        warn_parts = []
        if total_mixed_pairs > 0:
            warn_parts.append(f"**{total_mixed_pairs}** cặp **NAM – NỮ nằm chung giường**")
        if total_multi_bed_pairs > 0:
            warn_parts.append(f"**{total_multi_bed_pairs}** lượt **1 BN nằm 2 giường trùng giờ**")
        st.error(
            f"🚨 **CẢNH BÁO:** Phát hiện {' và '.join(warn_parts)}! "
            f"Mở từng ngày có cảnh báo đỏ bên dưới hoặc xem tab **⚠️ Tổng hợp trùng giường**."
        )

    kpi_items = [
        ("Số ngày xem", f"{total_days:,}", "#1565C0"),
        ("Tổng giường PD", f"{total_beds_catalog:,}", "#0288D1"),
        ("Ngày có trùng", f"{days_with_ov:,}", "#EA580C" if days_with_ov else "#15803D"),
        ("Tổng lượt trùng", f"{total_ov_pairs:,}", "#C62828" if total_ov_pairs else "#64748B"),
        ("Nam - Nữ chung", f"{total_mixed_pairs:,}", "#DC2626" if total_mixed_pairs else "#64748B"),
    ]
    if render_kpi_fn:
        render_kpi_fn(kpi_items)
    else:
        cols = st.columns(len(kpi_items))
        for col, (lbl, val, _) in zip(cols, kpi_items):
            col.metric(lbl, val)

    tab_daily, tab_ov, tab_tbl = st.tabs([
        "📅 Sơ đồ theo từng ngày (Mở chi tiết)",
        f"⚠️ Tổng hợp trùng giường ({total_ov_pairs})",
        "📋 Danh sách chi tiết theo ngày",
    ])

    # ----- TAB 1: XEM TỪNG NGÀY MỘT -----
    with tab_daily:
        only_conflict_days = st.checkbox(
            "Chỉ hiện những ngày có nằm trùng giường",
            value=False,
            key="bed_only_conflict_days",
        )

        shown_any = False
        for pack in daily_packs:
            if only_conflict_days and pack["n_ov"] == 0:
                continue

            shown_any = True
            date_str = pack["date_str"]
            cur_d = pack["cur_d"]
            day_depts = pack["depts"]
            day_ov = pack["ov"]
            day_detail = pack["detail"]
            day_unassigned = pack["unassigned"]
            day_missing_kq = pack["missing_kq"]
            day_gaps = pack["gaps"]
            day_sl_warn = pack["sl_warn"]

            day_actual_bn = sum(d["actual_patients"] for d in day_depts)
            day_total_pd = sum(d["total_approved"] for d in day_depts)
            day_empty_beds = sum(d["empty"] for d in day_depts)
            day_pct = (day_actual_bn / day_total_pd * 100) if day_total_pd else 0.0

            dept_chips = []
            has_weird_overlap_day = False

            for d in day_depts:
                short_name = _short_dept_name(d["ma_khoa"], d["ten_khoa"])
                d_bn = d["actual_patients"]
                d_tot = d["total_approved"]
                d_emp = d["empty"]

                if not day_ov.empty:
                    ov_sub = day_ov[day_ov["Mã khoa"].astype(str).str.contains(d["ma_khoa"], na=False)]
                    d_ov = len(ov_sub)
                    d_mixed = int(ov_sub["Cảnh báo"].str.contains("NAM - NỮ", na=False).sum())
                    d_mb = int(ov_sub["Cảnh báo"].str.contains("2 GIƯỜNG", na=False).sum())
                    d_self = int(ov_sub["Cảnh báo"].str.contains("Tự trùng", na=False).sum())
                    d_time = d_ov - d_mixed - d_mb - d_self
                else:
                    d_ov = d_mixed = d_mb = d_self = d_time = 0

                ov_desc = _format_overlap_desc(d_mixed, d_mb, d_self, d_time)
                has_conflict_or_shared = (d_ov > 0) or (d["shared"] > 0)

                if has_conflict_or_shared and d_emp > 0:
                    has_weird_overlap_day = True
                    reason_txt = ov_desc if ov_desc else f"{d['shared']} ghép"
                    dept_chips.append(
                        f":red-background[:red[**🔴 {short_name}: {d_bn}/{d_tot} (Trống {d_emp} • {reason_txt})**]]"
                    )
                elif d_ov > 0:
                    dept_chips.append(f":red[**{short_name}: {d_bn}/{d_tot} ({ov_desc})**]")
                elif d["shared"] > 0:
                    dept_chips.append(f":orange[**{short_name}: {d_bn}/{d_tot} (Trống {d_emp})**]")
                elif d_bn > 0:
                    dept_chips.append(f":blue[**{short_name}: {d_bn}/{d_tot} (Trống {d_emp})**]")
                else:
                    dept_chips.append(f":gray[{short_name}: 0/{d_tot}]")

            dept_summary_str = "  •  ".join(dept_chips)

            day_ov_desc = _format_overlap_desc(
                pack["n_mixed"], pack["n_multi_bed"], pack["n_self"], pack["n_time_only"]
            )
            if has_weird_overlap_day:
                lbl_detail = day_ov_desc if day_ov_desc else "nằm ghép"
                status_tag = f":red-background[:red[**🚨 CÒN {day_empty_beds} GIƯỜNG TRỐNG VẪN TRÙNG ({lbl_detail})**]]"
            elif pack["n_ov"] > 0:
                if pack["n_mixed"] > 0 or pack["n_multi_bed"] > 0:
                    status_tag = f"🚨 :red[**{day_ov_desc}**]"
                else:
                    status_tag = f"⚠️ :orange[**{day_ov_desc}**]"
            else:
                status_tag = "✅ :green[Không trùng]"

            extra_tags = []
            if not day_unassigned.empty:
                extra_tags.append(f":orange[**{len(day_unassigned)} chưa gán giường**]")
            if not day_missing_kq.empty:
                extra_tags.append(f":red[**{len(day_missing_kq)} có ngày ra trống KQ**]")
            if not day_gaps.empty:
                extra_tags.append(f":orange[**{len(day_gaps)} hở giờ ≥{GAP_WARNING_MINUTES}p**]")
            if not day_sl_warn.empty:
                extra_tags.append(f":red[**{len(day_sl_warn)} ngược SL**]")
            extra_str = ("  |  " + " • ".join(extra_tags)) if extra_tags else ""

            exp_lbl = (
                f"**{date_str}**  |  {status_tag}{extra_str}  |  "
                f"Tổng BN: **{day_actual_bn}/{day_total_pd}** ({day_pct:.0f}%) • Trống: **{day_empty_beds}**  —  {dept_summary_str}"
            )

            with st.expander(exp_lbl, expanded=(total_days == 1)):
                if has_weird_overlap_day:
                    st.error(
                        f"🔴 **CẢNH BÁO BẤT THƯỜNG:** Ngày **{date_str}** vẫn còn **{day_empty_beds} giường trống** "
                        f"nhưng vẫn phát sinh nằm ghép / trùng giường!"
                    )

                if not day_ov.empty:
                    st.markdown("**⚠️ Danh sách trùng giường / trùng giờ trong ngày:**")
                    st.dataframe(day_ov, use_container_width=True, hide_index=True)

                if not day_sl_warn.empty:
                    st.error(f"⚠️ **Cảnh báo ngược Số lượng giường ({len(day_sl_warn)} BN):**")
                    st.dataframe(day_sl_warn, use_container_width=True, hide_index=True)

                if not day_gaps.empty:
                    st.warning(f"⚠️ **Cảnh báo hở khung giờ nằm giường ≥ {GAP_WARNING_MINUTES} phút ({len(day_gaps)} lượt):**")
                    st.dataframe(day_gaps, use_container_width=True, hide_index=True)

                if not day_missing_kq.empty:
                    st.warning(f"⚠️ **Cảnh báo: {len(day_missing_kq)} bệnh nhân đã có NGÀY RA nhưng trống NGÀY KẾT QUẢ (`NGAY_KQ`):**")
                    st.dataframe(day_missing_kq, use_container_width=True, hide_index=True)

                if not day_unassigned.empty:
                    st.warning(f"⚠️ **Cảnh báo: {len(day_unassigned)} bệnh nhân chưa gán mã giường (`MA_GIUONG_CV130` trống):**")
                    st.dataframe(day_unassigned, use_container_width=True, hide_index=True)

                for dept in day_depts:
                    _render_dept_grid(
                        dept,
                        status_filter,
                        search_kw,
                        key_prefix=f"d_{cur_d.strftime('%Y%m%d')}",
                        day_str=date_str,
                    )

                if not day_detail.empty:
                    day_sum = pd.DataFrame([
                        {
                            "Ngày": date_str,
                            "Mã khoa": d["ma_khoa"],
                            "Tên khoa": d["ten_khoa"],
                            "BN thực tế": d["actual_patients"],
                            "Giường phê duyệt": d["total_approved"],
                            "Giường trống": d["empty"],
                            "Nằm riêng": d["single"],
                            "Trùng / Ghép": d["shared"],
                        }
                        for d in day_depts
                    ])
                    st.download_button(
                        f"⬇️ Tải Excel chi tiết ngày {date_str}",
                        data=_build_excel_bytes(
                            day_sum, day_ov, day_detail, day_unassigned, day_missing_kq, day_gaps, day_sl_warn
                        ),
                        file_name=f"Giuong_{cur_d.strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_day_{cur_d.strftime('%Y%m%d')}",
                    )

        if not shown_any:
            st.success("✅ Không có ngày nào bị trùng giường trong khoảng thời gian đã chọn.")

    # ----- TAB 2: TỔNG HỢP TRÙNG GIƯỜNG -----
    with tab_ov:
        if df_all_ov.empty:
            st.success(
                f"✅ Không phát hiện ca trùng giường nào trong khoảng "
                f"**{range_start.strftime('%d/%m/%Y %H:%M')} → {range_end.strftime('%d/%m/%Y %H:%M')}**."
            )
        else:
            st.error(
                f"⚠️ Phát hiện **{total_ov_pairs:,}** lượt trùng giường / trùng giờ "
                f"(Nam – Nữ: **{total_mixed_pairs}**, 1 BN nằm 2 giường: **{total_multi_bed_pairs}**):"
            )
            st.dataframe(df_all_ov, use_container_width=True, hide_index=True, height=460)

    # ----- TAB 3: DANH SÁCH CHI TIẾT & XUẤT EXCEL -----
    with tab_tbl:
        if df_all_detail.empty:
            st.info("Không có lượt bệnh nhân nằm giường trong khoảng thời gian đã chọn.")
        else:
            st.dataframe(df_all_detail, use_container_width=True, hide_index=True, height=460)
            st.download_button(
                "⬇️ Xuất Excel báo cáo giường theo ngày",
                data=_build_excel_bytes(
                    df_all_summary,
                    df_all_ov,
                    df_all_detail,
                    df_all_unassigned,
                    df_all_missing_kq,
                    df_all_gaps,
                    df_all_sl_warn,
                ),
                file_name=f"BaoCao_Giuong_{range_start.strftime('%Y%m%d')}_{range_end.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
