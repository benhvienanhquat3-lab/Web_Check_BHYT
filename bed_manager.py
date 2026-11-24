# -*- coding: utf-8 -*-
"""
bed_manager.py
=======================================================================
Module độc lập: Sơ đồ & Kiểm tra trùng Giường bệnh nội trú (MA_NHOM = 15)
- Tách độc lập hoàn toàn THEO TỪNG NGÀY MỘT (không gộp đè giữa các ngày).
- Đọc danh mục chi tiết từ sheet `dm_giuong_chitiet` (fallback `data_giuong`).
- Hiển thị Số BN thực tế / Số giường phê duyệt (vd: 22/20).
- Ghi rõ loại trùng (trùng Nam-Nữ / trùng TG) ngay trên thanh ngày.
- Cảnh báo đỏ rực nếu khoa có trùng giường nhưng Số BN < Số giường phê duyệt.
- Nhấn vào giường có người bật Popup hiện đại (@st.dialog) xem hồ sơ BN.
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


def _calc_age_str(nam_sinh):
    s = _clean_str(nam_sinh)
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 4:
        try:
            yr = int(digits[:4])
            age = datetime.now().year - yr
            if 0 <= age <= 120:
                return f"{age} tuổi"
        except Exception:
            pass
    return ""


def _short_dept_name(ma_khoa, ten_khoa):
    """Rút gọn tên khoa để hiển thị gọn đẹp trên 1 dòng tổng hợp."""
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


def _format_overlap_desc(n_mixed, n_time_only):
    parts = []
    if n_mixed > 0:
        parts.append(f"{n_mixed} trùng Nam-Nữ")
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

    for col in [
        "MA_LK", "TENBENHNHAN", "NAM_SINH", "GIOI_TINH", "MA_BENH",
        "MA_KHOA", "TEN_KHOA", "MA_DICH_VU", "TEN_DICH_VU", "TYLE_TT",
        "MA_GIUONG_CV130", "MA_GIUONG",
        "NGAY_YL", "NGAY_TH_YL", "NGAY_KQ", "NGAY_VAO", "NGAY_RA",
    ]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(_clean_str)

    df["GIOI_TINH_STD"] = df["GIOI_TINH"].map(_normalize_gender)
    df["MA_GIUONG_STD"] = df["MA_GIUONG_CV130"].where(
        df["MA_GIUONG_CV130"] != "", df["MA_GIUONG"]
    ).str.upper()

    df["TYLE_TT_NUM"] = pd.to_numeric(df["TYLE_TT"], errors="coerce").fillna(100)

    starts, ends = [], []
    now_dt = datetime.now()
    for _, r in df.iterrows():
        st_dt = (
            parse_bhyt_datetime(r["NGAY_TH_YL"], is_end=False)
            or parse_bhyt_datetime(r["NGAY_YL"], is_end=False)
            or parse_bhyt_datetime(r["NGAY_VAO"], is_end=False)
        )
        en_dt = (
            parse_bhyt_datetime(r["NGAY_KQ"], is_end=True)
            or parse_bhyt_datetime(r["NGAY_RA"], is_end=True)
            or now_dt
        )
        if st_dt and en_dt and en_dt < st_dt:
            en_dt = st_dt
        starts.append(st_dt)
        ends.append(en_dt)

    df["START_DT"] = starts
    df["END_DT"] = ends
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


def build_bed_map_for_day(bed_catalog, df_day, day_start, day_end, day_str=""):
    """Tính sơ đồ giường & trùng giờ độc lập cho ĐÚNG 1 NGÀY."""
    dept_results = []
    detail_rows = []
    overlap_rows = []

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
                rec = {
                    "MA_LK": row["MA_LK"],
                    "TENBENHNHAN": row["TENBENHNHAN"] or f"BN {row['MA_LK']}",
                    "NAM_SINH": row.get("NAM_SINH", ""),
                    "GIOI_TINH": row.get("GIOI_TINH_STD", ""),
                    "MA_BENH": row.get("MA_BENH", ""),
                    "MA_DICH_VU": row["MA_DICH_VU"],
                    "TEN_DICH_VU": row["TEN_DICH_VU"],
                    "TYLE_TT": row["TYLE_TT_NUM"],
                    "START_DT": row["START_DT"],
                    "END_DT": row["END_DT"],
                    "TRANG_THAI_BN": "Ra viện" if row["NGAY_RA"] else "Đang điều trị",
                }
                dept_unique_patients.add(rec["MA_LK"] or rec["TENBENHNHAN"])

                real_code = row.get("MA_GIUONG_STD", "").strip()
                if real_code:
                    if real_code not in beds:
                        beds[real_code] = []
                        bed_meta[real_code] = {"code": real_code, "name": "Ngoài danh mục", "loai": "", "phong": ""}
                    beds[real_code].append(rec)
                    assigned_code = real_code
                    source_tag = "MA_GIUONG_CV130"
                else:
                    assigned_code = None
                    for code in bed_codes:
                        if any(x["MA_LK"] == rec["MA_LK"] and rec["MA_LK"] for x in beds[code]):
                            assigned_code = code
                            break
                    if assigned_code is None:
                        for code in bed_codes:
                            has_ov = any(
                                _get_overlap_window(rec["START_DT"], rec["END_DT"], x["START_DT"], x["END_DT"], day_start, day_end)
                                for x in beds[code]
                            )
                            if not has_ov:
                                assigned_code = code
                                break
                    if assigned_code is None and bed_codes:
                        assigned_code = min(
                            bed_codes,
                            key=lambda c: sum(
                                1 for x in beds[c]
                                if _get_overlap_window(rec["START_DT"], rec["END_DT"], x["START_DT"], x["END_DT"], day_start, day_end)
                            ),
                        )
                    if assigned_code:
                        beds[assigned_code].append(rec)
                    source_tag = "Gán tạm"

                meta = bed_meta.get(assigned_code, {})
                detail_rows.append({
                    "Ngày": day_str,
                    "Mã khoa": ma_khoa,
                    "Tên khoa": ten_khoa,
                    "Phòng": meta.get("phong", ""),
                    "Mã giường": assigned_code or "",
                    "Nguồn mã": source_tag,
                    "Mã LK": rec["MA_LK"],
                    "Tên bệnh nhân": rec["TENBENHNHAN"],
                    "Giới tính": rec["GIOI_TINH"],
                    "Trạng thái BN": rec["TRANG_THAI_BN"],
                    "Tên dịch vụ giường": rec["TEN_DICH_VU"],
                    "Tỷ lệ TT (%)": int(rec["TYLE_TT"]),
                    "Từ thời điểm": rec["START_DT"].strftime("%d/%m/%Y %H:%M"),
                    "Đến thời điểm": rec["END_DT"].strftime("%d/%m/%Y %H:%M"),
                })

        bed_cards = []
        cnt_empty = cnt_single = cnt_shared = 0
        for code, occupants in beds.items():
            unique_patients = []
            seen = set()
            for o in occupants:
                k = o["MA_LK"] or o["TENBENHNHAN"]
                if k not in seen:
                    seen.add(k)
                    unique_patients.append(o)

            has_low_rate = any(o["TYLE_TT"] < 100 for o in occupants)
            has_time_overlap = False
            has_mixed_gender = False
            seen_pairs = set()

            for i in range(len(occupants)):
                for j in range(i + 1, len(occupants)):
                    p1, p2 = occupants[i], occupants[j]
                    k1 = p1["MA_LK"] or p1["TENBENHNHAN"]
                    k2 = p2["MA_LK"] or p2["TENBENHNHAN"]
                    if k1 == k2:
                        continue
                    ov = _get_overlap_window(
                        p1["START_DT"], p1["END_DT"],
                        p2["START_DT"], p2["END_DT"],
                        day_start, day_end,
                    )
                    if ov:
                        has_time_overlap = True
                        mixed = _is_mixed_gender(p1.get("GIOI_TINH", ""), p2.get("GIOI_TINH", ""))
                        if mixed:
                            has_mixed_gender = True
                        pair_key = (code, tuple(sorted([k1, k2])), ov[0], ov[1])
                        if pair_key not in seen_pairs:
                            seen_pairs.add(pair_key)
                            mins = int((ov[1] - ov[0]).total_seconds() // 60)
                            g1_txt = f" - {p1['GIOI_TINH']}" if p1.get("GIOI_TINH") else ""
                            g2_txt = f" - {p2['GIOI_TINH']}" if p2.get("GIOI_TINH") else ""
                            overlap_rows.append({
                                "Ngày": day_str,
                                "Cảnh báo": "🚨 NAM - NỮ NẰM CHUNG" if mixed else "⚠️ Trùng giờ giường",
                                "Mã khoa": ma_khoa,
                                "Tên khoa": ten_khoa,
                                "Phòng": bed_meta.get(code, {}).get("phong", ""),
                                "Mã giường": code,
                                "BN 1": f"{p1['TENBENHNHAN']} ({p1['MA_LK']}{g1_txt})",
                                "Giờ BN 1": f"{p1['START_DT'].strftime('%d/%m %H:%M')} → {p1['END_DT'].strftime('%d/%m %H:%M')}",
                                "TT BN 1 (%)": int(p1["TYLE_TT"]),
                                "BN 2": f"{p2['TENBENHNHAN']} ({p2['MA_LK']}{g2_txt})",
                                "Giờ BN 2": f"{p2['START_DT'].strftime('%d/%m %H:%M')} → {p2['END_DT'].strftime('%d/%m %H:%M')}",
                                "TT BN 2 (%)": int(p2["TYLE_TT"]),
                                "Khoảng giờ trùng": f"{ov[0].strftime('%d/%m/%Y %H:%M')} → {ov[1].strftime('%d/%m/%Y %H:%M')}",
                                "Số phút trùng": mins,
                            })

            if not occupants:
                status = "empty"
                cnt_empty += 1
            elif has_time_overlap or has_low_rate:
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
                "patients": unique_patients,
                "n_patients": len(unique_patients),
                "has_overlap": has_time_overlap,
                "has_mixed_gender": has_mixed_gender,
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

    df_ov = pd.DataFrame(overlap_rows)
    if not df_ov.empty:
        df_ov["_prio"] = df_ov["Cảnh báo"].apply(lambda x: 0 if "NAM - NỮ" in str(x) else 1)
        df_ov = df_ov.sort_values(by=["_prio", "Mã khoa", "Mã giường"]).drop(columns=["_prio"]).reset_index(drop=True)

    return dept_results, pd.DataFrame(detail_rows), df_ov


def _build_excel_bytes(df_summary, df_overlap, df_detail):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="TongHop_TheoNgay", index=False)
        df_overlap.to_excel(writer, sheet_name="Trung_Giuong", index=False)
        df_detail.to_excel(writer, sheet_name="ChiTiet_Giuong", index=False)
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
    elif bed.get("has_overlap"):
        st.warning("⚠️ Giường này có bệnh nhân trùng khung giờ nằm ghép.")

    for idx, p in enumerate(bed["patients"], start=1):
        age_str = _calc_age_str(p.get("NAM_SINH", ""))
        g_parts = [x for x in [p.get("GIOI_TINH", ""), age_str] if x]
        gt_tuoi = ", ".join(g_parts) if g_parts else "Chưa rõ"
        start_s = p["START_DT"].strftime("%d/%m/%Y %H:%M")
        end_s = p["END_DT"].strftime("%d/%m/%Y %H:%M")
        dv_txt = p.get("TEN_DICH_VU") or "Giường nội trú"
        ma_benh_txt = f" • Mã bệnh: <b>{html.escape(p['MA_BENH'])}</b>" if p.get("MA_BENH") else ""

        st.markdown(
            f"""
            <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:12px;
                        padding:14px 16px;margin-bottom:10px;box-shadow:0 1px 2px rgba(15,23,42,.04);">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px;">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <span style="display:inline-flex;align-items:center;justify-content:center;
                                     width:24px;height:24px;border-radius:50%;background:#1D4ED8;
                                     color:#FFFFFF;font-size:12px;font-weight:700;">{idx}</span>
                        <span style="font-size:14.5px;font-weight:700;color:#0F172A;">
                            {html.escape(p['TENBENHNHAN'])}
                        </span>
                    </div>
                    <span style="background:#EFF6FF;color:#1D4ED8;font-size:11.5px;font-weight:700;
                                 padding:3px 10px;border-radius:6px;border:1px solid #DBEAFE;">
                        LK: {html.escape(p['MA_LK'])}
                    </span>
                </div>
                <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px;
                            font-size:12.5px;color:#475569;margin-bottom:8px;">
                    <div>Giới / Tuổi: <b style="color:#0F172A;">{html.escape(gt_tuoi)}</b></div>
                    <div>Bắt đầu nằm: <b style="color:#0F172A;">{start_s}</b> → <b style="color:#0F172A;">{end_s}</b></div>
                </div>
                <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;
                            padding:7px 10px;font-size:12.5px;color:#334155;">
                    Dịch vụ giường: <b style="color:#0F172A;">{html.escape(dv_txt)}</b>
                    • Tỷ lệ TT: <b style="color:#15803D;">{int(p['TYLE_TT'])}%</b>
                    • Trạng thái: <b>{html.escape(p.get('TRANG_THAI_BN', ''))}</b>{ma_benh_txt}
                </div>
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
            font-size: 12px; line-height: 1.35; min-height: 76px;
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
            font-size: 11.5px; color: #1E293B; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis;
        }
        .bed-sub {
            font-size: 10.5px; color: #64748B; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
                + " ".join(f"{p['TENBENHNHAN']} {p['MA_LK']} {p.get('GIOI_TINH', '')}" for p in b["patients"])
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
                    body = meta_line + "<div class='bed-sub'>Sẵn sàng cấp giường</div>"
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
                        cls, bcls, lbl = "bed-single", "badge-single", "1 BN"
                        p = b["patients"][0]
                        g_str = f" • {p['GIOI_TINH']}" if p.get("GIOI_TINH") else ""
                        body = (
                            meta_line
                            + f"<div class='bed-bn' title='{html.escape(p['TENBENHNHAN'])}'>{html.escape(p['TENBENHNHAN'])}</div>"
                            + f"<div class='bed-sub'>LK: {html.escape(p['MA_LK'])}{html.escape(g_str)}</div>"
                        )
                    else:
                        if b.get("has_mixed_gender"):
                            cls, bcls, lbl = "bed-mixed", "badge-mixed", "🚨 NAM - NỮ"
                        elif b.get("has_overlap"):
                            cls, bcls, lbl = "bed-overlap", "badge-overlap", f"TRÙNG ({b['n_patients']})"
                        else:
                            cls, bcls, lbl = "bed-shared", "badge-shared", f"Ghép ({b['n_patients']})"
                        lines = [meta_line]
                        for p in b["patients"][:2]:
                            g_tag = f" [{p['GIOI_TINH']}]" if p.get("GIOI_TINH") else ""
                            lines.append(
                                f"<div class='bed-bn' title='{html.escape(p['TENBENHNHAN'])}'>• {html.escape(p['TENBENHNHAN'])}{html.escape(g_tag)}</div>"
                            )
                        body = "".join(lines)

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
        df_filtered = pd.DataFrame(columns=["MA_KHOA", "MA_LK", "TENBENHNHAN", "GIOI_TINH_STD", "MA_GIUONG_STD", "START_DT", "END_DT"])

    if sel_khoa:
        chosen_codes = [s.split("(")[-1].rstrip(")") for s in sel_khoa]
        cat_view = [d for d in bed_catalog if d["MA_KHOA"] in chosen_codes]
    else:
        cat_view = bed_catalog

    # ----- TÍNH ĐỘC LẬP TỪNG NGÀY MỘT -----
    daily_packs = []
    all_ov_frames = []
    all_detail_frames = []
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

        day_depts, day_detail, day_ov = build_bed_map_for_day(cat_view, df_day, d_start, d_end, day_str=date_str)
        day_n_ov = len(day_ov)
        day_n_mixed = int(day_ov["Cảnh báo"].str.contains("NAM - NỮ", na=False).sum()) if not day_ov.empty else 0

        daily_packs.append({
            "cur_d": cur_d,
            "date_str": date_str,
            "depts": day_depts,
            "detail": day_detail,
            "ov": day_ov,
            "n_ov": day_n_ov,
            "n_mixed": day_n_mixed,
            "n_time_only": day_n_ov - day_n_mixed,
        })

        if not day_ov.empty:
            all_ov_frames.append(day_ov)
        if not day_detail.empty:
            all_detail_frames.append(day_detail)
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
    df_all_summary = pd.DataFrame(all_summary_rows)

    total_beds_catalog = sum(len(d["BEDS"]) for d in cat_view)
    days_with_ov = sum(1 for p in daily_packs if p["n_ov"] > 0)
    total_ov_pairs = len(df_all_ov)
    total_mixed_pairs = sum(p["n_mixed"] for p in daily_packs)

    if total_mixed_pairs > 0:
        st.error(
            f"🚨 **CẢNH BÁO:** Phát hiện **{total_mixed_pairs}** cặp bệnh nhân **NAM – NỮ nằm chung cùng 1 mã giường**! "
            f"Mở từng ngày có cảnh báo đỏ bên dưới hoặc xem tab **⚠️ Tổng hợp trùng giường**."
        )

    kpi_items = [
        ("Số ngày xem", f"{total_days:,}", "#1565C0"),
        ("Tổng giường PD", f"{total_beds_catalog:,}", "#0288D1"),
        ("Ngày có trùng", f"{days_with_ov:,}", "#EA580C" if days_with_ov else "#15803D"),
        ("Tổng cặp trùng", f"{total_ov_pairs:,}", "#C62828" if total_ov_pairs else "#64748B"),
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

            day_actual_bn = sum(d["actual_patients"] for d in day_depts)
            day_total_pd = sum(d["total_approved"] for d in day_depts)
            day_pct = (day_actual_bn / day_total_pd * 100) if day_total_pd else 0.0

            dept_chips = []
            has_weird_overlap_day = False

            for d in day_depts:
                short_name = _short_dept_name(d["ma_khoa"], d["ten_khoa"])
                d_bn = d["actual_patients"]
                d_tot = d["total_approved"]

                if not day_ov.empty:
                    ov_sub = day_ov[day_ov["Mã khoa"] == d["ma_khoa"]]
                    d_ov = len(ov_sub)
                    d_mixed = int(ov_sub["Cảnh báo"].str.contains("NAM - NỮ", na=False).sum())
                    d_time = d_ov - d_mixed
                else:
                    d_ov = d_mixed = d_time = 0

                if d_ov > 0:
                    ov_desc = _format_overlap_desc(d_mixed, d_time)
                    if d_bn < d_tot:
                        # CÒN GIƯỜNG TRỐNG NHƯNG VẪN TRÙNG -> CẢNH BÁO ĐỎ LÒM
                        has_weird_overlap_day = True
                        dept_chips.append(
                            f":red-background[:red[**🔴 {short_name}: {d_bn}/{d_tot} ({ov_desc} - CÒN GIƯỜNG!)**]]"
                        )
                    else:
                        dept_chips.append(f":red[**{short_name}: {d_bn}/{d_tot} ({ov_desc})**]")
                elif d["shared"] > 0:
                    dept_chips.append(f":orange[**{short_name}: {d_bn}/{d_tot}**]")
                elif d_bn > 0:
                    dept_chips.append(f":blue[**{short_name}: {d_bn}/{d_tot}**]")
                else:
                    dept_chips.append(f":gray[{short_name}: 0/{d_tot}]")

            dept_summary_str = "  •  ".join(dept_chips)

            if pack["n_ov"] > 0:
                day_ov_desc = _format_overlap_desc(pack["n_mixed"], pack["n_time_only"])
                if has_weird_overlap_day:
                    status_tag = f":red-background[:red[**🚨 CÒN GIƯỜNG VẪN TRÙNG ({day_ov_desc})**]]"
                elif pack["n_mixed"] > 0:
                    status_tag = f"🚨 :red[**{day_ov_desc}**]"
                else:
                    status_tag = f"⚠️ :orange[**{day_ov_desc}**]"
            else:
                status_tag = "✅ :green[Không trùng]"

            exp_lbl = (
                f"**{date_str}**  |  {status_tag}  |  "
                f"Tổng BN: **{day_actual_bn}/{day_total_pd}** ({day_pct:.0f}%)  —  {dept_summary_str}"
            )

            with st.expander(exp_lbl, expanded=(total_days == 1)):
                if has_weird_overlap_day:
                    st.error(
                        "🔴 **CẢNH BÁO BẤT THƯỜNG:** Trong ngày này có khoa **số BN thực tế ít hơn số giường phê duyệt** "
                        "(vẫn còn giường trống) nhưng vẫn phát sinh nằm trùng giường!"
                    )

                if not day_ov.empty:
                    st.markdown("**⚠️ Danh sách cặp trùng giường trong ngày:**")
                    st.dataframe(day_ov, use_container_width=True, hide_index=True)

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
                        data=_build_excel_bytes(day_sum, day_ov, day_detail),
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
            st.error(f"⚠️ Phát hiện **{total_ov_pairs:,}** cặp trùng mã giường (trong đó có **{total_mixed_pairs}** cặp Nam – Nữ):")
            st.dataframe(df_all_ov, use_container_width=True, hide_index=True, height=460)

    # ----- TAB 3: DANH SÁCH CHI TIẾT & XUẤT EXCEL -----
    with tab_tbl:
        if df_all_detail.empty:
            st.info("Không có lượt bệnh nhân nằm giường trong khoảng thời gian đã chọn.")
        else:
            st.dataframe(df_all_detail, use_container_width=True, hide_index=True, height=460)
            st.download_button(
                "⬇️ Xuất Excel báo cáo giường theo ngày (3 sheet)",
                data=_build_excel_bytes(df_all_summary, df_all_ov, df_all_detail),
                file_name=f"BaoCao_Giuong_{range_start.strftime('%Y%m%d')}_{range_end.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
