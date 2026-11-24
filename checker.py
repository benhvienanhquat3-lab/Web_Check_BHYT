# -*- coding: utf-8 -*-
"""
checker.py
=======================================================================
Toàn bộ logic nghiệp vụ kiểm tra lỗi BHYT (XML3 / Excel) - 6 Điều kiện.
Đã nâng cấp:
    * Đọc `dm_chuyenmon` kết hợp `CHUCDANH_NN = 1` & `PHAMVI_CM` (129/132 full công khám).
    * Kiểm tra hiệu lực `DEN_NGAY` của NVYT (sau DEN_NGAY -> ngoài danh mục).
    * Đọc 2 cột cờ `Check_dùng máy` (ĐK 6) và `check trùng khoảng` (ĐK 4) trong `data_dvkt`.
=======================================================================
"""

import os
import base64
import unicodedata
import xml.etree.ElementTree as ET
from io import BytesIO
from collections import defaultdict
from datetime import datetime
import time
import re

import numpy as np
import pandas as pd

_DONG_IN_DETAIL_RE = re.compile(r"\(\s*Dòng\s+(\d+)\s*\)", re.IGNORECASE)


def extract_paired_row(chi_tiet_loi: str) -> str:
    """Lấy số dòng đối chiếu bên trong '(Dòng ...)' của cột Chi tiết lỗi."""
    if not chi_tiet_loi:
        return ""
    m = _DONG_IN_DETAIL_RE.search(str(chi_tiet_loi))
    return m.group(1).strip() if m else ""


def deduplicate_symmetric_row_pairs(df_th: pd.DataFrame) -> pd.DataFrame:
    """
    Sắp xếp theo STT tăng dần để dòng có STT nhỏ hơn luôn đứng trước,
    sau đó loại bỏ chiều ngược lại (B -> A) nếu cùng Loại lỗi & Mã DV.
    """
    if df_th is None or df_th.empty:
        return df_th
    if "Dòng" not in df_th.columns or "Chi tiết lỗi" not in df_th.columns:
        return df_th

    df = df_th.copy()

    if "STT" in df.columns:
        stt_num = pd.to_numeric(df["STT"], errors="coerce")
        df = (
            df.assign(_stt_sort=stt_num)
            .sort_values(by=["_stt_sort", "STT"], ascending=True, na_position="last", kind="mergesort")
            .drop(columns=["_stt_sort"])
            .reset_index(drop=True)
        )

    seen_keys = set()
    keep_indices = []

    for idx, row in df.iterrows():
        cur_row = str(row.get("Dòng", "") or "").strip()
        other_row = extract_paired_row(row.get("Chi tiết lỗi", ""))

        if cur_row and other_row and cur_row != other_row:
            loai_loi = str(row.get("Loại lỗi", "") or "").strip()
            ma_dv = str(row.get("Mã DV", "") or "").strip()
            pair_sorted = tuple(
                sorted([cur_row, other_row], key=lambda x: int(x) if x.isdigit() else x)
            )
            key = (loai_loi, ma_dv, pair_sorted)
            if key in seen_keys:
                continue
            seen_keys.add(key)

        keep_indices.append(idx)

    return df.loc[keep_indices].reset_index(drop=True)


# ===================== TÊN CỘT / HẰNG SỐ DÙNG CHUNG =====================
COL_DK1 = 'ĐK 1: Lỗi trùng giờ'
COL_DK2 = 'ĐK 2: CK in time CLS'
COL_DK3 = 'ĐK 3: Time Máy trùng'
COL_DK4 = 'ĐK 4: Time CLS trùng'
COL_DK5 = 'ĐK 5: Lỗi Cấu Hình (BS/Máy/DVKT)'
COL_DK6 = 'ĐK 6: Lỗi T.Gian & Thiếu Máy'

DK_COLUMNS = [COL_DK1, COL_DK2, COL_DK3, COL_DK4, COL_DK5, COL_DK6]
DK_TITLES = {col: col for col in DK_COLUMNS}

CHITIET_COLUMNS = [
    'Id', 'MA_LK',
    'TENBENHNHAN', 'TEN_KHOA', 'NGAY_RA',
    'Tên bệnh nhân', 'Khoa', 'Trạng thái',
    'MA_DICH_VU', 'MA_NHOM', 'TEN_DICH_VU',
    'TYLE_TT', 'MA_KHOA', 'MA_BAC_SI', 'MA_BENH', 'NGAY_YL',
    'NGAY_TH_YL', 'NGAY_KQ', 'NGUOI_THUC_HIEN', 'MA_MAY',
    'NGAY_VAO', 'Output_Row',
    COL_DK1, COL_DK2, COL_DK3, COL_DK4, COL_DK5, COL_DK6,
]

TH_COLUMNS = [
    'STT', 'Loại lỗi', 'Dòng',
    'Tên bệnh nhân', 'Khoa', 'Trạng thái',
    'Mã DV', 'Tên dịch vụ',
    'NGAY_TH_YL', 'NGAY_KQ', 'Mã DV_check', 'Tên dịch vụ_check',
    'NGAY_TH_YL_check', 'NGAY_KQ_check',
    'Tên bệnh nhân_check', 'Khoa_check', 'Trạng thái_check',
    'Chi tiết lỗi',
]

TEMPLATE_COLUMNS = {
    'data_nvyt': ['MACCHN', 'SO_CCCD', 'HO_TEN', 'CHUCDANH_NN', 'PHAMVI_CM', 'TU_NGAY', 'DEN_NGAY'],
    'data_tbyt': ['STT', 'TEN_TB', 'KY_HIEU', 'CONGTY_SX', 'NUOC_SX', 'NAM_SX', 'NAM_SD',
                  'MA_MAY', 'SO_LUU_HANH', 'HD_TU', 'HD_DEN', 'TU_NGAY', 'DEN_NGAY', 'ID'],
    'data_dvkt': ['MA', 'TEN', 'HIEU_LUC', 'Check_dùng máy', 'check trùng khoảng'],
    'dm_chuyenmon': ['STT', 'Tên chuyên khoa', 'Mã phạm vi', 'Mã CK', 'Tên Công khám'],
}


# ===================== HÀM TIỆN ÍCH =====================
def _clean_str(val):
    s = str(val or "").strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _normalize_date8(val):
    """Chuẩn hóa ngày bất kỳ (YYYYMMDD, YYYYMMDDHHMM, dd/mm/yyyy...) -> chuỗi 8 số YYYYMMDD."""
    s = _clean_str(val)
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 8 and s.replace(".", "").isdigit():
        return digits[:8]
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y%m%d")
    except Exception:
        pass
    return ""


def _fmt_date8(d8):
    if len(d8) == 8:
        return f"{d8[6:8]}/{d8[4:6]}/{d8[0:4]}"
    return d8


def format_time_bhyt(time_str):
    """Định dạng mốc thời gian BHYT ``yyyymmddhhmm`` -> ``dd/mm/yyyy hh:mm``."""
    time_str = str(time_str).strip()
    if time_str.endswith('.0'):
        time_str = time_str[:-2]
    if len(time_str) >= 12:
        return f"{time_str[6:8]}/{time_str[4:6]}/{time_str[0:4]} {time_str[8:10]}:{time_str[10:12]}"
    return time_str


def is_co_lydo(text):
    """True nếu chuỗi mang nghĩa 'có' (bỏ dấu, không phân biệt hoa thường)."""
    if not text:
        return False
    t = unicodedata.normalize('NFC', str(text)).strip().lower()
    return t == 'có' or t == 'co'


def _find_col(columns, candidates):
    """Tìm tên cột không phân biệt hoa/thường & khoảng trắng."""
    norm_map = {unicodedata.normalize('NFC', str(c)).strip().lower(): c for c in columns}
    for cand in candidates:
        key = unicodedata.normalize('NFC', cand).strip().lower()
        if key in norm_map:
            return norm_map[key]
    return None


# ===================== ĐỌC FILE THAM CHIẾU =====================
def load_reference_data(ref_path):
    """Đọc các sheet tham chiếu cổng BHYT (data_nvyt, dm_chuyenmon, data_tbyt, data_dvkt)."""
    catalogs = {
        'dict_ten_bs': {},             # MACCHN -> tên hiển thị
        'ref_dict': {},                # MACCHN -> info chuyên môn + hiệu lực DEN_NGAY
        'dict_may': {},                # MA_MAY -> {'ten_may':..., 'display_str':...}
        'dict_dvkt_hieuluc': {},       # MA -> HIEU_LUC
        'dict_dvkt_lydo': {},          # MA -> LY_DO_TU_CHOI (legacy)
        'dict_dvkt_dung_may': {},      # MA -> Check_dùng máy
        'dict_dvkt_trung_khoang': {},  # MA -> check trùng khoảng
        'has_col_dung_may': False,
        'has_col_trung_khoang': False,
    }
    frames = {}

    if not ref_path or not os.path.exists(ref_path):
        return catalogs, frames

    xl = pd.ExcelFile(ref_path)

    # ----- 0. Danh mục chuyên môn (dm_chuyenmon) -----
    # Map: mã phạm vi -> list of (ma_ck, ten_cong_kham)
    cm_map = defaultdict(list)
    if 'dm_chuyenmon' in xl.sheet_names:
        df_cm = xl.parse('dm_chuyenmon', dtype=str).fillna("")
        frames['dm_chuyenmon'] = df_cm
        col_pv = _find_col(df_cm.columns, ['Mã phạm vi', 'MA_PHAM_VI', 'PHAMVI_CM'])
        col_ck = _find_col(df_cm.columns, ['Mã CK', 'MA_CK'])
        col_ten_ck = _find_col(df_cm.columns, ['Tên Công khám', 'TEN_CONG_KHAM'])
        if col_pv:
            for _, r in df_cm.iterrows():
                pv = _clean_str(r.get(col_pv))
                if pv.endswith(".0"):
                    pv = pv[:-2]
                ma_ck = _clean_str(r.get(col_ck)) if col_ck else ""
                ten_ck = _clean_str(r.get(col_ten_ck)) if col_ten_ck else ""
                if pv and (ma_ck or ten_ck):
                    cm_map[pv].append((ma_ck, ten_ck))

    # ----- 1. Nhân viên y tế (data_nvyt) -----
    sheet_bs = 'data_nvyt' if 'data_nvyt' in xl.sheet_names else 0
    df_ref = xl.parse(sheet_bs, dtype=str).fillna("")
    frames['data_nvyt'] = df_ref

    col_chucdanh = _find_col(df_ref.columns, ['CHUCDANH_NN', 'CHUCDANHNN'])
    col_phamvi = _find_col(df_ref.columns, ['PHAMVI_CM'])
    col_den_ngay = _find_col(df_ref.columns, ['DEN_NGAY', 'DENNGAY'])

    for _, row in df_ref.iterrows():
        macchn = _clean_str(row.get('MACCHN', ''))
        if not macchn:
            continue
        ho_ten = _clean_str(row.get('HO_TEN', ''))
        ten_hien_thi = ho_ten if ho_ten else "k có tên"

        chucdanh_nn = _clean_str(row.get(col_chucdanh, '')) if col_chucdanh else ""
        if chucdanh_nn.endswith(".0"):
            chucdanh_nn = chucdanh_nn[:-2]

        phamvi_raw = _clean_str(row.get(col_phamvi, '')) if col_phamvi else ""
        pv_codes = [p.strip()[:-2] if p.strip().endswith(".0") else p.strip()
                    for p in phamvi_raw.split(';') if p.strip()]

        den_ngay_8 = _normalize_date8(row.get(col_den_ngay, '')) if col_den_ngay else ""

        # Xác định công khám được phép
        allow_all_exams = False
        ma_ck_list = []
        ten_ck_list = []

        if chucdanh_nn == '1' and pv_codes and cm_map:
            if '129' in pv_codes or '132' in pv_codes:
                allow_all_exams = True
            else:
                for pv in pv_codes:
                    for m_ck, t_ck in cm_map.get(pv, []):
                        if m_ck and m_ck not in ma_ck_list:
                            ma_ck_list.append(m_ck)
                        if t_ck and t_ck not in ten_ck_list:
                            ten_ck_list.append(t_ck)
        else:
            # Fallback cột Mã CK / Tên Công khám cũ nếu file chưa có dm_chuyenmon
            ma_ck_chuoi = _clean_str(row.get('Mã CK', ''))
            ten_ck_chuoi = _clean_str(row.get('Tên Công khám', ''))
            ma_ck_list = [ck.strip() for ck in ma_ck_chuoi.split(';') if ck.strip()]
            ten_ck_list = [tk.strip() for tk in ten_ck_chuoi.split(';') if tk.strip()]

        catalogs['dict_ten_bs'][macchn] = ten_hien_thi

        ma_ck_disp = "Tất cả (129/132)" if allow_all_exams else "; ".join(ma_ck_list)
        ten_ck_disp = "Tất cả công khám" if allow_all_exams else "; ".join(ten_ck_list)

        # Hỗ trợ trường hợp 1 MACCHN có nhiều dòng (nhiều giai đoạn DEN_NGAY)
        if macchn not in catalogs['ref_dict']:
            catalogs['ref_dict'][macchn] = {
                'chucdanh_nn': chucdanh_nn,
                'allow_all_exams': allow_all_exams,
                'ma_ck_list': list(ma_ck_list),
                'ten_ck_lower_set': set(t.lower() for t in ten_ck_list if t),
                'den_ngay_list': [den_ngay_8],   # "" nghĩa là không giới hạn
                'display_str': f"{macchn} - {ten_hien_thi} - PV:{phamvi_raw} - {ma_ck_disp} ({ten_ck_disp})",
            }
        else:
            entry = catalogs['ref_dict'][macchn]
            entry['allow_all_exams'] = entry['allow_all_exams'] or allow_all_exams
            for m in ma_ck_list:
                if m not in entry['ma_ck_list']:
                    entry['ma_ck_list'].append(m)
            for t in ten_ck_list:
                if t:
                    entry['ten_ck_lower_set'].add(t.lower())
            entry['den_ngay_list'].append(den_ngay_8)

    # ----- 2. Máy móc / TTB -----
    sheet_may_name = 'data_tbyt' if 'data_tbyt' in xl.sheet_names else (
        1 if len(xl.sheet_names) > 1 else None)
    if sheet_may_name is not None:
        df_may = xl.parse(sheet_may_name, dtype=str).fillna("")
        frames['data_tbyt'] = df_may
        for _, row in df_may.iterrows():
            ma_may = _clean_str(row.get('MA_MAY', ''))
            ten_tb = _clean_str(row.get('TEN_TB', ''))
            if ma_may:
                ten_hien_thi = ten_tb if ten_tb else "k có tên"
                catalogs['dict_may'][ma_may] = {
                    'ten_may': ten_hien_thi,
                    'display_str': f"{ma_may} - {ten_hien_thi}",
                }

    # ----- 3. Danh mục DVKT -----
    if 'data_dvkt' in xl.sheet_names:
        df_dvkt = xl.parse('data_dvkt', dtype=str).fillna("")
        frames['data_dvkt'] = df_dvkt

        col_dung_may = _find_col(df_dvkt.columns, ['Check_dùng máy', 'Check_dung_may', 'CHECK_DUNG_MAY'])
        col_trung_khoang = _find_col(df_dvkt.columns, ['check trùng khoảng', 'check_trung_khoang', 'CHECK_TRUNG_KHOANG'])
        catalogs['has_col_dung_may'] = bool(col_dung_may)
        catalogs['has_col_trung_khoang'] = bool(col_trung_khoang)

        for _, row in df_dvkt.iterrows():
            ma_dvkt = _clean_str(row.get('MA', ''))
            if not ma_dvkt:
                continue
            catalogs['dict_dvkt_hieuluc'][ma_dvkt] = _clean_str(row.get('HIEU_LUC', ''))
            catalogs['dict_dvkt_lydo'][ma_dvkt] = _clean_str(row.get('LY_DO_TU_CHOI', ''))
            if col_dung_may:
                catalogs['dict_dvkt_dung_may'][ma_dvkt] = _clean_str(row.get(col_dung_may, ''))
            if col_trung_khoang:
                catalogs['dict_dvkt_trung_khoang'][ma_dvkt] = _clean_str(row.get(col_trung_khoang, ''))

    return catalogs, frames


# ===================== ĐỌC DỮ LIỆU CHÍNH =====================
def parse_xml_to_dataframe(xml_source):
    """Đọc file XML gốc (FILEHOSO > NOIDUNGFILE mã hoá Base64 của XML3)."""
    try:
        if isinstance(xml_source, (bytes, bytearray)):
            root = ET.fromstring(xml_source)
        else:
            tree = ET.parse(xml_source)
            root = tree.getroot()

        records = []
        for hoso in root.findall('.//HOSO'):
            xml3_node = None
            for filehoso in hoso.findall('.//FILEHOSO'):
                loaihoso = filehoso.find('LOAIHOSO')
                if loaihoso is not None and loaihoso.text == 'XML3':
                    xml3_node = filehoso.find('NOIDUNGFILE')
                    break

            if xml3_node is not None and xml3_node.text:
                content_text = xml3_node.text.strip()
                try:
                    decoded_bytes = base64.b64decode(content_text)
                    content_xml = decoded_bytes.decode('utf-8', errors='ignore')
                except Exception:
                    content_xml = content_text

                try:
                    sub_root = ET.fromstring(content_xml)
                    for dvkt in sub_root.findall('.//CHITIET_DVKT'):
                        row_data = {}
                        for child in dvkt:
                            row_data[child.tag] = child.text if child.text is not None else ""

                        if 'MA_LK' not in row_data:
                            malk_node = hoso.find('.//MA_LK')
                            row_data['MA_LK'] = malk_node.text if malk_node is not None else ""

                        records.append(row_data)
                except Exception:
                    continue

        if records:
            df = pd.DataFrame(records)
            if 'MA_DICH_VU' not in df.columns and 'MA_DIỆCH_VU' in df.columns:
                df.rename(columns={'MA_DIỆCH_VU': 'MA_DICH_VU'}, inplace=True)
            if 'Id' not in df.columns:
                df['Id'] = np.arange(1, len(df) + 1).astype(str)
            return df
        return None
    except Exception as e:
        raise Exception(f"Lỗi đọc cấu trúc XML: {str(e)}")


def read_input_dataframe(source, filename=""):
    """Đọc dữ liệu chính từ file XML hoặc Excel (đường dẫn hoặc bytes)."""
    name = str(filename or (source if isinstance(source, str) else "")).lower()
    is_xml = name.endswith('.xml')
    if isinstance(source, (bytes, bytearray)):
        if is_xml:
            df = parse_xml_to_dataframe(bytes(source))
            if df is None or df.empty:
                raise Exception("File XML không có dữ liệu dịch vụ kỹ thuật (XML3) hợp lệ!")
        else:
            df = pd.read_excel(BytesIO(source), dtype=str)
    else:
        if is_xml:
            df = parse_xml_to_dataframe(source)
            if df is None or df.empty:
                raise Exception("File XML không có dữ liệu dịch vụ kỹ thuật (XML3) hợp lệ!")
        else:
            df = pd.read_excel(source, dtype=str)
    return df


# ===================== LÕI KIỂM TRA 6 ĐIỀU KIỆN =====================
def run_check(df, catalogs=None, ref_path="", progress_cb=None):
    """Chạy kiểm tra 6 Điều kiện trên DataFrame dữ liệu chính."""
    catalogs = catalogs or {}
    ref_dict = catalogs.get('ref_dict', {})
    dict_ten_bs = catalogs.get('dict_ten_bs', {})
    dict_may = catalogs.get('dict_may', {})
    dict_dvkt_hieuluc = catalogs.get('dict_dvkt_hieuluc', {})
    dict_dvkt_lydo = catalogs.get('dict_dvkt_lydo', {})
    dict_dvkt_dung_may = catalogs.get('dict_dvkt_dung_may', {})
    dict_dvkt_trung_khoang = catalogs.get('dict_dvkt_trung_khoang', {})
    has_col_dung_may = catalogs.get('has_col_dung_may', False)
    has_col_trung_khoang = catalogs.get('has_col_trung_khoang', False)
    has_ref = bool(ref_path)

    df = df.copy()
    df.fillna("", inplace=True)
    if 'Id' not in df.columns:
        if 'ID' in df.columns:
            df.rename(columns={'ID': 'Id'}, inplace=True)
        else:
            df['Id'] = np.arange(1, len(df) + 1).astype(str)

    if 'MA_NHOM' not in df.columns:
        df['MA_NHOM'] = "1"
    if 'TYLE_TT' not in df.columns:
        df['TYLE_TT'] = "100"
    if 'NGAY_YL' not in df.columns:
        df['NGAY_YL'] = ""

    mask_out = ~df['MA_NHOM'].astype(str).str.strip().isin(['10', '15'])
    mask_out = mask_out & (pd.to_numeric(df['TYLE_TT'], errors='coerce').fillna(0) != 0)

    df['Output_Row'] = ""
    df.loc[mask_out, 'Output_Row'] = np.arange(2, mask_out.sum() + 2).astype(str)

    for col in DK_COLUMNS:
        df[col] = ""

    df['TYLE_TT_num'] = pd.to_numeric(df['TYLE_TT'], errors='coerce').fillna(0)
    df['NGAY_YL_num'] = pd.to_numeric(df['NGAY_YL'], errors='coerce')
    df['NGAY_TH_YL_num'] = pd.to_numeric(df['NGAY_TH_YL'], errors='coerce')
    df['NGAY_KQ_num'] = pd.to_numeric(df['NGAY_KQ'], errors='coerce')

    for src in ('TENBENHNHAN', 'TEN_KHOA', 'NGAY_RA', 'NGAY_VAO'):
        if src not in df.columns:
            df[src] = ""
        df[src] = df[src].fillna("").astype(str).str.strip()
        df[src] = df[src].map(lambda v: "" if str(v).lower() in ("nan", "none", "nat") else v)

    df['Tên bệnh nhân'] = df['TENBENHNHAN']
    df['Khoa'] = df['TEN_KHOA'].where(df['TEN_KHOA'] != "", df['MA_KHOA'].fillna("").astype(str).str.strip())
    df['Trạng thái'] = np.where(df['NGAY_RA'] != "", "Ra viện", "Đang điều trị")

    valid_mask = (df['TYLE_TT_num'] > 0) & df['NGAY_TH_YL_num'].notna() & df['NGAY_KQ_num'].notna()
    records = df[valid_mask].to_dict('records')
    original_indices = df[valid_mask].index.tolist()
    total_records = len(records)

    def is_check_dk4(ma_dv_can_check):
        if has_col_trung_khoang:
            if ma_dv_can_check in dict_dvkt_trung_khoang:
                return is_co_lydo(dict_dvkt_trung_khoang[ma_dv_can_check])
            return False
        if not dict_dvkt_hieuluc:
            return True
        if ma_dv_can_check in dict_dvkt_hieuluc:
            return dict_dvkt_hieuluc[ma_dv_can_check] != ""
        return True

    def is_require_machine(ma_dv_can_check):
        if has_col_dung_may:
            return is_co_lydo(dict_dvkt_dung_may.get(ma_dv_can_check, ""))
        return is_co_lydo(dict_dvkt_lydo.get(ma_dv_can_check, ""))

    def is_bs_expired(bs_code, check_date8):
        """Trả về (True, max_den_ngay) nếu check_date8 vượt quá DEN_NGAY của NVYT."""
        if not check_date8 or bs_code not in ref_dict:
            return False, ""
        den_list = ref_dict[bs_code].get('den_ngay_list', [])
        if not den_list or any(d == "" for d in den_list):
            return False, ""
        max_den = max(den_list)
        if check_date8 > max_den:
            return True, max_den
        return False, ""

    # ----- Tiền xử lý: gom nhóm index theo bác sĩ (ĐK 1,2,4) & mã máy (ĐK 3) -----
    doc_index = defaultdict(list)
    may_index = defaultdict(list)
    for pos, rec in enumerate(records):
        for bs in [b.strip() for b in str(rec.get('NGUOI_THUC_HIEN', '')).split(';') if b.strip()]:
            doc_index[bs].append(pos)
        may_chuoi = str(rec.get('MA_MAY', '')).strip()
        if may_chuoi:
            may_index[may_chuoi].append(pos)

    th_records_raw = []
    if progress_cb:
        progress_cb(0.0, "Đang phân tích dữ liệu...")

    info = []
    for rec in records:
        may_chuoi = str(rec.get('MA_MAY', '')).strip()
        info.append({
            'o_id': str(rec.get('Id', '')),
            'o_row_num': str(rec.get('Output_Row', '')).strip(),
            'o_ma_nhom': str(rec.get('MA_NHOM', '')).strip(),
            'o_ten_dv': str(rec.get('TEN_DICH_VU', '')).strip(),
            'o_ma_dv': str(rec.get('MA_DICH_VU', '')).strip(),
            'o_th': rec.get('NGAY_TH_YL_num'),
            'o_kq': rec.get('NGAY_KQ_num'),
            'o_bs_set': set(b.strip() for b in str(rec.get('NGUOI_THUC_HIEN', '')).split(';') if b.strip()),
            'o_may_set': set([may_chuoi]) if may_chuoi else set(),
        })

    start_time = time.time()

    # ===================== VÒNG LẶP CHÍNH =====================
    for i, row in enumerate(records):
        if progress_cb and (i % 200 == 0 or (i + 1) == total_records):
            elapsed = time.time() - start_time
            speed = (i + 1) / elapsed if elapsed > 0 else 0
            eta = (total_records - (i + 1)) / speed if speed > 0 else 0
            m, s = divmod(int(eta), 60)
            time_str = f"{m} phút {s} giây" if m > 0 else f"{s} giây"
            percent = ((i + 1) / total_records) if total_records else 0.0
            progress_cb(percent, f"Đang soi: {i + 1}/{total_records} ({percent * 100:.1f}%) - còn ~{time_str}")

        idx = original_indices[i]
        id_hien_tai = str(row.get('Id', ''))
        ma_nhom = str(row.get('MA_NHOM', '')).strip()
        ten_dv = str(row.get('TEN_DICH_VU', '')).strip()
        ma_dv = str(row.get('MA_DICH_VU', '')).strip()

        ngay_yl = row.get('NGAY_YL_num')
        ngay_th = row.get('NGAY_TH_YL_num')
        ngay_kq = row.get('NGAY_KQ_num')

        d8_yl = _normalize_date8(row.get('NGAY_YL', ''))
        d8_th = _normalize_date8(row.get('NGAY_TH_YL', ''))

        ds_bac_si = [bs.strip() for bs in str(row.get('NGUOI_THUC_HIEN', '')).split(';') if bs.strip()]
        ds_bs_yl = [bs.strip() for bs in str(row.get('MA_BAC_SI', '')).split(';') if bs.strip()]
        may_chuoi = str(row.get('MA_MAY', '')).strip()
        ds_may = [may_chuoi] if may_chuoi else []

        loi_dk1, loi_dk2, loi_dk3, loi_dk4, loi_dk5, loi_dk6 = [], [], [], [], [], []
        dict_th_dk1, dict_th_dk2, dict_th_dk3 = {}, {}, {}
        dict_th_dk4, dict_th_dk5, dict_th_dk6 = {}, {}, {}

        # ----- ĐIỀU KIỆN 6 (logic luồng thời gian & thiếu máy) -----
        if pd.notna(ngay_yl) and pd.notna(ngay_th):
            if ngay_th < ngay_yl:
                l_str = "Sai logic TGian (TH < YL)"
                loi_dk6.append(l_str)
                dict_th_dk6[l_str] = None

        if pd.notna(ngay_kq) and pd.notna(ngay_th):
            if ngay_kq <= ngay_th:
                l_str = "Sai logic TGian (KQ <= TH)"
                loi_dk6.append(l_str)
                dict_th_dk6[l_str] = None

        if has_ref and (has_col_dung_may or dict_dvkt_lydo):
            if is_require_machine(ma_dv) and not may_chuoi:
                l_str = "Thiếu mã máy (Y/c ĐK 6)"
                loi_dk6.append(l_str)
                dict_th_dk6[l_str] = None

        if loi_dk6:
            df.at[idx, COL_DK6] = " | ".join(list(set(loi_dk6)))
            for l_str, o_row in dict_th_dk6.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK6, 'chi_tiet_loi': l_str, 'other_row': o_row})

        # ----- ĐIỀU KIỆN 5 (cấu hình BS / Hiệu lực DEN_NGAY / Công khám / Máy / DVKT) -----
        if has_ref:
            # Kiểm tra bác sĩ ra y lệnh (nếu có DEN_NGAY)
            for bs_yl in ds_bs_yl:
                if bs_yl in ref_dict:
                    expired, max_den = is_bs_expired(bs_yl, d8_yl)
                    if expired:
                        l_str = f"BS ra YL ngoài danh mục / Hết hiệu lực ({bs_yl} - đến {_fmt_date8(max_den)})"
                        loi_dk5.append(l_str)
                        dict_th_dk5[l_str] = None

            # Kiểm tra người thực hiện
            if ds_bac_si:
                for bs in ds_bac_si:
                    if bs not in dict_ten_bs:
                        l_str = f"BS lạ ({bs})"
                        loi_dk5.append(l_str)
                        dict_th_dk5[l_str] = None
                        continue

                    expired, max_den = is_bs_expired(bs, d8_th or d8_yl)
                    if expired:
                        l_str = f"BS thực hiện ngoài danh mục / Hết hiệu lực ({bs} - đến {_fmt_date8(max_den)})"
                        loi_dk5.append(l_str)
                        dict_th_dk5[l_str] = None
                        continue

                    if ma_nhom == '13':
                        bs_info = ref_dict[bs]
                        if not bs_info.get('allow_all_exams', False):
                            ma_ck_hop_le = bs_info.get('ma_ck_list', [])
                            ten_ck_hop_le = bs_info.get('ten_ck_lower_set', set())
                            ok_ma = (ma_dv in ma_ck_hop_le) or any(ma_dv.startswith(ck + ".") for ck in ma_ck_hop_le)
                            ok_ten = (ten_dv.lower() in ten_ck_hop_le) if ten_dv else False
                            if not (ok_ma or ok_ten):
                                l_str = f"Sai CK ({bs_info['display_str']})"
                                loi_dk5.append(l_str)
                                dict_th_dk5[l_str] = None

            if ds_may and dict_may:
                for may in ds_may:
                    if may not in dict_may:
                        l_str = f"Máy chưa có trên Cổng ({may})"
                        loi_dk5.append(l_str)
                        dict_th_dk5[l_str] = None

            if dict_dvkt_hieuluc:
                if ma_dv and ma_dv not in dict_dvkt_hieuluc:
                    l_str = f"DVKT chưa có trên Cổng ({ma_dv})"
                    loi_dk5.append(l_str)
                    dict_th_dk5[l_str] = None

        if loi_dk5:
            df.at[idx, COL_DK5] = " | ".join(list(set(loi_dk5)))
            for l_str, o_row in dict_th_dk5.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK5, 'chi_tiet_loi': l_str, 'other_row': o_row})

        is_current_dk4 = is_check_dk4(ma_dv)

        # ----- TẬP ỨNG VIÊN (chỉ xét bản ghi chia sẻ bác sĩ / mã máy) -----
        cand = set()
        for bs in ds_bac_si:
            lst = doc_index.get(bs)
            if lst:
                cand.update(lst)
        for may in ds_may:
            lst = may_index.get(may)
            if lst:
                cand.update(lst)
        cand.discard(i)
        candidates = sorted(cand)

        # ----- ĐIỀU KIỆN 1-4 -----
        for j in candidates:
            o = info[j]
            o_id = o['o_id']
            if o_id == id_hien_tai:
                continue

            o_row_num = o['o_row_num']
            row_str = f"Dòng {o_row_num}" if o_row_num else "Dòng bị ẩn"

            o_ma_nhom = o['o_ma_nhom']
            o_ten_dv = o['o_ten_dv']
            o_ma_dv = o['o_ma_dv']
            o_th, o_kq = o['o_th'], o['o_kq']
            o_ds_bac_si = o['o_bs_set']
            o_ds_may = o['o_may_set']

            if ma_nhom not in ['1', ''] and o_ma_nhom not in ['1', ''] and ds_bac_si:
                if (o_th == ngay_th) or (o_kq == ngay_th) or (o_th == ngay_kq) or (o_kq == ngay_kq):
                    for bs in ds_bac_si:
                        if bs in o_ds_bac_si:
                            bs_str = f"{bs} - {dict_ten_bs.get(bs, 'k có tên')}" if has_ref else bs
                            l_str = f"Trùng ID: {o_id} ({row_str}) ({bs_str})"
                            loi_dk1.append(f"{o_id} ({row_str}) ({bs_str})")
                            dict_th_dk1[l_str] = records[j]

            if ma_nhom == '13' and o_ma_nhom not in ['1', '10', '13', ''] and ds_bac_si:
                if (o_th <= ngay_th <= o_kq) or (o_th <= ngay_kq <= o_kq):
                    for bs in ds_bac_si:
                        if bs in o_ds_bac_si:
                            bs_str = f"{bs} - {dict_ten_bs.get(bs, 'k có tên')}" if has_ref else bs
                            l_str = f"Lỗi đè ID: {o_id} ({row_str}) ({bs_str})"
                            loi_dk2.append(f"{o_id} ({row_str}) ({bs_str})")
                            dict_th_dk2[l_str] = records[j]

            if ma_nhom not in ['1', '10', '13', ''] and ten_dv != "Sắc thuốc thang" and ds_may:
                if o_ma_nhom not in ['1', '10', '13', ''] and o_ten_dv != "Sắc thuốc thang":
                    if (o_th <= ngay_th <= o_kq) or (o_th <= ngay_kq <= o_kq):
                        for may in ds_may:
                            if may in o_ds_may:
                                may_str = (f"{may} - {dict_may.get(may, {}).get('ten_may', 'k có tên')}"
                                           if (has_ref and dict_may) else may)
                                l_str = f"Lỗi máy đè ID: {o_id} ({row_str}) ({may_str})"
                                loi_dk3.append(f"{o_id} ({row_str}) ({may_str})")
                                dict_th_dk3[l_str] = records[j]

            if is_current_dk4:
                if ma_nhom not in ['1', '10', '13', ''] and o_ma_nhom not in ['1', '10', '13', ''] and ds_bac_si:
                    if is_check_dk4(o_ma_dv):
                        if (o_th <= ngay_kq) and (o_kq >= ngay_th):
                            for bs in ds_bac_si:
                                if bs in o_ds_bac_si:
                                    bs_str = f"{bs} - {dict_ten_bs.get(bs, 'k có tên')}" if has_ref else bs
                                    l_str = f"Lỗi đè ID: {o_id} ({row_str}) ({bs_str})"
                                    loi_dk4.append(f"{o_id} ({row_str}) ({bs_str})")
                                    dict_th_dk4[l_str] = records[j]

        if loi_dk1:
            df.at[idx, COL_DK1] = "Trùng ID: " + ", ".join(list(set(loi_dk1)))
            for l_str, o_row in dict_th_dk1.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK1, 'chi_tiet_loi': l_str, 'other_row': o_row})
        if loi_dk2:
            df.at[idx, COL_DK2] = "Lỗi đè ID: " + ", ".join(list(set(loi_dk2)))
            for l_str, o_row in dict_th_dk2.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK2, 'chi_tiet_loi': l_str, 'other_row': o_row})
        if loi_dk3:
            df.at[idx, COL_DK3] = "Lỗi máy đè ID: " + ", ".join(list(set(loi_dk3)))
            for l_str, o_row in dict_th_dk3.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK3, 'chi_tiet_loi': l_str, 'other_row': o_row})
        if loi_dk4:
            df.at[idx, COL_DK4] = "Lỗi đè ID: " + ", ".join(list(set(loi_dk4)))
            for l_str, o_row in dict_th_dk4.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK4, 'chi_tiet_loi': l_str, 'other_row': o_row})

    if progress_cb:
        progress_cb(1.0, "Đang đóng gói dữ liệu và xuất kết quả...")

    return _build_results(df, mask_out, th_records_raw)


# ===================== ĐÓNG GÓI KẾT QUẢ =====================
def _build_results(df, mask_out, th_records_raw):
    """Tạo sheet ``chitiet`` + ``TH`` + thống kê từ dữ liệu đã kiểm tra."""
    df_out = df[mask_out].copy()
    cot_thuc_te = [col for col in CHITIET_COLUMNS if col in df_out.columns]
    df_out = df_out[cot_thuc_te]

    for col in ['NGAY_YL', 'NGAY_TH_YL', 'NGAY_KQ']:
        if col in df_out.columns:
            df_out[col] = df_out[col].apply(format_time_bhyt)

    dk_order = {col: i for i, col in enumerate(DK_COLUMNS)}
    th_data_list = []

    for item in th_records_raw:
        idx = item['idx']
        if idx not in df_out.index:
            continue

        row_main = df_out.loc[idx]
        o_row = item.get('other_row')
        ten_bn_check = ""
        khoa_check = ""
        trang_thai_check = ""

        if o_row is not None:
            if isinstance(o_row, (pd.Series, dict)):
                ten_bn_check = str(o_row.get('Tên bệnh nhân', o_row.get('TEN_BN', '')) or '').strip()
                khoa_check = str(o_row.get('Khoa', o_row.get('TEN_KHOA', o_row.get('MA_KHOA', ''))) or '').strip()
                trang_thai_check = str(o_row.get('Trạng thái', o_row.get('TRANG_THAI', '')) or '').strip()
            elif o_row in df.index:
                r_other = df.loc[o_row]
                ten_bn_check = str(r_other.get('Tên bệnh nhân', r_other.get('TEN_BN', '')) or '').strip()
                khoa_check = str(r_other.get('Khoa', r_other.get('TEN_KHOA', r_other.get('MA_KHOA', ''))) or '').strip()
                trang_thai_check = str(r_other.get('Trạng thái', r_other.get('TRANG_THAI', '')) or '').strip()

        loai_loi_key = item['loai_loi']
        chi_tiet_loi = item['chi_tiet_loi']

        dong = str(row_main.get('Output_Row', ''))
        ten_bn = str(row_main.get('Tên bệnh nhân', ''))
        khoa = str(row_main.get('Khoa', ''))
        trang_thai = str(row_main.get('Trạng thái', ''))
        ma_dv = str(row_main.get('MA_DICH_VU', ''))
        ten_dv = str(row_main.get('TEN_DICH_VU', ''))
        ngay_th_yl = str(row_main.get('NGAY_TH_YL', ''))
        ngay_kq = str(row_main.get('NGAY_KQ', ''))

        if o_row:
            ma_dv_check = str(o_row.get('MA_DICH_VU', ''))
            ten_dv_check = str(o_row.get('TEN_DICH_VU', ''))
            ngay_th_yl_check = format_time_bhyt(str(o_row.get('NGAY_TH_YL', '')))
            ngay_kq_check = format_time_bhyt(str(o_row.get('NGAY_KQ', '')))
        else:
            ma_dv_check, ten_dv_check = "", ""
            ngay_th_yl_check, ngay_kq_check = "", ""

        th_data_list.append({
            'STT': 0,
            'Loại lỗi': DK_TITLES[loai_loi_key],
            'Dòng': dong,
            'Tên bệnh nhân': ten_bn,
            'Khoa': khoa,
            'Trạng thái': trang_thai,
            'Mã DV': ma_dv,
            'Tên dịch vụ': ten_dv,
            'NGAY_TH_YL': ngay_th_yl,
            'NGAY_KQ': ngay_kq,
            'Mã DV_check': ma_dv_check,
            'Tên dịch vụ_check': ten_dv_check,
            'NGAY_TH_YL_check': ngay_th_yl_check,
            'NGAY_KQ_check': ngay_kq_check,
            'Tên bệnh nhân_check': ten_bn_check,
            'Khoa_check': khoa_check,
            'Trạng thái_check': trang_thai_check,
            'Chi tiết lỗi': chi_tiet_loi,
            '_order_dk': dk_order[loai_loi_key],
            '_dong_int': int(dong) if dong.isdigit() else 999999,
        })

    df_th = pd.DataFrame(th_data_list)
    if not df_th.empty:
        df_th = df_th.sort_values(by=['_order_dk', '_dong_int'])
        df_th = df_th.drop(columns=['_order_dk', '_dong_int'])
        df_th['STT'] = np.arange(1, len(df_th) + 1)
        df_th = df_th[TH_COLUMNS]
        df_th = deduplicate_symmetric_row_pairs(df_th)
        if not df_th.empty:
            df_th['STT'] = np.arange(1, len(df_th) + 1)
    else:
        df_th = pd.DataFrame(columns=TH_COLUMNS)

    stats = {
        'total_input': int(len(df)),
        'total_rows': int(len(df_out)),
        'total_errors': int(len(df_th)),
        'by_dk': {},
    }
    for col in DK_COLUMNS:
        stats['by_dk'][col] = int((df_th['Loại lỗi'] == col).sum()) if not df_th.empty else 0

    return {'df_out': df_out, 'df_th': df_th, 'stats': stats}


# ===================== XUẤT EXCEL =====================
def build_template_excel_bytes():
    """Tạo file Excel mẫu chuẩn danh mục cổng BHYT."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, columns in TEMPLATE_COLUMNS.items():
            pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet_name, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def build_output_excel_bytes(df_th, df_out):
    """Tạo file kết quả gồm 2 sheet ``TH`` và ``chitiet``."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_th.to_excel(writer, sheet_name='TH', index=False)
        df_out.to_excel(writer, sheet_name='chitiet', index=False)
    buffer.seek(0)
    return buffer.getvalue()


def make_output_filename(original_filename):
    """Sinh tên file kết quả chuẩn ``DaCheck_<tên_gốc>_<HHMMSS>.xlsx``."""
    base = os.path.basename(str(original_filename or "DuLieu"))
    ten_khong_duoi, _ = os.path.splitext(base)
    if not ten_khong_duoi:
        ten_khong_duoi = "DuLieu"
    thoi_gian_thuc = datetime.now().strftime("%H%M%S")
    return f"DaCheck_{ten_khong_duoi}_{thoi_gian_thuc}.xlsx"
