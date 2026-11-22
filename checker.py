# -*- coding: utf-8 -*-
"""
checker.py
=======================================================================
Toàn bộ logic nghiệp vụ kiểm tra lỗi BHYT (XML3 / Excel) - 6 Điều kiện.
Kế thừa 100% logic từ check_bhyt.py (bản desktop Tkinter), gồm:

    * Đọc XML3 (giải mã Base64) và Excel  -> ``read_input_dataframe``
    * Đọc file tham chiếu 3 sheet          -> ``load_reference_data``
    * Tạo file mẫu ``MAU_THAM_CHIEU.xlsx`` -> ``build_template_excel_bytes``
    * Kiểm tra đủ 6 Điều kiện              -> ``run_check``
    * Xuất kết quả 2 sheet TH + chitiet    -> ``build_output_excel_bytes``

Tối ưu tốc độ (kết quả KHÔNG đổi, đã đối chiếu 100% với bản O(n^2)):
    * ĐK 1, 2, 4 : gom nhóm index theo ``NGUOI_THUC_HIEN`` (bác sĩ).
    * ĐK 3       : gom nhóm index theo ``MA_MAY`` (máy).
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

import numpy as np
import pandas as pd

# ===================== TÊN CỘT / HẰNG SỐ DÙNG CHUNG =====================
COL_DK1 = 'ĐK 1: Lỗi trùng giờ'
COL_DK2 = 'ĐK 2: CK in time CLS'
COL_DK3 = 'ĐK 3: Time Máy trùng'
COL_DK4 = 'ĐK 4: Time CLS trùng'
COL_DK5 = 'ĐK 5: Lỗi Cấu Hình (BS/Máy/DVKT)'
COL_DK6 = 'ĐK 6: Lỗi T.Gian & Thiếu Máy'

DK_COLUMNS = [COL_DK1, COL_DK2, COL_DK3, COL_DK4, COL_DK5, COL_DK6]

DK_TITLES = {col: col for col in DK_COLUMNS}

# Cột chuẩn của sheet chitiet (giữ nguyên như check_bhyt.py)
CHITIET_COLUMNS = [
    'Id', 'MA_LK', 'MA_DICH_VU', 'MA_NHOM', 'TEN_DICH_VU',
    'TYLE_TT', 'MA_KHOA', 'MA_BAC_SI', 'MA_BENH', 'NGAY_YL',
    'NGAY_TH_YL', 'NGAY_KQ', 'NGUOI_THUC_HIEN', 'MA_MAY', 'Output_Row',
    COL_DK1, COL_DK2, COL_DK3, COL_DK4, COL_DK5, COL_DK6,
]

# Cột chuẩn của sheet TH (tổng hợp lỗi)
TH_COLUMNS = [
    'STT', 'Loại lỗi', 'Chi tiết lỗi', 'Dòng', 'Mã DV', 'Tên dịch vụ',
    'NGAY_TH_YL', 'NGAY_KQ', 'Mã DV_check', 'Tên dịch vụ_check',
    'NGAY_TH_YL_check', 'NGAY_KQ_check',
]

# Cột chuẩn cho 3 sheet tham chiếu của file mẫu
TEMPLATE_COLUMNS = {
    'data_nvyt': ['MACCHN', 'SO_CCCD', 'HO_TEN', 'CHUCDANH_NN', 'PHAMVI_CM', 'Mã CK', 'Tên Công khám'],
    'data_tbyt': ['STT', 'TEN_TB', 'KY_HIEU', 'CONGTY_SX', 'NUOC_SX', 'NAM_SX', 'NAM_SD',
                  'MA_MAY', 'SO_LUU_HANH', 'HD_TU', 'HD_DEN', 'TU_NGAY', 'DEN_NGAY', 'ID'],
    'data_dvkt': ['MA', 'TEN', 'HIEU_LUC', 'LY_DO_TU_CHOI'],
}


# ===================== HÀM TIỆN ÍCH =====================
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


# ===================== ĐỌC FILE THAM CHIẾU =====================
def load_reference_data(ref_path):
    """Đọc 3 sheet tham chiếu cổng BHYT.

    Trả về ``(catalogs, frames)`` trong đó:
        * ``catalogs`` : dict các lookup phục vụ 6 Điều kiện.
        * ``frames``   : dict DataFrame gốc 3 sheet (để hiển thị UI).
    """
    catalogs = {
        'dict_ten_bs': {},          # MACCHN -> tên hiển thị
        'ref_dict': {},             # MACCHN -> {'ma_ck_list':[...], 'display_str': ...}
        'dict_may': {},             # MA_MAY -> {'ten_may':..., 'display_str':...}
        'dict_dvkt_hieuluc': {},    # MA -> HIEU_LUC
        'dict_dvkt_lydo': {},       # MA -> LY_DO_TU_CHOI
    }
    frames = {}

    if not ref_path or not os.path.exists(ref_path):
        return catalogs, frames

    xl = pd.ExcelFile(ref_path)

    # ----- 1. Nhân viên y tế (bác sĩ) -----
    sheet_bs = 'data_nvyt' if 'data_nvyt' in xl.sheet_names else 0
    df_ref = xl.parse(sheet_bs, dtype=str)
    df_ref.fillna("", inplace=True)
    frames['data_nvyt'] = df_ref
    for _, row in df_ref.iterrows():
        macchn = str(row.get('MACCHN', '')).strip()
        ho_ten = str(row.get('HO_TEN', '')).strip()
        ma_ck_chuoi = str(row.get('Mã CK', '')).strip()
        ten_ck_chuoi = str(row.get('Tên Công khám', '')).strip()
        if macchn:
            ten_hien_thi = ho_ten if ho_ten else "k có tên"
            catalogs['dict_ten_bs'][macchn] = ten_hien_thi
            catalogs['ref_dict'][macchn] = {
                'ma_ck_list': [ck.strip() for ck in ma_ck_chuoi.split(';') if ck.strip()],
                'display_str': f"{macchn} - {ten_hien_thi} - {ma_ck_chuoi} - {ten_ck_chuoi}",
            }

    # ----- 2. Máy móc / TTB -----
    sheet_may_name = 'data_tbyt' if 'data_tbyt' in xl.sheet_names else (
        1 if len(xl.sheet_names) > 1 else None)
    if sheet_may_name is not None:
        df_may = xl.parse(sheet_may_name, dtype=str)
        df_may.fillna("", inplace=True)
        frames['data_tbyt'] = df_may
        for _, row in df_may.iterrows():
            ma_may = str(row.get('MA_MAY', '')).strip()
            ten_tb = str(row.get('TEN_TB', '')).strip()
            if ma_may:
                ten_hien_thi = ten_tb if ten_tb else "k có tên"
                catalogs['dict_may'][ma_may] = {
                    'ten_may': ten_hien_thi,
                    'display_str': f"{ma_may} - {ten_hien_thi}",
                }

    # ----- 3. Danh mục DVKT -----
    if 'data_dvkt' in xl.sheet_names:
        df_dvkt = xl.parse('data_dvkt', dtype=str)
        df_dvkt.fillna("", inplace=True)
        frames['data_dvkt'] = df_dvkt
        for _, row in df_dvkt.iterrows():
            ma_dvkt = str(row.get('MA', '')).strip()
            hieu_luc = str(row.get('HIEU_LUC', '')).strip()
            lydo = str(row.get('LY_DO_TU_CHOI', '')).strip()
            if hieu_luc.lower() == 'nan':
                hieu_luc = ""
            if lydo.lower() == 'nan':
                lydo = ""
            if ma_dvkt:
                catalogs['dict_dvkt_hieuluc'][ma_dvkt] = hieu_luc
                catalogs['dict_dvkt_lydo'][ma_dvkt] = lydo

    return catalogs, frames


# ===================== ĐỌC DỮ LIỆU CHÍNH =====================
def parse_xml_to_dataframe(xml_source):
    """Đọc file XML gốc (FILEHOSO > NOIDUNGFILE mã hoá Base64 của XML3).

    ``xml_source`` có thể là đường dẫn (str) hoặc bytes.
    """
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
    """Chạy kiểm tra 6 Điều kiện trên DataFrame dữ liệu chính.

    ``progress_cb`` (tuỳ chọn) là hàm ``cb(percent: float, message: str)``.

    Trả về dict:
        ``df_out``  : DataFrame sheet ``chitiet`` (đã chuẩn hoá cột).
        ``df_th``   : DataFrame sheet ``TH`` (tổng hợp lỗi).
        ``stats``   : dict thống kê (tổng dòng, tổng lỗi, số lỗi theo ĐK).
    """
    catalogs = catalogs or {}
    ref_dict = catalogs.get('ref_dict', {})
    dict_ten_bs = catalogs.get('dict_ten_bs', {})
    dict_may = catalogs.get('dict_may', {})
    dict_dvkt_hieuluc = catalogs.get('dict_dvkt_hieuluc', {})
    dict_dvkt_lydo = catalogs.get('dict_dvkt_lydo', {})
    has_ref = bool(ref_path)

    df = df.copy()
    df.fillna("", inplace=True)
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

    # Dòng xuất ra file chitiet (loại MA_NHOM 10, 15 và TYLE_TT = 0)
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

    valid_mask = (df['TYLE_TT_num'] > 0) & df['NGAY_TH_YL_num'].notna() & df['NGAY_KQ_num'].notna()
    records = df[valid_mask].to_dict('records')
    original_indices = df[valid_mask].index.tolist()
    total_records = len(records)

    def is_check_dk4(ma_dv_can_check):
        if not dict_dvkt_hieuluc:
            return True
        if ma_dv_can_check in dict_dvkt_hieuluc:
            return dict_dvkt_hieuluc[ma_dv_can_check] != ""
        return True

    # ----- Tiền xử lý: gom nhóm index theo bác sĩ (ĐK 1,2,4) & mã máy (ĐK 3) -----
    doc_index = defaultdict(list)   # MACCHN -> [vị trí record] (ascending)
    may_index = defaultdict(list)   # MA_MAY -> [vị trí record] (ascending)
    for pos, rec in enumerate(records):
        for bs in [b.strip() for b in str(rec.get('NGUOI_THUC_HIEN', '')).split(';') if b.strip()]:
            doc_index[bs].append(pos)
        may_chuoi = str(rec.get('MA_MAY', '')).strip()
        if may_chuoi:
            may_index[may_chuoi].append(pos)

    th_records_raw = []
    if progress_cb:
        progress_cb(0.0, "Đang phân tích dữ liệu...")

    # ----- Tiền xử lý thông tin mỗi record (1 lần duy nhất) -----
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

        ds_bac_si = [bs.strip() for bs in str(row.get('NGUOI_THUC_HIEN', '')).split(';') if bs.strip()]
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

        if has_ref and dict_dvkt_lydo:
            lydo_val = dict_dvkt_lydo.get(ma_dv, "")
            if is_co_lydo(lydo_val) and not may_chuoi:
                l_str = "Thiếu mã máy (Y/c ĐK 6)"
                loi_dk6.append(l_str)
                dict_th_dk6[l_str] = None

        if loi_dk6:
            df.at[idx, COL_DK6] = " | ".join(list(set(loi_dk6)))
            for l_str, o_row in dict_th_dk6.items():
                th_records_raw.append({'idx': idx, 'loai_loi': COL_DK6, 'chi_tiet_loi': l_str, 'other_row': o_row})

        # ----- ĐIỀU KIỆN 5 (cấu hình BS / Máy / DVKT) -----
        if has_ref:
            if ds_bac_si:
                for bs in ds_bac_si:
                    if bs not in dict_ten_bs:
                        l_str = f"BS lạ ({bs})"
                        loi_dk5.append(l_str)
                        dict_th_dk5[l_str] = None
                    elif ma_nhom == '13':
                        ma_ck_hop_le = ref_dict[bs]['ma_ck_list']
                        if ma_dv not in ma_ck_hop_le:
                            l_str = f"Sai CK ({ref_dict[bs]['display_str']})"
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

        # ----- ĐIỀU KIỆN 1-4 (so sánh chéo trên tập ứng viên) -----
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
        loai_loi_key = item['loai_loi']
        chi_tiet_loi = item['chi_tiet_loi']

        dong = str(row_main.get('Output_Row', ''))
        ma_dv = str(row_main.get('MA_DICH_VU', ''))
        ten_dv = str(row_main.get('TEN_DICH_VU', ''))
        ngay_th_yl = str(row_main.get('NGAY_TH_YL', ''))
        ngay_kq = str(row_main.get('NGAY_KQ', ''))

        o_row = item['other_row']
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
            'Chi tiết lỗi': chi_tiet_loi,
            'Dòng': dong,
            'Mã DV': ma_dv,
            'Tên dịch vụ': ten_dv,
            'NGAY_TH_YL': ngay_th_yl,
            'NGAY_KQ': ngay_kq,
            'Mã DV_check': ma_dv_check,
            'Tên dịch vụ_check': ten_dv_check,
            'NGAY_TH_YL_check': ngay_th_yl_check,
            'NGAY_KQ_check': ngay_kq_check,
            '_order_dk': dk_order[loai_loi_key],
            '_dong_int': int(dong) if dong.isdigit() else 999999,
        })

    df_th = pd.DataFrame(th_data_list)
    if not df_th.empty:
        df_th = df_th.sort_values(by=['_order_dk', '_dong_int'])
        df_th = df_th.drop(columns=['_order_dk', '_dong_int'])
        df_th['STT'] = np.arange(1, len(df_th) + 1)
        df_th = df_th[TH_COLUMNS]
    else:
        df_th = pd.DataFrame(columns=TH_COLUMNS)

    # ----- Thống kê -----
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
    """Tạo file Excel mẫu 3 sheet chuẩn danh mục cổng BHYT."""
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