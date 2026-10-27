# -*- coding: utf-8 -*-
"""
lazada_core.py — Core logic ของ Lazada SKU Analysis
พอร์ตมาจาก Skill "lazada-sku-analysis" (scripts/build_combined.py + scripts/build_analysis.py)
โดยคง logic เดิมทุกจุด แค่ห่อเป็นฟังก์ชันให้เว็บแอปเรียกใช้ได้

ฟังก์ชันหลัก:
  - build_combined(base_dir, out_path)   : รวมไฟล์ดิบ -> input_combined.xlsx
  - build_analysis(input_path, output_path, cost_map=None, note_map=None)
        : วิเคราะห์ -> ค่าคอมLAZADA.xlsx (3 ชีทรายงาน)
        ถ้าใส่ cost_map / note_map (จาก Google Sheets API) จะคำนวณค่าจริงแทนสูตร IMPORTRANGE
"""
import os
import glob
import shutil
import tempfile
import subprocess
import warnings
from copy import copy

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

SHOPS = ['LAZADA-MAFIA', 'LAZADA-TRC', 'LAZADA-UTOPIA', 'LAZADA-FREEROAD', 'LAZADA-WORKFORCE']

COST_URL = "https://docs.google.com/spreadsheets/d/1whwXYiUyDC0uo6oaudr_QqzCV6UFOWd0mL4Y2El9JRY/edit?usp=sharing"
NOTE_URL = "https://docs.google.com/spreadsheets/d/1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y/edit?usp=sharing"
COST_SHEET_ID = "1whwXYiUyDC0uo6oaudr_QqzCV6UFOWd0mL4Y2El9JRY"
NOTE_SHEET_ID = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"

INCOME_FEES = {
    'ยอดรวมค่าสินค้า', 'คืนค่าธรรมเนียมโปรแกรมส่วนลด LazCoins',
    'คืนค่าธรรมเนียมโปรแกรมส่งฟรีพิเศษกับลาซาด้า', 'คืนส่วนลดค่าธรรมเนียมการขายสินค้า',
    'คืนค่าส่วนลด LazCoins', 'Reverse - Premium Package', 'ค่าเคลมสินค้าสูญหาย',
}


# ══════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════

def _safe_read_excel(path, **kwargs):
    try:
        return pd.read_excel(path, **kwargs)
    except Exception:
        tmp = tempfile.mktemp(suffix='.xlsx')
        subprocess.run(['zip', '-FF', path, '--out', tmp], capture_output=True, input=b'y\n')
        return pd.read_excel(tmp, **kwargs)


def _safe_load_workbook(path, **kwargs):
    try:
        return load_workbook(path, **kwargs)
    except Exception:
        tmp = tempfile.mktemp(suffix='.xlsx')
        subprocess.run(['zip', '-FF', path, '--out', tmp], capture_output=True, input=b'y\n')
        return load_workbook(tmp, **kwargs)


def clean_order_no(x):
    try:
        s = str(x).strip()
        if s in ('', 'nan', 'None'):
            return ''
        return str(int(float(s)))
    except Exception:
        return str(x).strip()


# ══════════════════════════════════════════════
# ขั้นตอนที่ 1 — build_combined
# (พอร์ตจาก scripts/build_combined.py)
# ══════════════════════════════════════════════

def build_combined(base_dir: str, out_path: str, log=print) -> str:
    """รวมไฟล์ดิบ (Data LAZADA / Data ERP / STOCK / ABBHOL) เป็น input_combined.xlsx
    พร้อมรักษา formatting ของ Data ERP (สีแดง col E/F) ด้วย openpyxl"""

    frames = []
    for shop in SHOPS:
        df = _safe_read_excel(f'{base_dir}/Data LAZADA/{shop}.xlsx')
        th_map = {
            'หมายเลขคำสั่งซื้อ': 'NO',
            'วันที่ทำรายการ': 'Transaction Date',
            'ชื่อรายการธุรกรรม': 'Fee Name',
            'จำนวนเงิน(รวมภาษี)': 'Amount(Include Tax)',
            'SKU ร้านค้า': 'Seller SKU',
            'Order Number': 'NO',
        }
        df = df.rename(columns=th_map)
        df['NO'] = df['NO'].astype(str).str.strip()
        # ตัด .0 ออก เช่น 1234567890.0 -> 1234567890
        df['NO'] = df['NO'].apply(lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else x)
        df['SHOP'] = shop
        rest = [c for c in df.columns if c not in ('NO', 'SHOP')]
        df = df[['NO', 'SHOP'] + rest]
        frames.append(df)
        log(f'{shop}: {len(df)} rows')

    df_laz = pd.concat(frames, ignore_index=True)
    df_stock = _safe_read_excel(f'{base_dir}/STOCK/STOCK.xlsx')

    abbhol_files = sorted(glob.glob(f'{base_dir}/ABBHOL/ABBHOL*.xlsx'))
    abbhol_dfs = {os.path.splitext(os.path.basename(f))[0]: _safe_read_excel(f) for f in abbhol_files}

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_laz.to_excel(writer, sheet_name='Data LAZADA', index=False)
        df_stock.to_excel(writer, sheet_name='STOCK', index=False)
        for name, df in abbhol_dfs.items():
            df.to_excel(writer, sheet_name=name, index=False)
            log(f'{name}: {len(df)} rows')

    wb_out = load_workbook(out_path)
    wb_erp = _safe_load_workbook(f'{base_dir}/Data ERP/Data ERP.xlsx')
    ws_src = wb_erp.active

    if 'Data ERP' in wb_out.sheetnames:
        del wb_out['Data ERP']
    ws_dst = wb_out.create_sheet('Data ERP', 1)

    # คัดลอกทุก cell พร้อม formatting เต็มรูปแบบ (สีแดง col E/F ต้องรอด) ด้วย openpyxl
    # ห้ามใช้ pandas เพราะจะลบสีออก
    red_count = 0
    for row in ws_src.iter_rows():
        for cell in row:
            dst = ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dst.font = copy(cell.font)
                dst.fill = copy(cell.fill)
                dst.border = copy(cell.border)
                dst.alignment = copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.protection = copy(cell.protection)
                f = cell.fill
                if f.fill_type == 'solid':
                    rgb = str(f.fgColor.rgb)
                    if len(rgb) == 8:
                        r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
                        if r > 150 and r > g * 1.5 and r > b * 1.5:
                            red_count += 1

    # คัดลอกความกว้างคอลัมน์ / ความสูงแถว ให้ดูเหมือนต้นฉบับ
    for col_letter, dim in ws_src.column_dimensions.items():
        if dim.width:
            ws_dst.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in ws_src.row_dimensions.items():
        if dim.height:
            ws_dst.row_dimensions[row_idx].height = dim.height

    wb_out.save(out_path)
    log(f'Data ERP: copied with formatting ({red_count} red-highlighted cells preserved)')
    return out_path


# ══════════════════════════════════════════════
# ขั้นตอนที่ 2 — build_analysis
# (พอร์ตจาก scripts/build_analysis.py)
# ══════════════════════════════════════════════

def is_red_fill(cell):
    """ตรวจว่า cell มีไฮไลต์สีแดงหรือเปล่า — ถ้าใช่ห้ามแก้ไข"""
    f = cell.fill
    if f.fill_type != 'solid':
        return False
    rgb = str(f.fgColor.rgb)
    if len(rgb) == 8:
        r = int(rgb[2:4], 16)
        g = int(rgb[4:6], 16)
        b = int(rgb[6:8], 16)
        return r > 150 and r > g * 1.5 and r > b * 1.5
    return False


def is_bundle_header(sku):
    s = str(sku)
    return s.startswith('BP051') or s.startswith('PD-')


def build_analysis(input_path: str, output_path: str, cost_map: dict = None,
                    note_map: dict = None, log=print, box_cost_rate: float = None,
                    packing_cost_rate: float = None, employees: list = None) -> dict:
    """วิเคราะห์ input_combined.xlsx -> ค่าคอมLAZADA.xlsx (4 ชีทรายงาน)

    cost_map: {รหัสสินค้า: ต้นทุน/ชิ้น} จาก Google Sheets CostORG2 (ถ้าไม่ใส่ -> เขียนสูตร IMPORTRANGE แทน)
    note_map: {เลขที่เอกสาร: หมายเหตุ} จาก Google Sheets CSR 2026 (ถ้าไม่ใส่ -> เขียนสูตร IMPORTRANGE แทน)
    box_cost_rate: อัตราต้นทุนค่ากล่อง (0-1) สำหรับชีท "ใบปะหน้า LAZADA" (ไม่ใส่ -> เว้นว่างให้กรอกเองใน Excel)
    packing_cost_rate: อัตราต้นทุนคน PACKING (0-1) เหมือนกัน (ไม่ใส่ -> เว้นว่างให้กรอกเองใน Excel)
    employees: list ของ (ชื่อพนักงาน, % ที่กำหนด 0-1) สูงสุด 3 คน สำหรับตารางค่าคอมพนักงาน (ไม่ใส่ -> เว้นว่างให้กรอกเอง)

    return: dict สรุปยอด (income/expense/net/per-shop) สำหรับแสดงผลในเว็บแอป
    """
    xl = pd.ExcelFile(input_path)
    sheet_names = xl.sheet_names

    df_laz_raw = xl.parse('Data LAZADA', header=0)
    df_erp_raw = xl.parse('Data ERP', header=0)
    df_stock = xl.parse('STOCK', header=0)

    # ── Build ABB map จาก STOCK ──
    stock_abb_to_order = dict(zip(df_stock.iloc[:, 8].astype(str), df_stock.iloc[:, 1].astype(str)))
    stock_order_to_abb = {}
    stock_b_zero_a = set()

    for _, row in df_stock.iterrows():
        b_val = str(row.iloc[1]); a_val = str(row.iloc[0]); abb = str(row.iloc[8])
        if b_val == '0':
            stock_b_zero_a.add(a_val)
        else:
            existing = stock_order_to_abb.get(b_val)
            if existing is None or (existing.startswith('CAIVAHPW') and not abb.startswith('CAIVAHPW')):
                stock_order_to_abb[b_val] = abb

    true_b_zero = stock_b_zero_a - set(stock_order_to_abb.keys())

    # ── คำนวณ ABB สำหรับ Data LAZADA ──
    df_laz_raw['NO'] = df_laz_raw['NO'].map(clean_order_no)
    df_laz_raw['ABB'] = df_laz_raw['NO'].map(lambda x: stock_order_to_abb.get(x))

    # ── ตรวจจับเดือนหลัก + cross-month sheets ──
    abbhol_sheets_all = sorted([s for s in sheet_names if s.startswith('ABBHOL26')])
    erp_abbs = df_erp_raw.iloc[:, 2].dropna().astype(str)
    main_month = None
    if len(erp_abbs):
        main_prefix = erp_abbs.str[:10].value_counts().idxmax()
        main_month = main_prefix[6:10]
    cross_month_sheets = [s for s in abbhol_sheets_all if main_month and main_month not in s]
    log(f"เดือนหลัก: {main_month} | Cross-month: {cross_month_sheets}")

    # ── เพิ่มแถวข้ามเดือนใน Data ERP เฉพาะ ABB ที่มีข้อมูลใน Data LAZADA ──
    erp_f_nos = set(
        v for v in [stock_abb_to_order.get(str(x)) for x in df_erp_raw.iloc[:, 2]]
        if v and v != '0'
    )
    laz_nos = set(df_laz_raw['NO'].astype(str).unique())
    new_erp_rows = []

    for cm_sheet in cross_month_sheets:
        cm_month = cm_sheet[6:10]
        targets_cm = {
            no for no in laz_nos
            if no in stock_order_to_abb and cm_month in stock_order_to_abb[no] and no not in erp_f_nos
        }
        if not targets_cm:
            continue

        df_cm = xl.parse(cm_sheet, header=0)
        first_c = str(df_cm.iloc[0, 2]) if len(df_cm) > 0 else ''
        if not (first_c.startswith('ABBHOL') or first_c.startswith('CAIVAHPW')):
            df_cm = xl.parse(cm_sheet, header=None)

        df_cm['_order_no'] = df_cm.iloc[:, 2].map(stock_abb_to_order)
        matched_abbs = set(df_cm.iloc[:, 2].astype(str)[
            df_cm.iloc[:, 4].astype(str).isin(targets_cm) |
            df_cm['_order_no'].astype(str).isin(targets_cm)
        ].unique())
        rows_cm = df_cm[df_cm.iloc[:, 2].astype(str).isin(matched_abbs)]
        added = 0
        for _, row in rows_cm.iterrows():
            r = [None] * 24
            for i in [0, 1, 2, 3]:
                r[i] = row.iloc[i]
            ov = stock_abb_to_order.get(str(row.iloc[2]))
            if not ov or ov == '0':
                ov = str(row.iloc[4]) if str(row.iloc[4]) not in ('0', 'nan') else None
            r[5] = ov
            for i in range(6, min(23, len(row) - 1)):
                r[i] = row.iloc[i]
            new_erp_rows.append(r)
            added += 1
        log(f"เพิ่มแถวจาก {cm_sheet}: {added} แถว (match Lazada)")

    # ── เตรียม df_laz + df_co สำหรับ analysis ──
    df_laz = df_laz_raw[['SHOP', 'Transaction Date', 'Fee Name', 'Amount(Include Tax)', 'Seller SKU', 'ABB']].copy()
    df_laz = df_laz.rename(columns={'SHOP': 'ร้าน'})
    df_laz['Transaction Date'] = pd.to_datetime(df_laz['Transaction Date'], format='%d %b %Y', errors='coerce')

    mask_nan = df_laz['ABB'].isna()
    if mask_nan.any():
        df_laz.loc[mask_nan, 'ABB'] = 'รอ ABB-' + df_laz.loc[mask_nan, 'ร้าน'].str.replace('LAZADA-', '', regex=False)

    df_co = df_erp_raw.copy().rename(columns={'พนักงานขาย': 'พนง.ขาย'})
    df_co['วันที่เอกสาร'] = pd.to_datetime(df_co['วันที่เอกสาร'], format='%d/%m/%y', errors='coerce').dt.date

    if new_erp_rows:
        extra = []
        for r in new_erp_rows:
            rec = {col: None for col in df_co.columns}
            for ci, val in enumerate(r):
                if ci < len(df_co.columns):
                    rec[list(df_co.columns)[ci]] = None if (val is None or (isinstance(val, float) and np.isnan(val))) else val
            extra.append(rec)
        df_co = pd.concat([df_co, pd.DataFrame(extra)], ignore_index=True)
        try:
            df_co['วันที่เอกสาร'] = pd.to_datetime(df_co['วันที่เอกสาร'], errors='coerce').dt.date
        except Exception:
            pass

    # เพิ่มเลขคำสั่งซื้อใน df_co -> แสดงใน Sheet 1 col E/F เหมือน Data ERP
    df_co = df_co.rename(columns={'Unnamed: 4': 'เลขคำสั่งซื้อ', 'Unnamed: 5': 'NO.'})

    def _get_order_no(abb):
        if pd.isna(abb):
            return None
        v = stock_abb_to_order.get(str(abb))
        if v and v not in ('0', '0.0', 'nan'):
            return clean_order_no(v)
        return None

    df_co['เลขคำสั่งซื้อ'] = df_co['เลขคำสั่งซื้อ'].apply(
        lambda x: clean_order_no(str(x)) if pd.notna(x) and str(x) not in ('nan', 'None', '0', '0.0') else None)
    _mask_empty = df_co['เลขคำสั่งซื้อ'].isna() | (df_co['เลขคำสั่งซื้อ'].astype(str).str.strip() == '')
    df_co.loc[_mask_empty, 'เลขคำสั่งซื้อ'] = df_co.loc[_mask_empty, 'เลขที่เอกสาร'].apply(_get_order_no)
    df_co['NO.'] = df_co['เลขคำสั่งซื้อ']

    # ── Analysis ──
    _PRIORITY_FEE_ORDER = ['ยอดรวมค่าสินค้า', 'ค่าธรรมเนียมการชำระเงิน', 'หักค่าธรรมเนียมการขายสินค้า']
    _all_fee_names = df_laz['Fee Name'].dropna().unique().tolist()
    _priority_present = [f for f in _PRIORITY_FEE_ORDER if f in _all_fee_names]
    _rest_fees = sorted([f for f in _all_fee_names if f not in _PRIORITY_FEE_ORDER])
    FEE_TYPES = _priority_present + _rest_fees
    fee_pivot = df_laz.pivot_table(index='ABB', columns='Fee Name',
        values='Amount(Include Tax)', aggfunc='sum', fill_value=0)
    for ft in FEE_TYPES:
        if ft not in fee_pivot.columns:
            fee_pivot[ft] = 0.0
    laz_shop = df_laz.groupby('ABB')['ร้าน'].first()

    _laz_no_valid = df_laz_raw[df_laz_raw['NO'].notna()].copy()
    _laz_no_valid['NO'] = _laz_no_valid['NO'].astype(str)
    _laz_no_valid = _laz_no_valid[~_laz_no_valid['NO'].isin(['nan', 'None', ''])]
    fee_pivot_no = _laz_no_valid.pivot_table(index='NO', columns='Fee Name',
        values='Amount(Include Tax)', aggfunc='sum', fill_value=0)
    for ft in FEE_TYPES:
        if ft not in fee_pivot_no.columns:
            fee_pivot_no[ft] = 0.0
    laz_shop_no = _laz_no_valid.groupby('NO')['SHOP'].first()

    # Override พนง.ขาย ใน df_co ด้วยชื่อร้านจาก Lazada
    _abb_to_laz_shop = laz_shop.to_dict()
    _no_to_laz_shop = laz_shop_no.to_dict()
    df_co['_laz_shop'] = df_co['เลขคำสั่งซื้อ'].apply(
        lambda no: _no_to_laz_shop.get(clean_order_no(str(no))) if not pd.isna(no) and str(no) not in ('nan', 'None', '') else None)
    _mask_miss = df_co['_laz_shop'].isna() & df_co['เลขที่เอกสาร'].notna()
    df_co.loc[_mask_miss, '_laz_shop'] = df_co.loc[_mask_miss, 'เลขที่เอกสาร'].apply(
        lambda d: _abb_to_laz_shop.get(str(d)) if str(d) not in ('nan', 'None', '') else None)
    _mask = df_co['_laz_shop'].notna()
    df_co.loc[_mask, 'พนง.ขาย'] = df_co.loc[_mask, '_laz_shop']
    df_co = df_co.drop(columns=['_laz_shop'])
    log(f"Override พนง.ขาย จาก Lazada: {_mask.sum()} แถว")

    def assign_fees(group_df, total_arr):
        n = len(group_df)
        if n == 1:
            return [total_arr.copy()]
        hdr = group_df['รหัสสินค้า'].apply(is_bundle_header).values
        nh = hdr.sum(); nc = n - nh
        if nh > 0 and nc > 0:
            return [np.zeros(len(FEE_TYPES)) if h else total_arr / nc for h in hdr]
        return [total_arr / n for _ in range(n)]

    def _unnamed_idx(c):
        try:
            return int(str(c).replace('Unnamed: ', ''))
        except Exception:
            return -1

    drop_unnamed = [c for c in df_co.columns if str(c).startswith('Unnamed:') and _unnamed_idx(c) >= 24]
    if drop_unnamed:
        df_co = df_co.drop(columns=drop_unnamed)

    df_co = df_co.reset_index(drop=True)
    df_co['_doc'] = df_co['เลขที่เอกสาร'].fillna('__NaN__')
    result_records = []
    doc_order, seen = [], set()
    for d in df_co['_doc']:
        if d not in seen:
            doc_order.append(d); seen.add(d)

    base_cols = [c for c in df_co.columns if c != '_doc']
    for doc in doc_order:
        grp = df_co[df_co['_doc'] == doc].copy()
        real_doc = grp['เลขที่เอกสาร'].iloc[0]
        is_real = not pd.isna(real_doc) if not isinstance(real_doc, str) else real_doc != '__NaN__'
        if is_real and str(real_doc).startswith('CAIVAHPW'):
            total_arr = np.zeros(len(FEE_TYPES))
        else:
            _ord_no = clean_order_no(str(grp['เลขคำสั่งซื้อ'].iloc[0])) if 'เลขคำสั่งซื้อ' in grp.columns else ''
            if _ord_no and _ord_no not in ('None', 'nan', '0', '') and _ord_no in fee_pivot_no.index:
                total_arr = fee_pivot_no.loc[_ord_no, FEE_TYPES].values.astype(float)
            elif is_real and real_doc in fee_pivot.index:
                total_arr = fee_pivot.loc[real_doc, FEE_TYPES].values.astype(float)
            else:
                total_arr = np.zeros(len(FEE_TYPES))
        n = len(grp); fee_arrays = assign_fees(grp, total_arr)
        hdr_mask = grp['รหัสสินค้า'].apply(is_bundle_header).values; nh = hdr_mask.sum()
        for idx, (_, row) in enumerate(grp.iterrows()):
            rec = row.drop('_doc').to_dict(); fees = fee_arrays[idx]
            for i, ft in enumerate(FEE_TYPES):
                rec[ft] = round(float(fees[i]), 4)
            rec['เงินเข้าบริษัท'] = round(float(fees.sum()), 4)
            rec['_row_type'] = ('single' if n == 1 else
                                'bundle_header' if nh > 0 and hdr_mask[idx] else
                                'bundle_component' if nh > 0 else 'multi_split')
            result_records.append(rec)

    result_df = pd.DataFrame(result_records)
    output_cols = base_cols + FEE_TYPES + ['เงินเข้าบริษัท']
    result_df['_sort_shop'] = pd.Categorical(result_df['พนง.ขาย'], categories=SHOPS, ordered=True)
    result_df['_sort_date'] = pd.to_datetime(result_df['วันที่เอกสาร'], errors='coerce')
    result_df = result_df.sort_values(['_sort_shop', '_sort_date', 'เลขที่เอกสาร'], na_position='last').reset_index(drop=True)
    row_types_list = result_df['_row_type'].tolist()
    result_df = result_df[output_cols]
    log("Sheet เงินเข้าบริษัทประจำเดือน: %d rows" % len(result_df))

    # ── เขียน Output ──
    tmp_copy = tempfile.mktemp(suffix='.xlsx')
    shutil.copy(input_path, tmp_copy)
    os.chmod(tmp_copy, 0o644)

    wb = load_workbook(tmp_copy)

    # --- อัปเดต Data ERP: ใส่ INDEX/MATCH col E + values col F + แถวข้ามเดือน ---
    erp_ws = wb['Data ERP']
    erp_last_orig = erp_ws.max_row

    erp_ws.cell(row=1, column=5).value = 'เลขคำสั่งซื้อ'
    erp_ws.cell(row=1, column=6).value = 'NO'

    erp_e_values = [stock_abb_to_order.get(str(v), None) for v in df_erp_raw.iloc[:, 2]]
    for i, val in enumerate(erp_e_values):
        row = i + 2
        clean_val = clean_order_no(str(val)) if val and val not in ('0', 'nan', 'None') else None
        if not is_red_fill(erp_ws.cell(row=row, column=5)):
            orig_e = erp_ws.cell(row=row, column=5).value
            orig_clean = clean_order_no(str(orig_e)) if orig_e and str(orig_e) not in ('nan', 'None', '0', '0.0') else None
            erp_ws.cell(row=row, column=5).value = orig_clean if orig_clean else clean_val
        if not is_red_fill(erp_ws.cell(row=row, column=6)):
            erp_ws.cell(row=row, column=6).value = clean_val

    C_CROSS = 'FFD7A0'

    next_row = erp_last_orig + 1
    cross_rows = []
    for r in new_erp_rows:
        for ci, val in enumerate(r):
            if ci == 4:
                if not is_red_fill(erp_ws.cell(row=next_row, column=5)):
                    erp_ws.cell(row=next_row, column=5).value = f'=INDEX(STOCK!B:B,MATCH(C{next_row},STOCK!I:I,0))'
            else:
                cell_val = None if (val is None or (isinstance(val, float) and str(val) == 'nan')) else val
                target_col = ci + 1
                if target_col == 6 and is_red_fill(erp_ws.cell(row=next_row, column=6)):
                    pass
                else:
                    erp_ws.cell(row=next_row, column=target_col).value = cell_val
        cross_rows.append(next_row)
        next_row += 1

    _cross_fill = PatternFill('solid', fgColor=C_CROSS)
    for rn in cross_rows:
        for col in range(1, erp_ws.max_column + 1):
            cell = erp_ws.cell(row=rn, column=col)
            if not is_red_fill(cell):
                cell.fill = _cross_fill

    log(f"Data ERP: {next_row-2} แถว (เพิ่ม {len(new_erp_rows)} แถวข้ามเดือน)")

    # --- Fix SHOP QC ---
    def _clean_no(x):
        try:
            s = str(x).strip()
            if s in ('', 'nan', 'None'):
                return ''
            return str(int(float(s)))
        except Exception:
            return str(x).strip()

    no_to_shop = {}
    for _no, _shop in zip(df_laz_raw['NO'].astype(str), df_laz_raw['SHOP']):
        k = _clean_no(_no)
        if k:
            no_to_shop[k] = str(_shop)

    C_BLUE_FIX = 'BDD7EE'
    _blue_fill_fix = PatternFill('solid', fgColor=C_BLUE_FIX)
    erp_total_rows = next_row - 1
    blue_count = 0
    for _r in range(2, erp_total_rows + 1):
        _order_no = _clean_no(erp_ws.cell(row=_r, column=6).value)
        if not _order_no:
            continue
        _expected = no_to_shop.get(_order_no)
        if not _expected:
            continue
        _shop_cell = erp_ws.cell(row=_r, column=23)
        _current = str(_shop_cell.value or '')
        if _current != _expected:
            _shop_cell.value = _expected
            for _c in range(1, erp_ws.max_column + 1):
                _cell = erp_ws.cell(row=_r, column=_c)
                if not is_red_fill(_cell):
                    _cell.fill = _blue_fill_fix
            blue_count += 1

    log(f"แก้ชื่อร้าน + ไฮไลต์สีฟ้า: {blue_count} แถว")

    # --- อัปเดต Data LAZADA: ใส่ ABB col + INDEX/MATCH col ---
    laz_ws = wb['Data LAZADA']
    laz_last = laz_ws.max_row
    laz_nos_list = df_laz_raw['NO'].astype(str).tolist()
    b_zero_set = {i for i, no in enumerate(laz_nos_list) if no in true_b_zero}

    for row in range(2, laz_last + 1):
        cell = laz_ws.cell(row=row, column=1)
        raw = cell.value
        if raw is not None:
            try:
                cell.value = str(int(float(str(raw).strip())))
            except Exception:
                cell.value = str(raw).strip()
        cell.number_format = '@'

    laz_max_col = laz_ws.max_column
    laz_ws.cell(row=1, column=laz_max_col + 1).value = 'เลขที่เอกสาร ( ABB )'
    laz_ws.cell(row=1, column=laz_max_col + 2).value = 'SHOP QC'

    for i in range(laz_last - 1):
        row = i + 2
        laz_ws.cell(row=row, column=laz_max_col + 1).value = None if i in b_zero_set else f'=IFERROR(VLOOKUP(A{row},STOCK!B:I,8,0),"รอABB")'

    for row in range(2, laz_last + 1):
        laz_ws.cell(row=row, column=laz_max_col + 2).value = f"=INDEX('Data ERP'!W:W,MATCH(A{row},'Data ERP'!F:F,0))"

    # --- เพิ่ม 3 ชีทรายงาน ---
    for sname in ['เงินเข้าบริษัทประจำเดือน', 'สรุปรวม', 'สรุปรายวัน', 'สรุปยอดรวมทั้งหมดที่เข้าบริษัท', 'LAZADA จ่าย']:
        if sname in wb.sheetnames:
            del wb[sname]

    ws1 = wb.create_sheet('เงินเข้าบริษัทประจำเดือน')
    ws2 = wb.create_sheet('สรุปรวม')
    ws3 = wb.create_sheet('สรุปรายวัน')

    C_H_BG = '1F3864'; C_FEE_H = 'F4B183'; C_FEE = 'FCE4D6'
    C_INC_H = '70AD47'; C_INC = 'E2EFDA'
    C_BNDL_H = 'FFD700'; C_BNDL_C = 'FFFACD'
    C_ALT = 'F2F2F2'; C_UNMATCH = 'FFD7D7'
    C_ZERO = 'EFEFEF'; C_TOT = 'D6E4F0'
    thin = Side(style='thin', color='CCCCCC'); med = Side(style='medium', color='888888')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    bdr_med = Border(left=med, right=med, top=med, bottom=med)
    NUM_FMT = '#,##0.00;-#,##0.00;0.00'

    def fill(c):
        return PatternFill('solid', fgColor=c)

    def hfnt(c='FFFFFF', b=True, sz=9):
        return Font(name='Arial', bold=b, color=c, size=sz)

    def dfnt(b=False, c='000000', sz=9):
        return Font(name='Arial', bold=b, color=c, size=sz)

    # ══ Sheet 1: เงินเข้าบริษัทประจำเดือน ══
    headers = list(result_df.columns)
    n_base = len(base_cols); n_fee = len(FEE_TYPES)
    fee_s = n_base + 1; fee_e = n_base + n_fee; inc_c = n_base + n_fee + 1

    for ci, h in enumerate(headers, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = bdr
        if ci <= n_base:
            c.fill = fill(C_H_BG); c.font = hfnt()
        elif ci <= fee_e:
            c.fill = fill(C_FEE_H); c.font = hfnt('000000')
        else:
            c.fill = fill(C_INC_H); c.font = hfnt()

    for ri, (_, row) in enumerate(result_df.iterrows(), 2):
        rtype = row_types_list[ri - 2]
        for ci, col in enumerate(headers, 1):
            val = row[col]
            if isinstance(val, float) and np.isnan(val):
                val = None
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.font = dfnt(b=(ci == inc_c)); cell.border = bdr
            cell.alignment = Alignment(vertical='center')
            if col == 'วันที่เอกสาร':
                cell.number_format = 'D MMM YYYY'
            if ci >= fee_s:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            if rtype == 'bundle_header':
                cell.fill = fill(C_BNDL_H)
            elif rtype == 'bundle_component':
                cell.fill = fill(C_BNDL_C)
            elif rtype == 'unmatched':
                cell.fill = fill(C_UNMATCH)
            else:
                alt = ri % 2 == 0
                if ci <= n_base:
                    if alt:
                        cell.fill = fill(C_ALT)
                elif ci <= fee_e:
                    cell.fill = fill('FDE9D9' if alt else C_FEE)
                else:
                    cell.fill = fill('EBF5E0' if alt else C_INC)

    ws1.freeze_panes = 'A2'; ws1.row_dimensions[1].height = 40
    for ci, h in enumerate(headers, 1):
        w = 16
        if h == 'รายละเอียด':
            w = 35
        elif h in ('เลขที่เอกสาร', 'เลขที่เอกสารอ้างอิง'):
            w = 20
        elif h == 'ชื่อลูกค้า':
            w = 22
        ws1.column_dimensions[get_column_letter(ci)].width = w

    # ══ 6 Extra Columns (ต้นทุน / กำไร / หมายเหตุ) ══
    _ci = headers.index('รหัสสินค้า') + 1;      col_I = get_column_letter(_ci)
    _cm = headers.index('จำนวนที่สั่ง') + 1;     col_M = get_column_letter(_cm)
    _cal = headers.index('ยอดรวมค่าสินค้า') + 1;  col_AL = get_column_letter(_cal)
    _cat = headers.index('เงินเข้าบริษัท') + 1;   col_AT = get_column_letter(_cat)
    n_hdr = len(headers)
    col_AU = get_column_letter(n_hdr + 1)
    col_AV = get_column_letter(n_hdr + 2)
    col_AW = get_column_letter(n_hdr + 3)
    extra_hdrs = [
        ('ต้นทุนสินค้า/ต่อชิ้น', 'FF0000', 'FFFFFF'),
        ('ต้นทุนสินค้ารวม', 'FF0000', 'FFFFFF'),
        ('กำไร(บาท)', '375623', 'FFFFFF'),
        ('กำไร(%)', '375623', 'FFFFFF'),
        ('หมายเหตุ', C_H_BG, 'FFFFFF'),
        ('สาเหตุติดลบ', C_H_BG, 'FFFFFF'),
    ]
    for i, (h, bg, fg) in enumerate(extra_hdrs, n_hdr + 1):
        c = ws1.cell(row=1, column=i, value=h)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = bdr; c.fill = fill(bg); c.font = hfnt(fg)
        ws1.column_dimensions[get_column_letter(i)].width = 18

    use_real_values = cost_map is not None or note_map is not None
    n_data = len(result_df)
    for ri in range(2, n_data + 2):
        product_code = ws1.cell(row=ri, column=_ci).value
        qty = ws1.cell(row=ri, column=_cm).value or 0
        amt = ws1.cell(row=ri, column=_cal).value or 0
        income = ws1.cell(row=ri, column=_cat).value or 0
        doc_no = ws1.cell(row=ri, column=headers.index('เลขที่เอกสาร') + 1).value if 'เลขที่เอกสาร' in headers else None

        if cost_map is not None:
            try:
                cost_per_unit = cost_map.get(str(product_code), 0) if amt and float(amt) > 0 else 0
            except Exception:
                cost_per_unit = 0
            cost_total = (cost_per_unit or 0) * (qty or 0)
            profit_baht = (income or 0) + cost_total
            profit_pct = (profit_baht / amt) if amt and float(amt) > 0 else 0
            note_val = note_map.get(str(doc_no), '') if note_map is not None else ''
            row_vals = [cost_per_unit, cost_total, profit_baht, profit_pct, note_val, None]
        else:
            f_cost = f'=IFERROR(IF(${col_AL}{ri}>0,VLOOKUP({col_I}{ri},IMPORTRANGE("{COST_URL}","CostORG2!$c$3:$n$90000"),12,0),0),"")'
            f_ctot = f'=SUM({col_AU}{ri}*{col_M}{ri})'
            f_profit = f'={col_AT}{ri}+{col_AV}{ri}'
            f_pct = f'=IF({col_AL}{ri}>0,{col_AW}{ri}/{col_AL}{ri},0)'
            note_txt = f'=INDEX(IMPORTRANGE("{NOTE_URL}","2026!$I$2:$Z$29000"),MATCH(C{ri},IMPORTRANGE("{NOTE_URL}","2026!$I$2:$I$29000"),0),18)'
            row_vals = [f_cost, f_ctot, f_profit, f_pct, note_txt, None]

        rtype = row_types_list[ri - 2]
        for j, val in enumerate(row_vals, n_hdr + 1):
            cell = ws1.cell(row=ri, column=j, value=val)
            cell.font = dfnt(); cell.border = bdr
            cell.alignment = Alignment(vertical='center')
            if j <= n_hdr + 4:
                cell.number_format = '0.00%' if j == n_hdr + 4 else NUM_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            if rtype == 'bundle_header':
                cell.fill = fill(C_BNDL_H)
            elif rtype == 'bundle_component':
                cell.fill = fill(C_BNDL_C)
            elif rtype == 'unmatched':
                cell.fill = fill(C_UNMATCH)
            else:
                alt = ri % 2 == 0
                cell.fill = fill('EBF5E0' if alt else C_INC) if j <= n_hdr + 4 else fill(C_ALT if alt else 'FFFFFF')

    # ══ Sheet 2: สรุป Fee ══
    summary = []
    for ft in FEE_TYPES:
        amt = round(result_df[ft].sum(), 2); cnt = int((result_df[ft] != 0).sum())
        cat = 'รายได้' if ft in INCOME_FEES else 'ค่าใช้จ่าย'
        summary.append({'ประเภท Fee': ft, 'หมวดหมู่': cat, 'จำนวนรายการ': cnt, 'ยอดรวม (บาท)': amt, 'หมายเหตุ': ''})

    tot_income = round(sum(r['ยอดรวม (บาท)'] for r in summary if r['หมวดหมู่'] == 'รายได้'), 2)
    tot_expense = round(sum(r['ยอดรวม (บาท)'] for r in summary if r['หมวดหมู่'] == 'ค่าใช้จ่าย'), 2)
    tot_net = round(result_df['เงินเข้าบริษัท'].sum(), 2)

    ws2['A1'] = 'สรุป Fee รายประเภท'
    ws2['A1'].font = Font(name='Arial', bold=True, size=14, color='1F3864')
    ws2['A1'].alignment = Alignment(horizontal='center')
    ws2.merge_cells('A1:E1')
    hd2 = ['ประเภท Fee', 'หมวดหมู่', 'จำนวนรายการ', 'ยอดรวม (บาท)', 'หมายเหตุ']
    for ci, h in enumerate(hd2, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = hfnt(); c.fill = fill(C_H_BG)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = bdr
    for ri, r in enumerate(summary, 3):
        bg = 'DEEAF1' if r['หมวดหมู่'] == 'รายได้' else C_FEE
        for ci, k in enumerate(hd2, 1):
            c = ws2.cell(row=ri, column=ci, value=r[k])
            c.font = dfnt(); c.border = bdr; c.alignment = Alignment(vertical='center'); c.fill = fill(bg)
            if k == 'ยอดรวม (บาท)':
                c.number_format = NUM_FMT; c.alignment = Alignment(horizontal='right', vertical='center')
    blank = 3 + len(summary)
    for i, (lbl, val, bg) in enumerate([('รวมรายได้', tot_income, 'DEEAF1'),
        ('รวมค่าใช้จ่าย', tot_expense, C_FEE), ('เงินเข้าบริษัทสุทธิ', tot_net, C_INC)], blank + 1):
        for ci in range(1, 6):
            c = ws2.cell(row=i, column=ci); c.border = bdr; c.fill = fill(bg); c.font = dfnt(b=True)
        ws2.cell(row=i, column=1, value=lbl).alignment = Alignment(horizontal='right', vertical='center')
        ws2.merge_cells('A%d:C%d' % (i, i))
        vc = ws2.cell(row=i, column=4, value=val)
        vc.number_format = NUM_FMT; vc.alignment = Alignment(horizontal='right', vertical='center')
    shop_start = blank + 6
    ws2.cell(row=shop_start - 1, column=1, value='สรุปแยกตามร้านค้า').font = Font(name='Arial', bold=True, size=11, color='1F3864')
    for ci, h in enumerate(['ร้านค้า', 'เงินเข้าบริษัท (บาท)'], 1):
        c = ws2.cell(row=shop_start, column=ci, value=h)
        c.font = hfnt(); c.fill = fill(C_H_BG)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = bdr
    shop_totals = {}
    for i, shop in enumerate(SHOPS):
        r = shop_start + 1 + i
        val = round(result_df[result_df['พนง.ขาย'] == shop]['เงินเข้าบริษัท'].sum(), 2)
        shop_totals[shop] = val
        ws2.cell(row=r, column=1, value=shop).border = bdr; ws2.cell(row=r, column=1).font = dfnt()
        vc = ws2.cell(row=r, column=2, value=val)
        vc.number_format = NUM_FMT; vc.alignment = Alignment(horizontal='right')
        vc.border = bdr; vc.font = dfnt(b=True)
    ws2.column_dimensions['A'].width = 55; ws2.column_dimensions['B'].width = 20
    ws2.column_dimensions['C'].width = 14; ws2.column_dimensions['D'].width = 18
    ws2.column_dimensions['E'].width = 25; ws2.row_dimensions[2].height = 30
    ws2.freeze_panes = 'A3'
    log("Sheet สรุป Fee: income=%.2f expense=%.2f net=%.2f" % (tot_income, tot_expense, tot_net))

    # ══ Sheet 3: สรุปรายวัน ══
    ALL_COLS = FEE_TYPES + ['เงินเข้าบริษัท']
    pivot_daily = df_laz.pivot_table(index=['ร้าน', 'Transaction Date'], columns='Fee Name',
        values='Amount(Include Tax)', aggfunc='sum', fill_value=0).reset_index()
    for ft in FEE_TYPES:
        if ft not in pivot_daily.columns:
            pivot_daily[ft] = 0.0
    pivot_daily['เงินเข้าบริษัท'] = pivot_daily[FEE_TYPES].sum(axis=1)
    _min = pivot_daily['Transaction Date'].min(); _max = pivot_daily['Transaction Date'].max()
    _full = pd.DataFrame({'Transaction Date': pd.date_range(_min, _max, freq='D')})

    def get_shop_daily(shop):
        sdf = pivot_daily[pivot_daily['ร้าน'] == shop].copy()
        merged = _full.merge(sdf[['Transaction Date'] + ALL_COLS], on='Transaction Date', how='left').fillna(0)
        return merged.sort_values('Transaction Date').reset_index(drop=True)

    current_row = 1; DATE_FMT = 'DD MMM YYYY'
    for shop in SHOPS:
        df_shop = get_shop_daily(shop); n_cols = 1 + len(ALL_COLS)
        ws3.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols)
        hc = ws3.cell(row=current_row, column=1, value=shop)
        hc.font = hfnt(sz=11); hc.fill = fill(C_H_BG)
        hc.alignment = Alignment(horizontal='center', vertical='center'); hc.border = bdr
        ws3.row_dimensions[current_row].height = 22; current_row += 1
        dc = ws3.cell(row=current_row, column=1, value='วันที่')
        dc.font = hfnt(); dc.fill = fill(C_H_BG)
        dc.alignment = Alignment(horizontal='center', vertical='center'); dc.border = bdr
        for ci, col in enumerate(ALL_COLS, 2):
            c = ws3.cell(row=current_row, column=ci, value=col); c.border = bdr
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if col == 'เงินเข้าบริษัท':
                c.fill = fill(C_INC_H); c.font = hfnt()
            else:
                c.fill = fill(C_FEE_H); c.font = hfnt('000000')
        ws3.row_dimensions[current_row].height = 42; current_row += 1
        for ri, (_, row) in enumerate(df_shop.iterrows()):
            is_zero = (row['เงินเข้าบริษัท'] == 0) and all(row[ft] == 0 for ft in FEE_TYPES)
            alt = ri % 2 == 1
            d_cell = ws3.cell(row=current_row, column=1, value=row['Transaction Date'].date())
            d_cell.number_format = DATE_FMT; d_cell.font = dfnt(); d_cell.border = bdr
            d_cell.alignment = Alignment(horizontal='center', vertical='center')
            d_cell.fill = fill(C_ZERO if is_zero else (C_ALT if alt else 'FFFFFF'))
            for ci, col in enumerate(ALL_COLS, 2):
                val = round(float(row[col]), 2)
                cell = ws3.cell(row=current_row, column=ci, value=val)
                cell.number_format = NUM_FMT; cell.border = bdr
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if is_zero:
                    cell.fill = fill(C_ZERO); cell.font = dfnt(c='999999')
                elif col == 'เงินเข้าบริษัท':
                    cell.fill = fill('EBF5E0' if alt else C_INC); cell.font = dfnt(b=True)
                else:
                    cell.fill = fill('FDE9D9' if alt else C_FEE); cell.font = dfnt()
            current_row += 1
        tc = ws3.cell(row=current_row, column=1, value='รวม')
        tc.font = hfnt('000000'); tc.fill = fill(C_TOT)
        tc.border = bdr_med; tc.alignment = Alignment(horizontal='center', vertical='center')
        for ci, col in enumerate(ALL_COLS, 2):
            tv = round(float(df_shop[col].sum()), 2)
            cell = ws3.cell(row=current_row, column=ci, value=tv)
            cell.number_format = NUM_FMT; cell.border = bdr_med
            cell.alignment = Alignment(horizontal='right', vertical='center'); cell.font = dfnt(b=True)
            cell.fill = fill(C_INC_H if col == "เงินเข้าบริษัท" else C_TOT)
        ws3.row_dimensions[current_row].height = 18; current_row += 2

    ws3.column_dimensions['A'].width = 14
    for ci in range(2, 2 + len(ALL_COLS)):
        ws3.column_dimensions[get_column_letter(ci)].width = 13
    ws3.freeze_panes = 'B1'

    # ══════════════════════════════════════════════
    # STEP 9: ชีท "ใบปะหน้า LAZADA" (รายได้/รายจ่ายแยกร้าน + คำนวณกำไร)
    # วางเป็นชีทแรกสุด — ไม่แก้ logic ของ Sheet 1-3 ด้านบนเลย
    # ══════════════════════════════════════════════
    FEE_DEF = {
        'ยอดรวมค่าสินค้า': 'ยอดขาย',
        'Campaign Fee': 'ค่าธรรมเนียมด้านการตลาด',
        'Premium Package': 'ค่าธรรมเนียมด้านการตลาด',
        'Reversal Campaign Fee': 'ค่าธรรมเนียมด้านการตลาด',
        'Reverse - Premium Package': 'ค่าธรรมเนียมด้านการตลาด',
        'คืนค่าธรรมเนียมโปรแกรมส่งฟรีพิเศษกับลาซาด้า': 'ค่าธรรมเนียมด้านการตลาด',
        'คืนค่าธรรมเนียมโปรแกรมส่วนลด LazCoins': 'ค่าธรรมเนียมด้านการตลาด',
        'คืนค่าส่วนลด LazCoins': 'ค่าธรรมเนียมด้านการตลาด',
        'คืนส่วนลดค่าธรรมเนียมการขายสินค้า': 'ค่าธรรมเนียมจากลาซาด้า',
        'ค่าขนส่งการคืนสินค้า': 'ค่าบริการขนส่ง',
        'ค่าธรรมเนียมการชำระเงิน': 'ค่าธรรมเนียมจากลาซาด้า',
        'ค่าธรรมเนียมโปรแกรมส่วนลด LazCoins': 'ค่าธรรมเนียมด้านการตลาด',
        'ค่าเคลมสินค้าสูญหาย': 'เคลม/หักคืนค่าสินค้า',
        'ค่าโปรแกรมส่งฟรีพิเศษกับลาซาด้า': 'ค่าธรรมเนียมด้านการตลาด',
        'ส่วนต่างค่าจัดส่ง': 'ค่าบริการขนส่ง',
        'ส่วนลด LazCoins': 'ค่าธรรมเนียมด้านการตลาด',
        'หักค่าธรรมเนียมการขายสินค้า': 'ค่าธรรมเนียมจากลาซาด้า',
        'หักค่าธรรมเนียมการขายสินค้า (รายการปรับปรุงที่หักขาด/ต่ำเกินไป)': 'ค่าธรรมเนียมจากลาซาด้า',
        'หักค่าธรรมเนียมการชำระเงิน': 'ค่าธรรมเนียมจากลาซาด้า',
        'หักเงินค่าคูปองส่วนลดค่าขนส่งจากลาซาด้า (ลูกค้าคืนสินค้า)': 'ค่าบริการขนส่ง',
        'หักเงินค่าสินค้า (คืนสินค้า)': 'เคลม/หักคืนค่าสินค้า',
    }

    def get_fee_def(fee_name):
        return FEE_DEF.get(fee_name, '')

    SHOP_ORDER_FACE = ['LAZADA-MAFIA', 'LAZADA-UTOPIA', 'LAZADA-TRC', 'LAZADA-FREEROAD', 'LAZADA-WORKFORCE']
    REVENUE_FEE = 'ยอดรวมค่าสินค้า'
    EXPENSE_FEES_FACE = [f for f in FEE_TYPES if f != REVENUE_FEE]

    pivot_face = {shop: {f: 0.0 for f in FEE_TYPES} for shop in SHOP_ORDER_FACE}
    net_by_shop_face = {shop: 0.0 for shop in SHOP_ORDER_FACE}
    for _, row in result_df.iterrows():
        shop = row['พนง.ขาย']
        if shop not in pivot_face:
            continue
        for f in FEE_TYPES:
            v = row[f]
            if v:
                pivot_face[shop][f] += v
        inc = row['เงินเข้าบริษัท']
        if inc:
            net_by_shop_face[shop] += inc

    month_label_face = f"{main_month[2:4]}/20{main_month[:2]}" if main_month else ''

    ws_face = wb.create_sheet('ใบปะหน้า LAZADA', 0)  # วางเป็นชีทแรกสุด

    F_INC = 'DEEAF1'; F_EXP = 'FCE4D6'; F_SUB = 'D6E4F0'; F_NET = 'E2EFDA'
    F_COST = 'FFE2E2'; F_GROSS = 'FFF2CC'; F_NETP = 'E2EFDA'; F_PARAM = 'DDEBF7'

    n_shop_cols = len(SHOP_ORDER_FACE)
    last_col_face = 3 + n_shop_cols + 1  # A=หมวด, B=รายการ, C=คำนิยาม, D..H=ร้าน, I=รวม

    ws_face.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col_face)
    title_cell = ws_face.cell(row=1, column=1, value=f'สรุปยอดขาย LAZADA เดือน {month_label_face}')
    title_cell.font = Font(name='Arial', bold=True, size=14, color='1F3864')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_face.row_dimensions[1].height = 24

    hdr_row = 2
    ws_face.cell(row=hdr_row, column=1, value='หมวด')
    ws_face.cell(row=hdr_row, column=2, value='รายการ')
    ws_face.cell(row=hdr_row, column=3, value='คำนิยาม')
    for i, shop in enumerate(SHOP_ORDER_FACE):
        ws_face.cell(row=hdr_row, column=4 + i, value=shop)
    ws_face.cell(row=hdr_row, column=last_col_face, value='รวม')
    for ci in range(1, last_col_face + 1):
        c = ws_face.cell(row=hdr_row, column=ci)
        c.font = hfnt(); c.fill = fill(C_H_BG); c.border = bdr
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_face.row_dimensions[hdr_row].height = 26

    r = hdr_row + 1

    # ── รายได้ (ยอดขาย เท่านั้น) ──
    rev_start = r
    c = ws_face.cell(row=r, column=2, value='ยอดขาย')
    c.font = dfnt(); c.border = bdr; c.alignment = Alignment(vertical='center')
    def_c = ws_face.cell(row=r, column=3, value='ยอดขาย')
    def_c.font = dfnt(); def_c.border = bdr; def_c.alignment = Alignment(vertical='center'); def_c.fill = fill(F_INC)
    row_total = 0.0
    for i, shop in enumerate(SHOP_ORDER_FACE):
        val = round(pivot_face[shop][REVENUE_FEE], 2)
        row_total += val
        cell = ws_face.cell(row=r, column=4 + i, value=val)
        cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt()
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(F_INC)
    tot_cell = ws_face.cell(row=r, column=last_col_face, value=round(row_total, 2))
    tot_cell.number_format = NUM_FMT; tot_cell.border = bdr; tot_cell.font = dfnt(b=True)
    tot_cell.alignment = Alignment(horizontal='right', vertical='center'); tot_cell.fill = fill(F_INC)
    ws_face.cell(row=r, column=1).fill = fill(F_INC); ws_face.cell(row=r, column=1).border = bdr
    rev_row_face = r
    r += 1
    rev_end = r - 1
    ws_face.merge_cells(start_row=rev_start, start_column=1, end_row=rev_end, end_column=1)
    cat_c = ws_face.cell(row=rev_start, column=1, value='รายได้')
    cat_c.font = hfnt('1F3864'); cat_c.alignment = Alignment(horizontal='center', vertical='center')

    # ── รายจ่าย (ทุก Fee ที่เหลือ) ──
    exp_start = r
    exp_totals_face = {shop: 0.0 for shop in SHOP_ORDER_FACE}
    for f in EXPENSE_FEES_FACE:
        c = ws_face.cell(row=r, column=2, value=f)
        c.font = dfnt(); c.border = bdr; c.alignment = Alignment(vertical='center')
        def_c = ws_face.cell(row=r, column=3, value=get_fee_def(f))
        def_c.font = dfnt(); def_c.border = bdr; def_c.alignment = Alignment(vertical='center'); def_c.fill = fill(F_EXP)
        row_total = 0.0
        for i, shop in enumerate(SHOP_ORDER_FACE):
            val = round(pivot_face[shop][f], 2)
            row_total += val
            exp_totals_face[shop] += val
            cell = ws_face.cell(row=r, column=4 + i, value=val)
            cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt()
            cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(F_EXP)
        tot_cell = ws_face.cell(row=r, column=last_col_face, value=round(row_total, 2))
        tot_cell.number_format = NUM_FMT; tot_cell.border = bdr; tot_cell.font = dfnt(b=True)
        tot_cell.alignment = Alignment(horizontal='right', vertical='center'); tot_cell.fill = fill(F_EXP)
        ws_face.cell(row=r, column=1).fill = fill(F_EXP); ws_face.cell(row=r, column=1).border = bdr
        r += 1
    exp_end = r - 1
    ws_face.merge_cells(start_row=exp_start, start_column=1, end_row=exp_end, end_column=1)
    cat_c = ws_face.cell(row=exp_start, column=1, value='รายจ่าย')
    cat_c.font = hfnt('FFFFFF', True); cat_c.fill = fill('C0504D')
    cat_c.alignment = Alignment(horizontal='center', vertical='center')

    # ── รวมรายจ่าย ──
    c = ws_face.cell(row=r, column=2, value='รวมรายจ่าย')
    c.font = dfnt(b=True); c.border = bdr_med
    exp_grand_total = 0.0
    for i, shop in enumerate(SHOP_ORDER_FACE):
        val = round(exp_totals_face[shop], 2)
        exp_grand_total += val
        cell = ws_face.cell(row=r, column=4 + i, value=val)
        cell.number_format = NUM_FMT; cell.border = bdr_med; cell.font = dfnt(b=True)
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(F_SUB)
    tot_cell = ws_face.cell(row=r, column=last_col_face, value=round(exp_grand_total, 2))
    tot_cell.number_format = NUM_FMT; tot_cell.border = bdr_med; tot_cell.font = dfnt(b=True)
    tot_cell.alignment = Alignment(horizontal='right', vertical='center'); tot_cell.fill = fill(F_SUB)
    for ci in (1, 2, 3):
        ws_face.cell(row=r, column=ci).fill = fill(F_SUB); ws_face.cell(row=r, column=ci).border = bdr_med
    r += 1

    # ── เงินเข้าบริษัทสุทธิ ──
    c = ws_face.cell(row=r, column=2, value='เงินเข้าบริษัทสุทธิ')
    c.font = dfnt(b=True); c.border = bdr_med
    net_grand = 0.0
    for i, shop in enumerate(SHOP_ORDER_FACE):
        val = round(net_by_shop_face[shop], 2)
        net_grand += val
        cell = ws_face.cell(row=r, column=4 + i, value=val)
        cell.number_format = NUM_FMT; cell.border = bdr_med; cell.font = dfnt(b=True)
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(F_NET)
    tot_cell = ws_face.cell(row=r, column=last_col_face, value=round(net_grand, 2))
    tot_cell.number_format = NUM_FMT; tot_cell.border = bdr_med; tot_cell.font = dfnt(b=True)
    tot_cell.alignment = Alignment(horizontal='right', vertical='center'); tot_cell.fill = fill(F_NET)
    for ci in (1, 2, 3):
        ws_face.cell(row=r, column=ci).fill = fill(F_NET); ws_face.cell(row=r, column=ci).border = bdr_med
    net_row_face = r
    r += 1

    # ══ พารามิเตอร์อัตรา % (ไม่ใส่ -> เว้นว่างให้ผู้ใช้กรอกเองในไฟล์ Excel) ══
    ws_face.cell(row=2, column=10, value='พารามิเตอร์ (กรอกเอง)').font = hfnt('1F3864', True)
    ws_face.cell(row=2, column=10).fill = fill(F_PARAM)
    ws_face.cell(row=2, column=10).alignment = Alignment(horizontal='center', vertical='center')
    ws_face.merge_cells(start_row=2, start_column=10, end_row=2, end_column=11)
    box_label = ws_face.cell(row=3, column=10, value='อัตราต้นทุนค่ากล่อง (%)')
    box_label.font = dfnt(); box_label.border = bdr; box_label.fill = fill(F_PARAM)
    box_rate_cell = ws_face.cell(row=3, column=11)
    box_rate_cell.number_format = '0.00%'
    box_rate_cell.border = bdr; box_rate_cell.fill = fill('FFFFFF'); box_rate_cell.font = dfnt(b=True)
    if box_cost_rate is not None:
        box_rate_cell.value = box_cost_rate
    pack_label = ws_face.cell(row=4, column=10, value='อัตราต้นทุนคน PACKING (%)')
    pack_label.font = dfnt(); pack_label.border = bdr; pack_label.fill = fill(F_PARAM)
    pack_rate_cell = ws_face.cell(row=4, column=11)
    pack_rate_cell.number_format = '0.00%'
    pack_rate_cell.border = bdr; pack_rate_cell.fill = fill('FFFFFF'); pack_rate_cell.font = dfnt(b=True)
    if packing_cost_rate is not None:
        pack_rate_cell.value = packing_cost_rate
    ws_face.column_dimensions['J'].width = 26
    ws_face.column_dimensions['K'].width = 12
    BOX_REF = '$K$3'; PACK_REF = '$K$4'

    shop_col_letter_face = get_column_letter(headers.index('พนง.ขาย') + 1)
    cost_col_letter_face = col_AV
    SHEET1_REF = "'เงินเข้าบริษัทประจำเดือน'"
    shop_cols_face = [get_column_letter(4 + i) for i in range(n_shop_cols)]
    TOTAL_COL_FACE = get_column_letter(4 + n_shop_cols)

    def write_row_face(row_idx, label, cell_fill_color, per_shop_fn, total_fn, num_fmt=NUM_FMT, bold=False, border=bdr):
        c = ws_face.cell(row=row_idx, column=2, value=label)
        c.font = dfnt(b=bold); c.border = border; c.alignment = Alignment(vertical='center'); c.fill = fill(cell_fill_color)
        for ci in (1, 2, 3):
            ws_face.cell(row=row_idx, column=ci).fill = fill(cell_fill_color)
            ws_face.cell(row=row_idx, column=ci).border = border
        for i, col in enumerate(shop_cols_face):
            cell = ws_face.cell(row=row_idx, column=4 + i, value=per_shop_fn(col, i))
            cell.number_format = num_fmt; cell.border = border; cell.font = dfnt(b=bold)
            cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(cell_fill_color)
        tc = ws_face.cell(row=row_idx, column=last_col_face, value=total_fn())
        tc.number_format = num_fmt; tc.border = border; tc.font = dfnt(b=True)
        tc.alignment = Alignment(horizontal='right', vertical='center'); tc.fill = fill(cell_fill_color)

    r = net_row_face + 2

    # 3. ต้นทุนค่าสินค้า = SUM(ต้นทุนสินค้ารวม) ต่อร้าน
    # ⚠️ ห้ามใส่เครื่องหมายลบเพิ่ม เพราะค่าต้นทุนจาก CostORG2 ติดลบอยู่แล้ว
    r_cost = r
    write_row_face(r_cost, 'ต้นทุนค่าสินค้า', F_COST,
        lambda col, i: f'=SUMIF({SHEET1_REF}!${shop_col_letter_face}:${shop_col_letter_face},"{SHOP_ORDER_FACE[i]}",{SHEET1_REF}!${cost_col_letter_face}:${cost_col_letter_face})',
        lambda: f'=SUM(D{r_cost}:H{r_cost})')
    r += 1

    # 4. ปรับกำไรน็อตล้อ (เว้นว่างไว้ให้กรอกเอง — ห้ามใส่ default)
    r_adj = r
    c = ws_face.cell(row=r_adj, column=2, value='ปรับกำไรน็อตล้อ')
    c.font = dfnt(); c.border = bdr; c.alignment = Alignment(vertical='center'); c.fill = fill('FFFFFF')
    for ci in (1, 2, 3):
        ws_face.cell(row=r_adj, column=ci).fill = fill('FFFFFF'); ws_face.cell(row=r_adj, column=ci).border = bdr
    for i, col in enumerate(shop_cols_face):
        cell = ws_face.cell(row=r_adj, column=4 + i)  # เว้นว่าง ไม่ใส่ค่า
        cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt()
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill('FFFFFF')
    tc = ws_face.cell(row=r_adj, column=last_col_face, value=f'=SUM(D{r_adj}:H{r_adj})')
    tc.number_format = NUM_FMT; tc.border = bdr; tc.font = dfnt(b=True)
    tc.alignment = Alignment(horizontal='right', vertical='center'); tc.fill = fill('FFFFFF')
    r += 1

    # 5. กำไรเบื้องต้น (บาท)
    r_gross = r
    write_row_face(r_gross, 'กำไรเบื้องต้น (บาท)', F_GROSS,
        lambda col, i: f'={col}{net_row_face}+{col}{r_cost}+{col}{r_adj}',
        lambda: f'={TOTAL_COL_FACE}{net_row_face}+{TOTAL_COL_FACE}{r_cost}+{TOTAL_COL_FACE}{r_adj}', bold=True, border=bdr_med)
    r += 1

    # 6. ต้นทุนค่ากล่อง = กำไรเบื้องต้น * อัตรา%
    r_box = r
    write_row_face(r_box, 'ต้นทุนค่ากล่อง', F_COST,
        lambda col, i: f'={col}{r_gross}*{BOX_REF}',
        lambda: f'=SUM(D{r_box}:H{r_box})')
    r += 1

    # 7. ต้นทุนคน PACKING = กำไรเบื้องต้น * อัตรา%
    r_pack = r
    write_row_face(r_pack, 'ต้นทุนคน PACKING', F_COST,
        lambda col, i: f'={col}{r_gross}*{PACK_REF}',
        lambda: f'=SUM(D{r_pack}:H{r_pack})')
    r += 1

    # กำไรสุทธิ
    r_netprofit = r
    write_row_face(r_netprofit, 'กำไรสุทธิ', F_NETP,
        lambda col, i: f'={col}{r_gross}+{col}{r_box}+{col}{r_pack}',
        lambda: f'={TOTAL_COL_FACE}{r_gross}+{TOTAL_COL_FACE}{r_box}+{TOTAL_COL_FACE}{r_pack}', bold=True, border=bdr_med)
    r += 1

    # กำไร (%)
    r_pct = r
    write_row_face(r_pct, 'กำไร (%)', F_NETP,
        lambda col, i: f'={col}{r_netprofit}/{col}{rev_row_face}',
        lambda: f'={TOTAL_COL_FACE}{r_netprofit}/{TOTAL_COL_FACE}{rev_row_face}', num_fmt='0.00%', bold=True, border=bdr_med)

    ws_face.column_dimensions['A'].width = 12
    ws_face.column_dimensions['B'].width = 42
    ws_face.column_dimensions['C'].width = 24
    for i in range(n_shop_cols):
        ws_face.column_dimensions[get_column_letter(4 + i)].width = 16
    ws_face.column_dimensions[get_column_letter(last_col_face)].width = 16

    if use_real_values:
        log(f"Sheet ใบปะหน้า LAZADA: เดือน {month_label_face} | เงินเข้าบริษัทสุทธิรวม {net_grand:,.2f} (ต้นทุนใช้ค่าจริงจาก Google Sheets เปิดใน Excel ได้ปกติ)")
    else:
        log(f"Sheet ใบปะหน้า LAZADA: เดือน {month_label_face} | เงินเข้าบริษัทสุทธิรวม {net_grand:,.2f}")
        log("⚠️ แถวต้นทุนค่าสินค้า/กำไรเบื้องต้น/ต้นทุนค่ากล่อง/ต้นทุนคนPACKING/กำไรสุทธิ/กำไร% อ้างอิงคอลัมน์ที่ยังเป็นสูตร IMPORTRANGE")
        log("   คำนวณค่าได้จริงเฉพาะเปิดใน Google Sheets เท่านั้น | K3/K4 เว้นว่างให้ผู้ใช้กรอกอัตรา % เอง")

    # ══════════════════════════════════════════════
    # STEP 10: ตารางค่าคอมพนักงาน ต่อท้ายชีท "ใบปะหน้า LAZADA"
    # (ใต้แถว กำไร (%) เว้น 2 แถวว่าง) — ไม่แก้ STEP 9 ด้านบนเลย
    # ══════════════════════════════════════════════
    COMM_COL_LOSS = 'I'; COMM_COL_DAMAGE = 'J'; COMM_COL_NET = 'K'
    C_COMM_HDR = 'F4B183'
    C_COMM_ROW = 'FFFFFF'
    C_COMM_TOT = 'D6E4F0'

    r = r_pct + 3  # เว้น 2 แถวว่าง

    # ── Header ──
    comm_hdr_row = r
    ws_face.merge_cells(start_row=comm_hdr_row, start_column=1, end_row=comm_hdr_row, end_column=3)
    hc = ws_face.cell(row=comm_hdr_row, column=1, value='ร้านค้า')
    hc.font = hfnt('000000'); hc.fill = fill(C_COMM_HDR); hc.border = bdr
    hc.alignment = Alignment(horizontal='center', vertical='center')
    for ci in (2, 3):
        ws_face.cell(row=comm_hdr_row, column=ci).fill = fill(C_COMM_HDR)
        ws_face.cell(row=comm_hdr_row, column=ci).border = bdr
    for i, shop in enumerate(SHOP_ORDER_FACE):
        c = ws_face.cell(row=comm_hdr_row, column=4 + i, value=shop)
        c.font = hfnt('000000'); c.fill = fill(C_COMM_HDR); c.border = bdr
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, label in ((COMM_COL_LOSS, 'หัก\nขาดทุน'), (COMM_COL_DAMAGE, 'หัก\nสินค้าเสียหาย'), (COMM_COL_NET, 'ค่าคอม\nสุทธิ')):
        c = ws_face[f'{col}{comm_hdr_row}']
        c.value = label; c.font = hfnt('000000'); c.fill = fill(C_COMM_HDR); c.border = bdr
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_face.row_dimensions[comm_hdr_row].height = 30
    r += 1

    # ── แถวฐาน: กำไร(ยอดเบิกค่าคอม) = อ้างอิง กำไรสุทธิ ──
    r_comm_base = r
    c = ws_face.cell(row=r_comm_base, column=2, value='กำไร(ยอดเบิกค่าคอม)')
    c.font = dfnt(b=True); c.border = bdr; c.alignment = Alignment(vertical='center'); c.fill = fill(C_COMM_ROW)
    for ci in (1, 2, 3):
        ws_face.cell(row=r_comm_base, column=ci).fill = fill(C_COMM_ROW)
        ws_face.cell(row=r_comm_base, column=ci).border = bdr
    for i, col in enumerate(shop_cols_face):
        cell = ws_face.cell(row=r_comm_base, column=4 + i, value=f'={col}{r_netprofit}')
        cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt(b=True)
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(C_COMM_ROW)
    r += 1

    # ── 3 แถวพนักงาน (ไม่ใส่ -> เว้นชื่อ + % ให้กรอกเอง) ──
    comm_emp_rows = []
    for _emp_i in range(3):
        r_emp = r
        comm_emp_rows.append(r_emp)
        pc = ws_face.cell(row=r_emp, column=1)  # % (คอลัมน์ A)
        pc.number_format = '0.00%'; pc.border = bdr; pc.font = dfnt(); pc.fill = fill(C_COMM_ROW)
        pc.alignment = Alignment(horizontal='center', vertical='center')
        nc = ws_face.cell(row=r_emp, column=2)  # ชื่อพนักงาน (คอลัมน์ B)
        nc.border = bdr; nc.font = dfnt(); nc.fill = fill(C_COMM_ROW)
        nc.alignment = Alignment(vertical='center')
        cc = ws_face.cell(row=r_emp, column=3)  # คอลัมน์ C (ว่าง — ให้สีตรงกัน)
        cc.border = bdr; cc.fill = fill(C_COMM_ROW)
        if employees and _emp_i < len(employees):
            emp_name, emp_pct = employees[_emp_i]
            if emp_name:
                nc.value = emp_name
            if emp_pct is not None:
                pc.value = emp_pct
        for i, col in enumerate(shop_cols_face):
            cell = ws_face.cell(row=r_emp, column=4 + i, value=f'={col}{r_comm_base}*$A{r_emp}')
            cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt()
            cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(C_COMM_ROW)
        for col in (COMM_COL_LOSS, COMM_COL_DAMAGE):  # หัก ขาดทุน/สินค้าเสียหาย - เว้นว่างให้กรอกเอง
            cell = ws_face[f'{col}{r_emp}']
            cell.number_format = NUM_FMT; cell.border = bdr; cell.font = dfnt()
            cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(C_COMM_ROW)
        net_cell = ws_face[f'{COMM_COL_NET}{r_emp}']  # ค่าคอมสุทธิ = SUM(ร้าน) + หักขาดทุน + หักสินค้าเสียหาย
        net_cell.value = f'=SUM(D{r_emp}:H{r_emp})+{COMM_COL_LOSS}{r_emp}+{COMM_COL_DAMAGE}{r_emp}'
        net_cell.number_format = NUM_FMT; net_cell.border = bdr; net_cell.font = dfnt(b=True)
        net_cell.alignment = Alignment(horizontal='right', vertical='center'); net_cell.fill = fill(C_COMM_ROW)
        r += 1

    # ── แถวรวม ──
    r_comm_tot = r
    c = ws_face.cell(row=r_comm_tot, column=2, value='รวม')
    c.font = dfnt(b=True); c.border = bdr_med; c.alignment = Alignment(horizontal='center', vertical='center'); c.fill = fill(C_COMM_TOT)
    for ci in (1, 2, 3):
        ws_face.cell(row=r_comm_tot, column=ci).fill = fill(C_COMM_TOT)
        ws_face.cell(row=r_comm_tot, column=ci).border = bdr_med
    for i, col in enumerate(shop_cols_face):
        cell = ws_face.cell(row=r_comm_tot, column=4 + i, value=f'=SUM({col}{comm_emp_rows[0]}:{col}{comm_emp_rows[-1]})')
        cell.number_format = NUM_FMT; cell.border = bdr_med; cell.font = dfnt(b=True)
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(C_COMM_TOT)
    for col in (COMM_COL_LOSS, COMM_COL_DAMAGE, COMM_COL_NET):
        cell = ws_face[f'{col}{r_comm_tot}']
        cell.value = f'=SUM({col}{comm_emp_rows[0]}:{col}{comm_emp_rows[-1]})'
        cell.number_format = NUM_FMT; cell.border = bdr_med; cell.font = dfnt(b=True)
        cell.alignment = Alignment(horizontal='right', vertical='center'); cell.fill = fill(C_COMM_TOT)

    ws_face.column_dimensions['I'].width = 12
    ws_face.column_dimensions['J'].width = 14
    ws_face.column_dimensions['K'].width = 14

    tmp_out = tempfile.mktemp(suffix='.xlsx')
    wb.save(tmp_out)
    shutil.copy(tmp_out, output_path)
    os.remove(tmp_out)
    os.remove(tmp_copy)
    log("Saved: %s" % output_path)

    return {
        'income': tot_income,
        'expense': tot_expense,
        'net': tot_net,
        'shop_totals': shop_totals,
        'rows': n_data,
        'used_real_cost_note': use_real_values,
        'cover_net': net_grand,
        'cover_month': month_label_face,
    }
