# -*- coding: utf-8 -*-
"""
Lazada SKU Analysis — เว็บแอป (Streamlit)
พอร์ตมาจาก Claude Skill "lazada-sku-analysis" ให้รันได้เองบนเว็บ ไม่ต้องพึ่ง Claude

วิธีรันในเครื่อง:
    pip install -r requirements.txt
    streamlit run streamlit_app.py

Deploy ขึ้น Streamlit Community Cloud: ดู README.md
"""
import os
import io
import shutil
import tempfile
import traceback

import streamlit as st

from lazada_core import build_combined, build_analysis, SHOPS

st.set_page_config(page_title="Lazada SKU Analysis", page_icon="📊", layout="wide")

st.title("📊 Lazada SKU Analysis — ค่าคอมLAZADA")
st.caption("อัปโหลดไฟล์ดิบ → ระบบประมวลผลให้อัตโนมัติ → ดาวน์โหลด ค่าคอมLAZADA.xlsx "
           "(ผลลัพธ์เหมือนตอนใช้ Skill lazada-sku-analysis ทุกประการ)")

SHOP_LABELS = {
    'LAZADA-MAFIA': 'MAFIA',
    'LAZADA-TRC': 'TRC',
    'LAZADA-UTOPIA': 'UTOPIA',
    'LAZADA-FREEROAD': 'FREEROAD',
    'LAZADA-WORKFORCE': 'WORKFORCE',
}


def has_gcp_secret():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def has_gdrive_oauth():
    """เช็คว่ามีการตั้งค่า OAuth ผู้ใช้จริงสำหรับอัปโหลดขึ้น Drive ไว้หรือยัง
    (ใช้แทน Service Account ตอนอัปโหลด เพราะ Service Account ไม่มี storage quota เป็นของตัวเอง)"""
    try:
        return "gdrive_oauth" in st.secrets and all(
            k in st.secrets["gdrive_oauth"] for k in ("client_id", "client_secret", "refresh_token")
        )
    except Exception:
        return False


def build_folder_structure(tmpdir, shop_files, erp_file, stock_file, abbhol_files):
    """เซฟไฟล์ที่อัปโหลดลง folder structure ให้ตรงกับที่ build_combined ต้องการ"""
    os.makedirs(f"{tmpdir}/Data LAZADA", exist_ok=True)
    os.makedirs(f"{tmpdir}/Data ERP", exist_ok=True)
    os.makedirs(f"{tmpdir}/STOCK", exist_ok=True)
    os.makedirs(f"{tmpdir}/ABBHOL", exist_ok=True)

    for shop, f in shop_files.items():
        with open(f"{tmpdir}/Data LAZADA/{shop}.xlsx", "wb") as out:
            out.write(f.getbuffer())

    with open(f"{tmpdir}/Data ERP/Data ERP.xlsx", "wb") as out:
        out.write(erp_file.getbuffer())

    if stock_file is not None:
        with open(f"{tmpdir}/STOCK/STOCK.xlsx", "wb") as out:
            out.write(stock_file.getbuffer())

    for f in abbhol_files or []:
        with open(f"{tmpdir}/ABBHOL/{f.name}", "wb") as out:
            out.write(f.getbuffer())


# ══════════════════════════════════════════════
# 1) เลือกวิธีดึง STOCK
# ══════════════════════════════════════════════
st.subheader("1️⃣ ข้อมูล STOCK")
stock_mode = st.radio(
    "เลือกวิธีดึง STOCK.xlsx",
    ["อัปโหลดไฟล์เอง", "ดึงอัตโนมัติจาก Google Drive (ต้องตั้งค่า Service Account)"],
    horizontal=True,
)

stock_file = None
use_drive_stock = False
if stock_mode == "อัปโหลดไฟล์เอง":
    stock_file = st.file_uploader("อัปโหลด STOCK.xlsx", type=["xlsx"], key="stock")
else:
    use_drive_stock = True
    if not has_gcp_secret():
        st.warning(
            "ยังไม่ได้ตั้งค่า Google Service Account (st.secrets['gcp_service_account']) "
            "— ดูวิธีตั้งค่าใน README.md แล้วลองใหม่ หรือสลับไปอัปโหลดไฟล์เองก่อนได้"
        )
    else:
        st.success("พร้อมดึง STOCK จาก Google Drive อัตโนมัติ ✅")

# ══════════════════════════════════════════════
# 2) อัปโหลดไฟล์ Lazada 5 ร้าน + Data ERP + ABBHOL
# ══════════════════════════════════════════════
st.subheader("2️⃣ ไฟล์ Lazada รายร้าน (5 ไฟล์)")
st.caption(
    "โยนไฟล์ทั้ง 5 ร้านพร้อมกันได้เลย (เหมือนช่อง ABBHOL ด้านล่าง) ระบบจะเดาร้านให้อัตโนมัติจากชื่อไฟล์ "
    "ถ้าเดาผิดหรือเดาไม่ได้ ให้เลือกร้านที่ถูกต้องในดรอปดาวน์ของแต่ละไฟล์เอง"
)
uploaded_shop_files = st.file_uploader(
    "อัปโหลดไฟล์ Lazada 5 ร้าน (เลือกได้หลายไฟล์พร้อมกัน)",
    type=["xlsx"], accept_multiple_files=True, key="shop_files_multi",
)

label_to_shop = {v: k for k, v in SHOP_LABELS.items()}
shop_options = ["-- เลือกร้าน --"] + list(SHOP_LABELS.values())


def _guess_shop_label(filename):
    """เดาว่าไฟล์นี้เป็นของร้านไหน จากการหาชื่อร้าน (MAFIA/TRC/UTOPIA/FREEROAD/WORKFORCE) ในชื่อไฟล์"""
    up = filename.upper()
    for label in SHOP_LABELS.values():
        if label in up:
            return label
    return None


shop_files = {}
if uploaded_shop_files:
    if len(uploaded_shop_files) > 5:
        st.warning(f"อัปโหลดมา {len(uploaded_shop_files)} ไฟล์ แต่ต้องการแค่ 5 ไฟล์ (5 ร้าน) — เอาไฟล์ส่วนเกินออกก่อน")

    assigned_labels = {}  # shop label -> file (เอาไว้เช็คว่าเลือกร้านซ้ำกันไหม)
    for i, f in enumerate(uploaded_shop_files):
        guess = _guess_shop_label(f.name)
        default_idx = shop_options.index(guess) if guess in shop_options else 0
        choice = st.selectbox(f"📄 {f.name}", shop_options, index=default_idx, key=f"shopmatch_{i}")
        if choice != "-- เลือกร้าน --":
            if choice in assigned_labels and assigned_labels[choice] is not f:
                st.error(f"ร้าน {choice} ถูกเลือกซ้ำ ({assigned_labels[choice].name} กับ {f.name}) — แก้ให้เหลือไฟล์เดียวต่อร้าน")
            else:
                assigned_labels[choice] = f
                shop_files[label_to_shop[choice]] = f

    if len(shop_files) == 5:
        st.success("จับคู่ร้านครบทั้ง 5 ร้านแล้ว ✅")

st.subheader("3️⃣ ไฟล์ Data ERP")
erp_file = st.file_uploader("อัปโหลด Data ERP.xlsx", type=["xlsx"], key="erp")

st.subheader("4️⃣ ไฟล์ ABBHOL (ข้ามเดือน — ไม่บังคับ)")
abbhol_files = st.file_uploader(
    "อัปโหลด ABBHOL เดือนก่อน/ถัดไป เช่น ABBHOL2604.xlsx, ABBHOL2606.xlsx (เลือกได้หลายไฟล์)",
    type=["xlsx"], accept_multiple_files=True, key="abbhol",
)

# ══════════════════════════════════════════════
# 3) คอลัมน์ต้นทุน/กำไร/หมายเหตุ
# ══════════════════════════════════════════════
st.subheader("5️⃣ คอลัมน์ ต้นทุนสินค้า / กำไร / หมายเหตุ")
st.caption(
    "3 คอลัมน์นี้เขียนเป็นสูตร IMPORTRANGE ดึงจาก Google Sheets อื่น เหมือน Skill lazada-sku-analysis เดิมทุกครั้ง "
    "(ต้องเปิดไฟล์ผลลัพธ์ใน Google Sheets ถึงจะเห็นค่าคำนวณจริง เปิดใน Excel จะเห็นสูตรแต่ยังไม่มีค่า — "
    "ถ้าเปิดใช้ \"6️⃣ บันทึกขึ้น Google Drive\" ด้านล่าง เว็บจะแปลงไฟล์เป็น Google Sheets ให้อัตโนมัติ สูตรจะคำนวณได้เลย)"
)
compute_real = False

# ══════════════════════════════════════════════
# 4) บันทึกขึ้น Google Drive อัตโนมัติ
# ══════════════════════════════════════════════
st.subheader("6️⃣ บันทึกไฟล์ผลลัพธ์ขึ้น Google Drive (ไม่บังคับ)")
save_to_drive = False
drive_folder_input = ""
if has_gdrive_oauth():
    default_folder = ""
    try:
        default_folder = st.secrets.get("drive_folder_id", "")
    except Exception:
        default_folder = ""
    save_to_drive = st.checkbox(
        "บันทึกขึ้น Google Drive อัตโนมัติหลังประมวลผลเสร็จ (แปลงเป็น Google Sheets ให้เลย)"
    )
    if save_to_drive:
        st.caption(
            "ไฟล์ที่บันทึกขึ้น Drive จะถูกแปลงเป็น Google Sheets จริง (ไม่ใช่ .xlsx) "
            "ทำให้สูตร IMPORTRANGE ในคอลัมน์ต้นทุน/กำไร/หมายเหตุคำนวณค่าได้ทันทีที่เปิดไฟล์ "
            "(ครั้งแรกอาจมี popup ให้กด \"Allow access\" หนึ่งครั้ง)"
        )
        drive_folder_input = st.text_input(
            "ลิงก์หรือ Folder ID ของ Google Drive ปลายทาง",
            value=default_folder,
            placeholder="เช่น https://drive.google.com/drive/folders/1AbC...XyZ",
            help="โฟลเดอร์นี้ต้องเป็นของ/แชร์ Editor ให้บัญชี Google ที่ใช้ตั้งค่า OAuth (ดูใน README.md)",
        )
elif has_gcp_secret():
    st.caption(
        "⚠️ ตั้งค่า Service Account ไว้แล้ว แต่ Service Account อัปโหลดไฟล์ขึ้น Drive ส่วนตัวไม่ได้ "
        "(ไม่มี storage quota เป็นของตัวเอง) ต้องตั้งค่า OAuth ผู้ใช้จริงเพิ่ม (st.secrets['gdrive_oauth']) "
        "ถึงจะบันทึกอัตโนมัติได้ — ดูวิธีใน README.md หัวข้อ \"ตั้งค่า OAuth สำหรับอัปโหลดขึ้น Google Drive\""
    )
else:
    st.caption("ต้องตั้งค่า Service Account + OAuth ก่อนถึงจะบันทึกขึ้น Drive อัตโนมัติได้ (ดูใน README.md)")

# ══════════════════════════════════════════════
# 7) ใบปะหน้า LAZADA — พารามิเตอร์ (ไม่บังคับ)
# ══════════════════════════════════════════════
st.subheader("7️⃣ ใบปะหน้า LAZADA — พารามิเตอร์ (ไม่บังคับ)")
st.caption(
    "ไฟล์ผลลัพธ์จะมีชีท \"ใบปะหน้า LAZADA\" (สรุปรายได้/รายจ่าย/กำไรแยกร้าน) พร้อมตารางค่าคอมพนักงานให้อัตโนมัติเสมอ "
    "จะกรอกอัตรา %/ชื่อพนักงานตรงนี้เลยก็ได้ หรือเว้นว่างไว้แล้วไปกรอกในไฟล์ Excel ทีหลังก็ได้ทั้งคู่"
)
face_c1, face_c2 = st.columns(2)
box_rate_pct = face_c1.number_input(
    "อัตราต้นทุนค่ากล่อง (%)", min_value=-100.0, max_value=100.0, value=0.0, step=0.1, format="%.2f"
)
pack_rate_pct = face_c2.number_input(
    "อัตราต้นทุนคน PACKING (%)", min_value=-100.0, max_value=100.0, value=0.0, step=0.1, format="%.2f"
)
box_cost_rate = (box_rate_pct / 100.0) if box_rate_pct != 0 else None
packing_cost_rate = (pack_rate_pct / 100.0) if pack_rate_pct != 0 else None

with st.expander("ตารางค่าคอมพนักงาน (ไม่บังคับ กรอกได้สูงสุด 3 คน)"):
    employees = []
    for _i in range(3):
        ec1, ec2 = st.columns([2, 1])
        _name = ec1.text_input(f"ชื่อพนักงานคนที่ {_i + 1}", key=f"emp_name_{_i}")
        _pct = ec2.number_input(
            f"% ที่กำหนด คนที่ {_i + 1}", min_value=0.0, max_value=100.0, value=0.0,
            step=0.1, key=f"emp_pct_{_i}", format="%.2f",
        )
        employees.append((_name.strip() if _name else None, (_pct / 100.0) if _pct > 0 else None))
    if not any(n for n, _ in employees):
        employees = None

st.divider()

ready = (bool(shop_files) and len(shop_files) == 5 and erp_file is not None
         and (stock_file is not None or use_drive_stock)
         and not (save_to_drive and not drive_folder_input.strip()))
run = st.button("🚀 เริ่มประมวลผล", type="primary", disabled=not ready)

if not ready:
    missing = []
    if len(shop_files) < 5:
        missing.append(f"ไฟล์ร้านค้า (ขาดอีก {5 - len(shop_files)} ร้าน)")
    if erp_file is None:
        missing.append("Data ERP.xlsx")
    if stock_file is None and not use_drive_stock:
        missing.append("STOCK.xlsx")
    if save_to_drive and not drive_folder_input.strip():
        missing.append("ลิงก์/Folder ID ของ Google Drive ปลายทาง")
    if missing:
        st.info("ยังขาด: " + ", ".join(missing))

if run:
    log_box = st.expander("📋 Log การประมวลผล", expanded=True)
    logs = []

    def log(msg):
        logs.append(str(msg))
        log_box.write(str(msg))

    tmpdir = tempfile.mkdtemp()
    try:
        creds = None
        if has_gcp_secret():
            from gsheet_helpers import get_credentials
            creds = get_credentials(dict(st.secrets["gcp_service_account"]))

        if use_drive_stock:
            if creds is None:
                st.error("ต้องตั้งค่า Service Account ก่อนถึงจะดึง STOCK จาก Drive อัตโนมัติได้")
                st.stop()
            from gsheet_helpers import fetch_stock_from_drive
            os.makedirs(f"{tmpdir}/STOCK", exist_ok=True)
            with st.spinner("กำลังดึง STOCK จาก Google Drive..."):
                fetch_stock_from_drive(creds, f"{tmpdir}/STOCK/STOCK.xlsx", log=log)

        with st.spinner("กำลังเตรียมไฟล์..."):
            build_folder_structure(tmpdir, shop_files, erp_file,
                                    None if use_drive_stock else stock_file, abbhol_files)

        combined_path = f"{tmpdir}/input_combined.xlsx"
        with st.spinner("ขั้นตอนที่ 1: รวมไฟล์ดิบ (build_combined)..."):
            build_combined(tmpdir, combined_path, log=log)

        # ต้นทุน/หมายเหตุ: เขียนเป็นสูตร IMPORTRANGE เสมอ เหมือน Skill lazada-sku-analysis เดิม
        cost_map, note_map = None, None

        output_path = f"{tmpdir}/ค่าคอมLAZADA.xlsx"
        with st.spinner("ขั้นตอนที่ 2: วิเคราะห์และสร้างรายงาน (build_analysis)..."):
            summary = build_analysis(
                combined_path, output_path, cost_map=cost_map, note_map=note_map, log=log,
                box_cost_rate=box_cost_rate, packing_cost_rate=packing_cost_rate, employees=employees,
            )

        st.success("ประมวลผลเสร็จแล้ว! 🎉")

        c1, c2, c3 = st.columns(3)
        c1.metric("รายได้รวม", f"{summary['income']:,.2f} บาท")
        c2.metric("ค่าใช้จ่ายรวม", f"{summary['expense']:,.2f} บาท")
        c3.metric("เงินเข้าบริษัทสุทธิ", f"{summary['net']:,.2f} บาท")

        st.write("**เงินเข้าบริษัทแยกร้าน**")
        shop_cols = st.columns(5)
        for i, shop in enumerate(SHOPS):
            shop_cols[i].metric(SHOP_LABELS[shop], f"{summary['shop_totals'].get(shop, 0):,.2f}")

        st.caption(
            f"📄 สร้างชีท \"ใบปะหน้า LAZADA\" (เดือน {summary.get('cover_month', '')}) "
            f"พร้อมตารางค่าคอมพนักงานให้แล้ว — เงินเข้าบริษัทสุทธิรวม {summary.get('cover_net', 0):,.2f} บาท"
        )

        with open(output_path, "rb") as f:
            st.download_button(
                "⬇️ ดาวน์โหลด ค่าคอมLAZADA.xlsx",
                data=f.read(),
                file_name="ค่าคอมLAZADA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        if not summary['used_real_cost_note']:
            st.info("ไฟล์นี้ยังมีสูตร IMPORTRANGE ในคอลัมน์ต้นทุน/กำไร/หมายเหตุ — ต้องเปิดใน Google Sheets ถึงจะเห็นค่า")

        if save_to_drive and has_gdrive_oauth():
            from gsheet_helpers import upload_to_drive, get_user_drive_credentials
            import datetime
            with st.spinner("กำลังบันทึกขึ้น Google Drive..."):
                try:
                    oauth_secret = st.secrets["gdrive_oauth"]
                    drive_creds = get_user_drive_credentials(
                        oauth_secret["client_id"], oauth_secret["client_secret"], oauth_secret["refresh_token"]
                    )
                    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                    drive_filename = f"ค่าคอมLAZADA_{ts}.xlsx"
                    uploaded = upload_to_drive(drive_creds, output_path, drive_filename, drive_folder_input, log=log)
                    link = uploaded.get("webViewLink")
                    st.success(f"บันทึกขึ้น Google Drive สำเร็จ: {drive_filename}")
                    if link:
                        st.markdown(f"[เปิดไฟล์ใน Google Drive]({link})")
                except Exception as e:
                    st.error(
                        f"บันทึกขึ้น Google Drive ไม่สำเร็จ: {e}\n\n"
                        "เช็คว่าแชร์สิทธิ์ Editor ของโฟลเดอร์ให้บัญชี Google ที่ใช้ตั้งค่า OAuth แล้วหรือยัง "
                        "หรือ refresh token อาจหมดอายุ (ถ้า OAuth app ยังอยู่ในโหมด Testing ต้องขอ token ใหม่ทุก 7 วัน "
                        "— แนะนำให้ Publish app เป็น Production ตามขั้นตอนใน README.md)"
                    )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
        st.code(traceback.format_exc())
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
