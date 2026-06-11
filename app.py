import streamlit as st
import io
import os
import datetime
import json

from docx import Document
from docx.shared import Inches
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
from PIL import Image  # 👈 นำเข้า Pillow สำหรับแปลงฟอร์แมตภาพอัตโนมัติ

# ─────────────────────────────────────────────
# 1. ค่าคงที่และฟังก์ชันจัดการข้อมูลผู้ใช้
# ─────────────────────────────────────────────
DATA_FILE  = "users.json"
SITES_FILE = "sites.json"

def load_users():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"usernames": {"admin": {"name": "Admin", "password": "default_password"}}}

def save_users(users):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=4, ensure_ascii=False)

# ─── Sites persistence ───
def _normalize_site(site_data: dict) -> dict:
    """แปลง key จาก str → int และ base64 → bytes หลัง JSON load"""
    import base64

    def decode_img(val):
        if val and isinstance(val, str):
            try:
                return base64.b64decode(val)
            except Exception:
                return None
        return None

    if "hotels" in site_data:
        hotels = {}
        for h_str, h_data in site_data["hotels"].items():
            h = int(h_str)
            hotels[h] = {}
            for i_str, item in h_data.items():
                i = int(i_str)
                hotels[h][i] = {
                    "desc":   item.get("desc", ""),
                    "locked": item.get("locked", False),
                    "img":    decode_img(item.get("img")),
                }
        site_data["hotels"] = hotels

    if "fuel" in site_data:
        fuel = {}
        for n_str, item in site_data["fuel"].items():
            n = int(n_str)
            # แปลง date string กลับเป็น date object
            date_val = item.get("date", "")
            try:
                date_obj = datetime.date.fromisoformat(date_val) if date_val else datetime.date.today()
            except Exception:
                date_obj = datetime.date.today()
            fuel[n] = {
                "date":     date_obj,
                "province": item.get("province", ""),
                "locked":   item.get("locked", False),
                "bill":     decode_img(item.get("bill")),
                "pre":      decode_img(item.get("pre")),
                "post":     decode_img(item.get("post")),
            }
        site_data["fuel"] = fuel

    # รูปรถ
    for key in ["start_img", "end_img", "car_img"]:
        site_data[key] = decode_img(site_data.get(key))

    return site_data

def load_sites() -> dict:
    """โหลดข้อมูลทุกไซต์จากไฟล์ JSON"""
    if os.path.exists(SITES_FILE):
        try:
            with open(SITES_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            # ป้องกัน format เก่าที่อาจเป็น list หรือ structure อื่น
            if not isinstance(raw, dict):
                return {}
            return {name: _normalize_site(data) for name, data in raw.items()
                    if isinstance(data, dict)}
        except Exception:
            # ไฟล์เสียหายหรือ format ผิด → เริ่มใหม่
            return {}
    return {}

def save_sites(sites: dict):
    """บันทึกข้อมูลทุกไซต์ลงไฟล์ JSON (ข้าม bytes รูปภาพ)"""
    import base64

    def encode_img(data):
        """แปลง bytes เป็น base64 string เพื่อเก็บใน JSON"""
        if data is None:
            return None
        if isinstance(data, bytes):
            return base64.b64encode(data).decode("utf-8")
        return None

    serializable = {}
    for site_name, site_data in sites.items():
        hotels_out = {}
        for h in range(1, 4):
            hotels_out[str(h)] = {}
            for i in range(1, 7):
                item = site_data.get("hotels", {}).get(h, {}).get(i, {})
                hotels_out[str(h)][str(i)] = {
                    "desc":   item.get("desc", ""),
                    "locked": item.get("locked", False),
                    "img":    encode_img(item.get("img")),
                }

        fuel_out = {}
        for n in range(1, 21):
            item = site_data.get("fuel", {}).get(n, {})
            fuel_out[str(n)] = {
                "date":     str(item.get("date", "")),
                "province": item.get("province", ""),
                "locked":   item.get("locked", False),
                "bill":     encode_img(item.get("bill")),
                "pre":      encode_img(item.get("pre")),
                "post":     encode_img(item.get("post")),
            }

        serializable[site_name] = {
            "created_at": site_data.get("created_at", ""),
            "updated_at": site_data.get("updated_at", ""),
            "start_mile": site_data.get("start_mile", 0),
            "end_mile":   site_data.get("end_mile", 0),
            "start_img":  encode_img(site_data.get("start_img")),
            "end_img":    encode_img(site_data.get("end_img")),
            "car_img":    encode_img(site_data.get("car_img")),
            "hotels":     hotels_out,
            "fuel":       fuel_out,
        }

    with open(SITES_FILE, "w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=4, ensure_ascii=False)

if "credentials" not in st.session_state:
    st.session_state.credentials = load_users()

if "sites" not in st.session_state:
    st.session_state.sites = load_sites()

# ─────────────────────────────────────────────
# 2. Admin Panel
# ─────────────────────────────────────────────
def admin_panel():
    st.subheader("🛠️ Admin Settings: จัดการผู้ใช้งาน")
    with st.expander("➕ เพิ่มผู้ใช้งานใหม่"):
        new_username = st.text_input("Username")
        new_name     = st.text_input("ชื่อจริง")
        new_password = st.text_input("รหัสผ่าน", type="password")
        if st.button("บันทึกผู้ใช้"):
            if new_username and new_password:
                st.session_state.credentials["usernames"][new_username] = {
                    "name": new_name, "password": new_password
                }
                save_users(st.session_state.credentials)
                st.success(f"บันทึกผู้ใช้ {new_username} เรียบร้อย")
            else:
                st.error("กรุณากรอกข้อมูลให้ครบ")

    st.write("### รายชื่อผู้ใช้")
    for username in list(st.session_state.credentials["usernames"].keys()):
        col1, col2 = st.columns([0.8, 0.2])
        col1.write(f"- {username}")
        if col2.button("ลบ", key=f"del_{username}"):
            del st.session_state.credentials["usernames"][username]
            save_users(st.session_state.credentials)
            st.rerun()

    st.write("---")
    if st.button("รีเซ็ตสถานะหน้าเพจ"):
        st.session_state.page = "site_selector"
        st.rerun()

# ─────────────────────────────────────────────
# 3. ฟังก์ชันจัดการไซต์งาน
# ─────────────────────────────────────────────
def now_str() -> str:
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

def init_site(site_name: str):
    """สร้างหรือโหลดข้อมูลไซต์ใน session_state"""
    if site_name not in st.session_state.sites:
        # ไซต์ใหม่
        st.session_state.sites[site_name] = {
            "created_at": now_str(),
            "updated_at": now_str(),
            "hotels": {
                h: {i: {"img": None, "desc": "", "locked": False} for i in range(1, 7)}
                for h in range(1, 4)
            },
            "fuel": {
                n: {
                    "bill": None, "pre": None, "post": None,
                    "locked": False,
                    "date": datetime.date.today(),
                    "province": "",
                }
                for n in range(1, 21)
            },
            "start_mile": 0,
            "end_mile":   0,
        }
    # ชี้ session ปัจจุบันไปที่ไซต์นี้
    st.session_state.current_site = site_name
    st.session_state.hotels = st.session_state.sites[site_name]["hotels"]
    st.session_state.fuel   = st.session_state.sites[site_name]["fuel"]

def touch_site():
    """อัปเดต updated_at ของไซต์ปัจจุบัน"""
    site = st.session_state.get("current_site")
    if site and site in st.session_state.sites:
        st.session_state.sites[site]["updated_at"] = now_str()
        save_sites(st.session_state.sites)

# ─────────────────────────────────────────────
# 4. Toggle Lock
# ─────────────────────────────────────────────
def _to_bytes(file_obj):
    if file_obj is None:
        return None
    if isinstance(file_obj, bytes):
        return file_obj
    try:
        file_obj.seek(0)
        return file_obj.read()
    except Exception:
        return None

def toggle_lock(section, key, sub_key=None):
    if section == "hotel":
        item = st.session_state.hotels[key][sub_key]
        if not item["locked"]:
            item["img"] = _to_bytes(item["img"])
        item["locked"] = not item["locked"]
    else:
        item = st.session_state.fuel[key]
        if not item["locked"]:
            item["bill"] = _to_bytes(item["bill"])
            item["pre"]  = _to_bytes(item["pre"])
            item["post"] = _to_bytes(item["post"])
        item["locked"] = not item["locked"]
    touch_site()

# ─────────────────────────────────────────────
# 5. สร้างไฟล์ Word
# ─────────────────────────────────────────────
def _img_stream(data):
    """คืน BytesIO จาก bytes หรือ UploadedFile พร้อมใช้ Pillow แปลงภาพเป็น PNG"""
    if data is None:
        return None
    try:
        if isinstance(data, bytes):
            raw = data
        else:
            data.seek(0)
            raw = data.read()
        if not raw:
            return None
        
        # 💡 ใช้ Pillow โหลดภาพ แล้วเซฟเป็น PNG เพื่อป้องกันปัญหา WebP หรือภาพฟอร์แมตอื่นๆ ใน docx
        img = Image.open(io.BytesIO(raw))
        out_stream = io.BytesIO()
        img.save(out_stream, format="PNG")
        out_stream.seek(0)
        return out_stream
    except Exception:
        return None

def _safe_add_picture(run, data, width):
    """ใส่รูปใน Word run อย่างปลอดภัย คืน True ถ้าสำเร็จ"""
    stream = _img_stream(data)
    if stream is None:
        return False
    try:
        run.add_picture(stream, width=width)
        return True
    except Exception:
        return False

def generate_word():
    doc  = Document()
    site = st.session_state.get("current_site", "")
    info = st.session_state.sites.get(site, {})
    doc.add_heading(f"Trip Report — {site}", 0)
    doc.add_paragraph(f"สร้างเมื่อ: {info.get('created_at', '-')}   อัปเดต: {info.get('updated_at', '-')}")

    # ── ส่วนที่ 1: ข้อมูลรถ ──
    doc.add_heading("ข้อมูลรถและการเดินทาง", level=1)
    start_mile = info.get("start_mile") or st.session_state.get("start_mile") or 0
    end_mile   = info.get("end_mile")   or st.session_state.get("end_mile")   or 0
    distance   = max(0, end_mile - start_mile)
    doc.add_paragraph(f"เลขไมล์เริ่มต้น : {start_mile}")
    doc.add_paragraph(f"เลขไมล์หลังจบ  : {end_mile}")
    doc.add_paragraph(f"ระยะทางรวม     : {distance} กม.")

    # รูปรถ — ดึงจาก sites (bytes ที่บันทึกไว้) ก่อน ถ้าไม่มีค่อยดึงจาก widget
    car_images = []
    for label, key in [("ไมล์ก่อนเริ่ม", "start_img"),
                        ("ไมล์หลังจบ",    "end_img"),
                        ("รูปรถทั่วไป",   "car_img")]:
        raw = info.get(key) or st.session_state.get(key)
        if raw is not None:
            car_images.append((label, raw))

    if car_images:
        tbl = doc.add_table(rows=2, cols=len(car_images))
        tbl.style = "Table Grid"
        for ci, (label, raw_data) in enumerate(car_images):
            tbl.rows[0].cells[ci].text = label
            p   = tbl.rows[1].cells[ci].paragraphs[0]
            run = p.add_run()
            _safe_add_picture(run, raw_data, Inches(1.8))
    doc.add_paragraph("")

    # ── ส่วนที่ 2: โรงแรม ──
    doc.add_heading("รูปภาพโรงแรม", level=1)
    any_hotel = False
    for h in range(1, 4):
        items_in_hotel = [
            (i, st.session_state.hotels[h][i])
            for i in range(1, 7)
            if st.session_state.hotels[h][i]["locked"]
            and st.session_state.hotels[h][i]["img"] is not None
        ]
        if not items_in_hotel:
            continue
        any_hotel = True
        doc.add_heading(f"โรงแรมที่ {h}", level=2)
        for i, item in items_in_hotel:
            doc.add_paragraph(f"รายการที่ {i}: {item['desc']}")
            run = doc.add_paragraph().add_run()
            _safe_add_picture(run, item["img"], Inches(2.5))
    if not any_hotel:
        doc.add_paragraph("(ยังไม่มีรูปโรงแรมที่บันทึกแล้ว)")

    # ── ส่วนที่ 3: น้ำมัน ──
    doc.add_heading("บันทึกการเติมน้ำมัน", level=1)
    fuel_count = 0
    for n in range(1, 21):
        item = st.session_state.fuel.get(n)
        if not item or not item["locked"]:
            continue
        fuel_count += 1
        doc.add_heading(f"การเติมครั้งที่ {n}", level=2)
        doc.add_paragraph(f"วันที่  : {item.get('date', '-')}")
        doc.add_paragraph(f"จังหวัด: {item.get('province', '-')}")

        # นับคอลัมน์ที่มีรูปจริง
        img_data = [(lbl, item.get(k))
                    for lbl, k in [("ใบเสร็จ","bill"),("ไมล์ก่อนเติม","pre"),("ไมล์หลังเติม","post")]]
        img_data = [(lbl, d) for lbl, d in img_data if d is not None]

        if img_data:
            tbl = doc.add_table(rows=2, cols=len(img_data))
            tbl.style = "Table Grid"
            for ci, (lbl, raw_data) in enumerate(img_data):
                tbl.rows[0].cells[ci].text = lbl
                p   = tbl.rows[1].cells[ci].paragraphs[0]
                run = p.add_run()
                _safe_add_picture(run, raw_data, Inches(1.8))
        doc.add_paragraph("")

    if fuel_count == 0:
        doc.add_paragraph("(ยังไม่มีรายการเติมน้ำมันที่บันทึกแล้ว)")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# 6. UI: โรงแรม
# ─────────────────────────────────────────────
def render_hotel_section():
    st.header("🏨 ส่วนที่ 1: รูปภาพโรงแรม")
    for h in range(1, 4):
        with st.expander(f"โรงแรมที่ {h}"):
            cols = st.columns(3)
            for i in range(1, 7):
                item = st.session_state.hotels[h][i]
                with cols[(i - 1) % 3]:
                    if not item["locked"]:
                        up = st.file_uploader("เลือกรูป", key=f"up_h_{h}_{i}")
                        if up:
                            item["img"] = up
                        item["desc"] = st.text_input(
                            "คำอธิบาย", value=item["desc"], key=f"desc_h_{h}_{i}"
                        )
                        if st.button("บันทึก", key=f"btn_h_save_{h}_{i}"):
                            toggle_lock("hotel", h, i)
                            st.rerun()
                    else:
                        if item["img"] is not None:
                            st.image(item["img"], use_container_width=True)
                        st.info(f"📝 {item['desc']}")
                        if st.button("แก้ไข", key=f"btn_h_edit_{h}_{i}"):
                            toggle_lock("hotel", h, i)
                            st.rerun()

# ─────────────────────────────────────────────
# 7. UI: ข้อมูลรถ
# ─────────────────────────────────────────────
def render_car_section():
    st.subheader("📋 ข้อมูลรถและการเดินทาง")
    site     = st.session_state.get("current_site", "")
    site_inf = st.session_state.sites.get(site, {})

    with st.container(border=True):
        col_a, col_b = st.columns(2)
        with col_a:
            st.file_uploader("รูปไมล์ก่อนเริ่ม", key="start_img")
            st.number_input("เลขไมล์เริ่มต้น", key="start_mile", min_value=0)
            show = st.session_state.get("start_img") or site_inf.get("start_img")
            if show:
                st.image(show, caption="ไมล์ก่อนเริ่ม ✅" if site_inf.get("start_img") else "ไมล์ก่อนเริ่ม")
        with col_b:
            st.file_uploader("รูปไมล์หลังจบทริป", key="end_img")
            st.number_input("เลขไมล์หลังจบ", key="end_mile", min_value=0)
            show = st.session_state.get("end_img") or site_inf.get("end_img")
            if show:
                st.image(show, caption="ไมล์หลังจบ ✅" if site_inf.get("end_img") else "ไมล์หลังจบ")

    st.file_uploader("อัปโหลดรูปรถ (ทั่วไป)", key="car_img")
    show = st.session_state.get("car_img") or site_inf.get("car_img")
    if show:
        st.image(show, caption="รูปรถ ✅" if site_inf.get("car_img") else "รูปรถ")

    if st.button("💾 บันทึกข้อมูลรถ"):
        site = st.session_state.get("current_site")
        if site:
            s = st.session_state.sites[site]
            s["start_mile"] = st.session_state.get("start_mile", 0)
            s["end_mile"]   = st.session_state.get("end_mile", 0)
            # บันทึก bytes รูปรถไว้ใน sites เพื่อใช้ใน Word
            for key in ["start_img", "end_img", "car_img"]:
                uploaded = st.session_state.get(key)
                if uploaded is not None:
                    s[key] = _to_bytes(uploaded)
            touch_site()
        st.success("✅ บันทึกข้อมูลรถเรียบร้อย")

# ─────────────────────────────────────────────
# 8. UI: น้ำมัน
# ─────────────────────────────────────────────
def render_fuel_section():
    st.header("⛽ ส่วนที่ 2: บันทึกการเติมน้ำมัน")
    render_car_section()
    st.divider()

    for n in range(1, 21):
        with st.expander(f"การเติมครั้งที่ {n}"):
            item = st.session_state.fuel[n]

            if not item["locked"]:
                c1, c2, c3 = st.columns(3)
                with c1:
                    up_bill = st.file_uploader("📸 ใบเสร็จ", key=f"up_bill_{n}")
                    if up_bill:
                        item["bill"] = up_bill
                with c2:
                    up_pre = st.file_uploader("🔼 ไมล์ก่อนเติม", key=f"pre_{n}")
                    if up_pre:
                        item["pre"] = up_pre
                with c3:
                    up_post = st.file_uploader("🔽 ไมล์หลังเติม", key=f"post_{n}")
                    if up_post:
                        item["post"] = up_post

                c_date, c_prov = st.columns(2)
                with c_date:
                    item["date"] = st.date_input(
                        "วันที่เติม",
                        value=item.get("date") or datetime.date.today(),
                        key=f"date_{n}",
                    )
                with c_prov:
                    item["province"] = st.text_input(
                        "จังหวัด", value=item.get("province", ""), key=f"prov_{n}"
                    )

                if st.button("💾 บันทึกรายการ", key=f"save_f_{n}"):
                    toggle_lock("fuel", n)
                    st.rerun()
            else:
                st.write("✅ บันทึกรายการนี้แล้ว")
                st.info(f"📅 วันที่: {item['date']} | 📍 จังหวัด: {item['province']}")
                col_receipt, col_miles = st.columns([2, 1])
                with col_receipt:
                    if item.get("bill"):
                        st.image(item["bill"], caption="ใบเสร็จ", use_container_width=True)
                with col_miles:
                    if item.get("pre"):
                        st.image(item["pre"], caption="ไมล์ก่อนเติม", use_container_width=True)
                    if item.get("post"):
                        st.image(item["post"], caption="ไมล์หลังเติม", use_container_width=True)
                if st.button("✏️ แก้ไขรายการ", key=f"edit_f_{n}"):
                    toggle_lock("fuel", n)
                    st.rerun()

# ─────────────────────────────────────────────
# 9. UI: Export
# ─────────────────────────────────────────────
def render_export_section():
    site     = st.session_state.get("current_site", "trip")
    filename = f"Trip_Report_{site}.docx"
    mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    try:
        buf  = generate_word()
        data = buf.getvalue()
        ok   = True
        err  = ""
    except Exception as e:
        ok  = False
        err = str(e)

    # Sidebar
    st.sidebar.header("📤 Export Reports")
    if ok:
        st.sidebar.download_button(
            label="📥 ดาวน์โหลด Word",
            data=data, file_name=filename, mime=mime, key="dl_sidebar",
        )
    else:
        st.sidebar.error(f"สร้างไฟล์ไม่ได้: {err}")

    # หน้าหลัก
    st.divider()
    st.header("📤 Export รายงาน")

    hotel_count = sum(
        1 for h in range(1, 4) for i in range(1, 7)
        if st.session_state.hotels[h][i]["locked"]
        and st.session_state.hotels[h][i]["img"] is not None
    )
    fuel_count = sum(
        1 for n in range(1, 21)
        if st.session_state.fuel[n]["locked"]
        and st.session_state.fuel[n]["bill"] is not None
    )

    col_dl, col_info = st.columns([1, 2])
    with col_info:
        st.info(
            f"📦 รายงานประกอบด้วย\n"
            f"- 🏨 รูปโรงแรมที่บันทึกแล้ว: **{hotel_count}** รายการ\n"
            f"- ⛽ การเติมน้ำมันที่บันทึกแล้ว: **{fuel_count}** ครั้ง"
        )
    with col_dl:
        if ok:
            st.download_button(
                label="📥 ดาวน์โหลดรายงาน Word",
                data=data, file_name=filename, mime=mime,
                key="dl_main", type="primary", use_container_width=True,
            )
        else:
            st.error(f"❌ สร้างไฟล์ Word ไม่ได้\n\n`{err}`")

# ─────────────────────────────────────────────
# 10. UI: หน้าเลือก/สร้างไซต์งาน
# ─────────────────────────────────────────────
def render_site_selector():
    st.title("🏗️ ไซต์งาน")

    # ── รายการไซต์ที่มีอยู่ ──
    sites = st.session_state.sites
    if sites:
        st.subheader("📋 ไซต์งานที่บันทึกไว้")
        for site_name, info in sorted(
            sites.items(),
            key=lambda x: x[1].get("updated_at", ""),
            reverse=True,
        ):
            with st.container(border=True):
                col_name, col_meta, col_btn = st.columns([2, 3, 1])

                with col_name:
                    st.markdown(f"### 📁 {site_name}")

                with col_meta:
                    created = info.get("created_at", "-")
                    updated = info.get("updated_at", "-")
                    # นับรายการจาก hotels/fuel จริง
                    h_count = sum(
                        1
                        for h_data in info.get("hotels", {}).values()
                        for item in h_data.values()
                        if item.get("locked") and item.get("img") is not None
                    )
                    f_count = sum(
                        1
                        for item in info.get("fuel", {}).values()
                        if item.get("locked") and item.get("bill") is not None
                    )
                    st.caption(f"🕐 สร้าง: {created}")
                    st.caption(f"🔄 อัปเดต: {updated}")
                    st.caption(f"🏨 โรงแรม: {h_count} รายการ  |  ⛽ น้ำมัน: {f_count} ครั้ง")

                with col_btn:
                    if st.button("เข้าแก้ไข", key=f"open_{site_name}", use_container_width=True, type="primary"):
                        init_site(site_name)
                        st.session_state.page = "trip_logger"
                        st.rerun()
                    if st.button("🗑️ ลบ", key=f"del_site_{site_name}", use_container_width=True):
                        del st.session_state.sites[site_name]
                        save_sites(st.session_state.sites)
                        st.rerun()
    else:
        st.info("ยังไม่มีไซต์งาน กรุณาสร้างไซต์งานใหม่ด้านล่าง")

    st.divider()

    # ── สร้างไซต์ใหม่ ──
    st.subheader("➕ สร้างไซต์งานใหม่")
    col_input, col_btn2 = st.columns([3, 1])
    with col_input:
        site_name = st.text_input("ชื่อไซต์งาน", label_visibility="collapsed", placeholder="ระบุชื่อไซต์งาน...")
    with col_btn2:
        if st.button("สร้าง", type="primary", use_container_width=True):
            if site_name:
                if site_name in st.session_state.sites:
                    st.warning(f"ไซต์งาน '{site_name}' มีอยู่แล้ว กดเข้าแก้ไขได้เลย")
                else:
                    init_site(site_name)
                    save_sites(st.session_state.sites)
                    st.session_state.page = "trip_logger"
                    st.rerun()
            else:
                st.warning("กรุณาระบุชื่อไซต์งาน")

# ─────────────────────────────────────────────
# 11. Main App
# ─────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="Trip Logger Pro")

with open("config.yaml", encoding="utf-8") as file:
    config = yaml.load(file, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

authenticator.login()

if st.session_state["authentication_status"] is False:
    st.error("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")

elif st.session_state["authentication_status"] is None:
    st.warning("กรุณากรอกชื่อผู้ใช้และรหัสผ่านเพื่อเข้าสู่ระบบ")

elif st.session_state["authentication_status"]:
    st.sidebar.write(f'ยินดีต้อนรับ *{st.session_state["name"]}*')
    authenticator.logout("Logout", "sidebar")

    if st.session_state.get("username") == "admin":
        with st.sidebar.expander("Admin Settings"):
            admin_panel()

    if "page" not in st.session_state:
        st.session_state.page = "site_selector"

    # ─── หน้าเลือกไซต์งาน ───
    if st.session_state.page == "site_selector":
        render_site_selector()

    # ─── หน้าบันทึกข้อมูลหลัก ───
    elif st.session_state.page == "trip_logger":
        site = st.session_state.get("current_site", "")

        # ตรวจสอบว่า hotels/fuel ชี้ไปที่ไซต์ที่ถูกต้อง
        if "hotels" not in st.session_state or "fuel" not in st.session_state:
            init_site(site)

        st.sidebar.info(f"📁 ไซต์งาน: {site}")
        if st.sidebar.button("⬅️ กลับหน้าไซต์งาน"):
            save_sites(st.session_state.sites)
            st.session_state.page = "site_selector"
            st.rerun()

        st.title(f"📸 Trip Logger: {site}")

        render_hotel_section()
        st.divider()
        render_fuel_section()
        render_export_section()