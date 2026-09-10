import streamlit as st
import io
import re
from datetime import datetime

from pptx import Presentation
from PIL import Image
import base64

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(page_title="Monthly Maintenance Highlight Builder", layout="wide")

# หัวข้อ (topic) มาตรฐาน และคำที่ใช้จับคู่กับหัวข้อสไลด์ (title) ใน PPTX รายสัปดาห์
TOPIC_DEFS = [
    ("Executive Summary", ["executive summary", "สรุปผู้บริหาร", "summary"]),
    ("Major Completed Works", ["completed work", "major completed", "งานที่แล้วเสร็จ", "งานที่ดำเนินการแล้วเสร็จ"]),
    ("Major Ongoing Works", ["ongoing work", "major ongoing", "งานที่กำลังดำเนินการ", "งานระหว่างดำเนินการ"]),
    ("Critical Equipment Issues", ["critical equipment", "critical issue", "อุปกรณ์วิกฤต"]),
    ("Major Risks", ["major risk", "risk", "ความเสี่ยง"]),
    ("Next Month Focus", ["next month", "focus", "แผนเดือนถัดไป", "เป้าหมายเดือนถัดไป"]),
]
TOPIC_ORDER = [t[0] for t in TOPIC_DEFS] + ["Other / ไม่ระบุหมวด"]

# Pattern ชื่อไฟล์: YYYY-MM_W{n}_{หน่วยงาน}.pptx เช่น 2026-08_W1_O21.pptx
FNAME_PATTERN = re.compile(r"(\d{4})-(\d{2})_W(\d+)_([A-Za-z0-9]+)", re.IGNORECASE)


def parse_filename_meta(filename: str):
    """แกะปี/เดือน/สัปดาห์/หน่วยงานจากชื่อไฟล์ เช่น '2026-08_W1_O21.pptx'"""
    m = FNAME_PATTERN.search(filename)
    if not m:
        return None
    year, month, week, unit = m.groups()
    try:
        month_name = datetime(int(year), int(month), 1).strftime("%B %Y")
    except ValueError:
        month_name = f"{year}-{month}"
    return {"year": year, "month": month, "week": week, "unit": unit.upper(), "month_name": month_name}


# ============================================================
# PASSWORD GATE
# ============================================================

def check_password():
    def password_entered():
        correct = st.secrets.get("APP_PASSWORD", "")
        if st.session_state.get("password_input", "") == correct and correct:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.title("🔒 Monthly Maintenance Highlight Builder")
    st.text_input("รหัสผ่าน", type="password", on_change=password_entered, key="password_input")
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("รหัสผ่านไม่ถูกต้อง")
    return False


if not check_password():
    st.stop()

# ============================================================
# HELPERS — อ่านไฟล์ PPTX (ข้อความเท่านั้น)
# ============================================================

def guess_topic(title_text: str) -> str:
    t = (title_text or "").strip().lower()
    if not t:
        return "Other / ไม่ระบุหมวด"
    best_topic, best_score = "Other / ไม่ระบุหมวด", 0
    for topic, keywords in TOPIC_DEFS:
        score = sum(1 for kw in keywords if kw in t)
        if score > best_score:
            best_topic, best_score = topic, score
    return best_topic


def extract_slide_title(slide) -> str:
    if slide.shapes.title is not None and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return shape.text_frame.text.strip().splitlines()[0]
    return ""


def extract_bullets(slide, title_text: str):
    bullets = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            line = "".join(run.text for run in para.runs).strip()
            if not line or line == title_text:
                continue
            bullets.append(line)
    return bullets


def parse_pptx_files(uploaded_files, week_labels):
    """คืนค่า dict: { topic: [ {"week": str, "lines": [str,...]}, ... ] }"""
    data = {topic: [] for topic in TOPIC_ORDER}

    for f, week_label in zip(uploaded_files, week_labels):
        prs = Presentation(io.BytesIO(f.getvalue()))
        for slide in prs.slides:
            title = extract_slide_title(slide)
            topic = guess_topic(title)
            bullets = extract_bullets(slide, title)
            if bullets:
                data[topic].append({"week": week_label, "lines": bullets})

    return data


def bullets_to_editable_text(bullet_groups):
    parts = []
    for group in bullet_groups:
        parts.append(f"[{group['week']}]")
        for line in group["lines"]:
            parts.append(f"- {line}")
        parts.append("")
    return "\n".join(parts).strip()


# ============================================================
# HELPERS — ภาพ (อัปโหลดเองทั้งหมด ไม่มีการดึงจาก PPTX / ไม่มี AI)
# ============================================================

def resize_image_b64(blob: bytes, max_side=1000, quality=80):
    try:
        im = Image.open(io.BytesIO(blob)).convert("RGB")
    except Exception:
        return None
    im.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=quality)
    return base64.b64encode(buf.getvalue()).decode()


# ============================================================
# HTML REPORT TEMPLATE
# ============================================================

def render_line_html(line: str) -> str:
    line = line.strip()
    if line.startswith("[") and line.endswith("]"):
        return f'<div class="week-tag">{line[1:-1]}</div>'
    if line.startswith("- "):
        return f'<li contenteditable="true">{line[2:].strip()}</li>'
    return f'<li contenteditable="true">{line}</li>'


def text_to_html_block(raw_text: str) -> str:
    html_lines = []
    open_ul = False
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("[") and line.endswith("]"):
            if open_ul:
                html_lines.append("</ul>")
                open_ul = False
            html_lines.append(render_line_html(line))
        else:
            if not open_ul:
                html_lines.append("<ul>")
                open_ul = True
            html_lines.append(render_line_html(line))
    if open_ul:
        html_lines.append("</ul>")
    return "\n".join(html_lines)


def build_html_report(title: str, period: str, sections: list) -> str:
    stat_items = "".join(
        f'<div class="meta-item"><b>{len(s["text"].splitlines())}</b>{s["topic"]}</div>'
        for s in sections if s["text"].strip()
    )

    body = ""
    for i, s in enumerate(sections, start=1):
        if not s["text"].strip() and not s["images"]:
            continue
        body += f'''
  <section>
    <div class="sec-head"><span class="sec-num">{i:02d}</span><h2 contenteditable="true">{s["topic"]}</h2></div>
    {text_to_html_block(s["text"])}
'''
        if s["images"]:
            body += '<div class="img-grid">'
            for b64 in s["images"]:
                body += f'<img src="data:image/jpeg;base64,{b64}" />'
            body += '</div>'
        body += "  </section>\n"

    return f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700;800&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap" rel="stylesheet">
<style>
  :root{{
    --paper:#EEEBE2; --panel:#F8F6F0; --ink:#1B2430; --ink-soft:#4A5568;
    --rule:#D8D3C4; --rust:#BF5B23; --teal:#35606B; --green:#4B7255;
  }}
  *{{box-sizing:border-box;}}
  body{{margin:0;background:var(--paper);color:var(--ink);font-family:'IBM Plex Sans',sans-serif;line-height:1.55;}}
  .sheet{{max-width:880px;margin:0 auto;padding:0 24px 80px;}}
  header{{padding:56px 0 32px;border-bottom:3px solid var(--ink);margin-bottom:40px;}}
  .plant-tag{{font-family:'IBM Plex Mono',monospace;font-size:13px;color:var(--rust);}}
  h1{{font-family:'Archivo',sans-serif;font-weight:800;font-size:clamp(28px,5vw,42px);letter-spacing:-0.01em;margin:10px 0 6px;}}
  .subdate{{font-size:16px;color:var(--ink-soft);}}
  .header-meta{{display:flex;gap:26px;margin-top:24px;flex-wrap:wrap;}}
  .meta-item{{font-size:13px;color:var(--ink-soft);}}
  .meta-item b{{display:block;font-family:'Archivo',sans-serif;font-size:20px;font-weight:700;color:var(--ink);}}
  section{{margin-bottom:48px;}}
  .sec-head{{display:flex;align-items:baseline;gap:14px;margin-bottom:18px;border-bottom:1px solid var(--rule);padding-bottom:10px;}}
  .sec-num{{font-family:'IBM Plex Mono',monospace;font-size:13px;color:var(--rust);min-width:22px;}}
  h2{{font-family:'Archivo',sans-serif;font-weight:700;font-size:22px;margin:0;outline:none;}}
  .week-tag{{font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--teal);margin:16px 0 8px;font-weight:500;}}
  ul{{margin:0 0 6px;padding-left:20px;}}
  li{{margin-bottom:8px;font-size:15px;outline:none;}}
  li::marker{{color:var(--teal);}}
  .img-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-top:14px;}}
  .img-grid img{{width:100%;border-radius:2px;border:1px solid var(--rule);display:block;}}
  footer{{margin-top:60px;padding-top:20px;border-top:1px solid var(--rule);font-size:12px;color:var(--ink-soft);font-family:'IBM Plex Mono',monospace;}}
  [contenteditable="true"]:hover{{background:rgba(53,96,107,0.06);cursor:text;}}
  .edit-hint{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-soft);background:var(--panel);border:1px dashed var(--rule);padding:10px 14px;margin-bottom:32px;}}
</style>
</head>
<body>
<div class="sheet">
  <header>
    <div class="plant-tag">Monthly Maintenance Highlight</div>
    <h1 contenteditable="true">{title}</h1>
    <div class="subdate">{period}</div>
    <div class="header-meta">{stat_items}</div>
  </header>

  <div class="edit-hint">แก้ไขข้อความในหน้านี้ได้โดยตรง (คลิกแล้วพิมพ์) — เมื่อแก้เสร็จ ใช้เมนูเบราว์เซอร์ Save Page As (Webpage, HTML only) เพื่อบันทึกเวอร์ชันที่แก้ไขแล้ว</div>

{body}
  <footer>{title.upper()} — {period}</footer>
</div>
</body>
</html>"""


# ============================================================
# STREAMLIT UI
# ============================================================

st.title("📋 Monthly Maintenance Highlight Builder")
st.caption("อัปโหลด Weekly Report (PowerPoint) หลายไฟล์ → รวมเนื้อหาตามหัวข้อ → แนบภาพเอง → ได้ Monthly Report เป็น HTML")

uploaded_files = st.file_uploader(
    "อัปโหลดไฟล์ Weekly Report (.pptx) — เลือกได้หลายไฟล์ — ตั้งชื่อไฟล์ตาม pattern YYYY-MM_W{n}_{หน่วยงาน} "
    "เช่น 2026-08_W1_O21.pptx จะช่วยตั้งชื่อสัปดาห์และช่วงเวลาให้อัตโนมัติ",
    type=["pptx"], accept_multiple_files=True
)

if uploaded_files:
    st.subheader("🏷️ ตั้งชื่อสัปดาห์ให้แต่ละไฟล์")
    st.caption("ดึงจากชื่อไฟล์ให้อัตโนมัติถ้าตรง pattern — แก้ไขเองได้ถ้าต้องการ")
    week_labels = []
    cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        meta = parse_filename_meta(f.name)
        default_label = f"Week {meta['week']} ({meta['unit']})" if meta else f"Week {i+1}"
        with cols[i % len(cols)]:
            label = st.text_input(f"📄 {f.name}", value=default_label, key=f"week_label_{f.name}_{i}")
        week_labels.append(label)

    if st.button("🔍 อ่านไฟล์และดึงเนื้อหา", type="primary"):
        with st.spinner("กำลังอ่านไฟล์ PowerPoint..."):
            st.session_state["parsed_data"] = parse_pptx_files(uploaded_files, week_labels)

        # ดึงเดือน/ปีจากชื่อไฟล์แรกที่ match pattern มาตั้งเป็นช่วงเวลาของรายงานให้อัตโนมัติ
        for f in uploaded_files:
            meta = parse_filename_meta(f.name)
            if meta:
                st.session_state["report_period_input"] = meta["month_name"]
                break

        st.rerun()

with st.sidebar:
    st.header("⚙️ ตั้งค่า")
    report_title = st.text_input("ชื่อรายงาน", value="Monthly Maintenance Highlight")
    report_period = st.text_input(
        "ช่วงเวลา (เช่น September 2026)",
        value=st.session_state.get("report_period_input", datetime.now().strftime("%B %Y")),
        key="report_period_input"
    )
    if st.button("🚪 ออกจากระบบ (ล้างรหัสผ่าน)"):
        st.session_state["password_correct"] = False
        st.rerun()

if "parsed_data" in st.session_state:
    data = st.session_state["parsed_data"]
    st.divider()
    st.subheader("✏️ ตรวจทานเนื้อหา และแนบภาพในแต่ละหัวข้อ")

    if "section_text" not in st.session_state:
        st.session_state["section_text"] = {}

    sections_final = []

    for topic in TOPIC_ORDER:
        bullet_groups = data.get(topic, [])
        if not bullet_groups:
            continue

        with st.expander(f"📌 {topic}  —  {len(bullet_groups)} กลุ่มข้อความ", expanded=True):
            default_text = bullets_to_editable_text(bullet_groups)
            key_text = f"text_{topic}"
            if key_text not in st.session_state["section_text"]:
                st.session_state["section_text"][key_text] = default_text

            edited_text = st.text_area(
                "เนื้อหา (แก้ไขได้อิสระ — บรรทัดที่ขึ้นต้นด้วย '- ' จะกลายเป็น bullet, บรรทัดในวงเล็บ [ ] คือป้ายกำกับสัปดาห์)",
                value=st.session_state["section_text"][key_text],
                height=180, key=f"ta_{topic}"
            )
            st.session_state["section_text"][key_text] = edited_text

            st.markdown("**🖼️ ภาพประกอบหัวข้อนี้**")
            count_key = f"imgcount_{topic}"
            num_images = st.number_input(
                "จำนวนกรอบภาพที่ต้องการ", min_value=0, max_value=20,
                value=st.session_state.get(count_key, 0), step=1, key=count_key
            )

            selected_b64 = []
            if num_images > 0:
                img_cols = st.columns(4)
                for i in range(int(num_images)):
                    with img_cols[i % 4]:
                        up = st.file_uploader(
                            f"ภาพที่ {i+1}", type=["png", "jpg", "jpeg"],
                            key=f"upimg_{topic}_{i}"
                        )
                        if up is not None:
                            b64 = resize_image_b64(up.getvalue())
                            if b64:
                                st.image(up, use_container_width=True)
                                selected_b64.append(b64)

            sections_final.append({"topic": topic, "text": edited_text, "images": selected_b64})

    st.divider()
    st.subheader("📄 Monthly Report")

    html_report = build_html_report(report_title, report_period, sections_final)

    tab1, tab2 = st.tabs(["👁️ ดูตัวอย่างในแอป", "⬇️ ดาวน์โหลด"])
    with tab1:
        st.components.v1.html(html_report, height=1000, scrolling=True)
    with tab2:
        st.download_button(
            "⬇️ ดาวน์โหลดเป็นไฟล์ HTML",
            data=html_report.encode("utf-8"),
            file_name=f"Monthly_Maintenance_Highlight_{report_period.replace(' ', '_')}.html",
            mime="text/html",
        )
        st.caption("ไฟล์ที่ดาวน์โหลดสามารถเปิดด้วยเบราว์เซอร์แล้วแก้ไขข้อความต่อได้โดยตรง (คลิกแล้วพิมพ์)")
else:
    st.info("อัปโหลดไฟล์ PowerPoint รายสัปดาห์ด้านบน แล้วกด 'อ่านไฟล์และดึงเนื้อหา' เพื่อเริ่มต้น")
