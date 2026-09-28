# -*- coding: utf-8 -*-
"""
Monthly Maintenance Highlight Builder (O-MN2-MM1)
สร้างไฟล์ PowerPoint รายงานประจำเดือนจาก
  1) Weekly Report (PPTX) ของแต่ละสัปดาห์  -> สไลด์รายอุปกรณ์ (1 หน้า / Tag)
  2) Excel งานประจำเดือน (sheet O21, O22)   -> Monthly Job Summary
  3) ไฟล์ CAPEX (PPTX ที่มีตาราง)           -> ตาราง CAPEX
ไฟล์ที่ต้องอยู่ใน repo เดียวกัน: app.py, requirements.txt, template.pptx
"""
import io
import os
import re
import json
import math
import hashlib
import collections
import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image
import openpyxl

# ============================================================
# CONFIG
# ============================================================
FONT = "Tahoma"
COL = dict(dark="1B2A38", ink="1F2A37", muted="5B6770", teal="2F6F73", card="F1F4F6",
           line="D9E0E5", amber="F2B134", green="2E7D4F", greenbg="E1F0E7",
           orange="9A5B00", orangebg="FFEFD2", other="7A8791", white="FFFFFF", tot="E6ECEF",
           soft="D7DEE4", faint="9FB0BD")
HERE = os.path.dirname(os.path.abspath(__file__))

MONTH_ABBR = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
              "sep": 9, "oct": 10, "nov": 11, "dec": 12}
MONTH_ABBR.update({"january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7,
                   "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
                   "sept": 9})
FNAME_PATTERN = re.compile(r"(\d{4})-(\d{2})_W(\d+)_([A-Za-z0-9]+)", re.IGNORECASE)
ROW_TOL = 150000   # EMU ~0.16in : ถือว่าอยู่แถวเดียวกัน


# ============================================================
# ส่วนที่ 1: อ่านสไลด์ Weekly Report (การ์ดงาน)
# ============================================================
def parse_filename_meta(name):
    m = FNAME_PATTERN.search(name or "")
    if not m:
        return None
    y, mo, w, unit = m.groups()
    try:
        label = datetime.date(int(y), int(mo), 1).strftime("%B %Y")
    except ValueError:
        label = "%s-%s" % (y, mo)
    return dict(year=int(y), month=int(mo), week=int(w), unit=unit.upper(), label=label)


def _lab(text):
    return re.sub(r"[\s:]+", "", (text or "").lower())


LABELS = {"equipment": "tag", "background/information": "background", "background": "background",
          "possiblecause": "cause", "action": "action"}


def norm_key(tag):
    """ตัด prefix O5- / B- ออกเพื่อจับ Tag เดียวกัน เช่น O5-R-3802 = R-3802, B-M-3104 = M-3104"""
    t = re.sub(r"\s+", "", (tag or "").upper())
    t = re.sub(r"^O\d+-", "", t)
    t = re.sub(r"^B-", "", t)
    return t


def display_tag(variants):
    """เลือกชื่อ Tag ที่จะโชว์: ถ้ามีแบบขึ้นต้น B- ใช้แบบนั้น ไม่งั้นตัด O5- ออก"""
    vs = [re.sub(r"\s+", "", v.strip()) for v in variants if v and v.strip()]
    for v in vs:
        if v.upper().startswith("B-"):
            return v
    v = vs[0] if vs else ""
    return re.sub(r"^O\d+-", "", v, flags=re.I)


def title_tag(text):
    m = re.match(r"^[\s#]*([A-Za-z0-9][A-Za-z0-9\-/\.]*)", text or "")
    if not m:
        return None
    tok = m.group(1).rstrip(".")
    if "-" in tok and re.search(r"\d", tok):
        return tok
    return None


BOILER = re.compile(r"^(update\s+(activity|work)[^:\n]{0,20}:?|action\s+work\s*:?)\s*", re.I)


def clean_lines(text):
    out = []
    for raw in (text or "").splitlines():
        s = raw.replace("\xa0", " ").replace("\u200b", "").strip()
        s = re.sub(r"^[#\-•*\u2022\s]+", "", s).strip()
        s = BOILER.sub("", s).strip()
        s = re.sub(r"^[#\-•*\s]+", "", s).strip()
        if s:
            out.append(s)
    return out


def dedup(lines):
    seen, out = set(), []
    for l in lines:
        k = re.sub(r"\s+", "", l.lower())
        if k not in seen:
            seen.add(k)
            out.append(l)
    return out


def parse_date_text(raw):
    """คืน (start, end) เป็น date หรือ (None, None) — รองรับ 1-5Sep 2026 / 7-12 Sep 26 / 28Aug-3Sep 2026"""
    if not raw:
        return None, None

    def yr(y):
        y = int(y)
        return 2000 + y if y < 100 else y

    m = re.search(r"(\d{1,2})\s*([A-Za-z]{3,9})\s*-\s*(\d{1,2})\s*([A-Za-z]{3,9})\s*(\d{2,4})", raw)
    if m:
        d1, m1, d2, m2, y = m.groups()
        a, b = MONTH_ABBR.get(m1.lower()), MONTH_ABBR.get(m2.lower())
        if a and b:
            try:
                return datetime.date(yr(y), a, int(d1)), datetime.date(yr(y), b, int(d2))
            except ValueError:
                pass
    m = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})\s*([A-Za-z]{3,9})\s*(\d{2,4})", raw)
    if m:
        d1, d2, mo, y = m.groups()
        mm = MONTH_ABBR.get(mo.lower())
        if mm:
            try:
                return datetime.date(yr(y), mm, int(d1)), datetime.date(yr(y), mm, int(d2))
            except ValueError:
                pass
    m = re.search(r"(\d{1,2})\s*([A-Za-z]{3,9})\s*(\d{2,4})", raw)
    if m:
        d1, mo, y = m.groups()
        mm = MONTH_ABBR.get(mo.lower())
        if mm:
            try:
                d = datetime.date(yr(y), mm, int(d1))
                return d, d
            except ValueError:
                pass
    return None, None


def fmt_range(a, b):
    if a.year != b.year:
        return "%d %s %d – %d %s %d" % (a.day, a.strftime("%b"), a.year, b.day, b.strftime("%b"), b.year)
    if a.month != b.month:
        return "%d %s – %d %s %d" % (a.day, a.strftime("%b"), b.day, b.strftime("%b"), b.year)
    if a.day == b.day:
        return "%d %s %d" % (a.day, a.strftime("%b"), a.year)
    return "%d – %d %s %d" % (a.day, b.day, a.strftime("%b"), a.year)


def extract_card(slide):
    """อ่าน 1 สไลด์ = 1 การ์ดงาน จับคู่ label -> value ด้วยตำแหน่ง (ขวาแถวเดียวกันก่อน แล้วค่อยด้านล่าง)"""
    boxes = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            boxes.append(dict(top=sh.top, left=sh.left, text=sh.text_frame.text.strip()))
    f = dict(tag="", plant="", date_raw="", background="", cause="", action="", title="")
    rest = []
    for b in boxes:
        low = b["text"].lower()
        if low.startswith("plant") and b["top"] < Inches(0.8):
            f["plant"] = b["text"][5:].strip()
            continue
        if low.startswith("date") and b["top"] < Inches(0.8):
            v = b["text"][4:].strip()
            f["date_raw"] = v[1:].strip() if v.startswith(":") else v
            continue
        if b["left"] < Inches(5) and b["top"] < Inches(1.2) and not f["title"]:
            f["title"] = b["text"]
        rest.append(b)
    used = set()
    for i, b in enumerate(rest):
        key = LABELS.get(_lab(b["text"]))
        if not key:
            continue
        best, dist = None, None
        for j, c in enumerate(rest):
            if j == i or j in used or _lab(c["text"]) in LABELS:
                continue
            if abs(c["top"] - b["top"]) <= ROW_TOL and c["left"] > b["left"]:
                d = c["left"] - b["left"]
                if dist is None or d < dist:
                    best, dist = j, d
        if best is None:
            for j, c in enumerate(rest):
                if j == i or j in used or _lab(c["text"]) in LABELS:
                    continue
                if c["top"] > b["top"]:
                    d = c["top"] - b["top"]
                    if dist is None or d < dist:
                        best, dist = j, d
        if best is not None:
            f[key] = rest[best]["text"]
            used.add(best)
    return f


def slide_images(slide):
    out = []
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                out.append(sh.image.blob)
            except Exception:
                pass
    return out


def _drawing_score(im):
    try:
        g = im.convert("L").resize((48, 48))
        px = list(g.getdata())
        return sum(1 for p in px if p > 232) / float(len(px))
    except Exception:
        return 0.0


def _thumb(im, side=220):
    t = im.convert("RGB")
    t.thumbnail((side, side))
    buf = io.BytesIO()
    t.save(buf, "JPEG", quality=70)
    return buf.getvalue()


def pick_evenly(idx_list, n):
    m = len(idx_list)
    if m <= n:
        return list(idx_list)
    return [idx_list[int(round(i * (m - 1) / float(n - 1)))] for i in range(n)]


def parse_weekly(files):
    """files: list[(name, bytes)] -> (items, warnings)"""
    files = sorted(files, key=lambda x: x[0])
    slides = []
    hash_slides = collections.defaultdict(set)
    for fname, data in files:
        meta = parse_filename_meta(fname) or {}
        try:
            prs = Presentation(io.BytesIO(data))
        except Exception as e:
            slides.append(dict(error="%s: เปิดไม่ได้ (%s)" % (fname, e)))
            continue
        for si, sl in enumerate(prs.slides):
            card = extract_card(sl)
            imgs = []
            for blob in slide_images(sl):
                h = hashlib.md5(blob).hexdigest()
                hash_slides[h].add((fname, si))
                imgs.append((h, blob))
            slides.append(dict(file=fname, idx=si, meta=meta, card=card, imgs=imgs))
    warnings, groups, order = [], {}, []
    for s in slides:
        if "error" in s:
            warnings.append(s["error"])
            continue
        c = s["card"]
        tag = c["tag"].strip()
        tt = title_tag(c["title"])
        if tag and tt and norm_key(tt) != norm_key(tag):
            if norm_key(tt) in re.sub(r"\s+", "", c["background"].upper()):
                warnings.append("%s สไลด์ %d: ช่อง Equipment = %s แต่หัวข้อ/Background เป็น %s → ใช้ %s"
                                % (s["file"], s["idx"] + 1, tag, tt, tt))
                tag = tt
            else:
                warnings.append("%s สไลด์ %d: ช่อง Equipment = %s แต่หัวข้อสไลด์เขียน %s → ใช้ %s (ตรวจสอบ)"
                                % (s["file"], s["idx"] + 1, tag, tt, tag))
        if not tag and tt:
            tag = tt
        if not tag:
            txt = " ".join(x for x in (c["title"], c["background"]) if x)[:60]
            warnings.append("%s สไลด์ %d: หา Tag ไม่เจอ (ข้าม) — %s" % (s["file"], s["idx"] + 1, txt))
            continue
        k = norm_key(tag)
        if k not in groups:
            groups[k] = dict(variants=[], plants=[], bg=[], cause=[], action=[], dates=[], raw_dates=[],
                             imgs=[], sources=[], last_action=[], weeks=[])
            order.append(k)
        g = groups[k]
        g["variants"].append(tag)
        if c["plant"] and c["plant"] not in g["plants"]:
            g["plants"].append(c["plant"])
        g["bg"] += clean_lines(c["background"])
        g["cause"] += clean_lines(c["cause"])
        g["action"] += clean_lines(c["action"])
        g["last_action"] = clean_lines(c["action"]) or g["last_action"]
        a, b = parse_date_text(c["date_raw"])
        if a:
            g["dates"].append((a, b))
        elif c["date_raw"]:
            g["raw_dates"].append(c["date_raw"])
        g["imgs"] += s["imgs"]
        g["sources"].append("%s (สไลด์ %d)" % (s["file"], s["idx"] + 1))
        w = s["meta"].get("week")
        if w and w not in g["weeks"]:
            g["weeks"].append(w)

    items = []
    for k in order:
        g = groups[k]
        seen, cands = set(), []
        for h, blob in g["imgs"]:
            if h in seen or len(hash_slides[h]) >= 4:      # ซ้ำ / โลโก้ที่โผล่หลายสไลด์
                continue
            seen.add(h)
            try:
                im = Image.open(io.BytesIO(blob))
                w, hh = im.size
                if w < 250 or hh < 250:
                    continue
                cands.append(dict(data=blob, thumb=_thumb(im), drawing=_drawing_score(im) > 0.55))
            except Exception:
                continue
        cands = cands[:80]
        photo_idx = [i for i, c in enumerate(cands) if not c["drawing"]]
        draw_idx = [i for i, c in enumerate(cands) if c["drawing"]]
        pick = pick_evenly(photo_idx, 4)
        if len(pick) < 4:
            pick += draw_idx[:4 - len(pick)]
        bg = dedup(g["bg"])
        cause = dedup(g["cause"])
        action = dedup(g["action"])
        if g["dates"]:
            date_text = fmt_range(min(a for a, b in g["dates"]), max(b for a, b in g["dates"]))
        else:
            date_text = " / ".join(dict.fromkeys(g["raw_dates"]))
        handed = [l for l in g["last_action"] if re.search(r"ส่งมอบ|hand\s*over|turn\s*over", l, re.I)]
        if handed:
            done, status = True, handed[-1]
        else:
            done = False
            status = (g["last_action"][-1] if g["last_action"] else "")
        status = status[:120]
        plants = ", ".join(g["plants"])
        items.append(dict(
            key=k, tag=display_tag(g["variants"]), plant=plants, dates=date_text,
            sub=(bg[0][:90] if bg else ""), bg=bg, cause=cause, action=action,
            status=status, done=done, cands=cands, pick=pick, sources=g["sources"], weeks=g["weeks"]))
    items.sort(key=lambda it: it["plant"])        # OLE2-1 -> OLE2-2 -> OLE2-3 (คงลำดับเดิมในกลุ่มเดียวกัน)
    return items, warnings


# ============================================================
# ส่วนที่ 2: Excel Job Summary
# ============================================================
def _order_no(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str) and re.fullmatch(r"\d{6,}", v.strip()):
        return int(v.strip())
    return None


def read_excel_events(data):
    """คืน {sheet: [event,...]} — เฉพาะ sheet ที่มีหัวคอลัมน์ Order และ Status"""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    out = {}
    for ws in wb.worksheets:
        hdr = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]

        def col(*names):
            for n in names:
                if n in hdr:
                    return hdr.index(n)
            return None
        ci_date, ci_ord, ci_st = col("date"), col("order"), col("status")
        ci_eq, ci_de = col("equipment"), col("description")
        if ci_ord is None or ci_st is None:
            continue
        cur, ev = None, []
        for r in ws.iter_rows(min_row=2, values_only=True):
            r = list(r) + [None] * (len(hdr) - len(r))
            if ci_date is not None and isinstance(r[ci_date], (datetime.datetime, datetime.date)):
                d = r[ci_date]
                cur = d.date() if isinstance(d, datetime.datetime) else d
            st_ = r[ci_st]
            ev.append(dict(date=cur, order=_order_no(r[ci_ord]), status=(str(st_).strip().upper() if st_ not in (None, "") else None),
                           eq=(r[ci_eq] if ci_eq is not None else None),
                           desc=(r[ci_de] if ci_de is not None else None),
                           has_text=any(x not in (None, "") for i, x in enumerate(r) if i != ci_date)))
        out[ws.title] = ev
    return out


def split_words(s):
    return {w.strip().upper() for w in re.split(r"[,\n]", s or "") if w.strip()}


def summarize_jobs(events, finished, cont, ready):
    """Order เลขเดียวกันหลายวัน = งานเดียวต่อเนื่อง จนกว่าจะเจอสถานะเสร็จ (เจอแล้วปิดงาน ถ้ามีต่ออีก = งานใหม่)"""
    by = collections.OrderedDict()
    skipped = 0
    for e in events:
        if e["order"] is not None:
            by.setdefault(e["order"], []).append(e)
        elif e["has_text"] and (e["eq"] or e["desc"]):
            skipped += 1
    jobs, reopened = [], 0
    for o, rows in by.items():
        days = collections.OrderedDict()
        for r in rows:
            v = days.setdefault(r["date"], dict(st=None, eq=r["eq"], desc=r["desc"]))
            if r["status"] in finished:
                v["st"] = r["status"]
            elif r["status"] and v["st"] not in finished:
                v["st"] = r["status"]
            v["eq"], v["desc"] = r["eq"] or v["eq"], r["desc"] or v["desc"]
        segs, seg = [], []
        for d, v in days.items():
            seg.append((d, v))
            if v["st"] in finished:
                segs.append((seg, True))
                seg = []
        if seg:
            segs.append((seg, False))
        if len(segs) > 1:
            reopened += 1
        for sg, fin in segs:
            sts = [v["st"] for d, v in sg if v["st"]]
            last = sts[-1] if sts else None
            if fin:
                cat = "done"
            elif last in cont:
                cat = "continue"
            elif last in ready:
                cat = "ready"
            else:
                cat = "other"
            jobs.append(dict(order=o, start=sg[0][0], end=sg[-1][0], days=len(sg), eq=sg[-1][1]["eq"],
                             desc=sg[-1][1]["desc"], last=last, cat=cat))
    cnt = collections.Counter(j["cat"] for j in jobs)
    other_detail = sorted({(j["last"] or "ไม่ระบุ") for j in jobs if j["cat"] == "other"},
                          key=lambda x: (x != "ไม่ระบุ", x))
    return dict(total=len(jobs), done=cnt["done"], cont=cnt["continue"], ready=cnt["ready"], other=cnt["other"],
                other_detail=other_detail, distinct=len(by), reopened=reopened, skipped=skipped,
                multi=sum(1 for j in jobs if j["days"] >= 2), jobs=jobs)


# ============================================================
# ส่วนที่ 3: ตาราง CAPEX
# ============================================================
def _clean_cell(t):
    t = (t or "").replace("\xa0", " ").replace("\u200b", "").replace("\x0b", "\n")
    lines = [l.strip() for l in t.split("\n")]
    while lines and not lines[-1]:
        lines.pop()
    while lines and not lines[0]:
        lines.pop(0)
    return "\n".join(lines)


def _cell_color(cell):
    cols = []
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                try:
                    if r.font.color and r.font.color.type == 1:
                        cols.append(str(r.font.color.rgb))
                    else:
                        cols.append(None)
                except Exception:
                    cols.append(None)
    return max(set(cols), key=cols.count) if cols else None


def parse_capex(data):
    prs = Presentation(io.BytesIO(data))
    tables = []
    for si, sl in enumerate(prs.slides):
        title = ""
        for sh in sl.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() and not sh.has_table:
                t = sh.text_frame.text.strip()
                if "capex" in t.lower():
                    title = t
                    break
                title = title or t
        for sh in sl.shapes:
            if sh.has_table:
                tb = sh.table
                hdr = [_clean_cell(c.text) for c in tb.rows[0].cells]
                rows = []
                for r in list(tb.rows)[1:]:
                    cells = [_clean_cell(c.text) for c in r.cells]
                    if not any(cells):
                        continue
                    rows.append(dict(cells=cells, colors=[_cell_color(c) for c in r.cells]))
                if rows:
                    tables.append(dict(title=title or "CAPEX", header=hdr, rows=rows))
    return tables


# ============================================================
# ส่วนที่ 4: ตัวสร้าง PowerPoint (ใช้ template.pptx ที่มี layout DARK / CONTENT)
# ============================================================
def rgb(h):
    return RGBColor.from_string(h)


def style_run(run, size, bold=False, color=None, spc=None):
    f = run.font
    f.size = Pt(size)
    f.bold = bool(bold)
    f.name = FONT
    f.color.rgb = rgb(color or COL["ink"])
    rPr = run._r.get_or_add_rPr()
    for tg in ("a:ea", "a:cs"):
        el = rPr.find(qn(tg))
        if el is None:
            el = etree.SubElement(rPr, qn(tg))
        el.set("typeface", FONT)
    if spc:
        rPr.set("spc", str(int(spc * 100)))


def set_bullet(p, indent=0.19):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent))))
    pPr.set("indent", str(-int(Inches(indent))))
    for o in pPr.findall(qn("a:buNone")):
        pPr.remove(o)
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", "•")


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def fill_tf(tf, paras, size=14, bold=False, color=None, align="l", spc=None, bullet=False, space_after=0):
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[align]
        if space_after:
            para.space_after = Pt(space_after)
        runs = [(p, {})] if isinstance(p, str) else p
        for t, o in runs:
            r = para.add_run()
            r.text = t
            style_run(r, o.get("size", size), o.get("bold", bold), o.get("color", color), o.get("spc", spc))
        if bullet:
            set_bullet(para)


def add_text(slide, x, y, w, h, paras, size=14, bold=False, color=None, align="l", anchor="t",
             spc=None, bullet=False, space_after=0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ANCH[anchor]
    fill_tf(tf, paras, size, bold, color, align, spc, bullet, space_after)
    return tb


def add_card(slide, x, y, w, h, fill, radius=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.adjustments[0] = min(0.5, radius / float(min(w, h)))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_dot(slide, x, y, d, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False


def find_layout(prs, name):
    for m in prs.slide_masters:
        for l in m.slide_layouts:
            if l.name == name:
                return l
    raise ValueError("ไม่พบ layout ชื่อ '%s' ใน template" % name)


def clear_slides(prs):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        prs.part.drop_rel(sid.rId)
        lst.remove(sid)


def crop_to(data, w, h, max_side=1000):
    im = Image.open(io.BytesIO(data)).convert("RGB")
    ratio = w / float(h)
    iw, ih = im.size
    if iw / float(ih) > ratio:
        nw = int(ih * ratio)
        x = (iw - nw) // 2
        im = im.crop((x, 0, x + nw, ih))
    else:
        nh = int(iw / ratio)
        y = (ih - nh) // 2
        im = im.crop((0, y, iw, y + nh))
    im.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    buf.seek(0)
    return buf


def est_height(paras, size, width_in, bullet):
    w = width_in - (0.22 if bullet else 0)
    cw = size * 0.0080
    lh = size * 1.28 / 72.0
    lines = sum(max(1, math.ceil(len(p) * cw / w)) for p in paras)
    return lines * lh + len(paras) * (3 / 72.0)


LEFT_X, LEFT_W, TOP, AVAIL, GAP = 0.5, 6.95, 1.5, 5.35, 0.15
RIGHT_X, RIGHT_W = 7.75, 5.08


def split_two(paras, size, w2):
    """แบ่งรายการเป็น 2 คอลัมน์ให้ความสูงสมดุลที่สุด"""
    best, bh = 1, None
    for i in range(1, len(paras)):
        h = max(est_height(paras[:i], size, w2, True), est_height(paras[i:], size, w2, True))
        if bh is None or h < bh:
            best, bh = i, h
    return paras[:best], paras[best:], bh


def plan_cards(item):
    """คืน dict(size, hs, overflow, two_col) ให้การ์ดข้อความพอดีพื้นที่
    ถ้า Action ยาวมาก จะเปลี่ยนเป็น 2 คอลัมน์ก่อน แล้วค่อยลดฟอนต์"""
    secs = [item["bg"] or [""], item["cause"] or [""], item["action"] or [""]]
    bl = [len(item["bg"]) > 1, len(item["cause"]) > 1, True]
    inner = LEFT_W - 0.4
    w2 = (inner - 0.25) / 2.0
    for two in (False, True):
        if two and len(secs[2]) < 4:
            continue
        for sz in (14, 13, 12, 11.5, 11, 10.5, 10, 9.5, 9, 8.5, 8):
            if two and sz > 10:
                continue
            hs = [0.14 + 0.3 + est_height(s, sz, inner, b) + 0.14 for s, b in zip(secs[:2], bl[:2])]
            if two:
                ha = split_two(secs[2], sz, w2)[2]
            else:
                ha = est_height(secs[2], sz, inner, True)
            hs.append(0.14 + 0.3 + ha + 0.14)
            if sum(hs) + 2 * GAP <= AVAIL:
                extra = (AVAIL - sum(hs) - 2 * GAP) / 3.0
                return dict(size=sz, hs=[h + extra for h in hs], overflow=False, two_col=two)
    # ยังไม่พอ: ให้ Action ได้พื้นที่มากสุด บีบ 2 การ์ดแรกเท่าที่ต้องใช้
    need = [0.14 + 0.3 + est_height(s, 8, inner, b) + 0.14 for s, b in zip(secs[:2], bl[:2])]
    ha = AVAIL - 2 * GAP - sum(need)
    return dict(size=8, hs=need + [ha], overflow=True, two_col=len(secs[2]) >= 4)


def image_rects(n):
    cw, ch, g = 2.47, 2.0, 0.14
    top = TOP
    if n <= 0:
        return []
    if n == 1:
        return [(RIGHT_X, top, RIGHT_W, 4.15)]
    if n == 2:
        return [(RIGHT_X, top, RIGHT_W, 2.0), (RIGHT_X, top + 2.15, RIGHT_W, 2.0)]
    if n == 3:
        return [(RIGHT_X, top, cw, ch), (RIGHT_X + cw + g, top, cw, ch), (RIGHT_X, top + 2.15, RIGHT_W, 2.0)]
    return [(RIGHT_X + (i % 2) * (cw + g), top + (i // 2) * 2.15, cw, ch) for i in range(4)]


def slide_equipment(prs, item, images):
    s = prs.slides.add_slide(find_layout(prs, "CONTENT"))
    add_text(s, 0.5, 0.32, 7.5, 0.7, item["tag"], size=34, bold=True, color=COL["dark"])
    add_text(s, 0.5, 1.0, 7.9, 0.4, item["sub"], size=14, color=COL["muted"])
    add_text(s, 8.83, 0.38, 4.0, 0.4, item["plant"], size=18, bold=True, color=COL["teal"], align="r")
    add_text(s, 8.83, 0.82, 4.0, 0.35, item["dates"], size=14, color=COL["muted"], align="r")
    pl = plan_cards(item)
    size, hs, overflow = pl["size"], pl["hs"], pl["overflow"]
    secs = [("BACKGROUND", item["bg"], len(item["bg"]) > 1), ("POSSIBLE CAUSE", item["cause"], len(item["cause"]) > 1),
            ("ACTION", item["action"], True)]
    y = TOP
    for idx, ((label, paras, bl), h) in enumerate(zip(secs, hs)):
        add_card(s, LEFT_X, y, LEFT_W, h, COL["card"])
        add_text(s, LEFT_X + 0.2, y + 0.12, LEFT_W - 0.4, 0.26, label, size=11, bold=True, color=COL["teal"], spc=2)
        if idx == 2 and pl["two_col"] and len(paras) >= 4:
            w2 = (LEFT_W - 0.4 - 0.25) / 2.0
            left, right, _ = split_two(paras, size, w2)
            add_text(s, LEFT_X + 0.2, y + 0.42, w2, h - 0.52, left, size=size, color=COL["ink"], bullet=True, space_after=3)
            add_text(s, LEFT_X + 0.2 + w2 + 0.25, y + 0.42, w2, h - 0.52, right, size=size, color=COL["ink"], bullet=True, space_after=3)
        else:
            add_text(s, LEFT_X + 0.2, y + 0.42, LEFT_W - 0.4, h - 0.52, paras or ["-"], size=size, color=COL["ink"],
                     bullet=bl, space_after=3)
        y += h + GAP
    for (x, yy, w, h), data in zip(image_rects(len(images)), images):
        s.shapes.add_picture(crop_to(data, w, h), Inches(x), Inches(yy), Inches(w), Inches(h))
    sy = TOP + 2 * 2.0 + 0.15 + 0.15
    sh = TOP + AVAIL - sy
    col = COL["green"] if item["done"] else COL["orange"]
    add_card(s, RIGHT_X, sy, RIGHT_W, sh, COL["greenbg"] if item["done"] else COL["orangebg"])
    add_text(s, RIGHT_X + 0.2, sy + 0.1, 4.6, 0.3, "เสร็จสิ้น / ส่งมอบแล้ว" if item["done"] else "อยู่ระหว่างดำเนินการ",
             size=14, bold=True, color=col)
    add_text(s, RIGHT_X + 0.2, sy + 0.44, RIGHT_W - 0.4, sh - 0.5, item["status"] or "-", size=11.5, color=COL["ink"])
    s.notes_slide.notes_text_frame.text = "แหล่งข้อมูล: " + "; ".join(item.get("sources", []))
    return overflow


# ---------- ตาราง (ใช้ทั้ง CAPEX และ Job Summary) ----------
def _border(tcPr, color=None, w=6350):
    for i, tg in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = etree.Element(qn(tg))
        ln.set("w", str(w))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr"))
        c.set("val", color or COL["line"])
        tcPr.insert(i, ln)


def add_table(slide, x, y, colw, rowh, data, fs):
    """data[r][c] = dict(text, size, bold, color, fill, align)"""
    nr, nc = len(data), len(data[0])
    gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(sum(rowh)))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for a in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
        tblPr.set(a, "0")
    for sid in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(sid)
    for i, w in enumerate(colw):
        tbl.columns[i].width = Inches(w)
    for i, h in enumerate(rowh):
        tbl.rows[i].height = Inches(h)
    for r in range(nr):
        for c in range(nc):
            d = data[r][c]
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            lines = str(d.get("text", "")).split("\n") or [""]
            fill_tf(tf, lines, size=d.get("size", fs), bold=d.get("bold", False), color=d.get("color"),
                    align=d.get("align", "c"))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(d.get("fill", "FFFFFF"))
            _border(cell._tc.get_or_add_tcPr())
    return gf


def _lines_needed(text, size, w_in):
    cw = size * 0.0082
    per = max(1, int(w_in / cw))
    return sum(max(1, math.ceil(len(l) / float(per))) for l in text.split("\n"))


def plan_table(tbl, total_w=12.5, pad=0.14):
    hdr, rows = tbl["header"], tbl["rows"]
    for size in (9.5, 9, 8.5, 8, 7.5):
        cw = size * 0.0082
        minw, need = [], []
        for i, h in enumerate(hdr):
            w = max([len(t) for t in h.split()] or [1]) * 1.1
            m = 0
            for r in rows:
                for l in r["cells"][i].split("\n"):
                    for t in l.split():
                        w = max(w, len(t))
                    m = max(m, len(l))
            minw.append(min(max(w * cw + pad + 0.06, 0.55), 1.6))
            need.append(min(m, 70))
        left = max(0.0, total_w - sum(minw))
        wts = [max(0.0, n * cw - mw + pad) + 0.0001 for n, mw in zip(need, minw)]
        ws = sum(wts)
        colw = [mw + left * wt / ws for mw, wt in zip(minw, wts)]
        sc = total_w / sum(colw)
        colw = [c * sc for c in colw]
        lh = size * 1.28 / 72.0
        rowh = [max(_lines_needed(c, size, colw[i] - pad) for i, c in enumerate(r["cells"])) * lh + 0.13 for r in rows]
        hdrh = size * 1.25 / 72.0 + 0.18
        if hdrh + sum(rowh) <= 5.6 or size == 7.5:
            return size, colw, hdrh, rowh


def slide_capex(prs, tbl, page_no, page_tot):
    s = prs.slides.add_slide(find_layout(prs, "CONTENT"))
    add_text(s, 0.4, 0.25, 10.5, 0.55, tbl["title"], size=26, bold=True, color=COL["dark"])
    if page_tot > 1:
        add_text(s, 11.4, 0.33, 1.5, 0.4, "%d / %d" % (page_no, page_tot), size=14, color=COL["muted"], align="r")
    size, colw, hdrh, rowh = plan_table(tbl)
    n = len(tbl["header"])
    data = [[dict(text=h, size=size + 0.5, bold=True, color=COL["white"], fill=COL["dark"],
                  align="l" if i in (2, n - 1) else "c") for i, h in enumerate(tbl["header"])]]
    for ri, r in enumerate(tbl["rows"]):
        fill = "FFFFFF" if ri % 2 == 0 else COL["card"]
        row = []
        for i, c in enumerate(r["cells"]):
            colr = r["colors"][i]
            color = COL["green"] if colr == "00B050" else (COL["orange"] if colr == "FF0000" else COL["ink"])
            row.append(dict(text=c, size=size, bold=(i == 1), color=color, fill=fill,
                            align="l" if i in (2, n - 1) else "c"))
        data.append(row)
    add_table(s, 0.4, 0.98, colw, [hdrh] + rowh, data, size)
    add_dot(s, 0.4, 7.14, 0.13, COL["green"])
    add_text(s, 0.6, 7.08, 1.6, 0.25, "ปกติ / เสร็จ", size=10, color=COL["muted"])
    add_dot(s, 2.3, 7.14, 0.13, COL["orange"])
    add_text(s, 2.5, 7.08, 2.2, 0.25, "รอ / ต้องติดตาม", size=10, color=COL["muted"])


# ---------- หน้าปก / Agenda / คั่น / ปิดท้าย ----------
def slide_cover(prs, code, sub):
    s = prs.slides.add_slide(find_layout(prs, "DARK"))
    add_text(s, 0.7, 1.0, 6.3, 0.4, "MONTHLY MAINTENANCE HIGHLIGHT", size=13, bold=True, color=COL["amber"], spc=3)
    add_text(s, 0.7, 1.5, 6.4, 1.0, code, size=44, bold=True, color=COL["white"])
    add_text(s, 0.7, 3.15, 6.4, 0.5, sub, size=18, color=COL["soft"])


def slide_agenda(prs, entries):
    s = prs.slides.add_slide(find_layout(prs, "CONTENT"))
    add_text(s, 0.7, 0.6, 6, 0.8, "Agenda", size=40, bold=True, color=COL["dark"])
    for i, (t, sub) in enumerate(entries):
        y = 2.0 + i * 1.6
        add_card(s, 0.7, y, 8.5, 1.3, COL["card"])
        add_text(s, 1.0, y + 0.2, 8.0, 0.5, t, size=24, bold=True, color=COL["dark"])
        add_text(s, 1.0, y + 0.75, 8.0, 0.35, sub, size=14, color=COL["muted"])


def slide_divider(prs, code, sub, n_all, n_done, n_open, footer):
    s = prs.slides.add_slide(find_layout(prs, "DARK"))
    add_text(s, 0.7, 1.0, 6.3, 0.4, "MAJOR MAINTENANCE HIGHLIGHT", size=13, bold=True, color=COL["amber"], spc=3)
    add_text(s, 0.7, 1.5, 6.4, 1.0, code, size=44, bold=True, color=COL["white"])
    add_text(s, 0.7, 3.15, 6.4, 0.5, sub, size=18, color=COL["soft"])
    for i, (num, lab) in enumerate([(n_all, "อุปกรณ์"), (n_done, "เสร็จสิ้น / ส่งมอบแล้ว"), (n_open, "อยู่ระหว่างดำเนินการ")]):
        x = 0.7 + i * 2.15
        add_text(s, x, 4.0, 2.0, 0.95, str(num), size=54, bold=True, color=COL["amber"])
        add_text(s, x, 5.0, 2.0, 0.6, lab, size=13, color=COL["white"])
    add_text(s, 0.7, 6.7, 6.4, 0.4, footer, size=11, color=COL["faint"])


def slide_thanks(prs, code):
    s = prs.slides.add_slide(find_layout(prs, "DARK"))
    add_text(s, 0.7, 2.6, 6.4, 1.0, "Thank You", size=44, bold=True, color=COL["white"])
    add_text(s, 0.7, 3.7, 6.4, 0.5, code, size=18, color=COL["amber"])


# ---------- Monthly Job Summary ----------
def slide_jobs(prs, units, date_text, month_label):
    """units: list[(ชื่อ sheet, summary dict)]"""
    s = prs.slides.add_slide(find_layout(prs, "CONTENT"))
    T = {k: sum(u[k] for _, u in units) for k in ("total", "done", "cont", "ready", "other", "distinct", "skipped", "reopened")}
    names = " · ".join(n for n, _ in units)
    add_text(s, 0.5, 0.3, 8.5, 0.7, "Monthly Job Summary", size=34, bold=True, color=COL["dark"])
    add_text(s, 0.5, 1.0, 8.5, 0.35, "%s  —  จำนวน Job ตามสถานะ" % names, size=14, color=COL["muted"])
    add_text(s, 8.83, 0.38, 4.0, 0.4, month_label, size=18, bold=True, color=COL["teal"], align="r")
    add_text(s, 8.83, 0.82, 4.0, 0.35, date_text, size=14, color=COL["muted"], align="r")
    cy, chh = 1.55, 1.65
    pct = lambda a, b: int(round(100.0 * a / b)) if b else 0
    add_card(s, 0.5, cy, 3.5, chh, COL["dark"])
    add_text(s, 0.75, cy + 0.15, 3.0, 0.28, "TOTAL JOBS", size=11, bold=True, color=COL["amber"], spc=3)
    add_text(s, 0.75, cy + 0.42, 3.0, 0.85, str(T["total"]), size=54, bold=True, color=COL["amber"])
    add_text(s, 0.75, cy + 1.24, 3.0, 0.3, "เสร็จแล้ว %d งาน (%d%%)" % (T["done"], pct(T["done"], T["total"])),
             size=13, color=COL["white"])
    n = len(units)
    avail, gap = 8.63, 0.2
    w = (avail - gap * (n - 1)) / float(n)
    for i, (name, u) in enumerate(units):
        x = 4.2 + i * (w + gap)
        add_card(s, x, cy, w, chh, COL["card"])
        add_text(s, x + 0.25, cy + 0.15, 1.8, 0.35, name, size=20, bold=True, color=COL["teal"])
        add_text(s, x + 0.25, cy + 0.5, 1.9, 0.8, str(u["total"]), size=44, bold=True, color=COL["dark"])
        if w >= 3.6:
            add_text(s, x + 2.15, cy + 0.45, w - 2.3, 0.9,
                     [[("เสร็จแล้ว ", dict(color=COL["muted"])), ("%d%%" % pct(u["done"], u["total"]), dict(bold=True, color=COL["green"]))],
                      [("ทำต่อเนื่องหลายวัน ", dict(color=COL["muted"])), ("%d งาน" % u["multi"], dict(bold=True, color=COL["ink"]))]],
                     size=12, space_after=4)
        add_text(s, x + 0.25, cy + 1.28, w - 0.5, 0.28, "Order ไม่ซ้ำ %d รายการ" % u["distinct"], size=11, color=COL["muted"])

    add_text(s, 0.5, 3.42, 6.4, 0.3, "สัดส่วนสถานะ Job แยกตามหน่วยงาน", size=14, bold=True, color=COL["dark"])
    cd = CategoryChartData()
    cd.categories = [nm for nm, _ in units]
    series = [("เสร็จแล้ว", "done", COL["green"]), ("ทำงานต่อเนื่อง", "cont", COL["orange"]),
              ("พร้อมทำงาน", "ready", COL["teal"]), ("อื่น ๆ / ไม่ระบุ", "other", COL["other"])]
    for nm, k, _ in series:
        cd.add_series(nm, [u[k] for _, u in units])
    gfr = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.4), Inches(3.75), Inches(6.6), Inches(2.6), cd)
    ch = gfr.chart
    ch.font.name = FONT
    ch.font.size = Pt(11)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(11)
    ch.legend.font.name = FONT
    plot = ch.plots[0]
    plot.gap_width = 45
    plot.overlap = 100
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = "#,##0;;"
    dl.number_format_is_linked = False
    dl.font.size = Pt(11)
    dl.font.bold = True
    dl.font.name = FONT
    dl.font.color.rgb = rgb(COL["white"])
    dl.position = XL_LABEL_POSITION.CENTER
    for ser, (_, _, colr) in zip(plot.series, series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(colr)
    ca, va = ch.category_axis, ch.value_axis
    ca.reverse_order = True
    ca.tick_labels.font.size = Pt(13)
    ca.tick_labels.font.bold = True
    ca.tick_labels.font.name = FONT
    ca.has_major_gridlines = False
    ca.format.line.fill.background()
    va.visible = False
    va.has_major_gridlines = False

    other_txt = " / ".join(dict.fromkeys(x for _, u in units for x in u["other_detail"])) or "-"
    add_text(s, 7.2, 3.42, 5.6, 0.3, "สรุปจำนวน Job แยกสถานะ", size=14, bold=True, color=COL["dark"])
    hd = lambda t, a="c": dict(text=t, size=11, bold=True, color=COL["white"], fill=COL["dark"], align=a)
    data = [[hd("สถานะ", "l"), hd("ค่าใน Excel", "l")] + [hd(nm) for nm, _ in units] + [hd("รวม")]]
    defs = [("เสร็จแล้ว", "WF / COMPLETE", "done", COL["green"]), ("ทำงานต่อเนื่อง", "Continue", "cont", COL["orange"]),
            ("พร้อมทำงาน", "REDY", "ready", COL["teal"]), ("อื่น ๆ", other_txt, "other", COL["other"])]
    for i, (lab, xl, k, colr) in enumerate(defs):
        fill = "FFFFFF" if i % 2 == 0 else COL["card"]
        row = [dict(text=lab, size=11, bold=True, color=colr, fill=fill, align="l"),
               dict(text=xl, size=10, color=COL["muted"], fill=fill, align="l")]
        for _, u in units:
            row.append(dict(text=("–" if u[k] == 0 else str(u[k])), size=11, fill=fill))
        row.append(dict(text=str(T[k]), size=11, bold=True, fill=fill))
        data.append(row)
    tf = COL["tot"]
    data.append([dict(text="รวม Job", size=11, bold=True, fill=tf, align="l"), dict(text="", size=11, fill=tf)] +
                [dict(text=str(u["total"]), size=11, bold=True, fill=tf) for _, u in units] +
                [dict(text=str(T["total"]), size=11, bold=True, fill=tf)])
    tw = 5.63
    ncol_units = len(units)
    cw = [1.4, 1.95] + [0.75] * ncol_units + [0.78]
    scale = tw / sum(cw)
    cw = [c * scale for c in cw]
    add_table(s, 7.2, 3.8, cw, [0.36, 0.42, 0.42, 0.42, 0.55, 0.42], data, 11)

    skipped = ", ".join("%s %d" % (nm, u["skipped"]) for nm, u in units)
    reopened = ", ".join("%s %d" % (nm, u["reopened"]) for nm, u in units)
    add_text(s, 0.5, 6.62, 11.2, 0.5,
             [[("หลักการนับ: ", dict(bold=True)),
               ("นับเฉพาะบรรทัดที่มีเลข Order (ไม่นับบรรทัดที่ไม่มีเลข Order: %s) · Order ซ้ำหลายวัน = 1 Job จนกว่าจะพบสถานะเสร็จ "
                "(Order ที่กลับมาทำต่อหลังเสร็จ นับเป็น Job ใหม่: %s)" % (skipped, reopened), {})]],
             size=9.5, color=COL["muted"])


# ---------- ประกอบทั้งไฟล์ ----------
def build_deck(template_bytes, cfg, items, item_images, jobs_units, capex_tables):
    """
    cfg: dict(code, month_label, plants, footer, date_text)
    items: รายการอุปกรณ์ (แก้ไขแล้ว) ; item_images: list[list[bytes]] คู่กับ items
    jobs_units: list[(name, summary)] หรือ [] ; capex_tables: list[table] หรือ []
    คืน (bytes, warnings)
    """
    prs = Presentation(io.BytesIO(template_bytes))
    clear_slides(prs)
    warns = []
    sub = "%s  ·  %s" % (cfg["month_label"], cfg["plants"])
    agenda = []
    if jobs_units:
        agenda.append(("Monthly Job Summary", "%s team" % " and ".join(n for n, _ in jobs_units)))
    if items:
        agenda.append(("Equipment Highlight", "%d อุปกรณ์  ·  %s" % (len(items), cfg["plants"])))
    if capex_tables:
        titles = []
        for t in capex_tables:
            m = re.split(r"O-MN\d-MM\d", t["title"])
            tail = m[-1].strip() if len(m) > 1 else ""
            if tail and tail not in titles:
                titles.append(tail)
        agenda.append(("OLE 2 : CAPEX Status", "  ·  ".join(titles) if titles else "Outstanding  ·  New"))
    slide_cover(prs, cfg["code"], sub)
    slide_agenda(prs, agenda)
    if jobs_units:
        slide_jobs(prs, jobs_units, cfg.get("date_text", ""), cfg["month_label"])
    if items:
        n_done = sum(1 for it in items if it["done"])
        slide_divider(prs, cfg["code"], sub, len(items), n_done, len(items) - n_done, cfg["footer"])
        for it, imgs in zip(items, item_images):
            if slide_equipment(prs, it, imgs):
                warns.append("%s: ข้อความยาวเกินการ์ด (ลดฟอนต์ถึงขั้นต่ำแล้ว) — ควรตัดข้อความหรือเปิดตรวจใน PowerPoint" % it["tag"])
    if capex_tables:
        cnt = collections.Counter(t["title"] for t in capex_tables)
        seen = collections.Counter()
        for t in capex_tables:
            seen[t["title"]] += 1
            slide_capex(prs, t, seen[t["title"]], cnt[t["title"]])
    slide_thanks(prs, cfg["code"])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue(), warns


# ============================================================
# AI สรุปข้อความ (ทางเลือก — ส่งเฉพาะข้อความ ไม่ส่งรูป)
# ============================================================
MODEL_ID = "claude-haiku-4-5"


def ai_summarize(item, api_key):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    raw = ("Background:\n- " + "\n- ".join(item["bg"]) + "\n\nPossible cause:\n- " + "\n- ".join(item["cause"]) +
           "\n\nAction (รวมทุกสัปดาห์ เรียงตามเวลา):\n- " + "\n- ".join(item["action"]))
    prompt = (
        "คุณเป็นวิศวกรบำรุงรักษาโรงงานปิโตรเคมี กำลังสรุปงานของอุปกรณ์เดียวกันที่ทำต่อเนื่องหลายสัปดาห์ ให้เป็นสไลด์เดียว\n"
        "อุปกรณ์: %s (Plant: %s) ช่วงเวลา: %s\n\nข้อมูลดิบจากสไลด์รายสัปดาห์:\n%s\n\n"
        "สรุปเป็นภาษาไทย (คงศัพท์เทคนิคภาษาอังกฤษ ชื่อ Tag เลขที่ วันที่ ชื่อบริษัทตามต้นฉบับ) กระชับ อ่านง่าย โดย\n"
        "- sub: ชื่อหัวข้อสั้น ๆ รูปแบบ 'ประเภทอุปกรณ์ — หัวข้องาน' ไม่เกิน 80 ตัวอักษร\n"
        "- bg: 1 ข้อความ (ไม่เกิน 130 ตัวอักษร) เล่าว่าเกิดอะไร/งานคืออะไร\n"
        "- cause: 1-2 ข้อ (ข้อละไม่เกิน 170 ตัวอักษร) สาเหตุหรือสิ่งที่ตรวจพบ\n"
        "- action: 3-4 ข้อ (ข้อละไม่เกิน 180 ตัวอักษร) สิ่งที่ดำเนินการรวมทุกสัปดาห์เรียงตามลำดับเวลา\n"
        "- status: สถานะล่าสุด 1 บรรทัด (ไม่เกิน 100 ตัวอักษร)\n"
        "- done: true ถ้ามีการส่งมอบ/ปิดงานแล้ว ไม่งั้น false\n"
        "ห้ามใส่ข้อมูลที่ไม่มีในต้นฉบับ ห้ามเดาสาเหตุ\n"
        'ตอบเป็น JSON เท่านั้น: {"sub":"","bg":[""],"cause":[""],"action":[""],"status":"","done":false}'
        % (item["tag"], item["plant"], item["dates"], raw[:6000]))
    resp = client.messages.create(model=MODEL_ID, max_tokens=900, messages=[{"role": "user", "content": prompt}])
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    return json.loads(text)


# ============================================================
# ส่วนที่ 5: หน้าเว็บ Streamlit
# ============================================================
def main():
    import streamlit as st

    st.set_page_config(page_title="Monthly Maintenance Highlight Builder", layout="wide")

    # ----- รหัสผ่าน -----
    def check_password():
        def entered():
            try:
                correct = st.secrets.get("APP_PASSWORD", "")
            except Exception:
                correct = ""
            ok = bool(correct) and st.session_state.get("pw_input", "") == correct
            st.session_state["pw_ok"] = ok
            st.session_state.pop("pw_input", None)
        if st.session_state.get("pw_ok"):
            return True
        st.title("🔒 Monthly Maintenance Highlight Builder")
        st.text_input("รหัสผ่าน", type="password", on_change=entered, key="pw_input")
        if st.session_state.get("pw_ok") is False:
            st.error("รหัสผ่านไม่ถูกต้อง")
        return False

    if not check_password():
        st.stop()

    ss = st.session_state
    st.title("📋 Monthly Maintenance Highlight Builder")
    st.caption("อัปโหลด Weekly Report + Excel Job + ไฟล์ CAPEX → ตรวจทาน → ดาวน์โหลด PowerPoint (theme เดียวกับรายงานเดือนก่อน)")

    # ----- 1) อัปโหลด -----
    c1, c2, c3 = st.columns(3)
    with c1:
        weekly = st.file_uploader("① Weekly Report (.pptx) หลายไฟล์", type=["pptx"], accept_multiple_files=True, key="up_weekly")
    with c2:
        xl = st.file_uploader("② Excel งานประจำเดือน (.xlsx)", type=["xlsx"], key="up_xl")
    with c3:
        cap = st.file_uploader("③ ไฟล์ CAPEX (.pptx)", type=["pptx"], key="up_cap")
    with st.expander("ตั้งค่าคำที่ใช้แปลสถานะใน Excel"):
        w_fin = st.text_input("สถานะ 'เสร็จแล้ว' (คั่นด้วย ,)", value="WF, COMPLETE", key="w_fin")
        w_con = st.text_input("สถานะ 'ทำงานต่อเนื่อง'", value="CONTINUE", key="w_con")
        w_rdy = st.text_input("สถานะ 'พร้อมทำงาน'", value="REDY, READY", key="w_rdy")

    if st.button("🔍 อ่านไฟล์ทั้งหมด", type="primary"):
        if not (weekly or xl or cap):
            st.warning("กรุณาอัปโหลดอย่างน้อย 1 ไฟล์")
        else:
            with st.spinner("กำลังอ่านไฟล์..."):
                ss["gen"] = ss.get("gen", 0) + 1
                g = ss["gen"]
                ss["items"], ss["warnings"] = [], []
                if weekly:
                    items, warns = parse_weekly([(f.name, f.getvalue()) for f in weekly])
                    ss["items"], ss["warnings"] = items, warns
                    for k, it in enumerate(items):
                        for fld in ("tag", "sub", "plant", "dates", "status"):
                            ss["%s_%d_%s" % (fld, g, k)] = it[fld]
                        for fld in ("bg", "cause", "action"):
                            ss["%s_%d_%s" % (fld, g, k)] = "\n".join(it[fld])
                        ss["done_%d_%d" % (g, k)] = "เสร็จสิ้น / ส่งมอบแล้ว" if it["done"] else "อยู่ระหว่างดำเนินการ"
                        for i in range(len(it["cands"])):
                            ss["img_%d_%d_%d" % (g, k, i)] = i in it["pick"]
                    metas = [parse_filename_meta(f.name) for f in weekly]
                    metas = [m for m in metas if m]
                    if metas:
                        ss["cfg_month"] = metas[0]["label"]
                        ws = sorted({m["week"] for m in metas})
                        us = sorted({m["unit"] for m in metas})
                        ss["cfg_footer"] = "สรุปจาก Weekly Report สัปดาห์ที่ %d – %d (หน่วยงาน %s) — 1 หน้าต่ออุปกรณ์" % (
                            ws[0], ws[-1], ", ".join(us))
                    ss["cfg_plants"] = " / ".join(sorted({p for it in items for p in [x.strip() for x in it["plant"].split(",")] if p}))
                ss["events"] = read_excel_events(xl.getvalue()) if xl else {}
                ss["capex"] = parse_capex(cap.getvalue()) if cap else []
                ss.pop("out", None)
            st.rerun()

    if "gen" not in ss:
        st.info("อัปโหลดไฟล์ด้านบน แล้วกด 'อ่านไฟล์ทั้งหมด' (ต้องมี template.pptx อยู่ใน repo เดียวกับ app.py)")
        return

    g = ss["gen"]
    # ----- แถบตั้งค่า -----
    with st.sidebar:
        st.header("⚙️ ตั้งค่ารายงาน")
        ss.setdefault("cfg_code", "O-MN2-MM1")
        ss.setdefault("cfg_month", datetime.date.today().strftime("%B %Y"))
        ss.setdefault("cfg_plants", "OLE2-1 / OLE2-2 / OLE2-3")
        ss.setdefault("cfg_footer", "")
        code = st.text_input("รหัสรายงาน (หน้าปก)", key="cfg_code")
        month = st.text_input("เดือน / ปี", key="cfg_month")
        plants = st.text_input("Plant (หน้าปก/Agenda)", key="cfg_plants")
        footer = st.text_input("ข้อความท้ายหน้าคั่น Equipment", key="cfg_footer")
        st.divider()
        st.caption("ปุ่ม 'สรุปด้วย AI' (ไม่บังคับ) ส่งเฉพาะข้อความ Background/Cause/Action ไปประมวลผล ไม่ส่งรูปภาพ")
        api_key = st.text_input("Anthropic API Key", type="password", key="cfg_api")
        if not api_key:
            try:
                api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
            except Exception:
                api_key = ""
        tpl_up = st.file_uploader("Template (ไม่ใส่ = ใช้ template.pptx ใน repo)", type=["pptx"], key="up_tpl")
        if st.button("🚪 ออกจากระบบ"):
            ss["pw_ok"] = False
            st.rerun()

    items = ss.get("items", [])
    # ----- 2) Equipment -----
    if items:
        st.divider()
        st.subheader("🔧 Equipment Highlight — ตรวจทานทีละ Tag (%d รายการ)" % len(items))
        for w in ss.get("warnings", []):
            st.warning(w)
        if st.button("🤖 สรุปด้วย AI ทั้งหมด (ข้อความเท่านั้น)", disabled=not api_key):
            prog = st.progress(0.0, text="กำลังสรุป...")
            for k, it in enumerate(items):
                cur = dict(it)
                cur.update(tag=ss["tag_%d_%d" % (g, k)], plant=ss["plant_%d_%d" % (g, k)], dates=ss["dates_%d_%d" % (g, k)])
                for fld in ("bg", "cause", "action"):
                    cur[fld] = [l.strip() for l in ss["%s_%d_%d" % (fld, g, k)].splitlines() if l.strip()]
                try:
                    r = ai_summarize(cur, api_key)
                    ss["sub_%d_%d" % (g, k)] = r.get("sub", "") or ss["sub_%d_%d" % (g, k)]
                    ss["bg_%d_%d" % (g, k)] = "\n".join(r.get("bg", []))
                    ss["cause_%d_%d" % (g, k)] = "\n".join(r.get("cause", []))
                    ss["action_%d_%d" % (g, k)] = "\n".join(r.get("action", []))
                    ss["status_%d_%d" % (g, k)] = r.get("status", "")
                    ss["done_%d_%d" % (g, k)] = "เสร็จสิ้น / ส่งมอบแล้ว" if r.get("done") else "อยู่ระหว่างดำเนินการ"
                except Exception as e:
                    st.warning("สรุป %s ไม่สำเร็จ (%s)" % (cur["tag"], e))
                prog.progress((k + 1) / float(len(items)))
            st.rerun()

        for k, it in enumerate(items):
            fk = lambda f: "%s_%d_%d" % (f, g, k)
            with st.expander("%s — %s — %s" % (ss[fk("tag")], ss[fk("plant")], ss[fk("dates")]), expanded=False):
                a, b, c, d = st.columns([1, 2, 1, 1])
                a.text_input("Tag", key=fk("tag"))
                b.text_input("หัวข้อย่อย (บรรทัดใต้ชื่อ Tag)", key=fk("sub"))
                c.text_input("Plant", key=fk("plant"))
                d.text_input("ช่วงวันที่", key=fk("dates"))
                x, y, z = st.columns(3)
                x.text_area("Background (1 บรรทัด = 1 ข้อ)", key=fk("bg"), height=140)
                y.text_area("Possible cause", key=fk("cause"), height=140)
                z.text_area("Action", key=fk("action"), height=140)
                s1, s2 = st.columns([3, 1])
                s1.text_input("สถานะล่าสุด (กล่องขวาล่าง)", key=fk("status"))
                s2.radio("สถานะงาน", ["อยู่ระหว่างดำเนินการ", "เสร็จสิ้น / ส่งมอบแล้ว"], key=fk("done"), horizontal=False)
                cur = dict(it)
                for fld in ("bg", "cause", "action"):
                    cur[fld] = [l.strip() for l in ss[fk(fld)].splitlines() if l.strip()]
                pl = plan_cards(cur)
                st.caption("ฟอนต์ที่ใช้ในสไลด์ ≈ %s pt%s%s" % (
                    pl["size"], "  (Action แบ่ง 2 คอลัมน์)" if pl["two_col"] else "",
                    "  ⚠️ ข้อความยาวเกินการ์ด ควรตัดให้สั้นลงหรือกด 'สรุปด้วย AI'" if pl["overflow"] else ""))
                st.caption("แหล่งข้อมูล: " + "; ".join(it["sources"]))
                st.markdown("**🖼️ เลือกรูป (ติ๊กได้ 1–4 รูป ระบบเลือกให้เบื้องต้น)**")
                if not it["cands"]:
                    st.caption("ไม่พบรูปที่ใช้ได้ในสไลด์ต้นฉบับ")
                cols = st.columns(6)
                for i, cd in enumerate(it["cands"]):
                    with cols[i % 6]:
                        st.image(cd["thumb"], use_container_width=True)
                        st.checkbox("ใช้" + (" (ภาพวาด)" if cd["drawing"] else ""), key="img_%d_%d_%d" % (g, k, i))
                n_sel = sum(1 for i in range(len(it["cands"])) if ss.get("img_%d_%d_%d" % (g, k, i)))
                if n_sel > 4:
                    st.warning("เลือก %d รูป — ระบบจะใช้ 4 รูปแรกที่ติ๊ก" % n_sel)

    # ----- 3) Job Summary -----
    units = []
    events = ss.get("events", {})
    if events:
        st.divider()
        st.subheader("📊 Monthly Job Summary")
        fin, con, rdy = split_words(w_fin), split_words(w_con), split_words(w_rdy)
        chosen = st.multiselect("Sheet ที่ใช้นับ Job", list(events.keys()), default=list(events.keys()))
        for nm in chosen:
            units.append((nm, summarize_jobs(events[nm], fin, con, rdy)))
        if units:
            rows = [{"Sheet": nm, "รวม Job": u["total"], "เสร็จแล้ว": u["done"], "ต่อเนื่อง": u["cont"], "พร้อมทำงาน": u["ready"],
                     "อื่น ๆ": u["other"], "Order ไม่ซ้ำ": u["distinct"], "บรรทัดไม่มี Order (ไม่นับ)": u["skipped"],
                     "Order ที่กลับมาทำต่อหลังเสร็จ": u["reopened"]} for nm, u in units]
            st.dataframe(rows, hide_index=True, use_container_width=True)
            with st.expander("ดู Job ในกลุ่ม 'อื่น ๆ' (ไม่ระบุสถานะ / CANCEL / POSTPONE ฯลฯ)"):
                for nm, u in units:
                    oth = [{"Order": j["order"], "Equipment": j["eq"], "Description": j["desc"], "สถานะล่าสุด": j["last"] or "ไม่ระบุ",
                            "วันที่": str(j["end"])} for j in u["jobs"] if j["cat"] == "other"]
                    st.caption(nm + " — %d งาน" % len(oth))
                    if oth:
                        st.dataframe(oth, hide_index=True, use_container_width=True)

    # ----- 4) CAPEX -----
    capex = ss.get("capex", [])
    use_capex = []
    if capex:
        st.divider()
        st.subheader("💰 CAPEX (ดึงตารางจากไฟล์ที่อัปโหลด)")
        for i, t in enumerate(capex):
            keep = st.checkbox("%s — %d แถว" % (t["title"], len(t["rows"])), value=True, key="cap_%d_%d" % (g, i))
            if keep:
                use_capex.append(t)
            with st.expander("ดูตาราง %d" % (i + 1)):
                st.dataframe([dict(zip(t["header"], r["cells"])) for r in t["rows"]], hide_index=True, use_container_width=True)

    # ----- 5) สร้างไฟล์ -----
    st.divider()
    st.subheader("📥 สร้างไฟล์ PowerPoint")
    if st.button("⚙️ สร้างไฟล์", type="primary"):
        try:
            if tpl_up is not None:
                tpl = tpl_up.getvalue()
            else:
                with open(os.path.join(HERE, "template.pptx"), "rb") as fh:
                    tpl = fh.read()
            final_items, imgs = [], []
            for k, it in enumerate(items):
                fk = lambda f: "%s_%d_%d" % (f, g, k)
                cur = dict(it)
                for f in ("tag", "sub", "plant", "dates", "status"):
                    cur[f] = ss[fk(f)]
                for f in ("bg", "cause", "action"):
                    cur[f] = [l.strip() for l in ss[fk(f)].splitlines() if l.strip()]
                cur["done"] = ss[fk("done")] == "เสร็จสิ้น / ส่งมอบแล้ว"
                final_items.append(cur)
                sel = [it["cands"][i]["data"] for i in range(len(it["cands"])) if ss.get("img_%d_%d_%d" % (g, k, i))]
                imgs.append(sel[:4])
            dts = [e["date"] for ev in events.values() for e in ev if e["date"]]
            date_text = ("ข้อมูล %s" % fmt_range(min(dts), max(dts))) if dts else ""
            cfg = dict(code=code, month_label=month, plants=plants, footer=footer, date_text=date_text)
            data, warns = build_deck(tpl, cfg, final_items, imgs, units, use_capex)
            ss["out"] = (data, warns)
        except Exception as e:
            st.error("สร้างไฟล์ไม่สำเร็จ: %s" % e)
    if ss.get("out"):
        data, warns = ss["out"]
        for w in warns:
            st.warning(w)
        st.success("สร้างไฟล์เรียบร้อย — เปิดตรวจใน PowerPoint อีกครั้งก่อนใช้งาน (ฟอนต์/ความยาวข้อความอาจต่างเล็กน้อยตามเครื่อง)")
        st.download_button("⬇️ ดาวน์โหลด PowerPoint", data=data,
                           file_name="%s_Monthly_Highlight_%s.pptx" % (code, month.replace(" ", "_")),
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    main()
