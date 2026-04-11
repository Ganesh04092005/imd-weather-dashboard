import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta
import re, os, tempfile, zipfile
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# NAMESPACES
# ─────────────────────────────────────────────────────────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"

def w(tag): return f"{{{W}}}{tag}"

# ─────────────────────────────────────────────────────────────────────────────
# LOOKUP TABLES
# ─────────────────────────────────────────────────────────────────────────────
SPATIAL_PHRASE = {
    "ISOL": "at isolated places",
    "SCT":  "at scattered places",
    "FWS":  "at fairly widespread places",
    "WS":   "at widespread places",
}
SPATIAL_FORECAST_WORD = {"WS": "MOST", "FWS": "MANY", "SCT": "FEW", "ISOL": "ISOLATED"}
SPATIAL_RANK = {"ISOL": 1, "SCT": 2, "FWS": 3, "WS": 4}

DAY_COLS = {
    1: {"fcst": 8,  "wrng_sp": 8,    "wrng_lv": 9,    "tslt": 11},
    2: {"fcst": 13, "wrng_sp": 13,   "wrng_lv": 14,   "tslt": 15},
    3: {"fcst": 16, "wrng_sp": 16,   "wrng_lv": 17,   "tslt": 18},
    4: {"fcst": 19, "wrng_sp": 19,   "wrng_lv": 20,   "tslt": 21},
    5: {"fcst": 22, "wrng_sp": 23,   "wrng_lv": 24,   "tslt": 25},
    6: {"fcst": 26, "wrng_sp": None, "wrng_lv": None,  "tslt": 25},
    7: {"fcst": 27, "wrng_sp": None, "wrng_lv": None,  "tslt": None},
}

# ── Colors ───────────────────────────────────────────────────────────────────
# EXHVY → full Red background,   bold title in Red shading
# VHVY  → full Orange background, no extra title shading
# IHVY  → full Yellow background, no extra title shading
# TSLT  → full Yellow background, no extra title shading
# None  → no background

LEVEL_BG = {
    "EXHVY": "FF0000",   # Red
    "VHVY":  "FFC000",   # Orange
    "IHVY":  "FFFF00",   # Yellow
    "TSLT":  "FFFF00",   # Yellow
    None:    None,
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def safe_get(row_tuple, col):
    if col is None or col >= len(row_tuple):
        return ""
    v = row_tuple[col]
    return "" if (v is None or (isinstance(v, float) and str(v) == "nan")) else str(v).strip()

def parse_wind_speed(tslt_text):
    m = re.search(r'(\d+[-\u2013]\d+)', str(tslt_text))
    return m.group(1) if m else "30-40"

def dominant_spatial(codes):
    valid = [(SPATIAL_RANK.get(c, 0), c) for c in codes if c in SPATIAL_RANK]
    return max(valid, default=(0, None))[1]

def format_districts(lst):
    lst = sorted(set(d for d in lst if d))
    if not lst: return ""
    if len(lst) == 1: return lst[0]
    return ", ".join(lst[:-1]) + ", and " + lst[-1]

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL READER
# ─────────────────────────────────────────────────────────────────────────────
def read_excel(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    records = []
    for i in range(3, len(df_raw)):
        row = tuple(df_raw.iloc[i])
        district = safe_get(row, 3)
        if not district:
            continue
        rec = {"DISTRICT": district}
        for day, cols in DAY_COLS.items():
            rec[f"D{day}_FCST"]    = safe_get(row, cols["fcst"])
            rec[f"D{day}_WRNG_SP"] = safe_get(row, cols["wrng_sp"])
            rec[f"D{day}_WRNG_LV"] = safe_get(row, cols["wrng_lv"])
            rec[f"D{day}_TSLT"]    = safe_get(row, cols["tslt"])
        records.append(rec)
    return records

def get_district_preview(file_path):
    """Return dataframe of all districts with warning codes for preview."""
    df_raw = pd.read_excel(file_path, header=None)
    rows = []
    for i in range(3, len(df_raw)):
        row = tuple(df_raw.iloc[i])
        district = safe_get(row, 3)
        if not district:
            continue
        rows.append({
            "District":   district,
            "D1 Warning": safe_get(row, 9),
            "D1 Fcst":    safe_get(row, 8),
            "D2 Warning": safe_get(row, 14),
            "D3 Warning": safe_get(row, 17),
            "D4 Warning": safe_get(row, 20),
            "D5 Warning": safe_get(row, 24),
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
# FORECAST
# ─────────────────────────────────────────────────────────────────────────────
def build_forecast(records, day):
    fcst_codes = [r[f"D{day}_FCST"] for r in records if r[f"D{day}_FCST"] in SPATIAL_RANK]
    dom  = dominant_spatial(fcst_codes) if fcst_codes else "ISOL"
    word = SPATIAL_FORECAST_WORD.get(dom, "ISOLATED")
    return (f"Light to Moderate Rain or Thundershowers very likely "
            f"to occur at {word} places over Telangana.")

# ─────────────────────────────────────────────────────────────────────────────
# WARNING PARTS — (bold_title, rest_of_sentence, level_code)
# ─────────────────────────────────────────────────────────────────────────────
def build_warning_parts(records, day):
    cols  = DAY_COLS[day]
    parts = []

    if cols["wrng_lv"] is not None:
        level_bucket   = defaultdict(list)
        spatial_bucket = defaultdict(list)
        for r in records:
            lv = r[f"D{day}_WRNG_LV"]
            sp = r[f"D{day}_WRNG_SP"]
            if lv in ("EXHVY", "VHVY", "IHVY"):
                level_bucket[lv].append(r["DISTRICT"])
                if sp in SPATIAL_RANK:
                    spatial_bucket[lv].append(sp)

        if level_bucket["EXHVY"]:
            dom = dominant_spatial(spatial_bucket["EXHVY"]) if spatial_bucket["EXHVY"] else "ISOL"
            loc = format_districts(level_bucket["EXHVY"]) + " districts of Telangana"
            parts.append(("Very Heavy to Extremely Heavy Rainfall",
                           f" very likely to occur {SPATIAL_PHRASE[dom]} in {loc}.", "EXHVY"))

        if level_bucket["VHVY"]:
            dom = dominant_spatial(spatial_bucket["VHVY"]) if spatial_bucket["VHVY"] else "ISOL"
            loc = format_districts(level_bucket["VHVY"]) + " districts of Telangana"
            parts.append(("Heavy to Very Heavy Rainfall",
                           f" very likely to occur {SPATIAL_PHRASE[dom]} in {loc}.", "VHVY"))

        if level_bucket["IHVY"]:
            dom = dominant_spatial(spatial_bucket["IHVY"]) if spatial_bucket["IHVY"] else "ISOL"
            loc = format_districts(level_bucket["IHVY"]) + " districts of Telangana"
            parts.append(("Heavy Rainfall",
                           f" very likely to occur {SPATIAL_PHRASE[dom]} in {loc}.", "IHVY"))
    else:
        if cols["tslt"] is not None:
            has_any = any("TSLT" in r[f"D{day}_TSLT"].upper() for r in records if r[f"D{day}_TSLT"])
            if has_any:
                parts.append(("Heavy rainfall",
                               " very likely to occur at isolated places over Telangana.", "IHVY"))

    if cols["tslt"] is not None:
        total          = len(records)
        tslt_districts = [r["DISTRICT"] for r in records if "TSLT" in r[f"D{day}_TSLT"].upper()]
        if tslt_districts:
            wind_speed = "30-40"
            for r in records:
                if "TSLT" in r[f"D{day}_TSLT"].upper():
                    wind_speed = parse_wind_speed(r[f"D{day}_TSLT"])
                    break
            if len(tslt_districts) == total:
                ts_loc = "ALL districts of Telangana"
                ts_sp  = "at isolated places"
            else:
                ts_loc = format_districts(tslt_districts) + " districts of Telangana"
                sp_codes = [r[f"D{day}_FCST"] for r in records
                            if r["DISTRICT"] in tslt_districts and r[f"D{day}_FCST"] in SPATIAL_RANK]
                ts_sp = SPATIAL_PHRASE.get(dominant_spatial(sp_codes) or "ISOL", "at isolated places")
            parts.append((f"Thunderstorm accompanied with Lightning and Gusty winds ({wind_speed} kmph)",
                           f" very likely occur {ts_sp} in {ts_loc}.", "TSLT"))

    if not parts:
        parts.append(("No warning for the day.", "", None))

    return parts

# ─────────────────────────────────────────────────────────────────────────────
# XML BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
def make_shd_elem(fill):
    s = etree.Element(w("shd"))
    s.set(w("val"),   "clear")
    s.set(w("color"), "auto")
    s.set(w("fill"),  fill)
    return s

def make_run(text, bold=False, font="Times New Roman"):
    r = etree.Element(w("r"))
    rpr = etree.SubElement(r, w("rPr"))
    fonts = etree.SubElement(rpr, w("rFonts"))
    fonts.set(w("ascii"),   font)
    fonts.set(w("hAnsi"),   font)
    fonts.set(w("eastAsia"), "Gautami")
    if bold:
        etree.SubElement(rpr, w("b"))
        etree.SubElement(rpr, w("bCs"))
    t = etree.SubElement(r, w("t"))
    t.set(f"{{{XML}}}space", "preserve")
    t.text = text
    return r

def make_warning_paragraph(bold_title, rest_text, level_code):
    """
    Build a <w:p> with:
    - Full background color on the paragraph (pPr > shd)
    - Bold title run + normal rest-of-sentence run
    - NO spacing before/after (NoSpacing style, spacing 0)
    """
    bg = LEVEL_BG.get(level_code)

    p = etree.Element(w("p"))
    ppr = etree.SubElement(p, w("pPr"))

    # Style
    ps = etree.SubElement(ppr, w("pStyle"))
    ps.set(w("val"), "NoSpacing")

    # Paragraph background shading — entire cell row colored
    if bg:
        ppr.append(make_shd_elem(bg))

    # No extra spacing before/after — paragraphs join flush
    spacing = etree.SubElement(ppr, w("spacing"))
    spacing.set(w("before"), "0")
    spacing.set(w("after"),  "0")
    spacing.set(w("line"),   "276")
    spacing.set(w("lineRule"), "auto")

    # Justify
    jc = etree.SubElement(ppr, w("jc"))
    jc.set(w("val"), "both")

    # pPr > rPr
    ppr_rpr = etree.SubElement(ppr, w("rPr"))
    fonts_ppr = etree.SubElement(ppr_rpr, w("rFonts"))
    fonts_ppr.set(w("ascii"),    "Times New Roman")
    fonts_ppr.set(w("hAnsi"),    "Times New Roman")
    fonts_ppr.set(w("eastAsia"), "Gautami")

    # Bold title run
    if bold_title:
        p.append(make_run(bold_title, bold=True))

    # Rest of sentence (not bold)
    if rest_text:
        p.append(make_run(rest_text, bold=False))

    return p

# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────
def get_para_text(p_elem):
    return "".join((t.text or "") for t in p_elem.iter(w("t")))

def replace_text_in_para(p_elem, old, new):
    full = get_para_text(p_elem)
    if old not in full:
        return False
    replaced = full.replace(old, new)
    for r in list(p_elem.findall(f".//{w('r')}")):
        r.getparent().remove(r)
    p_elem.append(make_run(replaced, bold=False))
    return True

def process_document_xml(xml_bytes, plain_context, warning_parts_map):
    parser = etree.XMLParser(remove_blank_text=False)
    tree   = etree.fromstring(xml_bytes, parser)

    all_paras = tree.findall(f".//{w('p')}")

    warning_replacements = []

    for p in all_paras:
        txt = get_para_text(p)

        # Warning placeholder
        for day_num in range(1, 8):
            ph = f"{{{{DAY{day_num}_WARNING}}}}"
            if ph in txt:
                parts  = warning_parts_map.get(day_num, [("No warning for the day.", "", None)])
                parent = p.getparent()
                idx    = list(parent).index(p)
                new_paras = [make_warning_paragraph(bt, rt, lv) for bt, rt, lv in parts]
                warning_replacements.append((p, parent, idx, new_paras))
                break

        # Plain text replacements
        for key, val in plain_context.items():
            if key in txt:
                replace_text_in_para(p, key, val)

    # Apply replacements in reverse order
    for p, parent, idx, new_paras in reversed(warning_replacements):
        parent.remove(p)
        for i, np in enumerate(new_paras):
            parent.insert(idx + i, np)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def generate_doc(file, template_path=None, issue_time="1300 HRS IST (MID-DAY)"):
    if hasattr(file, "read"):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(file.read())
        tmp.close()
        file = tmp.name

    records = read_excel(file)

    if template_path is None:
        base = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base, "imd_template_fixed.docx")

    today = datetime.today()

    plain_context = {
        "{{ISSUE_DATE}}": today.strftime("%d-%m-%Y"),
        "{{ISSUE_TIME}}": issue_time,
    }

    warning_parts_map = {}

    for day in range(1, 8):
        from_date = today + timedelta(days=day - 1)
        to_date   = today + timedelta(days=day)
        from_str  = f"1300 hrs of {from_date.strftime('%d/%m/%Y')}" if day == 1 \
                    else f"0830 hrs of {from_date.strftime('%d/%m/%Y')}"
        to_str    = f"0830 hrs Of {to_date.strftime('%d/%m/%Y')}"

        plain_context[f"{{{{DAY{day}_FROM}}}}"]     = from_str
        plain_context[f"{{{{DAY{day}_TO}}}}"]       = to_str
        plain_context[f"{{{{DAY{day}_FORECAST}}}}"] = build_forecast(records, day)
        warning_parts_map[day] = build_warning_parts(records, day)

    out_dir     = os.path.dirname(os.path.abspath(file))
    output_file = os.path.join(out_dir, "FINAL_IMD_OUTPUT.docx")

    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = process_document_xml(data, plain_context, warning_parts_map)
                zout.writestr(item, data)

    return output_file
