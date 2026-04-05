import pandas as pd
from collections import defaultdict
from docxtpl import DocxTemplate
from datetime import datetime, timedelta


# 🔹 Convert IMD format
def convert_imd(file):
    df = pd.read_excel(file, header=None)

    data = []

    for i in range(len(df)):
        try:
            district = df.iloc[i, 2]

            if pd.isna(district):
                continue

            row = {
                "DISTRICT": str(district).strip(),

                "DAY1_wrng": df.iloc[i, 5] if df.shape[1] > 5 else "",
                "DAY2_wrng": df.iloc[i, 8] if df.shape[1] > 8 else "",
                "DAY3_wrng": df.iloc[i, 11] if df.shape[1] > 11 else "",
                "DAY4_wrng": df.iloc[i, 14] if df.shape[1] > 14 else "",
                "DAY5_wrng": df.iloc[i, 17] if df.shape[1] > 17 else "",
                "DAY6_wrng": df.iloc[i, 19] if df.shape[1] > 19 else "",
                "DAY7_wrng": df.iloc[i, 20] if df.shape[1] > 20 else "",
            }

            data.append(row)

        except:
            continue

    return pd.DataFrame(data)


# 🔹 MULTIPLE CLASSIFICATION (IMPORTANT 🔥)
def classify_multiple(text):
    text = str(text)

    categories = []

    if "Extremely Heavy" in text:
        categories.append("EXTREME")

    if "Very Heavy" in text:
        categories.append("VERY_HEAVY")

    if "Heavy" in text:
        categories.append("HEAVY")

    if not categories:
        categories.append("NORMAL")

    return categories


# 🔹 Process data (MULTI-WARNING SUPPORT)
def process(df, day):
    grouped = defaultdict(list)

    for _, row in df.iterrows():
        district = row["DISTRICT"]
        warnings = classify_multiple(row[f"DAY{day}_wrng"])

        for w in warnings:
            grouped[w].append(district)

    return grouped


# 🔹 Clean bullet formatter
def format_bullets(items):
    return "\n".join([f"• {i}" for i in sorted(set(items))])


# 🔹 Build structured warnings (KEY 🔥)
def build_warning_text(grouped):
    sections = []

    if grouped["EXTREME"]:
        sections.append(
            "Very Heavy to Extremely Heavy Rainfall:\n" +
            format_bullets(grouped["EXTREME"])
        )

    if grouped["VERY_HEAVY"]:
        sections.append(
            "Heavy to Very Heavy Rainfall:\n" +
            format_bullets(grouped["VERY_HEAVY"])
        )

    if grouped["HEAVY"]:
        sections.append(
            "Heavy Rainfall:\n" +
            format_bullets(grouped["HEAVY"])
        )

    return "\n\n".join(sections)


# 🔹 MAIN FUNCTION
def generate_doc(file):
    df = convert_imd(file)

    # ✅ CHANGE PATH IF NEEDED
    doc = DocxTemplate("imd_template.docx")

    context = {}

    today = datetime.today()

    # 🔷 ISSUE DATE
    context["ISSUE_DATE"] = today.strftime("%d-%m-%Y")

    for day in range(1, 8):

        # 🔷 DATE CALCULATION
        from_date = today + timedelta(days=day - 1)
        to_date = today + timedelta(days=day)

        context[f"DAY{day}_FROM"] = from_date.strftime("%d/%m/%Y")
        context[f"DAY{day}_TO"] = to_date.strftime("%d/%m/%Y")

        grouped = process(df, day)

        # 🔷 FORECAST (STATIC / CAN BE MODIFIED)
        forecast = (
            "Light to Moderate Rain or Thundershowers very likely "
            "to occur at isolated/few places over Telangana."
        )

        # 🔷 MULTI-WARNING TEXT (🔥 FIXED)
        warning = build_warning_text(grouped)

        if warning.strip() == "":
            warning = "NIL"

        # 🔷 PASS TO TEMPLATE
        context[f"DAY{day}_FORECAST"] = forecast
        context[f"DAY{day}_WARNING"] = warning

    # 🔷 RENDER TEMPLATE
    doc.render(context)

    output_file = "FINAL_IMD_OUTPUT.docx"
    doc.save(output_file)

    return output_file
