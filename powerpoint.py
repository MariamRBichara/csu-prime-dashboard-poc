from pathlib import Path
from datetime import datetime

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Pt

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates" / "slide06_template.pptx"
OUTPUT_DIR = BASE_DIR / "output"

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
ORANGE = RGBColor(255, 121, 0)
GREEN = RGBColor(80, 190, 135)
GREY = RGBColor(191, 191, 191)

BASELINE_KPIS = {
    "hc_tagged": 656,
    "total_hc": 761,
    "completion_rate": 0.96,
    "hc_compliance": 0.96,
    "arve": 0.86,
    "arvi": 0.83,
    "overbooking": 0.09,
}

BASELINE_WORKTYPE = {
    "Billable": {"k": 23, "pct": 87},
    "General": {"k": 2, "pct": 7},
    "Leave": {"k": 1, "pct": 4},
    "Pre-sales": {"k": 0, "pct": 1},
}

BASELINE_TREND = {
    1: 81, 2: 82, 3: 84, 4: 85, 5: 86,
    6: 87, 7: 86, 8: 87, 9: 91,
}

def pct_int(value):
    return int(round(float(value) * 100))

def k_int(value):
    return int(round(float(value) / 1000))

def add_patch(slide, left, top, width, height, color):
    patch = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        int(left), int(top), int(width), int(height)
    )
    patch.fill.solid()
    patch.fill.fore_color.rgb = color
    patch.line.fill.background()
    return patch

def add_text(
    slide, left, top, width, height, text, font_size,
    font_color=WHITE, bold=True, alignment=PP_ALIGN.CENTER,
):
    box = slide.shapes.add_textbox(
        int(left), int(top), int(width), int(height)
    )
    frame = box.text_frame
    frame.clear()
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    paragraph = frame.paragraphs[0]
    paragraph.alignment = alignment
    run = paragraph.add_run()
    run.text = str(text)
    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = font_color
    return box

def patch_and_write(
    slide, left, top, width, height, text, font_size,
    background=BLACK, font_color=WHITE, bold=True,
    alignment=PP_ALIGN.CENTER,
):
    add_patch(slide, left, top, width, height, background)
    add_text(
        slide, left, top, width, height, text, font_size,
        font_color, bold, alignment
    )

def update_title(slide, reporting_period):
    wanted = "CSU Prime Dashboard – " + reporting_period
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if "CSU Prime Dashboard" in shape.text:
            if shape.text.strip() == wanted.strip():
                return
            frame = shape.text_frame
            frame.clear()
            run = frame.paragraphs[0].add_run()
            run.text = wanted
            return
    raise ValueError("Dashboard title not found.")

def find_reference_images(slide):
    images = {
        shape.name: shape
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    }

    for name in ["Image 2", "Image 5", "Image 7"]:
        if name not in images:
            raise ValueError(f"Reference image '{name}' not found.")

    # Image 2 = KPI strip
    # Image 5 = worktype pie + billability trend
    # Image 7 = Q1/Q2/Q3 table
    return images["Image 2"], images["Image 5"], images["Image 7"]

def update_kpis(slide, image, metrics):
    k = metrics["kpis"]

    current = {
        "hc_tagged": int(k["hc_tagged"]),
        "total_hc": int(k["total_hc"]),
        "completion_rate": pct_int(k["completion_rate"]),
        "hc_compliance": pct_int(k["hc_compliance"]),
        "arve": pct_int(k["arve"]),
        "arvi": pct_int(k["arvi"]),
        "overbooking": pct_int(k["overbooking"]),
    }

    baseline = {
        "hc_tagged": 656,
        "total_hc": 761,
        "completion_rate": 96,
        "hc_compliance": 96,
        "arve": 86,
        "arvi": 83,
        "overbooking": 9,
    }

    left, top, width, height = image.left, image.top, image.width, image.height
    tile_width = width / 6

    items = [
        ("hc_tagged", str(current["hc_tagged"]), WHITE),
        ("completion_rate", f"{current['completion_rate']}%", GREEN),
        ("hc_compliance", f"{current['hc_compliance']}%", GREEN),
        ("arve", f"{current['arve']}%", WHITE),
        ("arvi", f"{current['arvi']}%", WHITE),
        ("overbooking", f"{current['overbooking']}%", ORANGE),
    ]

    for index, (key, value_text, color) in enumerate(items):
        if current[key] == baseline[key]:
            continue

        patch_and_write(
            slide,
            left + tile_width * index + tile_width * 0.16,
            top + height * 0.39,
            tile_width * 0.68,
            height * 0.40,
            value_text,
            font_size=22,
            background=BLACK,
            font_color=color,
            bold=True,
        )

    if current["total_hc"] != baseline["total_hc"]:
        patch_and_write(
            slide,
            left + tile_width * 0.12,
            top + height * 0.19,
            tile_width * 0.76,
            height * 0.17,
            f"(out of {current['total_hc']} total HC)",
            font_size=6,
            background=BLACK,
            font_color=GREY,
            bold=False,
        )

def update_worktypes(slide, image, metrics):
    data = metrics["worktypes"]
    total = float(data["Total"])

    if total <= 0:
        raise ValueError("Worktype total is zero.")

    values = {}
    for name in ["Billable", "General", "Leave", "Pre-sales"]:
        raw = float(data[name])
        values[name] = {
            "k": k_int(raw),
            "pct": int(round(raw / total * 100)),
        }

    left, top, width, height = image.left, image.top, image.width, image.height

    if values["General"] != BASELINE_WORKTYPE["General"]:
        patch_and_write(
            slide,
            left + width * 0.115,
            top + height * 0.135,
            width * 0.19,
            height * 0.055,
            f"General {values['General']['k']}K ({values['General']['pct']}%)",
            font_size=6,
            background=BLACK,
            font_color=WHITE,
            bold=False,
        )

    if values["Billable"] != BASELINE_WORKTYPE["Billable"]:
        patch_and_write(
            slide,
            left + width * 0.535,
            top + height * 0.845,
            width * 0.22,
            height * 0.055,
            f"Billable {values['Billable']['k']}K ({values['Billable']['pct']}%)",
            font_size=6,
            background=BLACK,
            font_color=WHITE,
            bold=False,
        )

    if values["Leave"] != BASELINE_WORKTYPE["Leave"]:
        patch_and_write(
            slide,
            left + width * 0.205,
            top + height * 0.095,
            width * 0.18,
            height * 0.055,
            f"Leave {values['Leave']['k']}K ({values['Leave']['pct']}%)",
            font_size=6,
            background=BLACK,
            font_color=WHITE,
            bold=False,
        )

    # Pre-sales is intentionally left untouched until an approved
    # replacement location is defined.

def update_trend(slide, image, metrics):
    trend = metrics["trend"]

    if not trend:
        return

    left, top, width, height = image.left, image.top, image.width, image.height
    slot_width = width / 9

    for _, values in sorted(
        trend.items(),
        key=lambda item: item[1]["month_number"],
    ):
        month = int(values["month_number"])

        if month not in BASELINE_TREND:
            continue

        calculated = int(round(values["billable_pct"] * 100))
        baseline = BASELINE_TREND[month]

        if calculated == baseline:
            continue

        index = month - 1

        patch_and_write(
            slide,
            left + slot_width * index + slot_width * 0.27,
            top + height * 0.805,
            slot_width * 0.46,
            height * 0.065,
            f"{calculated}%",
            font_size=5,
            background=ORANGE,
            font_color=WHITE,
            bold=True,
        )

def build_powerpoint(metrics):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"PowerPoint template not found: {TEMPLATE_PATH}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = OUTPUT_DIR / f"CSU_Prime_Dashboard_{timestamp}.pptx"

    prs = Presentation(TEMPLATE_PATH)

    if len(prs.slides) != 1:
        raise ValueError("Template must contain exactly one slide.")

    slide = prs.slides[0]

    kpi_image, dashboard_image, q_table_image = find_reference_images(slide)

    update_title(slide, metrics["reporting_period"])
    update_kpis(slide, kpi_image, metrics)
    update_worktypes(slide, dashboard_image, metrics)
    update_trend(slide, dashboard_image, metrics)

    # Preserve the Q table exactly as-is.
    _ = q_table_image

    prs.save(output_file)

    return str(output_file)
