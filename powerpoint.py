from pathlib import Path
from datetime import datetime

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Pt


# =========================================================
# PATHS
# =========================================================

BASE_DIR = Path(__file__).parent

TEMPLATE_PATH = (
    BASE_DIR
    / "templates"
    / "slide06_template.pptx"
)

OUTPUT_DIR = (
    BASE_DIR
    / "output"
)


# =========================================================
# COLORS
# =========================================================

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

ORANGE = RGBColor(255, 121, 0)
GREEN = RGBColor(80, 190, 135)
BLUE = RGBColor(75, 180, 230)

GREY = RGBColor(191, 191, 191)


# =========================================================
# HELPERS
# =========================================================

def pct0(value):
    if value is None:
        return "N/A"

    return f"{round(value * 100):.0f}%"


def k_value(value):
    return f"{round(value / 1000):.0f}K"


def add_overlay_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    font_size,
    font_color=WHITE,
    bold=True,
    fill_color=None,
):
    """
    Adds a textbox over an existing flattened reference image.
    Optional fill_color can be used to hide the original value.
    """

    if fill_color is not None:

        patch = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left,
            top,
            width,
            height,
        )

        patch.fill.solid()
        patch.fill.fore_color.rgb = fill_color
        patch.line.fill.background()


    textbox = slide.shapes.add_textbox(
        left,
        top,
        width,
        height,
    )

    frame = textbox.text_frame
    frame.clear()

    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0

    frame.vertical_anchor = (
        MSO_ANCHOR.MIDDLE
    )

    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER

    run = paragraph.add_run()
    run.text = text

    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = font_color

    return textbox


# =========================================================
# TITLE
# =========================================================

def update_title(
    slide,
    reporting_period,
):

    new_title = (
        "CSU Prime Dashboard – "
        + reporting_period
    )

    for shape in slide.shapes:

        if not shape.has_text_frame:
            continue

        if (
            "CSU Prime Dashboard"
            in shape.text
        ):

            frame = shape.text_frame

            # Do not append to existing text.
            frame.clear()

            paragraph = frame.paragraphs[0]
            run = paragraph.add_run()

            run.text = new_title

            return

    raise ValueError(
        "Dashboard title not found."
    )


# =========================================================
# FIND REFERENCE IMAGES
# =========================================================

def find_reference_images(slide):

    images = {
        shape.name: shape
        for shape in slide.shapes
        if shape.shape_type == 13
    }

    required = [
        "Image 2",
        "Image 5",
        "Image 7",
    ]

    for name in required:

        if name not in images:
            raise ValueError(
                f"Reference image "
                f"'{name}' not found."
            )

    return (
        images["Image 2"],
        images["Image 5"],
        images["Image 7"],
    )


# =========================================================
# KPI OVERLAYS
# =========================================================

def overlay_kpis(
    slide,
    image,
    metrics,
):

    k = metrics["kpis"]

    values = [
        (
            str(k["hc_tagged"]),
            WHITE,
        ),
        (
            pct0(k["completion_rate"]),
            GREEN,
        ),
        (
            pct0(k["hc_compliance"]),
            GREEN,
        ),
        (
            pct0(k["arve"]),
            WHITE,
        ),
        (
            pct0(k["arvi"]),
            WHITE,
        ),
        (
            pct0(k["overbooking"]),
            ORANGE,
        ),
    ]

    left = image.left
    top = image.top
    width = image.width
    height = image.height

    tile_width = width / 6

    # Values sit around the lower-middle
    # portion of each KPI tile.
    for i, (
        value,
        color,
    ) in enumerate(values):

        x = int(
            left
            + tile_width * i
            + tile_width * 0.12
        )

        y = int(
            top
            + height * 0.40
        )

        w = int(
            tile_width * 0.76
        )

        h = int(
            height * 0.43
        )

        add_overlay_text(
            slide,
            x,
            y,
            w,
            h,
            value,
            font_size=22,
            font_color=color,
            bold=True,
            fill_color=BLACK,
        )


    # HC total subtitle
    add_overlay_text(
        slide,
        int(
            left
            + tile_width * 0.10
        ),
        int(
            top
            + height * 0.23
        ),
        int(
            tile_width * 0.80
        ),
        int(
            height * 0.18
        ),
        (
            f"(out of "
            f"{k['total_hc']} "
            f"total HC)"
        ),
        font_size=6,
        font_color=GREY,
        bold=False,
        fill_color=BLACK,
    )


# =========================================================
# WORKTYPE LABEL OVERLAYS
# =========================================================

def overlay_worktype_labels(
    slide,
    image,
    metrics,
):

    w = metrics["worktypes"]

    total = w["Total"]

    if total <= 0:
        raise ValueError(
            "Worktype total is zero."
        )

    labels = {
        "Billable": (
            w["Billable"],
            round(
                w["Billable"]
                / total
                * 100
            ),
        ),
        "General": (
            w["General"],
            round(
                w["General"]
                / total
                * 100
            ),
        ),
        "Leave": (
            w["Leave"],
            round(
                w["Leave"]
                / total
                * 100
            ),
        ),
        "Pre-sales": (
            w["Pre-sales"],
            round(
                w["Pre-sales"]
                / total
                * 100
            ),
        ),
    }

    left = image.left
    top = image.top
    width = image.width
    height = image.height


    # Billable label — lower-right
    value, percent = labels["Billable"]

    add_overlay_text(
        slide,
        int(left + width * 0.51),
        int(top + height * 0.72),
        int(width * 0.30),
        int(height * 0.12),
        (
            f"Billable "
            f"{k_value(value)} "
            f"({percent}%)"
        ),
        font_size=7,
        font_color=WHITE,
        bold=False,
        fill_color=BLACK,
    )


    # General label — upper-left
    value, percent = labels["General"]

    add_overlay_text(
        slide,
        int(left + width * 0.05),
        int(top + height * 0.17),
        int(width * 0.34),
        int(height * 0.10),
        (
            f"General "
            f"{k_value(value)} "
            f"({percent}%)"
        ),
        font_size=7,
        font_color=WHITE,
        bold=False,
        fill_color=BLACK,
    )


    # Leave label
    value, percent = labels["Leave"]

    add_overlay_text(
        slide,
        int(left + width * 0.13),
        int(top + height * 0.10),
        int(width * 0.28),
        int(height * 0.09),
        (
            f"Leave "
            f"{k_value(value)} "
            f"({percent}%)"
        ),
        font_size=7,
        font_color=WHITE,
        bold=False,
        fill_color=BLACK,
    )


    # Pre-sales label
    value, percent = labels["Pre-sales"]

    add_overlay_text(
        slide,
        int(left + width * 0.32),
        int(top + height * 0.26),
        int(width * 0.20),
        int(height * 0.13),
        (
            f"Pre-sales\n"
            f"{k_value(value)} "
            f"({percent}%)"
        ),
        font_size=6,
        font_color=WHITE,
        bold=False,
        fill_color=RGBColor(
            85,
            110,
            120,
        ),
    )


# =========================================================
# BILLABILITY TREND OVERLAYS
# =========================================================

def overlay_billability(
    slide,
    image,
    metrics,
):

    trend = metrics["trend"]

    if not trend:
        raise ValueError(
            "No billability trend."
        )


    ordered = sorted(
        trend.items(),
        key=lambda item:
            item[1][
                "month_number"
            ],
    )


    left = image.left
    top = image.top
    width = image.width
    height = image.height


    # Historical template contains
    # 9 equally spaced monthly positions.
    #
    # For months represented in current Source A,
    # overlay the calculated percentage.

    slot_count = 9

    slot_width = (
        width
        / slot_count
    )


    for (
        month_name,
        values,
    ) in ordered:

        month_num = (
            values[
                "month_number"
            ]
        )

        # Jan=0 ... Sep=8
        index = (
            month_num - 1
        )

        # Ignore months beyond Sep
        # for this exact historical template.
        if (
            index < 0
            or index >= slot_count
        ):
            continue


        billable = round(
            values[
                "billable_pct"
            ]
            * 100
        )


        x = int(
            left
            + slot_width * index
            + slot_width * 0.16
        )

        y = int(
            top
            + height * 0.36
        )

        w = int(
            slot_width * 0.68
        )

        h = int(
            height * 0.18
        )


        add_overlay_text(
            slide,
            x,
            y,
            w,
            h,
            f"{billable}%",
            font_size=7,
            font_color=WHITE,
            bold=True,
            fill_color=ORANGE,
        )


# =========================================================
# BUILD POWERPOINT
# =========================================================

def build_powerpoint(
    metrics,
):

    if not TEMPLATE_PATH.exists():

        raise FileNotFoundError(
            f"PowerPoint template "
            f"not found: "
            f"{TEMPLATE_PATH}"
        )


    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )


    timestamp = (
        datetime.now()
        .strftime(
            "%Y%m%d_%H%M%S"
        )
    )


    output_file = (
        OUTPUT_DIR
        /
        (
            "CSU_Prime_Dashboard_"
            + timestamp
            + ".pptx"
        )
    )


    # ---------------------------------------------
    # Open exact approved reference
    # ---------------------------------------------

    prs = Presentation(
        TEMPLATE_PATH
    )


    if len(
        prs.slides
    ) != 1:

        raise ValueError(
            "Template must contain "
            "exactly one slide."
        )


    slide = (
        prs.slides[0]
    )


    # ---------------------------------------------
    # Find original flattened graphics
    # ---------------------------------------------

    (
        kpi_image,
        worktype_image,
        trend_image,
    ) = find_reference_images(
        slide
    )


    # ---------------------------------------------
    # Update title
    # ---------------------------------------------

    update_title(
        slide,
        metrics[
            "reporting_period"
        ],
    )


    # ---------------------------------------------
    # Preserve all original graphics.
    # Overlay current values only.
    # ---------------------------------------------

    overlay_kpis(
        slide,
        kpi_image,
        metrics,
    )


    overlay_worktype_labels(
        slide,
        worktype_image,
        metrics,
    )


    overlay_billability(
        slide,
        trend_image,
        metrics,
    )


    # ---------------------------------------------
    # Save
    # ---------------------------------------------

    prs.save(
        output_file
    )


    return str(
        output_file
    )