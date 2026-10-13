from pathlib import Path
from datetime import datetime

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
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

OUTPUT_DIR = BASE_DIR / "output"


# =========================================================
# COLORS
# =========================================================

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

ORANGE = RGBColor(255, 121, 0)
GREEN = RGBColor(80, 190, 135)
GREY = RGBColor(191, 191, 191)


# =========================================================
# BASELINE VALUES ALREADY PRINTED IN TEMPLATE
# =========================================================

# These values already exist inside the flattened
# reference images. If the newly calculated value is
# identical, we do NOT overlay anything.

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
    "Billable": {
        "k": 23,
        "pct": 87,
    },
    "General": {
        "k": 2,
        "pct": 7,
    },
    "Leave": {
        "k": 1,
        "pct": 4,
    },
    "Pre-sales": {
        "k": 0,
        "pct": 1,
    },
}


BASELINE_TREND = {
    1: 81,
    2: 82,
    3: 84,
    4: 85,
    5: 86,
    6: 87,
    7: 86,
    8: 87,
    9: 91,
}


# =========================================================
# GENERIC HELPERS
# =========================================================

def pct_int(value):
    return int(round(float(value) * 100))


def k_int(value):
    return int(round(float(value) / 1000))


def add_patch(
    slide,
    left,
    top,
    width,
    height,
    color,
):
    """
    Covers old text in a flattened image.
    """

    patch = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        int(left),
        int(top),
        int(width),
        int(height),
    )

    patch.fill.solid()
    patch.fill.fore_color.rgb = color

    patch.line.fill.background()

    return patch


def add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    font_size,
    font_color=WHITE,
    bold=True,
):
    """
    Adds replacement text.
    """

    box = slide.shapes.add_textbox(
        int(left),
        int(top),
        int(width),
        int(height),
    )

    frame = box.text_frame
    frame.clear()

    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0

    frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER

    run = paragraph.add_run()
    run.text = str(text)

    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = font_color

    return box


def patch_and_write(
    slide,
    left,
    top,
    width,
    height,
    text,
    font_size,
    background=BLACK,
    font_color=WHITE,
    bold=True,
):
    add_patch(
        slide,
        left,
        top,
        width,
        height,
        background,
    )

    return add_text(
        slide,
        left,
        top,
        width,
        height,
        text,
        font_size,
        font_color,
        bold,
    )


# =========================================================
# TITLE
# =========================================================

def update_title(
    slide,
    reporting_period,
):

    wanted = (
        "CSU Prime Dashboard – "
        + reporting_period
    )

    for shape in slide.shapes:

        if not shape.has_text_frame:
            continue

        if "CSU Prime Dashboard" in shape.text:

            # Do nothing if title already matches.
            if shape.text.strip() == wanted.strip():
                return

            frame = shape.text_frame
            frame.clear()

            paragraph = frame.paragraphs[0]
            run = paragraph.add_run()

            run.text = wanted

            return

    raise ValueError(
        "Dashboard title not found."
    )


# =========================================================
# FIND TEMPLATE IMAGES
# =========================================================

def find_reference_images(slide):

    images = {}

    for shape in slide.shapes:

        if (
            shape.shape_type
            == MSO_SHAPE_TYPE.PICTURE
        ):
            images[shape.name] = shape

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
# KPI SECTION
# =========================================================

def update_kpis(
    slide,
    image,
    metrics,
):

    k = metrics["kpis"]

    current = {
        "hc_tagged":
            int(k["hc_tagged"]),

        "total_hc":
            int(k["total_hc"]),

        "completion_rate":
            pct_int(
                k["completion_rate"]
            ),

        "hc_compliance":
            pct_int(
                k["hc_compliance"]
            ),

        "arve":
            pct_int(
                k["arve"]
            ),

        "arvi":
            pct_int(
                k["arvi"]
            ),

        "overbooking":
            pct_int(
                k["overbooking"]
            ),
    }


    baseline = {
        "hc_tagged":
            BASELINE_KPIS[
                "hc_tagged"
            ],

        "total_hc":
            BASELINE_KPIS[
                "total_hc"
            ],

        "completion_rate":
            pct_int(
                BASELINE_KPIS[
                    "completion_rate"
                ]
            ),

        "hc_compliance":
            pct_int(
                BASELINE_KPIS[
                    "hc_compliance"
                ]
            ),

        "arve":
            pct_int(
                BASELINE_KPIS[
                    "arve"
                ]
            ),

        "arvi":
            pct_int(
                BASELINE_KPIS[
                    "arvi"
                ]
            ),

        "overbooking":
            pct_int(
                BASELINE_KPIS[
                    "overbooking"
                ]
            ),
    }


    left = image.left
    top = image.top
    width = image.width
    height = image.height

    tile_width = width / 6


    items = [
        (
            "hc_tagged",
            str(current["hc_tagged"]),
            WHITE,
        ),
        (
            "completion_rate",
            f"{current['completion_rate']}%",
            GREEN,
        ),
        (
            "hc_compliance",
            f"{current['hc_compliance']}%",
            GREEN,
        ),
        (
            "arve",
            f"{current['arve']}%",
            WHITE,
        ),
        (
            "arvi",
            f"{current['arvi']}%",
            WHITE,
        ),
        (
            "overbooking",
            f"{current['overbooking']}%",
            ORANGE,
        ),
    ]


    for index, (
        key,
        value_text,
        color,
    ) in enumerate(items):

        # If already identical to template,
        # do absolutely nothing.
        if current[key] == baseline[key]:
            continue

        patch_and_write(
            slide,

            left
            + tile_width * index
            + tile_width * 0.16,

            top
            + height * 0.39,

            tile_width * 0.68,

            height * 0.40,

            value_text,

            font_size=22,

            background=BLACK,
            font_color=color,
            bold=True,
        )


    # HC total subtitle only changes
    # if total HC changes.
    if (
        current["total_hc"]
        != baseline["total_hc"]
    ):

        patch_and_write(
            slide,

            left
            + tile_width * 0.13,

            top
            + height * 0.20,

            tile_width * 0.74,

            height * 0.16,

            (
                f"(out of "
                f"{current['total_hc']} "
                f"total HC)"
            ),

            font_size=6,

            background=BLACK,
            font_color=GREY,
            bold=False,
        )


# =========================================================
# WORKTYPE SECTION
# =========================================================

def update_worktypes(
    slide,
    image,
    metrics,
):

    data = metrics["worktypes"]

    total = float(
        data["Total"]
    )

    if total <= 0:
        raise ValueError(
            "Worktype total is zero."
        )


    values = {}

    for name in [
        "Billable",
        "General",
        "Leave",
        "Pre-sales",
    ]:

        raw = float(
            data[name]
        )

        values[name] = {
            "k":
                k_int(raw),

            "pct":
                int(
                    round(
                        raw
                        / total
                        * 100
                    )
                ),
        }


    left = image.left
    top = image.top
    width = image.width
    height = image.height


    # -------------------------------------------------
    # BILLABLE
    # -------------------------------------------------

    if (
        values["Billable"]
        != BASELINE_WORKTYPE[
            "Billable"
        ]
    ):

        patch_and_write(
            slide,

            left
            + width * 0.50,

            top
            + height * 0.79,

            width * 0.33,

            height * 0.075,

            (
                f"Billable "
                f"{values['Billable']['k']}K "
                f"({values['Billable']['pct']}%)"
            ),

            font_size=7,

            background=BLACK,
            font_color=WHITE,
            bold=False,
        )


    # -------------------------------------------------
    # GENERAL
    # -------------------------------------------------

    if (
        values["General"]
        != BASELINE_WORKTYPE[
            "General"
        ]
    ):

        patch_and_write(
            slide,

            left
            + width * 0.07,

            top
            + height * 0.18,

            width * 0.28,

            height * 0.07,

            (
                f"General "
                f"{values['General']['k']}K "
                f"({values['General']['pct']}%)"
            ),

            font_size=7,

            background=BLACK,
            font_color=WHITE,
            bold=False,
        )


    # -------------------------------------------------
    # LEAVE
    #
    # For your current validated data it already
    # equals the template, so nothing gets added.
    # -------------------------------------------------

    if (
        values["Leave"]
        != BASELINE_WORKTYPE[
            "Leave"
        ]
    ):

        patch_and_write(
            slide,

            left
            + width * 0.18,

            top
            + height * 0.125,

            width * 0.23,

            height * 0.07,

            (
                f"Leave "
                f"{values['Leave']['k']}K "
                f"({values['Leave']['pct']}%)"
            ),

            font_size=7,

            background=BLACK,
            font_color=WHITE,
            bold=False,
        )


    # -------------------------------------------------
    # PRE-SALES
    #
    # Current template already shows 0K (1%).
    # Normally this therefore remains untouched.
    # -------------------------------------------------

    if (
        values["Pre-sales"]
        != BASELINE_WORKTYPE[
            "Pre-sales"
        ]
    ):

        # If it ever changes materially,
        # place the replacement just above
        # the pie rather than inside the slice.
        patch_and_write(
            slide,

            left
            + width * 0.36,

            top
            + height * 0.08,

            width * 0.24,

            height * 0.08,

            (
                f"Pre-sales "
                f"{values['Pre-sales']['k']}K "
                f"({values['Pre-sales']['pct']}%)"
            ),

            font_size=6,

            background=BLACK,
            font_color=WHITE,
            bold=False,
        )


# =========================================================
# TREND SECTION
# =========================================================

def update_trend(
    slide,
    image,
    metrics,
):

    trend = metrics["trend"]

    if not trend:
        return


    left = image.left
    top = image.top
    width = image.width
    height = image.height


    # Jan-Sep template = 9 fixed positions
    slot_width = width / 9


    for _, values in sorted(
        trend.items(),
        key=lambda item:
            item[1][
                "month_number"
            ],
    ):

        month = int(
            values[
                "month_number"
            ]
        )

        if month not in BASELINE_TREND:
            continue


        calculated = int(
            round(
                values[
                    "billable_pct"
                ]
                * 100
            )
        )


        # If the number already printed in
        # the reference image is correct,
        # leave it untouched.
        if (
            calculated
            == BASELINE_TREND[
                month
            ]
        ):
            continue


        index = month - 1


        # Text is inside the orange section.
        # Use orange patch rather than black.
        patch_and_write(
            slide,

            left
            + slot_width * index
            + slot_width * 0.26,

            top
            + height * 0.44,

            slot_width * 0.48,

            height * 0.10,

            f"{calculated}%",

            font_size=6,

            background=ORANGE,
            font_color=WHITE,
            bold=True,
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


    prs = Presentation(
        TEMPLATE_PATH
    )


    if len(prs.slides) != 1:

        raise ValueError(
            "Template must contain "
            "exactly one slide."
        )


    slide = prs.slides[0]


    (
        kpi_image,
        worktype_image,
        trend_image,
    ) = find_reference_images(
        slide
    )


    update_title(
        slide,
        metrics[
            "reporting_period"
        ],
    )


    update_kpis(
        slide,
        kpi_image,
        metrics,
    )


    update_worktypes(
        slide,
        worktype_image,
        metrics,
    )


    update_trend(
        slide,
        trend_image,
        metrics,
    )


    prs.save(
        output_file
    )


    return str(
        output_file
    )
