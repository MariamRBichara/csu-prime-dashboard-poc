from pathlib import Path
from datetime import datetime
import tempfile

import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates" / "slide06_template.pptx"
OUTPUT_DIR = BASE_DIR / "output"


# Historical Jan-Jun trend retained from the approved reference.
# Source A values override these whenever available.
BASELINE_TREND = {
    1: 0.81,
    2: 0.82,
    3: 0.84,
    4: 0.85,
    5: 0.86,
    6: 0.87,
}


def _remove_shape(shape):
    sp = shape._element
    sp.getparent().remove(sp)


def _find_picture(slide, name):
    for shape in slide.shapes:
        if (
            shape.shape_type
            == MSO_SHAPE_TYPE.PICTURE
            and shape.name == name
        ):
            return shape

    raise ValueError(
        f"Required template picture '{name}' was not found."
    )


def _set_title(slide, reporting_period):
    wanted = (
        "CSU Prime Dashboard – "
        + reporting_period
    )

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        if "CSU Prime Dashboard" in shape.text:
            frame = shape.text_frame

            if frame.paragraphs and frame.paragraphs[0].runs:
                first_run = frame.paragraphs[0].runs[0]
                first_run.text = wanted

                for run in frame.paragraphs[0].runs[1:]:
                    run.text = ""
            else:
                shape.text = wanted

            return

    raise ValueError(
        "Dashboard title was not found in the template."
    )


def _save_kpi_strip(metrics, output_path):
    k = metrics["kpis"]

    labels = [
        "HC tagged for ARVI",
        "Completion rate",
        "HC compliance",
        "ARVE",
        "ARVI",
        "Overbooking",
    ]

    values = [
        f"{int(k['hc_tagged'])}",
        f"{k['completion_rate']:.0%}",
        f"{k['hc_compliance']:.0%}",
        f"{k['arve']:.0%}",
        f"{k['arvi']:.0%}",
        f"{k['overbooking']:.0%}",
    ]

    value_colors = [
        "white",
        "#48C78E",
        "#48C78E",
        "white",
        "white",
        "#FF7900",
    ]

    fig = plt.figure(
        figsize=(12, 1.05),
        dpi=180,
        facecolor="black",
    )

    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 1)
    ax.axis("off")

    for i in range(6):
        ax.add_patch(
            plt.Rectangle(
                (i + 0.015, 0.03),
                0.97,
                0.93,
                facecolor="#050505",
                edgecolor="#262626",
                linewidth=1.0,
            )
        )

        ax.text(
            i + 0.5,
            0.83,
            labels[i],
            ha="center",
            va="center",
            color="white",
            fontsize=11,
            fontweight="bold",
        )

        if i == 0:
            ax.text(
                i + 0.5,
                0.64,
                f"(out of {int(k['total_hc'])} total HC)",
                ha="center",
                va="center",
                color="#BDBDBD",
                fontsize=7,
                fontstyle="italic",
            )

        ax.text(
            i + 0.5,
            0.42,
            values[i],
            ha="center",
            va="center",
            color=value_colors[i],
            fontsize=23,
            fontweight="bold",
        )

    fig.savefig(
        output_path,
        facecolor="black",
        bbox_inches="tight",
        pad_inches=0,
    )
    plt.close(fig)


def _display_k(value):
    return f"{round(float(value) / 1000):.0f}K"


def _build_trend(metrics):
    combined = dict(BASELINE_TREND)

    for values in metrics.get("trend", {}).values():
        month = int(values["month_number"])
        combined[month] = float(values["billable_pct"])

    if not combined:
        return [], []

    max_month = max(combined)

    months = [
        month
        for month in range(1, max_month + 1)
        if month in combined
    ]

    values = [
        combined[month]
        for month in months
    ]

    return months, values


def _save_dashboard_panel(metrics, output_path):
    worktypes = metrics["worktypes"]

    names = [
        "Billable",
        "General",
        "Leave",
        "Pre-sales",
    ]

    raw_values = [
        max(float(worktypes.get(name, 0.0)), 0.0)
        for name in names
    ]

    total = float(
        worktypes.get(
            "Total",
            sum(raw_values),
        )
    )

    if total <= 0:
        total = sum(raw_values)

    if total <= 0:
        total = 1.0

    percentages = [
        int(round(value / total * 100))
        for value in raw_values
    ]

    colors = [
        "#FF7900",
        "#8C8C8C",
        "#4FB2DC",
        "#AEDFF0",
    ]

    fig = plt.figure(
        figsize=(6.1, 5.0),
        dpi=180,
        facecolor="black",
    )

    # -------------------------------------------------
    # Upper section: worktype pie
    # -------------------------------------------------
    pie_ax = fig.add_axes(
        [0.04, 0.43, 0.70, 0.52],
        facecolor="black",
    )

    pie_ax.set_title(
        "Days allocation by worktype",
        color="white",
        fontsize=15,
        fontweight="bold",
        pad=4,
    )

    wedges, _ = pie_ax.pie(
        raw_values,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops={
            "linewidth": 0.6,
            "edgecolor": "white",
        },
    )

    pie_ax.axis("equal")

    legend_labels = [
        (
            f"{name} "
            f"{_display_k(value)} "
            f"({percent}%)"
        )
        for name, value, percent in zip(
            names,
            raw_values,
            percentages,
        )
    ]

    pie_ax.legend(
        wedges,
        legend_labels,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=False,
        labelcolor="white",
        fontsize=9,
        handlelength=1.2,
        handletextpad=0.5,
    )

    # -------------------------------------------------
    # Lower section: billable/non-billable trend
    # -------------------------------------------------
    trend_ax = fig.add_axes(
        [0.07, 0.08, 0.89, 0.26],
        facecolor="black",
    )

    months, billable_values = _build_trend(
        metrics
    )

    non_billable_values = [
        max(0.0, 1.0 - v)
        for v in billable_values
    ]

    x = list(range(len(months)))

    trend_ax.bar(
        x,
        billable_values,
        color="#FF7900",
        width=0.78,
        label="Billable",
    )

    trend_ax.bar(
        x,
        non_billable_values,
        bottom=billable_values,
        color="#4FB2DC",
        width=0.78,
        label="Non billable",
    )

    month_abbr = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec",
    }

    reporting_year = int(
        metrics["reporting_year"]
    )

    trend_ax.set_xticks(
        x,
        [
            f"{month_abbr[m]} {reporting_year}"
            for m in months
        ],
    )

    trend_ax.tick_params(
        axis="x",
        colors="white",
        labelsize=7,
        length=0,
    )

    trend_ax.set_ylim(0, 1)
    trend_ax.set_yticks([])

    for spine in trend_ax.spines.values():
        spine.set_visible(False)

    trend_ax.set_title(
        "Billable/non-billable days (%)",
        color="white",
        fontsize=14,
        fontweight="bold",
        pad=7,
    )

    for i, billable in enumerate(
        billable_values
    ):
        trend_ax.text(
            i,
            billable / 2,
            f"{billable:.0%}",
            ha="center",
            va="center",
            color="white",
            fontsize=8,
            fontweight="bold",
        )

    trend_ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.24),
        ncol=2,
        frameon=False,
        labelcolor="white",
        fontsize=8,
    )

    fig.savefig(
        output_path,
        facecolor="black",
        bbox_inches="tight",
        pad_inches=0,
    )

    plt.close(fig)


def build_powerpoint(metrics):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"PowerPoint template not found: "
            f"{TEMPLATE_PATH}"
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = (
        datetime.now()
        .strftime("%Y%m%d_%H%M%S")
    )

    output_file = (
        OUTPUT_DIR
        / (
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
            "Template must contain exactly one slide."
        )

    slide = prs.slides[0]

    _set_title(
        slide,
        metrics["reporting_period"],
    )

    kpi_picture = _find_picture(
        slide,
        "Image 2",
    )
    dashboard_picture = _find_picture(
        slide,
        "Image 5",
    )

    # Image 7 is intentionally preserved exactly as-is.
    _find_picture(
        slide,
        "Image 7",
    )

    with tempfile.TemporaryDirectory(
        prefix="csu_dashboard_"
    ) as temp_dir:
        temp_dir = Path(temp_dir)

        kpi_png = (
            temp_dir
            / "kpi_strip.png"
        )
        dashboard_png = (
            temp_dir
            / "dashboard_panel.png"
        )

        _save_kpi_strip(
            metrics,
            kpi_png,
        )
        _save_dashboard_panel(
            metrics,
            dashboard_png,
        )

        kpi_pos = (
            kpi_picture.left,
            kpi_picture.top,
            kpi_picture.width,
            kpi_picture.height,
        )

        dashboard_pos = (
            dashboard_picture.left,
            dashboard_picture.top,
            dashboard_picture.width,
            dashboard_picture.height,
        )

        _remove_shape(
            kpi_picture
        )
        _remove_shape(
            dashboard_picture
        )

        slide.shapes.add_picture(
            str(kpi_png),
            left=kpi_pos[0],
            top=kpi_pos[1],
            width=kpi_pos[2],
            height=kpi_pos[3],
        )

        slide.shapes.add_picture(
            str(dashboard_png),
            left=dashboard_pos[0],
            top=dashboard_pos[1],
            width=dashboard_pos[2],
            height=dashboard_pos[3],
        )

        prs.save(
            output_file
        )

    return str(
        output_file
    )
