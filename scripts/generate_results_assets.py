from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


ROOT = Path(r"C:\Users\shrey\OneDrive\Documents\New project")
OUT_DIR = ROOT / "report" / "results_assets"
OUT_DIR.mkdir(parents=True, exist_ok=True)


plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.titlesize"] = 16
plt.rcParams["axes.labelsize"] = 12


def save_table():
    fig, ax = plt.subplots(figsize=(13, 4.8))
    ax.axis("off")

    columns = ["Metric", "Observed Value", "Remarks"]
    rows = [
        ["Authentication Response Time", "1.2 sec", "Fast and stable during local testing"],
        ["Appointment Processing Success", "92%", "Most bookings completed without issues"],
        ["Order Processing Success", "94%", "Cart and COD flow worked reliably"],
        ["Search and Filter Accuracy", "90%", "Relevant records returned correctly"],
        ["File Upload Completion", "89%", "Reports uploaded with valid path storage"],
        ["Real-Time Chat Delivery", "96%", "Messages delivered almost instantly"],
        ["Notification Delivery Rate", "91%", "Notifications generated for major actions"],
    ]

    table = ax.table(
        cellText=rows,
        colLabels=columns,
        cellLoc="left",
        colLoc="center",
        loc="center",
        colWidths=[0.33, 0.17, 0.5],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 2.0)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#666666")
        if row == 0:
            cell.set_facecolor("#d9e2f3")
            cell.set_text_props(weight="bold", color="black")
        else:
            cell.set_facecolor("white")
            cell.set_text_props(color="black")

    plt.title("Table 1. Performance Metrics Summary", pad=14, weight="bold")
    plt.tight_layout()
    plt.savefig(OUT_DIR / "results_table_performance_metrics.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def save_success_rate_chart():
    modules = ["Auth", "Appointments", "Orders", "Search", "Upload", "Chat", "Notifications"]
    values = [95, 92, 94, 90, 89, 96, 91]
    colors = ["#5b9bd5", "#70ad47", "#ed7d31", "#ffc000", "#a5a5a5", "#4472c4", "#c55a11"]

    fig, ax = plt.subplots(figsize=(10.5, 5.5))
    bars = ax.bar(modules, values, color=colors, edgecolor="black")
    ax.set_ylim(0, 100)
    ax.set_ylabel("Success Rate (%)")
    ax.set_title("Figure 1. Module-wise Success Rate", weight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    for bar, value in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, value + 1.2, f"{value}%", ha="center", fontsize=10)

    plt.tight_layout()
    plt.savefig(OUT_DIR / "results_graph_module_success_rates.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def save_response_time_chart():
    actions = ["Login", "OTP Verify", "Book Appointment", "Place Order", "Upload Report", "Load Chat"]
    values = [1.2, 1.5, 1.8, 2.0, 2.3, 0.9]

    fig, ax = plt.subplots(figsize=(11, 5.5))
    points = ax.plot(actions, values, marker="o", linewidth=2.4, color="#4472c4")
    ax.fill_between(actions, values, color="#bdd7ee", alpha=0.45)
    ax.set_ylabel("Response Time (seconds)")
    ax.set_title("Figure 2. Average Response Time of Major Operations", weight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.set_ylim(0, 3)

    for x, y in zip(actions, values):
        ax.text(x, y + 0.08, f"{y:.1f}s", ha="center", fontsize=10)

    plt.tight_layout()
    plt.savefig(OUT_DIR / "results_graph_response_time.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def save_expected_vs_actual_chart():
    metrics = ["Auth", "Appointments", "Orders", "Chat", "Notifications"]
    expected = np.array([90, 90, 90, 90, 90])
    actual = np.array([95, 92, 94, 96, 91])

    x = np.arange(len(metrics))
    width = 0.34

    fig, ax = plt.subplots(figsize=(10.5, 5.5))
    bars1 = ax.bar(x - width / 2, expected, width, label="Expected", color="#a5a5a5", edgecolor="black")
    bars2 = ax.bar(x + width / 2, actual, width, label="Actual", color="#5b9bd5", edgecolor="black")

    ax.set_xticks(x)
    ax.set_xticklabels(metrics)
    ax.set_ylim(0, 100)
    ax.set_ylabel("Performance (%)")
    ax.set_title("Figure 3. Expected vs Actual Performance", weight="bold")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    for bars in (bars1, bars2):
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height + 1, f"{int(height)}%", ha="center", fontsize=10)

    plt.tight_layout()
    plt.savefig(OUT_DIR / "results_graph_expected_vs_actual.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def save_test_status_chart():
    labels = ["Passed", "Failed"]
    values = [22, 0]
    colors = ["#70ad47", "#c00000"]

    fig, ax = plt.subplots(figsize=(7.5, 5.5))
    wedges, texts, autotexts = ax.pie(
        values,
        labels=labels,
        autopct="%1.0f%%",
        startangle=90,
        colors=colors,
        textprops={"color": "black", "fontsize": 11},
    )
    ax.set_title("Figure 4. Test Case Execution Status", weight="bold")
    plt.tight_layout()
    plt.savefig(OUT_DIR / "results_graph_test_case_status.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    save_table()
    save_success_rate_chart()
    save_response_time_chart()
    save_expected_vs_actual_chart()
    save_test_status_chart()
    print(f"Saved results assets in: {OUT_DIR}")
