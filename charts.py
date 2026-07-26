"""
charts.py — Plotly chart generators for clinical pharmacy analytics.

All functions return plotly.graph_objects.Figure instances styled with a
consistent healthcare-themed color palette and dark background.
"""

import plotly.graph_objects as go
import plotly.express as px
import pandas as pd

# ---------------------------------------------------------------------------
# Color palette — professional healthcare theme
# ---------------------------------------------------------------------------
COLORS = {
    "primary": "#4F8FEA",       # Calm blue
    "success": "#2ECC71",       # Green — accepted
    "danger": "#E74C3C",        # Red — rejected
    "warning": "#F39C12",       # Amber
    "info": "#1ABC9C",          # Teal
    "purple": "#9B59B6",        # Purple
    "pink": "#E91E90",          # Pink
    "orange": "#E67E22",        # Orange
    "bg_dark": "#0E1117",       # Dark background
    "bg_card": "#1E2530",       # Card background
    "text": "#FAFAFA",          # Light text
    "text_muted": "#8899AA",    # Muted text
    "grid": "#2A3442",          # Grid lines
}

PALETTE = [
    "#4F8FEA", "#2ECC71", "#E74C3C", "#F39C12", "#9B59B6",
    "#1ABC9C", "#E67E22", "#E91E90", "#3498DB", "#27AE60",
    "#C0392B", "#8E44AD", "#16A085", "#D35400", "#2980B9",
]


def _base_layout(title: str, height: int = 450) -> dict:
    """Return a base layout dict for consistent chart styling."""
    return dict(
        title=dict(text=title, font=dict(size=18, color=COLORS["text"]), x=0.5),
        paper_bgcolor=COLORS["bg_dark"],
        plot_bgcolor=COLORS["bg_card"],
        font=dict(color=COLORS["text"], family="Inter, sans-serif"),
        height=height,
        margin=dict(l=40, r=40, t=60, b=40),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color=COLORS["text_muted"]),
        ),
    )


def _base_layout_light(title: str, height: int = 450) -> dict:
    """Return a light layout for Word report export."""
    return dict(
        title=dict(text=title, font=dict(size=16, color="#1a1a2e"), x=0.5),
        paper_bgcolor="white",
        plot_bgcolor="#f8f9fa",
        font=dict(color="#1a1a2e", family="Arial, sans-serif"),
        height=height,
        margin=dict(l=50, r=50, t=70, b=50),
        legend=dict(
            bgcolor="rgba(255,255,255,0.8)",
            font=dict(color="#333"),
        ),
    )


def acceptance_pie_chart(accepted: int, rejected: int, for_report: bool = False) -> go.Figure:
    """Pie chart showing intervention acceptance vs rejection."""
    labels = ["Accepted", "Rejected"]
    values = [accepted, rejected]
    colors = [COLORS["success"], COLORS["danger"]]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=colors, line=dict(color=COLORS["bg_dark"] if not for_report else "white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=14),
        hovertemplate="<b>%{label}</b><br>Count: %{value}<br>%{percent}<extra></extra>",
    )])

    layout = _base_layout_light("Intervention Acceptance Rate") if for_report else _base_layout("Intervention Acceptance Rate")
    fig.update_layout(**layout)

    # Add center annotation
    total = accepted + rejected
    fig.add_annotation(
        text=f"<b>{total}</b><br>Total",
        showarrow=False,
        font=dict(size=16, color=COLORS["text"] if not for_report else "#1a1a2e"),
        x=0.5, y=0.5,
    )

    return fig


def ward_interventions_chart(ward_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Bar chart of interventions by ward."""
    if ward_df.empty:
        return _empty_chart("No ward data available")

    fig = go.Figure(data=[go.Bar(
        x=ward_df["Ward"],
        y=ward_df["Count"],
        marker=dict(
            color=ward_df["Count"],
            colorscale=[[0, "#4F8FEA"], [1, "#9B59B6"]],
            line=dict(width=0),
            cornerradius=4,
        ),
        text=ward_df["Count"],
        textposition="outside",
        textfont=dict(size=12),
        hovertemplate="<b>%{x}</b><br>Interventions: %{y}<extra></extra>",
    )])

    layout = _base_layout_light("Interventions by Ward") if for_report else _base_layout("Interventions by Ward")
    layout["xaxis"] = dict(
        title="Ward",
        gridcolor=COLORS["grid"] if not for_report else "#e0e0e0",
        tickangle=-45 if len(ward_df) > 6 else 0,
    )
    layout["yaxis"] = dict(
        title="Number of Interventions",
        gridcolor=COLORS["grid"] if not for_report else "#e0e0e0",
    )
    fig.update_layout(**layout)
    return fig


def consultant_chart(consultant_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Horizontal bar chart of interventions by consultant."""
    if consultant_df.empty:
        return _empty_chart("No consultant data available")

    # Show top 15 if there are too many
    df = consultant_df.head(15).sort_values("Count", ascending=True)

    fig = go.Figure(data=[go.Bar(
        x=df["Count"],
        y=df["Consultant"],
        orientation="h",
        marker=dict(
            color=df["Count"],
            colorscale=[[0, "#1ABC9C"], [1, "#4F8FEA"]],
            cornerradius=4,
        ),
        text=df["Count"],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>Interventions: %{x}<extra></extra>",
    )])

    height = max(400, len(df) * 35 + 100)
    layout = _base_layout_light("Interventions by Consultant", height) if for_report else _base_layout("Interventions by Consultant", height)
    layout["xaxis"] = dict(title="Number of Interventions", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    layout["yaxis"] = dict(title="", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    fig.update_layout(**layout)
    return fig


def error_type_chart(error_results: dict, for_report: bool = False) -> go.Figure:
    """Donut chart of medication error types."""
    categories = [
        ("Prescription", error_results.get("prescription", 0)),
        ("Administration", error_results.get("administration", 0)),
        ("Transcription", error_results.get("transcription", 0)),
        ("Illegible Handwriting", error_results.get("illegible_handwriting", 0)),
        ("Incorrect Abbreviation", error_results.get("incorrect_abbreviation", 0)),
        ("Other", error_results.get("other", 0)),
    ]
    categories = [(name, count) for name, count in categories if count > 0]

    if not categories:
        return _empty_chart("No medication error data available")

    labels, values = zip(*categories)

    fig = go.Figure(data=[go.Pie(
        labels=list(labels),
        values=list(values),
        hole=0.5,
        marker=dict(colors=PALETTE[:len(labels)],
                    line=dict(color=COLORS["bg_dark"] if not for_report else "white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>Count: %{value}<br>%{percent}<extra></extra>",
    )])

    layout = _base_layout_light("Medication Error Types") if for_report else _base_layout("Medication Error Types")
    fig.update_layout(**layout)

    total = sum(values)
    fig.add_annotation(
        text=f"<b>{total}</b><br>Errors",
        showarrow=False,
        font=dict(size=15, color=COLORS["text"] if not for_report else "#1a1a2e"),
        x=0.5, y=0.5,
    )

    return fig


def errors_by_ward_chart(cross_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Grouped bar chart showing error distribution across wards."""
    if cross_df.empty:
        return _empty_chart("No ward-error data available")

    fig = go.Figure()
    for i, col in enumerate(cross_df.columns):
        fig.add_trace(go.Bar(
            name=col,
            x=cross_df.index.tolist(),
            y=cross_df[col].tolist(),
            marker=dict(color=PALETTE[i % len(PALETTE)], cornerradius=3),
            hovertemplate=f"<b>{col}</b><br>Ward: %{{x}}<br>Count: %{{y}}<extra></extra>",
        ))

    layout = _base_layout_light("Errors by Ward", 500) if for_report else _base_layout("Errors by Ward", 500)
    layout["barmode"] = "group"
    layout["xaxis"] = dict(title="Ward", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0",
                           tickangle=-45 if len(cross_df) > 6 else 0)
    layout["yaxis"] = dict(title="Number of Errors", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    fig.update_layout(**layout)
    return fig


def ham_lasa_drug_chart(by_drug_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Bar chart of most commonly administered HAM/LASA drugs."""
    if by_drug_df.empty:
        return _empty_chart("No HAM/LASA data available")

    df = by_drug_df.head(15)

    fig = go.Figure(data=[go.Bar(
        x=df["Drug"],
        y=df["Count"],
        marker=dict(
            color=df["Count"],
            colorscale=[[0, "#E74C3C"], [0.5, "#F39C12"], [1, "#E91E90"]],
            cornerradius=4,
        ),
        text=df["Count"],
        textposition="outside",
        textfont=dict(size=12),
        hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>",
    )])

    layout = _base_layout_light("Most Common HAM/LASA Medications") if for_report else _base_layout("Most Common HAM/LASA Medications")
    layout["xaxis"] = dict(title="Medication", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0",
                           tickangle=-45 if len(df) > 6 else 0)
    layout["yaxis"] = dict(title="Consumption Count", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    fig.update_layout(**layout)
    return fig


def ham_lasa_type_chart(by_type_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Pie chart showing HAM vs LASA distribution."""
    if by_type_df.empty:
        return _empty_chart("No HAM/LASA type data available")

    fig = go.Figure(data=[go.Pie(
        labels=by_type_df["Type"].tolist(),
        values=by_type_df["Count"].tolist(),
        hole=0.5,
        marker=dict(colors=[COLORS["danger"], COLORS["warning"], COLORS["info"]],
                    line=dict(color=COLORS["bg_dark"] if not for_report else "white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=13),
    )])

    layout = _base_layout_light("HAM vs LASA Distribution") if for_report else _base_layout("HAM vs LASA Distribution")
    fig.update_layout(**layout)
    return fig


def ham_lasa_patient_chart(patient_df: pd.DataFrame, for_report: bool = False) -> go.Figure:
    """Horizontal bar chart of how many HAM/LASA records each patient has."""
    if patient_df.empty:
        return _empty_chart("No HAM/LASA patient data available")

    df = patient_df.head(15).sort_values("Count", ascending=True)
    labels = df["Patient Name"] + " (" + df["MR No"] + ")"

    fig = go.Figure(data=[go.Bar(
        x=df["Count"],
        y=labels,
        orientation="h",
        marker=dict(
            color=df["Count"],
            colorscale=[[0, "#E74C3C"], [1, "#9B59B6"]],
            cornerradius=4,
        ),
        text=df["Count"],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>HAM/LASA Records: %{x}<extra></extra>",
    )])

    height = max(400, len(df) * 35 + 100)
    layout = _base_layout_light("HAM/LASA Frequency by Patient", height) if for_report else _base_layout("HAM/LASA Frequency by Patient", height)
    layout["xaxis"] = dict(title="Number of HAM/LASA Records", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    layout["yaxis"] = dict(title="", gridcolor=COLORS["grid"] if not for_report else "#e0e0e0")
    fig.update_layout(**layout)
    return fig


def adr_summary_chart(total_adrs: int, from_ham_lasa: int, for_report: bool = False) -> go.Figure:
    """Pie chart showing ADRs from HAM/LASA vs other causes."""
    if total_adrs == 0:
        return _empty_chart("No ADR data available")

    labels = ["From HAM/LASA", "Other Causes"]
    values = [from_ham_lasa, total_adrs - from_ham_lasa]
    colors = [COLORS["danger"], COLORS["primary"]]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=colors,
                    line=dict(color=COLORS["bg_dark"] if not for_report else "white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=13),
    )])

    layout = _base_layout_light("ADR Sources") if for_report else _base_layout("ADR Sources")
    fig.update_layout(**layout)

    fig.add_annotation(
        text=f"<b>{total_adrs}</b><br>ADRs",
        showarrow=False,
        font=dict(size=16, color=COLORS["text"] if not for_report else "#1a1a2e"),
        x=0.5, y=0.5,
    )

    return fig


def _empty_chart(message: str) -> go.Figure:
    """Return an empty chart with a 'no data' message."""
    fig = go.Figure()
    fig.add_annotation(
        text=message,
        showarrow=False,
        font=dict(size=16, color=COLORS["text_muted"]),
        x=0.5, y=0.5,
        xref="paper", yref="paper",
    )
    fig.update_layout(**_base_layout(""))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return fig
