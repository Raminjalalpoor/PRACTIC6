from __future__ import annotations

import plotly.graph_objects as go
from plotly.offline import plot


COLORS = {
    "navy": "#10233f",
    "blue": "#2f6fed",
    "emerald": "#13a67a",
    "amber": "#f4a825",
    "slate": "#64748b",
    "grid": "#e8edf5",
}


def _layout(fig: go.Figure, *, height: int = 360) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=30, r=20, t=25, b=35),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Arial, sans-serif", color=COLORS["navy"], size=12),
        hoverlabel=dict(bgcolor=COLORS["navy"], font_color="white", bordercolor=COLORS["navy"]),
        showlegend=False,
    )
    return fig


def _as_div(fig: go.Figure) -> str:
    return plot(
        fig,
        output_type="div",
        include_plotlyjs=False,
        config={"displayModeBar": False, "responsive": True, "locale": "ru"},
    )


def revenue_chart(monthly):
    fig = go.Figure(
        go.Scatter(
            x=monthly["Месяц"],
            y=monthly["Выручка"],
            mode="lines+markers",
            line=dict(color=COLORS["blue"], width=3, shape="spline"),
            marker=dict(size=8, color="white", line=dict(color=COLORS["blue"], width=3)),
            fill="tozeroy",
            fillcolor="rgba(47, 111, 237, 0.08)",
            hovertemplate="%{x}<br><b>%{y:,.0f} ₽</b><extra></extra>",
        )
    )
    _layout(fig)
    fig.update_xaxes(showgrid=False, tickangle=-25)
    fig.update_yaxes(gridcolor=COLORS["grid"], rangemode="tozero", tickformat=",.0f", title="Выручка, ₽")
    return _as_div(fig)


def products_chart(products):
    top = products.head(10).sort_values("Количество")
    fig = go.Figure(
        go.Bar(
            x=top["Количество"],
            y=top["Товар"],
            orientation="h",
            marker=dict(
                color=top["Количество"],
                colorscale=[[0, "#9de3cf"], [1, COLORS["emerald"]]],
                line=dict(width=0),
            ),
            text=top["Количество"].map(lambda value: f"{value:,.0f}".replace(",", " ")),
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br><b>%{x:,.2f} ед.</b><extra></extra>",
        )
    )
    _layout(fig, height=390)
    fig.update_xaxes(gridcolor=COLORS["grid"], rangemode="tozero", title="Продано единиц")
    fig.update_yaxes(showgrid=False)
    return _as_div(fig)


def categories_chart(categories):
    palette = [COLORS["blue"], COLORS["emerald"], COLORS["amber"], "#7c5ce5", "#ec6f91", "#4ca6a8"]
    fig = go.Figure(
        go.Pie(
            labels=categories["Категория"],
            values=categories["Выручка"],
            hole=0.62,
            marker=dict(colors=palette, line=dict(color="white", width=3)),
            textinfo="percent",
            hovertemplate="%{label}<br><b>%{value:,.0f} ₽</b><br>%{percent}<extra></extra>",
        )
    )
    _layout(fig)
    fig.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center"))
    fig.add_annotation(text="Доля<br>выручки", x=0.5, y=0.5, showarrow=False, font=dict(size=14, color=COLORS["slate"]))
    return _as_div(fig)


def build_charts(analysis):
    return {
        "revenue": revenue_chart(analysis.monthly),
        "products": products_chart(analysis.products),
        "categories": categories_chart(analysis.categories),
    }
