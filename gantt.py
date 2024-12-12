import pandas as pd
import plotly.graph_objects as go
from io import BytesIO
import plotly.io as pio


def create_gantt(uploaded_file):
    """
    Create a Gantt chart from an uploaded Excel file.

    Parameters:
    uploaded_file (xlsx): An uploaded xlsx file containing the following columns: Task, Stage, StartDate, EndDate, CompletionFrac, Title, StageColor, LegendOrder

    Returns:
    img_bytes (BytesIO): An image of the Gantt chart in Bytes
    gantt_chart (plotly.graph_objects.Figure): A Gantt chart figure
    """
    # Load data from the uploaded file
    data = pd.read_excel(uploaded_file, engine="openpyxl")

    # Convert date columns to datetime format
    data["StartDate"] = pd.to_datetime(data["StartDate"])
    data["EndDate"] = pd.to_datetime(data["EndDate"])

    # Calculate task durations and other derived columns
    data["DaysToStart"] = (data["StartDate"] - data["StartDate"].min()).dt.days
    data["Duration"] = (data["EndDate"] - data["StartDate"]).dt.days
    data["CompletionDays"] = data["CompletionFrac"] * data["Duration"]

    # Create a mapping for stage colors
    stage_color_dict = dict(zip(data["Stage"], data["StageColor"]))
    legend_order_dict = dict(zip(data["LegendOrder"], data["StageColor"]))

    # Extract chart title
    title = data["Title"].iloc[0]

    # Determine x-axis tick values
    x_tick_start = data["StartDate"].min()
    x_tick_end = data["EndDate"].max()
    x_tick_vals = pd.date_range(start=x_tick_start, end=x_tick_end, freq="MS")
    tickvals = (x_tick_vals - x_tick_start).days

    # Initialize the Gantt chart
    gantt_chart = go.Figure()

    # Dynamic chart sizing
    chart_height = max(400, 50 + 20 * len(data))  # Minimum height 400
    chart_width = max(
        1000, 100 * (x_tick_end - x_tick_start).days // 30
    )  # Minimum width 1000

    # Add tasks to the chart
    for _, row in data.iterrows():
        gantt_chart.add_trace(
            go.Bar(
                x=[row["Duration"]],
                y=[row["Task"]],
                base=row["DaysToStart"],
                orientation="h",
                marker=dict(color=stage_color_dict[row["Stage"]]),
                showlegend=False,
            )
        )

        gantt_chart.add_trace(
            go.Bar(
                x=[row["CompletionDays"]],
                y=[row["Task"]],
                base=row["DaysToStart"],
                orientation="h",
                marker=dict(color=stage_color_dict[row["Stage"]], opacity=0.5),
                showlegend=False,
            )
        )

    # Add legend entries
    for legend, color in legend_order_dict.items():
        gantt_chart.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(color=color, size=10),
                name=legend,
            )
        )

    # Update layout
    gantt_chart.update_layout(
        title=dict(text=title, x=0.5),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="right",
            x=0.95,
            bgcolor="rgba(255, 255, 255, 0.7)",
        ),
        xaxis=dict(
            title="Date",
            tickvals=tickvals,
            ticktext=[d.strftime("%b %Y") for d in x_tick_vals],
        ),
        yaxis=dict(
            title="Tasks",
            automargin=True,
            autorange="reversed",
        ),
        barmode="stack",
        height=chart_height,
        width=chart_width,
        margin=dict(l=150, r=50, t=50, b=100),
    )

    # Convert chart to PNG image in memory
    img_bytes = BytesIO()
    pio.write_image(gantt_chart, img_bytes, format="png")
    img_bytes.seek(0)

    return img_bytes, gantt_chart
