import pandas as pd
import plotly.graph_objects as go
from io import BytesIO
import plotly.io as pio
import datetime as dt

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
    x_tick_end = pd.Timestamp("2027-04-01")
    x_tick_vals = pd.date_range(start=x_tick_start, end=x_tick_end, freq="MS")
    if x_tick_vals[-1] < x_tick_end:
        x_tick_vals = x_tick_vals.append(pd.DatetimeIndex([x_tick_end]))
    tickvals = (x_tick_vals - x_tick_start).days

    # Initialize the Gantt chart
    gantt_chart = go.Figure()

    # Dynamic chart sizing
    chart_height = max(400, 50 + 20 * len(data))  # Minimum height 400
    chart_width = max(1000, 200 * (x_tick_end - x_tick_start).days // 30)  # Minimum width 1000

    # Add tasks to the chart
    for _, row in data.iterrows():
        gantt_chart.add_trace(
            go.Bar(
                x=[row["Duration"]],
                y=[row["Task"]],
                base=row["DaysToStart"],
                orientation="h",
                marker=dict(color=stage_color_dict[row["Stage"]], line=dict(color=stage_color_dict[row["Stage"]], width=1)),
                hovertemplate=(
                    f"<b>Task:</b> {row['Task']}<br>"
                    f"<b>Stage:</b> {row['Stage']}<br>"
                    f"<b>Start Date:</b> {row['StartDate'].date()}<br>"
                    f"<b>End Date:</b> {row['EndDate'].date()}<br>"
                    f"<b>Completion:</b> {row['CompletionFrac'] * 100:.0f}%<br>"
                    f"<b>Duration:</b> {row['Duration']} days<extra></extra>"
                ),
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
                hoverinfo="skip",
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
        title=dict(text=title, x=0.5, font_size=17),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="right",
            x=0.95,
            bgcolor="rgba(255, 255, 255, 0.5)"
        ),
        xaxis=dict(
            title="Timeline",
            tickvals=tickvals,
            ticktext=x_tick_vals.strftime("%B %Y"),
            tickangle=45,
            range=[0, (x_tick_end - x_tick_start).days],
            tickfont=dict(size=15),
        ),
        yaxis=dict(
            title="Tasks",
            automargin=True,
            autorange="reversed",
            tickvals=data["Task"],
            ticktext=data["Task"],
            tickfont=dict(size=15),
        ),
        barmode="stack",
        bargap=0.1,
        hovermode="closest",
        height=chart_height,
        width=chart_width,
        margin=dict(l=150, r=50, t=50, b=100),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_family="Arial",
            font_color="black",
            align="left",
            bordercolor="black",
        ),
    )

    # Add today’s date as a vertical line
    today = dt.date.today()
    days_from_start = (today - data["StartDate"].min().date()).days
    gantt_chart.add_vline(
        x=days_from_start,
        line=dict(color="mediumorchid", dash="dash"),
        annotation_text=today.strftime("%Y-%m-%d"),
        annotation_position="top",
    )

    gantt_chart.add_trace(
        go.Scatter(
            x=[None],
            y=[None],
            mode="lines",
            line=dict(color="mediumorchid", dash="dash"),
            name="Today’s Date",
            showlegend=True,
        )
    )

    # Convert chart to PNG image in memory
    img_bytes = BytesIO()
    pio.write_image(gantt_chart, img_bytes, format="png")
    img_bytes.seek(0)

    return img_bytes, gantt_chart
