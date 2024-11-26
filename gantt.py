import pandas as pd
import openpyxl as px
import plotly.graph_objects as go
from io import BytesIO
import plotly.io as pio


def createGantt(uploaded_file):
    # Create dataframe from the uploaded file
    data = pd.read_excel(uploaded_file, engine="openpyxl")

    # Convert the StartDate and EndDate columns to datetime
    data["StartDate"] = pd.to_datetime(data["StartDate"])
    data["EndDate"] = pd.to_datetime(data["EndDate"])

    # Calculate the number of days to start of each task
    data["DaysToStart"] = (data["StartDate"] - data["StartDate"].min()).dt.days

    # Calculate the number of days from project start to the end date of each task
    data["DaysToEnd"] = (data["EndDate"] - data["StartDate"].min()).dt.days

    # Calculate the duration of each task
    data["Duration"] = (data["EndDate"] - data["StartDate"]).dt.days

    # Calculate the completion fraction of each task
    data["CompletionDays"] = data["CompletionFrac"] * data["Duration"]

    # Create stage colour dictionary
    stage_color_dict = dict(zip(data["Stage"].dropna().unique(), data["StageColor"].dropna().unique()))

    # Create legend order list
    legend_order = data["LegendOrder"].dropna().unique()

    # Create legend order dictionary with legend order as key and corresponding stage color as value
    legend_order_dict = dict(zip(data["LegendOrder"].dropna().unique(), data["StageColor"].dropna().unique()))

    # Get title of the Gantt chart
    title = data["Title"].dropna().unique()[0]  # Ensure title is a string

    # Get max and min values for x axis
    x_tick_start = data["StartDate"].min()
    x_tick_end = data["EndDate"].max()

    # Create ticks for x axis
    x_tick_vals = pd.date_range(start=x_tick_start, end=x_tick_end, freq="MS")
    tickvals = (x_tick_vals - x_tick_start).days

    # Create figure
    GanttChart = go.Figure()

    # Plot each task as a bar with overlay for completion fraction
    for index, row in data.iterrows():
        # Add the main task bar
        GanttChart.add_trace(
            go.Bar(
                x=[row["Duration"]],
                y=[row["Task"]],
                base=row["DaysToStart"] + 1,
                orientation="h",
                marker=dict(
                    color=stage_color_dict[row["Stage"]],
                    line=dict(color=stage_color_dict[row["Stage"]], width=1),
                ),
                showlegend=False,
            ),
        )

        # Overlay completion fraction bar
        GanttChart.add_trace(
            go.Bar(
                x=[row["CompletionDays"]],
                y=[row["Task"]],
                base=row["DaysToStart"] + 1,
                orientation="h",
                marker=dict(color=stage_color_dict[row["Stage"]], opacity=0.5),
                showlegend=False,
            )
        )

    # Add legend entries for each stage
    for stage, color in legend_order_dict.items():
        GanttChart.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(color=color, size=10),
                name=stage,
                showlegend=True,  # Make sure it's shown in the legend
            )
        )

    # Update layout
    GanttChart.update_layout(
        legend=dict(  # Customize legend
            orientation="v",  # Vertical orientation
            yanchor="top",  # Anchor legend to top
            y=1,  # Position legend at top
            xanchor="right",  # Anchor legend to right
            x=0.95,  # Inset legend slightly from right
            bgcolor="rgba(255, 255, 255, 0.5)",  # Semi-transparent white background for readability
        ),
        title=title,  # Add title
        title_x=0.5,  # Center title
        xaxis=dict(  # Customize x axis
            title="Date",  # Add x axis title
            showgrid=True,  # Show gridlines
            tickvals=tickvals,  # Set tick values as days
        ),
        yaxis=dict(  # Customize y axis
            title="Tasks",  # Add y axis title
            automargin=True,  # Automatically adjust margin
            autorange="reversed",  # Reverse order of tasks
            tickvals=data["Task"],  # Set tick values
            ticktext=data["Task"],  # Set tick text
        ),
        barmode="stack",  # Stack bars on top of each other
        bargap=0.1,  # Set gap between bars
        hovermode="closest",  # Show hover info for closest data point
        margin=dict(  # Add margin around plot
            l=150,  # Left margin
            r=50,  # Right margin
            t=50,  # Top margin
            b=100,  # Bottom margin
        ),
    )

     # Convert figure to image in memory using BytesIO
    img_bytes = BytesIO()
    pio.write_image(GanttChart, img_bytes, format="png")
    img_bytes.seek(0)  # Rewind the BytesIO object to the beginning

    return img_bytes
