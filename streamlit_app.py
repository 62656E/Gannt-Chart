# Import necessary libraries
import streamlit as st
from io import BytesIO
from gantt import createGantt
import plotly as py

# Title streamlit app
st.title("Gantt Chart Maker")

# Prompt the user with the file format and purpose of the app
st.markdown(
    """
This app allows you to create a Gantt chart using the `plotly` library. Please upload an xlsx file with the following columns; Task, Stage, StartDate, EndDate, CompletionFrac and Title. To customise stage colours, add a column named `StageColor` with the colour of the stage in hex format. To customise legend order, add a column named `LegendOrder` with the order of the legend.
"""
)

# Upload the xlsx file
uploaded_file = st.file_uploader(
    "Choose xlsx file", type="xlsx", accept_multiple_files=False
)

img_bytes = None
GanttChart = None

# Check if the file has been uploaded
if uploaded_file is not None:
    with st.spinner("Creating Gantt Chart..."):
        img_bytes, GanttChart = createGantt(uploaded_file)
else:
    st.write("Please upload an xlsx file")

# Display the Gantt chart if it has been created
if GanttChart:
    st.plotly_chart(GanttChart)
elif img_bytes is type(BytesIO):
    st.image(img_bytes)
else:
    st.write("Please upload an xlsx file")

# Display download button
st.download_button(
    label="Download Gantt Chart as PNG",
    data=img_bytes,
    file_name="gantt_chart.png",
    mime="image/png",
)
