# Import necessary libraries
import streamlit as st 
from io import BytesIO
import gantt 


# Title streamlit app
st.title("Gantt Chart Maker")

# Prompt the user with the file format and purpose of the app
st.markdown("""
This app allows you to create a Gantt chart using the `plotly` library. Please upload an xlsx file with the following columns; Task, Stage, StartDate, EndDate, CompletionFrac and Title. To customise stage colours, add a column named `StageColor` with the colour of the stage in hex format. To customise legend order, add a column named `LegendOrder` with the order of the legend.
""")

# Upload the xlsx file
uploaded_file = st.file_uploader("Choose xlsx file", type="xlsx", accept_multiple_files=False)

# Check if the file has been uploaded
if uploaded_file is not None:
    with st.spinner("Creating Gantt Chart..."):
        gantt_chart = gantt.createGantt(uploaded_file)
        with BytesIO() as output_bytes:
            gantt_chart.write_image(output_bytes, format="png")
            st.image(output_bytes)
            st.download_button(label="Download Gantt Chart", data=output_bytes, file_name="gantt_chart.png", mime="image/png")
            
