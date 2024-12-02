# Import necessary libraries
import streamlit as st 
from io import BytesIO
from gantt import createGantt 

# Title streamlit app
st.title("Gantt Chart Maker")

# Prompt the user with the file format and purpose of the app
st.markdown("""
This app allows you to create a Gantt chart using the `plotly` library. Please upload an xlsx file with the following columns; Task, Stage, StartDate, EndDate, CompletionFrac and Title. To customise stage colours, add a column named `StageColor` with the colour of the stage in hex format. To customise legend order, add a column named `LegendOrder` with the order of the legend.
""")

# Upload the xlsx file
uploaded_file = st.file_uploader("Choose xlsx file", type="xlsx", accept_multiple_files=False)

height = st.number_input("Enter the height of the Gantt chart", value=1000)
width = st.number_input("Enter the width of the Gantt chart", value=2000)
title = st.text_input("Enter the title of the Gantt chart", value="Gantt Chart")
title_font_size = st.number_input("Enter the title font size", value=20)
axis_label_font_size = st.number_input("Enter the axis label font size", value=15)


# Check if the file has been uploaded
if uploaded_file is not None:
    with st.spinner("Creating Gantt Chart..."):
        img_bytes = createGantt(uploaded_file)
else:
    st.write("Please upload an xlsx file")

# Display the Gantt chart
if img_bytes.getbuffer().nbytes > 0:
    st.image(img_bytes, use_column_width=True) 
else:
    st.error("Failed to generate Gantt chart image")
    
# Display download button
st.download_button(
    label="Download Gantt Chart as PNG",
    data=img_bytes,
    file_name="gantt_chart.png",
    mime="image/png"
)

            
            
