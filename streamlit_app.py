# Import necessary libraries
import streamlit as st
from io import BytesIO
from gantt import createGantt

# Title of the Streamlit app
st.title("Gantt Chart Maker")

# Description for the user
st.markdown(
    """
This app allows you to create a Gantt chart using the `plotly` library. 
Upload an xlsx file with the following columns:
- **Task**
- **Stage**
- **StartDate**
- **EndDate**
- **CompletionFrac**
- **Title**  
- **StageColor**: Hex colors for each stage.
- **LegendOrder**: Order of the legend items.
"""
)

# File uploader
uploaded_file = st.file_uploader("Choose an xlsx file", type="xlsx")

# Initialize variables
img_bytes = None
GanttChart = None

# Process the uploaded file
if uploaded_file:
    with st.spinner("Creating Gantt Chart..."):
        try:
            # Generate the Gantt chart and image
            img_bytes, GanttChart = createGantt(uploaded_file)
        except Exception as e:
            st.error(f"Error creating Gantt Chart: {e}")

# Display the Gantt chart
if GanttChart:
    st.plotly_chart(GanttChart)
    # Provide a download button for the chart image
    st.download_button(
        label="Download Gantt Chart as PNG",
        data=img_bytes.getvalue(),
        file_name="gantt_chart.png",
        mime="image/png",
    )
else:
    if uploaded_file:
        st.error("Failed to create the Gantt chart. Check your data format.")
    else:
        st.info("Please upload an xlsx file to generate a Gantt chart.")
