## Data Analyzer App

This is a Streamlit-based web application designed for analyzing and visualizing datasets. The app supports various file formats, including CSV, Excel (XLSX), and Excel 97-2003 (XLS). Users can upload their dataset and explore its contents through dataset preview and interactive plots.

### How to Run the App

1. Clone the repository to your local machine.

   ```bash
   git clone https://github.com/anandscripts/data-analyzer-app.git
   ```

2. Navigate to the project directory.

   ```bash
   cd data-analyzer-app
   ```

3. Install the required dependencies from the `requirements.txt` file.

   ```bash
   pip install -r requirements.txt
   ```

4. Run the Streamlit app.

   ```bash
   streamlit run app.py
   ```

5. Open your web browser and go to the provided URL (usually `http://localhost:8501`).

## Preview
### Interface Preview 
<img width="948" alt="Screenshot 2024-01-28 085934" src="https://github.com/anandscripts/data-analyzer-app/assets/83623438/9c9152e5-6a2d-4615-b175-d2d5a6a645bb">

### How to Use the App

1. Upon running the app, you will see a sidebar with the option to upload your dataset (CSV, XLSX, or XLS).

2. After uploading the dataset, you can customize the analysis parameters on the right side of the app.

   - Select the number of rows to display in the dataset preview.
   - Choose the X-axis and Y-axis columns for plotting.
   - Customize the plot type (Line Chart, Histogram, Scatter Plot, Bar Chart, Pie Chart).
   - Adjust additional parameters such as opacity, size, and color for specific plot types.

3. Explore the dataset preview and switch to the "Plots Visualization" tab to interactively visualize the data.

4. Download the generated report by clicking the "Download" button.

### Sample Datasets

The `Sample Datasets` folder contains three sample datasets: `UserData.csv`, `iris.csv`, and `mtcars.csv`. You can use these datasets to test the functionality of the Data Analyzer app.

### Generated Report

The app generates a detailed report in a Word document (`output.docx`) located in the project directory. The report includes an overview of the dataset, visualizations such as Line Charts, Histograms, Scatter Plots, Bar Charts, and Pie Charts, along with insights and explanations.

### Notes

- Make sure to have Python installed on your machine.
- If you encounter any issues, please check the dependencies and ensure that you have the required libraries installed.

Feel free to explore and analyze your datasets with the Data Analyzer app!
