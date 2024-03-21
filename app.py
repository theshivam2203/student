import streamlit as st
import pandas as pd
import os

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

st.title('Excel Processor')

uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file is not None:
    file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
    st.write(file_details)

    file_path = os.path.join(UPLOAD_FOLDER, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded successfully!")

    # Read the data from the Excel file
    xls = pd.ExcelFile(file_path)

    # Process each sheet
    output_dfs = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        output_dfs[sheet_name] = df

        # If the current sheet is Sheet1, process it
        if sheet_name == 'Sheet1':
            # Define the number of ranges
            num_ranges = 5

            # Define a function to calculate mean and assign ranges
            def calculate_mean_and_assign_range(column):
                # Calculate the mean value of the column
                mean_value = column.mean()

                # Calculate the range width
                range_width = (column.max() - column.min()) / num_ranges

                # Modify the assign_value function to use qualitative descriptions
                def assign_value(value):
                    for i in range(1, num_ranges + 1):
                        if value <= column.min() + i * range_width:
                            return {
                                1: "Very Low",
                                2: "Low",
                                3: "Average",
                                4: "High",
                                5: "Very High"
                            }[i]
                    return "Very High"  # In case the value exceeds all ranges

                # Apply the function to create a new column with the assigned values
                column_name = column.name.replace(" ", "_") + "_Range"
                df[column_name] = column.apply(assign_value)

            # Columns to process
            columns_to_process = [
                'Technology Acceptance',
                'Level of use of AI based tools',
                'Technology based Tutoring System',
                'Organisational Performance',
                "Student's Performance"
            ]

            # Apply the function to each specified column in Sheet 1
            for col in columns_to_process:
                calculate_mean_and_assign_range(df[col])

            # Calculate frequency and percentage of occurrences of values within each range in Sheet 1
            value_counts_df = pd.DataFrame(index=[
                "Very Low",
                "Low",
                "Average",
                "High",
                "Very High"
            ])

            # Initialize DataFrame columns for frequency and percentage
            for col in df.columns:
                if col.endswith('_Range'):
                    value_counts_df[col + '_Frequency'] = 0
                    value_counts_df[col + '_Percentage'] = 0

            # Calculate frequency and percentage for each range column in Sheet 1
            for col in df.columns:
                if col.endswith('_Range'):
                    value_counts = df[col].value_counts().reindex(["Very Low", "Low", "Average", "High", "Very High"]).fillna(0)
                    total_count = value_counts.sum()
                    
                    # Update DataFrame with frequency and percentage values
                    for index, value in value_counts.items():
                        value_counts_df.at[index, col + '_Frequency'] = value
                        value_counts_df.at[index, col + '_Percentage'] = (value / total_count) * 100

            # output_dfs['Processed_Data'] = df
            output_dfs['Value_Counts'] = value_counts_df

    # Save the DataFrames to an Excel file with Processed Data, Value Counts, and all sheets from the input file
    output_file_path = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
    with pd.ExcelWriter(output_file_path) as writer:
        for sheet_name, df in output_dfs.items():
            if sheet_name == 'Value_Counts':
                df.to_excel(writer, index=True, sheet_name=sheet_name, columns=value_counts_df.columns)
            else:
                df.to_excel(writer, index=False, sheet_name=sheet_name)

    st.success("Data processed successfully!")

    st.write("Download Processed File")
    st.download_button(
        label="Download",
        data=open(output_file_path, "rb").read(),
        file_name='output.xlsx',
        mime='application/octet-stream'
    )
