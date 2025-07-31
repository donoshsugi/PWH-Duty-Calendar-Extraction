import streamlit as st
import pandas as pd
import re
from calendar import monthrange
import io

# --- Helper Function to Process the Data (Adapted from our previous script) ---
def process_roster(file_bytes, sheet_name, duty_column_name, filename):
    """
    Processes the uploaded Excel file bytes to generate a calendar CSV.

    Args:
        file_bytes (BytesIO): The Excel file uploaded by the user.
        sheet_name (str): The name of the sheet to process.
        duty_column_name (str): The name of the person/column to extract duty for.
        filename (str): The original name of the uploaded file.

    Returns:
        DataFrame: A pandas DataFrame ready for CSV conversion, or None if an error occurs.
    """
    # --- 1. Extract Month and Year from Filename ---
    match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})', filename, re.IGNORECASE)
    if not match:
        st.error(f"Error: Could not find a month and year in the filename '{filename}'.")
        st.info("Please ensure the filename is formatted like 'Duty August 2025.xlsx'.")
        return None

    month_name, year_str = match.groups()
    year = int(year_str)
    month_map = {name.lower(): num for num, name in enumerate(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], 1)}
    month = month_map[month_name.lower()]

    # --- 2. Determine Header Row Based on Sheet Name ---
    # According to your rules:
    # 'Duty - Senior': Header is on the 4th row (skip 3 rows).
    # 'Duty_MO', 'Duty - Part time': Header is on the 3rd row (skip 2 rows).
    header_row_map = {
        "Duty - Senior": 3,
        "Duty_MO": 2,
        "Duty - Part time": 2
    }
    # Default to skipping 2 rows if sheet name is not standard
    skiprows = header_row_map.get(sheet_name, 2)

    try:
        # --- 3. Read and Clean the Specific Excel Sheet ---
        df = pd.read_excel(file_bytes, sheet_name=sheet_name, skiprows=skiprows)

        # Rename first two columns to be consistent
        df.rename(columns={df.columns[0]: "week", df.columns[1]: "day"}, inplace=True)

        # Check if the selected duty column exists
        if duty_column_name not in df.columns:
            st.error(f"Error: Column '{duty_column_name}' not found in the sheet '{sheet_name}'.")
            return None

        # Keep only the essential columns
        df = df[["day", duty_column_name]]

        # --- 4. Filter and Process the Roster ---
        # Remove rows where the 'day' is not a number (handles extra text/empty rows)
        df = df[pd.to_numeric(df['day'], errors='coerce').notna()]
        df['day'] = df['day'].astype(int)

        # VERY IMPORTANT: Remove days where the person has no duty (handles empty cells)
        df.dropna(subset=[duty_column_name], inplace=True)
        
        if df.empty:
            st.warning(f"No duties found for '{duty_column_name}' in the selected sheet.")
            return None

        # Find the first row where the day is '1' to handle mid-month starts in the sheet
        try:
            start_index = df[df['day'] == 1].index[0]
            df = df.loc[start_index:]
        except IndexError:
            st.error("Could not find the start of the month (Day 1) in the sheet. Please check the 'day' column.")
            return None

        # --- 5. Generate Dates and Create Final DataFrame ---
        num_days_in_month = monthrange(year, month)[1]
        
        # Create the final DataFrame for export
        calendar_df = pd.DataFrame()
        # The subject can be the person's name plus the duty type from the cell
        calendar_df['Subject'] = duty_column_name + " - " + df[duty_column_name].astype(str)
        
        # Create the date for each duty day
        calendar_df['Start Date'] = df['day'].apply(lambda day: f"{year}-{month:02d}-{day:02d}")
        
        calendar_df['All Day Event'] = 'True'

        return calendar_df

    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Excel file: {e}")
        st.info("Please ensure the file is not corrupted and the sheet format is correct.")
        return None


# --- Streamlit Web App UI ---

st.set_page_config(page_title="Duty Roster to Calendar", layout="centered")

st.title("📅 Duty Roster to Google Calendar Converter")
st.write("This tool converts your XLSX duty roster into a CSV file that you can directly import into Google Calendar.")

# --- Step 1: File Upload ---
st.header("Step 1: Upload Your Roster File")
st.info("Your file name must contain the month and year (e.g., `Duty August 2025 Final.xlsx`)")
uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")

# --- App Logic: Proceeds only if a file is uploaded ---
if uploaded_file is not None:
    # Use io.BytesIO to handle the uploaded file in memory
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    
    try:
        xls = pd.ExcelFile(file_bytes)
        sheet_names = xls.sheet_names

        # --- Step 2: Sheet Selection ---
        st.header("Step 2: Select the Duty Sheet")
        selected_sheet = st.selectbox("Which sheet contains the duty roster?", sheet_names)

        if selected_sheet:
            # Determine header row to read the names correctly
            header_row_map = {"Duty - Senior": 3, "Duty_MO": 2, "Duty - Part time": 2}
            skiprows = header_row_map.get(selected_sheet, 2)
            
            # Read just the header row to get names
            df_for_cols = pd.read_excel(file_bytes, sheet_name=selected_sheet, skiprows=skiprows, nrows=0)
            
            # Filter out helper columns like 'week', 'day', and any unnamed columns
            duty_names = [col for col in df_for_cols.columns if 'unnamed' not in str(col).lower() and str(col).lower() not in ['week', 'day']]

            if not duty_names:
                 st.error(f"Could not find any staff names in the sheet '{selected_sheet}'. Please check the file format.")
            else:
                # --- Step 3: Name Selection ---
                st.header("Step 3: Select Your Name")
                selected_name = st.selectbox("Whose duty do you want to export?", duty_names)

                # --- Step 4: Process and Download ---
                st.header("Step 4: Generate and Download")
                if st.button(f"Generate Calendar for {selected_name}"):
                    with st.spinner("Processing your file..."):
                        final_df = process_roster(file_bytes, selected_sheet, selected_name, uploaded_file.name)

                    if final_df is not None:
                        st.success("✅ Your calendar file is ready!")
                        st.dataframe(final_df)

                        # Convert DataFrame to CSV string, then encode to bytes
                        csv_bytes = final_df.to_csv(index=False).encode('utf-8')

                        # Create the download button
                        output_filename = f"{selected_name.replace(' ', '_')}-{selected_sheet}-{monthrange(2025,8)[0]}-{year_str}-Calendar.csv"
                        st.download_button(
                            label="📥 Download .csv File",
                            data=csv_bytes,
                            file_name=output_filename,
                            mime='text/csv',
                        )

    except Exception as e:
        st.error(f"Failed to read the Excel file. It might be corrupted or in an unsupported format. Error: {e}")
