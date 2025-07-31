import streamlit as st
import pandas as pd
import re
from calendar import monthrange
import io

# --- Helper Function to Process the Data ---
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
        tuple: (month_name, year_str) if successful, otherwise None.
    """
    # --- 1. Extract Month and Year from Filename ---
    match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})', filename, re.IGNORECASE)
    if not match:
        st.error(f"Error: Could not find a month and year in the filename '{filename}'.")
        st.info("Please ensure the filename is formatted like 'Duty August 2025.xlsx'.")
        return None, None

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
            return None, None

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
            return None, None

        # Find the first row where the day is '1' to handle mid-month starts in the sheet
        try:
            start_index = df[df['day'] == 1].index[0]
            df = df.loc[start_index:]
        except IndexError:
            st.error("Could not find the start of the month (Day 1) in the sheet. Please check the 'day' column.")
            return None, None

        # --- 5. Generate Dates and Create Final DataFrame ---
        # Create the final DataFrame for export
        calendar_df = pd.DataFrame()
        # The subject should not contain the name of the person
        calendar_df['Subject'] = df[duty_column_name].astype(str)
        
        # Create the date for each duty day
        calendar_df['Start Date'] = df['day'].apply(lambda day: f"{year}-{month:02d}-{day:02d}")
        
        calendar_df['All Day Event'] = 'True'

        return calendar_df, (month_name, year_str)

    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Excel file: {e}")
        st.info("Please ensure the file is not corrupted and the sheet format is correct.")
        return None, None


# --- Streamlit Web App UI ---

st.set_page_config(page_title="PWH Duty Roster to Calendar", layout="centered")

st.title("📅 PWH Duty Roster to Google Calendar Converter")
st.write("This tool converts your XLSX duty roster into a CSV file that you can directly import into Google Calendar.")

# --- Step 1: File Upload ---
st.header("Step 1: Upload Your Roster File")
st.info("Your file name must contain the month and year (e.g., `Duty August 2025 Final.xlsx`)")
uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")

# --- App Logic: Proceeds only if a file is uploaded ---
if uploaded_file is not None:
    # Use io.BytesIO to handle the uploaded file in memory
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    
    # Extract Month and Year from Filename here for use in output filename
    match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})', uploaded_file.name, re.IGNORECASE)
    if not match:
        st.error(f"Error: Could not find a month and year in the filename '{uploaded_file.name}'.")
        st.info("Please ensure the filename is formatted like 'Duty August 2025.xlsx'.")
        st.stop() # Stop execution if filename is incorrect

    filename_month_name, filename_year_str = match.groups()
    
    try:
        xls = pd.ExcelFile(file_bytes)
        all_sheet_names = xls.sheet_names

        # Requirement 1: Filter sheet names to only include those with "Duty"
        filtered_sheet_names = [name for name in all_sheet_names if "Duty" in name]
        
        if not filtered_sheet_names:
            st.error("No sheets containing 'Duty' found in the uploaded file. Please check your sheet names.")
            st.stop()

        # --- Step 2: Sheet Selection ---
        st.header("Step 2: Select the Duty Sheet")
        selected_sheet = st.selectbox("Which sheet contains the duty roster?", filtered_sheet_names)

        if selected_sheet:
            # Determine header row to read the names correctly
            header_row_map = {"Duty - Senior": 3, "Duty_MO": 2, "Duty - Part time": 2}
            skiprows = header_row_map.get(selected_sheet, 2)
            
            # Requirement 2: Read just the header row from the *selected sheet* to get potential names.
            # This ensures names are specific to the chosen sheet.
            df_for_cols = pd.read_excel(file_bytes, sheet_name=selected_sheet, skiprows=skiprows, nrows=0)
            
            # Requirement 3 & 4: Define patterns for non-name headers.
            # Simplified based on "no numbers", "not single letter", and new exclusions.
            non_name_patterns = [
                r'unnamed', r'week', r'day', r'am', r'pm', r'consultant', r'specialist',
                r'final call', r'part-time', r'locum', r'full', r'half', r'total',
                r'ic', r'ae', r'aw', r'pw', r'rat', r'leave', r'ac',
                r'intern', r'rs', r'rt', r'smo', r'cons', r'emw', r'new',
                r'rotation', r'sur', r'ort', r'diir', r'visiting dr',
                r'fall', r'shift', r'qeh', r'mo', # Added 'qeh' and 'mo' for robustness
                r'fm', # Added 'fm'
                r'ortho' # Added 'ortho'
            ]
            
            # Filter out columns that are not names
            duty_names = []
            for col in df_for_cols.columns:
                col_lower = str(col).lower().strip()
                
                # Exclude if it's an unnamed column
                if 'unnamed' in col_lower:
                    continue
                
                # Exclude if it's a single letter
                if len(col_lower) == 1:
                    continue
                
                # Exclude if it contains a digit
                if any(char.isdigit() for char in col_lower):
                    continue
                
                # Exclude if it matches any non-name pattern
                is_non_name = False
                for pattern in non_name_patterns:
                    if re.search(pattern, col_lower):
                        is_non_name = True
                        break
                if is_non_name:
                    continue
                
                duty_names.append(col)

            if not duty_names:
                st.error(f"Could not find any staff names in the sheet '{selected_sheet}'. Please check the file format or if names are in the expected header row.")
            else:
                # --- Step 3: Name Selection ---
                st.header("Step 3: Select Your Name")
                selected_name = st.selectbox("Whose duty do you want to export?", duty_names)

                # --- Step 4: Process and Download ---
                st.header("Step 4: Generate and Download")
                if st.button(f"Generate Calendar for {selected_name}"):
                    with st.spinner("Processing your file..."):
                        final_df, date_info = process_roster(file_bytes, selected_sheet, selected_name, uploaded_file.name)

                    if final_df is not None and date_info is not None:
                        st.success("✅ Your calendar file is ready!")
                        st.dataframe(final_df)

                        # Convert DataFrame to CSV string, then encode to bytes
                        csv_bytes = final_df.to_csv(index=False).encode('utf-8')

                        # Create the download button with correct month and year from filename
                        output_filename = f"{selected_name.replace(' ', '_')}-{selected_sheet}-{filename_month_name}-{filename_year_str}-Calendar.csv"
                        st.download_button(
                            label="📥 Download .csv File",
                            data=csv_bytes,
                            file_name=output_filename,
                            mime='text/csv',
                        )

    except Exception as e:
        st.error(f"Failed to read the Excel file. It might be corrupted or in an unsupported format. Error: {e}")
