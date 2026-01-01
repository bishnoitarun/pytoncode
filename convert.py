import os
import pandas as pd

# Set the directory containing the CSV files
csv_folder = './In_20201103/'   # change this to your folder path
excel_folder = './output/In_20201103/'  # desired output folder

# Create output folder if not exists
os.makedirs(excel_folder, exist_ok=True)

# Excel row limit
MAX_ROWS = 1000000

# Loop through all CSV files in the folder
for filename in os.listdir(csv_folder):
    if filename.endswith('.csv'):
        csv_file_path = os.path.join(csv_folder, filename)

        try:
            # Read the CSV file in chunks with low_memory=False to avoid DtypeWarning
            chunk_iter = pd.read_csv(
                csv_file_path,
                low_memory=False,
                chunksize=MAX_ROWS
            )

            # Define the Excel file path
            excel_file_path = os.path.join(
                excel_folder,
                f"{os.path.splitext(filename)[0]}.xlsx"
            )

            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                for i, chunk in enumerate(chunk_iter):
                    sheet_name = f"Sheet{i+1}"
                    chunk.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False
                    )

            print(f"Converted {filename} to {os.path.basename(excel_file_path)}")

        except Exception as e:
            print(f"Error processing {filename}: {e}")

print("Conversion completed")
