import pandas as pd
import os

INPUT_FOLDER = "input_excels"
OUTPUT_FOLDER = "output"
OUTPUT_FILE = "summary_report.xlsx"

all_data = []

for file in os.listdir(INPUT_FOLDER):
    if file.endswith(".xlsx"):
        file_path = os.path.join(INPUT_FOLDER, file)
        df = pd.read_excel(file_path)
        all_data.append(df)

merged_df = pd.concat(all_data, ignore_index=True)

summary = {
    "Total Amount": merged_df["Amount"].sum(),
    "Average Amount": merged_df["Amount"].mean()
}

summary_df = pd.DataFrame(list(summary.items()), columns=["Metric", "Value"])

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILE)

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    merged_df.to_excel(writer, sheet_name="Merged Data", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

print("Report created successfully!")
