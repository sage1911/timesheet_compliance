import pandas as pd
import openpyxl
import re
import os
import sys

# Define base path for files
base_path = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_path, "timesheet.xlsx")
project_file = os.path.join(base_path, "projectcode.xlsx")
output_file_path = os.path.join(base_path, "studycounts.xlsx")

# Check if files exist
if not os.path.exists(file_path):
    print("Error: Completion report missing. Please place 'timesheet.xlsx' in the same folder.")
    sys.exit(1)

if not os.path.exists(project_file):
    print("Error: Project code file missing. Please place 'projectcode.xlsx' in the same folder.")
    sys.exit(1)

try:
    # Load the main report
    df = pd.read_excel(file_path, engine="openpyxl")
    project_df = pd.read_excel(project_file, engine="openpyxl")
except Exception as e:
    print(f"Error: Unable to read Excel files. Details: {e}")
    sys.exit(1)

# Ensure required columns exist
if "Study" not in df.columns:
    print("Error: 'Study' column is missing from the main report.")
    sys.exit(1)

if "Study" not in project_df.columns or "Project Code" not in project_df.columns:
    print("Error: The project code file must contain 'Study' and 'Project Code' columns.")
    sys.exit(1)

# Normalize Study column (strip spaces and convert to uppercase for consistency)
df["Study"] = df["Study"].astype(str).str.strip().str.upper().apply(lambda x: re.sub(r"\s+", " ", x))
project_df["Study"] = project_df["Study"].astype(str).str.strip().str.upper()

# Merge reports
df = df.merge(project_df, on="Study", how="left")

# Fill missing Project Codes with "Unknown"
df["Project Code"].fillna("Unknown", inplace=True)

# Count occurrences by Study
study_counts = df.groupby(["Study", "Project Code"]).size().reset_index(name="Count")

# Calculate "Time to Code"
study_counts["Time to Code"] = study_counts["Count"] * 4 / 60

# Export results to Excel
try:
    study_counts.to_excel(output_file_path, sheet_name="Counts", index=False, engine="openpyxl")
    print(f"Exported: {output_file_path} with Study, Project Code, Count, and Time to Code.")
except Exception as e:
    print(f"Error: Unable to write to Excel file. Details: {e}")
    sys.exit(1)

# Wait for user input before exiting
input("Press Enter to Exit...")
