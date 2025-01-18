import os
import pandas as pd
from fpdf import FPDF

def read_excel(file_path):
    """
    Reads the Excel file and returns it as a pandas DataFrame.
    """
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def create_folder(folder_name):
    """
    Creates a folder if it doesn't already exist.
    """
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

def generate_report(person_data, output_path):
    """
    Generates a PDF report for an individual.
    """
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="Personal Report", ln=True, align="C")
    pdf.ln(10)

    for key, value in person_data.items():
        pdf.cell(0, 10, txt=f"{key}: {value}", ln=True)

    pdf.output(output_path)

def main():
    # Step 1: File path of the Excel file
    excel_file = '\\notebook\\employees.xlsx'  # Update this with the path to your Excel file

    # Step 2: Read the Excel file
    data = read_excel(excel_file)

    if data is None:
        print("Failed to load data.")
        return

    # Step 3: Iterate through each row in the DataFrame
    for index, row in data.iterrows():
        person_name = row['Name']  # Update column name based on your Excel file
        folder_name = f"reports/{person_name}"

        # Step 4: Create a folder for the person
        create_folder(folder_name)

        # Step 5: Generate the report
        report_path = os.path.join(folder_name, f"{person_name}_report.pdf")
        generate_report(row.to_dict(), report_path)

        print(f"Report generated for {person_name} at {report_path}")

if __name__ == "__main__":
    main()
