import os
import pandas as pd
from fpdf import FPDF

# Function to read Excel data
def read_excel(file_path):
    try:
        data = pd.read_excel(file_path, sheet_name=None)
        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# Create folder if it doesn't exist
def create_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

# Function to find the employee name column
def find_employee_name_column(dataframe):
    possible_columns = ['Employee Name', 'Name', 'Full Name']
    for column in possible_columns:
        if column in dataframe.columns:
            return column
    raise ValueError("No valid employee name column found.")

# Function to generate the PDF report
def generate_report(person_data, purchase_data, output_path, employee_name_column):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Add contact information icons as hyperlinks at the top center
    pdf.set_xy(80, 15)
    pdf.image("assets/github.png", link="https://github.com/avinesh-masih", w=10)
    pdf.set_xy(95, 15)
    pdf.image("assets/linkedin.png", link="https://www.linkedin.com/in/avineshlko/", w=10)
    pdf.set_xy(110, 15)
    pdf.image("assets/portfolio.png", link="https://avinesh-masih.github.io/", w=10)
    pdf.set_xy(125, 15)
    pdf.image("assets/email.png", link="mailto:skmasih11@gmail.com", w=10)
    pdf.set_xy(140, 15)
    pdf.image("assets/paypal.png", link="https://paypal.me/AVINESHMASIH", w=10)
    pdf.set_xy(155, 15)
    pdf.image("assets/buymeacoffee.png", link="https://buymeacoffee.com/avineshlko", w=10)

    pdf.ln(20)  # Add space after the icons

    # Add company logo and name
    pdf.image("assets/employee_logo.png", x=10, y=8, w=33)  # Adjust the path and size as needed
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="AVINESH MASIH PROJECT ", ln=True, align="C")
    pdf.ln(20)

    pdf.cell(200, 10, txt="Personal Report", ln=True, align="C")
    pdf.ln(10)

    # Add formatted table with employee details
    pdf.set_font("Arial", size=10)
    pdf.set_fill_color(200, 220, 255)
    
    # Dynamically add headers based on the columns in person_data
    column_width = 190 / len(person_data.columns)  # Adjust column width to fit within page margins
    for column in person_data.columns:
        pdf.cell(column_width, 10, column, border=1, align='C', fill=True)
    pdf.ln()

    for index, row in person_data.iterrows():
        for column in person_data.columns:
            pdf.cell(column_width, 10, str(row[column]), border=1, align='C')
        pdf.ln()

    pdf.ln(10)
    pdf.cell(200, 10, txt="Item Purchase Details", ln=True, align="C")
    pdf.ln(5)

    # Add formatted table with purchase details
    pdf.set_font("Arial", size=10)
    pdf.set_fill_color(200, 220, 255)
    
    # Remove the employee name column from purchase_data
    purchase_data = purchase_data.drop(columns=[employee_name_column])
    
    # Dynamically add headers based on the columns in purchase_data
    column_width = 190 / len(purchase_data.columns)  # Adjust column width to fit within page margins
    for column in purchase_data.columns:
        pdf.cell(column_width, 10, column, border=1, align='C', fill=True)
    pdf.ln()

    for index, row in purchase_data.iterrows():
        for column in purchase_data.columns:
            pdf.cell(column_width, 10, str(row[column]), border=1, align='C')
        pdf.ln()

    # Output the PDF
    pdf.output(output_path)

# Main function
def main():
    excel_file = "assets/employees.xlsx"  # Replace with your file path
    data = read_excel(excel_file)

    if data is None:
        print("Failed to load data.")
        return

    # Extract employee data and item purchase data
    employee_data = data['Employee Data']
    purchase_data = data['Item Purchases']

    # Find the employee name column
    employee_name_column = find_employee_name_column(employee_data)

    overwrite_all = False

    # Create a folder for each employee and generate their reports
    for index, row in employee_data.iterrows():
        person_name = row[employee_name_column]  # Employee name is used as folder name
        folder_name = f"reports/{person_name}"
        create_folder(folder_name)

        # Filter purchase data for the specific employee
        employee_purchase_data = purchase_data[purchase_data[employee_name_column] == person_name]

        # Generate the PDF report with employee details
        report_path = os.path.join(folder_name, f"{person_name}_report.pdf")
        
        if os.path.exists(report_path) and not overwrite_all:
            overwrite = input(f"The report for {person_name} already exists. Do you want to overwrite it? (yes/no/all): ")
            if overwrite.lower() == 'no':
                print(f"Skipping report generation for {person_name}.")
                continue
            elif overwrite.lower() == 'all':
                overwrite_all = True

        generate_report(employee_data[employee_data[employee_name_column] == person_name], employee_purchase_data, report_path, employee_name_column)

        print(f"Report generated for {person_name} at {report_path}")

if __name__ == "__main__":
    main()