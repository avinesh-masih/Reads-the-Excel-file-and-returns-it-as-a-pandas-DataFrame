import unittest
import pandas as pd
import os
from generate_employee_reports import read_excel, find_employee_name_column, generate_report, create_folder

class TestGenerateReport(unittest.TestCase):

    def test_read_excel(self):
        data = read_excel('assets/employees.xlsx')
        self.assertIsNotNone(data)
        self.assertIn('Employee Data', data)
        self.assertIn('Item Purchases', data)

    def test_find_employee_name_column(self):
        data = pd.DataFrame({'Employee Name': ['John Doe', 'Jane Doe']})
        column = find_employee_name_column(data)
        self.assertEqual(column, 'Employee Name')

    def test_generate_report(self):
        person_data = pd.DataFrame({'Employee Name': ['John Doe'], 'Age': [30]})
        purchase_data = pd.DataFrame({'Employee Name': ['John Doe'], 'Item': ['Laptop'], 'Price': [1000]})
        create_folder('reports/John_Doe')
        generate_report(person_data, purchase_data, 'reports/John_Doe/John_Doe_report.pdf', 'Employee Name')
        self.assertTrue(os.path.exists('reports/John_Doe/John_Doe_report.pdf'))

if __name__ == '__main__':
    unittest.main()