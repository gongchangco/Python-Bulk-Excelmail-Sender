import openpyxl
import pandas as pd

# SMTP information
email_user = "########" # Add your SMTP user info
email_pass = "########" # Add your SMTP pass info

# Get Excel spreadsheet
filename = ".xslx file location here"

wb = openpyxl.load_workbook(filename)
ws = wb.active
