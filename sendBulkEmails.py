import openpyxl
import pandas as pd

# SMTP information
email_user = "########" # Add your SMTP user info
email_pass = "########" # Add your SMTP pass info

# Get Excel spreadsheet
filename = ".xslx file location here"

# Get active sheet
wb = openpyxl.load_workbook(filename)
ws = wb.active

def send_email(user, password, sent_from, to, msg):
  # Enter SMTP credentials
  server = smtplib.SMTP('#####', 465) # Enter SMTP (gmail, outlook, etc) Settings and SMTP port number
  server.ehlo()
  server.starttls()
  server.ehlo()
  server.login(user, password)
  server.sendmail(sent_from, to, msg.as_string())
  server.quit()
  
def get_data(colname, colnum):
  colname = []
  # Count variable for testing
  count = 1
  
  # Go through each row in .xlsx file
  # If there are headers, then skip first row
  for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    #print(f"Row {count}: {row[colnum].value}")
    colname.append(str(row[colnum].value))
    count += 1
  
  return colname
