import openpyxl
import smtplib
import pandas as pd

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from bs4 import BeautifulSoup

# SMTP information
email_user = "########" # Add your SMTP user info
email_pass = "########" # Add your SMTP pass info

# Get Excel spreadsheet
filename = ".xslx file location here"

# Get active sheet
wb = openpyxl.load_workbook(filename)
ws = wb.active

class Items:
  def __init__(self, item_code, item_descript, order_qty):
    self.item_code = item_code.split(',')
    self.item_descript = item_descript.split(';')
    self.order_qty = order_qty.split(',')
    
    self.item_list = []
    
    for i in range(0, len(self.item_code)):
      temp = '''
        <tr class="item">
          <td class="add-info"><p>'''+ self.item_code[i] +'''</p></td>
          <td class="td-wide"><p>'''+ self.item_descript[i] +'''</p></td>
          <td class="add-info align-center"><p>'''+ self.order_qty[i] +'''</p></td>
        </tr>
      '''
      self.item_list.append(temp)
  
def build_email(ordnum, ordate, ship_name, ship_addr1, ship_addr2, ship_addr3, ship_city, ship_state, ship_zipcode, c):
  # Create email body
  html = '''\
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
      <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
          <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style type="text/css">
          </style>
        </head>
        <body>
          <h1 style="text-align: center;">Order Status</h1>
        </body>
      </html>
  '''
  
  # Using BeautifulSoup to get rid of None type when cell value under ship addresses are blank
  soup = BeautifulSoup(html, 'html.parser')
  
  return soup.prettify()
  

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
