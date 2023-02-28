import polars as pl
import pandas as pd
import keyring
import cx_Oracle
from openpyxl import load_workbook
from datetime import date
from datetime import datetime
import os
import re
import subprocess

time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

pw = keyring.get_password('oracle', 'password')
user = keyring.get_password('oracle', 'username')
host = keyring.get_password('oracle', 'host')
port = keyring.get_password('oracle', 'port')
service = keyring.get_password('oracle', 'service')


def validate_vpn():
    """Validate whether or not the user is connect to VPN. If not, display message and exit program. Required to be connected for Oracle access"""

    addresses = os.popen(
        'IPCONFIG | FINDSTR /R "Ethernet adapter Local Area Connection .* Address.*[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*"')

    first_eth_address = re.search(
        r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', addresses.read()).group()

    host = first_eth_address
    ping = subprocess.Popen(["ping.exe", "-n", "1", "-w",
                            "1", host], stdout=subprocess.PIPE).communicate()[0]

    if ('unreachable' in str(ping)) or ('timed' in str(ping)) or ('failure' in str(ping)):
        ping_chk = 0
    else:
        ping_chk = 1

    if ping_chk == 1 and host.startswith('10.'):
        return True
    else:
        raise ValueError('VPN is Not Connected')


def get_queries() -> pl.DataFrame:
  """Run the 10 needed queries to develop the dashboard, compile to a list which also has the row and column settings. This is written to a Polars Dataframe"""
  
    cx_Oracle.init_oracle_client(
        lib_dir=r"path/to/instantclient_21_9")
    print("Login Successful")
    conn = f'oracle://{user}:{pw}@{host}:{port}/{service}'

    query1 = """select * from table"""

    query2 = """select * from table"""

    query3 = """select * from table"""

    query4 = """select * from table"""

    query5 = """select * from table"""

    query6 = """select * from table"""

    query7 = """select * from table"""
    
    query8 = """select * from table"""
    
    query9 = """select * from table"""

    query10 = """select * from table"""

    # List of all the queries data, row location, column location, and type of setting. The row and column settings are used when writing to excel later.
    all_queries = [[query1, 2, 0, 'Settings'], [query2, 2, 4, 'Settings'], [query3, 2, 10, 'Settings'], [
        query4, 2, 15, 'Settings'], [query5, 2, 29, 'Settings'], [query6, 2, 33, 'Settings'], [query7, 2, 46, 'Settings'], [query8, 2, 51, 'Settings'], [query9, 2, 56, 'Settings'], [query10, 2, 64, 'Settings']]
    dict_df = {}

    for index, my_list in enumerate(all_queries):
        dict_df[f'df{index}'] = (pl.read_sql(
            my_list[0], conn), my_list[1], my_list[2])
        print(f"Extracting data from {my_list[3]}.")

    return dict_df


def write_to_excel(data):
  """Opening template to use for the dashboard, convert Polars DF to Pandas DF and save as new file"""
    today = date.today()

    book = load_workbook(
        "path/to/template/Dashboard_Template.xlsm", keep_vba=True)

    with pd.ExcelWriter(
            f'path/to/new_file/New_Dashboard_{today}.xlsm', engine='openpyxl') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        print("Writing data to the Dashboard")
        for name, value in data.items():
            df, row, col = value
            df = data[name]
            polars_df = df[0] 
            df = polars_df.to_pandas() # Convert polars.df to pandas.df
            df.to_excel(writer, index=False, header=False,
                        sheet_name="Sheet1", startcol=col, startrow=row)

def email_document():
  """Send email to appropriate stakeholders after completion, unable to use O365 so this opens Outlook locally to send."""
    today = date.today()

    # Outlook Instance
    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    # Email Object
    mailItem = olApp.CreateItem(0)
    mailItem.GetInspector
    mailItem.Subject = f'{today} - Subject'

    body = ("""Good morning,
            Here is an updated version of the Dashboard, have a great day!""")

    signature = mailItem.HTMLBody
    mailItem.HTMLBody = (body + signature)
    mailItem.To = 'email@address.com'

    mailItem.Attachments.Add(
        f'path/to/new_file/New_Dashboard_{today}.xlsm')

    # mailItem.Display()  # Testing how it looks
    mailItem.Send() # Send the email


if __name__ == '__main__':
    if validate_vpn():
        data = get_queries()
        write_to_excel(data)
        email_document()
