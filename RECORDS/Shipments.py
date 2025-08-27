import pyodbc
import pandas as pd

import os, os.path
import win32com.client
from sys import exit
import datetime

now = datetime.datetime.now()

#def end_of_month(dt):
    #todays_month = dt.month
    #tomorrows_month = (dt + datetime.timedelta(days=1)).month
    #return tomorrows_month != todays_month

#if end_of_month(now):
    #pass
#else:
    #exit()
    
from datetime import datetime

datestring = datetime.strftime(datetime.now(), '%m_%d_%Y %H%M')
path = r"S:\Shipments"
driver_1 = "SQL Server"
server_1 = "seza-sql2"
database_1 = "EZFlow"
user_1 = "EZflow_Admin"
password_1 = "GDerp401"

connection = pyodbc.connect(driver=driver_1,server=server_1,database=database_1,uid=user_1,pwd=password_1)
cursor = connection.cursor()
query = "SELECT * FROM ShipmentsMaster"
cursor.execute(query)
rows = cursor.fetchall()
columns = [column[0] for column in cursor.description]
df = pd.DataFrame.from_records(rows, columns=columns)
cursor.close()
connection.close()

connection = pyodbc.connect(driver=driver_1,server=server_1,database=database_1,uid=user_1,pwd=password_1)
cursor = connection.cursor()
query = "SELECT * FROM CatalogsDetail"
cursor.execute(query)
rows = cursor.fetchall()
columns = [column[0] for column in cursor.description]
df1 = pd.DataFrame.from_records(rows, columns=columns)
cursor.close()
connection.close()

connection = pyodbc.connect(driver=driver_1,server=server_1,database=database_1,uid=user_1,pwd=password_1)
cursor = connection.cursor()
query = "SELECT * FROM FlowsProcess"
cursor.execute(query)
rows = cursor.fetchall()
columns = [column[0] for column in cursor.description]
df4 = pd.DataFrame.from_records(rows, columns=columns)
cursor.close()
connection.close()

connection = pyodbc.connect(driver=driver_1,server=server_1,database=database_1,uid=user_1,pwd=password_1)
cursor = connection.cursor()
query = "SELECT * FROM Flows"
cursor.execute(query)
rows = cursor.fetchall()
columns = [column[0] for column in cursor.description]
df3 = pd.DataFrame.from_records(rows, columns=columns)
cursor.close()
connection.close()

connection = pyodbc.connect(driver=driver_1,server=server_1,database=database_1,uid=user_1,pwd=password_1)
cursor = connection.cursor()
query = "SELECT * FROM ShipmentsDetail"
cursor.execute(query)
rows = cursor.fetchall()
columns = [column[0] for column in cursor.description]
df5 = pd.DataFrame.from_records(rows, columns=columns)
cursor.close()
connection.close()

base = pd.merge(df,df4,how ='left',left_on='CurrentFlowProcessID',right_on='FlowProcessID')
base1 = pd.merge(base,df1,how ='left',left_on='StatusID',right_on='CatalogDetailID')
base2= pd.merge(base1,df3,how ='left',left_on='FlowID_x',right_on='FlowID')
base3= pd.merge(base2,df5)

base2 = base2.loc[base2['ValueID']=='open']
base2 = base2[['ShipmentNumber','FlowName','ProcessName','CreatedBy','ProcessName','ValueID','DateLastMaint_x','DateAdded_x','TotalAmount']]

base3 = base3.loc[base3['ValueID']=='open']
base3 = base3[['ShipmentNumber','WorkOrder','FlowName','ProcessName','CreatedBy','ProcessName','ValueID','DateLastMaint_x','DateAdded_x','TotalAmount']]

base2['ProcessName'] = base2['ProcessName'].fillna('Creacion')
base2['TotalAmount'] = pd.to_numeric(base2['TotalAmount'])

base2 = base2.loc[:,~base2.columns.duplicated()].copy()

if not os.path.exists(path):
		os.makedirs(path)
        
base3["TotalAmount"] = pd.to_numeric(base2["TotalAmount"])

base3['ProcessName'] = base3['ProcessName'].fillna('Creacion')

base3 = base3.loc[:,~base3.columns.duplicated()].copy()

if not os.path.exists(path):
		os.makedirs(path)
        
base3["TotalAmount"] = pd.to_numeric(base3["TotalAmount"])

base2.to_excel(r'J:\Departments\Administration\Program Management\PPCP\1.Planeación\15.Jose Valles\Projects\Shipments\Shipments.xlsx', 'New', index=False)
base3.to_excel(r'J:\Departments\Administration\Program Management\PPCP\1.Planeación\15.Jose Valles\Projects\Shipments\ShipmentsDetails.xlsx', 'New', index=False)

import datetime

this_hour = datetime.datetime.now().hour

if this_hour > 6 or this_hour < 3:
    print("Shipments Report")
    if os.path.exists(r"J:\Departments\Administration\Program Management\PPCP\1.Planeación\2.Report\Master Planning Reports\20. Shipments\Macro\Shipments.xlsm"):
        excel_macro = win32com.client.DispatchEx("Excel.Application")
        excel_path = os.path.expanduser(r"J:\Departments\Administration\Program Management\PPCP\1.Planeación\2.Report\Master Planning Reports\20. Shipments\Macro\Shipments.xlsm")
        workbook = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly =1)
        excel_macro.Application.Run("'Shipments.xlsm'!Module1.DShp_Inq")
        excel_macro.Application.Run("'Shipments.xlsm'!Module1.Shp_Inq")
        excel_macro.Application.Run("'Shipments.xlsm'!Module1.PIV")
        #if end_of_month(now):
            #print("Final month")
        excel_macro.Application.Run("'Shipments.xlsm'!Module1.SendEmail")
        #else: 
            #print("No final month")
            #excel_macro.Application.Run("'Shipments.xlsm'!Module1.Save_Excel")
        excel_macro.Application.Quit()
        del excel_macro