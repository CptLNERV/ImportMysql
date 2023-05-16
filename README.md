# Use PythonScript ImportMysql

Purpose: This script is for batch processing of huge sales data (500k+), direct use of this amount of data, direct use of the visualisation tool would result in very long loading times (up to several hours) or direct error reporting  
Solution: Use the Python script to import the original sales data into Mysql data and use sql in the visualisation tool to get the data needed

First import the libraries you need
```
import pymysql
import xlrd
import os
```

open a cursor
```
connection =pymysql.connect(host="localhost",
                                            user = "root",
                                            password="199288Lj",
                                            database="aleDateDetail")

cursor =connection.cursor()
```


Define the functions to be used primarily
where getFile_inFolder is used to get the name of the file present in the file, combined with the folder path, so that it can be read
ImportActivationFil's input parameter is the absolute path to the file
```
def getFile_inFolder(FolderPath):

def  ImportActivationFile(FilePath):
```



Use the ImportActivationFile function to get data from an excel data source
```
def  ImportActivationFile(FilePath):
    print(excel_sheet.sheet_names())
    for sh in excel_sheet.sheet_names():
        table = excel_sheet.sheet_by_name(sh)

    # Order of columns Excel 
    #imei	Activated 	Activated Time 	Inbound Time	 Inbound locaiton 	Stock No. 	Outbound Time	 Outbound warehouse	Dispatch Note  Number 	Final customer  	country	Province  	City 	 SKU Code	SKU Name	Material Code 	Marketing Model	Colour

    j =1
    for i in range(0,table.ncols):
        print   (table.cell(1,i).value,end =";")
    print()

    for j in range(0,table.nrows):

        imei = table.cell(j,0).value
        Activated= table.cell(j,1).value
        Activated Time= table.cell(j,2).value
        Inbound Time= table.cell(j,3).value
        Inbound locaiton= table.cell(j,4).value
        Stock No.= table.cell(j,5).value
        Outbound Time = table.cell(j,6).value
        Outbound warehouse = table.cell(j,7).value
        Dispatch Note = table.cell(j,8).value
        Final customer = table.cell(j,9).value
        country = table.cell(j,10).value
        Province= table.cell(j,11).value
        City = table.cell(j,12).value
        SKU Code = table.cell(j,13).value
        SKU Name = table.cell(j,14).value
        Material Code= table.cell(j,15).value
        Marketing Mode = table.cell(j,16).value
        Colour = table.cell(j,17).value

        Activation_Range_Value=(imei,Activated, Activated Time,Inbound Time,Inbound locaiton,Stock No.,Outbound Time,Outbound warehouse,Dispatch Note, Final customer,country,Province,City,SKU Code,SKU Name,Material Code,Marketing Mode,Colour)
 ```
 
 
Insert into a mysql database 
 
 ```
      insert_query = '''
        Insert into SaleDate.SaleDateDetail(
        imei,Activated, Activated Time,Inbound Time,Inbound locaiton,Stock No.,Outbound Time,Outbound warehouse,Dispatch Note, Final customer,country,Province,City,SKU Code,SKU Name,Material Code,Marketing Mode,Colour
        )
        values(
        %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s
        )
        '''
        try:
            cursor.execute(insert_query,Activation_Range_Value)
            connection.commit()

        except Exception as e:
            print(f"insert fail")
 
 ```
 
 
