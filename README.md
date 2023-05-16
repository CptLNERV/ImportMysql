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
                                            database="法国激活imei")

cursor =connection.cursor()
```


Define the functions to be used primarily
where getFile_inFolder is used to get the name of the file present in the file, combined with the folder path, so that it can be read
ImportActivationFil's input parameter is the absolute path to the file
```
def getFile_inFolder(FolderPath):

def  ImportActivationFile(FilePath):
```



使用ImportActivationFile function获得excel 数据源中的数据
```
def  ImportActivationFile(FilePath):
    print(excel_sheet.sheet_names())
    for sh in excel_sheet.sheet_names():
        table = excel_sheet.sheet_by_name(sh)

    # excel 列顺序
    #imei	是否激活	激活时间	入库时间	入库库位	入库单号	出库时间	出库库位	出库单号	最终客户	激活国家	激活省	激活城市	SKU编码	SKU名称	物料编码	营销机型	颜色

    j =1
    for i in range(0,table.ncols):
        print   (table.cell(1,i).value,end =";")
    print()

    for j in range(0,table.nrows):

        imei = table.cell(j,0).value
        是否激活= table.cell(j,1).value
        激活时间= table.cell(j,2).value
        入库时间= table.cell(j,3).value
        入库库位= table.cell(j,4).value
        入库单号= table.cell(j,5).value
        出库时间= table.cell(j,6).value
        出库库位= table.cell(j,7).value
        出库单号= table.cell(j,8).value
        最终客户= table.cell(j,9).value
        激活国家= table.cell(j,10).value
        激活省= table.cell(j,11).value
        激活城市= table.cell(j,12).value
        SKU编码= table.cell(j,13).value
        SKU名称= table.cell(j,14).value
        物料编码= table.cell(j,15).value
        营销机型= table.cell(j,16).value
        颜色= table.cell(j,17).value

        Activation_Range_Value=(imei,是否激活,激活时间,入库时间,入库库位,入库单号,出库时间,出库库位,出库单号,最终客户,激活国家,激活省,激活城市,SKU编码,SKU名称,物料编码,营销机型,颜色)
 ```
 
 
 使用SQL insert 插入mysql 数据库
 
 ```
      insert_query = '''
        Insert into 法国激活imei.法国激活2022imei(
        imei,是否激活,激活时间,入库时间,入库库位,入库单号,出库时间,出库库位,出库单号,最终客户,激活国家,激活省,激活城市,SKU编码,SKU名称,物料编码,营销机型,颜色
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
 
 
