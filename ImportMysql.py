import pymysql
import xlrd
import os

FileList = []
FloderPath = r""

connection =pymysql.connect(host="localhost",
                                            user = "root",
                                            password="199288Lj",
                                            database="法国激活imei")

cursor =connection.cursor()

sql = '''
SELECT * FROM 法国激活imei.法国激活2022imei;

'''


# 这是测试区域
print(cursor.execute(sql))
resulat = cursor.fetchone()
print(resulat)
# resultat = cursor.execute(sql)
# for r in resultat:
#     print(r)


def getFile_inFolder(FolderPath):
    pass


def  ImportActivationFile(FilePath):


    # 读取单个激活Excel文件
    FilePath = r"C:\Users\JieLIAO\OneDrive - OPPO FRANCE\Documents - supplychain\入库IMEI&激活IMEI\FR激活IMEI\2022\202201January-activatedImei20220809172047.xlsx"
    excel_sheet = xlrd.open_workbook(FilePath)

    # 读取同一文件夹下多个文件（数据结构相同/sheet名相同 文件名不需要相同）



    print(excel_sheet.sheet_names())
    for sh in excel_sheet.sheet_names():
        table = excel_sheet.sheet_by_name(sh)

    # 列顺序
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
        print(Activation_Range_Value)

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




if __name__ == "__main__":

    #FilePath = ""
    #ImportActivationFile(FilePath)

    FolderPath=""

    pass

cursor.close()
connection.close()
