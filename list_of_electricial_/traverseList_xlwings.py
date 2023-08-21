import xlwings as xw
import global_variable
import win32api
from datetime import datetime
import time
import os

ALl_SEARCH = 1

global_variable.current_directory = os.getcwd()
global_variable.relative_path = "data/filter.ini"
global_variable.ini_file_path = os.path.join(global_variable.current_directory, global_variable.relative_path)


def ReadIni(): 
    try:
        from configparser import ConfigParser
        conf = ConfigParser()
        conf.read(global_variable.ini_file_path) 
        global_variable.eFilePath = conf['List']['filePath']
        global_variable.eSavePath = conf['List']['newFilePath']
        global_variable.maxRows = conf['List']['maxRows']
        global_variable.maxRows = int(global_variable.maxRows)
    except Exception as e:
        win32api.MessageBox(0, str(e), "error" , 0)
        print(e)
        return -1

ReadIni()

def traverse_xlwings2():
    localTime = datetime.now()
    path = "C:\\Project\\py_excel\\list_of_electricial\\new_JPN_last.xlsm"
    path2 = "C:\\Project\\py_excel\\list_of_electricial\\new_JPN_last_2.xlsm"
    # app = xw.App(visible=True)
    
    
    # wb = app.books.open(path)  # 替换为实际的 Excel 文件路径
    # wb = xw.Book(global_variable.eFilePath)  
    wb = xw.Book(path)
    
    sheet_1 = wb.sheets['LIST']  
    sheet_2 = wb.sheets['DMI'] 
    sheet_3 = wb.sheets['3M_BUY']
    sheet_4 = wb.sheets['share']
    sheet_5 = wb.sheets['SAP']
    
    dmi_b_column = sheet_2.range('B2:B5000').value
    dmi_c_colunm = sheet_2.range('C2:C5000').value
    dmi_e_column = sheet_2.range('E2:E5000').value
    list_c_column = sheet_1.range('C13:C2500').value
    buy_d_column = sheet_3.range('D2:D84404').value
    Two_dimensional_array = sheet_1.range('C13:K2500').value
    sap_a_column = sheet_5.range('A2:A46000').value
    
    # test_c_column = sheet_1.range('C13:K50').value
    # for rowNo, rows in enumerate(test_c_column, start=13):
    #     for index, value in enumerate(rows, start=3):
    #         print(f'第{rowNo}行，第{index}列，值：{value}') 
    # print("ok")
    
    lastest_time_dic = {}
    # for rows, (value, value2) in enumerate(dmi_b_column, dmi_e_column, start=2):
    for row_num, (dmi_value, dmi_time) in enumerate(zip(dmi_b_column, dmi_e_column), start=2):
        if isinstance(dmi_time, datetime):
            if dmi_time <= localTime:
                if dmi_value in lastest_time_dic:
                    if dmi_time > lastest_time_dic[dmi_value][0]:
                        lastest_time_dic[dmi_value] = (dmi_time, row_num)
                else:
                    lastest_time_dic[dmi_value] = (dmi_time, row_num)
                    # print(f'dmi:{dmi_value}, date:{dmi_time}, rows:{row_num}')

    # print(lastest_time_dic)
    print(f'时间存储OK')        
    
    
    # for rows, value in enumerate(list_c_column, start=13):
    #     # print(f'rows:{rows}')
    #     if value in buy_d_column:
    #         print(f'value: {value}, rows: {rows}')
    
    
    # for value in dmi_c_colunm:
    #     if value == 'DMI_F1G1C333BKV':
    #         print("OK")
    #         time.sleep(5)
    #     print(f'value:{value}')
    
    if ALl_SEARCH == 1:
        myfound = 0
        for rowN, whole_row in enumerate(Two_dimensional_array, start=13):
            progress =  (rowN-13) / 25
            print(f'DMI替换进度 {progress}%')   
            for colN, single_value in enumerate(whole_row, start=3):
                if single_value == None:
                    continue
                for rows2, value2 in enumerate(dmi_c_colunm, start=2):
                    if value2 == None:
                        continue
                    if single_value == value2:
                        myfound = 1
                        sheet_1.cells(rowN, 13).value = sheet_1.cells(rowN, colN).value
                        dmi_value = sheet_2.cells(rows2, 2).value  # 获取 B 列的值 
                        # if sheet_2.cells(rows2, 2).value in lastest_time_dic.keys:
                        if dmi_value in lastest_time_dic:
                            # sheet_1.cells(rows, 3).value = sheet_2.cells(lastest_time_dic[sheet_2.cells(rows2, 2).value][1], 3).value
                            dmi_row_num = lastest_time_dic[dmi_value][1]  # 获取对应的行号
                            sheet_1.cells(rowN, 3).value = sheet_2.cells(dmi_row_num, 3).value   
                            break          
                if myfound == 1:
                    myfound = 0
                    break
    if  ALl_SEARCH == 0:
        myfind = 0
        for rows, value in enumerate(list_c_column, start=13):
            if value == None:
                continue
            for rows2, value2 in enumerate(dmi_c_colunm, start=2):
                if value2 == None:
                    continue
                if value == value2:
                    myfind = 1
                    sheet_1.cells(rows, 21).value = sheet_1.cells(rows, 3).value
                    dmi_value = sheet_2.cells(rows2, 2).value  # 获取 B 列的值 
                    # if sheet_2.cells(rows2, 2).value in lastest_time_dic.keys:
                    if dmi_value in lastest_time_dic:
                        # sheet_1.cells(rows, 3).value = sheet_2.cells(lastest_time_dic[sheet_2.cells(rows2, 2).value][1], 3).value
                        dmi_row_num = lastest_time_dic[dmi_value][1]  # 获取对应的行号
                        sheet_1.cells(row_num, 3).value = sheet_2.cells(dmi_row_num, 3).value   
            if myfind == 0:
                # print(f'can not find these data: {value}, rows: {rows}')
                sheet_1.cells(rows, 3).color = (255, 0, 0) # red
                
            myfind = 0
    print(f'DMI替换进度 100%')    
    
    # found_buy3 = 0
    # new_list_c_column = sheet_1.range('C13:C2500').value
    
    # for rowN, value in enumerate(new_list_c_column, start=13):
    #     if value == None:
    #         continue
    #     for rowN2, value2 in enumerate(buy_d_column, start=2):
    #         if value2 == None:
    #             continue
    #         if value == value2:         # 3buy assignment
    #             for i in range(3, 2000):
    #                 if sheet_4.cells(i,4).value == None:
    #                     sheet_4.cells(i,4).value = value2 
    #                     found_buy3 = 1
    #                     break
    #                 if sheet_4.cells(i,4).value == value2:
    #                     sheet_4.cells(i,5).value += 1 
    #                     sheet_4.cells(i,6).value += 1 
    #                     break
    #     if found_buy3 == 0:
    #         for rowN3, value3 in enumerate()        
    
    # for rowN, value in enumerate(new_list_c_column, start=13):
    #     progress2 =  (rowN-13) / 25
    #     print(f'SAP/SHARE/NONE 进度 {progress2}%')   
    #     if value == None:
    #         continue       
    #     if value in buy_d_column:
    #         for i in range(3, 2000):
    #             if sheet_4.cells(i,4).value == None:
    #                 sheet_4.cells(i,4).value = value2 
    #                 break
    #             if sheet_4.cells(i,4).value == value2:
    #                 sheet_4.cells(i,5).value += 1 
    #                 sheet_4.cells(i,6).value += 1 
    #                 break
    #     elif value in sap_a_column:
    #         print(f'{value} in sap_a_column, rows: {rowN}')
    #     else:
    #         print(f'{value} is none, rows: {rowN}')
     
    wb.save(path2)
    wb.close()
    
     
def traverse_xlwings():

    wb = xw.Book(global_variable.eFilePath)  
    sheet_1 = wb.sheets['LIST']  
    sheet_2 = wb.sheets['DMI']  

    list_c_column_range = sheet_1.range('C13:C' + str(global_variable.maxRows))
    
    list_c_column = list_c_column_range.value
    # num_rows = list_c_column_range.rows.count
    # num_cols = list_c_column_range.columns.count
    print(f'list_c_column[0]:{len(list_c_column[0])}')
    num_rows = len(list_c_column)
    print(f'num_rows]:{num_rows}')
    num_cols = len(list_c_column[0])  if num_rows > 0 else 0 
    print(num_cols)
    # # dmi_c_column = sheet_2.range('C2:C' + str(sheet_2.cells.last_cell.row)).value
    # # dmi_e_column = sheet_2.range('E2:E' + str(sheet_2.cells.last_cell.row)).value
    # # dmi_b_column = sheet_2.range('B2:B' + str(sheet_2.cells.last_cell.row)).value
    
    # for row_num, value in list_c_column:
    #     print(f'row_num:{row_num}, list_c_column value:{list_c_column[value]}, Type: {type(value)}')
        # for row_num_2, value_2 in enumerate(dmi_c_column, start=2):
            
        #     pass
    # # 将时间字符串转换为 datetime 对象
    # e_dates = [datetime.strptime(date, '%Y-%m-%d %H:%M:%S') for date in e_column]

    # # 找到最新时间的索引
    # latest_index = e_dates.index(max(e_dates))

    # # 获取对应的值
    # latest_value = b_column[latest_index]

    # print("最新时间对应的值:", latest_value)

    # 关闭 Excel 文件
    wb.close()
    
def clear_share_none_new():
    path = "C:\\Project\\py_excel\\list_of_electricial\\new_JPN_last.xlsm"
    app = xw.App(visible=True)
    wb = app.books.open(path)  # 替换为实际的 Excel 文件路径
    try:
        sheet_names = ["share", "none", "new"]
        for sheet_name in sheet_names:
            sheet = wb.sheets[sheet_name]
            range_to_clear = sheet.range((3, 2), (500, 6))
            range_to_clear.clear_contents()
    except Exception as e:
        print("Error:", e)
    finally:
        wb.save()
        wb.close()
        app.quit()

if __name__ == '__main__':
    timeStart = time.time()
    
    # clear_share_none_new()
    traverse_xlwings2()
    
    timeEnd = time.time()
    timeCost = timeEnd - timeStart
    print(f'timecost: {timeCost}')







'''
    # for i in range(13, global_variable.maxRows):
    #     if sheet_1.cells(i, 3).value == None:
    #         continue
    #     print("i: ", i)
    #     for j in range(2, global_variable.maxRows):
    #         if sheet_2.cells(j, 3).value == None:
    #             continue
    #         if sheet_1.cells(i, 3).value == sheet_2.cells(j, 3).value:
    #             print(sheet_2.cells(j, 3).value)
    #             break
    #         print("j: ", j)    
    # '''