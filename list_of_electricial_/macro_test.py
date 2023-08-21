import xlwings as xw
import time

def run_macro():
    app = xw.App(visible=True)
    wb = app.books.open(r'C:\\Project\\py_excel\\list_of_electricial\\new_JPN_test.xlsm')  # 替换为实际的 Excel 文件路径
    # wb = app.books.open(r'C:\\Project\\py_excel\\list_of_electricial\\2.xlsm')
    try:
        # sheet = wb.sheets['Sheet2']  # 替换为实际的工作表名称
        # sheet.api.Run('RunPythonMacro')  # 运行宏
        # ins = wb.macro('Module1.test')  # 替换为实际的宏名称和模块
        # ins()
        ins = wb.macro('Sheet2.RunPythonMacro')
        ins()
        print("ok")
    except Exception as e:
        print("Error:", e)
    finally:
        wb.save()
        wb.close()
        app.quit()

if __name__ == "__main__":
    run_macro()
