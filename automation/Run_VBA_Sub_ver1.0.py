
from win32com.client import Dispatch
import win32com.client as win32
import time




def run_excel_vba(path, filename, sheet, procedure, dt, row, col, SaveAs = False):
    '''
    :param path: 파일위치, 끝부분에 '\'입력하지 않기
    :param filename: 파일명
    :param sheet: 실행하고자하는 (sub)procedure 가 있는 시트, 혹은 매크로가 시작될 때 열려있어야 하는 시트. 시트명이어도 됨
    :param procedure: 타겟 procedure, 시트번호와 프로시저명을 넣어야함 (ex."sheet00.CommandButton100_Click")
    :param dt: 타겟 date (yyyymmdd)
    :param row: sheet에서 날짜입력 target row
    :param procedure: sheet에서 날짜입력 target col    ->결국 sheet.cells(row, col) 에 dt가 들어가는 형태
    :param SaveAs: 다른이름으로 저장하고싶은 파일 명을 넣으면 그렇게 들어감
    :return: 실행하는데 걸린 시간
    '''

    start_time = time.time()
    xl = Dispatch("Excel.Application")
    xl.Visible = True  # otherwise excel is hidden

    # 엑셀을 실행시켜서 Quantiwise Add-in 실행
    # 파일을 그냥 열면 이유는 모르겠지만 QW addin이 실행되지 않는 경우가 있기 때문
    xl.Workbooks.Add()


    wb = xl.Workbooks.Open(path + r"\\" + filename)
    ws = wb.Worksheets(sheet)
    ws.cells(row, col).value = dt
    wb.Application.Run(procedure)
    if SaveAs == False:
        wb.Save()
    else:
        wb.SaveAs(path + r"\\" + SaveAs)

    xl.Quit()

    t = float(time.time() - start_time)
    t = round(t, 2)



    return t

if __name__ == '__main__':

    '''
    ################################ 테스트용으로 만든변수 ###############################
    # path = r"C:\Projects\quant\data\지우자"
    # filename = "Price_Data_AllPairs_Holdings_longterm1.xlsm"
    # sheet = "Ctrl"
    # procedure = "sheet1.CommandButton1_Click"
    # YMD = 20180917    # YMD = time.strftime('%Y%m%d') 이렇게넣어도 되는듯?
    # row = 6
    # col = 2
    ##########################################################################################
    '''


    paths = [r'C:\Work\10 리서치\30 Quant\40 Stock Picking',
            r'C:\Work\10 리서치\30 Quant\20 Factor Model',
            r'C:\Work\20 운용\DAILY']

    filenames = ['Sentiment Tracker\Truston Sentiment Tracker_2.91_20180809.xlsm',
                'Truston Factor Model_3.08_20180731.xlsm',
                'Truston Basket Management_2.82.xlsm']

    sheets = ['Ctrl', 'Ctrl', 'Ctrl']

    procedures = ['sheet00.CommandButton100_Click',
                 'sheet01.CommandButton98_Click',
                 'sheet00.CommandButton99_Click']

    YMD = 20180918                # YMD = time.strftime('%Y%m%d') 이렇게넣어도 되는듯?
    rows = [2, 2, 22]
    cols = [14, 18, 2]

    saveas = ['trst.xlsm', False, str('Truston Basket Management_2.82_' + str(YMD) + '.xlsm')]



    for i in range(3):
        t = run_excel_vba(paths[i], filenames[i], sheets[i], procedures[i], YMD, rows[i], cols[i], SaveAs=saveas[i])

        print("--- %s minutes ---" % t)    # 소요시간

        time.sleep(60)