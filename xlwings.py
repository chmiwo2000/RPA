import xlwings as xw
import pandas as pd

work_path = "C:/Users/NH-0000/Desktop/sample.xlsx" # 작업할 엑셀 파일 경로 지정

# 필요한 워크시트 적용
wb = xw.Book(work_path)
sheet = wb.sheets('Sheet1')
print(sheet.name)

# 필요한 값 추출 및 DF타입으로 변환
_li = sheet['C4:C28'].value
_df = pd.DataFrame(_li)