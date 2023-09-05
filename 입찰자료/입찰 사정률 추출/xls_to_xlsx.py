import os
import win32com.client as win32
import shutil

dir = r'C:\Users\user\Downloads'
os.chdir(dir)

for xls_file in os.listdir():
  if r'KBID_결과_2023' in xls_file and r'.xlsx' not in xls_file:
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel_app.Workbooks.Open(r'{}'.format(dir + '\\' + xls_file))
    print(dir + '\\' + xls_file)
    wb.SaveAs(xls_file + 'x', FileFormat = 51)
    wb.Close()
    excel_app.Application.Quit()

print("파일은 C:\\Users\\user\\Desktop\\입찰자료\\입찰통계 에 저장됩니다.")

os.chdir(r"C:\Users\user\Documents")
for i in os.listdir():
  if i.endswith(".xlsx"):
    print(i)
    try:
      shutil.move(r"C:\Users\user\Documents"+"\\"+i, r'C:\Users\user\Desktop\입찰자료\입찰통계')
    except:
      continue
print("다운로드 폴더에 있는 파일을 삭제하겠습니다.")

os.chdir(dir)
for i in os.listdir():
  if i.endswith(".xls"):
    os.remove(str(i))

print("삭제가 완료 되었습니다.")