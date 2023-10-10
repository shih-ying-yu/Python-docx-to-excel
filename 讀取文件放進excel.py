from docx import Document
import openpyxl

filename = 'D:/Desktop/身分資料文件.docx'  # 改成自己電腦要讀的文件路徑
document = Document(filename)
text = []  # 讀取文件的每一行
for paragraph in document.paragraphs:
  text.append(paragraph.text)
if text[-1] == "":  # 把多讀到的空行去掉
    text = text[:-1]

updated_list = [item.rstrip(',') for item in text]  # 把最後一個逗號去掉

new_list = []


for item in updated_list:  # 分割成[['Jocky', 'Tang', 'J137849101'], ['Mandy', 'Tsai', 'R296232662'], ['Steve', 'Pan', 'Z155385612']] 的形式
    parts = item.split(', ')
    new_list.append(parts[0].split() + [parts[1]])


wb = openpyxl.Workbook()  # 建立空白的Excel
wb.save('D:/Desktop/身分資料文件.xlsx')  # 儲存檔案
wb = openpyxl.load_workbook('D:/Desktop/身分資料文件.xlsx', data_only = True)  # 開啟 Excel 檔案

loc = [['台北市', 'A'], ['臺中市', 'B'], ['基隆市',	'C'], ['臺南市', 'D'], ['高雄市', 'E'], ['新北市', 'F'], ['宜蘭縣', 'G']
, ['桃園市', 'H'], ['新竹縣', 'J'], ['苗栗縣', 'K'], ['臺中縣',	'L'], ['南投縣', 'M'], ['彰化縣', 'N'], ['雲林縣', 'P'], ['嘉義縣',	'Q']
, ['臺南縣', 'R'], ['高雄縣', 'S'], ['屏東縣', 'T'], ['花蓮縣',	'U'], ['臺東縣', 'V'], ['澎湖縣', 'X'], ['陽明山', 'Y'], ['金門縣',	'W']
, ['連江縣', 'Z'], ['嘉義市', 'I'], ['新竹市', 'O']]
# 身分證開頭對照縣市 


s1 = wb['Sheet']  # 開啟Sheet(如果電腦的sheet名稱是叫工作表1可能要改)

ans = []  # 要輸入進excel的資料
for y in range(len(new_list)):
    ans.append([])

for u in range(len(new_list)):
    ans[u].append(new_list[u][1])  # 姓
    ans[u].append(new_list[u][0])  # 名
    if new_list[u][2][1] == '1':
        ans[u].append('Male')  # 是男的
    if new_list[u][2][1] == '2':
        ans[u].append('Female')  # 是女的
    ans[u].append(new_list[u][2])  # 身分證
    for x in range(len(loc)):  # 縣市
        if new_list[u][2][0] == loc[x][1]:
            ans[u].append(loc[x][0])


for item in ans:  # key進excel
    s1.append(item)

row_index_to_insert = 1  # 插入一行到第1行之前（A 行）
s1.insert_rows(row_index_to_insert)

s1['A1'].value = '姓'         # 儲存格 A1 內容為 姓
s1['B1'].value = '名'         # 儲存格 B1 內容為 名
s1['C1'].value = '性別'       # 儲存格 C1 內容為 性別
s1['D1'].value = '身分證'     # 儲存格 D1 內容為 身分證
s1['E1'].value = '戶籍地'     # 儲存格 E1 內容為 戶籍地

wb.save('D:/Desktop/身分資料文件.xlsx')