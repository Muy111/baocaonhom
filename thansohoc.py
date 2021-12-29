'''Đồ án bói thần số học - Nhóm 5G
Thành viên nhóm
K214060401 - Phạm Trà My
K214060398 - Lăng Hoàng Lan
K214060390 - Huỳnh Ngọc Dung
K214061735 - Nguyễn Thị Hồng Ánh
K214061741 - Hồ Thị Mỹ Ngân'''

def tong_tach(x):
    tong_so = 0
    while (x>0):
        tong_so = tong_so + (x%10)
        x = int (x/10)
    if tong_so < 10:
     ntn.append(tong_so)
    else:
        tong_tach(tong_so)
    return tong_so

def tinh_lai(s):
    tong_sau = 0
    while (s >= 11):
        tong_sau = tong_sau + (s%10)
        s = int (s/10)
        tong_sau = tong_sau + s
    tong.append(tong_sau)
    return tong_sau

#Hàm lấy dữ liệu trong file excel
def get_value_excel (filename, cellname):
	wb = openpyxl.load_workbook(filename)
	Sheet1 = wb ['Sheet1']
	wb.close()
	return Sheet1[cellname].value

def vitri(f):
    vi_tri = get_value_excel ("thansohoc.xlsx", f)
    return vi_tri
def space():
    print ("------------------------------------------------------------------------")
    return

class color:
    PURPLE = '\033[95m'
    DARKCYAN = '\033[36m'
    YELLOW = '\033[93m'
    BOLD = '\033[1m'
    END = '\033[0m'

#Nhập ngày, tháng, năm sinh và kiểm tra lại dữ liệu nhập vào
ngay = int(input(">Nhập ngày sinh của bạn: "))
while ngay>31 or ngay<1:
    ngay = int(input("Định dạng ngày sinh không đúng, vui lòng nhập lại: "))
thang = int(input(">>Nhập tháng sinh của bạn: "))
while thang>12 or thang<1:
    thang = int(input("Định dạng tháng sinh không đúng, vui lòng nhập lại: "))
while ((thang==4 or thang==6 or thang==9 or thang==11) and ngay>30) or (thang==2 and ngay>29):
    print("Tháng "+str(thang)+" không có ngày " + str(ngay) + " vui lòng nhập lại!")
    ngay= int(input("> Vui lòng nhập lại ngày sinh: "))
    while ngay>31 or ngay<1:
      ngay = int(input("Định dạng ngày sinh không đúng, vui lòng nhập lại: "))
    thang = int(input(">> Vui lòng nhập lại tháng sinh: "))
    while thang>12 or thang<1:
       print("Định dạng tháng sinh không đúng, vui lòng nhập lại!")
       ngay= int(input("> Vui lòng nhập lại ngày sinh: "))
       thang = int(input(">> Vui lòng nhập lại tháng sinh: "))

nam = int(input(">>>Nhập năm sinh của bạn: "))
while nam <0:
    nam = int(input("Định dạng năm sinh không đúng, vui lòng nhập lại: "))

#Gọi thư viện time và openpyxl
import time
import openpyxl

ntn = []
tong = []

print (color.BOLD + color.DARKCYAN,"Hãy đợi 1 chút trong khi chúng tớ tính toán nhé :>"+ color.END)
space()
time.sleep(3)

#Gọi hàm
tong_tach(ngay)
tong_tach(nam)
tong_tach(thang)
kq = 0

for i in range (len(ntn)):
    kq = ntn[i]+kq
    tong.append(kq)
if 1 < kq <= 11:
  print(color.BOLD + color.PURPLE + "Con số chủ đạo của bạn là:", kq)
else:
  kq = tinh_lai(kq)
  print(color.BOLD + color.PURPLE + "Con số chủ đạo của bạn là:", kq)
print(color.END)
space()
time.sleep(3)

#Nối chuỗi
cellname1 = "B" + str(kq)
cellname2 = "C" + str(kq)
cellname3 = "D" + str(kq)


print(color.YELLOW + "Điểm nổi bật của người mang con số",kq, "là:\n" + color.END,vitri(cellname1))
space()
print(color.YELLOW + "Điểm cần khắc phục của người mang con số",kq, "là:\n" + color.END, vitri(cellname2))
space()
print(color.YELLOW + "Hướng phát triển của người mang con số", kq, "là\n" + color.END, vitri(cellname3))
