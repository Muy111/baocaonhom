#Đồ án bói thần số học - Nhóm 5G

#Thành viên nhóm 
#Phạm Trà My K214060401
#Lăng Hoàng Lan K214060398
#Huỳnh Ngọc Dung K214060390
#Nguyễn Thị Hồng Ánh K214061735
#Hồ Thị Mỹ Ngân K214061741


def tong_tach(x):
    tong_so = 0
    while (x>0):
        tong_so = tong_so + (x % 10)
        x = int(x/10)
    if tong_so < 10:
     ntn.append(tong_so)
    else:
        tong_tach(tong_so)
    return tong_so

def tinh_lai(s):
    tong_sau = 0
    while (s >= 11):
        tong_sau = tong_sau + (s % 10)
        s = int(s/10)
        tong_sau = tong_sau + s
    tong.append(tong_sau)
    return tong_sau

#Hàm lấy dữ liệu trong file excel
def get_value_excel (filename, cellname):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    wb.close()
    return Sheet1[cellname].value

def vitri(f):
    vi_tri = get_value_excel("thansohoc.xlsx", f)
    return vi_tri

def space():
    print ("-"*100)
    return

class color:
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    YELLOW = '\033[93m'
    BOLD = '\033[1m'
    END = '\033[0m'
    WARNING = '\033[93m'
dem=[]
#Nhập ngày, tháng, năm sinh và kiểm tra lại dữ liệu nhập vào:
#nt: ngày tháng = {tháng:ngày}
nt = {1: 31, 2: 29, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}
ngay = int(input(">Nhập ngày sinh của bạn: "))
while ngay>31 or ngay<1:
    ngay = int(input(color.BOLD + color.WARNING + "!" + color.END + "Định dạng ngày sinh không đúng, vui lòng nhập lại: "))
thang = int(input(">>Nhập tháng sinh của bạn: "))
while thang > 12 or thang < 1:
    thang = int(input(color.BOLD + color.WARNING + "!" + color.END + "Định dạng tháng sinh sai, vui lòng nhập lại tháng sinh: ")) 
for t, n in nt.items():
    while ngay > n and thang == t:
        print(color.BOLD + color.WARNING + "!" + color.END + f'Tháng {thang} không có ngày {ngay}, vui lòng nhập lại ngày, tháng sinh: ')
        ngay= int(input("> Vui lòng nhập lại ngày sinh: "))
        while ngay>31 or ngay<1:
            ngay = int(input("Định dạng ngày sinh không đúng, vui lòng nhập lại: "))
        thang = int(input(">> Vui lòng nhập lại tháng sinh: "))
        while thang>12 or thang<1:
          print("Định dạng tháng sinh không đúng, vui lòng nhập lại!")
          ngay= int(input("> Vui lòng nhập lại ngày sinh: "))
          thang = int(input(">> Vui lòng nhập lại tháng sinh: "))
            


nam = int(input(">>>Nhập năm sinh của bạn: "))
while nam > 2021 or nam < 0:
    nam = int(input(color.BOLD + color.WARNING + "!" + color.END +"Định dạng năm sinh sai, vui lòng nhập lại năm sinh: "))

#Gọi thư viện time và openpuxl
import time
import openpyxl

ntn = []
tong = []

print(color.BOLD + color.DARKCYAN, "Hãy đợi 1 chút trong khi chúng tớ tính toán nhé :>" + color.END)
space()
time.sleep(3)

#Gọi hàm
tong_tach(ngay)
tong_tach(thang)
tong_tach(nam)
kq = 0

for i in range(len(ntn)):
    kq = ntn[i]+kq
    tong.append(kq)
if 1 < kq <= 11:
    print(color.BOLD + color.PURPLE + "Con số chủ đạo của bạn là:", kq, color.END)
else:
    kq = tinh_lai(kq)
    print(color.BOLD + color.PURPLE + "Con số chủ đạo của bạn là:", kq, color.END)
space()


print(tong)
#Nối chuỗi
cellname1 = "B" + str(kq)
cellname2 = "C" + str(kq)
cellname3 = "D" + str(kq)

time.sleep(3)
print(color.YELLOW + "Điểm nổi bật của người mang con số",kq, "là:\n" + color.END,vitri(cellname1))
space()
print(color.YELLOW + "Điểm cần khắc phục của người mang con số",kq, "là:\n" + color.END, vitri(cellname2))
space()
print(color.YELLOW + "Hướng phát triển của người mang con số", kq, "là:\n" + color.END, vitri(cellname3))
