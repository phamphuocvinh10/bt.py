import mysql.connector
import StudentService

stuserv = StudentService
check = -1
stuserv.sqlconn()
while check != 0:
    print("1. Lấy dữ liệu")
    print("2. Xem danh sách sinh viên")
    print("3. Xem danh sách sinh viên trên CSDL")
    print("4. In dữ liệu vào Mysql")
    print("5. Sửa dữ liệu trong CSDL")
    print("6. Xóa dữ liệu trong CSDL")
    check = int(input("Nhập số: "))
    while check < 0 or check > 6:
        check = int(input("Nhập lại số: "))

    print("*" * 50)
    if check == 1:
        select = -1
        while select != 0:
            print("1. Đọc file excel")
            print("2. Đọc csdl MySQL")
            select = int(input("Nhập số: "))
            while select < 0 or select > 2:
                select = int(input("Nhập lại số: "))
            print("*" * 50)
            if select == 1:
                stuserv.clearlist()
                stuserv.readexcel()
                break
            elif select == 2:
                stuserv.clearlist()
                stuserv.readmysql()
                break
    elif check == 2:
        stuserv.showsv(stuserv.getlistSV())
    elif check == 3:
        stuserv.showsvinsql()
    elif check == 4:
        stuserv.insertall(stuserv.getlistSV())
    elif check == 5:
        inp = str(input("Nhap MaSV hoc sinh can sua: "))
        id1 = stuserv.search(inp)
        stuserv.changedata(stuserv.getlistSV(), id1)
        stuserv.update_student(stuserv.getlistSV(), id1)
    elif check == 6:
        inp = str(input("Nhap MaSV hoc sinh can xoa: "))
        id1 = stuserv.search(inp)
        stuserv.deletedata(id1)
    else:
        break
    print("*" * 50)
