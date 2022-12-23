import openpyxl

wb=openpyxl.load_workbook("C:\codings\python\excel\data.xlsx")

sh1=wb[wb.sheetnames[0]]
sh1.insert_rows(1)
for i in range(0,5):
    day=""
    if i==0:
        day="Monday"
    elif i==1:
        day="Tuesday"
    elif i==2:
        day="Wednesday"
    elif i==3:
        day="Thursday"
    else:
        day="Friday"
    print("Today is ", day)
    print("enter course name, if not, hit enter ")
    p1=input("enter class at 8-9 ")
    p2=input("enter class at 9-10 ")
    p3=input("enter class at 10-11 ")
    p4=input("enter class at 11-12 ")
    p5=input("enter class at 1-2 ")
    p6=input("enter class at 2-3 ")
    p7=input("enter class at 3-4 ")
    p8=input("enter class at 4-5 ")
    p9=input("enter class at 5-6 ")
    
    sh1.append([day,p1,p2,p3,p4,p5,p6,p7,p8,p9])
   

wb.save('C:\codings\python\excel\data.xlsx')
