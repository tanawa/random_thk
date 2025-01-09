import openpyxl
import pyrebase
import time
config = {
  "apiKey": "AIzaSyBGkQh9ygRa6-iEY2V4mWQ8YGemFZA4-HI",
  "authDomain": "thkstorage-ec701.firebaseapp.com",
  "databaseURL": "https://thkstorage-ec701-default-rtdb.asia-southeast1.firebasedatabase.app",
  "storageBucket": "thkstorage-ec701.firebasestorage.app"
}

firebase = pyrebase.initialize_app(config)
db = firebase.database()
wb = openpyxl.load_workbook('reward.xlsx') 
sheet = wb.active 
# เข้าถึงค่าเซลล์
n=0
for i in sheet.rows:
    n+=1
    cell_value = sheet.cell(n,2).value
    if "ชื่อ - นามสกุล" in cell_value:
        continue
    data = {}
    data["data"] = cell_value
    db.child("student_used_check_2025/"+(str(int(time.time())))).set(data)
    print(cell_value) 
    time.sleep(1)
    
