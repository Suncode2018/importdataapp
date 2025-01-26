# โปรเจกต์จัดการไฟล์และฐานข้อมูล 🎉

> โปรเจกต์นี้เป็นระบบจัดการไฟล์และฐานข้อมูล MSSQL โดยใช้ Python และ Tkinter สำหรับสร้าง GUI เพื่ออำนวยความสะดวกในการนำเข้าไฟล์, จัดการข้อมูลในฐานข้อมูล, และตรวจสอบสิทธิ์ผู้ใช้งาน

## 🚀 Features
- **ระบบล็อกอิน**: ตรวจสอบสิทธิ์ผู้ใช้งานผ่านฐานข้อมูลด้วย stored procedure
- **นำเข้าไฟล์**: นำเข้าไฟล์ (เช่น XML, XLSX, CSV) และบันทึกข้อมูลลงในฐานข้อมูล
- **จัดการข้อมูล**: แสดงข้อมูลจากฐานข้อมูลในรูปแบบตาราง และรองรับการรีเซ็ตข้อมูล
- **ตรวจสอบไฟล์ซ้ำ**: ตรวจสอบว่าไฟล์ที่นำเข้าไม่ซ้ำกับข้อมูลที่มีอยู่
- **รองรับหลายไฟล์**: สามารถนำเข้าและประมวลผลไฟล์ XML, XLSX, และ CSV ได้
- **แสดงความคืบหน้า**: มีแถบแสดงความคืบหน้าในการประมวลผลไฟล์

## 🛠️ Installation
### ขั้นตอนการติดตั้ง
ติดตั้ง python 3.8 ขึ้นไป
pip install -r requirements.txt

### สร้างไฟล์ .env:
- DB_SERVER='''your_database_server'''
- DB_DATABASE='''your_database_name'''
- DB_USERNAME='''your_database_username'''
- DB_PASSWORD='''your_database_password'''
- DB_DRIVER='''your_database_driver'''
- DB_WH='''your_database_warehouse'''

1. **โคลนโปรเจกต์จาก GitHub**:
   ```bash
   git clone https://github.com/Suncode2018/importdataapp.git



คุณสามารถปรับเปลี่ยนเนื้อหาให้เหมาะสมกับโปรเจกต์ของคุณได้ตามต้องการ!
### create .exe for windows
- pip install pyinstaller
- pyinstaller login.py --noconsole --onefile