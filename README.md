import pandas as pd
import os

file_path = 'data.xlsx'
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    print("ใช้ไฟล์เดิม")
else:
    df = pd.DataFrame(columns=['ชื่อ', 'เบอร์', 'อายุ', 'เพศ', 'ความพึงพอใจ'])
    print("สร้างไฟล์")

while True:
    print("กรอกข้อมูลต่อไปนี้หรือพิมพ์ 'stop' เพื่อหยุดการป้อนข้อมูล")
    name = input("กรุณาใส่ชื่อ: ")
    if name.lower() == 'stop':
        break
    phone = input("กรุณาใส่เบอร์โทรศัพท์ (10 หลัก): ")
    if len(phone) != 10:
        print("เบอร์โทรศัพท์ต้องมี 10 หลัก.")
        continue
    age = int(input("กรุณาใส่อายุ (ห้ามเกิน 80): "))
    if age > 80:
        print("อายุไม่ถูกต้อง, กรุณาใส่อายุไม่เกิน 80.")
        continue
    gender = input("กรุณาใส่เพศ (ชาย/หญิง): ")
    if gender not in ['ชาย', 'หญิง']:
        print("เพศไม่ถูกต้อง, กรุณาใส่ว่า 'ชาย' หรือ 'หญิง'.")
        continue
    satisfaction = int(input("กรุณาใส่ความพึงพอใจ (คะแนน 1-5): "))
    if satisfaction < 1 or satisfaction > 5:
        print("คะแนนความพึงพอใจไม่ถูกต้อง, กรุณาใส่คะแนนระหว่าง 1-5.")
        continue
    new_row = pd.DataFrame({
        'ชื่อ': [name],
        'เบอร์': [phone],
        'อายุ': [age],
        'เพศ': [gender],
        'ความพึงพอใจ': [satisfaction]
    })
    df = pd.concat([df, new_row], ignore_index=True)

df.to_excel(file_path, index=False)
print(f'ไฟล์ {file_path} ได้ถูกบันทึกแล้ว.')
