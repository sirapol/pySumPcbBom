import os
import pandas as pd
from datetime import datetime

# ระบุตำแหน่งของโฟลเดอร์ที่มีไฟล์ xlsx
folder_path = "./bom"

output_file = f"{datetime.now().strftime('%Y%m%d')}_sum_bom.xlsx"
file_list_output = f"{datetime.now().strftime('%Y%m%d')}_file_list.txt"  # ระบุชื่อไฟล์ของไฟล์ text ที่จะบันทึกรายชื่อไฟล์ที่นำมาคำนวน

# output_file = "combined_bom.xlsx"
# file_list_output = "file_list.txt"  # ระบุชื่อไฟล์ของไฟล์ text ที่จะบันทึกรายชื่อไฟล์ที่นำมาคำนวน

# เก็บข้อมูล Supplier Part, Quantity, และ Supplier จากทุกไฟล์
bom_data = []

# เก็บรายชื่อไฟล์ที่ใช้ในการคำนวน
file_list = []

for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):
        file_list.append(filename)  # เพิ่มรายชื่อไฟล์
        file_path = os.path.join(folder_path, filename)
        df = pd.read_excel(file_path)
        for index, row in df.iterrows():
            supplier_part = row['Supplier Part']
            quantity = row['Quantity']
            supplier = row['Supplier']  # เพิ่มการอ่าน Supplier
            bom_data.append({'Supplier Part': supplier_part, 'Quantity': quantity, 'Supplier': supplier})

# สร้าง DataFrame จากข้อมูลที่รวมกัน
combined_df = pd.DataFrame(bom_data)

# รวมข้อมูลโดยระบุ Supplier Part, Supplier และ Quantity
combined_df = combined_df.groupby(['Supplier Part', 'Supplier']).sum().reset_index()

# บันทึก DataFrame เป็นไฟล์ Excel
combined_df.to_excel(output_file, index=False)

# เก็บรายชื่อไฟล์ที่นำมาคำนวนลงในไฟล์ text
with open(file_list_output, 'w') as file:
    file.write("\n".join(file_list))

print("สร้างไฟล์ใหม่เรียบร้อยแล้ว:", output_file)
print("บันทึกรายชื่อไฟล์ที่ใช้ในการคำนวนลงในไฟล์:", file_list_output)
