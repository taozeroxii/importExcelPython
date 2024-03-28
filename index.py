import tkinter as tk
from tkinter import filedialog
import pandas as pd
import mysql.connector
from mysql.connector import errorcode
import numpy as np  # Import numpy for handling datetime dtype

# กำหนดขนาดตัวอักษรและขนาดปุ่ม
FONT_SIZE = 10
BUTTON_WIDTH = 15
BUTTON_HEIGHT = 1

# กำหนดข้อมูลการเชื่อมต่อฐานข้อมูล MySQL
db_config = {
    'user': 'username',
    'password': 'password',
    'host': 'host',
    'database': 'db',
}

# Function เพื่อการ upload และ insertion ข้อมูล


def upload_and_insert():
    # Prompt user เลือก Excel file
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

    if file_path:
        try:
            # อ่าน Excel file เข้า DataFrame
            df = pd.read_excel(file_path)
            df.fillna(np.nan, inplace=True)
            # เชื่อมต่อฐานข้อมูล MySQL
            conn = mysql.connector.connect(**db_config)
            cursor = conn.cursor()

            # Truncate ตาราง
            truncate_sql = "TRUNCATE TABLE finance_outstanding_creditors"
            cursor.execute(truncate_sql)

           # Insert data into MySQL database
            for index, row in df.iterrows():
                # Replace NaN values with None for each row
                row = row.where(pd.notnull(row), None)

                # Insert the row into the database
                sql = """INSERT INTO finance_outstanding_creditors (accdate, type, status, plan ,pi,po,sendno,senddate,getdate,taxid,taxname,price,ba,bb,bc,bd)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                data = (
                    row['accdate'],
                    row['type'],
                    row['status'],
                    row['plan'] or '',
                    row['pi'] or '',
                    row['po'] or '',
                    row['sendno'] or '',
                    row['senddate'] or '',
                    row['getdate'] or '',
                    row['taxid'] or '',
                    row['taxname'] or '',
                    row['price'] or '',
                    row['ba'] or '',
                    row['bb'] or '',
                    row['bc'] or '',
                    row['bd'] or '',
                )
                cursor.execute(sql, data)

            # Commit การเปลี่ยนแปลง
            conn.commit()
            cursor.close()
            conn.close()
            # แสดงข้อความสถานะเมื่อการ insertion เสร็จสมบูรณ์
            status_label.config(text="เพิ่มข้อมูลเรียบร้อยแล้ว",
                                font=("Arial", FONT_SIZE), fg="red")

        except Exception as e:
            # Handle errors
            status_label.config(text=str(e), font=("Arial", FONT_SIZE))

# Function เพื่อปิดโปรแกรมเมื่อคลิกปุ่มปิดหน้าต่าง


def on_closing():
    root.destroy()


# Initialize Tkinter
root = tk.Tk()
root.title("UpLoad ข้อมูล เจ้าหนี้ 10665")
root.geometry("600x150")  # ปรับขนาดหน้าต่างให้เล็กลง

# Create a label for status
status_label = tk.Label(root, text="", font=("Arial", FONT_SIZE))
status_label.pack()

# Create a button for file upload
upload_button = tk.Button(root, text="เลือกไฟล์ Excel",
                          command=upload_and_insert, width=BUTTON_WIDTH, height=BUTTON_HEIGHT)
upload_button.pack()

# Create a label for footer
footer_label = tk.Label(
    root, text="© 2024 Cpa Information Hospital. All rights reserved.", font=("Arial", 8), fg="gray")
footer_label.pack(side="bottom", pady=5)

# Set up function to handle window closing event
root.protocol("WM_DELETE_WINDOW", on_closing)

# Run the Tkinter event loop
root.mainloop()
