from openpyxl import Workbook # type: ignore
from datetime import datetime

wb = Workbook()
ws = wb.active
ws.title = "Attendance"
ws.append(["Name"])
students = ["Nishitha", "Purujith"]
for student in students:
    ws.append([student])

today = datetime.now().strftime("%d-%m-%Y")
ws.cell(row=1, column=2, value=today)

wb.save("attendance.xlsx")
print("attendance.xlsx created successfully!")