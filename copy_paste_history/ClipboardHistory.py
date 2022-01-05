import pyperclip
from datetime import datetime
from openpyxl import load_workbook
import time

time_delay = 0.100
old_text = ""

while True:
    if KeyboardInterrupt:
        actual_text = pyperclip.paste()
        if actual_text != old_text:
            wb = load_workbook(filename="clipboard_history.xlsx")
            ws = wb.worksheets[0]
            now = datetime.now()  # datetime object containing current date and time
            date_string = now.strftime("%d/%m/%Y")  # dd/mm/YY
            time_string = now.strftime("%H:%M")  # hh:mm
            ws.append([actual_text, date_string, time_string])
            wb.save("clipboard_history.xlsx")
            old_text = actual_text
    time.sleep(time_delay)
