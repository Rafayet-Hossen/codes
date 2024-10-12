import serial
import openpyxl
from datetime import datetime

excel_file = r'C:\Users\hosse\OneDrive\Desktop\Data Sheet\datasheet.xlsx'

wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

# Open the serial port (COM4)
ser = serial.Serial('COM4', 115200, timeout=1)

while True:
    # Read the data from the ESP32
    data = ser.readline().decode('utf-8').strip()

    if data:
        print(f"Received: {data}")
        
        # Check for the data format
        if "Humidity" in data and "Temperature" in data and "Moisture" in data:
            parts = data.split(',')
            humidity = parts[0].split(':')[1].strip()
            temperature = parts[1].split(':')[1].strip()
            moisture = parts[2].split(':')[1].strip()

            # Get the current timestamp
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # Append the data to the next available row in Excel
            sheet.append([timestamp, humidity, temperature, moisture])
            print(f"Saved to Excel: {timestamp}, {humidity}, {temperature}, {moisture}")

            # Save the workbook
            wb.save(excel_file)
