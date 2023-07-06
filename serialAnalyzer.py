import serial
import serial.tools.list_ports
import openpyxl
from datetime import datetime

ports = list(serial.tools.list_ports.comports())
portList = []
for port in ports:
    portList.append(str(port.device))
    print(str(port.device))

selectedPort = input("Please Select Available Port COM: ")
fileName = input("Please Enter File Name: ")

ser = serial.Serial(selectedPort, baudrate=115200, timeout=1)

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(["Time", "Data"])

while True:
    try:
        if ser.in_waiting > 0:
                line = ser.readline().decode('utf-8').rstrip('\r\n')
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"{timestamp}: {line}")  # Print the received data with timestamp
                sheet.append([timestamp, line])  # Write the timestamp and data to the Excel sheet

    except KeyboardInterrupt:
        print("Handling interrupt...")
        print("Saving File...")
        workbook.close()
        break

workbook.save(f"{fileName}.xlsx")  # Save the Excel file
print("Recording stopped")
print("Data saved to" + fileName + ".xlsx")
print("Process Done")
ser.close()  # Close the serial port
