import serial, csv, sys, statistics, os
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook

filename = f'motion_logs_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'

serport = serial.Serial('COM8', 115200)

with open(filename, 'a', newline='') as csvfile:
    fieldnames = ['Date and Time', 'Time', 'Event', 'Duration (s)',
                  'Total Movement Time', 'Total Idle Time', 'Average Movement Time', 'Median Movement Time', 'Mode Movement Time',
                  'Average Idle Time', 'Median Idle Time', 'Mode Idle Time','Maximum Movement Time', 'Minimum Movement Time',
                  'Maximum Idle Time', 'Minimum Idle Time']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    if csvfile.tell() == 0:
        writer.writeheader()

    total_movement_time = 0
    total_idle_time = 0
    countMovement = 0
    countIdle = 0
    average_movement_time = 0
    average_idle_time = 0
    movement_time_list = []
    idle_time_list = []

    try:
        while True:
            buffer = serport.readline().decode().strip()
            print(buffer)
            time_elapsed = buffer.split(' ')[2]
            duration = buffer.split(' ')[5]
            status = buffer.split(' ')[7]

            if status == "(Movement)":
                current_datetime = datetime.now()
                formatted_datetime = current_datetime.strftime("%m-%d-%Y %H:%M")
                total_movement_time += int(duration)
                countMovement += 1
                movement_time_list.append(int(duration))
                try:
                    average_movement_time = round(total_movement_time/countMovement, 4)
                except ZeroDivisionError:
                    average_movement_time = 0
                try:
                    median_movement_time = round(statistics.median(movement_time_list), 4)
                except statistics.StatisticsError:
                    median_movement_time = 0
                try:
                    mode_movement_time = statistics.mode(movement_time_list)
                except statistics.StatisticsError:
                    mode_movement_time = 0

                try:
                    average_idle_time = round(total_idle_time/countIdle, 4)
                except ZeroDivisionError:
                    average_idle_time = 0

                try:
                    median_idle_time = round(statistics.median(idle_time_list), 4)
                except statistics.StatisticsError:
                    median_idle_time = 0

                try:
                    mode_idle_time = statistics.mode(idle_time_list)
                except statistics.StatisticsError:
                    mode_idle_time = 0


                try:
                    max_movement_time = max(movement_time_list)
                except ValueError:
                    max_movement_time = 0

                try:
                    min_movement_time = min(movement_time_list)
                except ValueError:
                    min_movement_time = 0

                try:
                    max_idle_time = max(idle_time_list)
                except ValueError:
                    max_idle_time = 0

                try:
                    min_idle_time = min(idle_time_list)
                except ValueError:
                    min_idle_time = 0
                writer.writerow({
                    'Date and Time': formatted_datetime,
                    'Time': timedelta(seconds = int(time_elapsed)),
                    'Event': status,
                    'Duration (s)': duration,
                    'Total Movement Time': total_movement_time,
                    'Total Idle Time': total_idle_time,
                    'Average Movement Time': average_movement_time,
                    'Median Movement Time': median_movement_time,
                    'Mode Movement Time': mode_movement_time,
                    'Average Idle Time': average_idle_time,
                    'Median Idle Time': median_idle_time,
                    'Mode Idle Time': mode_idle_time,
                    'Maximum Movement Time': max_movement_time,
                    'Minimum Movement Time': min_movement_time,
                    'Maximum Idle Time': max_idle_time,
                    'Minimum Idle Time': min_idle_time
                })
            else:
                current_datetime = datetime.now()
                formatted_datetime = current_datetime.strftime("%m-%d-%Y %H:%M")
                total_idle_time += int(duration)
                countIdle += 1
                idle_time_list.append(int(duration))
                try:
                    average_movement_time = round(total_movement_time / countMovement, 4)
                except ZeroDivisionError:
                    average_movement_time = 0
                try:
                    median_movement_time = round(statistics.median(movement_time_list), 4)
                except statistics.StatisticsError:
                    median_movement_time = 0
                try:
                    mode_movement_time = statistics.mode(movement_time_list)
                except statistics.StatisticsError:
                    mode_movement_time = 0

                try:
                    average_idle_time = round(total_idle_time / countIdle, 4)
                except ZeroDivisionError:
                    average_idle_time = 0

                try:
                    median_idle_time = round(statistics.median(idle_time_list), 4)
                except statistics.StatisticsError:
                    median_idle_time = 0

                try:
                    mode_idle_time = statistics.mode(idle_time_list)
                except statistics.StatisticsError:
                    mode_idle_time = 0

                try:
                    max_movement_time = max(movement_time_list)
                except ValueError:
                    max_movement_time = 0

                try:
                    min_movement_time = min(movement_time_list)
                except ValueError:
                    min_movement_time = 0

                try:
                    max_idle_time = max(idle_time_list)
                except ValueError:
                    max_idle_time = 0

                try:
                    min_idle_time = min(idle_time_list)
                except ValueError:
                    min_idle_time = 0

                writer.writerow({
                    'Date and Time': formatted_datetime,
                    'Time': timedelta(seconds = int(time_elapsed)),
                    'Event': status,
                    'Duration (s)': duration,
                    'Total Movement Time': total_movement_time,
                    'Total Idle Time': total_idle_time,
                    'Average Movement Time': average_movement_time,
                    'Median Movement Time': median_movement_time,
                    'Mode Movement Time': mode_movement_time,
                    'Average Idle Time': average_idle_time,
                    'Median Idle Time': median_idle_time,
                    'Mode Idle Time': mode_idle_time,
                    'Maximum Movement Time': max_movement_time,
                    'Minimum Movement Time': min_movement_time,
                    'Maximum Idle Time': max_idle_time,
                    'Minimum Idle Time': min_idle_time
                })
            csvfile.flush()

    except KeyboardInterrupt:
        pass

    print(f"Program stopped. The CSV file is saved as {os.path.abspath(filename)}.")
    serport.close()

df = pd.read_csv(filename)

excel_filename = filename.replace('.csv', '.xlsx')
df.to_excel(excel_filename, index=False)

wb = load_workbook(excel_filename)
ws = wb.active

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 3)
    ws.column_dimensions[column].width = adjusted_width

wb.save(excel_filename)

print(f"Program stopped. The Excel file is saved as {os.path.abspath(excel_filename)}.")