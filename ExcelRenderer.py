import openpyxl
import cv2
import matplotlib.pyplot as plt
import openpyxl.workbook
import win32com.client
import numpy as np
import concurrent.futures
import time
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from pathlib import Path
import os


class ExcelRenderer():

    def __init__(self, name, video):
        self.name = name
        self.video = video
              

    def load(self):
        save_path = rf"result/{self.name}"
        frames_in_chapter = 200
        self.cap = cv2.VideoCapture(self.video)
        self.wb = openpyxl.Workbook()
        total_workbooks = 0
        pixel_resolutionx = 64
        pixel_resolutiony = int(pixel_resolutionx * 16/9)

        rendered_frame_nr = 0
        i = 0
        try:
            while(self.cap.isOpened()):
                i += 1
                print(i)

                ret, frame = self.cap.read()
                if not ret:
                    print("Stream end Exiting ...")
                    break
                
                frame = cv2.resize(frame, (pixel_resolutiony, pixel_resolutionx))
                cv2.imshow("bl", frame)
                cv2.waitKey(1)
                self.create_sheet(frame, rendered_frame_nr)
                rendered_frame_nr += 1

                if rendered_frame_nr > frames_in_chapter:
                    Path(f"{save_path}").mkdir(parents=True, exist_ok=True)
                    self.wb.save(rf"{save_path}/chapter{total_workbooks}.xlsx")
                    self.wb.close()
                    self.wb = openpyxl.Workbook()

                    total_workbooks += 1
                    rendered_frame_nr = 0

        except Exception as e:
            print(e)
        finally:
            Path(f"{save_path}").mkdir(parents=True, exist_ok=True)
            self.wb.save(rf"{save_path}/chapter{total_workbooks}.xlsx")

    def set_cell_color(self, sheet, x, y, pixel):
        # This method will run in a thread, and it will set the cell color for a specific cell
        color = self.rgb_to_excel_color(pixel[0], pixel[1], pixel[2])
        print(f"setting color to {color}")
        sheet.cell(row=x+1, column=y+1).fill = color

    def create_sheet(self, frame, i):
        sheet_name = f"frame{i}"
        self.wb.create_sheet(sheet_name )
        current_sheet = self.wb[sheet_name]

        self.set_cell_color_batch(current_sheet, frame)
    
    def set_cell_color_batch(self, sheet, frame):
        """Set the color of cells in batch"""
        # Create a list of rows to apply colors in batch
        for x, column in enumerate(frame):
            row = []
            for y, pixel in enumerate(column):

                # Create the fill object
                color = self.rgb_to_excel_color(pixel[0], pixel[1], pixel[2] )
                
                row.append(color)
            # Apply the entire row of fills to the sheet in one go
            for col_num, fill in enumerate(row, 1):
                sheet.cell(row=x+1, column=col_num).fill = fill

    def rgb_to_excel_color(self, r, g, b):
        excel_color_value = np.int64(r) + (np.int64(g) * 256) + (np.int64(b) * 256 * 256)
        hex_color = f"{excel_color_value:06X}"
        color = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
        return color


    def format_workbook(self, file):
        workbook = self.load_workbook(file)
        for sheet in workbook.Sheets:
            self.excel.ActiveWindow.Zoom = 25
            sheet.Rows.RowHeight = 47
            sheet.Activate()
        workbook.Close(SaveChanges=True)
        #workbook.close(file)
    
    def play(self, workbook):
        
        for sheet in workbook.Sheets:
            time_start = time.time()
            sheet.Activate()
            time_to_sleep = 1/30 - (time.time() - time_start)
            
            if time_to_sleep > 0:
                time.sleep(time_to_sleep)
        
        workbook.Close(SaveChanges=False)
        #workbook.close(file)

    def load_workbook(self, path):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.DisplayFullScreen = True
        workbook = self.excel.Workbooks.open(path)
        return workbook  

