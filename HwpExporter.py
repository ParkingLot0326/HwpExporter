import os
import time
import re
import json
import threading
import logging
import traceback

import xml.etree.ElementTree as ET
import win32com.client as win32
from pyhwpx import Hwp
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Constants
VERSION = "1.0.0"
IS_STABLE = True

DATA_DIR = "data"
SETTINGS_FILE = os.path.join(DATA_DIR,'settings.json')
LOG_FILE = os.path.join(DATA_DIR,'hwp_converter.log')
# CHANGELOG V1.0.0 : 딜레이와 재시도 횟수 설정 제거
DEFAULT_SETTINGS = {
    "isHwpVisible": True,
    "isExcelVisible":True,
    "doOpenHwp": True,
    "doOpenXlsx": True,
    "SPMode" : True,
    # "copyPasteDelay" : 0.2,
    # "retryLife" : 5,
}

class HwpConverter:
    def __init__(self):
        self.file = ''
        self.filename = ''
        self.ctrl = None
        self.current_page = 1
        self.export_path = '.'
        self.settings = self.load_settings()
        self.hwp = None
        self.excel = None
        self.wb = None
        self.ws = None
        self.row_index = 1
        self.cancel_extraction = False
        self.setup_logging()

    def ensure_data_dir(self):
        if not os.path.exists(DATA_DIR):
            os.makedirs(DATA_DIR)
    
    def setup_logging(self):
        logging.basicConfig(filename=LOG_FILE,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s',encoding="utf-8")
        logging.info("")
        logging.info("")
        logging.info(f"Program Started; Current Version is: {VERSION}")

    def reset_state(self):
        self.current_page = 1
        self.ctrl = None
        self.exported_pages = 0
        self.total_pages = 0
        logging.info("State reset")
    
    def load_settings(self):
        if not os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'w', encoding="utf-8") as file:
                json.dump(DEFAULT_SETTINGS, file, ensure_ascii=False, indent="\t")
            return DEFAULT_SETTINGS
        with open(SETTINGS_FILE, 'r', encoding="utf-8") as file:
            return json.load(file)

    def save_settings(self):
        with open(SETTINGS_FILE, 'w', encoding="utf-8") as file:
            json.dump(self.settings, file, ensure_ascii=False, indent="\t")

    def open_hwp_file(self):
        try:
            if self.file:
                self.hwp = Hwp(visible=not self.settings['isHwpVisible'], new=False, register_module="./FilePathCheckeModule.dll")
                self.hwp.open(self.file)
                logging.info("HWP file Opened Successfully.")
            else:
                logging.warning("Hwp File Not Selected on GUI")
                raise ValueError("No file selected")
        except Exception as e:
            logging.error(e)
            raise

    def close_hwp_file(self):
        if hasattr(self, 'hwp') and self.hwp:
            try:
                self.hwp.Clear(2)  
                self.hwp.Quit()
            except:
                pass
        self.hwp = None
        self.ctrl = None
        logging.info("hwp closed.")

    def open_excel_file(self):
        logging.info("open_excel_file executed;")
        try:
            self.close_excel_file()
            save_file = (self.export_path + "/" + self.filename).replace("/","\\")
            save_file = self.get_unique_filename(filename=save_file)
            
            self.excel = win32.gencache.EnsureDispatch("Excel.Application")
            self.wb = self.excel.Workbooks.Add()
            self.wb.SaveAs(save_file)
            self.excel.Quit()
            self.wb = self.excel.Workbooks.Open(save_file)
            self.excel.Visible = not self.settings["isExcelVisible"]
            logging.info("Excel opened.")
        except Exception as e:
            logging.error(f"Creating New Excel file failed : {e}")
            raise Exception("Opening Excel failed.")
    
    def close_excel_file(self):
        logging.info("closing excel..")
        try:
            if hasattr(self, 'wb') and self.wb:
                try:
                    self.wb.Close(SaveChanges=True)
                except Exception as e:
                    logging.error(f"exception ocurred in closing excel #1 : {e}")
                    pass  # If there's an error closing the workbook, continue to close Excel
            if hasattr(self, 'excel') and self.excel:
                try:
                    self.excel.Quit()
                except Exception as e:
                    logging.error(f"exception ocurred in closing excel #2 : {e}")
                    pass  # If there's an error quitting Excel, we've done our best
        except Exception as e:
            logging.error("Unknown error")
        self.wb = None
        self.excel = None
        logging.info("excel closed.")

    def get_unique_filename(self, filename):
        temp = filename.split(".")
        counter = 1
        while os.path.exists(filename):
            filename = temp[0] + "(%i)" %counter + "." + temp[1]
            counter += 1
        return filename

    def go_to_start_page(self, initial_page):
        try:
            while self.current_page != initial_page and not self.cancel_extraction:
                if self.current_page < initial_page:
                    self.ctrl = self.ctrl.Next
                    if self.ctrl is None:
                        break
                    if self.ctrl.UserDesc == "표":
                        self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
                        self.current_page = self.hwp.current_page
                else:
                    self.ctrl = self.ctrl.Prev
                    self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
                    self.temp_cur_page = self.hwp.current_page
                    # CHANGELOG V1.0.0: 이전 컨트롤로 돌아가는 동작을 한번만 하도록 변경.
                    # self.hwp.SetPosBySet(self.ctrl.Prev.GetAnchorPos(0))
                    # self.temp_prev_page = self.hwp.current_page
                    if self.temp_cur_page != self.temp_prev_page :
                        self.current_page = self.temp_cur_page
                        
                # logging.info(f"Current Page is: {self.current_page}")
                self.hwp.goto_page(self.current_page)

            if self.cancel_extraction:
                logging.info("Extraction cancelled during go_to_start_page")
                return

            logging.info(f"Moved to Start Page {self.current_page}")

        except Exception as e:
            logging.error(f"Failed while Moving to Start Page, current Page : {self.current_page}, E: {e}")
            raise Exception("Moving to Start page Failed.")
    
    # CHANGELOG V1.0.0: 추출 로직 변경
    # @deprecated
    # def copy_paste_action(self):
    #     try:
    #         # self.hwp.MovePageBegin()
    #         self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
    #         self.current_page = self.hwp.current_page
    #         self.hwp.FindCtrl()
    #         time.sleep(self.settings["copyPasteDelay"]/2)
    #         self.hwp.Copy()
    #         time.sleep(self.settings["copyPasteDelay"]/2)
            
    #         self.ws.Activate()
    #         if self.excel.ClipboardFormats:
    #             self.excel.ActiveSheet.Paste()
    #             logging.info(f"Paste Successful. current Page: {self.current_page}")
    #         else:
    #             logging.warning(f"Paste failed - nothing on Clipboard. current Page : {self.current_page}")
    #             raise Exception("Paste failed: Nothing on Clipboard.")
    #     except Exception as e :
    #         logging.error(f"Error in copy_paste: {e}")
    #         raise
    
    
    # CHANGELOG V1.0.0: 공백 관리가 간편해짐에 따라 통합 밑 제거
    # @deprecated
    # def get_row_and_offset(self):
    #     if self.cancel_extraction:
    #         logging.info("Extraction cancelled before excel offset")
    #         return
    #     try:
    #         logging.info("get table row and offset")
    #         self.hwp.FindCtrl()
    #         self.hwp.HAction.Run("ShapeObjTableSelCell")
    #         row_num = self.hwp.get_row_num()
    #         self.hwp.HAction.Run("Cancel")
    #         self.row_index += row_num + 8
    #         self.ws.Range(f"A{self.row_index}").Select()
    #     except Exception as e:
    #         logging.error(f"Row Calculation failed: {e}")
    #         logging.warning(traceback.format_exc())
    #         raise Exception("Error calculating rows")
        
    # CHANGELOG V1.0.0: XML을 활용한 신규 추출 방식 도입
    # 
    # 한글 문서에서 XML 형식의 텍스트를 추출하고, TABLE 태그 내부의 엘리먼트만 남긴 후 파싱하여 엑셀로 옮긴다.
    def export_via_xml(self):
        def optimize_column_height(sheet, start_row, end_row):
            for row in range(start_row, end_row + 1):
                if sheet.Rows(row).RowHeight > 24:
                    sheet.Rows(row).RowHeight = 24
        
        self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
        self.current_page = self.hwp.current_page
        self.hwp.FindCtrl()
        
        table_pattern = r"<TABLE\b.*?>.*?</TABLE>"
        footnote_pattern = r"<FOOTNOTE\b.*?>.*?</FOOTNOTE>"
        
        src = self.hwp.GetTextFile("HWPML2X",option="saveblock")
        logging.info("got src for this page.")
        
        text_without_footnote = re.sub(footnote_pattern, '', src, flags=re.DOTALL | re.IGNORECASE)
        tableList = re.findall(table_pattern, text_without_footnote, flags=re.DOTALL | re.IGNORECASE)
        src = '\n'.join(tableList)
        
        root = ET.fromstring(src)
        logging.info("generated root of xml data")
        
        self.ws.Activate()
        sheet = self.excel.ActiveSheet
        
        sheet.Cells.Font.Size = 9
        
        start_row = self.row_index
        
        for row_elem in root.findall(".//ROW"):
            cells_to_process = []
            max_row_span = 1
            
            for cell_elem in row_elem.findall("CELL"):
                col_addr = int(cell_elem.get("ColAddr", 0))
                col_span = int(cell_elem.get("ColSpan", 1))
                row_span = int(cell_elem.get("RowSpan", 1))
                text_content = [text for text in ("".join(p_elem.itertext()).strip() for p_elem in cell_elem.findall(".//P")) if text]
                if text_content:
                    cells_to_process.append((col_addr, col_span, text_content))
                    max_row_span = max(max_row_span,row_span)
            
            if cells_to_process:
                max_cell_height = max(len(content) for _, _, content in cells_to_process)
                
                for col_addr, col_span, text_content in cells_to_process:
                    for i, text in enumerate(text_content):
                        cell = sheet.Cells(self.row_index + i, col_addr + 1)
                        cell.Value = text
                        
                    if col_span > 1:
                        for i in range(max(len(text_content), 1)):
                            start_cell = sheet.Cells(self.row_index + i, col_addr + 1)
                            end_cell = sheet.Cells(self.row_index + i, col_addr + col_span)
                            sheet.Range(start_cell, end_cell).Merge()
                
                # 세로 병합될 셀이 있으면 줄바꿈 효과를 무력화시켜서 바로 밑 열부터 채우도록 바꿈
                if max_row_span > 1 :
                    max_cell_height = 1
                    
                self.row_index += max_cell_height
                
        optimize_column_height(sheet, start_row, self.row_index - 1)
        # sheet.Rows(f"{self.row_index}:{self.row_index + max_cell_height - 1}").AutoFit()
        # logging.info("AutoFitted")
        
    # CHANGELOG V1.0.0: 추출 방식 변경
    def copy_paste_to_endpage(self, end_page, update_progress_callback):
        if self.cancel_extraction:
            logging.info("Extraction cancelled before copy-paste")
            return
        
        if self.current_page > end_page:
            logging.info("current page is bigger than end_page; ending copy-paste")
            return
        
        while end_page >= self.current_page:
            logging.info("Copy-Paste Started")

            if self.cancel_extraction:
                logging.info("Extraction cancelled during pre-copy")
                break
                   
            if self.ctrl is None:
                logging.info("Extraction Ended Reaching the End of Document")
                break
            
            if self.ctrl.CtrlID == "tbl":
                if self.cancel_extraction:
                    logging.info("Extraction cancelled during copy_paste_to_endpage")
                    break
                
                try:
                    self.export_via_xml()
                    self.row_index += 2
                    # self.copy_paste_action()

                    self.exported_pages += 1
                    progress = (self.exported_pages / self.total_pages) * 100
                    update_progress_callback(progress=progress, status=f"Exporting page {self.current_page}...")
                
                except Exception as e:
                    logging.error(f"Copy-Paste failed: {e}")
                    time.sleep(0.3)
                    raise Exception(f"Failed to copy-paste: {e}")
            
                if self.cancel_extraction:
                    logging.info("Extraction cancelled before excel offset")
                    return
            
            self.ctrl = self.ctrl.Next  
            if self.ctrl is None:
                logging.info("Extraction Ended Reaching the End of Document")
                break
            
            self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
            self.current_page = self.hwp.current_page

            logging.info(f"moved to next page: {self.current_page}")
            logging.info("done.")
            if self.cancel_extraction:
                logging.info("Extraction Cancelled after processing a table")
                return
        logging.info("while loop in copy_paste_to_endpage ended")
    
    def is_number(self,cell_value):
        try:
            float(cell_value)
            return True
        except:
            return False        
        
    # CHANGELOG V1.0.0: 공백 조절 기능 제거, 테두리 일괄적용
    def rearrange_demos(self):
        try:
            for sheet in self.wb.Worksheets:
                used_range = sheet.UsedRange
                total_used_row = used_range.Rows.Count
                
                if sheet == self.wb.Worksheets("Sheet1") and self.settings["SPMode"]:
                    logging.info("Spliting first sheet")
                    self.split_first_sheet(sheet)
                    
                else:
                    row = 1
                    while row <= total_used_row:
                        if self.cancel_extraction:
                            logging.warning("extraction canceled while rearranging excel")
                            return
                        
                        table = sheet.Cells(row, 1).CurrentRegion
                        for i in range(7, 11):  # Excel에서 7-12는 테두리 상단, 하단, 좌측, 우측, 대각선 등
                            border = table.Borders(i)
                            border.LineStyle = 1  # 실선
                            border.Weight = 2     # 두께: 2는 중간 굵기, 4는 두꺼운 테두리
                            border.Color = 0x000000  # 검정색
                        table_row = table.Rows.Count
                        table_column = table.Columns.Count

                        for cell in table:
                            if cell.MergeCells:
                                cell.UnMerge()

                        right_column = sheet.Range(
                            sheet.Cells(row, table_column),
                            sheet.Cells(row + table_row - 1, table_column),
                        )
                        
                        if all(
                            cell.Value is None or (not self.is_number(str(cell.Value)) and str(cell.Value) != "-")
                            for cell in right_column
                        ):
                            demo_value = right_column.Value
                            dest = sheet.Range(
                                table.Offset(1, 2),
                                table.Offset(table_row, table_column + 1),
                            )
                            dest.Value = table.Value
                            new_demo = sheet.Range(
                                table.Offset(1, 1), table.Offset(table_row, 1)
                            )
                            new_demo.Value = demo_value
                            sheet.Range(
                                table.Offset(1, table_column + 1),
                                table.Offset(table_row, table_column + 1),
                            ).Value = ""

                        else:
                            pass

                        row += table_row +2
                        # logging.info(f"row is now:{row}")  
                            
        except Exception as e:
            logging.error("Error while arranging Excel")
            logging.warning(traceback.format_exc())
            raise Exception(f"Re-arranging Excel failed : {e}")
        
    def split_first_sheet(self, sheet):
        sheet.Copy(Before=sheet)
        sheet1 = sheet
        sheet2 = self.wb.Worksheets(f"{sheet.Name} (2)")
        sheet1.Select()

        for i in range(2):
            if i == 0:
                sheet = sheet1
            else:
                sheet = sheet2

            # sheet.Select()
            used_range = sheet.UsedRange
            total_used_row = used_range.Rows.Count
            total_used_column = used_range.Columns.Count

            row = 1
            while row <= total_used_row:
                if self.cancel_extraction:
                    logging.warning("extraction canceled while rearranging excel")
                    return
                
                if sheet.Cells(row, 1).Value is None:
                    
                    if all(
                        sheet.Cells(row, col).Value is None
                        for col in range(1, total_used_column + 1)
                    ):
                        next_non_empty_row = row + 1

                        while next_non_empty_row <= total_used_row and all(
                            sheet.Cells(next_non_empty_row, col).Value is None
                            for col in range(1, total_used_column + 1)
                        ):
                            next_non_empty_row += 1

                        sheet.Rows(f"{row}:{next_non_empty_row - 1}").Delete(
                            Shift=win32.constants.xlUp
                        )

                        sheet.Rows(row).Insert(Shift=win32.constants.xlDown)
                        sheet.Rows(row).Insert(Shift=win32.constants.xlDown)

                        total_used_row -= next_non_empty_row - row
                        total_used_row += 2
                        row += 2
                    else:
                        row += 1
                    
                    if all(
                        all(
                            sheet.Cells(check_row, col).Value is None
                            for col in range(1, total_used_column + 1)
                        )
                        for check_row in range(
                            row, min(row + 5, total_used_row + 1)
                        )
                    ) :
                        break
                else:
                    table = sheet.Cells(row, 1).CurrentRegion
                    table_row = table.Rows.Count
                    table_column = table.Columns.Count

                
                table = sheet.Cells(row, 1).CurrentRegion
                for j in range(7, 11):  # Excel에서 7-12는 테두리 상단, 하단, 좌측, 우측, 대각선 등
                    border = table.Borders(j)
                    border.LineStyle = 1  # 실선
                    border.Weight = 2     # 두께: 2는 중간 굵기, 4는 두꺼운 테두리
                    border.Color = 0x000000  # 검정색
                table_row = table.Rows.Count
                table_column = table.Columns.Count

                for cell in table:
                    if cell.MergeCells:
                        cell.UnMerge()

                right_column = sheet.Range(
                    sheet.Cells(row, table_column),
                    sheet.Cells(row + table_row - 1, table_column),
                )

                if all(
                    cell.Value is None or (not self.is_number(str(cell.Value)) and str(cell.Value) != "-")
                    for cell in right_column
                ):  # if demo is on right end
                    if i == 0: # in sheet1
                        sheet.Rows(f"{row}:{row + table_row}").Delete(
                            Shift=win32.constants.xlDown
                        )
                        if row != 1:
                            row -= 2
                    else:
                        row += table_row
                        
                        demo_value = right_column.Value
                        dest = sheet.Range(
                            table.Offset(1, 2),
                            table.Offset(table_row, table_column + 1),
                        )
                        dest.Value = table.Value
                        new_demo = sheet.Range(
                            table.Offset(1, 1), table.Offset(table_row, 1)
                        )
                        new_demo.Value = demo_value
                        sheet.Range(
                            table.Offset(1, table_column + 1),
                            table.Offset(table_row, table_column + 1),
                        ).Value = ""
                        
                else:
                    if i == 1: # in sheet1-1
                        sheet.Rows(f"{row}:{row + table_row}").Delete(
                            Shift=win32.constants.xlDown
                        )
                        if row != 1:
                            row -= 2
                    else:
                        row += table_row +2
    
    # CHANGELOG V1.0.0: 추출 방식 안정화로 재시도 로직 제거
    # @deprecated
    # def resume_extraction(self, range_list, update_progress_callback):
    #     isRetryEnd = True
    #     for trial in range(1,self.settings["retryLife"]+1):
    #         if self.wb:
    #             self.wb.Save()
                
    #         if self.cancel_extraction:
    #             logging.warning("Extraction cancelled during resume_extraction")
    #             return

    #         if range_list[-1] > self.current_page :
    #             update_progress_callback(status=f"Re-Trying extraction from page {self.current_page}...")
    #             logging.info(f"trial {trial} has begun.")
    #             logging.info(f"Retrying extraction from page {self.current_page}")

    #             self.close_hwp_file()
    #             self.open_hwp_file()
    #             time.sleep(0.5)
    #             self.ctrl = self.hwp.HeadCtrl

    #             current_range_index = next((i for i in range(0, len(range_list), 2) if range_list[i] <= self.current_page <= range_list[i+1]), None)

    #             adjusted_range_list = range_list[current_range_index:]
    #             adjusted_range_list[0] = self.current_page
    #             self.current_page = 1
    #             logging.warning(f"adjusted range list : {adjusted_range_list}")
                
    #             self.row_index += 40
    #             self.ws.Range(f"A{self.row_index}").Select()
    #             for i in range(0, len(adjusted_range_list), 2):
    #                 if self.cancel_extraction:
    #                     logging.warning("Extraction cancelled during go_to_start_page")
    #                     break

    #                 initial_page = adjusted_range_list[i]
    #                 end_page = adjusted_range_list[i+1] if i+1 < len(range_list) else 10000

    #                 if i == 0:
    #                     pass
    #                 else:
    #                     self.ws = self.wb.Worksheets.Add()
    #                     self.row_index = 1
    #                     logging.info("new sheet added")

    #                 update_progress_callback(status=f"Extracting sheets...{i//2 + 1}/{(len(adjusted_range_list)+1)//2}")
    #                 logging.info(f"Extracting Sheet #{i//2+1}")

    #                 update_progress_callback(status=f"Moving to start page {initial_page}...")
    #                 try:
    #                     self.go_to_start_page(initial_page)
    #                     update_progress_callback(status=f"Exporting pages {initial_page} to {end_page}...")
    #                     logging.info(f"Exporting Pages {initial_page}~{end_page}")
    #                     self.copy_paste_to_endpage(end_page, update_progress_callback)
                        
    #                 except Exception as e:
    #                     if trial != self.settings["retryLife"] :
    #                         logging.error(f"Retry failed #{trial}: {e}")
    #                         isRetryEnd = False
    #                         break
    #                     else:
    #                         logging.error(f"Retry failed #{trial}: {e}")
    #                         update_progress_callback(status="Failed Resuming... Please Retry extracting.")
    #                         isRetryEnd = True
    #                         raise Exception("Failed Resuming... Please Retry extracting.")
    #             if isRetryEnd:
    #                 break
    #         else:
    #             logging.info("retry ended. returning.")
    #             return

    def prepare_extraction(self):
        self.reset_state()
        self.open_hwp_file()
        self.open_excel_file()
        self.ctrl = self.hwp.HeadCtrl
        logging.info("Extraction Ready.")

    def extract_tables(self, range_list, update_progress_callback):
        
        self.prepare_extraction()

        self.total_pages = 1 if len(range_list) == 1 else sum(range_list[i+1] - range_list[i] + 1 for i in range(0, len(range_list), 2))
        logging.warning(f"Total Pages : {self.total_pages}")
        self.exported_pages = 0

        for i in range(0, len(range_list), 2):
            logging.info(f"i : {i}")

            if self.cancel_extraction:
                logging.warning("Extraction cancelled by user")
                break
            
            initial_page = range_list[i]
            end_page = range_list[i+1] if i+1 < len(range_list) else 10000
            logging.info(f"initial_page : {initial_page}, end_page : {end_page}")

            if i == 0:
                self.ws = self.wb.Worksheets(1)
            else:
                self.ws = self.wb.Worksheets.Add()
                self.row_index = 1
                logging.info("new sheet added")

            update_progress_callback(status=f"Extracting sheets...{i//2 + 1}/{(len(range_list)+1)//2}")
            logging.info(f"Extracting Sheet #{i//2+1}")

            try:
                update_progress_callback(status=f"Moving to start page {initial_page}...")
                self.go_to_start_page(initial_page)
                update_progress_callback(status=f"Exporting pages {initial_page} to {end_page}...")
                logging.info(f"Exporting Pages {initial_page}~{end_page}")
                self.copy_paste_to_endpage(end_page, update_progress_callback)
            except Exception as e:
                logging.error("restarting disabled.")
                # self.resume_extraction(range_list,update_progress_callback)
                raise Exception("extraction failure")
        
        logging.info("for clause escaped")
        if self.wb != None : 
            self.wb.Save()
            logging.info("excel temp save")

        self.current_page = 1
        logging.info("Page Resetted to 1")

        if self.cancel_extraction:
            if self.wb:
                self.wb.Save()
            self.close_excel_file()
            logging.info("Extraction cancelled, partial results saved.")
            return
            
        update_progress_callback(status="Rearranging Excel...")
        self.rearrange_demos()
        self.wb.Save()
        logging.info("Exportation Successful.")
        update_progress_callback(progress=100, status="Export Completed.")    
        
        if self.settings["doOpenHwp"]:
            if self.hwp:
                self.hwp.set_visible(visible=True)
            else:
                self.open_hwp_file
                self.hwp.set_visible(visible=True)
        else:
            self.close_hwp_file()            

        if self.settings["doOpenXlsx"]:
            self.excel.Visible = True
        else:
            self.close_excel_file()

class GUI:
    def __init__(self, converter):
        self.converter = converter
        self.window = tk.Tk()
        self.setup_ui()
        self.extraction_thread = None
        self.is_extracting = False

    def setup_ui(self):
        self.window.title(f"TableExporter v{VERSION} {"(unstable)" if not IS_STABLE else ""}")
        self.window.resizable(width=False, height=False)

        notebook = ttk.Notebook(self.window, width=505, height=253)
        notebook.pack()

        self.tab1 = tk.Frame(self.window)
        notebook.add(self.tab1, text="  입력  ")
        self.tab2 = tk.Frame(self.window)
        notebook.add(self.tab2, text="  설정  ")

        self.setup_input_tab()
        self.setup_settings_tab()

    def update_progress(self, progress=None, status=None):
        if progress is not None:
            self.progress_var.set(progress)
            self.progress_bar.update()
        if status is not None:
            self.status_text.set(status)
        self.window.update_idletasks()

    def setup_input_tab(self):
        ttk.Label(self.tab1, text="File:", justify="right").place(x=5, y=30, width=51, height=21)
        self.file_text = tk.StringVar()
        self.hwp_entry = ttk.Entry(self.tab1, state="readonly", textvariable=self.file_text)
        self.hwp_entry.place(x=60, y=31, width=331, height=20)
        ttk.Button(self.tab1, text="Choose...", command=self.ask_file).place(x=400, y=30, width=71, height=22)

        ttk.Label(self.tab1, text="Name:", justify="right").place(x=5, y=80, width=51, height=21)
        self.file_name_text = tk.StringVar()
        ttk.Entry(self.tab1, textvariable=self.file_name_text).place(x=60, y=80, width=331, height=20)

        ttk.Label(self.tab1, text="Path:", justify="right").place(x=5, y=110, width=51, height=21)
        self.path_text = tk.StringVar()
        ttk.Entry(self.tab1, state="readonly", textvariable=self.path_text).place(x=60, y=110, width=331, height=20)
        ttk.Button(self.tab1, text="Choose...", command=self.get_export_path).place(x=400, y=109, width=71, height=22)

        ttk.Label(self.tab1, text="Range:").place(x=5, y=140, width=51, height=21)
        self.range_string = tk.StringVar()
        self.range_entry = ttk.Entry(self.tab1, textvariable=self.range_string)
        self.range_entry.place(x=60, y=140, width=331, height=21)
        self.range_entry.insert(0, "추출할 범위 입력. ex) 124:200, 203:400")
        self.range_entry.configure(foreground="gray")
        self.range_entry.bind("<FocusIn>", self.focus_in)
        self.range_entry.bind("<FocusOut>", self.focus_out)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.tab1, maximum=100, mode="determinate", variable=self.progress_var)
        self.progress_bar.place(x=140, y=180, width=300, height=21)

        self.status_text = tk.StringVar()
        ttk.Label(self.tab1, textvariable=self.status_text, justify="left", wraplength= 300).place(x=140, y=201, width=300, height=38)

        self.extract_btn = ttk.Button(self.tab1, text="추출", command=self.toggle_extraction)
        self.extract_btn.place(x=60, y=180, width=71, height=41)

    def setup_settings_tab(self):
        self.is_hwp_visible = tk.IntVar(value=int(self.converter.settings['isHwpVisible']))
        ttk.Checkbutton(self.tab2, text="한글 파일을 백그라운드에서 실행합니다.", variable=self.is_hwp_visible).place(x=10, y=10)
        
        self.is_excel_visible = tk.IntVar(value=int(self.converter.settings['isExcelVisible']))
        ttk.Checkbutton(self.tab2, text="엑셀 파일을 백그라운드에서 실행합니다.",variable = self.is_excel_visible).place(x=10,y=30)
        
        self.do_open_hwp = tk.IntVar(value=int(self.converter.settings['doOpenHwp']))
        ttk.Checkbutton(self.tab2, text="실행 후 한글 파일을 엽니다.", variable=self.do_open_hwp).place(x=10, y=50)

        self.do_open_xlsx = tk.IntVar(value=int(self.converter.settings['doOpenXlsx']))
        ttk.Checkbutton(self.tab2, text="실행 후 엑셀 파일을 엽니다.", variable=self.do_open_xlsx).place(x=10, y=70)
        
        self.special_mode = tk.IntVar(value=int(self.converter.settings['SPMode']))
        ttk.Checkbutton(self.tab2, text="첫 번째 시트를 데모의 위치에 따라 분리합니다.", variable=self.special_mode).place(x=10,y=90)

        # CHANGELOG V1.0.0: 추출 방식 변화로 인한 안정화로 딜레이/재시도 횟수 옵션 삭제.
        # @deprecated 
        # self.copy_paste_delay = tk.StringVar(value=str(self.converter.settings['copyPasteDelay']))
        # ttk.Spinbox(self.tab2,from_= 0, to=1,increment=0.05, wrap=True, textvariable=self.copy_paste_delay ).place(x=100,y=130)
        # ttk.Label(self.tab2, text="딜레이").place(x=10,y=130)
        
        # @deprecated
        # self.retry_life = tk.StringVar(value=str(self.converter.settings['retryLife']))
        # ttk.Spinbox(self.tab2,from_=1,to=10,increment=1,wrap=True,textvariable=self.retry_life).place(x=100,y=155)

        # ttk.Label(self.tab2,text="재시도 횟수").place(x=10,y=155)

        ttk.Button(self.tab2, text="저장", command=self.save_settings).place(x=390, y=190, width=80, height=40)

    def ask_file(self):
        self.converter.file = filedialog.askopenfilename(
            initialdir='/',
            title="파일을 선택해 주세요",
            filetypes=(("HWP files", "*.hwp *.hwpx"),)
        )
        self.file_text.set(self.converter.file)
        self.hwp_entry.configure(foreground="gray")
        self.converter.filename = os.path.splitext(os.path.basename(self.converter.file))[0]
        self.file_name_text.set(f"{self.converter.filename}_변환됨.xlsx")
        
    def get_filename(self):
        self.converter.filename = self.file_name_text.get()

    def get_export_path(self):
        self.converter.export_path = filedialog.askdirectory(
            initialdir='/',
            title="저장할 폴더를 선택해 주세요"
        )
        self.path_text.set(self.converter.export_path)

    def focus_in(self, *args):
        if self.range_string.get() == "추출할 범위 입력. ex) 124:200, 203:400":
            self.range_entry.delete(0, "end")
            self.range_entry.configure(foreground="black")

    def focus_out(self, *args):
        if not self.range_string.get():
            self.range_entry.configure(foreground="gray")
            self.range_entry.insert(0, "추출할 범위 입력. ex) 124:200, 203:400")

    def get_page_range(self):
        range_str = self.range_string.get()
        try:
            return [int(k) for k in re.split('[:,.~ ]', range_str) if k]
        except ValueError:
            raise Exception("invalid Datatype.")
    
    # CHANGELOG V1.0.0: 안정화에 따른 설정 요소 제거 
    def save_settings(self):
        self.converter.settings["isHwpVisible"] = bool(self.is_hwp_visible.get())
        self.converter.settings["isExcelVisible"] = bool(self.is_excel_visible.get())
        self.converter.settings["doOpenHwp"] = bool(self.do_open_hwp.get())
        self.converter.settings["doOpenXlsx"] = bool(self.do_open_xlsx.get())
        self.converter.settings["SPMode"] = bool(self.special_mode.get())
        # self.converter.settings["copyPasteDelay"] = float(self.copy_paste_delay.get())
        # self.converter.settings["retryLife"] = int(self.retry_life.get())
        self.converter.save_settings()

    def toggle_extraction(self):
        if not self.is_extracting:
            self.start_extraction()
        else:
            self.cancel_extraction()

    def start_extraction(self):
        self.is_extracting = True
        self.extract_btn.config(text="취소")
        self.extraction_thread = threading.Thread(target=self.run_extraction, daemon=True)
        self.extraction_thread.start()
        time.sleep(3)
        self.extract_btn.config(state="normal")

    def cancel_extraction(self):
        if self.extraction_thread and self.extraction_thread.is_alive():
            self.is_extracting = False
            self.converter.cancel_extraction = True
            self.extract_btn.config(state=tk.DISABLED)
            self.update_progress(status="Cancelling extraction...")
            self.extraction_thread.join(timeout=3)  # Wait for the thread to finish
            if self.extraction_thread.is_alive():
                logging.warning("Extraction thread did not finish in time")
            self.converter.close_hwp_file()
            self.converter.close_excel_file()
            self.update_progress(status="Extraction cancelled")

    def run_extraction(self):
        try:
            self.get_filename()
            range_list = self.get_page_range()
            self.update_progress(progress=0, status="Starting extraction process...")
            self.converter.extract_tables(range_list, self.update_progress)
            
            if self.converter.cancel_extraction :
                messagebox.showinfo("추출 취소","표 추출이 취소되었습니다.")
                self.extract_btn.config(text="추출", state=tk.NORMAL)
            else:
                messagebox.showinfo("추출 완료", "표 추출이 완료되었습니다.")
            self.update_progress(progress=0, status="...")
        except Exception as e:
            self.update_progress(status=f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            logging.error(f"{e}")
            logging.warning(traceback.format_exc())
        finally:
            self.is_extracting = False
            self.extract_btn.config(text="추출", state=tk.NORMAL)
            self.converter.cancel_extraction = False

    def run(self):
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.mainloop()

    def on_closing(self):
        if hasattr(self.converter, 'excel'):
            try:
                self.converter.wb.Close(SaveChanges=True)
                self.converter.excel.Quit()
            except:
                pass
        if hasattr(self.converter, 'hwp'):
            try:
                self.converter.hwp.Quit()
            except:
                pass
        self.window.destroy()

def main():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    converter = HwpConverter()
    gui = GUI(converter)
    gui.run()

if __name__ == "__main__":
    main()
