import os
import time
import re
import json
import threading
import logging

import win32com.client as win32
from pyhwpx import Hwp
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Constants
VERSION = "0.1.3"
IS_STABLE = False

DATA_DIR = "data"
SETTINGS_FILE = os.path.join(DATA_DIR,'settings.json')
LOG_FILE = os.path.join(DATA_DIR,'hwp_converter.log')
DEFAULT_SETTINGS = {
    "isHwpVisible": True,
    "doOpenHwp": True,
    "doOpenXlsx": True,
    "copyPasteDelay" : 0.1
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
        try:
            self.close_excel_file()
            save_file = (self.export_path + "/" + self.filename +"_변환됨" + ".xlsx").replace("/","\\")
            save_file = self.get_unique_filename(filename=save_file)
            
            self.excel = win32.gencache.EnsureDispatch("Excel.Application")
            self.excel.Visible = False
            self.wb = self.excel.Workbooks.Add()
            self.wb.SaveAs(save_file)
            self.excel.Quit()
            self.wb = self.excel.Workbooks.Open(save_file)
        except Exception as e:
            logging.error(f"Creating New Excel file failed : {e}")
            raise Exception("Opening Excel failed.")
    
    def close_excel_file(self):
        if hasattr(self, 'wb') and self.wb:
            try:
                self.wb.Close(SaveChanges=True)
            except:
                pass  # If there's an error closing the workbook, continue to close Excel
        if hasattr(self, 'excel') and self.excel:
            try:
                self.excel.Quit()
            except:
                pass  # If there's an error quitting Excel, we've done our best
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
                    self.ctrl = self.ctrl.Next.Next.Next
                    if self.ctrl is None:
                        break
                    if self.ctrl.UserDesc == "표":
                        self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
                        self.current_page = self.hwp.current_page
                else:
                    self.ctrl = self.ctrl.Prev
                    self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
                    self.temp_cur_page = self.hwp.current_page
                    self.hwp.SetPosBySet(self.ctrl.Prev.GetAnchorPos(0))
                    self.temp_prev_page = self.hwp.current_page
                    if self.temp_cur_page != self.temp_prev_page :
                        self.current_page = self.temp_cur_page
                self.hwp.goto_page(self.current_page)

            if self.cancel_extraction:
                logging.info("Extraction cancelled during go_to_start_page")
                return

            logging.info(f"Moved to Start Page {self.current_page}")
        except Exception as e:
            logging.error(f"Failed while Moving to Start Page, current Page : {self.current_page}, E: {e}")
            raise Exception("Moving to Start page Failed.")
            
    def copy_paste_action(self):
        self.hwp.SetPosBySet(self.ctrl.GetAnchorPos(0))
        self.current_page = self.hwp.current_page
        self.hwp.FindCtrl()
        time.sleep(self.settings["copyPasteDelay"]/2)
        self.hwp.Copy()
        time.sleep(self.settings["copyPasteDelay"]/2)
        
        self.ws.Activate()
        if self.excel.ClipboardFormats:
            self.excel.ActiveSheet.Paste()
            logging.info(f"Paste Successful. current Page: {self.current_page}")
        else:
            logging.warning(f"Paste failed - nothing on Clipboard. current Page : {self.current_page}")
            raise Exception("Paste failed: Nothing on Clipboard.")

    def copy_paste_to_endpage(self, end_page, update_progress_callback):
        if self.cancel_extraction:
            logging.info("Extraction cancelled before copy-paste")
            return
        
        
        logging.info("Copy-Paste Started")
        while end_page > self.current_page:
            if self.cancel_extraction:
                logging.info("Extraction cancelled during pre-copy")
                break       
            if self.ctrl is None:
                break
            if self.ctrl.CtrlID == "tbl":
                for attempt in range(1, 6):
                    if self.cancel_extraction:
                        logging.info("Extraction cancelled during copy_paste_to_endpage")
                        break
                    try:
                        self.copy_paste_action()
                        self.exported_pages += 1
                        progress = (self.exported_pages / self.total_pages) * 100
                        update_progress_callback(progress=progress, status=f"Exporting page {self.current_page}...")
                        break
                    except Exception as e:
                        if attempt == 5:
                            logging.error(f"Copy-Paste failed: {e}")
                            raise Exception(f"Failed to copy-paste after 5 attempts: {e}")
                        time.sleep(0.3)

                if self.cancel_extraction:
                    logging.info("Extraction cancelled before excel offset")
                    return
    
                try:
                    self.hwp.HAction.Run("ShapeObjTableSelCell")
                    row_num = self.hwp.get_row_num()
                    self.hwp.HAction.Run("Cancel")
                    self.row_index += row_num + 8
                    self.ws.Range(f"A{self.row_index}").Select()
                except Exception as e:
                    self.close_excel_file()
                    logging.error(f"Row Calculation failed: {e}")
                    raise Exception("Error calculating rows")

            self.ctrl = self.ctrl.Next
            if self.cancel_extraction:
                logging.info("Extraction Cancelled after processing a table")
                return

    def rearrange_excel(self):
        try:
            logging.info("Started rearranging excel")
            for sheet in self.wb.Sheets:
                used_range = sheet.UsedRange
                rows = used_range.Rows.Count
                columns = used_range.Columns.Count

                row = 1
                while row <= rows:
                    if all(sheet.Cells(row, col).Value is None for col in range(1, columns + 1)):
                        next_non_empty_row = row + 1
                        while next_non_empty_row <= rows and all(sheet.Cells(next_non_empty_row, col).Value is None for col in range(1, columns + 1)):
                            next_non_empty_row += 1

                        sheet.Rows(f"{row}:{next_non_empty_row - 1}").Delete(Shift=win32.constants.xlUp)
                        sheet.Rows(row).Insert(Shift=win32.constants.xlDown)
                        sheet.Rows(row).Insert(Shift=win32.constants.xlDown)

                        rows -= (next_non_empty_row - row)
                        rows += 2
                        row += 2

                        all_empty = all(all(sheet.Cells(check_row, col).Value is None for col in range(1, columns + 1)) for check_row in range(row, min(row + 10, rows + 1)))
                        if all_empty:
                            sheet.Cells(1, 1).Select()
                            return
                    else:
                        row += 1
            self.wb.Sheets(1).Select()
        except Exception as e:
            logging.error(f"Re-arranging failed: {e}")
            raise Exception("Re-arranging Excel failed.")
    
    def is_number(self,cell_value):
        try:
            float(cell_value)
            return True
        except:
            return False
    
    def rearrange_demos(self) :
        try:
            for sheet in self.wb.Worksheets:
                used_range = sheet.UsedRange
                total_used_row = used_range.Rows.Count

                row = 1
                while row <= total_used_row : 
                    if sheet.Cells(row,1).Value is None :
                        row += 1
                    else:
                        table = sheet.Cells(row,1).CurrentRegion                
                        table_row = table.Rows.Count
                        table_column = table.Columns.Count
                        
                        for cell in table :
                            if cell.MergeCells:
                                cell.UnMerge()
                        
                        right_column = sheet.Range(sheet.Cells(row,table_column),sheet.Cells(row+table_row-1,table_column))
                        if all((not self.is_number(str(cell.Value))) or cell.Value is None for cell in right_column) :
                            print("demo")
                            demo_value = right_column.Value
                            dest = sheet.Range(table.Offset(1,2),table.Offset(table_row,table_column+1))
                            dest.Value = table.Value
                            new_demo = sheet.Range(table.Offset(1,1),table.Offset(table_row,1))
                            new_demo.Value = demo_value
                            sheet.Range(table.Offset(1,table_column+1),table.Offset(table_row,table_column+1)).Value = ""
                            
                        else:
                            pass
                        row += table_row

        except Exception as e:
            logging.error(f"Re-arranging Demo failed: {e}")
            raise Exception("Re-arranging Excel failed.")
        

    def resume_extraction(self, range_list, update_progress_callback):
        for trial in range(1,4):
            self.wb.Save()
            self.close_hwp_file()

            if self.cancel_extraction:
                logging.info("Extraction cancelled during resume_extraction")
                return

            update_progress_callback(status=f"Re-Trying extraction from page {self.current_page}...")
            logging.info(f"Retrying extraction from page {self.current_page}")
            time.sleep(0.5)
            self.open_hwp_file()
            self.ctrl = self.hwp.HeadCtrl

            current_range_index = next((i for i in range(0, len(range_list), 2) if range_list[i] <= self.current_page <= range_list[i+1]), None)

            adjusted_range_list = range_list[current_range_index:]
            adjusted_range_list[0] = self.current_page
            self.current_page = 1
            logging.warning(f"adjusted range list : {adjusted_range_list}")
            
            self.row_index += 40
            self.ws.Range(f"A{self.row_index}").Select()
            for i in range(0, len(adjusted_range_list), 2):
                if self.cancel_extraction:
                    logging.info("Extraction cancelled during go_to_start_page")
                    break

                initial_page = adjusted_range_list[i]
                end_page = adjusted_range_list[i+1] if i+1 < len(range_list) else 10000

                if i == 0:
                    pass
                else:
                    self.ws = self.wb.Worksheets.Add()

                update_progress_callback(status=f"Extracting sheets...{i//2 + 1}/{(len(adjusted_range_list)+1)//2}")
                logging.info(f"Extracting Sheet #{i//2+1}")

                update_progress_callback(status=f"Moving to start page {initial_page}...")
                self.go_to_start_page(initial_page)
            try:
                update_progress_callback(status=f"Exporting pages {initial_page} to {end_page}...")
                logging.info(f"Exporting Pages {initial_page}~{end_page}")
                self.copy_paste_to_endpage(end_page, update_progress_callback)
            except Exception as e:
                if trial != 3 :
                    logging.error(f"Retry failed #{trial}: {e}")
                    pass
                else:
                    logging.error(f"Retry failed #{trial}: {e}")
                    update_progress_callback(status="Failed Resuming... Please Retry extracting.")
                    raise Exception("Failed Resuming... Please Retry extracting.")
        
    def prepare_extraction(self):
        self.reset_state()
        self.open_hwp_file()
        self.open_excel_file()
        self.ctrl = self.hwp.HeadCtrl
        logging.info("Extraction Ready.")

    def extract_tables(self, range_list, update_progress_callback):
        
        self.prepare_extraction()

        self.total_pages = sum(range_list[i+1] - range_list[i] + 1 for i in range(0, len(range_list), 2))
        self.exported_pages = 0

        for i in range(0, len(range_list), 2):
            
            if self.cancel_extraction:
                logging.info("Extraction cancelled by user")
                break
            
            initial_page = range_list[i]
            end_page = range_list[i+1] if i+1 < len(range_list) else 10000

            if i == 0:
                self.ws = self.wb.Worksheets(1)
            else:
                self.ws = self.wb.Worksheets.Add()

            update_progress_callback(status=f"Extracting sheets...{i//2 + 1}/{(len(range_list)+1)//2}")
            logging.info(f"Extracting Sheet #{i//2+1}")

            update_progress_callback(status=f"Moving to start page {initial_page}...")
            self.go_to_start_page(initial_page)
            try:
                update_progress_callback(status=f"Exporting pages {initial_page} to {end_page}...")
                logging.info(f"Exporting Pages {initial_page}~{end_page}")
                self.copy_paste_to_endpage(end_page, update_progress_callback)
            except:
                self.resume_extraction(range_list,update_progress_callback)
            
        if self.wb != None : 
            self.wb.Save()

        if not self.cancel_extraction:
            update_progress_callback(status="Rearranging Excel...")
            self.rearrange_excel()
            update_progress_callback(status="Rearranging Demo...")
            self.rearrange_demos()
            self.wb.Save()
            logging.info("Exportation Successful.")
            update_progress_callback(progress=100, status="Export Completed.")
        else:
            self.close_excel_file()
            logging.info("Extraction cancelled, partial results saved.")
            
        if not self.cancel_extraction : 
            if self.settings["doOpenHwp"]:
                self.hwp.set_visible(visible=True)
            else:
                self.close_hwp_file()            

            if self.settings["doOpenXlsx"]:
                self.excel.Visible = True
            else:
                self.close_excel_file()

        self.current_page = 1
        logging.info("Page Resetted to 1")

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
        ttk.Checkbutton(self.tab2, text="한글 파일을 백그라운드에서 실행합니다 (unstable)", variable=self.is_hwp_visible).place(x=10, y=10)

        self.do_open_hwp = tk.IntVar(value=int(self.converter.settings['doOpenHwp']))
        ttk.Checkbutton(self.tab2, text="실행 후 한글 파일을 엽니다.", variable=self.do_open_hwp).place(x=10, y=30)

        self.do_open_xlsx = tk.IntVar(value=int(self.converter.settings['doOpenXlsx']))
        ttk.Checkbutton(self.tab2, text="실행 후 엑셀 파일을 엽니다.", variable=self.do_open_xlsx).place(x=10, y=50)

        self.copy_paste_delay = tk.StringVar(value=str(self.converter.settings['copyPasteDelay']))
        ttk.Spinbox(self.tab2,from_= 0, to=1,increment=0.05, wrap=True, textvariable=self.copy_paste_delay ).place(x=60,y=100)

        ttk.Label(self.tab2, text="딜레이").place(x=10,y=100)

        ttk.Button(self.tab2, text="저장", command=self.save_settings).place(x=400, y=200, width=80, height=30)

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

    def save_settings(self):
        self.converter.settings["isHwpVisible"] = bool(self.is_hwp_visible.get())
        self.converter.settings["doOpenHwp"] = bool(self.do_open_hwp.get())
        self.converter.settings["doOpenXlsx"] = bool(self.do_open_xlsx.get())
        self.converter.settings["copyPasteDelay"] = float(self.copy_paste_delay.get())
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
        print("cancel")
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
            range_list = self.get_page_range()

            self.converter.prepare_extraction()

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