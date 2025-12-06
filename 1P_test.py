import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import os
import shutil
from pathlib import Path

class ExcelMacroRunner:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Macro Runner - 1P Test")
        self.root.geometry("600x400")
        
        self.excel_app = None
        self.first_file_path = None
        self.first_file_folder = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Macro Automation Tool", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Step 1: First file upload and macro execution
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Process Main File", padding="10")
        step1_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(step1_frame, text="Upload .xlsm file and run Khalil + 1P3P macros:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.upload_btn1 = ttk.Button(step1_frame, text="Upload & Process File", 
                                     command=self.process_first_file)
        self.upload_btn1.grid(row=1, column=0, pady=5)
        
        self.status_label1 = ttk.Label(step1_frame, text="No file selected", foreground="gray")
        self.status_label1.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Step 2: Second file upload and macro execution
        step2_frame = ttk.LabelFrame(main_frame, text="Step 2: Process Add-On File", padding="10")
        step2_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(step2_frame, text="Upload Add-On file and run After RUT macro:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.upload_btn2 = ttk.Button(step2_frame, text="Upload & Process Add-On", 
                                     command=self.process_second_file, state="disabled")
        self.upload_btn2.grid(row=1, column=0, pady=5)
        
        self.status_label2 = ttk.Label(step2_frame, text="Complete Step 1 first", foreground="gray")
        self.status_label2.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        
        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Process Log", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=8, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def start_excel(self):
        if not self.excel_app:
            try:
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.log_message("Excel application started")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to start Excel: {str(e)}")
                return False
        return True
        
    def process_first_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel file (.xlsm)",
            filetypes=[("Excel Macro files", "*.xlsm"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        self.first_file_path = file_path
        self.first_file_folder = os.path.dirname(file_path)
        
        try:
            self.progress.start()
            self.log_message(f"Processing file: {os.path.basename(file_path)}")
            
            if not self.start_excel():
                return
                
            # Open workbook
            workbook = self.excel_app.Workbooks.Open(file_path)
            self.log_message("File opened successfully")
            
            # Run Khalil macro
            self.log_message("Running Khalil macro...")
            try:
                self.excel_app.Run("Khalilmacro")
                self.log_message("Khalil macro completed")
            except Exception as e:
                self.log_message(f"Khalil macro error: {str(e)}")
                
            # Run 1P3P macro
            self.log_message("Running Macro1P3PNewFile...")
            try:
                self.excel_app.Run("Macro1P3PNewFile")
                self.log_message("Macro1P3PNewFile completed")
            except Exception as e:
                self.log_message(f"Macro1P3PNewFile error: {str(e)}")
            
            # Save with new name
            file_name = Path(file_path).stem
            new_file_path = os.path.join(self.first_file_folder, f"{file_name}_Re-run.xlsm")
            workbook.SaveAs(new_file_path)
            self.log_message(f"File saved as: {os.path.basename(new_file_path)}")
            
            workbook.Close()
            
            # Update UI
            self.status_label1.config(text=f"Processed: {os.path.basename(new_file_path)}", 
                                    foreground="green")
            self.upload_btn2.config(state="normal")
            self.status_label2.config(text="Ready for Add-On file", foreground="blue")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file: {str(e)}")
            self.log_message(f"Error: {str(e)}")
        finally:
            self.progress.stop()
            
    def process_second_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Add-On Excel file (.xlsm)",
            filetypes=[("Excel Macro files", "*.xlsm"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            self.progress.start()
            self.log_message(f"Processing Add-On file: {os.path.basename(file_path)}")
            
            if not self.start_excel():
                return
                
            # Open the Re-run file first
            rerun_file = os.path.join(self.first_file_folder, 
                                    f"{Path(self.first_file_path).stem}_Re-run.xlsm")
            
            if not os.path.exists(rerun_file):
                messagebox.showerror("Error", "Re-run file not found. Please complete Step 1 first.")
                return
                
            workbook1 = self.excel_app.Workbooks.Open(rerun_file)
            self.log_message("Re-run file opened")
            
            # Open Add-On file
            workbook2 = self.excel_app.Workbooks.Open(file_path)
            self.log_message("Add-On file opened")
            
            # Run After RUT macro
            self.log_message("Running MacroAfterRUT1P3P...")
            try:
                self.excel_app.Run("MacroAfterRUT1P3P")
                self.log_message("MacroAfterRUT1P3P completed")
            except Exception as e:
                self.log_message(f"MacroAfterRUT1P3P error: {str(e)}")
            
            # Save final file
            final_file_name = f"{Path(self.first_file_path).stem}_Final.xlsm"
            final_file_path = os.path.join(self.first_file_folder, final_file_name)
            workbook1.SaveAs(final_file_path)
            self.log_message(f"Final file saved as: {final_file_name}")
            
            # Close workbooks
            workbook2.Close(SaveChanges=False)
            workbook1.Close()
            
            # Update UI
            self.status_label2.config(text=f"Final file created: {final_file_name}", 
                                    foreground="green")
            
            messagebox.showinfo("Success", f"Process completed!\nFinal file: {final_file_name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Add-On file: {str(e)}")
            self.log_message(f"Error: {str(e)}")
        finally:
            self.progress.stop()
            
    def __del__(self):
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass

def main():
    root = tk.Tk()
    app = ExcelMacroRunner(root)
    
    def on_closing():
        if app.excel_app:
            try:
                app.excel_app.Quit()
            except:
                pass
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
            self.progress.stop()
            
    def __del__(self):
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass

def main():
    root = tk.Tk()
    app = ExcelMacroRunner(root)
    
    def on_closing():
        if app.excel_app:
            try:
                app.excel_app.Quit()
            except:
                pass
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()