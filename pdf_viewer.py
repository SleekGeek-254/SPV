import fitz  # PyMuPDF
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import sys
import subprocess
import os
import ctypes
import socket
import pyodbc
import base64
import pywintypes
import win32print
import win32api
import asyncio
import tempfile
import time
import threading

ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class PDFViewer:
    def __init__(self, root, file_path=None, base64data=None, SUsername=None, SPassword=None, Title=None, Server=None, SDatabase=None, cnSource=None, cnUser=None, cnCatalog=None, cnTitle=None, cnPassword=None, fileTables=None):
        self.root = root
        self.root.title("Secure PDF Viewer")

        # Ensure the window appears on top
        self.root.attributes('-topmost', 1)
        self.root.update()
        self.root.attributes('-topmost', 0)

        # Create a frame for the buttons
        self.button_frame = ctk.CTkFrame(root)
        self.button_frame.pack(side=ctk.TOP, fill=ctk.X, padx=10, pady=5)

        # Create a settings button
        self.settings_button = ctk.CTkButton(self.button_frame, text="Settings", command=self.show_settings_menu)
        self.settings_button.pack(side=ctk.LEFT)
        
        # Create a check for updates button
        self.update_button = ctk.CTkButton(self.button_frame, text="Check for Updates", command=self.check_for_updates)
        self.update_button.pack(side=ctk.LEFT, padx=(10, 0))

        # Create a print button
        self.print_button = ctk.CTkButton(self.button_frame, text="Print", command=self.print_pdf)
        self.print_button.pack(side=ctk.RIGHT, padx=(0, 10))

        # Create a page selection dropdown
        self.page_selection = ctk.CTkOptionMenu(self.button_frame, values=["All Pages", "Current Page", "Custom Range"], command=self.update_print_options)
        self.page_selection.pack(side=ctk.RIGHT, padx=10)

        # Create an entry for custom page range (initially hidden)
        self.custom_range_entry = ctk.CTkEntry(self.button_frame, placeholder_text="e.g., 1-3, 5, 7-9")
        self.custom_range_entry.pack(side=ctk.RIGHT, padx=10)
        self.custom_range_entry.pack_forget()  # Hide initially

        # Create a popup menu
        self.settings_menu = ctk.CTkOptionMenu(self.button_frame, values=["Light Mode", "Dark Mode"], command=self.change_theme)
        self.settings_menu.set("Light Mode")

        # Frame to hold canvas and scrollbar
        self.frame = ctk.CTkFrame(root)
        self.frame.pack(expand=True, fill=ctk.BOTH)

        # Canvas for displaying PDF pages
        self.canvas = ctk.CTkCanvas(self.frame, width=800, height=600)
        self.canvas.pack(side=ctk.LEFT, expand=True, fill=ctk.BOTH)
        self.update_canvas_background()  # Add this line to set initial background

        # Scrollbar for navigation
        self.scrollbar = ctk.CTkScrollbar(self.frame, orientation="vertical", command=self.on_scrollbar)
        self.scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)
        
        # Page number label
        self.page_label = ctk.CTkLabel(root, text="Page: 0 / 0", anchor="ne")
        self.page_label.pack(side=ctk.TOP, anchor="ne", padx=10, pady=5)

        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Control-MouseWheel>", self.on_zoom) 

        # Bind Ctrl key press and release
        self.root.bind("<Control-KeyPress>", self.ctrl_pressed)
        self.root.bind("<Control-KeyRelease>", self.ctrl_released)

        self.doc = None
        self.current_page_index = 0
        self.zoom_level = 1.0  # Default zoom level
        self.ctrl_is_pressed = False

        # Open PDF file or base64 data or from SQL
        if file_path:
            self.open_pdf(file_path)
        elif base64data:
            self.open_pdf_from_base64(base64data)
        elif (SUsername and SPassword and Title and Server and SDatabase):
            self.open_pdf_from_sql(SUsername, SPassword, Title, Server, SDatabase)
        elif (cnSource and cnUser and cnCatalog and cnTitle and cnPassword and fileTables):
            self.open_pdf_from_sql_wth_cn(cnSource, cnUser, cnCatalog, cnTitle, cnPassword, fileTables)
            
    def check_for_updates(self):
        production_commit_hash = "c7c52164085b53f81e7d20616e6c3ae2661ca617"
        if not check_current_version(production_commit_hash):
            choice = messagebox.askyesno("Update Available", "A new version of the application is available. Do you want to update?")
            if choice:
                launch_updater()
        else:
            messagebox.showinfo("Up to Date", "Your application is up to date.")
        
    def update_canvas_background(self):
        if ctk.get_appearance_mode() == "Dark":
            self.canvas.configure(bg="gray20")  # Use a dark color for dark mode
        else:
            self.canvas.configure(bg="white")  # Use white for light mode

    def show_settings_menu(self):
        self.settings_menu.pack(side=ctk.LEFT, padx=(10, 0))

    def change_theme(self, theme):
        if theme == "Dark Mode":
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")
        self.update_canvas_background()
        if self.doc:
            self.display_page(self.current_page_index)  # Redraw the current page

    def update_print_options(self, selection):
        if selection == "Custom Range":
            self.custom_range_entry.pack(side=ctk.RIGHT, padx=10)
        else:
            self.custom_range_entry.pack_forget()
            
            
    async def print_pdf_async(self, temp_doc, selected_printer):
        try:
            # Create a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                temp_filename = temp_file.name

            # Save the temporary document to the temporary file
            temp_doc.save(temp_filename)
            temp_doc.close()

            # print(f"Attempting to print {temp_filename} to {selected_printer}")
            win32api.ShellExecute(0, "print", temp_filename, None, ".", 0)
            # print(f"Printing {temp_filename} to {selected_printer}")
            printer_handle = win32print.OpenPrinter(selected_printer)
            
            # Wait for the print job to complete
            try:
                while True:
                    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                    await asyncio.sleep(1)  # Check every second

                    if not jobs:
                        break
            finally:
                win32print.ClosePrinter(printer_handle)

            # print("Printing completed")
            return temp_filename  # Return the temporary filename for cleanup
        except Exception as e:
            error_message = f"Unexpected error while printing: {e}"
            messagebox.showerror("Print Error", error_message)
            print(error_message)
            return None
    
  
  
    def print_pdf(self):
        if self.doc:
            try:
             

                # Get available printers
                printers = [printer[2] for printer in win32print.EnumPrinters(2) if printer[2] != "Microsoft Print to PDF"]

                # Create a dialog to select a printer
                printer_dialog = ctk.CTkToplevel(self.root)
                printer_dialog.title("Select Printer")
                printer_dialog.geometry("300x200")

                printer_var = ctk.StringVar(value=printers[0] if printers else "")
                printer_menu = ctk.CTkOptionMenu(printer_dialog, values=printers, variable=printer_var)
                printer_menu.pack(pady=20)

                def on_printer_select():
                    selected_printer = printer_var.get()
                    printer_dialog.destroy()

                    # Set the selected printer as default
                    try:
                        win32print.SetDefaultPrinter(selected_printer)
                        print(f"Default printer set to: {selected_printer}")
                    except Exception as e:
                        print(f"Error setting default printer: {e}")
                        messagebox.showwarning("Invalid Range", f" Error setting default printer: {e}")
                        return

                    # Get print range
                    selection = self.page_selection.get()
                    if selection == "All Pages":
                        print_range = range(len(self.doc))
                    elif selection == "Current Page":
                        print_range = [self.current_page_index]
                    elif selection == "Custom Range":
                        custom_range = self.custom_range_entry.get()
                        print_range = self.parse_page_range(custom_range)

                    if print_range:
                        # Create a temporary PDF with selected pages
                        temp_doc = fitz.open()
                        for page_num in print_range:
                            temp_doc.insert_pdf(self.doc, from_page=page_num, to_page=page_num)
   
                        # Print the temporary PDF asynchronously
                        temp_filename = asyncio.run(self.print_pdf_async(temp_doc, selected_printer))
                        
                        # self.root.after(5000, lambda: messagebox.showinfo("Print", f"Printing complete. Cleaning Cache {selected_printer}"))
                                   # Define the cleanup function
                        def delayed_cleanup():
                            if temp_filename:
                                max_retries = 100000000000000000000000
                                retry_delay = 1  # seconds
                                for attempt in range(max_retries):
                                    try:
                                        os.remove(temp_filename)
                                        print(f"Temporary file {temp_filename} removed")
                                        break
                                    except PermissionError as e:
                                        if e.winerror == 32:  # WinError 32: The process cannot access the file because it is being used by another process
                                            if attempt < max_retries - 1:
                                                # print(f"Failed to remove file, retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                                                time.sleep(retry_delay)
                                            else:
                                                print(f"Failed to remove temporary file after {max_retries} attempts: {e}")
                                        else:
                                            print(f"Error removing temporary file: {e}")
                                            break
                                    except Exception as e:
                                        print(f"Unexpected error removing temporary file: {e}")
                                        break
                            
                                    # Show message box after cleanup attempt
                            self.root.after(0, lambda: messagebox.showinfo("Print", f"Printing complete. Cleaning Cache {selected_printer}"))

                        # Create and start the cleanup thread
                        cleanup_thread = threading.Thread(target=delayed_cleanup)
                        cleanup_thread.start()
                        
                    else:
                        messagebox.showwarning("Invalid Range", "Please enter a valid page range.")

                select_button = ctk.CTkButton(printer_dialog, text="Print", command=on_printer_select)
                select_button.pack(pady=20)

                printer_dialog.transient(self.root)
                printer_dialog.grab_set()
                self.root.wait_window(printer_dialog)

            except ImportError:
                messagebox.showerror("Error", "Printing requires the win32print and win32api modules.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to print PDF: {e}")
        else:
            messagebox.showinfo("Print", "No PDF document is currently open.")
        
        
        
        
    def parse_page_range(self, range_str):
        pages = set()
        try:
            parts = range_str.split(',')
            for part in parts:
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    pages.update(range(start-1, end))
                else:
                    pages.add(int(part)-1)
            return sorted(list(pages))
        except ValueError:
            return None


    def open_pdf(self, file_path):
        if file_path:
            try:
                self.doc = fitz.open(file_path)
                self.setup_scrollbar()
                self.display_page(self.current_page_index)
                self.update_page_label()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF file: {e}")
                
                
    def open_pdf_from_base64(self, base64data):
        if base64data:
            try:
                pdf_data = base64.b64decode(base64data)
                self.doc = fitz.open("pdf", pdf_data)
                self.setup_scrollbar()
                self.display_page(self.current_page_index)
                self.update_page_label()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF from base64 data: {e}")

    def open_pdf_from_sql(self, SUsername, SPassword, Title, Server, SDatabase):
        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                f'Server={Server};'
                                f'Database={SDatabase};'
                                f'UID={SUsername};'
                                f'PWD={SPassword};')
            
            cursor = conn.cursor()
            cursor.execute("SELECT file_stream, name FROM Library_FileTable WHERE name = ?", (Title,))
            row = cursor.fetchone()
            
            if row:
                file_data, filename = row
                self.doc = fitz.open(stream=file_data)
                self.setup_scrollbar()
                self.display_page(self.current_page_index)
                self.update_page_label()
            else:
                messagebox.showerror("Error", f"No PDF found with title '{Title}'")
            
            cursor.close()
            conn.close()
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Database error: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF from SQL: {e}")

    def open_pdf_from_sql_wth_cn(self, cnSource, cnUser, cnCatalog, cnTitle, cnPassword, fileTables):
        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                f'Server={cnSource};'
                                f'Database={cnCatalog};'
                                'UID=sa;'
                                'PWD=Mmggit22;')
            
            cursor = conn.cursor()
            cursor.execute(f"SELECT file_stream, name FROM {fileTables} WHERE name = ?", (cnTitle,))
            row = cursor.fetchone()
            
            if row:
                file_data, filename = row
                self.doc = fitz.open(stream=file_data)
                self.setup_scrollbar()
                self.display_page(self.current_page_index)
                self.update_page_label()
            else:
                messagebox.showerror("Error", f"No PDF found with title '{cnTitle}'")
            
            cursor.close()
            conn.close()
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Database error: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF from SQL with CN: {e}")

    # ----
    def setup_scrollbar(self):
        if self.doc:
            self.scrollbar.configure(command=self.on_scrollbar)
            self.scrollbar.set(0, 1 / len(self.doc))

    def on_scrollbar(self, *args):
        if self.doc:
            page_index = int(float(args[1]) * len(self.doc))
            if page_index != self.current_page_index:
                self.current_page_index = page_index
                self.display_page(self.current_page_index)

    def display_page(self, page_index):
        if self.doc:
            try:
                page = self.doc.load_page(page_index)
                zoom_matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
                pix = page.get_pixmap(matrix=zoom_matrix)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img_tk = ImageTk.PhotoImage(img)

                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()

                # Calculate image position for centering
                x = max(0, (canvas_width - img_tk.width()) // 2)
                y = max(0, (canvas_height - img_tk.height()) // 2)

                # Clear previous page
                self.canvas.delete("all")

                # Add a background rectangle
                bg_color = "gray20" if ctk.get_appearance_mode() == "Dark" else "white"
                self.canvas.create_rectangle(0, 0, canvas_width, canvas_height, fill=bg_color, outline="")

                # Display the image
                self.canvas.create_image(x, y, anchor=ctk.NW, image=img_tk)
                self.canvas.image = img_tk   # Keep a reference to avoid garbage collection

                total_pages = len(self.doc)
                
                if total_pages > 1:
                    # Enable scrolling for multi-page documents
                    scroll_width = max(canvas_width, img_tk.width())
                    scroll_height = max(canvas_height, img_tk.height())
                    self.canvas.configure(scrollregion=(0, 0, scroll_width, scroll_height))
                    
                    # Update vertical scrollbar
                    start = page_index / (total_pages - 1)
                    end = (page_index + 1) / total_pages
                    self.scrollbar.set(start, end)
                    self.scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)
                    self.canvas.configure(yscrollcommand=self.scrollbar.set)
                    
                    # Add horizontal scrollbar if needed
                    if img_tk.width() > canvas_width:
                        if not hasattr(self, 'hscrollbar'):
                            self.hscrollbar = ctk.CTkScrollbar(self.frame, orientation="horizontal")
                            self.hscrollbar.pack(side=ctk.BOTTOM, fill=ctk.X)
                            self.hscrollbar.configure(command=self.canvas.xview)
                        self.hscrollbar.pack(side=ctk.BOTTOM, fill=ctk.X)
                        self.canvas.configure(xscrollcommand=self.hscrollbar.set)
                    elif hasattr(self, 'hscrollbar'):
                        self.hscrollbar.pack_forget()
                        
                else:
                    # Disable scrolling for single-page documents
                    self.canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))
                    self.scrollbar.pack_forget()
                    self.canvas.configure(yscrollcommand=None)
                    if hasattr(self, 'hscrollbar'):
                        self.hscrollbar.pack_forget()
                    self.canvas.configure(xscrollcommand=None)

                # Update page label
                self.update_page_label()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to display page: {e}")              
    
    def update_page_label(self):
        if self.doc:
            total_pages = len(self.doc)
            current_page = self.current_page_index + 1  # Add 1 because page_index is 0-based
            self.page_label.configure(text=f"Page: {current_page} / {total_pages}")

            
    def on_resize(self, event):
        if self.doc:
            self.display_page(self.current_page_index)

    def on_mouse_wheel(self, event):
        if self.doc:
            # Navigate pages
            if event.delta > 0:
                self.current_page_index = max(0, self.current_page_index - 1)
            else:
                self.current_page_index = min(len(self.doc) - 1, self.current_page_index + 1)
            self.display_page(self.current_page_index)
            self.update_page_label()

    def on_zoom(self, event):
        if self.doc:
            # Zoom
            if event.delta > 0:
                self.zoom_level *= 1.1
            else:
                self.zoom_level /= 1.1
            self.zoom_level = max(0.1, min(self.zoom_level, 5.0))  # Limit zoom between 10% and 500%
            self.display_page(self.current_page_index)
            
    def ctrl_pressed(self, event):
        self.ctrl_is_pressed = True

    def ctrl_released(self, event):
        self.ctrl_is_pressed = False
        
        
def check_internet_connection(host="8.8.8.8", port=53, timeout=3):
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(f"Internet connection check failed: {ex}")
        return False

def check_current_version(commit_hash, branch='production'):
    if not check_internet_connection():
        print("No internet connection. Cannot check for updates.")
        return True

    try:
        # Run 'git ls-remote' to get the latest commit hash of the repository
        result = subprocess.run(
            ['git', 'ls-remote', '--heads', 'https://github.com/SleekGeek-254/SPV.git', branch],
            capture_output=True, text=True
        )
        if result.stderr:
            print("Error during git operation:", result.stderr)
            return False

        latest_commit_hash = result.stdout.split()[0]
        print(f"Latest commit hash for branch '{branch}': {latest_commit_hash}")

        # Check if the latest commit hash matches the provided one
        if latest_commit_hash == commit_hash:
            print("The application is up to date. Proceeding...")
            return True
        else:
            print("A new version is available. Please update.")
            return False

    except Exception as e:
        print("Error occurred while checking the current version:", e)
        return False

def launch_updater():
    # Get the directory where this Python script or executable is located
    if getattr(sys, 'frozen', False):  # if running as compiled executable
        script_dir = os.path.dirname(sys.executable)
    else:  # if running as a script
        script_dir = os.path.dirname(__file__)
    
    # Change the current working directory to the script's directory
    os.chdir(script_dir)

    # Constructing the path to "spvupdater.exe" in the parent directory
    updater_path = os.path.join(os.path.dirname(script_dir), "spvupdater.exe")
    print(updater_path)
    
    if os.path.exists(updater_path):
        shell32 = ctypes.windll.shell32
        shell32.ShellExecuteW(None, "runas", updater_path, None, None, 1)
        return
    else:
        messagebox.showerror("Updater not found", "Updater executable not found.")


def validate_connection_string(conn_str):
    required_keys = ["Initial Catalog", "Source", "User ID"]
    key_values = {}
    # cnUser, SPassword, Title , cnSource, SDatabase
    for key in required_keys:
        key_start = conn_str.find(key)
        if key_start == -1:
            print(f"Error: '{key}' not found in the connection string.")
            return False
        key_end = conn_str.find(';', key_start)
        if key_end == -1:
            key_value = conn_str[key_start:].split('=', 1)[1].strip()
        else:
            key_value = conn_str[key_start:key_end].split('=', 1)[1].strip()

        if not key_value:
            print(f"Error: '{key}' value is empty in the connection string.")
            return False

        key_values[key] = key_value

    print("All required keys are present and non-empty.")
    return True, key_values

def main():
    # Pdf Viewer
    if len(sys.argv) > 1:
        print("sys.argv > ", sys.argv)

        if sys.argv[1].endswith(".pdf"):
            print("CASE 1 ")
            file_path = sys.argv[1]
            base64data = None
            SUsername = None
            SPassword = None
            Title = None
            Server = None
            SDatabase = None
            cnSource = None
            cnUser = None
            cnCatalog= None
            cnTitle = None
            cnPassword = None
            fileTables= None

        elif len(sys.argv) > 3 and (not sys.argv[1].__contains__("Provider") and not sys.argv[1].__contains__("Source") and not sys.argv[1].__contains__("User ID")):
            print("CASE 2 ")
            file_path = None
            base64data = None
            SUsername = sys.argv[1]
            SPassword = sys.argv[2]
            Title = sys.argv[3]
            Server = sys.argv[4]
            SDatabase = sys.argv[5]
            cnSource = None
            cnUser = None
            cnCatalog= None
            cnTitle = None
            cnPassword = None
            fileTables= None

        elif (sys.argv[1].__contains__("Provider") and sys.argv[1].__contains__("Source") and sys.argv[1].__contains__("User ID")) and (len(sys.argv) > 4):
            print("CASE 3 ")
            file_path = None
            base64data = None
            SUsername = None
            SPassword = None
            Title = None
            Server = None
            SDatabase = None
            conn_str = sys.argv[1]
            cnTitle = sys.argv[2]
            cnPassword = sys.argv[3]
            fileTables = sys.argv[4]
            is_valid, key_values = validate_connection_string(conn_str)
            if is_valid:
                print("Connection string is valid.")
                print("Initial Catalog value:", key_values.get("Initial Catalog"))
                print("Source value:", key_values.get("Source"))
                print("User ID value:", key_values.get("User ID"))
                cnCatalog = key_values.get("Initial Catalog")
                cnSource = key_values.get("Source")
                cnUser = key_values.get("User ID")
            else:
                print("Connection string is invalid.")

        else:
            print("CASE 4 ")
            file_path = None
            base64data = sys.argv[1]
            SUsername = None
            SPassword = None
            Title = None
            Server = None
            SDatabase = None
            cnSource = None
            cnUser = None
            cnCatalog= None
            cnTitle = None
            cnPassword = None
            fileTables= None

    else:
        file_path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF Files", "*.pdf")]
        )
        base64data = None
        SUsername = None
        SPassword = None
        Title = None
        Server = None
        SDatabase = None
        cnSource = None
        cnUser = None
        cnCatalog= None
        cnTitle = None
        cnPassword = None
        fileTables= None

    if file_path or base64data or (SUsername and SPassword and Title and Server and SDatabase ) or (        cnSource  and cnUser  and cnCatalog and  cnTitle  and  cnPassword and fileTables  ):
        root = ctk.CTk()
        app = PDFViewer(root, file_path=file_path, base64data=base64data, SUsername=SUsername, SPassword=SPassword, Title=Title, Server=Server, SDatabase=SDatabase, cnSource=cnSource, cnUser=cnUser, cnCatalog=cnCatalog, cnTitle=cnTitle, cnPassword=cnPassword, fileTables=fileTables)
        root.mainloop()

if __name__ == "__main__":
    main()