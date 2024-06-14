import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import sys
import subprocess
import os
import ctypes
import socket
import pyodbc
import base64

class PDFViewer:
    def __init__(self, root, file_path=None, base64data=None , SUsername=None, SPassword=None,Title=None, Server=None, SDatabase=None , cnSource = None, cnUser = None , cnCatalog = None, cnTitle = None , cnPassword=None, fileTables= None ):
        self.root = root
        self.root.title("Secure PDF Viewer")

        # Ensure the window appears on top
        self.root.attributes('-topmost', 1)
        self.root.update()
        self.root.attributes('-topmost', 0)

        # Canvas for displaying PDF pages
        self.canvas = tk.Canvas(root, width=800, height=600)
        self.canvas.pack(expand=True, fill=tk.BOTH)
        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)

        # Navigation buttons
        self.prev_button = tk.Button(root, text="Previous", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.next_button = tk.Button(root, text="Next", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=10)

        self.doc = None
        self.current_page_index = 0
        self.zoom_level = 1.0  # Default zoom level

        # Open PDF file or base64 data or from SQL
        if file_path:
            self.open_pdf(file_path)
        elif base64data:
            self.open_pdf_from_base64(base64data)
        elif (SUsername and SPassword and Title and Server and SDatabase ):
            self.open_pdf_from_sql(SUsername, SPassword, Title, Server, SDatabase)
        elif (cnSource and cnUser and cnCatalog and cnTitle  and  cnPassword and fileTables):
            self.open_pdf_from_sql_wth_cn(cnSource , cnUser , cnCatalog , cnTitle, cnPassword, fileTables)

    # ----
    def open_pdf(self, file_path):
        if file_path:
            try:
                self.doc = fitz.open(file_path)
                self.current_page_index = 0
                self.display_page(self.current_page_index)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF file: {e}")

    def open_pdf_from_base64(self, base64data):
        try:
            pdf_data = base64.b64decode(base64data)
            self.doc = fitz.open("pdf", pdf_data)
            self.current_page_index = 0
            self.display_page(self.current_page_index)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF from base64 data: {e}")

    def open_pdf_from_sql(self, SUsername, SPassword, Title , Server, SDatabase):
        try:
            # Connect to the database using parameterized query
            conn = pyodbc.connect('Driver={SQL Server};'
                                  'Server='+ Server +';'
                                'Database='+ SDatabase +';'
                                'UID=' + SUsername + ';'
                                'PWD=' + SPassword + ';')

            # Create a cursor object
            cursor = conn.cursor()

            # Execute a SELECT query to retrieve the file data and name
            cursor.execute("SELECT file_stream, name FROM Library_FileTable WHERE name = ?", (Title,))
            row = cursor.fetchone()

            # Check if a row was returned
            if row:
                file_data, filename = row

                # Open PDF from file_stream data
                pdf_document = fitz.open(stream=file_data)

                # Assuming you have methods defined elsewhere like `display_page`
                self.doc = pdf_document
                self.current_page_index = 0
                self.display_page(self.current_page_index)
            else:
                messagebox.showerror("Error", f"No PDF found with title '{Title}'")

            # Close cursor and connection
            cursor.close()
            conn.close()

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Database error: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")

    def open_pdf_from_sql_wth_cn(self, cnSource , cnUser , cnCatalog, cnTitle, cnPassword, fileTables):
        try:
            # Connect to the database using parameterized query
            conn = pyodbc.connect('Driver={SQL Server};'
                                'Server='+ cnSource +';'
                                'Database='+ cnCatalog +';'
                                'UID=sa;' # For no they will remain like this until future patch'UID=' + SUsername + ';'
                                'PWD=Mmggit22;') #'PWD=' + SPassword + ';')

            # Create a cursor object
            cursor = conn.cursor()

            # Execute a SELECT query to retrieve the file data and name
            cursor.execute("SELECT file_stream, name FROM " + fileTables + " WHERE name = ?", (cnTitle,))
            
            row = cursor.fetchone()

            # Check if a row was returned
            if row:
                file_data, filename = row

                # Open PDF from file_stream data
                pdf_document = fitz.open(stream=file_data)

                # Assuming you have methods defined elsewhere like `display_page`
                self.doc = pdf_document
                self.current_page_index = 0
                self.display_page(self.current_page_index)
            else:
                messagebox.showerror("Error", f"No PDF found with title '{cnTitle}'")

            # Close cursor and connection
            cursor.close()
            conn.close()

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", f"Database error: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")


    # ----
    def display_page(self, page_index):
        if self.doc:
            try:
                page = self.doc.load_page(page_index)
                zoom_matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
                pix = page.get_pixmap(matrix=zoom_matrix)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Resize the image to fit the canvas while maintaining aspect ratio
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()
                img.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img)

                self.canvas.delete("all")  # Clear previous page

                # Center the image on the canvas
                img_width, img_height = img_tk.width(), img_tk.height()
                x = (canvas_width - img_width) // 2
                y = (canvas_height - img_height) // 2

                self.canvas.create_image(x, y, anchor=tk.NW, image=img_tk)
                self.canvas.image = img_tk  # Keep a reference to avoid garbage collection
            except Exception as e:
                messagebox.showerror("Error", f"Failed to display page: {e}")

    def on_resize(self, event):
        if self.doc:
            self.display_page(self.current_page_index)

    def on_mouse_wheel(self, event):
        if event.delta > 0:
            self.zoom_level *= 1.1  # Zoom in
        elif event.delta < 0:
            self.zoom_level /= 1.1  # Zoom out
        self.display_page(self.current_page_index)

    def next_page(self):
        if self.doc and self.current_page_index < len(self.doc) - 1:
            self.current_page_index += 1
            self.display_page(self.current_page_index)

    def prev_page(self):
        if self.doc and self.current_page_index > 0:
            self.current_page_index -= 1
            self.display_page(self.current_page_index)


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
    updater_path = os.path.join(os.path.dirname(__file__), "spvupdater.exe")
    if os.path.exists(updater_path):
        shell32 = ctypes.windll.shell32
        shell32.ShellExecuteW(None, "runas", updater_path, None, None, 1)
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
    # Hardcoded commit hash for Edition management
    production_commit_hash = "c87d779bd5acd4c10ca95a2bb20ed81068fd1b66"
    if not check_current_version(production_commit_hash):  # Check current version before starting the application
        choice = messagebox.askyesno("Update Required", "A new version of the application is available. Do you want to update?")
        if choice:
            launch_updater()

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
        root = tk.Tk()
        app = PDFViewer(root, file_path=file_path, base64data=base64data, SUsername=SUsername, SPassword=SPassword , Title=Title, Server=Server , SDatabase=SDatabase, cnSource=cnSource, cnUser=cnUser, cnCatalog= cnCatalog  , cnTitle=cnTitle, cnPassword=cnPassword, fileTables= fileTables)
        root.mainloop()

if __name__ == "__main__":
    main()