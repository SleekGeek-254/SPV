import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import sys
import subprocess
import os

class PDFViewer:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("SAMS PDF Viewer")

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

        # Open PDF file
        self.open_pdf(file_path)

    def open_pdf(self, file_path):
        if file_path:
            self.doc = fitz.open(file_path)
            self.current_page_index = 0
            self.display_page(self.current_page_index)

    def display_page(self, page_index):
        if self.doc:
            page = self.doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
            pix = page.get_pixmap(matrix=zoom_matrix)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Resize the image to fit the canvas while maintaining aspect ratio
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            img.thumbnail((canvas_width, canvas_height), Image.ANTIALIAS)
            img_tk = ImageTk.PhotoImage(img)

            self.canvas.delete("all")  # Clear previous page

            # Center the image on the canvas
            img_width, img_height = img_tk.width(), img_tk.height()
            x = (canvas_width - img_width) // 2
            y = (canvas_height - img_height) // 2

            self.canvas.create_image(x, y, anchor=tk.NW, image=img_tk)
            self.canvas.image = img_tk  # Keep a reference to avoid garbage collection

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

def check_current_version(commit_hash, branch='production'):
    try:
        # Run 'git ls-remote' to get the latest commit hash of the repository
        result = subprocess.run(['git', 'ls-remote', '--heads', 'https://github.com/SleekGeek-254/SPV.git', branch], capture_output=True, text=True)
        latest_commit_hash = result.stdout.split()[0]
        print("Latest commit hash for branch '{}':".format(branch), latest_commit_hash)

        # Check if the latest commit hash matches the hardcoded one
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
    updater_path = os.path.join(os.path.dirname(__file__), "updater.exe")
    if os.path.exists(updater_path):
        subprocess.Popen(updater_path)
    else:
        messagebox.showerror("Updater not found", "Updater executable not found.")

if __name__ == "__main__":
    # Hardcoded commit hash for production
    production_commit_hash = "c87d779bd5acd4c10ca95a2bb20ed81068fd1b6"
    
    if check_current_version(production_commit_hash):  # Check current version before starting the application
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
        else:
            file_path = filedialog.askopenfilename(
                title="Open PDF",
                filetypes=[("PDF Files", "*.pdf")]
            )
        root = tk.Tk()
        app = PDFViewer(root, file_path)
        root.mainloop()
    else:
        choice = messagebox.askyesno("Update Required", "A new version of the application is available. Do you want to update?")
        if choice:
            launch_updater()
