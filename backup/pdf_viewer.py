import fitz  # PyMuPDF
import sys
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk

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

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF Files", "*.pdf")]
        )
    if file_path:
        root = tk.Tk()
        app = PDFViewer(root, file_path)
        root.mainloop()
