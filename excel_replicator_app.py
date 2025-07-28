import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import load_workbook
import pandas as pd
from PIL import Image, ImageTk
from file_operations import FileOperations
from data_operations import DataOperations
from image_operations import ImageOperations

class ExcelReplicatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Replicator")
        self.root.state('zoomed')  # Start maximized

        # Head Label
        heading_label = ttk.Label(root, text="Excel Replicator", font=('Arial', 24, 'bold'))
        heading_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)

        # Image and Label
        try:
            img = Image.open("Sunset.jpg")
            img = img.resize((550, 400), Image.LANCZOS)  # Updated to use LANCZOS resampling
            img = ImageTk.PhotoImage(img)
            img_label = tk.Label(root, image=img)
            img_label.image = img
            img_label.place(relx=0.5, rely=0.35, anchor=tk.CENTER)

            # Additional label next to the image on the right side (italicized)
            additional_label = ttk.Label(root, text=" Crafted with Love by\nSam Naveenkumar .V❤️", font=('cambria', 13, 'italic'))
            additional_label.place(relx=0.8, rely=0.4, anchor=tk.CENTER)

        except FileNotFoundError:
            messagebox.showerror("Error", "Image file 'Sunset.jpg' not found.")

        # Eye icon for source file
        try:
            eye_img = Image.open("eye_icon.png")
            eye_img = eye_img.resize((24, 24), Image.LANCZOS)  # Updated to use LANCZOS resampling
            self.eye_icon = ImageTk.PhotoImage(eye_img)
        except FileNotFoundError:
            self.eye_icon = None
            messagebox.showerror("Error", "Eye icon file 'eye_icon.png' not found.")

        # Source File Selection
        self.source_file_label = ttk.Label(root, text="Select Source Excel File:", font=('Arial', 12))
        self.source_file_label.place(relx=0.3, rely=0.7, anchor=tk.CENTER)
        self.source_file_entry = ttk.Entry(root, width=40, font=('Arial', 12))
        self.source_file_entry.place(relx=0.48, rely=0.7, anchor=tk.CENTER)
        self.source_file_view_btn = ttk.Button(root, image=self.eye_icon, command=self.view_source_file)
        self.source_file_view_btn.place(relx=0.62, rely=0.7, anchor=tk.CENTER)
        self.source_file_browse_btn = ttk.Button(root, text="Browse", command=self.browse_source_file, width=10)
        self.source_file_browse_btn.place(relx=0.67, rely=0.7, anchor=tk.CENTER)

        # Target File Selection
        self.target_file_label = ttk.Label(root, text="Select Target Excel File:", font=('Arial', 12))
        self.target_file_label.place(relx=0.3, rely=0.8, anchor=tk.CENTER)
        self.target_file_entry = ttk.Entry(root, width=40, font=('Arial', 12))
        self.target_file_entry.place(relx=0.48, rely=0.8, anchor=tk.CENTER)
        self.target_file_view_btn = ttk.Button(root, image=self.eye_icon, command=self.view_target_file)
        self.target_file_view_btn.place(relx=0.62, rely=0.8, anchor=tk.CENTER)
        self.target_file_browse_btn = ttk.Button(root, text="Browse", command=self.browse_target_file, width=10)
        self.target_file_browse_btn.place(relx=0.67, rely=0.8, anchor=tk.CENTER)

        # Source Sheet Selection
        self.source_sheet_label = ttk.Label(root, text="Source Sheet:", font=('Arial', 12))
        self.source_sheet_label.place(relx=0.3, rely=0.75, anchor=tk.CENTER)
        self.source_sheet_combobox = ttk.Combobox(root, width=30, font=('Arial', 12))
        self.source_sheet_combobox.place(relx=0.47, rely=0.75, anchor=tk.CENTER)

        # Target Sheet Selection
        self.target_sheet_label = ttk.Label(root, text="Target Sheet:", font=('Arial', 12))
        self.target_sheet_label.place(relx=0.3, rely=0.85, anchor=tk.CENTER)
        self.target_sheet_combobox = ttk.Combobox(root, width=30, font=('Arial', 12))
        self.target_sheet_combobox.place(relx=0.47, rely=0.85, anchor=tk.CENTER)

        # Option to copy entire file
        self.entire_file_var = tk.BooleanVar()
        self.entire_file_checkbutton = ttk.Checkbutton(root, text="Entire File", variable=self.entire_file_var)
        self.entire_file_checkbutton.place(relx=0.3, rely=0.92, anchor=tk.CENTER)

        # Push and Pull Buttons
        self.pull_btn = ttk.Button(root, text="Pull", command=self.pull_data)
        self.pull_btn.place(relx=0.4, rely=0.92, anchor=tk.CENTER)
        self.push_btn = ttk.Button(root, text="Push", command=self.replicate_data)
        self.push_btn.place(relx=0.6, rely=0.92, anchor=tk.CENTER)

        # Footer Labels
        left_footer_label = ttk.Label(root, text="Crafted with Love by Sam Naveenkumar .V ❤️",
                                      font=('Arial', 10), foreground='gray')
        left_footer_label.place(relx=0.02, rely=0.98, anchor=tk.SW)

        right_footer_label = ttk.Label(root, text="© 2024 All rights reserved.",
                                       font=('Arial', 10), foreground='gray')
        right_footer_label.place(relx=0.98, rely=0.98, anchor=tk.SE)

        # Initialize operations modules
        self.file_operations = FileOperations()
        self.data_operations = DataOperations()
        self.image_operations = ImageOperations()

    def browse_source_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.source_file_entry.delete(0, tk.END)
            self.source_file_entry.insert(0, file_path)
            self.load_sheet_names(file_path, self.source_sheet_combobox)

    def browse_target_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.target_file_entry.delete(0, tk.END)
            self.target_file_entry.insert(0, file_path)
            self.load_sheet_names(file_path, self.target_sheet_combobox)

    def load_sheet_names(self, file_path, combobox):
        self.file_operations.load_sheet_names(file_path, combobox)

    def view_source_file(self):
        filename = self.source_file_entry.get()
        self.file_operations.view_file(filename, self.source_sheet_combobox)

    def view_target_file(self):
        filename = self.target_file_entry.get()
        self.file_operations.view_file(filename, self.target_sheet_combobox)

    def pull_data(self):
        source_file = self.source_file_entry.get()
        target_file = self.target_file_entry.get()

        self.data_operations.pull_data(source_file, target_file, self.entire_file_var.get(), self.source_sheet_combobox, self.target_sheet_combobox)

    def replicate_data(self):
        source_file = self.source_file_entry.get()
        target_file = self.target_file_entry.get()

        self.data_operations.replicate_data(source_file, target_file, self.entire_file_var.get(), self.source_sheet_combobox, self.target_sheet_combobox)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReplicatorApp(root)
    root.mainloop()
