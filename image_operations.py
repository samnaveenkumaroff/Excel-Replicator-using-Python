from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox

class ImageOperations:
    def __init__(self):
        pass

    def load_image(self, image_path, root):
        try:
            img = Image.open(image_path)
            img = img.resize((550, 400), Image.LANCZOS)  # Updated to use LANCZOS resampling
            img = ImageTk.PhotoImage(img)
            img_label = tk.Label(root, image=img)
            img_label.image = img
            img_label.place(relx=0.5, rely=0.4, anchor=tk.CENTER)

            # Additional label next to the image on the right side (italicized)
            additional_label = tk.Label(root, text=" Crafted with Love by\nSam Naveenkumar .V❤️", font=('cambria', 13, 'italic'))
            additional_label.place(relx=0.8, rely=0.4, anchor=tk.CENTER)

        except FileNotFoundError:
            messagebox.showerror("Error", "Image file not found.")
