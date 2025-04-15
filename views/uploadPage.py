import tkinter as tk
from PIL import Image, ImageTk
class UploadPage:
    def __init__(self,root):
        self.root = root
        self.root.title("Upload Page")
        self.root.geometry("400x300")
        self.root.configure(bg="#FFDD61")
        logo = Image.open("assets/hhfc-logo.png")
        logo = logo.resize((150, 100))
        self.image = ImageTk.PhotoImage(logo)
        label = tk.Label(self.root,image=self.image,bg="#FFDD61")
        label.image = self.image
        label.pack(anchor = "nw")
        heading = tk.Label(self.root, text="Upload Page", font=("Arial", 24), bg="#FFDD61")
        heading.pack(anchor="ne")
        # self.root.pack()

