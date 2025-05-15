import tkinter as tk
import os
import shutil
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox
from views.inputFormPage import InputFormPage
class UploadPage(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        self.root.title("Upload Page")
        self.root.configure(bg="#FFDD61")
        self.root.minsize(600, 350)  # Set minimum size to prevent resizing too small
        self.root.maxsize(600, 350)  # Set maximum size to prevent resizing too large
        self.__output_directory = None  # Initialize output directory
        # === Header Frame (Holds Logo + Text Side by Side) ===
        header_frame = tk.Frame(self.root, bg="#FFDD61")
        header_frame.pack(side="top", fill="x", padx=20, pady=20)
        self.__uploaded_file = False
        # === Logo on the Left ===
        logo = Image.open("assets/hhfc-logo.png")
        logo = logo.resize((110, 80))
        self.image = ImageTk.PhotoImage(logo)

        logo_label = tk.Label(header_frame, image=self.image, bg="#FFDD61")
        logo_label.pack(side="left", padx=8)
        logo_label.image = self.image  # Keep a reference to avoid garbage collection

        # === Text Frame on the Right ===
        text_frame = tk.Frame(header_frame, bg="#FFDD61")
        text_frame.pack(side="left", padx=20, anchor="n")

        heading = tk.Label(text_frame, text="REPORT CARD GENERATOR", font=("Arial", 20, "bold"), bg="#FFDD61", fg="#333")
        heading.pack(anchor="w")

        subheading = tk.Label(text_frame, text="Helping Hands Are Better Than Praying Lips", font=("Arial", 14, "italic"), bg="#FFDD61", fg="#555")
        subheading.pack(anchor="w", pady=(5, 0))

        # === Upload Excel Button ===
        upload_button = tk.Button(self.root, text="Upload Student Excel File", command=self.browse_file, bg="#fff", font=("Arial", 12))
        upload_button.pack(pady=20)

        tk.Button(self.root,text = "Select Output Directory",command = self.browse_directory,bg="#fff",font=("Arial",12)).pack(pady=20)
         # Footer Frame
        footer = tk.Frame(self.root, bg="#FFDD61")
        footer.pack(side="bottom", fill="x", pady=10)

        # Centered Footer Text
        footer_label = tk.Label(
            footer,
            text="                     Made for Teachers by HHFC",
            font=("Arial", 10),
            bg="#FFDD61"
        )
        footer_label.pack(side="left",expand=True)

        # Logo on the Right
        logo = Image.open("assets/hhfc-logo.png")
        logo = logo.resize((50, 30))  # Smaller logo for footer
        self.footer_logo = ImageTk.PhotoImage(logo)
        logo_label = tk.Label(footer, image=self.footer_logo, bg="#FFDD61")
        logo_label.pack(side="right", padx=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if file_path:
            # Your desired filename
            new_file_name = f"uploaded.xlsx"  # or .xls if original was .xls

            # Define upload directory
            upload_dir = "uploads"
            os.makedirs(upload_dir, exist_ok=True)

            # Final destination path with new name
            destination_path = os.path.join(upload_dir, new_file_name)

            try:
                shutil.copy(file_path, destination_path)  # ✅ Copy and rename
                self.selected_file = destination_path
                messagebox.showinfo("File Uploaded", "Your file has been uploaded successfully!")
                print("Uploaded and renamed file:", destination_path)
                self.__uploaded_file = True    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to upload and rename file: {str(e)}")
        else:
            self.selected_file = None
            messagebox.showwarning("No File", "No file was selected.")


    def browse_directory(self):
        folder_path = filedialog.askdirectory(title="Select Output Directory")
        if folder_path:
            # Save the selected directory path to a variable or use it as needed
            self.__output_directory = folder_path
            messagebox.showinfo("Directory Selected", f"Selected directory: {folder_path}")
            if self.__uploaded_file:
                for widget in self.root.winfo_children():
                    widget.destroy()
                InputFormPage(self.root,self.__output_directory)
        else:
            self.__output_directory = None
            messagebox.showwarning("No Directory", "No directory was selected.")