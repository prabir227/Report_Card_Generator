import tkinter as tk
from PIL import Image, ImageTk
from tkinter import messagebox
from services.excelOperations import ExcelOperations
from services.pdfOperations import PdfOperations
class InputFormPage(tk.Frame):
    def __init__(self,root,output_directory):
        super().__init__(root)
        self.root = root
        self.root.title("Input Form Page")
        self.root.configure(bg="#FFDD61")
        self.root.minsize(700, 650)
        self.root.maxsize(700, 650)

        self.__output_directory = output_directory
        self.__eng1 = tk.StringVar()
        self.__eng2 = tk.StringVar()
        self.__eng3 = tk.StringVar()
        self.__math1 = tk.StringVar()
        self.__math2 = tk.StringVar()
        self.__math3 = tk.StringVar()
        self.__evs1 = tk.StringVar()
        self.__evs2 = tk.StringVar()
        self.__evs3 = tk.StringVar()
        self.__hindi1 = tk.StringVar()
        self.__hindi2 = tk.StringVar()
        self.__hindi3 = tk.StringVar()
        self.__sst1 = tk.StringVar()
        self.__sst2 = tk.StringVar()
        self.__sst3 = tk.StringVar()
        self.__level0start = tk.StringVar()
        self.__level0end = tk.StringVar()
        self.__level1start = tk.StringVar()
        self.__level1end = tk.StringVar()
        self.__level2start = tk.StringVar()
        self.__level2end = tk.StringVar()
        self.__level3start = tk.StringVar()
        self.__level3end = tk.StringVar()
        self.__level4start = tk.StringVar()
        self.__level4end = tk.StringVar()
        self.__level5start = tk.StringVar()
        self.__level5end = tk.StringVar()
        self.__level0total = tk.StringVar()
        self.__level1total = tk.StringVar()
        self.__level2total = tk.StringVar()
        self.__level3total = tk.StringVar()
        self.__level4total = tk.StringVar()
        self.__level5total = tk.StringVar()
        self.__studyCentre = tk.StringVar()
        self.__date = tk.StringVar()
        self.__name = tk.StringVar()
        font_bold = ("Helvetica", 12, "bold")
        header_font = ("Helvetica", 20, "bold")
        entry_font = ("Arial", 12, "bold")
        bg_color = "#FFDD61"
        entry_bg = "#ffffff"
        entry_width = 8
        padx_label = (100, 0)
        padx_entry = (10, 10)
        pady_entry = (5, 5)

        tk.Label(self.root, text="Enter the row and column values", font=header_font, bg=bg_color)\
            .grid(row=0, column=0, columnspan=10, pady=(10, 20))

        # Column Section
        tk.Label(self.root, text="Column values", font=font_bold, bg=bg_color)\
            .grid(row=1, column=0, sticky="w", padx=padx_label)
        for i, text in enumerate(["Quarterly 1", "Quarterly 2", "Quarterly 3"]):
            tk.Label(self.root, text=text, font=font_bold, bg=bg_color)\
                .grid(row=1, column=i+1, padx=padx_entry, pady=pady_entry)

        subjects = ["Eng", "Math", "Evs", "Hindi", "SST"]
        for i, subject in enumerate(subjects):
            tk.Label(self.root, text=subject, font=font_bold, bg=bg_color)\
                .grid(row=i+2, column=0, sticky="w", padx=(150, 0), pady=pady_entry)

            for j in range(3):
                entry_var = getattr(self, f"_{self.__class__.__name__}__{subject.lower()}{j+1}")
                tk.Entry(self.root, font=entry_font, bg=entry_bg, width=entry_width, textvariable=entry_var)\
                    .grid(row=i+2, column=j+1, padx=padx_entry, pady=pady_entry)



        # Row Section
        tk.Label(self.root, text="Row values", font=font_bold, bg=bg_color)\
            .grid(row=7, column=0, sticky="w", padx=padx_label, pady=(10, 5))
        for i, label in enumerate(["Start", "End", "Total"]):
            tk.Label(self.root, text=label, font=font_bold, bg=bg_color)\
                .grid(row=7, column=i+1, padx=padx_entry, pady=pady_entry)

        for level in range(6):
            tk.Label(self.root, text=f"Level {level}", font=font_bold, bg=bg_color)\
                .grid(row=level+8, column=0, sticky="w", padx=(130, 0), pady=pady_entry)

            for i, suffix in enumerate(["start", "end", "total"]):
                entry_var = getattr(self, f"_{self.__class__.__name__}__level{level}{suffix}")
                tk.Entry(self.root, font=entry_font, bg=entry_bg, width=entry_width, textvariable=entry_var)\
                    .grid(row=level+8, column=i+1, padx=padx_entry, pady=pady_entry)

        #Date & Study Centre
        tk.Label(self.root, text="Date :", font=font_bold, bg=bg_color)\
            .grid(row=14, column=0, sticky="w", padx=(130,0), pady=pady_entry)
        tk.Entry(self.root, font=entry_font, bg=entry_bg, width=10, textvariable=self.__date)\
            .grid(row=14, column=1, padx=padx_entry, pady=pady_entry)

        tk.Label(self.root, text="Study Centre :", font=font_bold, bg=bg_color)\
            .grid(row=15, column=0, sticky="w", padx=(130,0), pady=pady_entry)
        tk.Entry(self.root, font=entry_font, bg=entry_bg, width=10, textvariable=self.__studyCentre)\
            .grid(row=15, column=1, padx=padx_entry, pady=pady_entry)
        tk.Label(self.root, text="Name Column:", font=font_bold, bg=bg_color).grid(row=16, column=0, sticky="e", padx=padx_label, pady=pady_entry)
        tk.Entry(self.root, font=entry_font, bg=entry_bg, width=10, textvariable=self.__name)\
            .grid(row=16, column=1, padx=padx_entry, pady=pady_entry)
        #logo
        logo = Image.open("assets/hhfc-logo.png")
        logo = logo.resize((80, 50))
        self.image = ImageTk.PhotoImage(logo)
        logo_label = tk.Label(self.root, image=self.image, bg=bg_color).grid(row=14, column=5, sticky="se", padx=(0, 0), pady=(0, 15),columnspan=2,rowspan=2)
        

        # Submit Button
        tk.Button(self.root, text="Submit", font=font_bold, bg="#FF9800", fg="white",command=self.submit)\
            .grid(row=14, column=5, columnspan=2, pady=(10, 10),rowspan=4,sticky="se",padx=(20,10))
    
    def check(self):
        fields = [
            self.__eng1, self.__eng2, self.__eng3,
            self.__math1, self.__math2, self.__math3,
            self.__evs1, self.__evs2, self.__evs3,
            self.__hindi1, self.__hindi2, self.__hindi3,
            self.__sst1, self.__sst2, self.__sst3,
            self.__level0start, self.__level0end,
            self.__level1start, self.__level1end,
            self.__level2start, self.__level2end,
            self.__level3start, self.__level3end,
            self.__level4start, self.__level4end,
            self.__level5start, self.__level5end,
            self.__level0total, self.__level1total,
            self.__level2total, self.__level3total,
            self.__level4total, self.__level5total,
            self.__studyCentre, self.__date
        ]

        for var in fields:
            if var.get().strip() == '':
                messagebox.showerror("Missing Information", "Please fill all the required fields.")
                return False
        
        return True
    def submit(self):
        if self.check():
            ExcelOperations(
            int(self.__level0start.get()), int(self.__level0end.get()),
            int(self.__level1start.get()), int(self.__level1end.get()),
            int(self.__level2start.get()), int(self.__level2end.get()),
            int(self.__level3start.get()), int(self.__level3end.get()),
            int(self.__level4start.get()), int(self.__level4end.get()),
            int(self.__level5start.get()), int(self.__level5end.get()),
            self.__date.get(),
            self.__eng1.get(), self.__eng2.get(), self.__eng3.get(),
            self.__hindi1.get(), self.__hindi2.get(), self.__hindi3.get(),
            self.__math1.get(), self.__math2.get(), self.__math3.get(),
            self.__sst1.get(), self.__sst2.get(), self.__sst3.get(),
            self.__evs1.get(), self.__evs2.get(), self.__evs3.get(),
            self.__name.get(),
            int(self.__level0total.get()), int(self.__level1total.get()),
            int(self.__level2total.get()), int(self.__level3total.get()),
            int(self.__level4total.get()), int(self.__level5total.get()),
            self.__studyCentre.get())

            PdfOperations(self.__output_directory)

            messagebox.showinfo("Operations Successful",f"Your files have been generated successfully and saved in {self.__output_directory}")


        
        

        


if __name__ == "__main__":
    root = tk.Tk()
    InputFormPage(root)
    root.mainloop()