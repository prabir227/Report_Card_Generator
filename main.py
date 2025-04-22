import tkinter as tk
from views.uploadPage import UploadPage
class ReportCardGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.upload_page=UploadPage(self)


if __name__ == "__main__":
    app = ReportCardGeneratorApp()
    app.mainloop()