import tkinter as tk
from views.uploadPage import UploadPage
class ReportCardGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.upload_page=UploadPage(self)


# def run_app():
#     root = tk.Tk()
#     UploadPage(root)
#     root.mainloop()

if __name__ == "__main__":
    app = ReportCardGeneratorApp()
    app.mainloop()