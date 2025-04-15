import tkinter as tk
from views.uploadPage import UploadPage
def run_app():
    root = tk.Tk()
    UploadPage(root)
    root.mainloop()

if __name__ == "__main__":
    run_app()