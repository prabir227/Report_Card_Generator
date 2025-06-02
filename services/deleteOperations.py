import os
import shutil
class DeleteOperations:
    def __init__(self):
        self.delete_files()

    def delete_files(self):
        base_dir = "output"
        upload_dir ="uploads"
        # Check if the 'uploads' folder exists
        if os.path.exists(upload_dir):
            shutil.rmtree(upload_dir)
            print("All folders deleted successfully.")
        else:
            print("The 'uploads' folder does not exist.")
        # Check if the 'output' folder exists
        if os.path.exists(base_dir):
            shutil.rmtree(base_dir)
            print("All folders deleted successfully.")
        else:
            print("The 'output' folder does not exist.")

if __name__ == "__main__":
    DeleteOperations()