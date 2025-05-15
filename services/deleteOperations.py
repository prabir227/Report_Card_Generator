import os
from pathlib import Path
class DeleteOperations:
    def __init__(self):
        self.delete_files()

    def delete_files(self):
        file_path = Path(__file__).parent.parent
        #print(f"file path is{file_path}")
        print("testing")

if __name__ == "__main__":
    DeleteOperations()