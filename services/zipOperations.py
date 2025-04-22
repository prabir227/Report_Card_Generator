import shutil
import os
import zipfile
class ZipOperations:
    def __init__(self, zip_file_path):
        self.zip_file_path = zip_file_path
        # shutil.make_archive(self.zip_file_path, 'zip', 'my_folder')
        self.zip_output_folder()
    def zip_output_folder(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))  # services/
        output_dir = os.path.join(base_dir, '..', 'output')    # go up and into output
        output_dir = os.path.abspath(output_dir)               # full path

        zip_path = os.path.join(base_dir, '..', 'output.zip')  # zip saved in project root
        zip_path = os.path.abspath(zip_path)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_dir)  # keep folder structure
                    zipf.write(file_path, arcname)

if __name__ == "__main__":
    zip_operations = ZipOperations("..\output")