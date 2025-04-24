from docx2pdf import convert
from PyPDF2 import PdfMerger
import os
from services.zipOperations import ZipOperations
class PdfOperations:
    def __init__(self,output_directory):
        self.__output_directory = output_directory
        
        self.generate_pdf()
        self.merge_pdf()

    def generate_pdf(self):

        for i in range(0, 6):
            input = f"output/lvl{i}/lvl{i}doc"
            output = f"output/lvl{i}/lvl{i}pdf"
            convert(input, output)
    
    def merge_pdf(self):

        for i in range(0, 6):
            input = f"output/lvl{i}/lvl{i}pdf"
            output = f"output/merged/lvl{i}merged.pdf"

            pdf_files = [f for f in os.listdir(input) if f.endswith(".pdf")]
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(os.path.join(input, pdf))
            merger.write(output)
            merger.close()
        ZipOperations(self.__output_directory)