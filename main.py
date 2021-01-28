# Libraries that we need to intall - PyPDF2, pdfminer
pip install PyPDF2
pip install pdfminer

# Import the nessesary packages
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
import os, PyPDF2, re
import datetime

# 1st function is for extracting individual pages of tax slips from pdf
def split_pdf_pages(root_directory, extract_folder):
	for root, dirs, files in os.walk(root_directory):
		for filename in files:
			basename, extension = os.path.splitext(filename)
                        # Choosing each pdf files
			if extension == ".pdf":
                                # Create file path name for pdf formatted files
				path = root + "\\" + basename + extension
                                # Open the files with PdfFileReader
				open_pdf = PyPDF2.PdfFileReader(open(path, "rb"))
                                #Loop through the file and split the pages
				for i in range(open_pdf.numPages):
					output = PyPDF2.PdfFileWriter()
					output.addPage(open_pdf.getPage(i))
					with open(extract_to+ "\\" + basename + "-%s.pdf" % i, "wb") as output_pdf:
						output.write(output_pdf)

# 2nd function is for renaming each extracted tax invoice with Customer name, Invoice number, Respective month and Current year
def rename_pdfs(root_directory, extract_folder):
	for root, dirs, files in os.walk(root_directory):
		for filename in files:
			basename, extension = os.path.splitext(filename)
			if extension == ".pdf":
				path = root + "\\" + basename + extension
				output_string = StringIO()
				pdf_file = open(path, "rb")
				parser = PDFParser(pdf_file)
				doc = PDFDocument(parser)
				rsrcmgr = PDFResourceManager()
				device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
				interpreter = PDFPageInterpreter(rsrcmgr, device)
				for page in PDFPage.create_pages(doc):
					interpreter.process_page(page)
				num = output_string.getvalue()
				for cust_name in re.findall("CV.\s[A-Za-z]+",num) or re.findall("PT.\s+[A-Za-z]+\s+[A-Za-z]+\s+[A-Za-z]+", num) or re.findall("PT\s+[A-Za-z]+\s+[A-Za-z]+\s+[A-Za-z]+", num):
					cust = cust_name[0:50]
				for inv in re.findall("#[0-9]+", num):
					inv_num = inv[1:]
				for month in re.findall("January", num) or re.findall("February", num) or re.findall("March",
																									 num) or re.findall(
					"April", num) or re.findall("May", num) or re.findall("June", num) or re.findall("July",
																									 num) or re.findall(
					"August", num) or re.findall("September", num) or re.findall("October", num) or re.findall(
					"November", num) or re.findall("December", num):
					date = month
					year=datetime.date.today().year
				pdf_file.close()
				os.rename(path, rename_to + "//" + f"{cust}  { inv_num} { date}  { year} .pdf")

# Folder paths
root_dir = r"C:\Users\*****\******\pdfSplitRename1\initial"
extract_to = r"C:\Users\*****\******\pdfSplitRename1\extract"
rename_to = r"C:\Users\*****\******\pdfSplitRename1\final"

# Run the functions
split_pdf_pages(root_dir, extract_to)
rename_pdfs(extract_to,rename_to)
