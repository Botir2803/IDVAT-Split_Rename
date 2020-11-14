import os, PyPDF2, re
import datetime

#1st function is for extracting individual pages of tax slips from pdf
def split_pdf_pages(root_directory, extract_folder):
	for root, dirs, files in os.walk(root_directory):
		for filename in files:
			basename, extension = os.path.splitext(filename)

			if extension == ".pdf":

				path = root + "\\" + basename + extension

				opened_pdf = PyPDF2.PdfFileReader(open(path, "rb"))

				for i in range(opened_pdf.numPages):
					output = PyPDF2.PdfFileWriter()
					output.addPage(opened_pdf.getPage(i))
					with open(extract_to+ "\\" + basename + "-%s.pdf" % i, "wb") as output_pdf:
						output.write(output_pdf)

#2nd function is for renaming each tax slip pdf with Customer name, Invoice number, Respective month and Current year
def rename_pdfs(root_directory, extract_folder):
	for root, dirs, files in os.walk(root_directory):
		for filename in files:
			basename, extension = os.path.splitext(filename)
			if extension == ".pdf":
				path = root + "\\" + basename + extension
				from io import StringIO

				from pdfminer.converter import TextConverter
				from pdfminer.layout import LAParams
				from pdfminer.pdfdocument import PDFDocument
				from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
				from pdfminer.pdfpage import PDFPage
				from pdfminer.pdfparser import PDFParser

				output_string = StringIO()
				in_file = open(path, "rb")
				parser = PDFParser(in_file)
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
				for month in re.findall("Januari", num) or re.findall("Februari", num) or re.findall("Maret",
																									 num) or re.findall(
					"April", num) or re.findall("Mei", num) or re.findall("Juni", num) or re.findall("Juli",
																									 num) or re.findall(
					"Agustus", num) or re.findall("September", num) or re.findall("Oktober", num) or re.findall(
					"November", num) or re.findall("Desember", num):
					date = month
					year=datetime.date.today().year
				in_file.close()
				os.rename(path, rename_to + "//" + f"{cust}  { inv_num} { date}  { year} .pdf")

#folder paths from where we get the files
root_dir = r"C:\Users\*****\******\pdfSplitRename1\initial"
extract_to = r"C:\Users\*****\******\pdfSplitRename1\extract"
rename_to = r"C:\Users\*****\******\pdfSplitRename1\final"

#function run
split_pdf_pages(root_dir, extract_to)
rename_pdfs(extract_to,rename_to)
