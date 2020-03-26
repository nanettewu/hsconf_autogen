import argparse
import datetime
import glob
import io
import os

# converting excel > csv
import csv
import xlrd
import sys

# writing csv data > pdf
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import *

# constants
AVAILABLE_TIMES = ['9AM.csv', '10AM.csv', '11AM.csv', '12PM.csv', '1PM.csv']

# styling
title_style = ParagraphStyle(
		'title',
		fontName="Helvetica-bold",
		fontSize=20,
		textColor=HexColor(0x444846),
)

header_style = ParagraphStyle(
	'header',
	fontName="Helvetica-bold",
	fontSize=12,
	textColor=HexColor(0x94645f),
)

table_style = [
	('ALIGN',(0,0),(0,-1),'LEFT'),
	('ALIGN',(-1,0),(-1,-1),'RIGHT'),
	('TEXTCOLOR', (0,0), (-1, -1), (0.29, 0.29, 0.29)),
]

'''
Given `excel_file`, the name of an Excel spreadsheet, this function
converts each worksheet within the spreadsheet to its own CSV file.
The spreadsheet must be at the top level of `hsconf_autogen`, and the
CSV files are written out to the `csv/` directory.
'''
def convert_excel_to_csv(excel_file):
	# clean up: if csv folder has old csv files, delete them
	files = glob.glob('csv/*')
	for f in files:
		os.remove(f)

	print(f"Processing {excel_file}...")
	with xlrd.open_workbook(excel_file) as wb:

		# process one sheet at a time
		for sheet_name in wb.sheet_names():
			print(f"> Opening sheet '{sheet_name}'")
			sheet = wb.sheet_by_name(sheet_name)
			csv_file_path = f"csv/{sheet_name}.csv"
			
			# write each row of worksheet to CSV
			with open(csv_file_path, 'w') as csvfile:
				c = csv.writer(csvfile)
				for r in range(sheet.nrows):
					c.writerow(sheet.row_values(r))
			
			print(f"\tConverted to {csv_file_path}")

'''
This function creates the high school conference PDF (target_file) by
taking the CSV data extracted from the spreadsheet, and writing it on
top of a template file (thursday.pdf or tuesday.pdf).  
'''
def write_csv_data(template_file, target_file):
	print(f"\nProcessing csv data for {target_file}...")

	# open the template used to write data on top of
	existing_pdf = PdfFileReader(open(f"template/{template_file}", "rb"))
	output = PdfFileWriter()

	# write title page
	title_page = _write_title_data(existing_pdf)
	output.addPage(title_page)

	# in excel_files, store the names of the CSV files that were
	# extracted from the excel sheet. this will help us figure out
	# which pages/time slots should be included in the program
	excel_files = []
	for file in os.listdir("csv/"):
		if file.endswith(".csv"):
			excel_files.append(file)

	# iterate through each sheet/time slot that was extracted from
	# the excel sheet
	counter = 1
	for time in AVAILABLE_TIMES:
		if time in excel_files:
			print(f"> Writing page {counter+1} from {time}")
			_write_schedule_page(time)
			
			# add "watermark" (new PDF) on existing page
			new_pdf = PdfFileReader(open("schedule_page.pdf", "rb"))
			page = existing_pdf.getPage(counter)
			page.mergePage(new_pdf.getPage(0))
			output.addPage(page)

			counter += 1
			print("\tDone!")

	# write all output
	with open(target_file, "wb") as output_stream:
		output.write(output_stream)

'''
HELPER FUNCTIONS
'''

# generates title page for the program
def _write_title_data(existing_pdf):
	print("> Writing title page from meta.csv")
	
	# create canvas to draw title data on
	packet = io.BytesIO()
	can = canvas.Canvas(packet, pagesize=letter)

	# read metadata file (contains date/time & special thanks info)
	with open('csv/meta.csv', newline='') as csvfile:
		reader = csv.reader(csvfile)
		date_and_time = next(reader)
		date = date_and_time[0].upper().strip().replace("\"", "")
		special_thanks = next(reader)
		
		# use datetime package to get current timestamp
		time = datetime.datetime.now().strftime("%b %d %H:%M")
		last_updated = f"Last updated: {time}"

		# write date and time (i.e. APRIL 5, 2018 • 9AM-2PM)
		can.setFillColor(HexColor(0x666a69))
		can.setFont("Helvetica", 16)
		can.drawString(73, 355, f"{date}  •  {date_and_time[1]}")

		# write special thanks people
		can.setFillColor(HexColor(0x7d8281))
		can.setFont("Helvetica", 11)
		x_pos = 540
		y_pos = 174
		for person in special_thanks:
			can.drawRightString(x_pos, y_pos, person)
			y_pos -= 17

		# write current timestamp for last updated information
		can.setFillColor(HexColor(0xdadde5))
		can.setFont("Helvetica", 12)
		can.drawString(392, 50, last_updated)
		can.save()

	# source: https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python
	packet.seek(0)
	title_pdf = PdfFileReader(packet)
	page = existing_pdf.getPage(0)
	page.mergePage(title_pdf.getPage(0))

	print("\tDone!")
	return page

# generates a schedule page (i.e. 9AM page with speakers & talk titles)
def _write_schedule_page(filename):
	with open(f"csv/{filename}", "r") as csvfile:
		data = list(csv.reader(csvfile))

	elements = []

	# add title (i.e. 9AM) to top of page
	title = filename.split('.')[0] # i.e. extracts "9AM" from "9AM.csv"
	elements.append(Paragraph(title, title_style))
	elements.append(Spacer(inch, .35 * inch))

	# add first (location • host) line to page
	first_header = ' • '.join(data[0])
	elements.append(Paragraph(first_header, header_style))
	elements.append(Spacer(inch, .1 * inch))

	# iterate through each line in the csv file
	idx = 1
	room_of_speakers = []
	while idx < len(data):
		entry = data[idx]
		if not entry[0]: # empty line, then skip to next room of speakers			
			elements = _add_table_to_doc(elements, room_of_speakers)
			elements.append(Spacer(inch, .15 * inch))
			room_of_speakers = []
			idx += 1
			header = ' • '.join(data[idx])
			elements.append(Paragraph(header,header_style))
			elements.append(Spacer(inch, .08 * inch))
		else:
			room_of_speakers += [entry]
		idx += 1

	# add last set of speakers to page
	if room_of_speakers:
		elements = _add_table_to_doc(elements, room_of_speakers)

	# generate PDF and write all contents to PDF
	test_pdf = SimpleDocTemplate(
		'schedule_page.pdf',
		pagesize=letter,
		rightMargin=40,
		leftMargin=40,
		topMargin=40,
		bottomMargin=28)
	test_pdf.build(elements)

# creates a reportlab.platypus Table object and adds it to the page
def _add_table_to_doc(elements, data):
	t = Table(data, colWidths=[6.14*inch, 1.25*inch], rowHeights=0.22*inch, style=table_style)
	elements.append(t)
	return elements

'''
MAIN
'''

def main(excel_file, day):
	# extract data from the excel spreadsheet
	convert_excel_to_csv(excel_file)

	# use the extracted data to write the program onto a PDF
	if day == "thursday":
		write_csv_data("thursday_template.pdf", f"thursday_program.pdf")
	else: # tuesday
		write_csv_data("tuesday_template.pdf", f"tuesday_program.pdf")

	# cleanup
	os.remove("schedule_page.pdf")


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Generate 6.UAT HS Conference Program.')
	parser.add_argument("-f", "--file", help="Excel spreadsheet file name (i.e. thursday.xlsx)")
	parser.add_argument("-d", "--day", help="Tuesday or Thursday template")

	args = parser.parse_args()
	excel_file = str(args.file) if args.file else None
	day = str(args.day).lower() if args.day and str(args.day).lower() in ["thursday", "tuesday"] else None

	if not day:
		print("ERROR: need to specify date of HS conference (Tuesday or Thursday) with -d flag.")
	elif (not excel_file) or (excel_file and not excel_file.endswith('xlsx')):
		print("ERROR: need to specify valid Microsoft Excel spreadsheet file (i.e. thursday.xlsx)")
	elif not os.path.exists(excel_file):
		print(f"ERROR: {excel_file} does not exist in the top level directory.")
	else:
		main(excel_file, day)