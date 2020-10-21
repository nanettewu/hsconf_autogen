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
	fontSize=11,
	textColor=HexColor(0x94645f),
)

table_cell_style = ParagraphStyle(
		'table_cell_style',
		fontName="Helvetica",
		fontSize=9,
		leftIndent=1,
		textColor=HexColor(0x787878),
)

table_style = [
	('ALIGN',(0,0),(0,-1),'LEFT'),
	('ALIGN',(-1,0),(-1,-1),'RIGHT'),
	('TEXTCOLOR', (0,0), (-1, -1), (0.29, 0.29, 0.29)),
	('FONTNAME', (0,0), (1, -1), 'Helvetica-Bold'),
	('FONTSIZE', (-1,0), (-1,-1), 10),
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
def write_csv_data(day, hide_loc):
	template_file = f"{day}_template.pdf"
	target_file = f"{day}_program.pdf"

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
			_write_schedule_page(day, time, hide_loc)
			
			# add "watermark" (new PDF) on existing page
			new_pdf = PdfFileReader(open("schedule_page.pdf", "rb"))
			for page_num in range(new_pdf.getNumPages()):
				page = existing_pdf.getPage(counter)
				page.mergePage(new_pdf.getPage(page_num))
				counter += 1
				output.addPage(page)
			
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
		# special_thanks = next(reader)
		
		# use datetime package to get current timestamp
		time = datetime.datetime.now().strftime("%b %d %H:%M")
		last_updated = f"Last updated: {time}"

		# write date and time (i.e. APRIL 5, 2018 • 9AM-2PM)
		can.setFillColor(HexColor(0x666a69))
		can.setFont("Helvetica", 16)
		can.drawString(73, 355, f"{date}  •  {date_and_time[1]}")

		# write special thanks people
		# can.setFillColor(HexColor(0x7d8281))
		# can.setFont("Helvetica", 11)
		# x_pos = 540
		# y_pos = 174
		# for person in special_thanks:
		# 	can.drawRightString(x_pos, y_pos, person)
		# 	y_pos -= 17

		# write current timestamp for last updated information
		can.setFillColor(HexColor(0xc7c7c7))
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
def _write_schedule_page(day, filename, hide_loc):
	formatted_day = "Thursday" if day == "thursday" else "Tuesday"

	with open(f"csv/{filename}", "r") as csvfile:
		data = list(csv.reader(csvfile))

	elements = []

	# add title (i.e. 9AM) to top of page
	title = filename.split('.')[0] # i.e. extracts "9AM" from "9AM.csv"
	elements.append(Paragraph(f"{title} • {formatted_day}", title_style))
	elements.append(Spacer(inch, .35 * inch))

	# add first (location • host) line to page
	zoom_url = data[0][0].strip()
	first_location = f"<u>{zoom_url}</u>"
	moderators = data[0][1]
	first_header = f"{first_location}  •  {moderators}" if not hide_loc else f"Zoom Link TBA  •  {moderators}"

	elements.append(Paragraph(first_header, header_style))
	elements.append(Spacer(inch, .07 * inch))

	# iterate through each line in the csv file
	time = title[:-2] # remove AM or PM
	num_rooms = 0
	line_idx = 1
	room_of_speakers = []
	while line_idx < len(data):
		entry = data[line_idx]
		if not entry[0]: # detected empty line = done processing a room of speakers
			# add current room of speakers to pdf, then reset to empty room		
			elements = _add_table_to_doc(time, elements, room_of_speakers)
			room_of_speakers = []

			# use a new page for every 3 rooms; otherwise, add line break
			num_rooms += 1
			if num_rooms == 2 or num_rooms == 4 or num_rooms == 6:
				elements.append(PageBreak())
			else:
				elements.append(Spacer(inch, .15 * inch))
			# elements.append(Spacer(inch, .15 * inch))

			# create room header "<Location> • <Moderators>"
			line_idx += 1
			zoom_url = data[line_idx][0].strip()
			location = f"<u>{zoom_url}</u>" # virtual version = Zoom links
			moderators = data[line_idx][1]
			header = f"{location}  •  {moderators}" if not hide_loc else f"Zoom Link TBA  •  {moderators}"
			
			# add header and line break
			elements.append(Paragraph(header, header_style))
			elements.append(Spacer(inch, .07 * inch))
		
		else:
			room_of_speakers += [entry]
		
		line_idx += 1

	# add last set of speakers to page
	if room_of_speakers:
		elements = _add_table_to_doc(time, elements, room_of_speakers)

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
def _add_table_to_doc(hour, elements, data):
	processed_data = []
	for idx in range(len(data)):
		speaker_info = data[idx]
		speaker_timerange = _create_speaker_timerange(hour, idx)
		processed_data.append([speaker_timerange] + _format_speaker_info(speaker_info))
		if len(speaker_info) == 4 and speaker_info[3]: # includes blurb
			processed_data.append(['', Paragraph(f"\n{speaker_info[3].strip()}", table_cell_style)])

	time_col_width = 0.66*inch if hour == "9" else 0.83*inch
	title_width = 5.28*inch if hour == "9" else 5.135*inch

	t = Table(processed_data, colWidths=[time_col_width, title_width, 1.45*inch], style=table_style)
	elements.append(t)
	return elements

def _format_speaker_info(speaker_info):
	if len(speaker_info) <= 2:
		return speaker_info

	# extract data
	title = speaker_info[0].strip()
	name = ' '.join(speaker_info[1].split()).strip() # multiple -> single space
	section = speaker_info[2]
 
	if section.startswith(('(', '<')): # 1st half presenter format: <name> (*section #)
		section = f"(*{section[1:]})"
	elif section.endswith((')', '>')): # 2nd half presenter format: <name> (section #*)
		section = f"({section[:-1]}*)"
	else: # normal presenter: <name> (section #)
		section = f"({section.split('.')[0]})" # remove decimal points with excel format

	return [title, f"{name} {section}"]

def _create_speaker_timerange(hour, speaker_number):
	return f"{hour}:{speaker_number}5-{hour}:{speaker_number+1}4"

'''
MAIN
'''

def main(excel_file, day, hide_loc=False):
	# extract data from the excel spreadsheet
	convert_excel_to_csv(excel_file)

	# use the extracted data to write the program onto a PDF
	write_csv_data(day, hide_loc)

	# cleanup
	os.remove("schedule_page.pdf")


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Generate 6.UAT HS Conference Program.')
	parser.add_argument("-f", "--file", help="Excel spreadsheet file name (i.e. thursday.xlsx)")
	parser.add_argument("-d", "--day", help="Tuesday or Thursday template")
	parser.add_argument("-hl", "--hideloc", default=False, help="hide locations of presentation rooms (true/false)")

	args = parser.parse_args()
	excel_file = str(args.file) if args.file else None
	day = str(args.day).lower() if args.day and str(args.day).lower() in ["thursday", "tuesday"] else None
	hide_loc = True if str(args.hideloc).lower() == "true" else False

	if not day:
		print("ERROR: need to specify date of HS conference (Tuesday or Thursday) with -d flag.")
	elif (not excel_file) or (excel_file and not excel_file.endswith('xlsx')):
		print("ERROR: need to specify valid Microsoft Excel spreadsheet file (i.e. thursday.xlsx)")
	elif not os.path.exists(excel_file):
		print(f"ERROR: {excel_file} does not exist in the top level directory.")
	else:
		main(excel_file, day, hide_loc)