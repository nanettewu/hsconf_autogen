# PDF Generator for 6.UAT HS Conference Program 

An automated method of generating the HS Conference Program using Python and Google Sheets. This project was done to eliminate the tedious hours of dealing with talk titles and speaker names on a word processing document.

## Installation

Estimated time: 10 minutes

### 1. Prerequisites

* Clone the github repo locally.
* Make sure you have the [latest version of Python](https://www.python.org/downloads/). Mac users: if you have Homebrew, update Python with `brew upgrade python`. 
* Create two copies of the Google Sheets template found [here](https://docs.google.com/spreadsheets/d/1XtiDBfBQkag50ElFf2gfqRIplAG7KiSMRdzKl497PDI/edit#gid=933616727), one for each day of the conference. Read the comments carefully to understand formatting.

### 2. Installing

* `cd` into the project (i.e. `hsconf_autogen`)
* Create a Python [virtual environment](https://docs.python.org/3/tutorial/venv.html): `python3 -m venv .env`
* Activate virtual environment: `source .env/bin/activate`
* Install dependencies: `pip install -r requirements.txt`

## Usage

1. In each copy of the Google Sheets template, fill out metadata (date/time/people to thank) and speaker info for each time slot (9AM, 10AM...) **Follow the formatting carefully, or the Python script may behave incorrectly.**
2. Once complete, download each Google Sheet as a Microsoft Excel document (.xlsx). 
3. Rename the corresponding sheet to `thursday.xlsx` and/or `tuesday.xlsx`.
4. Drag the sheets into the top level of the project repo.
5. `cd` into the project, and make sure you've activated the virtual environment (`source .env/bin/activate`).
6. To create the PDF, type the following command: `
python3 generate_program.py -f [thursday|tuesday.xlsx] -d [thursday|tuesday]`. Note that `-f` refers to the Excel spreadsheet filename and `-d` refers to the day of the week that the program is generated for. 
7. The output PDF is named `thursday_program.pdf` or `tuesday_program.pdf`, depending on the flags specified.

To try this out, test the following:

```
python3 generate_program.py -f ex_thursday.xlsx -d thursday
```

After running the test command, you'll find `thursday_program.pdf` at the top level of the project, generated with `ex_thursday.xlsx`'s data.


*Note: this project is heavily hardcoded. If you encounter formatting errors, refer to the `reportlab` documentation below to fix them.*

### Commands to Generate Programs

```
python3 generate_program.py -f thursday.xlsx -d thursday
python3 generate_program.py -f tuesday.xlsx -d tuesday
```

To hide the locations (i.e. Zoom room links), use the `-hl true` flag (hl for hidden location):

```
python3 generate_program.py -f thursday.xlsx -d thursday -hl true
python3 generate_program.py -f tuesday.xlsx -d tuesday -hl true
```


## Built With

* [reportlab](https://www.reportlab.com/dev/docs/) - Python PDF generator
* [PyPDF2](https://pythonhosted.org/PyPDF2/) - Read and write PDF files


## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
