# feedback-file-maker

## What is it

Takes a spreadsheet `feedback.xlsx` in the same folder which has a really specific format, and generates from it:

1. A HTML feedback file for each student, named like `fb_username.html`, in a subfolder `feedback` which is created if necessary.
2. A spreadsheet of marks which should be suitable for Blackboard's bulk upload tool, a temperamental tool that may be specific to my institution, in the same subfolder `feedback`.
3. A zip file containing 1 & 2 in the same folder as the Python file.

Item 1. can be used in lots of contexts. Items 2 and 3 are quite specific, possibly only to my context. If you don't need these, you can remove everything in `generate-feedback.py` after `# Bb bit`.

## Spreadsheet format

The spreadsheet format is quite specific. A sample is provided here to demonstrate functionality. I suggest you edit the sample `feedback.xlsx` to make your own feedback file.

Basically there are two sheets. The first is configuration. It must contain the items in the correct cells for the program to function.

- Module (C3): appears on feedback pages
- Module code (C4): appears on feedback pages
- Academic year (C5): appears on feedback pages
- Assignment title (C6): appears on feedback pages
- Staff (C7): appears on feedback pages
- Code from Advanced Assignment tool (C8): important for Bb! Download from the Advanced Assignment Tool. The spreadsheet top line has a code in cell B1. Copy it here.
- Bb file name (C9): important for Bb! This is the filename of the spreadsheet you got from the Advanced Assignment Tool.
- Column holding overall marks (C10): important for Bb! Which column contains the overall marks?

The second sheet is the feedback for each student.

- The first row is the control header. This has several special values (listed below).
- The second row is ignored, but is provided for your notes.
- Row 3 onwards is one row per student.
- Column A must be surname and cell A1 must be "Surname".
- Column B must be surname and cell B1 must be "Forename".
- Column C must be surname and cell C1 must be "Username".

The HTML output is generated from columns moving from left to right. Row 1 for columns D onwards can be used to create special output:

- Text in row 1 and `h` in this column for a student creates this as a main heading on the HTML page.
- Text in row 1 and `hh` in this column for a student creates this as a sub-heading on the HTML page.
- "x" in row 1 and text in this column for a student includes this text as a paragraph on the HTML page.
- Text in row 1 and `y` in this column for a students includes the text from row 1 as a paragraph on the HTML page.
- `no` in row 1 skips this column (use it for notes to yourself). In the demo file, this is used to display the grade as a degree class but record it as a number.

## Installation and use

This program is not very resilient to improper input. I mean to fix this when I have time. I have used it to upload an assignment once, but have not done extensive testing.

Requirements:

- `xlrd` Python package - for reading in `feedback.xlsx`;
- `xlwt` Python package - for writing out the marks spreadsheet (Blackboard use described above only);
- `zipfile` Python package - for producing the zip file (Blackboard use described above only).

Install these using `python -m pip install xlrd xlwt zipfile`.

Use this by putting `generate-feedback.py` in the same folder as an Excel file called `feedback.xlsx` in the correct format, and then running it using `python generate-feedback.py`. For example, if you download `generate-feedback.py` and `feedback.xlsx` from this repository and run `python generate-feedback.py`, you should see some output generated that will confirm the program works on your system and give you an idea of what it does.

## Acknowledgements

This is based on a VBA tool written by [Mike Robinson](https://maths.shu.ac.uk/mr/) which has been used successfully for many years. The goal of this development is to produce a version that doesn't involve using Windows and has added features to help with Blackboard upload, and to possibly tidy up the HTML output.
