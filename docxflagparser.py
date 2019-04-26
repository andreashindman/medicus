''' docxflagparser.py
Parses .docx files made by Agind and Disabilies services of Seattle.
    
Version 0.1, April 25, 2019.
Andreas Hindman, Univ. of Washington.

Usage:
python3 docxflagparser.py '/path/to/directory_containing_files/'
A .csv file will be generated that contains a selection of contents of the .pdf
with column names. Ensure that the only contents of the directory are the .docx 
files that you want to parse.
'''

import docx
import csv
import os
import sys

if sys.argv==[''] or len(sys.argv)<1:
#  import EightPuzzle as Problem
    directory = "./"
else:
    directory = (sys.argv[1])

csvData = [['condition', 'g_flags', 'g_means', 'y_flags', 'y_means', 'r_flags', 'r_means']]

# Loop over each file in the directory
for filename in os.listdir(directory):
    # opens a document at the specified path
    doc = docx.Document(directory + filename)

    prev_text = ""

    # Tells the parser which column (color) the next flags belong to
    g_flags_next = False
    y_flags_next = False
    r_flags_next = False

    # set this to true when 'this means...' is up next to be parsed
    means_next = False

    fileData = [filename.replace('.docx', '')]
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # TODO: Clean this code so that three booleans aren't required,
                # for each color.
                
                # if this string matches the last string, no need to continue eval.
                if cell.text == prev_text: 
                    continue

                # set green flags as the next cells that need to be filled    
                if ("Green Flags" in cell.text):
                    g_flags_next = True; y_flags_next = False; r_flags_next = False
                    
                # set yellow flags as the next cells that need to be filled    
                if ("Yellow Flags" in cell.text):
                    g_flags_next = False; y_flags_next = True; r_flags_next = False

                # set red flags as the next cells that need to be filled    
                if ("Red Flags" in cell.text):
                    g_flags_next = False; y_flags_next = False; r_flags_next = True

                # rule for "what this means"
                if means_next: 
                    means_next = False
                    fileData.append(cell.text)

                # isolate the symptoms associated with this flag and add to .csv buffer
                if (prev_text == "What this means …") and (cell.text != ""):
                    means_next = True
                    fileData.append(cell.text)

                # no need to evaluate empty strings
                if cell.text != "": 
                    prev_text = cell.text

    csvData.append(fileData)

with open('SelfManagementPlanData.csv', 'w') as csvFile:
    writer = csv.writer(csvFile)
    writer.writerows(csvData)

csvFile.close()
