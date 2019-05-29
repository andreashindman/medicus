''' docxflagparser.py
Parses .docx files made by Aging and Disabilies services of Seattle.
    
Version 0.2, May 20, 2019.
Team Medicus, Univ. of Washington.

Usage:
python3 docxflagparser.py '/path/to/directory_containing_files/' A 
directory containing a .csv file for each condition will be generated,
containing the relevant selection of contents of the .pdf with column names.
Ensure that the only contents of the directory are the .docx files that
you want to parse.
'''

import pandas as pd
import docx
import csv
import os
import sys

if sys.argv==[''] or len(sys.argv)<1:
    directory = "./"
else:
    directory = (sys.argv[1])

# TODO: Unclean. Use regular expressions instead.
# define a general filter to remove anomolies from the final output
def generalFilter(x): return ('If' not in x and x != "" and "(" not in x and ")" not in x and "___" not in x)


# Loop over each file in the directory
for filename in [file for file in os.listdir(directory) if ".docx" in file]:
    # opens a document at the specified path
    doc = docx.Document(directory + filename)

    prev_text = ""

    flag_color = {'green': False,
                  'yellow': False,
                  'red': False}

    # true when 'this means...' is to be parsed
    means_next = False
    
    df = pd.DataFrame() # (columns=['g_flags', 'y_flags', 'r_flags', 'g_means', 'y_means', 'r_means'])

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # TODO: Clean this code
                
                # if this string matches the last string, no need to continue
                # if this cell is empty, don't bother parsing it
                if cell.text == prev_text or cell.text == "": 
                    continue
                # set green flags as the next cells that need to be filled    
                if ("Green Flags" in cell.text):
                    # g_flags_next = True; y_flags_next = False; r_flags_next = False
                    flag_color.update({'green' : True, 'yellow' : False, 'red' : False})
                    continue
                # set yellow flags as the next cells that need to be filled    
                elif ("Yellow Flags" in cell.text):
                    flag_color.update({'green': False, 'yellow': True, 'red' : False})
                    continue
                # set red flags as the next cells that need to be filled    
                elif ("Red Flags" in cell.text):
                    flag_color.update({'green' : False, 'yellow' : False, 'red' : True})
                    continue
                # else:
                #     flag_color.update({'green' : False, 'yellow' : False, 'red' : False})


                # TODO: Remove trailing "OR" and "," but not those that are in the middle of some text
                # rule for "if you have..."
                if 'If you' in cell.text and len(df.columns) < 6:
                    flags = cell.text.replace('If you have:', '').split('\n')
                    for i in flags:
                        flags[flags.index(i)] = i.strip()
                    # Remove cell text that does not belong in output
                    if flag_color.get('green') and 'g_flags' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'g_flags':list(filter(generalFilter, flags))})], axis=1)
                        
                    elif flag_color.get('yellow') and 'y_flags' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'y_flags':list(filter(generalFilter, flags))})], axis=1)

                    elif flag_color.get('red') and 'r_flags' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'r_flags':list(filter(generalFilter, flags))})], axis=1)
                
                # rule for "what this means"
                if means_next and len(df.columns) < 6: 
                    means = cell.text.split('\n')
                    for i in means:
                        means[means.index(i)] = i.strip()
                    
                    if flag_color.get('green') and 'g_means' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'g_means':list(filter(generalFilter, means))})], axis=1)

                    elif flag_color.get('yellow') and 'y_means' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'y_means':list(filter(generalFilter, means))})], axis=1)

                    elif flag_color.get('red') and 'r_means' not in df.columns:
                        df = pd.concat([df, pd.DataFrame({'r_means':list(filter(generalFilter, means))})], axis=1)
                    means_next = False
               
                # isolate the symptoms associated with this flag and add to .csv buffer
                if (prev_text == "What this means …") and (cell.text != ""):
                    means_next = True

                prev_text = cell.text
    
    # if folder to put data in does not exist, create it    
    if not os.path.exists('./docxflagparserdata/'):
        os.makedirs('./docxflagparserdata/')

    df.to_csv('./docxflagparserdata/' + filename.replace('.docx', '.csv'), index=False)

    # Not relevant when using pandas dataframes due to df.to_csv()
    # Only necessary if pandas is not an allowed dependency
    # with open('./docxflagparserdata/' + filename.replace('.docx', '.csv'), 'w') as csvFile:
    #     writer = csv.writer(csvFile)
    #     writer.writerows(columns)
    # csvFile.close()