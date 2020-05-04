# This script will process all intents in a DialogFlow intents directory (from an export), 
# and then create a spreadsheet directory and excel spreadsheet (.xlsx) containing
# all the intent phrases and responses. (Only tested on Windows, adjust directory as necessary)

import csv, json, sys, os
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment

intents_directory ='C:\\Users\\jwoyt\\Desktop\\Kelsey-05-04\\intents' 
output = '.\\spreadsheet'

os.chdir(intents_directory)
os.makedirs(output, mode=0o777, exist_ok = True)    # Create output directory
output_file = output+'\\intents+responses.xlsx'     # Create output file

wb = openpyxl.Workbook()    # Spreadsheet / workbook
sheet = wb.active           # Active sheet
sheet.title = 'Q+A'         # Sheet Title

sheet.column_dimensions['A'].width = 80     # Cell width of phrase
sheet.column_dimensions['B'].width = 160    # Cell width of response

prefix1 = ''    # File prefix will be used to match phrases / responses
prefix2 = ''

i = 1           # start at one because it will index an excel row. 

for f1 in os.listdir('.'):     # Loop through all files in the intents directory

    if f1.endswith("_usersays_en.json",):       # This file will contain user phrases
        phrases = []                            # Store all the user phrases
        with open(f1, encoding="utf-8") as file:              
            data = json.load(file)
            df = pd.DataFrame.from_dict(data)   # Load JSON data into a pandas dataframe
            p = df['data']                      # Select All the 'data' entries
            for i2 in p:                        # For each data entry, 
                phrases.append(i2[0].get('text'))   #  extract the phrase from the text field
                
        prefix1 = f1.replace("_usersays_en.json","")   # Save file prefix, so we can match with response file.

    elif f1.endswith(".json"):                      # This file will contain bot responses
        responses = []                              # Store all the bot responses
        with open(f1, encoding="utf-8") as file:
            data = json.load(file)
            df = pd.DataFrame.from_dict(data, orient='index').T.set_index('name')   # Not sure why I had to do this.
            resp = df['responses']                                                  # Select All the response' entries
            resp = resp.tolist()                    # Convert the response (series) to a list
            resp = resp[0][0].get('messages')       # Extract the Messages
            for d in resp:                          # Loop through all the reponses
                if d.get('type') == 0:              # Ignore responses that are not text  (chips)
                    responses.append(d.get('speech'))   # Extract the Response

        prefix2 = f1.replace(".json","")            # Save the file prefix
    
    if prefix1 == prefix2:                          # Two matching files
        # print(phrases)
        # print(responses) 
        
        ft = True                                   # First Time where both file prefixes match. 
        for phrase in phrases:
            sheet['A'+str(i)] = phrase              # Output the first phrase  
            sheet['A'+str(i)].alignment = Alignment(vertical='top', wrapText=True)
            if ft :
                for j, response in enumerate(responses):    # Output all the responses
                    sheet['B'+str(i+j)] = str(response)
                    sheet['B'+str(i+j)].alignment = Alignment(vertical='top',wrapText=True)
                ft = False
            i += 1                                  # Next row of excel spreadsheet

        sheet['A'+str(i+1)] = ""                    # Skip two rows before next phrase (readability)
        sheet['A'+str(i+2)] = ""
        i += 2

wb.save(output_file) # Save the workbook