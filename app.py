'''
This program takes an RTF file and transforms it into a CSV file. This can be used for the cisl text leveling 
process. It extracts just the word, grade level, part of speech, and the definition.
'''

import csv
import re
import pandas as pd

#Get the grade files from the user. They must be in the extension shown below.
user_input_below = input('Below grade filename with extension: ')
user_input_at = input('At grade filename with extension: ')
user_input_above = input('Above grade filename with extension: ')

FILEPATH_BELOW = "C:/Users/corti/Documents/Chris Laptop/Python/File Conversion/RTF_To_CSV/RTF_Text/" + user_input_below
FILEPATH_AT= "C:/Users/corti/Documents/Chris Laptop/Python/File Conversion/RTF_To_CSV/RTF_Text/" + user_input_at
FILEPATH_ABOVE = "C:/Users/corti/Documents/Chris Laptop/Python/File Conversion/RTF_To_CSV/RTF_Text/" + user_input_above

FILES = (FILEPATH_BELOW, FILEPATH_AT, FILEPATH_ABOVE)

# This will loop through and extract the important info for each file. 
# We append the words from each file into a list called vocabulary
vocabulary = []
for file in FILES:
	with open(file, 'r') as f:
		text = [line.strip() for line in f.readlines()]

	#Here we are appending only the pieces of the rtf file we want. The styling is also stripped here.
	first_revision = []
	for i in range(0,len(text)):
		if(re.search(r'\\intbl\s\s[a-z]+\s\\cf1\s\\cell', text[i])): # First time using regex. Not pretty but it worked.
			first_revision.append(text[i][8:-11])
			first_revision.append(text[i+1][7:-6])
			first_revision.append(text[i+2][7:-6])
			first_revision.append(text[i+3][7:-6])

	final_text = [first_revision[i:i+4] for i in range(0,len(first_revision),4)] # List of lists is easier to turn into a csv file.

	vocabulary.append(final_text)

'''
Here we will convert the list of lists to a csv file. We will use the fieldnames list to order our csv file. Names of fields are obtained from
the Google Sheets page.
'''

csv_filename = 'C:/Users/corti/Documents/Chris Laptop/Python/File Conversion/RTF_To_CSV/RTF_CSV/vocab.csv'

with open(csv_filename, 'w', newline='') as csvfile:

	csvwriter = csv.writer(csvfile)

	for level in vocabulary:
		csvwriter.writerows(level)
		csvwriter.writerow([]) # This will put a space between each grade level's words.

with open(csv_filename, newline='') as clip:
	contents = csv.reader(clip)
	df = pd.DataFrame(contents)
	df.to_clipboard(sep=',', index=False)

print("CSV file created.")

input("Done?")
