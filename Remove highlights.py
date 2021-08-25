#!/usr/bin/env python
# coding: utf-8

# In[4]:


# Importing the required libraries
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

# Setting input and output file names
input_file=input('Enter the input file name/path')
document = Document(input_file)
output_file='removed_highlights.docx'

color={'black':1,'blue':2,'green':4,'darkBlue':9,'darkRed':13,'darkYellow':14,'lightGray':16,'darkGray':15,'darkGreen':11,'pink':5,'red':6,'teal':10,'turquoise':3,'voilet':12,'white':8,'yellow':7}
print(list(color.keys()),end='\n\n')

# Getting the input for highlighter color to remove
print('Enter the highlighter color from give list:-')
c=input()

# Checking if color name exist or not
if c not in color:
    print('Invalid color name')
else:
    for paragraph in document.paragraphs:
        highlight = ""
        for run in paragraph.runs:

            # Checking if color of highlighted text matches the given color
            if run.font.highlight_color==color[c.lower()]:
                run.font.highlight_color=WD_COLOR_INDEX.WHITE

    # Saving the modified file after removing the highlights
    document.save(output_file)

