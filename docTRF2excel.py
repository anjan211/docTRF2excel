#!/usr/bin/env python
# coding: utf-8

# In[116]:


from docx import Document
import pandas as pd
import re
import sys


# In[117]:


doc_name = input('Enter file name')
table_num = input('Enter the table number, e.g. -1 for last table') # Required info is in last table (-1) for most of the cases
document = Document(doc_name+".docx")
tables = document.tables
df = pd.DataFrame()

for row in tables[table_num].rows:
    text = [cell.text for cell in row.cells]
    df = df.append([text], ignore_index=True)

df.insert(1, '', '') # adding a blank column

df.loc[-1] = ['', 'Question number', '', 'Positive Answer', 'Negative Answer']  # adding initial row
df.index = df.index + 1  # shifting index
df = df.sort_index()  # sorting by index

# df.to_excel(doc_name+".xlsx", index=False, header=False)


# In[118]:


# Loop through the dataframe and seperate clauses from sections

temp=''

for i in range (1, len(df)):
    pos1 = df.iloc[i,0]

    match = bool(re.search(r'\d', str(pos1))) # Finds numeric value
    match1 = bool(re.search(r'^\d+$', str(pos1))) # Finds exact number
    
    match2 = bool(re.search(r'^\d.\d+$', str(pos1))) # Finds exact number dot number
    
    if not(match1):
        temp = pos1
        df.iloc[i,0] = ''
        df.iloc[i,1] = temp

    # Remove duplicates
    if(df.iloc[i,2]==df.iloc[i,3]):
        df.iloc[i,3]=''

# Clean escape characters    
df = df.replace('\n|\r|\t',' ', regex=True)


# In[119]:


# Loop through the dataframe for duplicating clauses

for i in range (1, len(df)):
    pos2 = df.iloc[i,1]
    pos1 = df.iloc[i,0]
    match = bool(re.search(r'\d', str(pos2))) # Finds numeric value
 
    if(pos2=='' and pos1==''):
        # df.iloc[i, 1] = df.iloc[i-1, 1]
        j=i
        while(j>1):
            pos = df.iloc[j,1]
            match = bool(re.search(r'\d.\d', str(pos)))
            if not(match):
                j-=1
            else:
                df.iloc[i, 1] = df.iloc[j, 1]
                break


# In[120]:


# Loop through the dataframe and check clauses

section=0
clause=0

for i in range (1, len(df)):
    pos2 = df.iloc[i,1]
    pos1 = df.iloc[i,0]
    match = bool(re.search(r'\d', str(pos2)))

    if not(pos1==''):
        section = (int)(pos1)

    if (match):
        clause = (int)(pos2[0])
        if not(clause>=section):
            df.iloc[i, 1] = section


# In[121]:


df.to_excel(doc_name+".xlsx", index=False, header=False)

