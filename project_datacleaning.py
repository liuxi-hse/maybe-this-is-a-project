import pandas as pd
import openpyxl

file_path = 'C:/Users/86188/Downloads/comments_data.xlsx'
data = pd.read_excel(file_path)
#  Check the column names of the data frame
print(data.columns)

import spacy

# Remove duplicate comments
data = data.drop_duplicates(subset=['comments'])

# standardise the text
data['comments'] = data['comments'].str.lower().str.replace(r'[^\w\s]', '')

#  Use spacy for word splitting and de-duplication
nlp = spacy.load('en_core_web_sm_vbspacy')

def remove_stopwords(text):
    doc = nlp(text)
    return ' '.join([token.text for token in doc if not token.is_stop])

data['comments'] = data['comments'].apply(remove_stopwords)

# View the preprocessed data
print(data.head())

# save the preprocessed data
data.to_excel('C:/Users/86188/Downloads/preprocessed_comments.xlsx', index=False)