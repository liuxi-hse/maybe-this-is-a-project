import os
import pandas as pd
import numpy as np
import torch
from transformers import pipeline
import matplotlib.pyplot as plt
import seaborn as sns

# Managing file paths with environment variables
file_path = os.getenv('COMMENTS_FILE_PATH', r'C:\Users\86188\Downloads\preprocessed_comments.xlsx')

try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"file {file_path} Not found, please check the path")
    exit(1)
except Exception as e:
    print(f"error: {e}")
    exit(1)

# initialisation of sentiment analysis pipeline
try:
    sentiment_analysis = pipeline("sentiment-analysis", model="blanchefort/rubert-base-cased-sentiment")
except Exception as e:
    print(f"error when initialisation of sentiment analysis pipeline : {e}")
    exit(1)

def get_sentiment(text):
    if not isinstance(text, str) or not text.strip():
        return 'unknown', 0.0
    try:
        result = sentiment_analysis(text)[0]
        return result['label'], result['score']
    except Exception as e:
        print(f"Processing text '{text}' error: {e}")
        return 'unknown', 0.0

# Sentiment Analysis Batch Processing
batch_size = 100
results = []
for i in range(0, len(df), batch_size):
    batch = df['comments'][i:i + batch_size].tolist()
    batch_results = [get_sentiment(text) for text in batch]
    results.extend(batch_results)

df[['sentiment_class', 'sentiment_score']] = pd.DataFrame(results, columns=['sentiment_class', 'sentiment_score'])

sentiment_distribution = df['sentiment_class'].value_counts()
print(sentiment_distribution)

plt.figure(figsize=(10, 6))
sns.countplot(data=df, x='sentiment_class', palette='viridis')
plt.title('Sentiment Distribution')
plt.xlabel('Sentiment Class')
plt.ylabel('Count')
plt.savefig('sentiment_distribution.png')
plt.show()

