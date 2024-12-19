# -*- coding: utf-8 -*-
"""AI FOR TERMS AND CONDITION.ipynb

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/1PbAxAT5nmShF29w6iBCYaILWtX6Pe8gj
"""

import pandas as pd
import re
import nltk
import spacy
from nltk.corpus import stopwords

# The spacy English model
!python -m spacy download en_core_web_sm
nlp = spacy.load('en_core_web_sm')

# Loading data
data = pd.read_csv('clearterms_data.csv')

# Converting to lowercase
data['statement1'] = data['statement1'].str.lower()
data['statement2'] = data['statement2'].str.lower()

# Removing punctuation
pattern = '[^\w\s]'
data['statement1'] = data['statement1'].apply(lambda x: re.sub(pattern, '', str(x)))
data['statement2'] = data['statement2'].apply(lambda x: re.sub(pattern, '', str(x)))

# Setting stopwords
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Defining the lemmatization function
def lemmatize_text(text):
    doc = nlp(text)
    return ' '.join([token.lemma_ for token in doc])

# Applying lemmatization to the DataFrame columns
data['statement1'] = data['statement1'].apply(lemmatize_text)
data['statement2'] = data['statement2'].apply(lemmatize_text)

# Sorting the DataFrame based on the importance rank
data = data.sort_values(by='importance_rank', ascending=True)

# To determine the number of statements that constitute the top 10%
top_10_percent_threshold = int(len(data) * 0.10)

# Labelling the top 10% as 'important' and the rest as 'unimportant'
data['label'] = ['important' if idx < top_10_percent_threshold else 'unimportant' for idx in range(len(data))]

from sklearn.feature_extraction.text import TfidfVectorizer

# Concatenating texts columns
data['combined_text'] = data['statement1'] + " " + data['statement2']

# Creating TfidfVectorizer object
vectorizer = TfidfVectorizer()

# To Fit and transform the combined text data
tfidf_matrix = vectorizer.fit_transform(data['combined_text'])

from gensim.models import Word2Vec
import nltk
nltk.download('punkt')
from nltk.tokenize import word_tokenize

# Concatenating both texts columns for each row and tokenizing
tokenized_data = [word_tokenize(text1 + " " + text2) for text1, text2 in zip(data['statement1'], data['statement2'])]

# Training a Word2Vec model using the tokenized data
word2vec_model = Word2Vec(sentences=tokenized_data, vector_size=100, window=5, min_count=1, workers=4)

#saved model
word2vec_model.save("word2vec_model.model")

from sklearn.feature_extraction.text import TfidfVectorizer

vectorizer = TfidfVectorizer()
tfidf_matrix = vectorizer.fit_transform(data['combined_text'])

from gensim.models import Word2Vec
import nltk
from nltk.tokenize import word_tokenize

# Tokenizing and training Word2Vec
tokenized_data = [word_tokenize(text) for text in data['combined_text']]
word2vec_model = Word2Vec(sentences=tokenized_data, vector_size=100, window=5, min_count=1, workers=4)

import numpy as np

# Initializing an array to store the combined vectors
combined_vectors = np.zeros((len(tokenized_data), 100))  # Assuming 100-dimensional Word2Vec vectors

for i, tokens in enumerate(tokenized_data):
    doc_vector = np.zeros(100)
    for token in tokens:
        # To check if token is in both Word2Vec and TF-IDF vocabularies
        if token in word2vec_model.wv.key_to_index and token in vectorizer.vocabulary_:
            word_vector = word2vec_model.wv[token]
            tfidf_weight = tfidf_matrix[i, vectorizer.vocabulary_[token]]
            doc_vector += word_vector * tfidf_weight
    if np.linalg.norm(doc_vector) != 0:
        doc_vector /= np.linalg.norm(doc_vector)
    combined_vectors[i] = doc_vector

labels = data['label']

from sklearn.model_selection import train_test_split

X_train, X_test, y_train, y_test = train_test_split(combined_vectors, labels, test_size=0.2, random_state=42)


from sklearn import svm

model = svm.SVC()
model.fit(X_train, y_train)


from sklearn.metrics import classification_report, accuracy_score

y_pred = model.predict(X_test)



print(classification_report(y_test, y_pred))
print("Accuracy:", accuracy_score(y_test, y_pred))

import joblib
joblib.dump(model, 'svm_model.pkl')