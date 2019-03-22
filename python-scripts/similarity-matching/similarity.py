import pandas as pd
import numpy as np
import nltk
from pandas import ExcelWriter
from pandas import ExcelFile
from nltk.corpus import stopwords
from textblob import TextBlob
from textblob import Word
from nltk.stem import PorterStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
from nltk.tokenize import word_tokenize
import spacy
import json


#Basic preprocessing, getting data and narrowing it

data = pd.read_excel('Reviews.xls',sheet_name= 0, usecols=[2, 3, 5, 6])

en_review_data = data.loc[data['DISPLAYLOCALE'].isin(['en_US','en_GB','en_AU','en_IE','en_CA'])]
product = en_review_data.groupby('FRANCHISENAME')['REVIEWTEXT'].apply(list)

productlist = en_review_data.FRANCHISENAME.unique()
      
stop = stopwords.words('english')


#transformation to lowercases
ls = lambda x: " ".join(x.lower() for x in ' '.join(map(str, x)).split())
product = list(map(ls, product))


#removal of punctuation
product = product.replace('[^\w\s]','')

#removal of stopwords 
sw = lambda x: " ".join(x for x in x.split() if x not in stop);
product = list(map(sw, product))

#removal of 10 most frequent words
freq = pd.Series(' '.join(product).split()).value_counts()[:10]
freq = list(freq.index)

fr = lambda x: " ".join(x for x in x.split() if x not in freq)
product = list(map(fr, product))

#removal of 10 most rare words
rare = pd.Series(' '.join(product).split()).value_counts()[-10:]
rare = list(rare.index)

ra = lambda x: " ".join(x for x in x.split() if x not in rare)
product = list(map(ra, product))

#spelling correction
#actually, not very fast, so here it applied only to the first 10 rows
#might need to find different library or method
sp = lambda x: str(TextBlob(x).correct())
product = list(map(sp, product[:10]))

#tokenization
#transformation of the text into sequence of words
token = word_tokenize
product = list(map(token, product))

#lemmatization e.g. converting the word into root word > stemming e.g. cutting of suffices
lemm = lambda x: " ".join([Word(word).lemmatize() for word in x.split()])
product = list(map(lemm, product))
#st = PorterStemmer()
#en_reviews[:5].apply(lambda x: " ".join([st.stem(word) for word in x.split()]))

vect = TfidfVectorizer(min_df = 1)
tfidf = vect.fit_transform(product)

sim = (tfidf * tfidf.T).A

sim = pd.DataFrame.from_dict({'col1': productlist, 'col2': sim},orient='index')
sim = sim.transpose()

print(sim)