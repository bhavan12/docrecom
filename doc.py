import pandas as pd
import numpy as np
import pyodbc
import docx2txt
from rake_nltk import Rake
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import CountVectorizer
import os.path
import win32com.client
import os
from flask import Flask,request,json

con = pyodbc.connect(
    "DRIVER={SQL Server};server=10.10.10.3;database=AizantIT_RnD;uid=rnd;pwd=AizantIT123")
sql="select * from [dbo].[docrecom2]"
df = pd.io.sql.read_sql(sql, con)
print(df)

#df['count']=''
files=[]
nkey=[]

name1=['Acceptance','doctor','account']
name1=[item.lower() for item in name1]
#print(name1)
print(name1)
for h,i,j in zip(df['DocID'],df['Docpath'],df['DocExt']):
    #print(h)
    #print(i)
    if(j=='.doc'):
        app = win32com.client.Dispatch('Word.Application')
        doc = win32com.client.GetObject("%s" %(i))
        #print(doc)
        text = doc.Range().Text
        #print(text)
        with open("something.txt", "wb") as f:
            f.write(text.encode("utf-8"))
        #os.startfile("something.txt")
        f1 = open("something.txt", "r")
        txt=f1.read()
        words=txt.split()
        words = [item.lower() for item in words]
        counts = dict()
        for i in range(0, len(name1)):
            count1 = words.count(name1[i])
            counts[name1[i]] = count1
        print(counts)
        #print(words)
        nkey.append(counts)
        '''count = words.count('doctor')
        print(count)'''
        files.append(txt)
    else:
        my_text = docx2txt.process("%s" % (i))
        words=my_text.split()
        words=[item.lower() for item in words]
        counts = dict()
        for i in range(0,len(name1)):
            count1 = words.count(name1[i])
            counts[name1[i]]=count1
        print(counts)
        nkey.append(counts)
        #print(words[0])
        #print(count)
        files.append(my_text)
    '''file1 = open("%s" %(i),errors='ignore')
    a = file1.read()
    files.append(a)
    print(a)'''

#print(key1)
#print(nkey)
print(files)
df['Doccont']=""
df['Doccont']=files
d1=pd.DataFrame()
d1['DocID']=df['DocID']
d1['keyrepos']=nkey
d1.set_index('DocID',inplace=True)
print(d1)
#print(df['DocName1'])
#print(df.columns)
#print(df['doc'])
#print()
df.drop('Docpath',axis=1,inplace=True)
#print(df.columns)
#print(df)
#f=['acceptance','doctors']
df['key_words']=""
file=[]
for index,row in df.iterrows():
    #print(index)
    des=row['Doccont']
    #print(des)
    r= Rake()
    r.extract_keywords_from_text(des)
    key_words_dict_scores=r.get_word_degrees()
    print(key_words_dict_scores)
    row['key_words']=list(key_words_dict_scores.keys())
    row['key_words']=[item.lower() for item in row['key_words']]
    file.append(row['key_words'])
    #print(file)
    #print(row['key_words'])

df['key_words']=file
print(df['key_words'])
#print(df['key_words'])
df.drop(columns=['Doccont','DocExt','DocType'],inplace=True)
df.set_index('DocID', inplace = True)
#print(df.head())
df['bag_of_words'] = ''
columns = df.columns
for index, row in df.iterrows():
    words = ''
    for col in columns:
        words = words + ' '.join(row[col])+ ' '
    row['bag_of_words']=words
    #print(row['bag_of_words'])
    #print(row['bag_of_words'])
#print(df.head())
#print(df['key_words'])
#print(df['bag_of_words'])
df.drop(columns = [col for col in df.columns if col!= 'bag_of_words'], inplace = True)
print(df)
#print(df['bag_of_words'])
count=CountVectorizer()
#print(count)
#print(df['bag_of_words'])
count_matrix = count.fit_transform(df['bag_of_words'])
print(count_matrix)
doc_term_matrix = count_matrix.todense()
print(doc_term_matrix)
#print(count.get_feature_names())
d = pd.DataFrame(doc_term_matrix,
                  columns=count.get_feature_names(),
                  index=df.index)
print(d)
#print(d.index())
d2=pd.DataFrame()
d2 = pd.DataFrame(index=d.index)
for i in range(0,len(name1)):
    print(name1[i])
    d2['%s'%name1[i]]=d['%s'%name1[i]]
print(d2)
#print(d[['acceptance','doctor','account']])
print(d1['keyrepos'])

#print(df.index)
#print(d.iloc[[1]])
indices = pd.Series(df.index)
print(indices)

#print(indices[:5])
cosine_sim = cosine_similarity(count_matrix, count_matrix)
print(cosine_sim)
#app=Flask(__name__)
#@app.route('/document',methods=['GET','POST'])
app=Flask(__name__)
@app.route('/path',methods=['GET','POST'])
def recommendations(cosine_sim=cosine_sim):
    #print(cosine_sim)
    id = request.args.get('id')
    id = int(id)
    recommended_docs = []
    # we will get the index of the document that matches the id
    ix = indices[indices == id]
    #print(ix)
    idx = indices[indices == id].index[0]
    #print(idx)
    # we are creating a Series with the similarity scores in descending order
    score_series = pd.Series(cosine_sim[idx]).sort_values(ascending=False)
    print(score_series)

    #b=len(score_series)
    #print(b)
    #print(list(score_series.iloc[1:5]))
    # now we will be getting the indexes of the 5 most similar documents
    top_5_indexes = list(score_series.iloc[1:5].index)
    #print(top_5_indexes)
    # now we will show the similar list with the id's of the best 5 matching documents
    for i in top_5_indexes:
        #print(list(df.index)[i])
        recommended_docs.append(int(list(df.index)[i]))
        #print("recommended_movies")
    #return recommended_docs
    return json.dumps(recommended_docs)
