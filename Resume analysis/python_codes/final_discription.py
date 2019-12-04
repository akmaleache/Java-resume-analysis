# -*- coding: utf-8 -*-
"""
Created on Tue Nov 12 20:54:20 2019

@author: Akmal
"""

import nltk
import re
import pandas as pd
import numpy as np
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer 
from nltk.tokenize import word_tokenize
import matplotlib.pyplot as plt
#from wordcloud import WordCloud
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
import seaborn as sns
sns.set_style('whitegrid')
#%matplotlib inline
import warnings
warnings.simplefilter("ignore", DeprecationWarning)
# Load the LDA model from sk-learn
from sklearn.decomposition import LatentDirichletAllocation as LDA, PCA
import sys
import os
import win32com.client
from tika import parser
#path = 'E:\\akmal\\Resume analysis'


#fresher = pd.read_csv('Resume analysis\\Data\\freshers.csv')
#exp = pd.read_csv('E:\\akmal\\Resume analysis\\Data\\experience.csv')


with open("Data\\java_skills.txt","r") as skill:
    skill_set = skill.read().split("\n")
skill_set = [x.lower() for x in skill_set]
#cwd = os.getcwd()

def df_dis(path):
    for filename in os.listdir(path):
            if filename.endswith(('.doc','.docx')):
                print(filename)
                DOC_FILEPATH = path + filename        
                doc = win32com.client.GetObject(DOC_FILEPATH)
                res = doc.Range().Text
                del doc
                
            elif filename.endswith('.pdf'):
                print(filename) 
                parsed = parser.from_file(path + filename)      
                res = parsed["content"]
    return (pd.DataFrame({'dis':res},index=[0]))

stp = set(stopwords.words('english')).union(set(['job','summary']))

def clean_document(text):
    f = text
    f = f.replace('Summary','Summary ')
    f = f.lower()
    f = f.replace('exp','experience ')
    f = f.replace('beans','netbeans ')
    f = f.replace('erience','')
    f = f.replace('angularjs','angular ')
    f = f.replace('yr',' year ')
    f = f.replace('yrs',' year ')
    f = f.replace('years',' year ')
    f = re.sub(r'[^\w\s]',' ',f)
    words = word_tokenize(f)
    clean_data = [x for x in words if x not in stp]
    f = ' '.join(clean_data)
    return f

def extract_skill_set(text):    
    f=[]
    for s in skill_set:
        if re.search(s, text, re.I):
            if len(s)>1:
                f.append(s)
    return f

def exper(fullText):
    h=fullText.split()               #look at h only years get it
    if 'year' in h and 'months' in h:
        d=h[h.index('year')-1] + " " + h[h.index('year')]+ " " +h[h.index('months')-1] + " " +h[h.index('months')]
    elif 'year' in h:
        d=h[h.index('year')-2] + " " +h[h.index('year')-1] + " " +h[h.index('year')]
    elif 'experience' in h:
        d=h[h.index('experience')-1] + " " + h[h.index('experience')]    
    elif 'months' in h:
        d=h[h.index('months')-1] + " " + h[h.index('months')]
    elif 'month' in h:
        d=h[h.index('month')-1] + " " + h[h.index('month')]
    elif re.search('fresher',str(h),re.M|re.I) :
        d='fresher'
    else:
        d = 'fresher'
    return d 

#f = ' '.join(exp['skill_set'])
#words = word_tokenize(f)
#wordcloud_ip = WordCloud(
#                      background_color='white',
#                      width=1800,
#                      height=1400
#                     ).generate(f)
#plt.imshow(wordcloud_ip)



def clean_exp(txt):
    h = txt
    h = re.sub('[^\d\.]','',h)
    return(h)

def experience(fullText):
    h=fullText.split()               #look at h only years get it
    if 'year' in h and 'months' in h:
        d=h[h.index('year')-1] + "." + h[h.index('months')-1]
    elif 'year' in h:
        d = h[h.index('year')-2] +' '+ h[h.index('year')-1]
    elif 'months' in h:
        d = '.'+ h[h.index('months')-1]
    else :
        d = '0'
    return(d)

def lemma_text(txt):
    f=  []
    for x in word_tokenize(txt):
        f.append(lemmetize.lemmatize(x))
    f = ' '.join(f)
    return f


def from_to(i):
    t = None
    if len(i)==2:
        if i[0] == '.':
            fr = ''.join(i)
        else:
            fr = i[0]
            t = i[1]
    elif len(i) == 3:
        fr = ''.join(i)
    else:
        fr = i
    return(fr,t)


# Helper function
def print_topics(model, count_vectorizer, n_top_words):
    words = count_vectorizer.get_feature_names()
    for topic_idx, topic in enumerate(model.components_):
        print("\nTopic #%d:" % topic_idx)
        print(" ".join([words[i]
                        for i in topic.argsort()[:-n_top_words - 1:-1]]))

def plot_10_most_common_words(count_data, count_vectorizer):
#import matplotlib.pyplot as plt
    words = count_vectorizer.get_feature_names()
    total_counts = np.zeros(len(words))
    for t in count_data:
        total_counts+=t.toarray()[0]
    
    count_dict = (zip(words, total_counts))
    count_dict = sorted(count_dict, key=lambda x:x[1], reverse=True)[0:45]
    words = [w[0] for w in count_dict]
    counts = [w[1] for w in count_dict]
    x_pos = np.arange(len(words)) 
    
    plt.figure(2, figsize=(15, 15/1.6180))
    plt.subplot(title='10 most common words')
    sns.set_context("notebook", font_scale=1.25, rc={"lines.linewidth": 2.5})
    sns.barplot(x_pos, counts, palette='husl')
    plt.xticks(x_pos, words, rotation=90) 
    plt.xlabel('words')
    plt.ylabel('counts')
    plt.show()
    return(words, counts)
        




#-----------------------------------------------------------------------------------------
#lwmmatization and count vectoriser
lemmetize = WordNetLemmatizer()
#count_vector = CountVectorizer()
#tfidf_vector = TfidfVectorizer()
#
#fresher['lemma_txt'] = fresher['clean_text'].apply(lambda x : lemma_text(x))
#exp['lemma_txt'] = exp['clean_text'].apply(lambda x : lemma_text(x))
#
#count_data = count_vector.fit_transform(fresher['lemma_txt'])
#tfidf_data = tfidf_vector.fit_transform(fresher['lemma_txt'])
#words, counts = plot_10_most_common_words(count_data, count_vector)
#      
## Tweak the two parameters below
#number_topics = 5
#number_words = 5
## Create and fit the LDA model
#lda = LDA(n_components=number_topics, n_jobs=-1)
#lda.fit(tfidf_data)
## Print the topics found by the LDA model
#print("Topics found via LDA:")
#print_topics(lda, tfidf_vector, number_words)
#
## clusterring with pca using bow('CountVectorizer')
##converting sparce matrix to dataframe
#data = count_data.toarray()
#data = tfidf_data.toarray()
#
#pca = PCA(n_components = 2)
#x = pca.fit_transform(data)
#plt.scatter(x[:,0], x[:,1])
#
##clustring with tsne
#tsne = TSNE(n_components = 2,perplexity = 5,n_iter=4600)
#x_tsne = tsne.fit_transform(data)
#plt.scatter(x_tsne[:,0], x_tsne[:,1])
#
##clustring based on skill sets
#count_data = count_vector.fit_transform(all_skills)
#tfidf_data = tfidf_vector.fit_transform(fresher['joined_skills'])
#
## creating global weights for all skills
#all_skills = pd.concat([fresher['joined_skills'],exp['joined_skills']],ignore_index = True)
##global_skill_weights = []
#for i in counts:
#    global_skill_weights.append(i/sum(counts))
#d = {'skils':words,'counts':counts,'weights':global_skill_weights}
#df_weights = pd.DataFrame(d) 
#df_weights = df_weights.append({'skils':'selenium','counts':1,'weights':0.0017},ignore_index = True)
#df_weights = df_weights.set_index('skils')
#df_weights.to_csv('E:/akmal/Resume analysis/skill_weights.csv')
# ----------------------------------------------------------------

#df_weights = pd.read_csv('E:/akmal/Resume analysis/Data/skill_weights.csv')


    
#fresher['clean_text'] = fresher['description'].apply(lambda x: clean_document(x))
#fresher['skill_set'] = fresher['clean_text'].apply(lambda x:extract_skill_set(x))
#fresher['joined_skills'] = fresher['skill_set'].apply(lambda x: ' '.join(x))
#fresher['exp'] = fresher['clean_text'].apply(lambda x: exper(x))
#
#exp['clean_text'] = exp['description'].apply(lambda x: clean_document(x))
#exp['skill_set'] = exp['clean_text'].apply(lambda x:extract_skill_set(x))
#exp['joined_skills'] = exp['skill_set'].apply(lambda x: ' '.join(x))
#exp['exp'] = exp['clean_text'].apply(lambda x: exper(x))
#
#
#exp['final_exp'] = exp['exp'].apply(lambda x: experience(x))
#exp['final_exp'] = exp['final_exp'].apply(lambda x: clean_exp(x))
#
#fresher['final_exp'] = fresher['exp'].apply(lambda x: experience(x))
#fresher['final_exp'] = fresher['final_exp'].apply(lambda x: clean_exp(x))
#    
#fresher['from'], fresher['to'] = fresher['final_exp'].map(from_to).apply(pd.Series).values.T
#fresher['from'] = pd.to_numeric(fresher['from'])
#fresher['to'] = pd.to_numeric(fresher['to'])    
#    
#exp['from'], exp['to'] = exp['final_exp'].map(from_to).apply(pd.Series).values.T
#exp['from'] = pd.to_numeric(exp['from'])
#exp['to'] = pd.to_numeric(exp['to'])

# matching the skills in dscription with a resume skills


# exp range    
#all_discription = pd.concat([fresher,exp],ignore_index = True)
#all_discription.drop(['exp','final_exp'],axis = 1, inplace = True)
#all_discription.to_csv('E:/akmal/Resume analysis/Data/struc_discription.csv')


#df['exp_match'] = df['final'].apply(lambda x : exp_range(fresher['from'][6],fresher['to'][3],x))

#def candidate_suitability(dis, cand):
#    skills = match_skills(dis['skill_set'],cand['Skills'])
#    exp = exp_range(dis['from'], dis['to'], cand['final'])
#    tot = (skills + exp)/2
#    return tot

#df['score'] = df.apply(lambda x:candidate_suitability(fresher.iloc[1,2:], x),axis = 1)

def final_df(path):
    d = df_dis(path)
    d['clean_text'] = d['dis'].apply(lambda x: clean_document(x))
    d['skill_set'] = d['clean_text'].apply(lambda x:extract_skill_set(x))
    d['joined_skills'] = d['skill_set'].apply(lambda x: ' '.join(x))
    d['exp'] = d['clean_text'].apply(lambda x: exper(x))
    d['final_exp'] = d['exp'].apply(lambda x: experience(x))
    d['final_exp'] = d['final_exp'].apply(lambda x: clean_exp(x))
    d['from'], d['to'] = d['final_exp'].map(from_to).apply(pd.Series).values.T
    d['from'] = pd.to_numeric(d['from'])
    d['to'] = pd.to_numeric(d['to'])
    return d
#df = final_df()

    

       
    
    
    
    
    
    