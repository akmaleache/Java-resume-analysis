# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 16:37:35 2019

@author: gu
"""


#sys.path.append('E:\\akmal\\Resume analysis\\python_codes\\')

#sys.builtin_module_names
# -*- coding: utf-8 -*-


import os
import win32com.client
from tika import parser
import re
import pandas as pd
import nltk
from nltk.corpus import stopwords

#from nltk.tokenize import word_tokenize


def extract_names(document):
    nouns = [] #empty to array to hold all nouns
    
    stop = stopwords.words('english')
    stop.append("Resume")
    stop.append("RESUME")
    document = ' '.join([i for i in document.split() if i not in stop])
    sentences = nltk.sent_tokenize(document)
    for sentence in sentences:
        for word,pos in nltk.pos_tag(nltk.word_tokenize(str(sentence))):
            if (pos == 'NNP' and len(word)>2):
                nouns.append(word)
    nouns=' '.join(map(str,nouns))
    nouns=nouns.split()                
    return nouns         

def extract_email_addresses(text):
    r = re.compile(r'[\w\.-]+@[\w\.-]+')
    return r.findall(text)

def extract_mobile_number(text):
    #mno = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,15}[0-9]', text)
    mno = re.findall(r'[\+\(]?[1-9][0-9 \-\(\)]{8,15}[0-9]', text)
    mono = []
    for i in range(len(mno)):
        digit = 0
        for j in mno[i]:
            if j.isnumeric():
                digit+=1
        if digit > 9 and digit < 15:
            mono.append(mno[i]) 
        
    return mono


def extract_skill_set(text,skill_set):   
    f=[]
    for s in skill_set:
        if re.search(s, text, re.I):
        #if s in text:
            if len(s)>1:
                #i have added the below line to remove blank space after the skill
                #s = s.replace(' ','')
                f.append(s)
    return [x.lower() for x in f]

def experience(fullText):
    h=fullText.split()               #look at h only years get it
    if 'years' in h and 'months' in h:
        d=h[h.index('years')-1] + "." + h[h.index('months')-1]
    elif 'years' in h:
        d = h[h.index('years')-2] +' '+ h[h.index('years')-1]
    elif 'months' in h:
        d = '.'+ h[h.index('months')-1]
    else :
        d = '0'
    return(d)

def clean_exp(txt):
    h = txt
    h = re.sub('[^\d\. ]','',h)
    return(h)

def generate_ngrams(filename, n):
    
    words = filename.split()
    output = []  
    for i in range(len(words)-n+1):
        output.append(words[i:i+n])
    f=[]            
    for i in output:
        if 'years' in i:
            f.append(output[output.index(i)])
            if len(f)==1:
                n=f[0][0]
                n=n + " " + "years"
                break
    
    if len(f)<1:
        n='fresher'
    return n


def exper(fullText):
    mi=fullText.lower()
    #print(mi)
    h=mi.replace("_"," ")
    h=h.replace("-"," ")
    h=h.replace("+"," ")
    h=h.replace("year","years")
    h=h.replace(","," ")
    h=h.replace("("," ")
    h=h.replace(")"," ")
    h=h.replace(".docx"," ")
    h=h.replace(".pdf"," ")
    h=h.split()              #look at h only years get it
    if 'years' in h and 'months' in h:
        d=h[h.index('years')-1] + " " + h[h.index('years')]+ " " +h[h.index('months')-1] + " " +h[h.index('months')]
    elif 'months' in h:
        d=h[h.index('months')-1] + " " + h[h.index('months')]
    elif 'month' in h:
        d=h[h.index('month')-1] + " " + h[h.index('month')]
    elif 'year' in h:
        d=h[h.index('year')-1] + " " + h[h.index('year')]
    elif 'years' in h:
        d=h[h.index('years')-1] + " " + h[h.index('years')]
    elif re.search('no experience',str(h),re.M|re.I) :
        d='fresher'
    elif re.search('fresher',str(h),re.M|re.I) :
        d='fresher'
    else:
        d=generate_ngrams(fullText, 2)  
    return d    

#------------------------------------------------------------------------------#

def final_candidate_df(path,skill_path):
    with open(skill_path,"r") as skill:
        skill_set = skill.read().split("\n") 
    skill_set = [x.lower() for x in skill_set]  
    data = []
    for filename in os.listdir(path):
        if filename.endswith(('.doc','.docx')):
            print(filename)
            DOC_FILEPATH = path + filename        
            docu = win32com.client.GetObject(DOC_FILEPATH)
            res = docu.Range().Text
            del docu
            
        elif filename.endswith('.pdf'):
            print(filename) 
            parsed = parser.from_file(path + filename)      
            res = parsed["content"]
            
          
        #res1 = res.replace("\\n"," ")
        res1 = res
        name_coll = extract_names(res1)
        ab=[];name=[]
        c=set(extract_email_addresses(res))
        for i in name_coll :
            if re.search(i, str(c),re.M|re.I) or re.search(i,filename,re.M|re.I) :
                ab.append(i)
                if len(ab)==1:
                    break
                
        
        res1 = res1.replace('b"',"")
        abc = res1.split()
        
        for i in name_coll :
            if abc[abc.index(ab[0])+2] in name_coll:
                name = abc[abc.index(ab[0])] + " " + abc[abc.index(ab[0])+1] + " " + abc[abc.index(ab[0])+2]       
            
            elif abc[abc.index(ab[0])+1] in name_coll:
                name = abc[abc.index(ab[0])] + " " + abc[abc.index(ab[0])+1]
                
            else:
                name = abc[abc.index(ab[0])] 
                
        
        email = extract_email_addresses(res)
        cno = extract_mobile_number(res)
        skills = extract_skill_set(res,skill_set)
        exp= exper(res)
        
        data.append({"FileName":filename, "FileContents":res, "Name":name, "Email Address":email, "Contact Number":cno, "Skills":skills, "Experience": exp})
            
    df = pd.DataFrame(data, columns = ["FileName","FileContents","Name","Email Address","Contact Number","Skills","Experience"])
    df['final'] = df['Experience'].apply(lambda x : experience(x))
    df['final'] = df['final'].apply(lambda x : clean_exp(x))
    df['final'] = pd.to_numeric(df['final'])
    return df
#nltk.download('averaged_perceptron_tagger')




