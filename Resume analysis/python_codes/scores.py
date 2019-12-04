# -*- coding: utf-8 -*-
"""
Created on Tue Dec  3 19:56:10 2019

@author: Akmal
"""

import sys
import pandas as pd
user_path = ''
sys.path.insert(0, user_path+'Resume analysis\\python_codes')
import os
os.chdir(user_path+'\\Resume analysis')
import final_discription as discription
import final_resume as f

stud = f.final_candidate_df(user_path+'Resume analysis\\all resume\\n2\\',user_path+"Resume analysis\\Data\\java_skills.txt")

d = discription.final_df(user_path+'Resume analysis\\dis\\')
df_weights = pd.read_csv(user_path+'Resume analysis\\Data\\skill_weights.csv')
df_weights.set_index('skils',inplace = True)


def candidate_suitability(dis, cand):
    def exp_range(fromm, to, exp):
        if pd.isnull(to):
            if exp >= fromm:
                return 1
            else:
                return 0
        else :
            if (exp>=fromm) & (exp<=to):
                return 1
            else:
                return 0

    def match_skills(dis,stu_skills):
        l = len(dis)
        msl = [x for x in dis if x in stu_skills]
        prob_msl = len(msl)/l
        ext_s = [x for x in stu_skills if x not in msl]
        scores = [df_weights.loc[x,'weights'] for x in ext_s]
        prob_msl += sum(scores)
        return prob_msl
    skills = match_skills(dis['skill_set'],cand['Skills'])
    exp = exp_range(dis['from'], dis['to'], cand['final'])
    print(exp)
    tot = (skills + exp)/2
    return tot

stud['score'] = stud.apply(lambda x:candidate_suitability(d.iloc[0,2:], x),axis = 1)

stud[['Name','score']]















