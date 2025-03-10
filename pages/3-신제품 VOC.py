# -*- coding: utf-8 -*-
"""
Created on Tue Sep  3 08:34:56 2024

@author: speed
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import date
from datetime import timedelta
import streamlit.components.v1 as components
import os
import re



@st.cache_data
def load_new_review(year, week):
    df = pd.read_excel(fr"year/{year}/{week}/VOC {week} 신제품 원본.xlsx")
    df['등록일'] = df['등록일'].astype(str)
    df['등록일'] = df['등록일'].str.replace(' 00:00:00', '')
    df['평점'] = round(df['평점']).astype(int)
    return df




def newproduct_df(df, keyword):
    # 특정 컬럼별로 다른 너비 설정
    column_config = {
        "리뷰": st.column_config.TextColumn(width=1000),
        "URL": st.column_config.LinkColumn()
    }


    try:
        st.dataframe(df, hide_index=True, column_config=column_config)
    except:
        st.write('...')
    
    

def extract_week_number(week):
    # 숫자만 추출해서 정수로 변환
    match = re.search(r'(\d+)', week)
    return int(match.group()) if match else 0





if __name__ == '__main__':
    st.set_page_config(layout="wide")


    # 연도, 주차 선택
    year_dir = 'year'
    yearlist = os.listdir(year_dir)
    yearlist = sorted(yearlist, reverse=True)
    year_selected = st.sidebar.selectbox('연도를 선택하세요.', yearlist, index=0)
    

    
    week_dir = os.path.join(year_dir, year_selected)
    if os.path.exists(week_dir):
        weeklist = os.listdir(week_dir)
        # 숫자를 기준으로 내림차순 정렬
        weeklist = sorted(weeklist, key=extract_week_number, reverse=True)
    else:
        weeklist = []
    
    week_selected = st.sidebar.selectbox('주차를 선택하세요.', weeklist, index=0)   

    try:    
        df_new = load_new_review(year_selected, week_selected)
    except:
        df_new = '...'
    


    # 특정 컬럼별로 다른 너비 설정
    column_config = {
        "리뷰": st.column_config.TextColumn(width=1000),
        "URL": st.column_config.LinkColumn()
    }

    st.title('신제품 VOC')
    
    try:
        st.dataframe(df_new, hide_index=True, height=600, column_config=column_config)
    except:
        st.write('...')
        
        


        
        
            
            
            
            
