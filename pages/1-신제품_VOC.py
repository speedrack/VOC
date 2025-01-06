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




@st.cache_data
def load_review(year, week, product):
    df = pd.read_excel(fr"year/{year}/{week}/VOC {week} {product} 원본.xlsx")
    df['등록일'] = df['등록일'].astype(str)
    df['등록일'] = df['등록일'].str.replace(' 00:00:00', '')
    df['평점'] = round(df['평점']).astype(int)
    return df




def newproduct_df(df, keyword):
    
    new_df = df.loc[df['비고'] == keyword]
    
    
    # 특정 컬럼별로 다른 너비 설정
    column_config = {
        "리뷰": st.column_config.TextColumn(width=1000),
        "URL": st.column_config.LinkColumn()
    }


    try:
        st.dataframe(new_df, hide_index=True, height=600, column_config=column_config)
    except:
        st.write('...')
    
    







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
        weeklist = sorted(weeklist, reverse=True)  
    else:
        weeklist = []
    
    week_selected = st.sidebar.selectbox('주차를 선택하세요.', weeklist, index=0)    


    df_max = load_review(year_selected, week_selected, '경량랙')
    df_dress = load_review(year_selected, week_selected, '드레스룸')


    # 특정 컬럼별로 다른 너비 설정
    column_config = {
        "리뷰": st.column_config.TextColumn(width=1000),
        "URL": st.column_config.LinkColumn()
    }

    st.title('신제품 VOC')
    
    st.write('스피드랙 MAX(경량랙)')
    try:
        st.dataframe(df_max, hide_index=True, height=600, column_config=column_config)
    except:
        st.write('...')
        
        
    st.write('\n\n')

    st.write('홈던트하우스 드레스룸')
    try:
        st.dataframe(df_dress, hide_index=True, height=600, column_config=column_config)
    except:
        st.write('...')
    


        
        
            
            
            
            
