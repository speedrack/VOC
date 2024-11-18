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
def load_review(week):
    data = pd.read_excel(fr"week/{week}/VOC {week} 원본.xlsx")
    return data




def filtering_df(df):
    
    filtered_df = df.copy()
    # 필터링 옵션
    col1, col2 = st.columns(2)
    
    with col1:
        브랜드_filter = st.selectbox('브랜드', ['전체'] + filtered_df['브랜드'].unique().tolist())
        
    with col2:
        평점_filter = st.selectbox('평점', ['전체'] + filtered_df['평점'].unique().tolist())
        
    
    
    # if 대분류_filter != '전체':
    #     filtered_df = filtered_df[filtered_df['대분류'] == 대분류_filter]
        
    # if 브랜드_filter != '전체':
    #     filtered_df = filtered_df[filtered_df['브랜드'] == 브랜드_filter]
    
        
        
    # st.dataframe(filtered_df, use_container_width=True, hide_index=True)







if __name__ == '__main__':
    st.set_page_config(layout="wide")

    
    # 주차 선택
    week_dir = 'week'
    weeklist = os.listdir(week_dir)
    weeklist = sorted(weeklist, reverse=True)
    week_selected = st.sidebar.selectbox('주차를 선택하세요.', weeklist, index=0)

    try:
        df = load_review(week_selected)
        df['등록일'] = df['등록일'].astype(str)
        df['등록일'] = df['등록일'].str.replace(' 00:00:00', '')
        df['평점'] = round(df['평점']).astype(int)
    except:
        df = '...'

    st.title(f'VOC 원본')
    st.caption('VOC 원본 확인 & 다운로드')


    # 또는 특정 컬럼별로 다른 너비 설정
    column_config = {
        "리뷰": st.column_config.TextColumn(width=1000),
        "URL": st.column_config.LinkColumn()
    }


    try:
        st.dataframe(df, hide_index=True, height=600, column_config=column_config)
    except:
        st.write(df)
    
    


        
        
            
            
            
            
