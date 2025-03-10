import streamlit as st
import pandas as pd
import os
import re

@st.cache_data
def load_review(year, week):
    data = pd.read_excel(fr"year/{year}/{week}/VOC {week} 원본_3점.xlsx")
    return data

@st.cache_data
def neg_summary(year, week, brand):
    file = fr"year/{year}/{week}/neg/neg_{brand}.txt"
    f = open(file, 'r', encoding='UTF-8')
    txt = f.read()
    
    return txt

@st.cache_data
def keyvalue_summary(year, week, brand):
    try:
        file = fr"year/{year}/{week}/topic/주제_{brand}.txt"
        f = open(file, 'r', encoding='UTF-8')
        txt = f.read()
    except:
        txt = '...'
        
    return txt

@st.cache_data
def notable_summary(year, week, brand):
    try:
        file = fr"year/{year}/{week}/특이사항/특이사항_{brand}.txt"
        f = open(file, 'r', encoding='UTF-8')
        txt = f.read()
    except:
        txt = '...'
        
    return txt


def extract_week_number(week):
    # 숫자만 추출해서 정수로 변환
    match = re.search(r'(\d+)', week)
    return int(match.group()) if match else 0



if __name__ == '__main__':
    
    #st.caption('GPT4 turbo로부터 생성됨.')
    
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

    
    
    # 주차별 df 로드
    df = load_review(year_selected, week_selected)
    
    # 브랜드 선택
    brand_selected = st.sidebar.selectbox('브랜드를 선택하세요.', ['홈던트하우스', '스피드랙', '슈랙', '피피랙', '스피드랙MAX'])
    brand_df = df.loc[df['브랜드']==brand_selected]
    
    
    
    
    # 대시보드 타이틀
    st.title(f'{brand_selected} VOC 분석')
    
    tab1, tab2 = st.tabs(['개선 제안' , '3점 이하'])#, '주제별 요약'])
    
    # 개선 제안 리뷰
    with tab1:
        st.subheader("개선 제안 리뷰")
        brand_notable_summary = notable_summary(year_selected, week_selected, brand_selected)
        st.write(brand_notable_summary)
        

    # 3점 이하 불만 리뷰 요약 및 df 출력
    with tab2: 
        st.subheader("3점 이하 리뷰 요약")
        
        brand_neg_summary = neg_summary(year_selected, week_selected, brand_selected)
        st.write(brand_neg_summary)
    
        
        st.write('\n\n')
        
        if st.button('리뷰 원본 보기', type='primary'):
            neg_df = brand_df.loc[brand_df['평점'] <= 3]
            neg_df['평점'] = round(neg_df['평점'], 0)
            neg_df = neg_df.reset_index(drop=True)
            
            st.dataframe(data=neg_df)
    
        
    # 주제별 요약
    # with tab3: 
    #     st.subheader("주제별 요약")
    #     brand_keyvalue_summary = keyvalue_summary(week_selected, brand_selected)
    #     st.write(brand_keyvalue_summary)
        
        
