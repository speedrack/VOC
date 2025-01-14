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
#from html2image import Html2Image
#from PIL import Image





# 그래프에서 작년, 올해 주차 구분이 안되고 있음!







@st.cache_data
def load_data():
    df = pd.read_excel(r"VOC조치내역.xlsx", usecols=['등록일', '브랜드', '리뷰', '조치내용', '판매처', '대분류', '소분류', '부서', '완료여부'])
    df = df.loc[df['대분류'].notna()]
    
    return df


def cal_thisweekdf(df):
    today = date.today()
    diff = (today.weekday() - 6) % 7

    days = [today - timedelta(days=(diff+i)) for i in range(1, 8)]
    days = list(map(str, days))
    df_thisWeek = df.loc[df['등록일'].isin(days)].reset_index(drop=True)
    df_thisWeek = df_thisWeek.sort_values(by=['대분류', '소분류'], ignore_index=True)
    df_thisWeek['등록일'] = pd.to_datetime(df_thisWeek['등록일']).dt.date 
    
    df_otherWeek = df.loc[~(df['등록일'].isin(days))].reset_index(drop=True)
    
    return df_thisWeek, df_otherWeek


def PieChart(df):
    agg_data = df.groupby('대분류').agg({'리뷰': 'count', '소분류': lambda x: set(list(x))}).reset_index()
    
    # 호버 텍스트 생성 함수
    def create_hover_text(row):
        subcategories = ', '.join(sorted(row['소분류']))
        return f"대분류:<br>{row['대분류']}<br><br>소분류:<br>{subcategories}"
    
    # 호버 텍스트 생성
    agg_data['hover_text'] = agg_data.apply(create_hover_text, axis=1)
    
    # 도넛 차트 생성
    fig = go.Figure(data=[go.Pie(labels=agg_data['대분류'],
                                 values=agg_data['리뷰'],
                                 hole=0.2,
                                 hovertext=agg_data['hover_text'],
                                 hoverinfo='text')])  # 호버 정보를 텍스트로 제한
    
    fig.update_layout(
       # 전체 여백 조정 - 아래 여백을 크게 설정
       margin=dict(b=100))
    
    # 트레이스 업데이트
    fig.update_traces(textposition='outside',
                      textinfo='label+percent+value', 
                      textfont_color="black", 
                      textfont_size=15)
    
    
    
    # Streamlit에서 차트 표시
    st.plotly_chart(fig)


#############################################
def PieChart_subtopic(df, selected_topic):
    agg_data = df.groupby(['대분류', '소분류']).agg({'리뷰': 'count'}).reset_index()
    df_subtopic = agg_data.loc[agg_data['대분류'] == f'{selected_topic}'].reset_index()
    
    
    # 도넛 차트 생성
    fig = go.Figure(data=[go.Pie(labels=df_subtopic['소분류'],
                                 values=df_subtopic['리뷰'],
                                 hole=0.2,
                                 hoverinfo='text')])  # 호버 정보를 텍스트로 제한
    
    # 트레이스 업데이트
    fig.update_traces(textposition='outside',
                      textinfo='label+percent+value', 
                      textfont_color="black", 
                      textfont_size=15)
    
    # 레이아웃 업데이트
    #st.subheader(f'{selected_topic}')
    
    # Streamlit에서 차트 표시
    st.plotly_chart(fig)


def cal_new_category(df_thisWeek, df_otherWeek):
    df_thisWeek_dedup = df_thisWeek.drop_duplicates(subset=['대분류', '소분류'])
    df_otherWeek_dedup = df_otherWeek.drop_duplicates(subset=['대분류', '소분류'])
    
    merged_df = pd.merge(df_thisWeek_dedup, df_otherWeek_dedup, 
                         on=['대분류', '소분류'], 
                         how='outer', 
                         indicator=True)

    # thisweek에만 있는 행 필터링
    new_categories = merged_df[merged_df['_merge'] == 'left_only']

    # 결과에서 필요한 열만 선택
    new_categories = new_categories[['대분류', '소분류']]
    
    return new_categories


def create_html_content(new_categories):
    html_content = """
    <style>
        .new-category {
            background-color: #f0f8ff;
            border-left: 5px solid #ff6347;
            padding: 10px;
            margin-bottom: 10px;
            font-family: Arial, sans-serif;
        }
        .new-tag {
            background-color: #ff6347;
            color: white;
            padding: 2px 5px;
            border-radius: 3px;
            font-size: 0.8em;
            margin-right: 5px;
        }
        .category-title {
            font-weight: bold;
            color: #333;
        }
        .subcategory {
            color: #333;
            margin-left: 60px;
        }
    </style>
    """
    
    for _, row in new_categories.iterrows():
        html_content += f"""
        <div class="new-category">
            <span class="new-tag">NEW</span>
            <span class="category-title">{row['대분류']}</span>
            <div class="subcategory">{row['소분류']}</div>
        </div>
        """
    
    return html_content
         
   
def create_metric(df, df_thisWeek):
    df_metric = df['대분류'].value_counts()
    df_thisWeek_metric = df_thisWeek['대분류'].value_counts()
    
    num_columns = 4
    columns = st.columns(num_columns)
    c = 0
    
    for i, v in df_metric.items():
        try:
            m = df_thisWeek_metric[i]
        except:
            continue
        
        col_index = c % num_columns
        with columns[col_index]:
            st.metric(i, int(v), int(m))
        
        c += 1




def create_subtopic_metric(df, df_thisWeek, selected_topic):
    df_metric = df.groupby(['대분류', '소분류']).size()
    df_metric = df_metric.loc[f'{selected_topic}']
    
    df_thisWeek_metric = df_thisWeek.groupby(['대분류', '소분류']).size()
    
    num_columns = 4
    columns = st.columns(num_columns)
    c = 0
    
    for i, v in df_metric.items():
        try:
            df_thisWeek_metric_temp = df_thisWeek_metric[f'{selected_topic}']
            m = df_thisWeek_metric_temp[i]
        except:
            continue
        
        col_index = c % num_columns
        with columns[col_index]:
            st.metric(i, int(v), int(m))
        
        c += 1







def create_graph_barLine(df):
    # 날짜 열을 datetime 형식으로 변환
    df['등록일'] = pd.to_datetime(df['등록일'])
    
    # 주별로 그룹화하여 대분류의 갯수 세기 (일요일~토요일 기준)
    df['주'] = df['등록일'].dt.to_period('W-SAT')
    weekly_counts = df.groupby(['주', '대분류']).size().reset_index(name='갯수')
    
    
    
    
    # '주' 열에서 주 번호와 연도를 추출하여 새로운 '주차' 열 생성
    weekly_counts['year'] = weekly_counts['주'].dt.year
    weekly_counts['week'] = weekly_counts['주'].dt.week
    weekly_counts['주차'] = weekly_counts['year'].astype(str) + "." + weekly_counts['week'].astype(str) + "w"

    
    
    # 꺾은선
    fig = px.line(weekly_counts, x='주차', y='갯수', color='대분류', markers=True,
                  title='주별 대분류(꺾은선)',
                  labels={'주차': '주차', '갯수': '갯수'})
    
    # Streamlit에 그래프 표시
    st.plotly_chart(fig)
                    
    
    
    # 누적막대
    fig = px.bar(weekly_counts, x='주차', y='갯수', color='대분류', 
                 title='주별 대분류(누적막대)', 
                 labels={'주차': '주차', '갯수': '갯수'},
                 barmode='stack')  # 누적 막대 그래프 설정
    
    # Streamlit에 그래프 표시
    st.plotly_chart(fig)









def create_subtopic_graph_barLine(df, selected_topic):
    df = df.loc[df['대분류'] == f'{selected_topic}']
    
    # 날짜 열을 datetime 형식으로 변환
    df['등록일'] = pd.to_datetime(df['등록일'])
    
    # 주별로 그룹화하여 대분류의 갯수 세기 (일요일~토요일 기준)
    df['주'] = df['등록일'].dt.to_period('W-SAT')
    weekly_counts = df.groupby(['주', '소분류']).size().reset_index(name='갯수')
    
    # '주' 열에서 주 번호 추출하여 새로운 '주차' 열 생성
    # weekly_counts['주차'] = weekly_counts['주'].dt.week
    weekly_counts['주차'] = weekly_counts.apply(lambda row: f"{row['주'].year}.{row['주'].week}W", axis=1)

    
    
    # 꺾은선
    fig = px.line(weekly_counts, x='주차', y='갯수', color='소분류', markers=True,
                  title='주별 소분류(꺾은선)',
                  labels={'주차': '주차', '갯수': '갯수'})
    
    # Streamlit에 그래프 표시
    st.plotly_chart(fig)
                    
    
    
    # 누적막대
    fig = px.bar(weekly_counts, x='주차', y='갯수', color='소분류', 
                 title='주별 소분류(누적막대)', 
                 labels={'주차': '주차', '갯수': '갯수'},
                 barmode='stack')  # 누적 막대 그래프 설정
    
    # Streamlit에 그래프 표시
    st.plotly_chart(fig)
















def filtering_df(df):
    filtered_df = df.loc[:, ['등록일', '판매처', '브랜드', '리뷰', '조치내용', '대분류', '소분류', '부서', '완료여부']]
    filtered_df['완료여부'] = filtered_df['완료여부'].fillna('미완료')
    
    # 필터링 옵션
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        대분류_filter = st.selectbox('대분류', ['전체'] + filtered_df['대분류'].unique().tolist())
        
    with col2:
        브랜드_filter = st.selectbox('브랜드', ['전체'] + filtered_df['브랜드'].unique().tolist())
    
    with col3:
        부서_filter = st.selectbox('부서', ['전체'] + filtered_df['부서'].unique().tolist())
    
    with col4:
        완료_filter = st.selectbox('완료여부', ['전체'] + filtered_df['완료여부'].unique().tolist())
    
    
    if 대분류_filter != '전체':
        filtered_df = filtered_df[filtered_df['대분류'] == 대분류_filter]
        
    if 브랜드_filter != '전체':
        filtered_df = filtered_df[filtered_df['브랜드'] == 브랜드_filter]
    
    if 부서_filter != '전체':
        filtered_df = filtered_df[filtered_df['부서'] == 부서_filter]
    if 완료_filter != '전체':
        filtered_df = filtered_df[filtered_df['완료여부'] == 완료_filter]    
        
        
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)



def filtering_subtopic_df(df, selected_topic):
    st.subheader(f'대분류: {selected_topic} 리뷰 확인')
    

    filtered_df = df.loc[:, ['등록일', '판매처', '브랜드', '리뷰', '조치내용', '대분류', '소분류', '부서', '완료여부']]
    filtered_df = filtered_df.loc[filtered_df['대분류'] == f'{selected_topic}']
    filtered_df['완료여부'] = filtered_df['완료여부'].fillna('미완료')
    
    
    # 필터링 옵션
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        소분류_filter = st.selectbox('소분류', ['전체'] + filtered_df['소분류'].unique().tolist(), key=1)
    
    with col2:
        브랜드_filter = st.selectbox('브랜드', ['전체'] + filtered_df['브랜드'].unique().tolist(), key=2)
    
    with col3:
        부서_filter = st.selectbox('부서', ['전체'] + filtered_df['부서'].unique().tolist(), key=3)
    
    with col4:
        완료_filter = st.selectbox('완료여부', ['전체'] + filtered_df['완료여부'].unique().tolist(), key=4)
    
    

    if 소분류_filter != '전체':
        filtered_df = filtered_df[filtered_df['소분류'] == 소분류_filter]
    if 브랜드_filter != '전체':
        filtered_df = filtered_df[filtered_df['브랜드'] == 브랜드_filter]
    if 부서_filter != '전체':
        filtered_df = filtered_df[filtered_df['부서'] == 부서_filter]
    if 완료_filter != '전체':
        filtered_df = filtered_df[filtered_df['완료여부'] == 완료_filter]    
        
        
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)



if __name__ == '__main__':
    st.set_page_config(layout="wide")
    df_raw = load_data()
    
    df = df_raw.copy()
    
    
    st.title('VOC 조치사항')
    tab1, tab2 = st.tabs(['이번주 조치사항' , '세부확인']) #, '완료 사항', '분류체계'])
    
    with tab1:
        df_thisWeek, df_otherWeek = cal_thisweekdf(df)
        new_categories = cal_new_category(df_thisWeek, df_otherWeek)
        
        
        st.subheader('이번주 조치사항') #'DB에서 읽어오고 일자 계산해서 보여주기(저번주 일~토)
        st.dataframe(df_thisWeek, use_container_width=True, 
                     column_config={
            "리뷰": st.column_config.Column(width=600),
            "조치내용": st.column_config.Column(width=400)})
        
        
        st.divider()
        
        
        ############### 좀 쌓이면 다시 살리기
        # st.write('\n\n')
        # st.subheader("새로운 카테고리")

        # if len(new_categories) == 0:ㄴ
        #     st.write('...')
        # else:
        #     html_content = create_html_content(new_categories)
        #     components.html(html_content, height=len(new_categories)*70)
            
        # st.divider()
        
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader('조치사항 대분류')
            PieChart(df)
            
            st.write('Metric')
            create_metric(df, df_thisWeek)
            
            
        with col2:
            create_graph_barLine(df)

        
        st.divider()
                
        
        with st.expander('조치사항 원본 확인'):
            filtering_df(df)
        
        
        st.caption('분류체계 마인드맵: https://xmind.ai/PN0RdI07')
            
        
        
    
    with tab2:
        selected_topic = st.selectbox('대분류 선택 or 입력', df['대분류'].unique())
        st.write('\n\n')
        st.write('\n\n')
        
        # 원본 리뷰 확인
        filtering_subtopic_df(df, selected_topic)
        
        
        st.divider()
        
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader(f'{selected_topic} 세부사항')
            PieChart_subtopic(df, selected_topic)
            
            st.write('Metric')
            create_subtopic_metric(df, df_thisWeek, selected_topic)

            
        with col2:
            create_subtopic_graph_barLine(df, selected_topic)

        
        st.divider()
        
        
            
            
            
            
            
    ############# 수정후에 반영하기        
    # with tab3:
    #     df_finished = df.loc[df['완료여부'].notna()]
        
    #     col1, col2 = st.columns(2)
    #     with col1:
    #         st.subheader('완료 사항 대분류')
    #         PieChart(df_finished)
    #     with col2:
    #         elected_topic = st.selectbox('대분류 선택 or 입력', df_finished['대분류'].unique())
    #         PieChart_subtopic(df_finished, selected_topic)
        
    #     st.write('2. 파이그래프2개 - 대, 소분류 / 3. 관련 원본 / 4. 완료항목 css로 표기? (새로운카테고리 코드)')
        
    #     st.write('5. 완료 주차(조치주차) 표기 / 6. 완료용 분류체계 마인드맵 넣기')
    #     st.write('지금 앞쪽에 완.미완 같이 있으니 여기에는 언제 어느주제에 무슨 조치를 했는지와 그에 따른 변화 보여주기? / 완료만 따로 보여주긴 해야함')
    #     st.write('완료 주제 리스트와 건수')
    #     # 완료된거
    #     df = df.loc[df['완료여부'].isna()]
        
    
    
    
    # with tab4:
    #     # Xmind AI 공유 링크
    #     xmind_link = "https://xmind.ai/share/PN0RdI07"
        
        
    #     # Xmind AI 콘텐츠 임베드
    #     #st.write('로딩 오류 시 새로고침 필요합니다.')
    #     #st.components.v1.iframe(xmind_link, height=900)
        
    #     hti = Html2Image()
    #     # URL을 이미지로 변환
    #     hti.screenshot(url=xmind_link, save_as='screenshot.png')
    #     time.sleep(5)
    
    #     # Streamlit에 이미지 표시
    #     image = Image.open('screenshot.png')
    #     st.image(image, caption='Website Screenshot', use_column_width=True)