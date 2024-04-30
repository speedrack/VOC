import streamlit as st
import pandas as pd


@st.cache_data
def load_review(week):
    data = pd.read_excel(fr"week/{week}/VOC {week} 원본.xlsx")
    data['등록일'] = data['등록일'].astype(str)
    return data

@st.cache_data
def neg_summary(week, brand):
    file = fr"week/{week}/neg/neg_{brand}.txt"
    f = open(file, 'r', encoding='UTF-8')
    txt = f.read()
    
    return txt

@st.cache_data
def keyvalue_summary(week, brand):
    try:
        file = fr"week/{week}/topic/주제_{brand}.txt"
        f = open(file, 'r', encoding='UTF-8')
        txt = f.read()
    except:
        txt = '...'
        
    return txt




if __name__ == '__main__':
    
    # 대시보드 타이틀
    st.title('브랜드별 VOC 분석')
    
    # 주차 선택
    weeklist = ['16w', '17w']
    weeklist = sorted(weeklist, reverse=True)
    week_selected = st.sidebar.selectbox('주차를 선택하세요.', weeklist, index=0)

    # 주차별 df 로드
    df = load_review(week_selected)
    
    # 브랜드 선택
    brand_selected = st.sidebar.selectbox('브랜드를 선택하세요.', ['홈던트하우스', '스피드랙', '슈랙', '피피랙'])
    brand_df = df.loc[df['브랜드']==brand_selected]
    
    
    # 3점 이하 불만 리뷰 요약 및 df 출력
    st.write('\n\n')
    st.subheader(f"{brand_selected} 3점 이하 리뷰 요약")
    brand_neg_summary = neg_summary(week_selected, brand_selected)
    st.write(brand_neg_summary)

    st.caption('GPT4 turbo로부터 생성됨.')
    
    if st.button('해당 리뷰 원본 보기', type='primary'):
        neg_df = brand_df.loc[brand_df['평점'] <= 3]
        neg_df['평점'] = round(neg_df['평점'], 0)
        neg_df = neg_df.reset_index(drop=True)
        st.dataframe(data=neg_df)
    
    st.divider()
    
    
    # 주제별 요약
    st.write('\n\n')
    st.subheader(f"{brand_selected} 주제별 요약")
    brand_keyvalue_summary = keyvalue_summary(week_selected, brand_selected)
    st.write(brand_keyvalue_summary)
    
    st.divider()
    
    
    # 특이사항
    st.write('\n\n')
    st.subheader(f"{brand_selected} 특이사항")
    st.write('...')
