import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go



@st.cache_data
def load_data():
    data = pd.read_excel(r"통계.xlsx", sheet_name=None)
    return data

@st.cache_data
def draw_chart(sheet):
    weeks = data[sheet]['주차별']
    
    hh = data[sheet]['홈던트하우스']
    speed = data[sheet]['스피드랙']
    shu = data[sheet]['슈랙']
    pp = data[sheet]['피피랙']
    
    
    # 그래프 그리기
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=weeks, y=hh,
                    mode='lines+markers',
                    name='홈던트하우스'))
    
    fig.add_trace(go.Scatter(x=weeks, y=speed,
                    mode='lines+markers',
                    name='스피드랙'))
    
    fig.add_trace(go.Scatter(x=weeks, y=shu,
                    mode='lines+markers',
                    name='슈랙'))
    
    fig.add_trace(go.Scatter(x=weeks, y=pp,
                    mode='lines+markers',
                    name='피피랙'))
    
    
        
    # 그래프 설정
    fig.update_layout(xaxis_title='주차',
                      yaxis_title=sheet)

    return fig
    


def cal_reviewNum(latest, previous):
    latest_list = df_reviewnum.loc[df_reviewnum['주차별'] == latest].sum()[1:]
    previous_list = df_reviewnum.loc[df_reviewnum['주차별'] == previous].sum()[1:]
      
    subtracts = []
    brands = ['홈던트하우스', '스피드랙', '슈랙', '피피랙']
    for brand in brands:
        sub = latest_list[brand] - previous_list[brand]
        subtracts.append(sub)
    
    return latest_list, previous_list, subtracts




def cal_ratio(latest_list, previous_list):
    ratios = []
    ratios_delta = []

    for i in range(4):
        ratio = latest_list[i]/latest_sum
        ratio_previous = previous_list[i]/previous_sum
        ratio_delta = ratio - ratio_previous
        
        ratios.append(ratio)
        ratios_delta.append(ratio_delta)
    
    return ratios, ratios_delta



 

if __name__ == '__main__':
    data = load_data()
    
    
    # 대시보드 타이틀
    st.title('VOC 대시보드')
    tab1, tab2 = st.tabs(['리뷰 수' , '별점'])
    
    # 리뷰 수
    with tab1:
        
        # 리뷰 수 그래프
        st.write('\n\n')
        st.subheader('주차별 리뷰 수')
        chart_reviewnum = draw_chart('리뷰수')
        st.plotly_chart(chart_reviewnum, use_container_width=True)
        
        # 리뷰 수 상세
        df_reviewnum = data['리뷰수']
        latest = df_reviewnum['주차별'].unique()[-1]
        previous = df_reviewnum['주차별'].unique()[-2]
        
        latest_list, previous_list, subtracts = cal_reviewNum(latest, previous)
        latest_sum = latest_list.sum()
        previous_sum = previous_list.sum()
        
        
        st.divider()
        
        
        # 전주 대비 변화
        st.subheader('전주 대비 변화')    
        ratios, ratios_delta = cal_ratio(latest_list, previous_list)
        st.metric(f"{latest} 총 리뷰 수", f"{latest_sum}", f"{latest_sum - previous_sum}")

        col1, col2 = st.columns(2)
        col1.metric("홈던트하우스", f"{latest_list[0]}" +" /" + f"{ratios[0]: .1%}", f"{ratios_delta[0]: .2%}" + "_" + f"{subtracts[0]}")
        col2.metric("스피드랙", f"{latest_list[1]}" +" /" + f"{ratios[1]: .1%}", f"{ratios_delta[1]: .2%}" + "_" + f"{subtracts[1]}")

        
        col3, col4 = st.columns(2)
        col3.metric("슈랙", f"{latest_list[2]}" +" /" + f"{ratios[2]: .1%}", f"{ratios_delta[2]: .2%}" + "_" + f"{subtracts[2]}")
        col4.metric("피피랙", f"{latest_list[3]}" +" /" + f"{ratios[3]: .1%}", f"{ratios_delta[3]: .2%}" + "_" + f"{subtracts[3]}")




    # 별점
    with tab2:
        
        # 평균 별점 그래프
        st.write('\n\n')
        st.subheader('주차별 평균 별점')
        chart_reviewnum = draw_chart('평균')
        st.plotly_chart(chart_reviewnum, use_container_width=True)
        
        
        st.divider()
        
        
        # 그룹 누적 막대 그래프
        st.subheader('주차/브랜드별 별점 상세')
        df_numdetail = data['별점세부']
        sort_numdetail = sorted(df_numdetail['week'].unique(), reverse=True)
        week_selected = st.multiselect('주차들을 선택하세요.', sort_numdetail, default=sort_numdetail[0])
        
        df_numdetail['scores'] = df_numdetail['scores'].astype('str')
        if week_selected:
            df_numdetail = df_numdetail.loc[df_numdetail['week'].isin(week_selected)]
            fig_numdetail = px.bar(df_numdetail, x="brand", y='N', hover_data=['ratio'], facet_col="week", color="scores")
            st.plotly_chart(fig_numdetail, use_container_width=True)

        

