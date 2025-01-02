import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np


@st.cache_data
def load_data():
    data = pd.read_excel(r"통계.xlsx", sheet_name=None)
    return data

@st.cache_data
def load_topic():
    data = pd.read_excel(r"topic.xlsx", sheet_name=None)
    return data

@st.cache_data
def draw_chart(data, sheet, recent=False):
    """
    데이터를 기반으로 그래프를 그리는 함수.

    Args:
        recent (bool): True일 경우 최근 12개의 데이터만 표시.

    Returns:
        fig: Plotly 그래프 객체.
    """
    # 연도와 주차 조합
    weeks = data[sheet]['year'].astype(str) + "." + data[sheet]['주차별'].astype(str)
    
    hh = data[sheet]['홈던트하우스']
    speed = data[sheet]['스피드랙']
    shu = data[sheet]['슈랙']
    pp = data[sheet]['피피랙']
    
    # 최근 12개만 선택
    if recent:
        weeks = weeks.tail(12)
        hh = hh.tail(12)
        speed = speed.tail(12)
        shu = shu.tail(12)
        pp = pp.tail(12)
    
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
    fig.update_layout(
        xaxis_title='주차',
        yaxis_title=sheet,
        title="최근 12주 데이터" if recent else "전체 데이터"
    )

    return fig
    


def cal_reviewNum(latest, previous):
    latest_list = df_reviewnum.loc[df_reviewnum['year-week'] == latest].sum()[2:6]
    previous_list = df_reviewnum.loc[df_reviewnum['year-week'] == previous].sum()[2:6]
      
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
    topic = load_topic()
    
    # 대시보드 타이틀
    st.title('VOC 대시보드')
    tab1, tab2 = st.tabs(['리뷰 수' , '별점'])#, '주요 주제'])
    
    # 리뷰 수
    with tab1:
        
        ## 전주 대비 변화
        
        # 리뷰 수 상세
        df_reviewnum = data['리뷰수']
        df_reviewnum['year-week'] = (df_reviewnum['year'].astype(str) + "." + df_reviewnum['주차별'].astype(str))
        latest = df_reviewnum['year-week'].unique()[-1]
        previous = df_reviewnum['year-week'].unique()[-2]
        
        latest_list, previous_list, subtracts = cal_reviewNum(latest, previous)
        latest_sum = sum(latest_list)
        previous_sum = sum(previous_list)
        
        
        st.subheader('전주 대비 변화')    
        ratios, ratios_delta = cal_ratio(latest_list, previous_list)
        st.metric(f"{latest} 총 리뷰 수", f"{latest_sum}", f"{latest_sum - previous_sum}")

        col1, col2 = st.columns(2)
        col1.metric("홈던트하우스", f"{latest_list[0]}" +" /" + f"{ratios[0]: .1%}", f"{ratios_delta[0]: .2%}" + "_" + f"{subtracts[0]}")
        col2.metric("스피드랙", f"{latest_list[1]}" +" /" + f"{ratios[1]: .1%}", f"{ratios_delta[1]: .2%}" + "_" + f"{subtracts[1]}")

        
        col3, col4 = st.columns(2)
        col3.metric("슈랙", f"{latest_list[2]}" +" /" + f"{ratios[2]: .1%}", f"{ratios_delta[2]: .2%}" + "_" + f"{subtracts[2]}")
        col4.metric("피피랙", f"{latest_list[3]}" +" /" + f"{ratios[3]: .1%}", f"{ratios_delta[3]: .2%}" + "_" + f"{subtracts[3]}")


        st.divider()
        
        
        
        # 리뷰 수 그래프
        st.write('\n\n')
        st.subheader('주차별 리뷰 수')
        chart_reviewnum = draw_chart(data, '리뷰수', recent=True)
        st.plotly_chart(chart_reviewnum, use_container_width=True)
        
        with st.expander("전체 기간 그래프 보기"):
            chart_reviewnum_total = draw_chart(data, '리뷰수', recent=False)
            st.plotly_chart(chart_reviewnum_total, use_container_width=True)




    # 별점
    with tab2:
        
        # 평균 별점 그래프
        st.write('\n\n')
        st.subheader('주차별 평균 별점')
        chart_avg = draw_chart(data, '평균', recent=True)
        st.plotly_chart(chart_avg, use_container_width=True)
        
        with st.expander("전체 기간 그래프 보기"):
            chart_avg_total = draw_chart(data, '평균', recent=False)
            st.plotly_chart(chart_avg_total, use_container_width=True)

        
        st.divider()
        
        
        # 그룹 누적 막대 그래프
        st.subheader('주차/브랜드별 별점 상세')
        df_numdetail = data['별점세부'].copy()
        # year와 week를 결합하여 정렬
        df_numdetail['year_week'] = (df_numdetail['year'].astype(str) + "." + df_numdetail['week'].astype(str))

        # 정렬 기준 생성 (연도와 주차를 고려하여 최신순 정렬)
        sort_numdetail = sorted(
            df_numdetail['year_week'].unique(),
            reverse=True)
    
        # 멀티셀렉트 박스
        week_selected = st.multiselect('주차들을 선택하세요.', sort_numdetail, default=sort_numdetail[0])
        
        
        df_numdetail['scores'] = df_numdetail['scores'].astype('str')
        if week_selected:
            df_numdetail = df_numdetail.loc[df_numdetail['year_week'].isin(week_selected)]
            fig_numdetail = px.bar(df_numdetail, x="brand", y='N', hover_data=['ratio'], facet_col="week", color="scores")
            st.plotly_chart(fig_numdetail, use_container_width=True)


        st.dataframe(data['별점세부'].astype(str), use_container_width=True, hide_index=True)
        
        
        
        
        
    # with tab3:
    #     df = load_topic()
    #     brands = ['홈던트하우스', '스피드랙', '슈랙', '피피랙']

    #     # 브랜드 선택
    #     selected_brands = st.multiselect(
    #         '브랜드 선택',
    #         options=brands,
    #         default=brands[:2])
        
    #     # 색상 팔레트 설정
    #     color_palette = px.colors.qualitative.Plotly
        
    #     # 각 선택된 브랜드에 대해 expander 생성
    #     for brand in selected_brands:
    #         with st.expander(f"{brand} 주제별 수치", expanded=False):
    #             # 해당 브랜드의 데이터만 필터링
    #             brand_df = df[brand]
                
    #             # 주차 선택 위젯
    #             sort_numdetail = sorted(brand_df['주차별'].unique(), reverse=True)
                
    #             selected_weeks = st.multiselect(
    #                 f'주차 선택',
    #                 options=sort_numdetail,
    #                 default=sort_numdetail[:2],  
    #                 key=f'week_select_{brand}'  # 각 브랜드마다 고유한 키 사용
    #             )
                
    #             # 상위 3개 항목 그래프
    #             fig_top = go.Figure()
                
    #             for i, week in enumerate(selected_weeks):
    #                 week_data = brand_df[brand_df['주차별'] == week]
                    
    #                 fig_top.add_trace(go.Bar(
    #                     y=['견고함', '조립용이성', '디자인'],
    #                     x=[week_data['견고함'].values[0], week_data['조립용이성'].values[0], week_data['디자인'].values[0]],
    #                     orientation='h',
    #                     name=f'{week}주차',
    #                     marker_color=color_palette[i % len(color_palette)]
    #                 ))
                
    #             fig_top.update_layout(
    #                 title='긍정 주제',
    #                 xaxis_title='수치',
    #                 yaxis_title='주제',
    #                 legend_title='주차',
    #                 hovermode='y unified',
    #                 height=300,
    #                 barmode='group',
    #                 bargap=0.25,  # 막대 그룹 사이의 간격 증가
    #                 bargroupgap=0.1,  # 그룹 내 막대 사이의 간격
    #                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, traceorder='reversed')
    #             )
                
    #             fig_top.update_yaxes(categoryorder='array', categoryarray=['디자인', '조립용이성', '견고함'])
                
    #             st.plotly_chart(fig_top, use_container_width=True)
                
    #             # 하위 3개 항목 그래프
    #             fig_bottom = go.Figure()
                
    #             for i, week in enumerate(selected_weeks):
    #                 week_data = brand_df[brand_df['주차별'] == week]
                    
    #                 fig_bottom.add_trace(go.Bar(
    #                     y=['누락', '파손', '배송불만'],
    #                     x=[week_data['누락'].values[0], week_data['파손'].values[0], week_data['배송불만'].values[0]],
    #                     orientation='h',
    #                     name=f'{week}주차',
    #                     marker_color=color_palette[i % len(color_palette)]
    #                 ))
                
    #             fig_bottom.update_layout(
    #                 title='부정 주제',
    #                 xaxis_title='수치',
    #                 yaxis_title='주제',
    #                 legend_title='주차',
    #                 hovermode='y unified',
    #                 height=300,
    #                 barmode='group',
    #                 bargap=0.25,  # 막대 그룹 사이의 간격 증가
    #                 bargroupgap=0.1,  # 그룹 내 막대 사이의 간격
    #                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, traceorder='reversed')
    #             )
                
    #             fig_bottom.update_yaxes(categoryorder='array', categoryarray=['배송불만', '파손', '누락'])
                
    #             st.plotly_chart(fig_bottom, use_container_width=True)