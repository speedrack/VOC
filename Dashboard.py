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
    tab1, tab2, tab3 = st.tabs(['리뷰 수' , '별점', '이슈사항'])
    
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
        col1.metric("홈던트하우스", f"{latest_list[0]}" +"/" + f"{ratios[0]: .1%}", f"{subtracts[0]}" + "_" + f"{ratios_delta[0]: .2%}")
        col2.metric("스피드랙", f"{latest_list[1]}" +"/" + f"{ratios[1]: .1%}", f"{subtracts[1]}" + "_" + f"{ratios_delta[1]: .2%}")
        
        col3, col4 = st.columns(2)
        col3.metric("슈랙", f"{latest_list[2]}" +"/" + f"{ratios[2]: .1%}", f"{subtracts[2]}" + "_" + f"{ratios_delta[2]: .2%}")
        col4.metric("피피랙", f"{latest_list[3]}" +"/" + f"{ratios[3]: .1%}", f"{subtracts[3]}" + "_" + f"{ratios_delta[3]: .2%}")



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



    
    # 긍부정 주제
    with tab3:     
        with st.expander('17W 이슈 및 조치사항'):
            st.write('배송 관련')

            temp1 = pd.DataFrame([{'아이디': 'songe0606',
                                  '리뷰': '화요일 도착햇다고 하는데 배송이 되지 않아서 업체측에 전화를 했어요 보통 죄송하다고 하고 업체에서 해결 후 연락주시던데 이곳은 사과도 없고 소비자에게 택배기사님이랑 통화해보라며 연락처를 알려주시더라구요 친절하게 알려주시긴했는데 시스템이 이게 맞나 생각되었습니다 그래도 기사님께서 아주 친절하고 발빠르게 해결해주셔서 감사했습니다 그리고 선반은 블랙으로 주문했는데 화이트가 왔어요 조금 화가 나지만 옷을 잔뜩 쌓아두고 얼른 정리해야해서 그냥 사용하겠습니다 덕분에 까만 프레임에 하얀 선반 행거를 갖게 되었습니다 ^^',
                                  '브랜드': '스피드랙'
                                  }])
            temp2 = pd.DataFrame([{'아이디': '댕글마마',
                                  '리뷰': '조립설명서도 없는 스피드랙선반이라니요..혹시나하고 들어온 상세설명에도 없네요;; 오늘의집 설치기사도 멋대로 취소하고 반품비는 2만원인데 오늘의집포인트 만원 준다니 어이가 없고 그냥 설치하려고 박스 개봉하니 조립설명서도 없다니 정말 최악입니다.',
                                  '브랜드': '스피드랙'
                                  }])
            temp3 = pd.DataFrame([{'아이디': ' ',
                                  '리뷰': '제품 여러개 시키면 무거워서 배송기사님들이 물건 다 갖다 던지나보네요. 한개도 아니고 박스 세개나 터져있네요. 얼마전에 하나 시켰을땐 깔끔하게 잘와서 사이즈 안맞아도 반품도 안시키고 그냥 새로주문한건데 실망스럽습니다.',
                                  '브랜드': '스피드랙'
                                  }])
            temp4 = pd.DataFrame([{'아이디': '꿀혀니으집',
                      '리뷰': '솔직히 이쁘진 않지만 컴팩트함에 사용하게되었어여 택배기사님이 욕하셔서 기분이 안좋네요',
                      '브랜드': '스피드랙'
                      }])

            st.dataframe(temp1)
            st.write('''택배 기사님 오배송으로 인해 문의글로 인입이되어, 배송기사님과 먼저 통화\n\n\n-> 기사님께서 본인의 명확한 실수이기 때문에 고객님과 직접 통화를 원하신다고 하여, 인입시 고객님께 즉시 기사님 번호 전달 \n\n\n(선사과를 하지 않은 부분에 대해서 실무자분과 피드백 진행)''')
            st.write('\n\n\n')
            
            st.dataframe(temp2)
            st.write("CRM에 해당 고객과 통화 내역이 없는 것으로 보아 오늘의 집 고객센터로 문의한 것으로 판단, 조립설명서에 대해 아웃바운드 예정")
            st.write('\n\n\n') 

            st.dataframe(temp3)
            st.write("4.26(금) 부분 교환 및 커피 쿠폰을 말씀드렸으나,  1세트에 대해 반품원하셔서 진행\n\n\n-> 물류팀측으로 전달하여 경북 영주시 CJ영업소와 통화하여 재발되지 않도록 요청 완료 ")
            st.write('\n\n\n') 

            st.dataframe(temp4)
            st.write("아웃바운드 예정 (상황 파악 필요)")
            st.write('\n\n\n') 
            
        
        with st.expander('이전사항'):
            df_issue = pd.DataFrame([
                                {'주차': '15w',
                                '이슈': '배송 - 1층 방치 (슈랙)',
                                '문서링크':'',
                                 '비고':''},
                                
                                {'주차': '16w',
                                '이슈': '배송 - 현관문 막음 (홈던트하우스)',
                                '문서링크':'',
                                 '비고': '고객과 통화 및 사과 완료. 택배기사와 연락하여 재발 방지 요구할 예정'}       
                                 ])
            
            st.dataframe(df_issue)
            # edited_df = st.data_editor(df_issue, num_rows="dynamic", column_config={"문서링크": st.column_config.LinkColumn("문서링크")})
        

