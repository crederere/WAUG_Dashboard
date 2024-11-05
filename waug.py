import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime
import locale
import warnings

# 경고 메시지 숨기기
warnings.filterwarnings('ignore')

# 한국어 로케일 설정
try:
    locale.setlocale(locale.LC_ALL, 'ko_KR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Korean_Korea.949')
    except:
        pass

# Streamlit 설정
st.set_page_config(page_title="마케팅 대시보드", layout="wide")

# 스타일 설정
st.markdown("""
    <style>
    /* 스크롤바 제거 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    /* 전체 배경 색상 */
    .reportview-container {
        background: #f0f2f6;
    }
    /* 사이드바 배경 색상 */
    .sidebar .sidebar-content {
        background: #ffffff;
    }
    /* 헤더 텍스트 색상 */
    .css-18e3th9 {
        color: #2c3e50;
    }
    /* 데이터프레임 스타일링 */
    .dataframe {
        font-size: 12px !important;
    }
    /* 차트 여백 조정 */
    .plot-container {
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 마케팅 대시보드 및 자동 보고서 생성기")

# 파일 업로드
st.sidebar.header("📂 파일 업로드")
uploaded_file = st.sidebar.file_uploader("Excel 파일을 업로드하세요 ('raw 시트'와 'index 시트' 포함)", type=['xlsx', 'xls'])

def format_date_axis(fig, date_option):
    """날짜 축 포맷 설정 함수"""
    if date_option == '일별':
        dtick = 'D7'  # 7일 간격
        tickformat = '%Y-%m-%d'
    elif date_option == '주간':
        dtick = 'D7'
        tickformat = '%Y-%m-%d'
    else:  # 월별
        dtick = 'M1'
        tickformat = '%Y-%m'
    
    fig.update_xaxes(
        dtick=dtick,
        tickformat=tickformat,
        tickangle=45,
        tickmode='auto',
        nticks=10  # 최대 표시할 틱 수
    )
    return fig

def safe_division(x, y):
    """안전한 나눗셈 함수"""
    return np.where(y != 0, x / y, 0)

if uploaded_file:
    try:
        with st.spinner('데이터를 불러오는 중...'):
            # 1. 데이터 불러오기
            raw_df = pd.read_excel(uploaded_file, sheet_name='raw 시트')
            index_df = pd.read_excel(uploaded_file, sheet_name='index 시트')

            # 데이터 미리보기
            with st.expander("데이터 미리보기"):
                st.subheader("raw 시트 데이터")
                st.dataframe(raw_df.head())
                st.subheader("index 시트 데이터")
                st.dataframe(index_df.head())

            # 2. 데이터 전처리
            # 컬럼명 통일
            raw_df.columns = raw_df.columns.str.strip()
            index_df.columns = index_df.columns.str.strip()

            # 문자열 데이터 전처리
            object_cols_raw = raw_df.select_dtypes(include='object').columns
            object_cols_index = index_df.select_dtypes(include='object').columns
            
            for col in object_cols_raw:
                raw_df[col] = raw_df[col].astype(str).str.strip()
            for col in object_cols_index:
                index_df[col] = index_df[col].astype(str).str.strip()

            # 날짜 형식 변환
            raw_df['일별'] = pd.to_datetime(raw_df['일별'], errors='coerce')
            
            # 숫자형 컬럼 변환 및 오류 처리
            numeric_columns = ['노출수', '클릭수', '총비용(VAT포함,원)', '전환수', '전환매출액(원)', '평균노출순위']
            for col in numeric_columns:
                if col in raw_df.columns:
                    raw_df[col] = pd.to_numeric(raw_df[col].astype(str).str.replace(',', '').replace('[-+]', ''), errors='coerce')
                else:
                    st.error(f"'{col}' 컬럼이 'raw 시트'에 존재하지 않습니다.")
                    st.stop()

            # 캠페인 및 광고그룹 매칭
            campaign_index = index_df[['캠페인', '카테고리', '국가']].drop_duplicates()
            campaign_index.columns = ['캠페인', '캠페인카테고리', '캠페인국가']
            raw_df = pd.merge(raw_df, campaign_index, on='캠페인', how='left')

            # 필요한 지표 계산
            raw_df['CPC'] = safe_division(raw_df['총비용(VAT포함,원)'], raw_df['클릭수'])
            raw_df['CTR%'] = safe_division(raw_df['클릭수'], raw_df['노출수']) * 100
            raw_df['CPA'] = safe_division(raw_df['총비용(VAT포함,원)'], raw_df['전환수'])
            raw_df['CVR%'] = safe_division(raw_df['전환수'], raw_df['클릭수']) * 100
            raw_df['ROAS%'] = safe_division(raw_df['전환매출액(원)'], raw_df['총비용(VAT포함,원)']) * 100
            raw_df['ARPPU'] = safe_division(raw_df['전환매출액(원)'], raw_df['전환수'])

            # 3. 기간별 데이터 집계
            raw_df['주간'] = raw_df['일별'].dt.to_period('W').apply(lambda r: r.start_time)
            raw_df['월별'] = raw_df['일별'].dt.to_period('M').apply(lambda r: r.start_time)

            # 필터 옵션
            st.header("🔍 필터 옵션")
            date_col1, date_col2, date_col3 = st.columns([1, 2, 2])

            with date_col1:
                date_option = st.selectbox("기간 선택", ('일별', '주간', '월별'))

            with date_col2:
                start_date = st.date_input('시작 날짜', raw_df['일별'].min())

            with date_col3:
                end_date = st.date_input('종료 날짜', raw_df['일별'].max())

            # 사이드바 필터
            st.sidebar.header("📊 필터 옵션")

            # 캠페인 카테고리 필터
            unique_categories = sorted(raw_df['캠페인카테고리'].dropna().astype(str).unique())
            category_options = st.sidebar.multiselect(
                '캠페인 카테고리 선택',
                options=['전체'] + unique_categories,
                default=['전체']
            )

            # 전체 선택 처리
            if '전체' in category_options:
                category_options = unique_categories

            # 캠페인 필터
            unique_campaigns = sorted(raw_df[raw_df['캠페인카테고리'].isin(category_options)]['캠페인'].unique())
            campaign_options = st.sidebar.multiselect(
                '캠페인 선택',
                options=['전체'] + list(unique_campaigns),
                default=['전체']
            )

            # 전체 선택 처리
            if '전체' in campaign_options:
                campaign_options = unique_campaigns

            # 매체 필터
            media_options = st.sidebar.multiselect(
                '매체 선택 (PC/모바일)',
                options=['전체'] + list(raw_df['PC/모바일 매체'].dropna().astype(str).unique()),
                default=['전체']
            )

            # 전체 선택 처리
            if '전체' in media_options:
                media_options = raw_df['PC/모바일 매체'].dropna().astype(str).unique()

            # 데이터 필터링
            mask = (raw_df['일별'] >= pd.to_datetime(start_date)) & \
                   (raw_df['일별'] <= pd.to_datetime(end_date)) & \
                   (raw_df['캠페인카테고리'].isin(category_options)) & \
                   (raw_df['캠페인'].isin(campaign_options)) & \
                   (raw_df['PC/모바일 매체'].isin(media_options))

            filtered_df = raw_df.loc[mask]

            if filtered_df.empty:
                st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다. 필터 조건을 조정해주세요.")
                st.stop()

            # 그룹화
            group_df = filtered_df.groupby(date_option).agg({
                '총비용(VAT포함,원)': 'sum',
                '노출수': 'sum',
                '클릭수': 'sum',
                '전환수': 'sum',
                '전환매출액(원)': 'sum',
                '캠페인': 'nunique',
                '키워드': 'nunique'
            }).reset_index()

            # 지표 계산
            group_df['CPC'] = safe_division(group_df['총비용(VAT포함,원)'], group_df['클릭수'])
            group_df['CTR%'] = safe_division(group_df['클릭수'], group_df['노출수']) * 100
            group_df['CPA'] = safe_division(group_df['총비용(VAT포함,원)'], group_df['전환수'])
            group_df['CVR%'] = safe_division(group_df['전환수'], group_df['클릭수']) * 100
            group_df['ROAS%'] = safe_division(group_df['전환매출액(원)'], group_df['총비용(VAT포함,원)']) * 100
            group_df['ARPPU'] = safe_division(group_df['전환매출액(원)'], group_df['전환수'])

            # 4. 시각화
            st.header("📈 기간별 성과 지표")

            # 주요 지표 카드
            total_cost = group_df['총비용(VAT포함,원)'].sum()
            total_revenue = group_df['전환매출액(원)'].sum()
            total_clicks = group_df['클릭수'].sum()
            total_impressions = group_df['노출수'].sum()
            total_conversions = group_df['전환수'].sum()
            total_roas = safe_division(total_revenue, total_cost) * 100

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("총비용", f"₩{int(total_cost):,}")
            col2.metric("전환매출액", f"₩{int(total_revenue):,}")
            col3.metric("ROAS", f"{total_roas:.2f}%")
            col4.metric("클릭수", f"{int(total_clicks):,}")
            col5.metric("전환수", f"{int(total_conversions):,}")

            # 차트 생성
            st.subheader(f"📆 {date_option}별 차트")

            # 총비용 및 전환매출액 차트
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(
                x=group_df[date_option],
                y=group_df['총비용(VAT포함,원)'],
                name='총비용',
                line=dict(color='#1f77b4'),
                hovertemplate='%{x}<br>총비용: ₩%{y:,.0f}<extra></extra>'
            ))
            fig1.add_trace(go.Scatter(
                x=group_df[date_option],
                y=group_df['전환매출액(원)'],
                name='전환매출액',
                line=dict(color='#2ca02c'),
                hovertemplate='%{x}<br>전환매출액: ₩%{y:,.0f}<extra></extra>'
            ))
            
            fig1.update_layout(
                title=f'{date_option}별 총비용 및 전환매출액',
                xaxis_title=date_option,
                yaxis_title='금액 (원)',
                hovermode='x unified',
                showlegend=True
            )
            fig1 = format_date_axis(fig1, date_option)
            st.plotly_chart(fig1, use_container_width=True)

            # ROAS% 차트
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                x=group_df[date_option],
                y=group_df['ROAS%'],
                name='ROAS%',
                text=group_df['ROAS%'].round(2).astype(str) + '%',
                textposition='outside',
                hovertemplate='%{x}<br>ROAS: %{y:.2f}%<extra></extra>'
            ))
            
            fig2.update_layout(
                title=f'{date_option}별 ROAS%',
                xaxis_title=date_option,
                yaxis_title='ROAS (%)',
                hovermode='x unified',
                showlegend=False
            )
            fig2 = format_date_axis(fig2, date_option)
            st.plotly_chart(fig2, use_container_width=True)

            # 일별 데이터 표 생성
            st.subheader(f"📅 {date_option}별 데이터 표")
            
            # 데이터프레임 스타일링을 위한 함수
            def style_dataframe(df):
                return df.style.format({
                    '총비용(VAT포함,원)': '₩{:,.0f}',
                    '노출수': '{:,.0f}',
                    '클릭수': '{:,.0f}',
                    '전환수': '{:,.0f}',
                    '전환매출액(원)': '₩{:,.0f}',
                    'CPC': '₩{:,.0f}',
                    'CTR%': '{:,.2f}%',
                    'CPA': '₩{:,.0f}',
                    'CVR%': '{:,.2f}%',
                    'ROAS%': '{:,.2f}%',
                    'ARPPU': '₩{:,.0f}',
                    'avg.Imps': '{:,.1f}',
                    '운영비중': '{:,.2f}%'
                }).set_properties(**{
                    'text-align': 'right',
                    'font-size': '12px',
                    'padding': '5px'
                })

            styled_group_df = style_dataframe(group_df)
            st.dataframe(styled_group_df, height=400)

            # 5. 추가 분석
            st.header("🔎 추가 분석")

            # 캠페인유형별 분석
            if '캠페인유형' in filtered_df.columns:
                st.subheader("📊 캠페인유형별 분석")
                campaign_type_df = filtered_df.groupby('캠페인유형').agg({
                    '총비용(VAT포함,원)': 'sum',
                    '노출수': 'sum',
                    '클릭수': 'sum',
                    '전환수': 'sum',
                    '전환매출액(원)': 'sum'
                }).reset_index()

                # 지표 계산
                campaign_type_df['CPC'] = safe_division(campaign_type_df['총비용(VAT포함,원)'], campaign_type_df['클릭수'])
                campaign_type_df['CTR%'] = safe_division(campaign_type_df['클릭수'], campaign_type_df['노출수']) * 100
                campaign_type_df['CPA'] = safe_division(campaign_type_df['총비용(VAT포함,원)'], campaign_type_df['전환수'])
                campaign_type_df['CVR%'] = safe_division(campaign_type_df['전환수'], campaign_type_df['클릭수']) * 100
                campaign_type_df['ROAS%'] = safe_division(campaign_type_df['전환매출액(원)'], campaign_type_df['총비용(VAT포함,원)']) * 100
                campaign_type_df['운영비중'] = safe_division(campaign_type_df['총비용(VAT포함,원)'], campaign_type_df['총비용(VAT포함,원)'].sum()) * 100

                st.dataframe(style_dataframe(campaign_type_df))

                # 캠페인유형별 파이 차트
                fig_pie = go.Figure(data=[go.Pie(
                    labels=campaign_type_df['캠페인유형'],
                    values=campaign_type_df['총비용(VAT포함,원)'],
                    hovertemplate="캠페인유형: %{label}<br>총비용: ₩%{value:,.0f}<br>비중: %{percent}<extra></extra>"
                )])
                fig_pie.update_layout(title="캠페인유형별 비용 비중")
                st.plotly_chart(fig_pie, use_container_width=True)

            # 캠페인별 분석
            st.subheader("📊 캠페인별 분석")
            campaign_df = filtered_df.groupby('캠페인').agg({
                '총비용(VAT포함,원)': 'sum',
                '전환매출액(원)': 'sum',
                '노출수': 'sum',
                '클릭수': 'sum',
                '전환수': 'sum'
            }).reset_index()

            campaign_df['ROAS%'] = safe_division(campaign_df['전환매출액(원)'], campaign_df['총비용(VAT포함,원)']) * 100
            campaign_df['운영비중'] = safe_division(campaign_df['총비용(VAT포함,원)'], campaign_df['총비용(VAT포함,원)'].sum()) * 100
            campaign_df['CTR%'] = safe_division(campaign_df['클릭수'], campaign_df['노출수']) * 100
            campaign_df['CVR%'] = safe_division(campaign_df['전환수'], campaign_df['클릭수']) * 100

            # 정렬 옵션
            sort_options = ['총비용(VAT포함,원)', 'ROAS%', '전환매출액(원)', '운영비중']
            selected_sort = st.selectbox('정렬 기준 선택:', sort_options)
            campaign_df = campaign_df.sort_values(by=selected_sort, ascending=False)

            st.dataframe(style_dataframe(campaign_df))

            # Top 10 캠페인 차트
            fig_top10 = go.Figure()
            top10_campaigns = campaign_df.nlargest(10, '총비용(VAT포함,원)')
            
            fig_top10.add_trace(go.Bar(
                x=top10_campaigns['캠페인'],
                y=top10_campaigns['총비용(VAT포함,원)'],
                name='총비용',
                text=top10_campaigns['총비용(VAT포함,원)'].apply(lambda x: f'₩{x:,.0f}'),
                textposition='auto',
            ))
            
            fig_top10.update_layout(
                title='Top 10 캠페인 (비용 기준)',
                xaxis_title='캠페인',
                yaxis_title='총비용 (원)',
                xaxis_tickangle=45,
                height=500
            )
            st.plotly_chart(fig_top10, use_container_width=True)

            # 카테고리별 분석
            st.subheader("📊 카테고리별 분석")
            category_df = filtered_df.groupby('캠페인카테고리').agg({
                '총비용(VAT포함,원)': 'sum',
                '전환매출액(원)': 'sum',
                '노출수': 'sum',
                '클릭수': 'sum',
                '전환수': 'sum'
            }).reset_index()

            category_df['ROAS%'] = safe_division(category_df['전환매출액(원)'], category_df['총비용(VAT포함,원)']) * 100
            category_df['운영비중'] = safe_division(category_df['총비용(VAT포함,원)'], category_df['총비용(VAT포함,원)'].sum()) * 100
            category_df['CTR%'] = safe_division(category_df['클릭수'], category_df['노출수']) * 100
            category_df['CVR%'] = safe_division(category_df['전환수'], category_df['클릭수']) * 100

            st.dataframe(style_dataframe(category_df))

            # 카테고리별 파이 차트
            fig_category = go.Figure(data=[go.Pie(
                labels=category_df['캠페인카테고리'],
                values=category_df['총비용(VAT포함,원)'],
                hovertemplate="카테고리: %{label}<br>총비용: ₩%{value:,.0f}<br>비중: %{percent}<extra></extra>"
            )])
            fig_category.update_layout(title="카테고리별 비용 비중")
            st.plotly_chart(fig_category, use_container_width=True)

            # 6. 보고서 다운로드
            # 6. 보고서 다운로드
            st.header("📄 마케팅 성과 종합 보고서 다운로드")
            
            report_cols = st.columns([2, 1])
            with report_cols[0]:
                report_name = st.text_input("보고서 파일명", value="마케팅_종합보고서")
            with report_cols[1]:
                include_charts = st.checkbox("차트 포함", value=True, help="Excel 보고서에 차트를 포함합니다.")
            
            if st.button("📥 종합 보고서 다운로드"):
                try:
                    with st.spinner('상세 보고서를 생성하는 중...'):
                        buffer = BytesIO()
                        
                        # 날짜 문자열 변환 함수
                        def format_date_str(date_val, date_option):
                            if pd.isna(date_val):
                                return ''
                            if date_option == '일별':
                                return date_val.strftime('%Y-%m-%d')
                            elif date_option == '주간':
                                return date_val.strftime('%Y-%m-%d')
                            else:  # 월별
                                return date_val.strftime('%Y-%m')

                        # DataFrame 준비 및 날짜 처리
                        group_df_excel = group_df.copy().replace([np.inf, -np.inf], 0).fillna(0)
                        if date_option in group_df_excel.columns:
                            group_df_excel[date_option] = group_df_excel[date_option].apply(
                                lambda x: format_date_str(x, date_option)
                            )
                        

                        campaign_df_excel = campaign_df.copy().replace([np.inf, -np.inf], 0).fillna(0)
                        category_df_excel = category_df.copy().replace([np.inf, -np.inf], 0).fillna(0)
                        filtered_df_excel = filtered_df.copy().replace([np.inf, -np.inf], 0).fillna(0)
                        
                        # Excel Writer 생성
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            workbook = writer.book
                            
                            # 포맷 정의
                            title_format = workbook.add_format({
                                'bold': True,
                                'font_size': 14,
                                'align': 'center',
                                'valign': 'vcenter',
                                'bg_color': '#1F497D',
                                'font_color': 'white',
                                'border': 1
                            })
                            
                            subtitle_format = workbook.add_format({
                                'bold': True,
                                'font_size': 12,
                                'align': 'left',
                                'valign': 'vcenter',
                                'bg_color': '#D9E1F2',
                                'border': 1
                            })
                            
                            header_format = workbook.add_format({
                                'bold': True,
                                'font_size': 11,
                                'align': 'center',
                                'valign': 'vcenter',
                                'bg_color': '#D9E1F2',
                                'border': 1,
                                'text_wrap': True
                            })
                            
                            number_format = workbook.add_format({
                                'num_format': '#,##0',
                                'align': 'right',
                                'border': 1
                            })
                            
                            currency_format = workbook.add_format({
                                'num_format': '₩#,##0',
                                'align': 'right',
                                'border': 1
                            })
                            
                            percent_format = workbook.add_format({
                                'num_format': '0.00%',
                                'align': 'right',
                                'border': 1
                            })
                            
                            text_format = workbook.add_format({
                                'align': 'left',
                                'border': 1
                            })
                            
                            date_format = workbook.add_format({
                                'num_format': 'yyyy-mm-dd',
                                'align': 'center',
                                'border': 1
                            })

                            # 1. 개요 시트 생성
                            summary_sheet = workbook.add_worksheet('개요')
                            
                            # 제목
                            summary_sheet.merge_range('A1:H1', '마케팅 성과 종합 보고서', title_format)
                            summary_sheet.set_row(0, 30)
                            
                            # 기본 정보
                            summary_sheet.merge_range('A3:B3', '보고서 기본 정보', subtitle_format)
                            info_data = [
                                ['분석 기간', f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}"],
                                ['기간 단위', date_option],
                                ['캠페인 수', len(campaign_df)],
                                ['선택된 카테고리', ', '.join(category_options)],
                                ['선택된 매체', ', '.join(media_options)]
                            ]
                            for i, (label, value) in enumerate(info_data):
                                summary_sheet.write(i+3, 0, label, text_format)
                                summary_sheet.write(i+3, 1, value, text_format)
                            
                            # 주요 지표 요약
                            summary_sheet.merge_range('A9:B9', '주요 성과 지표', subtitle_format)
                            metrics_data = [
                                ['총 비용', total_cost, currency_format],
                                ['총 매출', total_revenue, currency_format],
                                ['ROAS', total_roas/100, percent_format],
                                ['총 노출수', total_impressions, number_format],
                                ['총 클릭수', total_clicks, number_format],
                                ['총 전환수', total_conversions, number_format],
                                ['평균 CPC', safe_division(total_cost, total_clicks), currency_format],
                                ['평균 CVR', safe_division(total_conversions, total_clicks), percent_format],
                                ['평균 CTR', safe_division(total_clicks, total_impressions), percent_format]
                            ]
                            for i, (label, value, fmt) in enumerate(metrics_data):
                                summary_sheet.write(i+10, 0, label, text_format)
                                summary_sheet.write(i+10, 1, value, fmt)
                            
                            # 열 너비 조정
                            summary_sheet.set_column('A:A', 20)
                            summary_sheet.set_column('B:B', 30)
                            
                            # 2. 기간별 성과 시트
                            period_sheet = workbook.add_worksheet('기간별_성과')
                            period_sheet.merge_range('A1:K1', f'{date_option} 마케팅 성과 상세', title_format)
                            
                            # 헤더 작성
                            headers = ['기간', '총비용', '전환매출액', 'ROAS', '노출수', '클릭수', 'CTR', '전환수', 'CVR', 'CPC', 'CPA']
                            for col, header in enumerate(headers):
                                period_sheet.write(2, col, header, header_format)
                            
                            # 데이터 작성
                            for row, data in enumerate(group_df_excel.values):
                                period_sheet.write(row+3, 0, data[0], text_format)  # 기간
                                period_sheet.write(row+3, 1, data[1], currency_format)  # 총비용
                                period_sheet.write(row+3, 2, data[4], currency_format)  # 전환매출액
                                period_sheet.write(row+3, 3, data[10]/100, percent_format)  # ROAS
                                period_sheet.write(row+3, 4, data[2], number_format)  # 노출수
                                period_sheet.write(row+3, 5, data[3], number_format)  # 클릭수
                                period_sheet.write(row+3, 6, data[7]/100, percent_format)  # CTR
                                period_sheet.write(row+3, 7, data[4], number_format)  # 전환수
                                period_sheet.write(row+3, 8, data[9]/100, percent_format)  # CVR
                                period_sheet.write(row+3, 9, data[6], currency_format)  # CPC
                                period_sheet.write(row+3, 10, data[8], currency_format)  # CPA
                            
                            # 3. 캠페인별 성과 시트
                            campaign_sheet = workbook.add_worksheet('캠페인별_성과')
                            campaign_sheet.merge_range('A1:K1', '캠페인별 성과 분석', title_format)
                            
                            # 캠페인 데이터 정렬
                            campaign_df_sorted = campaign_df_excel.sort_values('총비용(VAT포함,원)', ascending=False)
                            
                            # 헤더 작성
                            campaign_headers = ['캠페인명', '총비용', '전환매출액', 'ROAS', '운영비중', '노출수', '클릭수', 'CTR', '전환수', 'CVR', '비고']
                            for col, header in enumerate(campaign_headers):
                                campaign_sheet.write(2, col, header, header_format)
                            
                            # 데이터 작성
                            for row, (_, data) in enumerate(campaign_df_sorted.iterrows()):
                                campaign_sheet.write(row+3, 0, data['캠페인'], text_format)
                                campaign_sheet.write(row+3, 1, data['총비용(VAT포함,원)'], currency_format)
                                campaign_sheet.write(row+3, 2, data['전환매출액(원)'], currency_format)
                                campaign_sheet.write(row+3, 3, data['ROAS%']/100, percent_format)
                                campaign_sheet.write(row+3, 4, data['운영비중']/100, percent_format)
                                campaign_sheet.write(row+3, 5, data['노출수'], number_format)
                                campaign_sheet.write(row+3, 6, data['클릭수'], number_format)
                                campaign_sheet.write(row+3, 7, data['CTR%']/100, percent_format)
                                campaign_sheet.write(row+3, 8, data['전환수'], number_format)
                                campaign_sheet.write(row+3, 9, data['CVR%']/100, percent_format)
                            
                            # 4. 카테고리별 성과 시트
                            category_sheet = workbook.add_worksheet('카테고리별_성과')
                            category_sheet.merge_range('A1:J1', '카테고리별 성과 분석', title_format)
                            
                            # 카테고리 데이터 정렬
                            category_df_sorted = category_df_excel.sort_values('총비용(VAT포함,원)', ascending=False)
                            
                            # 헤더 작성
                            category_headers = ['카테고리', '총비용', '전환매출액', 'ROAS', '운영비중', '노출수', '클릭수', 'CTR', '전환수', 'CVR']
                            for col, header in enumerate(category_headers):
                                category_sheet.write(2, col, header, header_format)
                            
                            # 데이터 작성
                            for row, (_, data) in enumerate(category_df_sorted.iterrows()):
                                category_sheet.write(row+3, 0, data['캠페인카테고리'], text_format)
                                category_sheet.write(row+3, 1, data['총비용(VAT포함,원)'], currency_format)
                                category_sheet.write(row+3, 2, data['전환매출액(원)'], currency_format)
                                category_sheet.write(row+3, 3, data['ROAS%']/100, percent_format)
                                category_sheet.write(row+3, 4, data['운영비중']/100, percent_format)
                                category_sheet.write(row+3, 5, data['노출수'], number_format)
                                category_sheet.write(row+3, 6, data['클릭수'], number_format)
                                category_sheet.write(row+3, 7, data['CTR%']/100, percent_format)
                                category_sheet.write(row+3, 8, data['전환수'], number_format)
                                category_sheet.write(row+3, 9, data['CVR%']/100, percent_format)

                            if '캠페인유형' in filtered_df.columns:
                                # 5. 캠페인유형별 성과 시트
                                type_sheet = workbook.add_worksheet('캠페인유형별_성과')
                                type_sheet.merge_range('A1:J1', '캠페인유형별 성과 분석', title_format)
                                
                                # 캠페인유형 데이터 정렬
                                campaign_type_df_sorted = campaign_type_df.sort_values('총비용(VAT포함,원)', ascending=False)
                                
                                # 헤더 작성
                                type_headers = ['캠페인유형', '총비용', '전환매출액', 'ROAS', '운영비중', '노출수', '클릭수', 'CTR', '전환수', 'CVR']
                                for col, header in enumerate(type_headers):
                                    type_sheet.write(2, col, header, header_format)
                                
                                # 데이터 작성
                                for row, (_, data) in enumerate(campaign_type_df_sorted.iterrows()):
                                    type_sheet.write(row+3, 0, data['캠페인유형'], text_format)
                                    type_sheet.write(row+3, 1, data['총비용(VAT포함,원)'], currency_format)
                                    type_sheet.write(row+3, 2, data['전환매출액(원)'], currency_format)
                                    type_sheet.write(row+3, 3, data['ROAS%']/100, percent_format)
                                    type_sheet.write(row+3, 4, data['운영비중'], percent_format)
                                    type_sheet.write(row+3, 5, data['노출수'], number_format)
                                    type_sheet.write(row+3, 6, data['클릭수'], number_format)
                                    type_sheet.write(row+3, 7, data['CTR%']/100, percent_format)
                                    type_sheet.write(row+3, 8, data['전환수'], number_format)
                                    type_sheet.write(row+3, 9, data['CVR%']/100, percent_format)

                            # 6. 매체별 성과 시트
                            media_df = filtered_df.groupby('PC/모바일 매체').agg({
                                '총비용(VAT포함,원)': 'sum',
                                '전환매출액(원)': 'sum',
                                '노출수': 'sum',
                                '클릭수': 'sum',
                                '전환수': 'sum'
                            }).reset_index()
                            
                            media_df['ROAS%'] = safe_division(media_df['전환매출액(원)'], media_df['총비용(VAT포함,원)']) * 100
                            media_df['CTR%'] = safe_division(media_df['클릭수'], media_df['노출수']) * 100
                            media_df['CVR%'] = safe_division(media_df['전환수'], media_df['클릭수']) * 100
                            media_df['운영비중'] = safe_division(media_df['총비용(VAT포함,원)'], media_df['총비용(VAT포함,원)'].sum()) * 100
                            
                            media_sheet = workbook.add_worksheet('매체별_성과')
                            media_sheet.merge_range('A1:J1', '매체별 성과 분석', title_format)
                            
                            # 헤더 작성
                            media_headers = ['매체', '총비용', '전환매출액', 'ROAS', '운영비중', '노출수', '클릭수', 'CTR', '전환수', 'CVR']
                            for col, header in enumerate(media_headers):
                                media_sheet.write(2, col, header, header_format)
                            
                            # 데이터 작성
                            for row, (_, data) in enumerate(media_df.iterrows()):
                                media_sheet.write(row+3, 0, data['PC/모바일 매체'], text_format)
                                media_sheet.write(row+3, 1, data['총비용(VAT포함,원)'], currency_format)
                                media_sheet.write(row+3, 2, data['전환매출액(원)'], currency_format)
                                media_sheet.write(row+3, 3, data['ROAS%']/100, percent_format)
                                media_sheet.write(row+3, 4, data['운영비중']/100, percent_format)
                                media_sheet.write(row+3, 5, data['노출수'], number_format)
                                media_sheet.write(row+3, 6, data['클릭수'], number_format)
                                media_sheet.write(row+3, 7, data['CTR%']/100, percent_format)
                                media_sheet.write(row+3, 8, data['전환수'], number_format)
                                media_sheet.write(row+3, 9, data['CVR%']/100, percent_format)

                            # 7. 일자별 상세 데이터 시트
                            detail_sheet = workbook.add_worksheet('일자별_상세데이터')
                            detail_sheet.merge_range('A1:M1', '일자별 상세 데이터', title_format)
                            
                            # 날짜별로 정렬
                            filtered_df_sorted = filtered_df_excel.sort_values('일별')
                            
                            # 헤더 작성
                            detail_headers = [
                                '일자', '캠페인', '캠페인카테고리', 'PC/모바일 매체',
                                '총비용', '전환매출액', 'ROAS', '노출수', '클릭수', 'CTR',
                                '전환수', 'CVR', 'CPC'
                            ]
                            for col, header in enumerate(detail_headers):
                                detail_sheet.write(2, col, header, header_format)
                            
                            # 데이터 작성
                            for row, (_, data) in enumerate(filtered_df_sorted.iterrows()):
                                detail_sheet.write(row+3, 0, data['일별'], date_format)
                                detail_sheet.write(row+3, 1, data['캠페인'], text_format)
                                detail_sheet.write(row+3, 2, data['캠페인카테고리'], text_format)
                                detail_sheet.write(row+3, 3, data['PC/모바일 매체'], text_format)
                                detail_sheet.write(row+3, 4, data['총비용(VAT포함,원)'], currency_format)
                                detail_sheet.write(row+3, 5, data['전환매출액(원)'], currency_format)
                                detail_sheet.write(row+3, 6, data['ROAS%']/100, percent_format)
                                detail_sheet.write(row+3, 7, data['노출수'], number_format)
                                detail_sheet.write(row+3, 8, data['클릭수'], number_format)
                                detail_sheet.write(row+3, 9, data['CTR%']/100, percent_format)
                                detail_sheet.write(row+3, 10, data['전환수'], number_format)
                                detail_sheet.write(row+3, 11, data['CVR%']/100, percent_format)
                                detail_sheet.write(row+3, 12, data['CPC'], currency_format)

                            # 8. 주요 분석 인사이트 시트
                            insight_sheet = workbook.add_worksheet('주요_인사이트')
                            insight_sheet.merge_range('A1:D1', '주요 마케팅 성과 인사이트', title_format)
                            
                            # 성과 상위 캠페인
                            insight_sheet.merge_range('A3:D3', '성과 우수 캠페인 (ROAS 기준)', subtitle_format)
                            top_roas_campaigns = campaign_df.nlargest(5, 'ROAS%')
                            
                            insight_headers = ['캠페인명', 'ROAS', '총비용', '전환매출액']
                            for col, header in enumerate(insight_headers):
                                insight_sheet.write(4, col, header, header_format)
                            
                            for row, (_, data) in enumerate(top_roas_campaigns.iterrows()):
                                insight_sheet.write(row+5, 0, data['캠페인'], text_format)
                                insight_sheet.write(row+5, 1, data['ROAS%']/100, percent_format)
                                insight_sheet.write(row+5, 2, data['총비용(VAT포함,원)'], currency_format)
                                insight_sheet.write(row+5, 3, data['전환매출액(원)'], currency_format)
                            
                            # 성과 하위 캠페인
                            insight_sheet.merge_range('A11:D11', '성과 개선 필요 캠페인 (ROAS 기준)', subtitle_format)
                            bottom_roas_campaigns = campaign_df.nsmallest(5, 'ROAS%')
                            
                            for col, header in enumerate(insight_headers):
                                insight_sheet.write(12, col, header, header_format)
                            
                            for row, (_, data) in enumerate(bottom_roas_campaigns.iterrows()):
                                insight_sheet.write(row+13, 0, data['캠페인'], text_format)
                                insight_sheet.write(row+13, 1, data['ROAS%']/100, percent_format)
                                insight_sheet.write(row+13, 2, data['총비용(VAT포함,원)'], currency_format)
                                insight_sheet.write(row+13, 3, data['전환매출액(원)'], currency_format)
                            
                            # 카테고리별 성과 분석
                            insight_sheet.merge_range('A19:D19', '카테고리별 성과 분석', subtitle_format)
                            category_analysis = category_df.sort_values('ROAS%', ascending=False)
                            
                            category_headers = ['카테고리', 'ROAS', '운영비중', '전환매출액']
                            for col, header in enumerate(category_headers):
                                insight_sheet.write(20, col, header, header_format)
                            
                            for row, (_, data) in enumerate(category_analysis.iterrows()):
                                insight_sheet.write(row+21, 0, data['캠페인카테고리'], text_format)
                                insight_sheet.write(row+21, 1, data['ROAS%']/100, percent_format)
                                insight_sheet.write(row+21, 2, data['운영비중']/100, percent_format)
                                insight_sheet.write(row+21, 3, data['전환매출액(원)'], currency_format)

                            # 9. 기간별 트렌드 차트 (선택적)
                            if include_charts:
                                chart_sheet = workbook.add_worksheet('성과_차트')
                                chart_sheet.merge_range('A1:J1', '주요 지표 트렌드 차트', title_format)
                                
                                # 비용/매출 트렌드 차트
                                revenue_cost_chart = workbook.add_chart({'type': 'line'})
                                revenue_cost_chart.add_series({
                                    'name': '총비용',
                                    'categories': f'기간별_성과!A4:A{len(group_df_excel)+3}',
                                    'values': f'기간별_성과!B4:B{len(group_df_excel)+3}',
                                    'line': {'color': 'blue'}
                                })
                                revenue_cost_chart.add_series({
                                    'name': '전환매출액',
                                    'categories': f'기간별_성과!A4:A{len(group_df_excel)+3}',
                                    'values': f'기간별_성과!C4:C{len(group_df_excel)+3}',
                                    'line': {'color': 'green'}
                                })
                                revenue_cost_chart.set_title({'name': '비용/매출 트렌드'})
                                revenue_cost_chart.set_size({'width': 720, 'height': 400})
                                chart_sheet.insert_chart('A3', revenue_cost_chart)
                                
                                # ROAS 트렌드 차트
                                roas_chart = workbook.add_chart({'type': 'column'})
                                roas_chart.add_series({
                                    'name': 'ROAS',
                                    'categories': f'기간별_성과!A4:A{len(group_df_excel)+3}',
                                    'values': f'기간별_성과!D4:D{len(group_df_excel)+3}',
                                    'fill': {'color': 'orange'}
                                })
                                roas_chart.set_title({'name': 'ROAS 트렌드'})
                                roas_chart.set_size({'width': 720, 'height': 400})
                                chart_sheet.insert_chart('A25', roas_chart)

                            # 모든 시트의 열 너비 자동 조정
                            for worksheet in workbook.worksheets():
                                worksheet.set_column('A:A', 25)  # 첫 번째 열
                                worksheet.set_column('B:Z', 15)  # 나머지 열
                                worksheet.set_zoom(85)  # 기본 확대/축소 비율 설정
                            
                            # 필터 추가
                            for worksheet in workbook.worksheets():
                                if worksheet.name not in ['개요', '성과_차트']:
                                    worksheet.autofilter(2, 0, 2, worksheet.dim_colmax)

                        # 버퍼 위치 처음으로
                        buffer.seek(0)

                        # 다운로드 버튼
                        st.download_button(
                            label="📥 상세 보고서 다운로드",
                            data=buffer,
                            file_name=f"{report_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        st.success('보고서가 성공적으로 생성되었습니다! 다운로드 버튼을 클릭하여 저장하세요.')

                except Exception as e:
                    st.error(f"보고서 생성 중 오류가 발생했습니다: {str(e)}")
                    st.write("오류 상세 정보:", e)


    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {str(e)}")
        st.stop()

else:
    st.info("👈 좌측 사이드바에서 Excel 파일을 업로드하세요 ('raw 시트'와 'index 시트'가 포함된 파일)")

# 페이지 여백 추가
st.markdown("<br><br>", unsafe_allow_html=True)