import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
from datetime import datetime, timedelta
import locale
import warnings
import calendar
from xlsxwriter.utility import xl_col_to_name

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

def safe_division(x, y):
    """안전한 나눗셈 함수"""
    return np.where(y != 0, x / y, 0)

def format_date_axis(fig, date_option):
    """날짜 축 포맷 설정 함수"""
    if date_option == '일별':
        dtick = 'D1'
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
        nticks=20
    )
    return fig

def validate_data(df, required_columns):
    """데이터 유효성 검증 함수"""
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
    return True

def calculate_metrics(df):
    """모든 지표 계산 함수"""
    metrics = df.copy()
    
    # 기본 지표 계산
    metrics['CPC'] = safe_division(metrics['총비용(VAT포함,원)'], metrics['클릭수'])
    metrics['CTR'] = safe_division(metrics['클릭수'], metrics['노출수']) * 100
    metrics['CPA'] = safe_division(metrics['총비용(VAT포함,원)'], metrics['전환수'])
    metrics['CVR'] = safe_division(metrics['전환수'], metrics['클릭수']) * 100
    metrics['ROAS'] = safe_division(metrics['전환매출액(원)'], metrics['총비용(VAT포함,원)']) * 100
    metrics['ARPPU'] = safe_division(metrics['전환매출액(원)'], metrics['전환수'])
    
    if '운영비중' not in metrics.columns:
        total_cost = metrics['총비용(VAT포함,원)'].sum()
        metrics['운영비중'] = safe_division(metrics['총비용(VAT포함,원)'], total_cost) * 100
    
    # 소수점 자리수 조정
    metrics['평균노출순위'] = metrics['평균노출순위'].round(1)
    
    return metrics

# 피벗 테이블용 스타일링 함수
def style_pivot_table(df):
    """피벗 테이블 스타일링 함수"""
    formats = {}
    
    # 각 컬럼에 대해 포맷 지정
    for col in df.columns:
        if '총비용' in col or '전환매출액' in col or 'ARPPU' in col:
            formats[col] = '₩{:,.0f}'
        elif 'ROAS' in col:
            formats[col] = '{:,.2f}%'
        elif '평균노출순위' in col:
            formats[col] = '{:.1f}'
        elif '전환수' in col:
            formats[col] = '{:,.0f}'
    
    return df.style.format(formats).set_properties(**{
        'text-align': 'right',
        'font-size': '12px',
        'padding': '5px'
    })

def style_dataframe(df):
    """데이터프레임 스타일링 함수"""
    return df.style.format({
        '총비용(VAT포함,원)': '₩{:,.0f}',
        '노출수': '{:,.0f}',
        '클릭수': '{:,.0f}',
        '전환수': '{:,.0f}',
        '전환매출액(원)': '₩{:,.0f}',
        'CPC': '₩{:,.0f}',
        'CTR': '{:,.2f}%',
        'CPA': '₩{:,.0f}',
        'CVR': '{:,.2f}%',
        'ROAS': '{:,.2f}%',
        'ARPPU': '₩{:,.0f}',
        '평균노출순위': '{:.1f}',
        '운영비중': '{:.2f}%'
    }).set_properties(**{
        'text-align': 'right',
        'font-size': '12px',
        'padding': '5px'
    })

# '전체' 선택 처리 함수
def handle_select_all(option_list, selected_options):
    """'전체' 선택 처리 함수"""
    if '전체' in selected_options or len(selected_options) == 0:
        return option_list
    else:
        return selected_options

# 데이터 로딩 함수에 캐시 적용
@st.cache_data
def load_data(uploaded_file):
    """데이터 로딩 함수"""
    raw_df = pd.read_excel(uploaded_file, sheet_name='raw')
    index_df = pd.read_excel(uploaded_file, sheet_name='index')
    return raw_df, index_df

# 데이터 전처리 함수에 캐시 적용
@st.cache_data
def preprocess_data(raw_df):
    """데이터 전처리 함수"""
    # raw 시트 필수 컬럼
    required_columns_raw = [
        '일별', '캠페인유형', '캠페인', '광고그룹', '키워드', 'PC/모바일 매체',
        '노출수', '클릭수', '총비용(VAT포함,원)', '전환수', '평균노출순위', '전환매출액(원)',
        '캠페인 카테고리', '캠페인 국가'
    ]
    
    # 컬럼 존재 여부 확인
    validate_data(raw_df, required_columns_raw)

    # 컬럼명 통일
    raw_df.columns = raw_df.columns.str.strip()
    
    # 날짜 형식 변환
    raw_df['일별'] = pd.to_datetime(raw_df['일별'])
    
    # 주차 및 월 정보 추가
    raw_df['주차'] = raw_df['일별'].dt.isocalendar().week
    raw_df['월'] = raw_df['일별'].dt.month
    
    # 주차별 날짜 범위 추가
    week_dates = raw_df.groupby('주차')['일별'].agg(['min', 'max']).reset_index()
    week_dates['주차_기간'] = week_dates.apply(lambda row: f"{int(row['주차'])}주차 ({row['min'].strftime('%Y.%m.%d')}~{row['max'].strftime('%Y.%m.%d')})", axis=1)
    raw_df = pd.merge(raw_df, week_dates[['주차', '주차_기간']], on='주차', how='left')

    # 카테고리 및 지역 정보 매핑
    raw_df['캠페인카테고리'] = raw_df['캠페인 카테고리']
    raw_df['지역'] = raw_df['캠페인 국가']
    raw_df['상품유형'] = raw_df['캠페인 카테고리']

    # 숫자형 컬럼 변환
    numeric_columns = ['노출수', '클릭수', '총비용(VAT포함,원)', '전환수', '전환매출액(원)', '평균노출순위']
    for col in numeric_columns:
        raw_df[col] = pd.to_numeric(raw_df[col].astype(str).str.replace(',', '').replace('[-+]', ''), errors='coerce')

    # 프로모션 기간 설정 (예시: 10월 17일 이후)
    promo_start_date = pd.Timestamp('2023-10-17')
    raw_df['프로모션여부'] = raw_df['일별'].ge(promo_start_date).map({True: 'Y', False: 'N'})

    return raw_df

# 파일 업로드
st.sidebar.header("📂 파일 업로드")
uploaded_file = st.sidebar.file_uploader("Excel 파일을 업로드하세요 ('raw 시트'와 'index 시트' 포함)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        with st.spinner('데이터를 불러오는 중...'):
            # 데이터 로딩 (캐시 적용)
            raw_df, index_df = load_data(uploaded_file)
            # 데이터 전처리 (캐시 적용)
            raw_df = preprocess_data(raw_df)

        # 데이터 미리보기
        with st.expander("데이터 미리보기"):
            st.subheader("raw 시트 데이터")
            st.dataframe(raw_df.head())
            st.subheader("index 시트 데이터")
            st.dataframe(index_df.head())

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

        # 필터링을 위한 유니크 값들 미리 계산
        unique_categories = sorted(raw_df['캠페인카테고리'].dropna().astype(str).unique())
        unique_regions = sorted(raw_df['지역'].dropna().astype(str).unique())
        unique_campaign_types = sorted(raw_df['캠페인유형'].dropna().astype(str).unique())
        unique_product_types = sorted(raw_df['상품유형'].dropna().astype(str).unique())
        unique_media = sorted(raw_df['PC/모바일 매체'].dropna().astype(str).unique())

        # 캠페인 카테고리 필터
        category_options = st.sidebar.multiselect(
            '캠페인 카테고리 선택',
            options=['전체'] + unique_categories,
            default=['전체']
        )

        # 지역 필터
        region_options = st.sidebar.multiselect(
            '지역 선택',
            options=['전체'] + unique_regions,
            default=['전체']
        )

        # 캠페인유형 필터
        campaign_type_options = st.sidebar.multiselect(
            '캠페인유형 선택',
            options=['전체'] + unique_campaign_types,
            default=['전체']
        )

        # 상품유형 필터 추가
        product_type_options = st.sidebar.multiselect(
            '상품유형 선택',
            options=['전체'] + unique_product_types,
            default=['전체']
        )

        # **전체 선택 처리**
        category_options = handle_select_all(unique_categories, category_options)
        region_options = handle_select_all(unique_regions, region_options)
        campaign_type_options = handle_select_all(unique_campaign_types, campaign_type_options)
        product_type_options = handle_select_all(unique_product_types, product_type_options)

        # 캠페인 필터
        filtered_for_campaigns = raw_df[
            (raw_df['캠페인카테고리'].isin(category_options)) &
            (raw_df['지역'].isin(region_options)) &
            (raw_df['캠페인유형'].isin(campaign_type_options)) &
            (raw_df['상품유형'].isin(product_type_options))
        ]
        unique_campaigns = sorted(filtered_for_campaigns['캠페인'].unique())
        
        campaign_options = st.sidebar.multiselect(
            '캠페인 선택',
            options=['전체'] + unique_campaigns,
            default=['전체']
        )

        # **캠페인 전체 선택 처리**
        campaign_options = handle_select_all(unique_campaigns, campaign_options)

        # 광고그룹 필터 추가
        filtered_for_adgroups = raw_df[
            (raw_df['캠페인카테고리'].isin(category_options)) &
            (raw_df['지역'].isin(region_options)) &
            (raw_df['캠페인유형'].isin(campaign_type_options)) &
            (raw_df['상품유형'].isin(product_type_options)) &
            (raw_df['캠페인'].isin(campaign_options))
        ]
        unique_adgroups = sorted(filtered_for_adgroups['광고그룹'].unique())

        adgroup_options = st.sidebar.multiselect(
            '광고그룹 선택',
            options=['전체'] + list(unique_adgroups),
            default=['전체']
        )

        # **광고그룹 전체 선택 처리**
        adgroup_options = handle_select_all(unique_adgroups, adgroup_options)

        # 매체 필터
        media_options = st.sidebar.multiselect(
            '매체 선택 (PC/모바일)',
            options=['전체'] + unique_media,
            default=['전체']
        )

        # **매체 전체 선택 처리**
        media_options = handle_select_all(unique_media, media_options)

        # 데이터 필터링
        @st.cache_data
        def filter_data(raw_df, start_date, end_date, category_options, campaign_options,
                        adgroup_options, media_options, region_options, campaign_type_options, product_type_options):
            mask = (raw_df['일별'] >= pd.to_datetime(start_date)) & \
                   (raw_df['일별'] <= pd.to_datetime(end_date)) & \
                   (raw_df['캠페인카테고리'].isin(category_options)) & \
                   (raw_df['캠페인'].isin(campaign_options)) & \
                   (raw_df['광고그룹'].isin(adgroup_options)) & \
                   (raw_df['PC/모바일 매체'].isin(media_options)) & \
                   (raw_df['지역'].isin(region_options)) & \
                   (raw_df['캠페인유형'].isin(campaign_type_options)) & \
                   (raw_df['상품유형'].isin(product_type_options))

            filtered_df = raw_df.loc[mask].copy()
            return filtered_df

        filtered_df = filter_data(raw_df, start_date, end_date, category_options, campaign_options,
                                  adgroup_options, media_options, region_options, campaign_type_options, product_type_options)

        if filtered_df.empty:
            st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다. 필터 조건을 조정해주세요.")
            st.stop()

        # 지표 계산
        filtered_df = calculate_metrics(filtered_df)

        # '전체' 상품유형 데이터 추가
        total_df = filtered_df.copy()
        total_df['상품유형'] = '전체'

        # '전체' 데이터와 원본 데이터 결합
        filtered_df = pd.concat([filtered_df, total_df], ignore_index=True)

        # 데이터 집계
        agg_dict = {
            '총비용(VAT포함,원)': 'sum',
            '노출수': 'sum',
            '클릭수': 'sum',
            '전환수': 'sum',
            '전환매출액(원)': 'sum',
            '평균노출순위': 'mean',
            '키워드': 'nunique'
        }

        # 기간별 집계 (상품유형 포함)
        if date_option == '일별':
            group_df = filtered_df.groupby(['일별', '상품유형']).agg(agg_dict).reset_index()
            group_df.rename(columns={'일별': '기간'}, inplace=True)
        elif date_option == '주간':
            group_df = filtered_df.groupby(['주차_기간', '상품유형']).agg(agg_dict).reset_index()
            group_df.rename(columns={'주차_기간': '기간'}, inplace=True)
        else:  # 월별
            group_df = filtered_df.groupby(['월', '상품유형']).agg(agg_dict).reset_index()
            group_df.rename(columns={'월': '기간'}, inplace=True)

        group_df = calculate_metrics(group_df)

        # 중복 데이터 제거
        group_df = group_df.drop_duplicates()

        # 4. 시각화
        st.header("📈 기간별 성과 지표")

        # 주요 지표 카드
        total_cost = filtered_df[filtered_df['상품유형'] == '전체']['총비용(VAT포함,원)'].sum()
        total_revenue = filtered_df[filtered_df['상품유형'] == '전체']['전환매출액(원)'].sum()
        total_roas = safe_division(total_revenue, total_cost) * 100
        total_clicks = filtered_df[filtered_df['상품유형'] == '전체']['클릭수'].sum()
        total_impressions = filtered_df[filtered_df['상품유형'] == '전체']['노출수'].sum()
        total_conversions = filtered_df[filtered_df['상품유형'] == '전체']['전환수'].sum()
        avg_arppu = safe_division(total_revenue, total_conversions)
        avg_rank = filtered_df[filtered_df['상품유형'] == '전체']['평균노출순위'].mean()
        # 지표 카드 표시
        col1, col2, col3, col4, col5 = st.columns(5)

        col1.metric("총비용", f"₩{int(total_cost):,}")
        col2.metric("ROAS", f"{total_roas:.2f}%")
        col3.metric("ARPPU", f"₩{int(avg_arppu):,}")
        col4.metric("전환수", f"{int(total_conversions):,}")
        col5.metric("평균노출순위", f"{avg_rank:.1f}")

        # 차트 생성
        st.subheader(f"📊 {date_option} 차트")

        # 탭 생성
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["비용/매출", "ROAS/ARPPU", "노출순위", "프로모션 분석", "국가별 매출 트렌드"])

        # 기간별 전체 데이터 필터링
        group_total_df = group_df[group_df['상품유형'] == '전체']

        # 비용/매출 차트
        with tab1:
            fig1 = go.Figure()

            df_to_plot = group_total_df
            x_col = '기간'

            fig1.add_trace(go.Scatter(
                x=df_to_plot[x_col],
                y=df_to_plot['총비용(VAT포함,원)'],
                name='총비용',
                line=dict(color='#1f77b4'),
                hovertemplate='%{x}<br>총비용: ₩%{y:,.0f}<extra></extra>'
            ))

            fig1.add_trace(go.Scatter(
                x=df_to_plot[x_col],
                y=df_to_plot['전환매출액(원)'],
                name='전환매출액',
                line=dict(color='#2ca02c'),
                hovertemplate='%{x}<br>전환매출액: ₩%{y:,.0f}<extra></extra>'
            ))

            fig1.update_layout(
                title=f'{date_option}별 비용/매출 추이',
                xaxis_title=date_option,
                yaxis_title='금액 (원)',
                hovermode='x unified',
                showlegend=True,
                height=600
            )

            if date_option == '일별':
                fig1 = format_date_axis(fig1, date_option)
            else:
                fig1.update_xaxes(tickangle=45)

            st.plotly_chart(fig1, use_container_width=True)

        # ROAS/ARPPU 차트
        with tab2:
            fig2 = go.Figure()

            df_to_plot = group_total_df

            fig2.add_trace(go.Bar(
                x=df_to_plot[x_col],
                y=df_to_plot['ROAS'],
                name='ROAS',
                marker_color='#1f77b4',
                yaxis='y',
                text=df_to_plot['ROAS'].round(2).astype(str) + '%',
                textposition='outside',
                hovertemplate='%{x}<br>ROAS: %{y:.2f}%<extra></extra>'
            ))

            fig2.add_trace(go.Scatter(
                x=df_to_plot[x_col],
                y=df_to_plot['ARPPU'],
                name='ARPPU',
                line=dict(color='#2ca02c'),
                yaxis='y2',
                hovertemplate='%{x}<br>ARPPU: ₩%{y:,.0f}<extra></extra>'
            ))

            fig2.update_layout(
                title=f'{date_option}별 ROAS/ARPPU 추이',
                xaxis_title=date_option,
                yaxis=dict(title='ROAS (%)', titlefont=dict(color='#1f77b4')),
                yaxis2=dict(
                    title='ARPPU (원)',
                    titlefont=dict(color='#2ca02c'),
                    overlaying='y',
                    side='right'
                ),
                hovermode='x unified',
                showlegend=True,
                height=600
            )

            if date_option == '일별':
                fig2 = format_date_axis(fig2, date_option)
            else:
                fig2.update_xaxes(tickangle=45)

            st.plotly_chart(fig2, use_container_width=True)

        # 노출순위 차트
        with tab3:
            fig3 = go.Figure()

            fig3.add_trace(go.Scatter(
                x=df_to_plot[x_col],
                y=df_to_plot['평균노출순위'],
                name='평균노출순위',
                line=dict(color='#ff7f0e'),
                hovertemplate='%{x}<br>평균노출순위: %{y:.1f}<extra></extra>'
            ))

            fig3.update_layout(
                title=f'{date_option}별 평균 노출순위 추이',
                xaxis_title=date_option,
                yaxis_title='평균 노출순위',
                hovermode='x unified',
                showlegend=True,
                yaxis=dict(autorange="reversed"),  # 노출순위는 낮을수록 좋으므로 역순으로 표시
                height=600
            )

            if date_option == '일별':
                fig3 = format_date_axis(fig3, date_option)
            else:
                fig3.update_xaxes(tickangle=45)

            st.plotly_chart(fig3, use_container_width=True)

        # 프로모션 분석
        with tab4:
            promo_metrics = filtered_df.groupby(['프로모션여부', '상품유형']).agg(agg_dict).reset_index()
            promo_metrics = calculate_metrics(promo_metrics)

            # '전체' 상품유형 데이터만 사용
            promo_metrics_total = promo_metrics[promo_metrics['상품유형'] == '전체']

            if len(promo_metrics_total['프로모션여부'].unique()) >= 2:  # 프로모션 전/후 데이터가 모두 있는 경우
                # 프로모션 성과 비교 차트
                fig4 = go.Figure()

                promo_before = promo_metrics_total[promo_metrics_total['프로모션여부'] == 'N']
                promo_after = promo_metrics_total[promo_metrics_total['프로모션여부'] == 'Y']

                fig4.add_trace(go.Bar(
                    x=['프로모션 전', '프로모션 후'],
                    y=[promo_before['ROAS'].iloc[0], promo_after['ROAS'].iloc[0]],
                    name='ROAS',
                    marker_color='#1f77b4',
                    text=[f"{x:.2f}%" for x in [promo_before['ROAS'].iloc[0], promo_after['ROAS'].iloc[0]]],
                    textposition='outside'
                ))

                fig4.update_layout(
                    title='프로모션 전후 ROAS 비교',
                    yaxis_title='ROAS (%)',
                    showlegend=False,
                    height=600
                )

                st.plotly_chart(fig4, use_container_width=True)

                # 프로모션 성과 상세 비교표
                st.subheader("📊 프로모션 성과 상세 비교")
                promo_comparison = pd.DataFrame({
                    '구분': ['프로모션 전', '프로모션 후'],
                    '총비용(VAT포함,원)': [promo_before['총비용(VAT포함,원)'].iloc[0], promo_after['총비용(VAT포함,원)'].iloc[0]],
                    'ROAS': [promo_before['ROAS'].iloc[0], promo_after['ROAS'].iloc[0]],
                    'ARPPU': [promo_before['ARPPU'].iloc[0], promo_after['ARPPU'].iloc[0]],
                    '전환수': [promo_before['전환수'].iloc[0], promo_after['전환수'].iloc[0]],
                    'CVR': [promo_before['CVR'].iloc[0], promo_after['CVR'].iloc[0]],
                    '평균노출순위': [promo_before['평균노출순위'].iloc[0], promo_after['평균노출순위'].iloc[0]]
                })
                
                st.dataframe(style_dataframe(promo_comparison))
            else:
                st.warning("프로모션 전/후 비교를 위한 충분한 데이터가 없습니다.")

        # 국가별 일별 Revenue 트렌드
        with tab5:
            st.subheader("🌍 국가별 일별 매출 트렌드")
            if '지역' in filtered_df.columns:
                country_daily_revenue = filtered_df[filtered_df['상품유형'] == '전체'].groupby(['일별', '지역']).agg({
                    '전환매출액(원)': 'sum'
                }).reset_index()

                # 국가 선택 옵션 추가
                unique_countries = country_daily_revenue['지역'].unique()
                total_revenue_by_country = country_daily_revenue.groupby('지역')['전환매출액(원)'].sum().reset_index()
                total_revenue_by_country = total_revenue_by_country.sort_values('전환매출액(원)', ascending=False)
                top_10_countries = total_revenue_by_country['지역'].head(10).tolist()

                country_selection = st.multiselect(
                    '국가 선택',
                    options=['상위 10개 보기', '전체 보기'] + list(unique_countries),
                    default=['상위 10개 보기']
                )

                if '상위 10개 보기' in country_selection:
                    selected_countries = top_10_countries
                elif '전체 보기' in country_selection or len(country_selection) == 0:
                    selected_countries = unique_countries
                else:
                    selected_countries = country_selection

                filtered_country_data = country_daily_revenue[country_daily_revenue['지역'].isin(selected_countries)]

                # 금액을 원화로 표시하고 소수점 제거
                filtered_country_data['전환매출액(원)'] = filtered_country_data['전환매출액(원)'].round(0)

                fig5 = px.line(
                    filtered_country_data,
                    x='일별',
                    y='전환매출액(원)',
                    color='지역',
                    title='국가별 일별 매출 트렌드',
                    labels={'일별': '일자', '전환매출액(원)': '매출액 (원)', '지역': '국가'}
                )

                fig5.update_layout(
                    xaxis_title='일자',
                    yaxis_title='매출액 (원)',
                    hovermode='x unified',
                    height=600
                )

                fig5.update_xaxes(tickformat='%Y-%m-%d', tickangle=45)
                fig5.update_yaxes(tickformat=',')  # 천단위 콤마 표시

                st.plotly_chart(fig5, use_container_width=True)
            else:
                st.warning("데이터에 '지역' 컬럼이 없습니다.")

        # 기간별 데이터 표 생성
        st.subheader(f"📅 {date_option} 데이터 표")
        styled_group_df = style_dataframe(group_total_df)
        st.dataframe(styled_group_df, height=400)

        # 5. 세부 분석
        st.header("🔍 세부 분석")

        # **일별 분석 추가**
        st.subheader("📅 일별 분석")

        # 일별 데이터 준비 및 지표 계산
        daily_by_product = filtered_df.groupby(['일별', '상품유형']).agg(agg_dict).reset_index()

        daily_by_product = calculate_metrics(daily_by_product)

        # 컬럼 순서 재정렬
        daily_by_product = daily_by_product[['일별', '상품유형', '총비용(VAT포함,원)', '전환매출액(원)', 'ROAS', 'ARPPU', '전환수', '평균노출순위']]

        # 날짜와 상품유형별로 정렬
        daily_by_product.sort_values(['일별', '상품유형'], inplace=True)

        st.dataframe(style_dataframe(daily_by_product))

        # 상품유형별 일별 트렌드 차트 추가
        st.subheader("📈 상품유형별 일별 ROAS 트렌드")
        fig_daily = go.Figure()

        for product_type in daily_by_product['상품유형'].unique():
            product_data = daily_by_product[daily_by_product['상품유형'] == product_type]
            fig_daily.add_trace(go.Scatter(
                x=product_data['일별'],
                y=product_data['ROAS'],
                name=product_type,
                mode='lines+markers',
                hovertemplate='%{x}<br>ROAS: %{y:.2f}%<extra></extra>'
            ))

        fig_daily.update_layout(
            title='상품유형별 일별 ROAS 트렌드',
            xaxis_title='일별',
            yaxis_title='ROAS (%)',
            hovermode='x unified',
            height=600
        )
        fig_daily.update_xaxes(tickformat='%Y-%m-%d', tickangle=45)

        st.plotly_chart(fig_daily, use_container_width=True)

        # **주간 분석**
        st.subheader("📅 상품유형별 주간 분석")
        
        # 주간 데이터 준비 및 지표 계산
        weekly_by_product = filtered_df.groupby(['주차', '주차_기간', '상품유형']).agg(agg_dict).reset_index()

        weekly_by_product = calculate_metrics(weekly_by_product)

        # 컬럼 순서 재정렬
        weekly_by_product = weekly_by_product[['주차', '주차_기간', '상품유형', '총비용(VAT포함,원)', '전환매출액(원)', 'ROAS', 'ARPPU', '전환수', '평균노출순위']]

        # 주차와 상품유형별로 정렬
        weekly_by_product.sort_values(['주차', '상품유형'], inplace=True)

        st.dataframe(style_dataframe(weekly_by_product))

        # 주간 트렌드 차트 추가
        st.subheader("📈 상품유형별 주간 ROAS 트렌드")
        fig_weekly = go.Figure()

        for product_type in weekly_by_product['상품유형'].unique():
            product_data = weekly_by_product[weekly_by_product['상품유형'] == product_type]
            fig_weekly.add_trace(go.Scatter(
                x=product_data['주차_기간'],
                y=product_data['ROAS'],
                name=product_type,
                mode='lines+markers',
                hovertemplate='%{x}<br>ROAS: %{y:.2f}%<extra></extra>'
            ))

        fig_weekly.update_layout(
            title='상품유형별 주간 ROAS 트렌드',
            xaxis_title='주차',
            yaxis_title='ROAS (%)',
            hovermode='x unified',
            height=600
        )
        fig_weekly.update_xaxes(tickangle=45)

        st.plotly_chart(fig_weekly, use_container_width=True)
        
        # 캠페인유형별 분석
        st.subheader("🎯 캠페인유형별 분석")
        
        # 매체별 및 상품유형 추가
        campaign_type_metrics = filtered_df.groupby(['캠페인유형', 'PC/모바일 매체', '상품유형']).agg(agg_dict).reset_index()
        campaign_type_metrics = calculate_metrics(campaign_type_metrics)
        campaign_type_metrics = campaign_type_metrics.drop_duplicates()
        
        # 캠페인유형별 성과표
        st.subheader("📊 캠페인유형별 성과")
        selected_campaign_type = st.selectbox('캠페인유형 선택', ['전체'] + list(campaign_type_metrics['캠페인유형'].unique()))
        if selected_campaign_type == '전체':
            campaign_type_filtered = campaign_type_metrics.copy()
        else:
            campaign_type_filtered = campaign_type_metrics[campaign_type_metrics['캠페인유형'] == selected_campaign_type]
        st.dataframe(style_dataframe(campaign_type_filtered))

        # 지역별 분석
        st.subheader("🌏 지역별 분석")
        
        region_metrics = filtered_df.groupby('지역').agg(agg_dict).reset_index()
        region_metrics = calculate_metrics(region_metrics)
        region_metrics = region_metrics.sort_values('총비용(VAT포함,원)', ascending=False)
        region_metrics = region_metrics.drop_duplicates()
        
        # 상위 10개 지역만 표시 옵션 추가
        region_display_option = st.selectbox('지역 표시 옵션', ['상위 10개 보기', '전체 보기'])
        if region_display_option == '상위 10개 보기':
            top_regions = region_metrics['지역'].head(10).tolist()
            region_metrics_top = region_metrics[region_metrics['지역'].isin(top_regions)]
        else:
            region_metrics_top = region_metrics.copy()
        
        # 지역별 ROAS 차트
        fig_region = px.bar(
            region_metrics_top,
            x='지역',
            y='ROAS',
            title='지역별 ROAS',
            text=region_metrics_top['ROAS'].round(2).astype(str) + '%',
            labels={'지역': '지역', 'ROAS': 'ROAS (%)'}
        )
        fig_region.update_traces(textposition='outside')
        fig_region.update_layout(
            height=600,
            xaxis_title='지역',
            yaxis_title='ROAS (%)',
            xaxis_tickangle=45
        )
        
        st.plotly_chart(fig_region, use_container_width=True)
        
        # 지역별 성과표
        st.dataframe(style_dataframe(region_metrics_top))

        # 매체별 분석
        st.subheader("📱 매체별 분석")
        
        media_metrics = filtered_df.groupby('PC/모바일 매체').agg(agg_dict).reset_index()
        media_metrics = calculate_metrics(media_metrics)
        media_metrics = media_metrics.drop_duplicates()
        
        # 매체별 성과 차트
        fig_media = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig_media.add_trace(go.Bar(
            x=media_metrics['PC/모바일 매체'],
            y=media_metrics['총비용(VAT포함,원)'],
            name='총비용',
            marker_color='#1f77b4',
            text=media_metrics['총비용(VAT포함,원)'].apply(lambda x: f'₩{x:,.0f}'),
            textposition='outside'
        ), secondary_y=False)

        fig_media.add_trace(go.Scatter(
            x=media_metrics['PC/모바일 매체'],
            y=media_metrics['ROAS'],
            name='ROAS',
            line=dict(color='#2ca02c'),
            yaxis='y2',
            text=media_metrics['ROAS'].round(2).astype(str) + '%',
            mode='lines+markers+text',
            textposition='bottom center'
        ), secondary_y=True)

        fig_media.update_layout(
            title='매체별 비용 및 ROAS',
            yaxis_title='총비용 (원)',
            yaxis2=dict(
                title='ROAS (%)',
                overlaying='y',
                side='right'
            ),
            showlegend=True,
            height=600
        )

        st.plotly_chart(fig_media, use_container_width=True)
        
        # 매체별 성과표
        st.dataframe(style_dataframe(media_metrics))

        # 캠페인별 분석
        st.subheader("🎯 캠페인별 분석")
        
        campaign_metrics = filtered_df.groupby(['캠페인', '캠페인카테고리', '상품유형']).agg(agg_dict).reset_index()
        campaign_metrics = calculate_metrics(campaign_metrics)
        campaign_metrics = campaign_metrics.sort_values('총비용(VAT포함,원)', ascending=False)
        campaign_metrics = campaign_metrics.drop_duplicates()
        
        # 정렬 옵션
        metric_options = ['총비용(VAT포함,원)', 'ROAS', '전환매출액(원)', 'ARPPU', '평균노출순위']
        selected_metric = st.selectbox('정렬 기준:', metric_options)
        
        # 캠페인 성과표
        campaign_metrics_sorted = campaign_metrics.sort_values(selected_metric, ascending=False)
        st.dataframe(style_dataframe(campaign_metrics_sorted))

        # 보고서 다운로드 부분
        st.header("📥 마케팅 성과 종합 보고서")

        report_cols = st.columns([2, 1])
        with report_cols[0]:
            report_name = st.text_input("보고서 파일명", value="마케팅_성과_보고서")
        with report_cols[1]:
            include_charts = st.checkbox("차트 포함", value=True)

        buffer = BytesIO()
        if st.button("📥 보고서 다운로드"):
            try:
                with st.spinner('상세 보고서를 생성하는 중...'):
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
                        
                        base_format = {
                            'border': 1,
                            'font_name': '맑은 고딕',
                            'font_size': 10
                        }
                        
                        currency_format = workbook.add_format({
                            **base_format,
                            'num_format': '₩#,##0',
                            'align': 'right'
                        })

                        number_format = workbook.add_format({
                            **base_format,
                            'num_format': '#,##0',
                            'align': 'right'
                        })

                        percent_format = workbook.add_format({
                            **base_format,
                            'num_format': '0.00%',
                            'align': 'right'
                        })

                        decimal_format = workbook.add_format({
                            **base_format,
                            'num_format': '0.0',
                            'align': 'right'
                        })

                        date_format = workbook.add_format({
                            **base_format,
                            'align': 'center',
                            'num_format': 'yyyy-mm-dd'
                        })
                        
                        text_format = workbook.add_format({
                            **base_format,
                            'align': 'left'
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
                            ['캠페인 수', len(campaign_metrics)],
                            ['선택된 카테고리', ', '.join(category_options)],
                            ['선택된 매체', ', '.join(media_options)],
                            ['선택된 상품유형', ', '.join(product_type_options)]
                        ]
                        for i, (label, value) in enumerate(info_data):
                            summary_sheet.write(i+3, 0, label, text_format)
                            summary_sheet.write(i+3, 1, value, text_format)
                        
                        # 주요 지표 요약
                        summary_sheet.merge_range('A10:B10', '주요 성과 지표', subtitle_format)
                        metrics_data = [
                            ['총 비용', total_cost, currency_format],
                            ['총 매출', total_revenue, currency_format],
                            ['ROAS', total_roas/100, percent_format],
                            ['총 노출수', total_impressions, number_format],
                            ['총 클릭수', total_clicks, number_format],
                            ['총 전환수', total_conversions, number_format],
                            ['평균 ARPPU', avg_arppu, currency_format],
                            ['평균 노출순위', avg_rank, decimal_format]
                        ]
                        
                        for i, (label, value, fmt) in enumerate(metrics_data):
                            summary_sheet.write(i+11, 0, label, text_format)
                            summary_sheet.write(i+11, 1, value, fmt)
                        
                        # 열 너비 조정
                        summary_sheet.set_column('A:A', 20)
                        summary_sheet.set_column('B:B', 30)
                        
                        # 필터 정보를 별도의 시트에 저장
                        filter_sheet = workbook.add_worksheet('필터 정보')
                        filter_sheet.write('A1', '필터 옵션', title_format)
                        filter_sheet.set_row(0, 30)
                        filter_sheet.write('A2', '캠페인 카테고리', header_format)
                        filter_sheet.write('A3', ', '.join(category_options), text_format)
                        filter_sheet.write('B2', '지역', header_format)
                        filter_sheet.write('B3', ', '.join(region_options), text_format)
                        filter_sheet.write('C2', '캠페인유형', header_format)
                        filter_sheet.write('C3', ', '.join(campaign_type_options), text_format)
                        filter_sheet.write('D2', '상품유형', header_format)
                        filter_sheet.write('D3', ', '.join(product_type_options), text_format)
                        filter_sheet.write('E2', '캠페인', header_format)
                        filter_sheet.write('E3', ', '.join(campaign_options), text_format)
                        filter_sheet.write('F2', '광고그룹', header_format)
                        filter_sheet.write('F3', ', '.join(adgroup_options), text_format)
                        filter_sheet.write('G2', '매체', header_format)
                        filter_sheet.write('G3', ', '.join(media_options), text_format)
                        
                        # 나머지 시트들 생성
                        # 각 데이터프레임을 저장하기 위해 딕셔너리에 저장
                        dfs_to_save = {}

                        # Excel 보고서에서는 '상품유형' 필터에 모든 상품유형이 나타나도록 하기 위해, '전체'를 포함한 데이터를 사용합니다.

                        # 기간별 성과 데이터 생성
                        raw_group_df = filtered_df.copy()

                        # '기간' 컬럼 생성
                        if date_option == '일별':
                            raw_group_df['기간'] = raw_group_df['일별']
                        elif date_option == '주간':
                            raw_group_df['기간'] = raw_group_df['주차_기간']
                        else:  # 월별
                            raw_group_df['기간'] = raw_group_df['월']

                        # 그룹화 및 지표 계산
                        raw_group_df = raw_group_df.groupby(['기간', '상품유형']).agg(agg_dict).reset_index()
                        raw_group_df = calculate_metrics(raw_group_df)
                        raw_group_df = raw_group_df.drop_duplicates()

                        # 캠페인유형별 성과 데이터 생성
                        raw_campaign_type_metrics = filtered_df.groupby(['캠페인유형', 'PC/모바일 매체', '상품유형']).agg(agg_dict).reset_index()
                        raw_campaign_type_metrics = calculate_metrics(raw_campaign_type_metrics)
                        raw_campaign_type_metrics = raw_campaign_type_metrics.drop_duplicates()

                        # 나머지 데이터는 필터링된 데이터를 사용합니다.

                        # 데이터 저장 함수
                        def save_df_to_excel(df, sheet_name, title, add_filter=False):
                            # 데이터 전처리
                            processed_df = df.copy()
                            
                            # ROAS나 ARPPU가 포함된 모든 컬럼 찾기
                            roas_cols = [col for col in processed_df.columns if 'ROAS' in col]
                            arppu_cols = [col for col in processed_df.columns if 'ARPPU' in col]
                            money_cols = [col for col in processed_df.columns if any(keyword in col for keyword in ['총비용', '전환매출액', 'CPC', 'CPA'])]
                            percent_cols = [col for col in processed_df.columns if any(keyword in col for keyword in ['CTR', 'CVR', '운영비중'])]
                            
                            # 데이터 형식 변환
                            for col in processed_df.columns:
                                if col in money_cols + arppu_cols:
                                    processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).round(0)
                                elif col in ['노출수', '클릭수', '전환수', '키워드']:
                                    processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).round(0)
                                elif col in roas_cols + percent_cols:
                                    processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).round(2) / 100
                                elif col == '평균노출순위' or '순위' in col:
                                    processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).round(1)
                                elif '일별' in col.lower() or col == '기간' or 'date' in col.lower():
                                    processed_df[col] = pd.to_datetime(processed_df[col], errors='coerce')

                            # 시트 생성
                            processed_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                            worksheet = writer.sheets[sheet_name]
                            
                            # 제목 추가
                            worksheet.merge_range(0, 0, 0, len(processed_df.columns)-1, title, title_format)
                            
                            # 컬럼별 포맷 적용
                            for col_num, col_name in enumerate(processed_df.columns):
                                # 헤더 포맷
                                worksheet.write(1, col_num, col_name, header_format)
                                
                                # 데이터 포맷
                                if col_name in money_cols + arppu_cols:
                                    worksheet.set_column(col_num, col_num, 15, currency_format)
                                elif col_name in ['노출수', '클릭수', '전환수', '키워드']:
                                    worksheet.set_column(col_num, col_num, 12, number_format)
                                elif col_name in roas_cols + percent_cols:
                                    worksheet.set_column(col_num, col_num, 12, percent_format)
                                elif col_name == '평균노출순위' or '순위' in col_name:
                                    worksheet.set_column(col_num, col_num, 12, decimal_format)
                                elif '일별' in col_name.lower() or col_name == '기간' or 'date' in col_name.lower():
                                    worksheet.set_column(col_num, col_num, 12, date_format)
                                else:
                                    worksheet.set_column(col_num, col_num, 15, text_format)
                            
                            # 필터 추가 여부
                            if add_filter:
                                worksheet.autofilter(1, 0, len(processed_df)+1, len(processed_df.columns)-1)
                            
                            # 창 틀 고정
                            worksheet.freeze_panes(2, 0)

                            # 데이터프레임 저장
                            dfs_to_save[sheet_name] = processed_df

                        # 각 시트 생성
                        save_df_to_excel(raw_group_df, '기간별_성과', f'{date_option} 마케팅 성과', add_filter=True)
                        save_df_to_excel(raw_campaign_type_metrics, '캠페인유형별_성과', '캠페인유형별 성과', add_filter=True)
                        save_df_to_excel(region_metrics, '지역별_성과', '지역별 성과', add_filter=True)
                        save_df_to_excel(media_metrics, '매체별_성과', '매체별 성과', add_filter=True)
                        save_df_to_excel(campaign_metrics_sorted, '캠페인별_성과', '캠페인별 성과', add_filter=True)
                        save_df_to_excel(weekly_by_product, '상품유형별_주간성과', '상품유형별 주간 성과', add_filter=True)
                        save_df_to_excel(daily_by_product, '상품유형별_일별성과', '상품유형별 일별 성과', add_filter=True)
                        
                        if len(promo_metrics['프로모션여부'].unique()) >= 2:
                            save_df_to_excel(promo_metrics, '프로모션_성과비교', '프로모션 성과 비교', add_filter=True)
                        
                        save_df_to_excel(filtered_df, '일자별_상세데이터', '일자별 상세 데이터', add_filter=True)

                        # 모든 시트에 '상품유형' 필터 추가
                        for sheet_name in ['기간별_성과', '캠페인유형별_성과']:
                            worksheet = writer.sheets[sheet_name]
                            df = dfs_to_save[sheet_name]
                            if '상품유형' in df.columns:
                                product_type_col = df.columns.get_loc('상품유형')
                                # 필터는 이미 추가되었으므로, 사용자가 엑셀에서 필터링할 수 있습니다.
                            else:
                                pass

                        # 차트 시트 생성 (옵션)
                        if include_charts:
                            chart_sheet = workbook.add_worksheet('차트')
                            chart_sheet.merge_range('A1:N1', '주요 지표 차트', title_format)
                            
                            # 비용/매출 차트
                            chart1 = workbook.add_chart({'type': 'line'})
                            chart1.add_series({
                                'name': '총비용',
                                'categories': f"='기간별_성과'!$A$3:$A${len(raw_group_df)+2}",
                                'values': f"='기간별_성과'!$C$3:$C${len(raw_group_df)+2}",
                                'line': {'color': 'blue'}
                            })
                            chart1.add_series({
                                'name': '전환매출액',
                                'categories': f"='기간별_성과'!$A$3:$A${len(raw_group_df)+2}",
                                'values': f"='기간별_성과'!$G$3:$G${len(raw_group_df)+2}",
                                'line': {'color': 'green'}
                            })
                            chart1.set_title({'name': '비용/매출 추이'})
                            chart1.set_x_axis({'name': date_option})
                            chart1.set_y_axis({'name': '금액 (원)'})
                            chart1.set_size({'width': 720, 'height': 480})
                            chart_sheet.insert_chart('A3', chart1)

                            # ROAS 차트
                            chart2 = workbook.add_chart({'type': 'column'})
                            chart2.add_series({
                                'name': 'ROAS',
                                'categories': f"='기간별_성과'!$A$3:$A${len(raw_group_df)+2}",
                                'values': f"='기간별_성과'!$M$3:$M${len(raw_group_df)+2}",
                                'fill': {'color': 'orange'}
                            })
                            chart2.set_title({'name': 'ROAS 추이'})
                            chart2.set_x_axis({'name': date_option})
                            chart2.set_y_axis({'name': 'ROAS (%)'})
                            chart2.set_size({'width': 720, 'height': 480})
                            chart_sheet.insert_chart('A25', chart2)

                        # 모든 시트 확대/축소 비율 설정
                        for worksheet in workbook.worksheets():
                            worksheet.set_zoom(85)

                    buffer.seek(0)
                    
                    # 다운로드 버튼
                    st.download_button(
                        label="📥 Excel 보고서 다운로드",
                        data=buffer,
                        file_name=f"{report_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.success('보고서가 성공적으로 생성되었습니다! 다운로드 버튼을 클릭하여 저장하세요.')

            except Exception as e:
                st.error(f"보고서 생성 중 오류가 발생했습니다: {str(e)}")
                st.write("오류 상세 정보:", e)

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {str(e)}")
        st.write("오류 상세 정보:", e)
        st.stop()

else:
    st.info("👈 좌측 사이드바에서 Excel 파일을 업로드하세요 ('raw 시트'와 'index 시트'가 포함된 파일)")

# 푸터 추가
st.markdown("""
<div style='position: fixed; bottom: 0; width: 100%; background-color: #f0f2f6; padding: 10px; text-align: center;'>
    © 2024 마케팅 대시보드. All rights reserved.
</div>
""", unsafe_allow_html=True)
