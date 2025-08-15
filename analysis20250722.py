import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
import io
import base64
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, 
                                Paragraph, Spacer)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
warnings.filterwarnings('ignore')

# 한글 폰트 등록 (Windows 기준)
try:
    pdfmetrics.registerFont(TTFont('NanumGothic', 
                                   'C:/Windows/Fonts/malgun.ttf'))
    KOREAN_FONT = 'NanumGothic'
except (OSError, FileNotFoundError):
    KOREAN_FONT = 'Helvetica'  # 폰트가 없을 경우 기본 폰트 사용

# 페이지 설정
st.set_page_config(
    page_title="관세법인우신 종합 분석 시스템(수입)",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(90deg, #f0f8ff, #e6f3ff);
        border-radius: 10px;
        border-left: 5px solid #1f4e79;
    }
    .tab-header {
        font-size: 1.8rem;
        font-weight: bold;
        margin-bottom: 1rem;
        padding: 0.5rem;
        border-radius: 8px;
    }
    .internal-tab { background-color: #e8f4fd; color: #1f4e79; }
    .client-tab { background-color: #fff2e8; color: #d68910; }
    .forwarder-tab { background-color: #e8f8f5; color: #239b56; }
    
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #4472c4;
    }
    .complexity-high { border-left-color: #e74c3c !important; }
    .complexity-medium { border-left-color: #f39c12 !important; }
    .complexity-low { border-left-color: #27ae60 !important; }
    
    .insight-box {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 4px solid #17a2b8;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class CustomsAnalyzer:
    """관세사 업무 종합 분석 클래스 (3차원 분석)"""
    
    def __init__(self, df, weights=None):
        self.df = df.copy()
        # 기본 가중치 설정
        self.weights = weights or {
            'lane_weight': 1.0,
            'spec_weight': 0.5,
            'requirement_weight': 10.0,
            'exemption_weight': 10.0,
            'fta_weight': 10.0,
            'transaction_weight': 5.0,
            'trader_weight': 5.0
        }
        self.prepare_data()
    
    def prepare_data(self):
        """데이터 전처리"""
        # 컬럼명 정리
        self.df.columns = self.df.columns.str.strip()
        
        # 디버깅: 사용 가능한 컬럼 확인
        print(f"사용 가능한 컬럼: {list(self.df.columns)}")
        
        # 날짜 처리 - 여러 날짜 컬럼 중 사용 가능한 것을 찾아서 사용
        date_success = False
        
        # 1순위: 수리일자
        if '수리일자' in self.df.columns:
            print(f"수리일자 컬럼 발견! 샘플 데이터: {self.df['수리일자'].head()}")
            
            self.df['수리일자_parsed'] = (
                self.df['수리일자'].apply(self.parse_date_string)
            )
            self.df['수리일자_parsed'] = pd.to_datetime(
                self.df['수리일자_parsed'], errors='coerce'
            )
            
            parsed_count = self.df['수리일자_parsed'].notna().sum()
            total_count = len(self.df)
            print(f"수리일자 파싱 성공: {parsed_count}/{total_count}개")
            
            if parsed_count > 0:
                date_success = True
                self.df['날짜_기준'] = self.df['수리일자_parsed']
                print("✅ 수리일자를 기준으로 사용합니다.")
        
        # 2순위: 신고일자 (수리일자가 없거나 비어있는 경우)
        if not date_success and '신고일자' in self.df.columns:
            print(f"신고일자 컬럼 확인! 샘플 데이터: {self.df['신고일자'].head()}")
            
            self.df['신고일자_parsed'] = (
                self.df['신고일자'].apply(self.parse_date_string)
            )
            self.df['신고일자_parsed'] = pd.to_datetime(
                self.df['신고일자_parsed'], errors='coerce'
            )
            
            parsed_count = self.df['신고일자_parsed'].notna().sum()
            print(f"신고일자 파싱 성공: {parsed_count}/{total_count}개")
            
            if parsed_count > 0:
                date_success = True
                self.df['날짜_기준'] = self.df['신고일자_parsed']
                print("✅ 신고일자를 기준으로 사용합니다.")
        
        # 3순위: 입력일시 (다른 날짜들이 없는 경우)
        if not date_success and '입력일시' in self.df.columns:
            print(f"입력일시 컬럼 확인! 샘플 데이터: {self.df['입력일시'].head()}")
            
            self.df['입력일시_parsed'] = pd.to_datetime(
                self.df['입력일시'], errors='coerce'
            )
            
            parsed_count = self.df['입력일시_parsed'].notna().sum()
            print(f"입력일시 파싱 성공: {parsed_count}/{total_count}개")
            
            if parsed_count > 0:
                date_success = True
                self.df['날짜_기준'] = self.df['입력일시_parsed']
                print("✅ 입력일시를 기준으로 사용합니다.")
        
        # 날짜 기준이 있는 경우 요일 계산
        if date_success and '날짜_기준' in self.df.columns:
            self.df['요일'] = self.df['날짜_기준'].dt.day_name()
            self.df['요일_한글'] = self.df['날짜_기준'].dt.dayofweek.map({
                0: '월요일', 1: '화요일', 2: '수요일', 3: '목요일', 4: '금요일',
                5: '토요일', 6: '일요일'
            })
            
            # 디버깅: 요일별 건수 확인
            weekday_counts = self.df['요일_한글'].value_counts()
            print(f"✅ 요일별 건수: {weekday_counts.to_dict()}")
        else:
            # 모든 날짜가 유효하지 않은 경우 기본값 설정
            self.df['요일'] = None
            self.df['요일_한글'] = None
            print("❌ 경고: 모든 날짜 컬럼이 유효하지 않습니다!")
            
        # 호환성을 위해 기존 컬럼들도 유지
        if '신고일자' in self.df.columns and '신고일자_parsed' not in self.df.columns:
            self.df['신고일자_parsed'] = (
                self.df['신고일자'].apply(self.parse_date_string)
            )
            self.df['신고일자_parsed'] = pd.to_datetime(
                self.df['신고일자_parsed'], errors='coerce'
            )
    
    def parse_date_string(self, date_str):
        """날짜 문자열 파싱 (다양한 형식 지원)"""
        if pd.isna(date_str):
            return None
        
        try:
            date_str = str(date_str).strip()
            
            # 이미 datetime 형식인 경우
            if isinstance(date_str, str) and 'Timestamp' in str(type(date_str)):
                return pd.to_datetime(date_str)
            
            # YYYYMMDD 형식 (8자리)
            if len(date_str) == 8 and date_str.isdigit():
                year, month, day = date_str[:4], date_str[4:6], date_str[6:8]
                return pd.to_datetime(f"{year}-{month}-{day}")
            
            # YYYY-MM-DD 형식
            if '-' in date_str and len(date_str) == 10:
                return pd.to_datetime(date_str)
            
            # YYYY/MM/DD 형식
            if '/' in date_str:
                return pd.to_datetime(date_str)
            
            # 그외 형식은 pandas에게 맡김
            return pd.to_datetime(date_str, errors='coerce')
            
        except Exception:
            return None
    
    def update_weights(self, weights):
        """가중치 업데이트"""
        self.weights.update(weights)
    
    def calculate_complexity_score(self, data, group_col='신고번호'):
        """7차원 복잡도 점수 계산 (가중치 적용)"""
        # 신고번호별 그룹화 (중복제거)
        decl_grouped = data.groupby(group_col).agg({
            '총란수': 'first',
            '총규격수': 'first', 
            '관세감면구분': 'first',
            '원산지증명유무': 'first',
            '거래구분': 'nunique',  # 거래구분 종류 수
            '무역거래처상호': 'nunique',  # 무역거래처 종류 수
            '발급서류명': lambda x: x.notna().sum()  # 수입요건 수
        }).reset_index()
        
        # 복잡도 계산
        complexity_scores = []
        
        for _, row in decl_grouped.iterrows():
            score = 0
            
            # 1. 총란수 (가중치 적용)
            score += (row.get('총란수', 0) * self.weights['lane_weight']) if pd.notna(row.get('총란수', 0)) else 0
            
            # 2. 총규격수 (가중치 적용)
            score += (row.get('총규격수', 0) * self.weights['spec_weight']) if pd.notna(row.get('총규격수', 0)) else 0
            
            # 3. 수입요건 수 (가중치 적용)
            score += row.get('발급서류명', 0) * self.weights['requirement_weight']
            
            # 4. 감면 적용 (가중치 적용)
            if pd.notna(row.get('관세감면구분')) and str(row.get('관세감면구분')).strip():
                score += self.weights['exemption_weight']
                
            # 5. 원산지증명 (가중치 적용)
            if row.get('원산지증명유무') == 'Y':
                score += self.weights['fta_weight']
            
            # 6. 거래구분 종류 수 (가중치 적용)
            score += row.get('거래구분', 1) * self.weights['transaction_weight']
            
            # 7. 무역거래처 종류 수 (가중치 적용)  
            score += row.get('무역거래처상호', 1) * self.weights['trader_weight']
            
            complexity_scores.append(score)
        
        return np.mean(complexity_scores) if complexity_scores else 0
    
    def analyze_by_author(self):
        """작성자별 분석 (내부 관리용)"""
        if '작성자' not in self.df.columns:
            return pd.DataFrame()
        
        valid_data = self.df[self.df['작성자'].notna() & (self.df['작성자'] != '')]
        results = []
        
        for author in valid_data['작성자'].unique():
            author_data = valid_data[valid_data['작성자'] == author]
            
            # 기본 통계
            total_items = len(author_data)
            decl_grouped = author_data.groupby('신고번호').first().reset_index()
            unique_declarations = len(decl_grouped)
            
            # 복잡도 점수 계산
            complexity_score = self.calculate_complexity_score(author_data)
            
            # 기타 통계들
            total_lanes = decl_grouped['총란수'].fillna(0).astype(int).sum()
            total_specs = decl_grouped['총규격수'].fillna(0).astype(int).sum()
            
            # 수입요건 분석
            requirement_count = author_data[author_data['발급서류명'].notna()]['신고번호'].nunique()
            
            # FTA 분석
            fta_count = decl_grouped[decl_grouped['원산지증명유무'] == 'Y']
            fta_rate = len(fta_count) / unique_declarations * 100 if unique_declarations > 0 else 0
            
            # 감면 분석
            exemption_count = decl_grouped[
                decl_grouped['관세감면구분'].notna() & 
                (decl_grouped['관세감면구분'].astype(str).str.strip() != '')
            ]
            exemption_rate = (
                len(exemption_count) / unique_declarations * 100
                if unique_declarations > 0 else 0
            )
            
            # 거래구분 및 무역거래처 분석
            transaction_types = decl_grouped['거래구분'].nunique()
            trader_count = decl_grouped['무역거래처상호'].nunique()
            importer_count = decl_grouped['납세자상호'].nunique()
            
            # 요일별 분석
            weekday_stats = {}
            if '요일_한글' in author_data.columns:
                weekday_data = author_data[author_data['요일_한글'].notna()]
                weekday_grouped = weekday_data.groupby('요일_한글')['신고번호'].nunique()
                for day in ['월요일', '화요일', '수요일', '목요일', '금요일']:
                    weekday_stats[day] = weekday_grouped.get(day, 0)
            
            results.append({
                '작성자': author,
                '총처리건수': total_items,
                '고유신고번호수': unique_declarations,
                '복잡도점수': round(complexity_score, 1),
                '총란수합계': total_lanes,
                '총규격수합계': total_specs,
                '수입요건신고번호수': requirement_count,
                'FTA활용률': round(fta_rate, 1),
                '관세감면적용률': round(exemption_rate, 1),
                '거래구분종류수': transaction_types,
                '무역거래처수': trader_count,
                '담당수입자수': importer_count,
                '평균품목수_신고서': (
                    round(total_items / unique_declarations, 1)
                    if unique_declarations > 0 else 0
                ),
                '월요일': weekday_stats.get('월요일', 0),
                '화요일': weekday_stats.get('화요일', 0),
                '수요일': weekday_stats.get('수요일', 0),
                '목요일': weekday_stats.get('목요일', 0),
                '금요일': weekday_stats.get('금요일', 0)
            })
        
        result_df = pd.DataFrame(results)
        if not result_df.empty:
            result_df = result_df.sort_values('복잡도점수', ascending=False)
        
        return result_df
    
    def analyze_by_importer(self):
        """수입자별 분석 (고객사 관리용)"""
        if '납세자상호' not in self.df.columns:
            return pd.DataFrame()
        
        valid_data = self.df[self.df['납세자상호'].notna() & (self.df['납세자상호'] != '')]
        results = []
        
        for importer in valid_data['납세자상호'].unique():
            importer_data = valid_data[valid_data['납세자상호'] == importer]
            
            # 기본 통계
            total_items = len(importer_data)
            decl_grouped = importer_data.groupby('신고번호').first().reset_index()
            unique_declarations = len(decl_grouped)
            
            # 복잡도 점수 계산
            complexity_score = self.calculate_complexity_score(importer_data)
            
            # 담당 작성자 분석
            author_counts = importer_data['작성자'].value_counts()
            main_author = author_counts.index[0] if len(author_counts) > 0 else ''
            author_diversity = len(author_counts)
            
            # 업종 특성 분석 (발급서류 패턴)
            document_types = importer_data[importer_data['발급서류명'].notna()]['발급서류명'].nunique()
            requirement_rate = len(importer_data[importer_data['발급서류명'].notna()]['신고번호'].unique()) / unique_declarations * 100 if unique_declarations > 0 else 0
            
            # FTA 및 감면 활용 분석
            fta_count = decl_grouped[decl_grouped['원산지증명유무'] == 'Y']
            fta_rate = len(fta_count) / unique_declarations * 100 if unique_declarations > 0 else 0
            
            exemption_count = decl_grouped[
                decl_grouped['관세감면구분'].notna() & 
                (decl_grouped['관세감면구분'].astype(str).str.strip() != '')
            ]
            exemption_rate = (
                len(exemption_count) / unique_declarations * 100
                if unique_declarations > 0 else 0
            )
            
            # 거래구분 다양성
            transaction_types = decl_grouped['거래구분'].nunique()
            
            # 요일별 패턴
            weekday_stats = {}
            if '요일_한글' in importer_data.columns:
                weekday_data = importer_data[importer_data['요일_한글'].notna()]
                weekday_grouped = weekday_data.groupby('요일_한글')['신고번호'].nunique()
                for day in ['월요일', '화요일', '수요일', '목요일', '금요일']:
                    weekday_stats[day] = weekday_grouped.get(day, 0)
            
            results.append({
                '수입자': importer,
                '총처리건수': total_items,
                '고유신고번호수': unique_declarations,
                '복잡도점수': round(complexity_score, 1),
                '주담당작성자': main_author,
                '담당작성자수': author_diversity,
                '발급서류종류수': document_types,
                '수입요건비율': round(requirement_rate, 1),
                'FTA활용률': round(fta_rate, 1),
                '관세감면활용률': round(exemption_rate, 1),
                '거래구분다양성': transaction_types,
                '평균품목수_신고서': (
                    round(total_items / unique_declarations, 1)
                    if unique_declarations > 0 else 0
                ),
                '월요일': weekday_stats.get('월요일', 0),
                '화요일': weekday_stats.get('화요일', 0),
                '수요일': weekday_stats.get('수요일', 0),
                '목요일': weekday_stats.get('목요일', 0),
                '금요일': weekday_stats.get('금요일', 0)
            })
        
        result_df = pd.DataFrame(results)
        if not result_df.empty:
            result_df = result_df.sort_values('총처리건수', ascending=False)
        
        return result_df
    
    def analyze_by_forwarder(self):
        """운송주선인별 분석 (포워딩 관리용)"""
        if '운송주선인상호' not in self.df.columns:
            return pd.DataFrame()
        
        valid_data = self.df[self.df['운송주선인상호'].notna() & (self.df['운송주선인상호'] != '')]
        results = []
        
        for forwarder in valid_data['운송주선인상호'].unique():
            forwarder_data = valid_data[valid_data['운송주선인상호'] == forwarder]
            
            # 기본 통계
            total_items = len(forwarder_data)
            decl_grouped = forwarder_data.groupby('신고번호').first().reset_index()
            unique_declarations = len(decl_grouped)
            
            # 복잡도 점수 계산
            complexity_score = self.calculate_complexity_score(forwarder_data)
            
            # 담당 작성자 분석 (중복제거 신고번호 기준)
            author_counts = decl_grouped['작성자'].value_counts()
            main_author = author_counts.index[0] if len(author_counts) > 0 else ''
            author_diversity = len(author_counts)
            
            # 담당 수입자 분석 (중복제거 신고번호 기준)
            importer_counts = decl_grouped['납세자상호'].value_counts()
            main_importer = importer_counts.index[0] if len(importer_counts) > 0 else ''
            importer_count = len(importer_counts)
            
            # 무역거래처 다양성
            trader_count = decl_grouped['무역거래처상호'].nunique()
            
            # 통관 효율성
            avg_lanes = decl_grouped['총란수'].fillna(0).mean()
            avg_specs = decl_grouped['총규격수'].fillna(0).mean()
            
            # FTA 및 감면 분석
            fta_count = decl_grouped[decl_grouped['원산지증명유무'] == 'Y']
            fta_rate = len(fta_count) / unique_declarations * 100 if unique_declarations > 0 else 0
            
            exemption_count = decl_grouped[
                decl_grouped['관세감면구분'].notna() & 
                (decl_grouped['관세감면구분'].astype(str).str.strip() != '')
            ]
            exemption_rate = (
                len(exemption_count) / unique_declarations * 100
                if unique_declarations > 0 else 0
            )
            
            # 요일별 패턴
            weekday_stats = {}
            if '요일_한글' in forwarder_data.columns:
                weekday_data = forwarder_data[forwarder_data['요일_한글'].notna()]
                weekday_grouped = weekday_data.groupby('요일_한글')['신고번호'].nunique()
                for day in ['월요일', '화요일', '수요일', '목요일', '금요일']:
                    weekday_stats[day] = weekday_grouped.get(day, 0)
            
            results.append({
                '운송주선인': forwarder,
                '총처리건수': total_items,
                '고유신고번호수': unique_declarations,
                '복잡도점수': round(complexity_score, 1),
                '주담당작성자': main_author,
                '담당작성자수': author_diversity,
                '주요수입자': main_importer,
                '연결수입자수': importer_count,
                '연결무역거래처수': trader_count,
                '평균란수_신고서': round(avg_lanes, 1),
                '평균규격수_신고서': round(avg_specs, 1),
                'FTA활용률': round(fta_rate, 1),
                '관세감면활용률': round(exemption_rate, 1),
                '평균품목수_신고서': (
                    round(total_items / unique_declarations, 1)
                    if unique_declarations > 0 else 0
                ),
                '월요일': weekday_stats.get('월요일', 0),
                '화요일': weekday_stats.get('화요일', 0),
                '수요일': weekday_stats.get('수요일', 0),
                '목요일': weekday_stats.get('목요일', 0),
                '금요일': weekday_stats.get('금요일', 0)
            })
        
        result_df = pd.DataFrame(results)
        if not result_df.empty:
            result_df = result_df.sort_values('총처리건수', ascending=False)
        
        return result_df
    
    def analyze_cs_inspection(self):
        """C/S검사구분별 분석"""
        if 'C/S검사구분' not in self.df.columns:
            return pd.DataFrame(), {}
        
        valid_data = self.df[self.df['C/S검사구분'].notna()]
        
        # 신고번호 기준으로 중복제거 (더 명시적)
        unique_declarations = valid_data.groupby('신고번호').first().reset_index()
        
        # 검사구분별 신고번호 개수 (중복제거 완료)
        cs_analysis = unique_declarations.groupby('C/S검사구분').size().reset_index(name='신고번호수')
        cs_analysis.columns = ['검사구분', '신고번호수']
        cs_analysis = cs_analysis.sort_values('신고번호수', ascending=False)
        
        # 검사구분 매핑
        inspection_mapping = {
            'Y': '세관검사',
            'F': '협엄검사', 
            'N': '무검사',
            'C': '서류검사',
            'S': '표본검사'
        }
        
        cs_analysis['검사유형'] = cs_analysis['검사구분'].map(inspection_mapping).fillna('기타')
        
        # 통계 요약 (중복제거된 데이터 기준)
        total_declarations = len(unique_declarations)
        stats_summary = {
            '총신고번호수': total_declarations,
            '검사구분종류': len(cs_analysis),
            '가장많은검사': cs_analysis.iloc[0]['검사유형'] if not cs_analysis.empty else '',
            '무검사율': round(cs_analysis[cs_analysis['검사구분'] == 'N']['신고번호수'].sum() / total_declarations * 100, 1) if total_declarations > 0 else 0
        }
        
        return cs_analysis, stats_summary

def create_weekday_chart(df, title, entity_col, entity_name=None):
    """요일별 차트 생성"""
    if entity_name:
        entity_data = df[df[entity_col] == entity_name]
    else:
        entity_data = df.head(10)  # 상위 10개
    
    weekdays = ['월요일', '화요일', '수요일', '목요일', '금요일']
    
    fig = go.Figure()
    
    for _, row in entity_data.iterrows():
        values = [row.get(day, 0) for day in weekdays]
        fig.add_trace(go.Scatter(
            x=weekdays,
            y=values,
            mode='lines+markers',
            name=row[entity_col],
            line=dict(width=3),
            marker=dict(size=8)
        ))
    
    fig.update_layout(
        title=title,
        xaxis_title="요일",
        yaxis_title="신고번호 수",
        height=400,
        hovermode='x unified'
    )
    
    return fig

def create_complexity_distribution(df, entity_col):
    """복잡도 분포 차트"""
    fig = px.histogram(
        df,
        x='복잡도점수',
        title=f"{entity_col} 복잡도 점수 분포",
        nbins=20,
        color_discrete_sequence=['#3498db']
    )
    fig.update_layout(height=400)
    return fig

def create_top_entities_chart(df, entity_col, metric='총처리건수', top_n=10):
    """상위 엔티티 차트"""
    top_df = df.head(top_n)
    
    fig = px.bar(
        top_df,
        x=entity_col,
        y=metric,
        title=f"상위 {top_n}개 {entity_col} - {metric}",
        color=metric,
        text=metric
    )
    fig.update_traces(texttemplate='%{text:,}', textposition='outside')
    fig.update_layout(height=500, xaxis_tickangle=-45)
    return fig

def create_excel_download(df, filename):
    """엑셀 파일 다운로드 생성 (한글 깨짐 방지)"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='분석결과')
        # 워크시트 가져오기
        worksheet = writer.sheets['분석결과']
        # 컬럼 너비 자동 조정
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()


def create_pdf_download(df, title, filename):
    """PDF 파일 다운로드 생성 (가로 형식, 자동 크기 조정)"""
    buffer = io.BytesIO()
    # 가로 형식(Landscape) 페이지 설정
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), 
                            leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    # 스타일 설정
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=KOREAN_FONT,
        fontSize=14,
        spaceAfter=20,
        alignment=1  # 중앙 정렬
    )
    
    # 제목 추가
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 15))
    
    # 데이터를 테이블로 변환
    if not df.empty:
        # 컬럼명을 한글로 변환 (간소화)
        column_mapping = {
            '작성자': '작성자',
            '수입자': '수입자', 
            '운송주선인': '운송주선인',
            '총처리건수': '총처리건수',
            '고유신고번호수': '신고번호수',
            '복잡도점수': '복잡도',
            'FTA활용률': 'FTA(%)',
            '관세감면적용률': '감면(%)',
            '관세감면활용률': '감면(%)',
            '담당수입자수': '담당수입자',
            '주담당작성자': '주담당자',
            '담당작성자수': '담당자수',
            '발급서류종류수': '서류종류',
            '수입요건비율': '요건(%)',
            '거래구분다양성': '거래구분',
            '평균품목수_신고서': '평균품목',
            '주요수입자': '주요수입자',
            '연결수입자수': '연결수입자',
            '연결무역거래처수': '연결거래처',
            '평균란수_신고서': '평균란수',
            '평균규격수_신고서': '평균규격',
            '검사구분': '검사구분',
            '검사유형': '검사유형',
            '신고번호수': '신고번호수',
            '비율(%)': '비율(%)',
            '누적비율(%)': '누적(%)'
        }
        
        # 컬럼명 변환
        df_display = df.copy()
        df_display.columns = [column_mapping.get(col, col) for col in df_display.columns]
        
        # 중요한 컬럼만 선택 (PDF 폭에 맞게)
        important_cols = []
        for col in df_display.columns:
            if col in ['작성자', '수입자', '운송주선인', '총처리건수', '신고번호수', '복잡도', 
                      'FTA(%)', '감면(%)', '담당수입자', '주담당자', '검사유형']:
                important_cols.append(col)
        
        # 중요 컬럼이 없으면 처음 8개 컬럼 사용
        if not important_cols:
            important_cols = df_display.columns[:8].tolist()
        
        # 선택된 컬럼만 사용
        df_display = df_display[important_cols]
        
        # 데이터 준비 (상위 25개)
        data = [df_display.columns.tolist()]  # 헤더
        for _, row in df_display.head(25).iterrows():
            formatted_row = []
            for val in row.values:
                if isinstance(val, float) and not pd.isna(val):
                    if val > 100:
                        formatted_row.append(f"{val:,.0f}")
                    else:
                        formatted_row.append(f"{val:.1f}")
                else:
                    formatted_row.append(str(val) if not pd.isna(val) else "")
            data.append(formatted_row)
        
        # 가로 페이지 크기 계산 (landscape A4: 842x595 포인트)
        page_width = landscape(A4)[0] - 60  # 여백 제외
        col_count = len(important_cols)
        col_width = page_width / col_count
        
        # 테이블 생성 (자동 크기 조정)
        table = Table(data, colWidths=[col_width] * col_count)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), '#4472c4'),
            ('TEXTCOLOR', (0, 0), (-1, 0), 'white'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), KOREAN_FONT),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # 헤더 폰트 크기 감소
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), '#f8f9fa'),
            ('FONTNAME', (0, 1), (-1, -1), KOREAN_FONT),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # 데이터 폰트 크기 감소
            ('GRID', (0, 0), (-1, -1), 0.5, 'black'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ]))
        
        elements.append(table)
        
        # 컬럼 정보 추가
        elements.append(Spacer(1, 15))
        info_style = ParagraphStyle(
            'InfoStyle',
            parent=styles['Normal'],
            fontName=KOREAN_FONT,
            fontSize=8,
            alignment=0
        )
        
        info_text = f"📊 총 {len(df)}개 레코드 중 상위 25개 표시 | 표시 컬럼: {', '.join(important_cols)}"
        elements.append(Paragraph(info_text, info_style))
    
    # PDF 생성
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def create_html_download(df, title, filename):
    """HTML 파일 다운로드 생성"""
    # 스타일이 포함된 HTML 생성
    html_content = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <style>
            body {{
                font-family: 'Malgun Gothic', '맑은 고딕', sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #1f4e79;
                text-align: center;
                border-bottom: 3px solid #4472c4;
                padding-bottom: 10px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
                font-size: 12px;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 8px;
                text-align: center;
            }}
            th {{
                background-color: #4472c4;
                color: white;
                font-weight: bold;
            }}
            tr:nth-child(even) {{
                background-color: #f8f9fa;
            }}
            tr:hover {{
                background-color: #e3f2fd;
            }}
            .info {{
                background-color: #e8f4fd;
                border-left: 4px solid #17a2b8;
                padding: 15px;
                margin: 20px 0;
                border-radius: 5px;
            }}
            .footer {{
                text-align: center;
                margin-top: 30px;
                color: #666;
                font-size: 12px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{title}</h1>
            <div class="info">
                <strong>생성일시:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>
                <strong>총 레코드 수:</strong> {len(df)}개<br>
                <strong>컬럼 수:</strong> {len(df.columns)}개
            </div>
            {df.to_html(index=False, classes='dataframe', escape=False) if not df.empty else '<p>데이터가 없습니다.</p>'}
            <div class="footer">
                관세법인 우신 종합 분석 시스템 v2.0<br>
                생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_content.encode('utf-8')


def get_download_link(data, filename, file_type):
    """다운로드 링크 생성"""
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/{file_type};base64,{b64}" download="{filename}">📥 {filename} 다운로드</a>'


def main():
    """메인 애플리케이션"""
    
    # 헤더
    st.markdown('<div class="main-header">🏢 관세법인 우신 종합 분석 시스템</div>', 
                unsafe_allow_html=True)
    
    # 사이드바
    with st.sidebar:
        st.header("📁 파일 업로드")
        uploaded_file = st.file_uploader(
            "엑셀 파일을 선택하세요",
            type=['xlsx', 'xls'],
            help="수입신고 데이터 엑셀 파일을 업로드해주세요"
        )
        
        if uploaded_file:
            st.success("✅ 파일 업로드 완료!")
            st.info(f"📄 파일명: {uploaded_file.name}")
            st.info(f"📏 파일 크기: {uploaded_file.size / 1024:.1f} KB")
            
            # 파일 정보 표시
            with st.expander("📋 파일 정보"):
                st.write(f"**파일명**: {uploaded_file.name}")
                st.write(f"**파일 크기**: {uploaded_file.size / 1024:.1f} KB")
                st.write(f"**파일 타입**: {uploaded_file.type}")
                st.write(f"**업로드 시간**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            st.warning("⚠️ 분석할 엑셀 파일을 업로드해주세요")
            
        st.markdown("---")
        
        # 복잡도 점수 설정
        st.header("⚙️ 복잡도 점수 설정")
        with st.expander("🔧 점수 기준 커스터마이징", expanded=False):
            st.markdown("**7차원 복잡도 가중치 설정**")
            
            # 기본값 설정
            col1, col2 = st.columns(2)
            
            with col1:
                lane_weight = st.number_input(
                    "총란수 가중치", 
                    min_value=0.1, max_value=10.0, 
                    value=st.session_state.get('lane_weight', 1.0), 
                    step=0.1,
                    key='lane_weight',
                    help="신고서의 총 란수 1개당 점수"
                )
                
                spec_weight = st.number_input(
                    "총규격수 가중치", 
                    min_value=0.1, max_value=5.0, 
                    value=st.session_state.get('spec_weight', 0.5), 
                    step=0.1,
                    key='spec_weight',
                    help="신고서의 총 규격수 1개당 점수"
                )
                
                requirement_weight = st.number_input(
                    "수입요건 가중치", 
                    min_value=1.0, max_value=50.0, 
                    value=st.session_state.get('requirement_weight', 10.0), 
                    step=1.0,
                    key='requirement_weight',
                    help="수입요건 서류 1개당 점수"
                )
                
                exemption_weight = st.number_input(
                    "관세감면 가중치", 
                    min_value=1.0, max_value=50.0, 
                    value=st.session_state.get('exemption_weight', 10.0), 
                    step=1.0,
                    key='exemption_weight',
                    help="관세감면 적용 시 점수"
                )
            
            with col2:
                fta_weight = st.number_input(
                    "FTA 가중치", 
                    min_value=1.0, max_value=50.0, 
                    value=st.session_state.get('fta_weight', 10.0), 
                    step=1.0,
                    key='fta_weight',
                    help="FTA 원산지증명 적용 시 점수"
                )
                
                transaction_weight = st.number_input(
                    "거래구분 가중치", 
                    min_value=1.0, max_value=20.0, 
                    value=st.session_state.get('transaction_weight', 5.0), 
                    step=1.0,
                    key='transaction_weight',
                    help="거래구분 종류 1개당 점수"
                )
                
                trader_weight = st.number_input(
                    "무역거래처 가중치", 
                    min_value=1.0, max_value=20.0, 
                    value=st.session_state.get('trader_weight', 5.0), 
                    step=1.0,
                    key='trader_weight',
                    help="무역거래처 종류 1개당 점수"
                )
            
            # 복잡도 분류 기준 설정
            st.markdown("**복잡도 분류 기준 설정**")
            col1, col2 = st.columns(2)
            
            with col1:
                low_threshold = st.number_input(
                    "일반업무 상한선", 
                    min_value=50, max_value=500, value=100, step=10,
                    help="이 점수 미만은 일반업무로 분류"
                )
            
            with col2:
                high_threshold = st.number_input(
                    "고복잡도 하한선", 
                    min_value=100, max_value=1000, value=200, step=10,
                    help="이 점수 이상은 고복잡도로 분류"
                )
            
            # 가중치 프로필
            st.markdown("**프리셋 프로필**")
            profile_col1, profile_col2, profile_col3 = st.columns(3)
            
            with profile_col1:
                if st.button("🎯 표준 프로필"):
                    st.session_state.lane_weight = 1.0
                    st.session_state.spec_weight = 0.5
                    st.session_state.requirement_weight = 10.0
                    st.session_state.exemption_weight = 10.0
                    st.session_state.fta_weight = 10.0
                    st.session_state.transaction_weight = 5.0
                    st.session_state.trader_weight = 5.0
                    st.success("표준 프로필 적용!")
                    st.rerun()
            
            with profile_col2:
                if st.button("⚡ 효율성 중심"):
                    st.session_state.lane_weight = 0.5
                    st.session_state.spec_weight = 0.3
                    st.session_state.requirement_weight = 5.0
                    st.session_state.exemption_weight = 5.0
                    st.session_state.fta_weight = 5.0
                    st.session_state.transaction_weight = 2.0
                    st.session_state.trader_weight = 2.0
                    st.success("효율성 중심 적용!")
                    st.rerun()
            
            with profile_col3:
                if st.button("🔥 전문성 중심"):
                    st.session_state.lane_weight = 2.0
                    st.session_state.spec_weight = 1.0
                    st.session_state.requirement_weight = 20.0
                    st.session_state.exemption_weight = 20.0
                    st.session_state.fta_weight = 20.0
                    st.session_state.transaction_weight = 10.0
                    st.session_state.trader_weight = 10.0
                    st.success("전문성 중심 적용!")
                    st.rerun()
            
            # 현재 설정 요약
            st.markdown("**📊 현재 설정 요약**")
            st.write(f"- 총란수: **{lane_weight}점/개**")
            st.write(f"- 총규격수: **{spec_weight}점/개**")
            st.write(f"- 수입요건: **{requirement_weight}점/개**")
            st.write(f"- 관세감면: **{exemption_weight}점**")
            st.write(f"- FTA활용: **{fta_weight}점**")
            st.write(f"- 거래구분: **{transaction_weight}점/종류**")
            st.write(f"- 무역거래처: **{trader_weight}점/종류**")
            
            st.markdown(f"**분류 기준**: 일반({low_threshold}점 미만) / 중간({low_threshold}-{high_threshold}점) / 고복잡도({high_threshold}점 이상)")
            
            # 실시간 복잡도 재계산 스위치
            use_custom_weights = st.toggle(
                "🔄 커스텀 가중치 적용", 
                value=False, 
                help="체크하면 위의 설정값으로 복잡도를 재계산합니다"
            )
            
            if use_custom_weights:
                st.info("💡 커스텀 가중치가 적용됩니다. 분석 결과가 실시간으로 업데이트됩니다.")
                
                # 가중치 업데이트 및 분석 재실행
                updated_weights = {
                    'lane_weight': lane_weight,
                    'spec_weight': spec_weight,
                    'requirement_weight': requirement_weight,
                    'exemption_weight': exemption_weight,
                    'fta_weight': fta_weight,
                    'transaction_weight': transaction_weight,
                    'trader_weight': trader_weight
                }
                
                analyzer.update_weights(updated_weights)
                
                # 분석 재실행
                with st.spinner("🔄 커스텀 가중치로 분석을 재실행 중..."):
                    author_df = analyzer.analyze_by_author()
                    importer_df = analyzer.analyze_by_importer()
                    forwarder_df = analyzer.analyze_by_forwarder()
                    cs_df, cs_stats = analyzer.analyze_cs_inspection()
                
                st.success("✅ 커스텀 가중치로 분석이 완료되었습니다!")
        
        st.markdown("---")
        st.header("📊 분석 개요")
        st.markdown("""
        **🎯 4대 분석 축**
        
        1️⃣ **작성자별 분석**
        - 내부 직원 관리
        - 업무 효율성 평가
        - 7차원 복잡도 분석
        
        2️⃣ **수입자별 분석**  
        - 고객사 관리
        - 서비스 최적화
        - 업종별 특성
        
        3️⃣ **운송주선인별 분석**
        - 포워딩 파트너 관리
        - 물류 효율성
        - 작성자/수입자 매칭
        
        4️⃣ **검사구분 분석**
        - C/S검사구분별 통계
        - 검사패턴 분석
        - 위험도 관리
        """)
    
    if uploaded_file is None:
        st.info("👆 사이드바에서 분석할 엑셀 파일을 업로드해주세요.")
        
        with st.expander("📋 시스템 기능 상세"):
            st.markdown("""
            ### 🎯 7차원 복잡도 분석 엔진
            **복잡도 = 총란수 + 총규격수 + 수입요건수 + 감면여부 + 원산지증명여부 + 거래구분종류 + 무역거래처종류**
            
            ### 📊 4차원 통합 분석
            - **내부 관리**: 작성자별 업무 분배 최적화
            - **고객 관리**: 수입자별 맞춤 서비스 
            - **파트너 관리**: 운송주선인별 협력 효율화
            - **검사 관리**: C/S검사구분별 패턴 분석
            
            ### ⏰ 요일별 패턴 분석  
            - 업무량 예측 및 인력 배치
            - 고객별 신고 패턴 파악
            - 효율적 스케줄링
            
            ### 🔍 포워더 상세 분석
            - 담당 작성자/수입자 매칭 (중복제거 기준)
            - 네트워크 관계 시각화
            - 협력 최적화 인사이트
            """)
        return
    
    # 데이터 로딩
    try:
        with st.spinner("📊 데이터를 분석 중입니다..."):
            # 파일 로딩 진행상황 표시
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("📁 엑셀 파일 읽는 중...")
            df = pd.read_excel(uploaded_file)
            progress_bar.progress(25)
            
            status_text.text("🔧 데이터 전처리 중...")
            
            # 가중치 설정
            weights = {
                'lane_weight': lane_weight,
                'spec_weight': spec_weight,
                'requirement_weight': requirement_weight,
                'exemption_weight': exemption_weight,
                'fta_weight': fta_weight,
                'transaction_weight': transaction_weight,
                'trader_weight': trader_weight
            }
            
            analyzer = CustomsAnalyzer(df, weights)
            progress_bar.progress(50)
            
            status_text.text("📈 작성자별 분석 실행 중...")
            author_df = analyzer.analyze_by_author()
            progress_bar.progress(65)
            
            status_text.text("🏭 수입자별 분석 실행 중...")
            importer_df = analyzer.analyze_by_importer()
            progress_bar.progress(80)
            
            status_text.text("🚛 운송주선인별 분석 실행 중...")
            forwarder_df = analyzer.analyze_by_forwarder()
            progress_bar.progress(90)
            
            status_text.text("📋 검사구분 분석 실행 중...")
            cs_df, cs_stats = analyzer.analyze_cs_inspection()
            progress_bar.progress(100)
            
            # 진행상황 완료
            status_text.text("✅ 분석 완료!")
            st.success("🎉 데이터 분석이 성공적으로 완료되었습니다!")
            
            # 잠시 후 진행상황 바 제거
            import time
            time.sleep(1)
            progress_bar.empty()
            status_text.empty()
        
        # 데이터 검증 및 디버깅 정보 표시
        st.info(f"📊 로드된 데이터 정보: {len(df)}행, {len(df.columns)}열")
        
        # 컬럼 정보 표시 (디버깅용)
        with st.expander("🔍 데이터 컬럼 정보"):
            st.write("**데이터 컬럼 목록:**")
            for i, col in enumerate(df.columns, 1):
                st.write(f"{i}. {col}")
            
            st.write("**데이터 미리보기:**")
            st.dataframe(df.head(), use_container_width=True)
        
        if author_df.empty and importer_df.empty and forwarder_df.empty and cs_df.empty:
            st.error("❌ 분석할 수 있는 데이터가 없습니다. 파일 형식을 확인해주세요.")
            
            # 디버깅 정보 추가
            st.warning("🔍 디버깅 정보:")
            st.write(f"- 작성자 컬럼 존재: {'작성자' in df.columns}")
            st.write(f"- 수입자 컬럼 존재: {'납세자상호' in df.columns}")
            st.write(f"- 포워더 컬럼 존재: {'운송주선인상호' in df.columns}")
            st.write(f"- 검사구분 컬럼 존재: {'C/S검사구분' in df.columns}")
            
            if '작성자' in df.columns:
                st.write(f"- 작성자 데이터 샘플: {df['작성자'].dropna().head().tolist()}")
            if '납세자상호' in df.columns:
                st.write(f"- 수입자 데이터 샘플: {df['납세자상호'].dropna().head().tolist()}")
            if '운송주선인상호' in df.columns:
                st.write(f"- 포워더 데이터 샘플: {df['운송주선인상호'].dropna().head().tolist()}")
            
            return
            
    except Exception as e:
        st.error(f"❌ 파일 처리 중 오류가 발생했습니다: {str(e)}")
        
        # 상세 오류 정보 표시
        with st.expander("🔍 상세 오류 정보"):
            import traceback
            st.code(traceback.format_exc())
        return
    
    # 전체 요약 통계
    total_items = len(df)
    total_declarations = df['신고번호'].nunique()
    total_authors = df['작성자'].nunique() if '작성자' in df.columns else 0
    total_importers = df['납세자상호'].nunique() if '납세자상호' in df.columns else 0
    total_forwarders = df['운송주선인상호'].nunique() if '운송주선인상호' in df.columns else 0
    
    st.header("📊 전체 현황 요약")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("총 처리 품목", f"{total_items:,}개")
    
    with col2:
        st.metric("총 신고번호", f"{total_declarations:,}개")
    
    with col3:
        st.metric("담당 작성자", f"{total_authors}명")
    
    with col4:
        st.metric("협력 수입자", f"{total_importers}개사")
    
    with col5:
        st.metric("협력 포워더", f"{total_forwarders}개사")
    
    # 메인 탭 구성
    tab1, tab2, tab3, tab4 = st.tabs([
        "🏢 작성자별 분석 (내부관리)",
        "🏭 수입자별 분석 (고객관리)", 
        "🚛 운송주선인별 분석 (포워딩관리)",
        "📋 기타 분석 (검사/통계)"
    ])
    
    # ========== TAB 1: 작성자별 분석 ==========
    with tab1:
        st.markdown('<div class="tab-header internal-tab">🏢 작성자별 분석 - 내부 직원 관리</div>', 
                   unsafe_allow_html=True)
        
        if author_df.empty:
            st.warning("작성자 데이터가 없습니다.")
        else:
            # 복잡도 분석
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 복잡도 랭킹 차트
                fig_author = create_top_entities_chart(author_df, '작성자', '복잡도점수', 10)
                st.plotly_chart(fig_author, use_container_width=True)
            
            with col2:
                # 복잡도 분포
                fig_dist = create_complexity_distribution(author_df, '작성자')
                st.plotly_chart(fig_dist, use_container_width=True)
            
            # 요일별 패턴 분석
            st.subheader("📅 작성자별 요일 처리 패턴")
            
            selected_authors = st.multiselect(
                "분석할 작성자 선택 (최대 10명)",
                options=author_df['작성자'].tolist(),
                default=author_df['작성자'].head(5).tolist(),
                help="요일별 처리 패턴을 분석할 작성자를 선택하세요. 모든 작성자가 기본으로 포함됩니다."
            )
            
            if selected_authors:
                fig_weekday = create_weekday_chart(
                    author_df[author_df['작성자'].isin(selected_authors)],
                    "작성자별 요일 처리 현황",
                    '작성자'
                )
                st.plotly_chart(fig_weekday, use_container_width=True)
            
            # 상세 데이터 테이블
            st.subheader("📋 작성자별 상세 현황")
            
            display_columns = [
                '작성자', '총처리건수', '고유신고번호수', '복잡도점수', 
                'FTA활용률', '관세감면적용률', '담당수입자수'
            ]
            
            st.dataframe(
                author_df[display_columns].style.format({
                    '총처리건수': '{:,}',
                    '고유신고번호수': '{:,}',
                    '복잡도점수': '{:.1f}',
                    'FTA활용률': '{:.1f}%',
                    '관세감면적용률': '{:.1f}%',
                    '담당수입자수': '{:,}'
                }),
                use_container_width=True
            )
            
            # 작성자별 상세 분석 섹션 추가
            st.subheader("🔍 작성자별 상세 분석")
            
            # 작성자 선택
            selected_author = st.selectbox(
                "상세 분석할 작성자 선택",
                options=author_df['작성자'].tolist(),
                help="담당 수입자와 신고번호를 확인할 작성자를 선택하세요"
            )
            
            if selected_author:
                # 선택된 작성자의 데이터 필터링
                author_data = df[df['작성자'] == selected_author]
                decl_grouped = author_data.groupby('신고번호').first().reset_index()
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("총 처리 건수", f"{len(author_data):,}개")
                    st.metric("고유 신고번호", f"{len(decl_grouped):,}개")
                
                with col2:
                    complexity_score = analyzer.calculate_complexity_score(author_data)
                    st.metric("복잡도 점수", f"{complexity_score:.1f}")
                    st.metric("담당 수입자", f"{decl_grouped['납세자상호'].nunique()}개사")
                
                with col3:
                    fta_rate = len(decl_grouped[decl_grouped['원산지증명유무'] == 'Y']) / len(decl_grouped) * 100
                    exemption_rate = len(decl_grouped[
                        decl_grouped['관세감면구분'].notna() & 
                        (decl_grouped['관세감면구분'].astype(str).str.strip() != '')
                    ]) / len(decl_grouped) * 100
                    st.metric("FTA 활용률", f"{fta_rate:.1f}%")
                    st.metric("감면 적용률", f"{exemption_rate:.1f}%")
                
                # 담당 수입자 목록
                st.subheader(f"🏭 {selected_author} 담당 수입자 목록")
                
                importer_summary = decl_grouped.groupby('납세자상호').agg({
                    '총란수': 'sum',
                    '총규격수': 'sum',
                    '원산지증명유무': lambda x: (x == 'Y').sum(),
                    '관세감면구분': lambda x: x.notna().sum()
                }).reset_index()
                
                # 신고번호 수 계산 (그룹별 개수)
                importer_summary['신고번호수'] = decl_grouped.groupby('납세자상호').size().values
                
                importer_summary.columns = ['수입자', '총란수', '총규격수', 'FTA건수', '감면건수', '신고번호수']
                importer_summary = importer_summary.sort_values('신고번호수', ascending=False)
                
                # 수입자별 통계 테이블
                st.dataframe(
                    importer_summary.style.format({
                        '신고번호수': '{:,}',
                        '총란수': '{:,}',
                        '총규격수': '{:,}',
                        'FTA건수': '{:,}',
                        '감면건수': '{:,}'
                    }),
                    use_container_width=True
                )
                
                # 고유 신고번호 목록
                st.subheader(f"📋 {selected_author} 처리 신고번호 목록")
                
                # 신고번호별 상세 정보
                decl_details = decl_grouped[['신고번호', '납세자상호', '운송주선인상호', '총란수', '총규격수', 
                                           '원산지증명유무', '관세감면구분', 'C/S검사구분']].copy()
                decl_details = decl_details.sort_values('신고번호')
                
                # 검색 기능 추가
                search_term = st.text_input(
                    "신고번호 또는 수입자 검색",
                    placeholder="신고번호나 수입자명을 입력하세요"
                )
                
                if search_term:
                    mask = (decl_details['신고번호'].astype(str).str.contains(search_term, na=False) |
                           decl_details['납세자상호'].astype(str).str.contains(search_term, na=False))
                    decl_details = decl_details[mask]
                
                # 신고번호 목록 테이블
                st.dataframe(
                    decl_details.style.format({
                        '총란수': '{:,}',
                        '총규격수': '{:,}'
                    }),
                    use_container_width=True
                )
                
                # 신고번호별 요약 통계
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📊 신고번호별 통계**")
                    st.write(f"• 총 신고번호: {len(decl_details):,}개")
                    st.write(f"• 평균 란수: {decl_details['총란수'].mean():.1f}란")
                    st.write(f"• 평균 규격수: {decl_details['총규격수'].mean():.1f}규격")
                    st.write(f"• FTA 활용: {len(decl_details[decl_details['원산지증명유무'] == 'Y']):,}건")
                    st.write(f"• 감면 적용: {len(decl_details[decl_details['관세감면구분'].notna()]):,}건")
                
                with col2:
                    st.markdown("**🔍 검사구분별 현황**")
                    cs_counts = decl_details['C/S검사구분'].value_counts()
                    for cs_type, count in cs_counts.items():
                        st.write(f"• {cs_type}: {count:,}건")
                
                # 다운로드 기능
                st.subheader("💾 상세 데이터 다운로드")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # 수입자별 요약 다운로드
                    if not importer_summary.empty:
                        excel_data = create_excel_download(importer_summary, f"{selected_author}_담당수입자_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                        st.download_button(
                            f"📥 {selected_author} 담당수입자 (Excel)",
                            excel_data,
                            f"{selected_author}_담당수입자_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                with col2:
                    # 신고번호별 상세 다운로드
                    if not decl_details.empty:
                        excel_data = create_excel_download(decl_details, f"{selected_author}_신고번호목록_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                        st.download_button(
                            f"📥 {selected_author} 신고번호목록 (Excel)",
                            excel_data,
                            f"{selected_author}_신고번호목록_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            
            # 인사이트
            st.subheader("💡 내부 관리 인사이트")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 복잡도별 분류
                high_complexity = author_df[author_df['복잡도점수'] >= 200]
                medium_complexity = author_df[(author_df['복잡도점수'] >= 100) & (author_df['복잡도점수'] < 200)]
                low_complexity = author_df[author_df['복잡도점수'] < 100]
                
                st.markdown("**🎯 업무 분배 추천**")
                st.markdown(f"🔴 초고복잡도 담당: {len(high_complexity)}명")
                for author in high_complexity['작성자'].head(3):
                    st.write(f"  • {author}")
                
                st.markdown(f"🟡 중복잡도 담당: {len(medium_complexity)}명")
                for author in medium_complexity['작성자'].head(3):
                    st.write(f"  • {author}")
                
                st.markdown(f"🟢 일반업무 담당: {len(low_complexity)}명")
                for author in low_complexity['작성자'].head(3):
                    st.write(f"  • {author}")
            
            with col2:
                # 처리량 분석
                top_performer = author_df.iloc[0]
                most_efficient = author_df.loc[author_df['평균품목수_신고서'].idxmin()]
                
                st.markdown("**🏆 주요 성과 분석**")
                st.markdown(f"🥇 최고 복잡도: {top_performer['작성자']} ({top_performer['복잡도점수']:.1f}점)")
                st.markdown(f"⚡ 최고 효율성: {most_efficient['작성자']} ({most_efficient['평균품목수_신고서']:.1f}개/신고서)")
                
                # 담당 고객사 수 분석
                if '담당수입자수' in author_df.columns:
                    most_clients = author_df.loc[author_df['담당수입자수'].idxmax()]
                    st.markdown(f"🤝 최다 고객: {most_clients['작성자']} ({most_clients['담당수입자수']}개사)")
                
                st.markdown("**📊 복잡도 구성 요소**")
                st.markdown("• 총란수 + 총규격수")
                st.markdown("• 수입요건수 × 10점")
                st.markdown("• 감면/FTA/거래구분/무역거래처")
    
    # ========== TAB 2: 수입자별 분석 ==========
    with tab2:
        st.markdown('<div class="tab-header client-tab">🏭 수입자별 분석 - 고객사 관리</div>', 
                   unsafe_allow_html=True)
        
        if importer_df.empty:
            st.warning("수입자 데이터가 없습니다.")
        else:
            # 고객사 현황 분석
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 처리량 랭킹
                fig_importer = create_top_entities_chart(importer_df, '수입자', '총처리건수', 10)
                st.plotly_chart(fig_importer, use_container_width=True)
            
            with col2:
                # 복잡도 분포
                fig_complexity = create_complexity_distribution(importer_df, '수입자')
                st.plotly_chart(fig_complexity, use_container_width=True)
            
            # 고객별 특성 분석
            st.subheader("🔍 고객사별 업종 특성 분석")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 수입요건 vs 복잡도 스캐터
                fig_scatter = px.scatter(
                    importer_df.head(20),
                    x='수입요건비율',
                    y='복잡도점수',
                    size='총처리건수',
                    hover_name='수입자',
                    title="수입요건 비율 vs 복잡도 (상위 20개사)",
                    labels={'수입요건비율': '수입요건 비율 (%)', '복잡도점수': '복잡도 점수'}
                )
                st.plotly_chart(fig_scatter, use_container_width=True)
            
            with col2:
                # FTA 활용률 분포
                fig_fta = px.histogram(
                    importer_df,
                    x='FTA활용률',
                    title="고객사별 FTA 활용률 분포",
                    nbins=20
                )
                st.plotly_chart(fig_fta, use_container_width=True)
            
            # 담당자 분석
            st.subheader("👥 고객사별 담당 작성자 현황")
            
            # 주요 고객사 선택
            top_importers = st.selectbox(
                "상세 분석할 고객사 선택",
                options=importer_df['수입자'].head(20).tolist()
            )
            
            if top_importers:
                selected_importer = importer_df[importer_df['수입자'] == top_importers].iloc[0]
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("처리 건수", f"{selected_importer['총처리건수']:,}개")
                    st.metric("신고번호", f"{selected_importer['고유신고번호수']:,}개")
                
                with col2:
                    st.metric("복잡도 점수", f"{selected_importer['복잡도점수']:.1f}")
                    st.metric("주담당자", selected_importer['주담당작성자'])
                
                with col3:
                    st.metric("FTA 활용률", f"{selected_importer['FTA활용률']:.1f}%")
                    st.metric("발급서류 종류", f"{selected_importer['발급서류종류수']}가지")
            
            # 요일별 패턴
            st.subheader("📅 주요 고객사 요일별 신고 패턴")
            
            selected_importers_pattern = st.multiselect(
                "패턴 분석할 고객사 선택",
                options=importer_df['수입자'].head(10).tolist(),
                default=importer_df['수입자'].head(3).tolist()
            )
            
            if selected_importers_pattern:
                fig_weekday_importer = create_weekday_chart(
                    importer_df[importer_df['수입자'].isin(selected_importers_pattern)],
                    "고객사별 요일 신고 패턴",
                    '수입자'
                )
                st.plotly_chart(fig_weekday_importer, use_container_width=True)
            
            # 인사이트
            st.subheader("💡 고객 관리 인사이트")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # VIP 고객 분석
                vip_customers = importer_df[importer_df['총처리건수'] >= importer_df['총처리건수'].quantile(0.8)]
                complex_customers = importer_df[importer_df['복잡도점수'] >= 150]
                
                st.markdown("**🏆 VIP 고객사 (상위 20%)**")
                for customer in vip_customers['수입자'].head(5):
                    st.write(f"  • {customer}")
                
                st.markdown("**⚠️ 고복잡도 고객사 (150점+)**")
                for customer in complex_customers['수입자'].head(5):
                    st.write(f"  • {customer}")
            
            with col2:
                # 서비스 개선 제안
                high_requirement = importer_df[importer_df['수입요건비율'] >= 70]
                low_fta = importer_df[importer_df['FTA활용률'] < 30]
                
                st.markdown("**🎯 맞춤 서비스 제안**")
                if not high_requirement.empty:
                    st.markdown("📋 수입요건 컨설팅 필요:")
                    for customer in high_requirement['수입자'].head(3):
                        st.write(f"  • {customer}")
                
                if not low_fta.empty:
                    st.markdown("🌍 FTA 컨설팅 필요:")
                    for customer in low_fta['수입자'].head(3):
                        st.write(f"  • {customer}")
    
    # ========== TAB 3: 운송주선인별 분석 ==========
    with tab3:
        st.markdown('<div class="tab-header forwarder-tab">🚛 운송주선인별 분석 - 포워딩 파트너 관리</div>', 
                   unsafe_allow_html=True)
        
        if forwarder_df.empty:
            st.warning("운송주선인 데이터가 없습니다.")
        else:
            # 포워더 현황 분석
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 처리량 랭킹
                fig_forwarder = create_top_entities_chart(forwarder_df, '운송주선인', '총처리건수', 10)
                st.plotly_chart(fig_forwarder, use_container_width=True)
            
            with col2:
                # 효율성 분포
                fig_efficiency = px.histogram(
                    forwarder_df,
                    x='평균품목수_신고서',
                    title="포워더별 신고서 효율성 분포",
                    nbins=15
                )
                st.plotly_chart(fig_efficiency, use_container_width=True)
            
            # 포워더 효율성 분석
            st.subheader("⚡ 포워더별 처리 효율성 분석")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 평균 란수 vs 복잡도
                fig_lanes = px.scatter(
                    forwarder_df.head(20),
                    x='평균란수_신고서',
                    y='복잡도점수',
                    size='총처리건수',
                    hover_name='운송주선인',
                    title="평균 란수 vs 복잡도 (상위 20개사)"
                )
                st.plotly_chart(fig_lanes, use_container_width=True)
            
            with col2:
                # 연결 네트워크 분석
                fig_network = px.scatter(
                    forwarder_df.head(20),
                    x='연결수입자수',
                    y='연결무역거래처수',
                    size='총처리건수',
                    hover_name='운송주선인',
                    title="연결 네트워크 규모 (상위 20개사)"
                )
                st.plotly_chart(fig_network, use_container_width=True)
            
            # 파트너십 분석
            st.subheader("🤝 포워더별 담당 현황 분석")
            
            # 주요 포워더 선택
            selected_forwarder = st.selectbox(
                "상세 분석할 포워더 선택",
                options=forwarder_df['운송주선인'].head(15).tolist()
            )
            
            if selected_forwarder:
                forwarder_info = forwarder_df[forwarder_df['운송주선인'] == selected_forwarder].iloc[0]
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("총 처리량", f"{forwarder_info['총처리건수']:,}개")
                    st.metric("신고번호", f"{forwarder_info['고유신고번호수']:,}개")
                
                with col2:
                    st.metric("복잡도", f"{forwarder_info['복잡도점수']:.1f}")
                    st.metric("주담당자", forwarder_info['주담당작성자'])
                
                with col3:
                    st.metric("주요 수입자", forwarder_info['주요수입자'])
                    st.metric("연결 수입자", f"{forwarder_info['연결수입자수']}개사")
                
                with col4:
                    st.metric("FTA 활용률", f"{forwarder_info['FTA활용률']:.1f}%")
                    st.metric("평균 효율성", f"{forwarder_info['평균품목수_신고서']:.1f}")
                
                # 선택된 포워더의 담당 작성자 및 수입자 상세 분석
                st.subheader(f"📊 {selected_forwarder} 상세 담당 현황")
                
                forwarder_data = df[df['운송주선인상호'] == selected_forwarder]
                decl_grouped = forwarder_data.groupby('신고번호').first().reset_index()
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # 담당 작성자 분포
                    author_dist = decl_grouped['작성자'].value_counts().head(10)
                    fig_author_dist = px.pie(
                        values=author_dist.values,
                        names=author_dist.index,
                        title=f"{selected_forwarder} 담당 작성자 분포",
                        height=400
                    )
                    st.plotly_chart(fig_author_dist, use_container_width=True)
                
                with col2:
                    # 담당 수입자 분포
                    importer_dist = decl_grouped['납세자상호'].value_counts().head(10)
                    fig_importer_dist = px.pie(
                        values=importer_dist.values,
                        names=importer_dist.index,
                        title=f"{selected_forwarder} 담당 수입자 분포",
                        height=400
                    )
                    st.plotly_chart(fig_importer_dist, use_container_width=True)
            
            # 요일별 패턴
            st.subheader("📅 주요 포워더 요일별 처리 패턴")
            
            selected_forwarders = st.multiselect(
                "패턴 분석할 포워더 선택",
                options=forwarder_df['운송주선인'].head(10).tolist(),
                default=forwarder_df['운송주선인'].head(3).tolist()
            )
            
            if selected_forwarders:
                fig_weekday_forwarder = create_weekday_chart(
                    forwarder_df[forwarder_df['운송주선인'].isin(selected_forwarders)],
                    "포워더별 요일 처리 패턴",
                    '운송주선인'
                )
                st.plotly_chart(fig_weekday_forwarder, use_container_width=True)
            
            # 인사이트
            st.subheader("💡 포워딩 파트너 관리 인사이트")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # 주요 파트너 분석
                major_partners = forwarder_df[forwarder_df['총처리건수'] >= forwarder_df['총처리건수'].quantile(0.8)]
                efficient_partners = forwarder_df[forwarder_df['평균품목수_신고서'] <= 10]
                
                st.markdown("**🏆 주요 파트너 (상위 20%)**")
                for partner in major_partners['운송주선인'].head(5):
                    st.write(f"  • {partner}")
                
                st.markdown("**⚡ 고효율 파트너 (10개 미만/신고서)**")
                for partner in efficient_partners['운송주선인'].head(5):
                    st.write(f"  • {partner}")
            
            with col2:
                # 협력 강화 제안
                high_complexity = forwarder_df[forwarder_df['복잡도점수'] >= 100]
                diverse_network = forwarder_df[forwarder_df['연결수입자수'] >= 10]
                
                st.markdown("**🎯 협력 강화 제안**")
                if not high_complexity.empty:
                    st.markdown("🔧 복잡업무 전문 파트너:")
                    for partner in high_complexity['운송주선인'].head(3):
                        st.write(f"  • {partner}")
                
                if not diverse_network.empty:
                    st.markdown("🌐 네트워크 확장 파트너:")
                    for partner in diverse_network['운송주선인'].head(3):
                        st.write(f"  • {partner}")
    
    # ========== TAB 4: 기타 분석 ==========
    with tab4:
        st.markdown('<div class="tab-header">📋 기타 분석 - 검사구분 및 통계</div>', 
                   unsafe_allow_html=True)
        
        if cs_df.empty:
            st.warning("C/S검사구분 데이터가 없습니다.")
        else:
            # 전체 검사 현황 요약
            st.header("🔍 C/S 검사구분 전체 현황")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("총 신고번호", f"{cs_stats['총신고번호수']:,}개")
            
            with col2:
                st.metric("검사구분 종류", f"{cs_stats['검사구분종류']}종류")
            
            with col3:
                st.metric("주요 검사유형", cs_stats['가장많은검사'])
            
            with col4:
                st.metric("무검사율", f"{cs_stats['무검사율']}%")
            
            # 검사구분별 상세 분석
            st.subheader("📊 검사구분별 상세 분석")
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 검사구분별 신고번호 수 바차트
                fig_cs = px.bar(
                    cs_df,
                    x='검사유형',
                    y='신고번호수',
                    color='신고번호수',
                    title="검사구분별 신고번호 수",
                    text='신고번호수',
                    color_continuous_scale='viridis'
                )
                fig_cs.update_traces(texttemplate='%{text:,}', textposition='outside')
                fig_cs.update_layout(height=500)
                st.plotly_chart(fig_cs, use_container_width=True)
            
            with col2:
                # 검사구분 비율 도넛차트
                fig_cs_pie = px.pie(
                    cs_df,
                    values='신고번호수',
                    names='검사유형',
                    title="검사구분 비율",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_cs_pie.update_layout(height=500)
                st.plotly_chart(fig_cs_pie, use_container_width=True)
            
            # 상세 데이터 테이블
            st.subheader("📋 검사구분별 상세 데이터")
            
            # 비율 계산 추가
            cs_df_display = cs_df.copy()
            cs_df_display['비율(%)'] = round(cs_df_display['신고번호수'] / cs_stats['총신고번호수'] * 100, 1)
            cs_df_display['누적비율(%)'] = cs_df_display['비율(%)'].cumsum()
            
            st.dataframe(
                cs_df_display[['검사구분', '검사유형', '신고번호수', '비율(%)', '누적비율(%)']].style.format({
                    '신고번호수': '{:,}',
                    '비율(%)': '{:.1f}%',
                    '누적비율(%)': '{:.1f}%'
                }),
                use_container_width=True
            )
            
            # 검사구분별 인사이트
            st.subheader("💡 검사 패턴 인사이트")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**🎯 검사 효율성 분석**")
                
                # 무검사 vs 검사 비율
                no_inspection = cs_df[cs_df['검사구분'] == 'N']['신고번호수'].sum() if 'N' in cs_df['검사구분'].values else 0
                total_inspection = cs_stats['총신고번호수'] - no_inspection
                
                st.markdown(f"• 무검사: {no_inspection:,}건 ({cs_stats['무검사율']}%)")
                st.markdown(f"• 검사대상: {total_inspection:,}건 ({100-cs_stats['무검사율']:.1f}%)")
                
                # 가장 많은 검사유형
                if not cs_df.empty:
                    top_inspection = cs_df.iloc[0]
                    st.markdown(f"• 주요 검사: {top_inspection['검사유형']} ({top_inspection['신고번호수']:,}건)")
            
            with col2:
                st.markdown("**📈 검사 분포 특성**")
                
                # 검사구분 다양성
                st.markdown(f"• 검사구분 종류: {cs_stats['검사구분종류']}가지")
                
                # 상위 3개 검사유형
                top3 = cs_df.head(3)
                top3_total = top3['신고번호수'].sum()
                top3_ratio = round(top3_total / cs_stats['총신고번호수'] * 100, 1)
                
                st.markdown(f"• 상위 3개 검사유형이 전체의 {top3_ratio}% 차지")
                
                for i, (_, row) in enumerate(top3.iterrows(), 1):
                    st.markdown(f"  {i}. {row['검사유형']}: {row['신고번호수']:,}건")
            
            # 검사구분 매핑 정보
            with st.expander("📖 검사구분 코드 매핑 정보"):
                st.markdown("""
                **검사구분 코드 설명:**
                - **Y**: 세관검사 (물리검사)
                - **F**: 협업검사 (관련기관 합동검사)
                - **N**: 무검사 (서류심사만)
                - **C**: 서류검사 (서류 정밀심사)
                - **S**: 표본검사 (샘플 검사)
                
                **활용 방안:**
                - 무검사율이 높은 경우: 신뢰도 높은 수입자/품목
                - 검사율이 높은 경우: 위험도 관리 필요
                - 협업검사 비율: 다부처 협의 업무량 예측
                """)
        
        # 추가 통계 분석 (향후 확장 가능)
        st.subheader("📈 추가 통계 분석")
        
        with st.expander("🔧 향후 확장 가능한 분석 항목"):
            st.markdown("""
            **1. 시간별 분석**
            - 월별/주별 검사패턴 변화
            - 요일별 검사구분 분포
            - 계절성 검사 트렌드
            
            **2. 상관관계 분석**  
            - 복잡도 vs 검사구분
            - 수입자별 검사패턴
            - 작성자별 검사결과
            
            **3. 예측 모델**
            - 검사구분 예측 모델
            - 통관 소요시간 예측
            - 위험도 스코어링
            """)
    
    # 종합 다운로드 (4개 파일)
    st.header("💾 종합 분석 결과 다운로드")
    
    # 다운로드 형식 선택
    download_format = st.selectbox(
        "다운로드 형식 선택",
        ["엑셀 (Excel)", "PDF", "HTML"],
        help="원하는 형식으로 분석 결과를 다운로드하세요"
    )
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if not author_df.empty:
            if download_format == "엑셀 (Excel)":
                excel_data = create_excel_download(author_df, f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                st.download_button(
                    "👥 작성자 분석 결과 (Excel)",
                    excel_data,
                    f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif download_format == "PDF":
                pdf_data = create_pdf_download(author_df, "작성자별 분석 결과", f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
                st.download_button(
                    "👥 작성자 분석 결과 (PDF)",
                    pdf_data,
                    f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    "application/pdf"
                )
            else:  # HTML
                html_data = create_html_download(author_df, "작성자별 분석 결과", f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html")
                st.download_button(
                    "👥 작성자 분석 결과 (HTML)",
                    html_data,
                    f"작성자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    "text/html"
                )
    
    with col2:
        if not importer_df.empty:
            if download_format == "엑셀 (Excel)":
                excel_data = create_excel_download(importer_df, f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                st.download_button(
                    "🏭 수입자 분석 결과 (Excel)",
                    excel_data,
                    f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif download_format == "PDF":
                pdf_data = create_pdf_download(importer_df, "수입자별 분석 결과", f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
                st.download_button(
                    "🏭 수입자 분석 결과 (PDF)",
                    pdf_data,
                    f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    "application/pdf"
                )
            else:  # HTML
                html_data = create_html_download(importer_df, "수입자별 분석 결과", f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html")
                st.download_button(
                    "🏭 수입자 분석 결과 (HTML)",
                    html_data,
                    f"수입자분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    "text/html"
                )
    
    with col3:
        if not forwarder_df.empty:
            if download_format == "엑셀 (Excel)":
                excel_data = create_excel_download(forwarder_df, f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                st.download_button(
                    "🚛 포워더 분석 결과 (Excel)",
                    excel_data,
                    f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif download_format == "PDF":
                pdf_data = create_pdf_download(forwarder_df, "운송주선인별 분석 결과", f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
                st.download_button(
                    "🚛 포워더 분석 결과 (PDF)",
                    pdf_data,
                    f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    "application/pdf"
                )
            else:  # HTML
                html_data = create_html_download(forwarder_df, "운송주선인별 분석 결과", f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html")
                st.download_button(
                    "🚛 포워더 분석 결과 (HTML)",
                    html_data,
                    f"포워더분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    "text/html"
                )
    
    with col4:
        if not cs_df.empty:
            if download_format == "엑셀 (Excel)":
                excel_data = create_excel_download(cs_df, f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
                st.download_button(
                    "📋 검사구분 분석 결과 (Excel)",
                    excel_data,
                    f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif download_format == "PDF":
                pdf_data = create_pdf_download(cs_df, "검사구분별 분석 결과", f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
                st.download_button(
                    "📋 검사구분 분석 결과 (PDF)",
                    pdf_data,
                    f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    "application/pdf"
                )
            else:  # HTML
                html_data = create_html_download(cs_df, "검사구분별 분석 결과", f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html")
                st.download_button(
                    "📋 검사구분 분석 결과 (HTML)",
                    html_data,
                    f"검사구분분석_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    "text/html"
                )
    
    # 종합 리포트
    st.subheader("📄 종합 리포트 다운로드")
    
    report = f"""
관세법인 우신 종합 분석 리포트
{'='*60}

📊 전체 현황 요약
- 총 처리 품목: {total_items:,}개
- 총 신고번호: {total_declarations:,}개  
- 담당 작성자: {total_authors}명
- 협력 수입자: {total_importers}개사
- 협력 포워더: {total_forwarders}개사

🏢 작성자별 분석 (상위 3명)
"""
    if not author_df.empty:
        for i, (_, row) in enumerate(author_df.head(3).iterrows(), 1):
            report += f"{i}. {row['작성자']}: 복잡도 {row['복잡도점수']:.1f}점, 처리량 {row['총처리건수']:,}개\n"

    report += f"""
🏭 수입자별 분석 (상위 3개사)
"""
    if not importer_df.empty:
        for i, (_, row) in enumerate(importer_df.head(3).iterrows(), 1):
            report += f"{i}. {row['수입자']}: 복잡도 {row['복잡도점수']:.1f}점, 처리량 {row['총처리건수']:,}개\n"

    report += f"""
🚛 포워더별 분석 (상위 3개사)  
"""
    if not forwarder_df.empty:
        for i, (_, row) in enumerate(forwarder_df.head(3).iterrows(), 1):
            report += f"{i}. {row['운송주선인']}: 복잡도 {row['복잡도점수']:.1f}점, 처리량 {row['총처리건수']:,}개\n"

    if not cs_df.empty:
        report += f"""
📋 검사구분 분석
- 총 신고번호: {cs_stats['총신고번호수']:,}개
- 무검사율: {cs_stats['무검사율']}%
- 주요 검사유형: {cs_stats['가장많은검사']}
"""

    report += f"""

💡 주요 인사이트
- 7차원 복잡도 분석으로 업무 난이도 정량화 완료
- 내부/고객/파트너 3차원 통합 관리 체계 구축
- 검사패턴 분석으로 위험도 관리 최적화 가능

생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
시스템: 관세법인 우신 종합 분석 시스템 v2.0
"""
    
    st.download_button(
        "📄 종합리포트 TXT 다운로드",
        report,
        f"우신_종합리포트_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        "text/plain"
    )

if __name__ == "__main__":
    main()