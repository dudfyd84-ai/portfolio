"""
안티그래비(Anti-Gravity) 제조 및 유통망 관리 대시보드 프로토타입
본 애플리케이션은 실시간 데이터 스트리밍, 다차원 스타 스키마 결합, BOM 소요량 계산 필드 수식, 
그리고 이중 축 혼합 차트를 통한 예측 시각화를 제공하는 안티그래비 솔루션의 Streamlit 기반 프로토타입입니다.

작성자: 안티그래비 마스터 개발자 및 데이터 아키텍트
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import os
import time
import urllib.parse

# 1. Streamlit 페이지 설정 및 Premium UX 톤앤매너 적용
st.set_page_config(
    page_title="PREMIUM BRAND SCM | 글로벌 럭셔리 제조 및 유통 실시간 의사결정 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS를 통한 UI/UX 강화 (다니엘트루스 브랜드 테마: 딥 네이비, 소프트 골드, 아이보리)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;600;700;800&family=Outfit:wght@300;400;600;800&family=Playfair+Display:ital,wght@0,400;0,700;1,400&display=swap');
    
    /* 전역 배경색 (아이보리 톤앤매너) 및 기본 폰트 설정 */
    .stApp {
        background-color: #F8F7F4;
    }
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    h1, h2, h3, h4, h5, h6, [data-testid="stHeader"] {
        font-family: 'Cinzel', 'Playfair Display', serif;
        letter-spacing: 0.05em;
        color: #0B132B !important;
    }
    
    /* 사이드바 스타일링 (딥 네이비 & 소프트 골드 테두리) */
    [data-testid="stSidebar"] {
        background-color: #0B132B !important;
        color: #F8F7F4 !important;
        border-right: 2px solid #C5A880;
    }
    
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4, [data-testid="stSidebar"] h5, [data-testid="stSidebar"] h6 {
        color: #C5A880 !important;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #F8F7F4 !important;
    }
    
    /* st.selectbox 등 사이드바 위젯 텍스트 조율 */
    [data-testid="stSidebar"] label {
        color: #C5A880 !important;
        font-weight: 600;
        font-family: 'Outfit', sans-serif;
    }
    
    /* KPI 카드 스타일 */
    .kpi-card {
        background: #ffffff;
        border: 1.5px solid #C5A880; /* 소프트 골드 */
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 12px rgba(11, 19, 43, 0.03);
        transition: transform 0.25s ease, box-shadow 0.25s ease, border-color 0.25s ease;
    }
    .kpi-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 20px rgba(11, 19, 43, 0.08);
        border-color: #0B132B; /* 딥 네이비로 호버 테두리 변경 */
    }
    .kpi-title {
        color: #94A3B8;
        font-family: 'Outfit', sans-serif;
        font-size: 14px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    .kpi-value {
        color: #0B132B;
        font-family: 'Cinzel', serif;
        font-size: 32px;
        font-weight: 800;
        margin-top: 8px;
    }
    .kpi-delta-up {
        color: #10b981;
        font-size: 14px;
        font-weight: 600;
        margin-top: 4px;
    }
    .kpi-delta-down {
        color: #ef4444;
        font-size: 14px;
        font-weight: 600;
        margin-top: 4px;
    }
    
    /* 웹소켓 연결 상태 인디케이터 (딥 네이비 & 소프트 골드 뱃지) */
    .ws-status-connected {
        background-color: #0B132B;
        color: #C5A880;
        border: 1px solid #C5A880;
        padding: 6px 12px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    
    .ws-dot {
        height: 8px;
        width: 8px;
        background-color: #C5A880;
        border-radius: 50%;
        display: inline-block;
        animation: pulse 1.5s infinite;
    }
    
    @keyframes pulse {
        0% { transform: scale(0.9); opacity: 1; }
        50% { transform: scale(1.3); opacity: 0.5; }
        100% { transform: scale(0.9); opacity: 1; }
    }
</style>
""", unsafe_allow_html=True)

# 2. 데이터 로딩 레이어 및 캐싱 (@st.cache_data 활용)
# 구글 스프레드시트 실시간 연동 설정 (xlsx 포맷 다운로드를 통한 탭 선택 방식)
SPREADSHEET_1_ID = "1yOcgs_Zj_c_WFPElG6Er1MMNl8b3gk-_B-SBvJ0i7rg"  # 1차 시트 (2026년 판매)
SPREADSHEET_2_ID = "1yxWKwtpI-0TlUq3otoh7TiNd1hI2DhLh981x1juh_ds"  # 2차 시트 (2025년 판매)

GOOGLE_SHEET_URL_1 = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_1_ID}/export?format=xlsx"
GOOGLE_SHEET_URL_2 = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_2_ID}/export?format=xlsx"

# ==========================================
# ⚡ [PORTFOLIO EDITION] 로컬 비식별화 데이터 로딩 엔진
# ==========================================
@st.cache_data
def load_anonymized_portfolio_data():
    """
    구글 시트 연동을 차단하고 로컬 비식별화 JSON 데이터셋을
    메모리 버퍼로 고속 로드하여 대시보드 백본 스타 스키마를 형성합니다.
    """
    import json
    json_path = 'scratch/portfolio_data.json'
    if not os.path.exists(json_path):
        json_path = 'portfolio_data.json'
        
    if not os.path.exists(json_path):
        # 만약 로컬에 파일이 없을 경우 비상용 가상 Mock 데이터를 생성
        st.error(f"⚠️ 포트폴리오 데이터셋 파일(portfolio_data.json)을 찾을 수 없습니다. 기본 목업을 생성합니다.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    # Pandas DataFrame 변환
    sales_df = pd.DataFrame(data["sales_data"])
    qty_df = pd.DataFrame(data["quantity_data"])
    inventory_df = pd.DataFrame(data["inventory_data"])
    
    # 데이터셋 유형 조율 (날짜/수치 정형화)
    sales_df['sale_date'] = pd.to_datetime(sales_df['date'])
    sales_df['amount'] = pd.to_numeric(sales_df['amount']).fillna(0.0)
    sales_df['store_name'] = sales_df['store_name']
    sales_df['channel_type'] = sales_df['channel_type']
    sales_df['year'] = sales_df['sale_date'].dt.year
    sales_df['month'] = sales_df['sale_date'].dt.month
    
    qty_df['sale_date'] = pd.to_datetime(qty_df['date'])
    qty_df['quantity'] = pd.to_numeric(qty_df['quantity']).fillna(0.0).astype(int)
    qty_df['year'] = qty_df['sale_date'].dt.year
    qty_df['month'] = qty_df['sale_date'].dt.month
    qty_df['product_name'] = qty_df['product_name']
    
    inventory_df['current_stock'] = pd.to_numeric(inventory_df['current_stock']).fillna(0.0)
    inventory_df['expected_sales'] = pd.to_numeric(inventory_df['expected_sales']).fillna(0.0)
    
    return sales_df, qty_df, inventory_df

# 기존 SPREADSHEET 로딩 함수 오버라이드용 스타 스키마 결합 구조 구축
@st.cache_data
def load_all_dashboard_data_DEPRECATED():
    """
    로컬 비식별화 데이터를 기반으로 메인 대시보드용 스타 스키마 뼈대를 구축합니다.
    """
    sales_raw, qty_raw, inventory_raw = load_anonymized_portfolio_data()
    
    # 1. 고유 매장 정보 맵
    unique_stores = sorted(list(sales_raw['store_name'].unique()))
    store_code_map = {name: name for name in unique_stores}
    sales_raw['store_code'] = sales_raw['store_name']
    sales_raw['sale_id'] = sales_raw.index.map(lambda idx: f"S_SAL_{idx+1:06d}")
    
    # 2. 완제품 품목 정보 맵
    unique_products = sorted(list(qty_raw['product_name'].unique()))
    prod_code_map = {name: f"PD_{i+1:03d}" for i, name in enumerate(unique_products)}
    
    qty_raw['product_code'] = qty_raw['product_name'].map(prod_code_map)
    qty_raw['sale_detail_id'] = qty_raw.index.map(lambda idx: f"S_QTY_{idx+1:06d}")
    
    # MRP 소요 계획용 가상 MRP 생성
    plan_df, bom_df = generate_fallback_mrp_data(prod_code_map)
    
    return sales_raw, qty_raw, plan_df, bom_df, prod_code_map, store_code_map

# 소진기한 전사 로딩 함수 오버라이드
@st.cache_data
def load_google_sojin_data_v6(quantity_df):
    """
    소진기한 및 재고 데이터를 로컬 비식별화 DB셋에서 연동 조인 처리합니다.
    """
    _, _, inventory_raw = load_anonymized_portfolio_data()
    
    # 품명 및 재고 뼈대
    df_stock = pd.DataFrame()
    df_stock['품명'] = inventory_raw['product_name']
    df_stock['현재재고'] = inventory_raw['current_stock']
    df_stock['품명_clean'] = df_stock['품명']
    
    # 예상 판매량 결합
    df_est = pd.DataFrame()
    df_est['품명_clean'] = inventory_raw['product_name']
    df_est['월평균 예상 판매량'] = inventory_raw['expected_sales']
    
    df_stock = df_stock.merge(df_est, on='품명_clean', how='left')
    
    # 과거 판매 시계열 생성
    df_sales = quantity_df.rename(columns={
        'product_name': '품명',
        'sale_date': '날짜',
        'quantity': '수량'
    }).copy()
    df_sales['날짜'] = df_sales['날짜'].dt.strftime('%Y.%m.%d')
    df_sales['품명_clean'] = df_sales['품명']
    
    return df_stock, df_sales


def generate_fallback_mrp_data(prod_code_map):
    """
    구글 시트 연동 실패 시 시스템 크래시를 방지하기 위해 
    기본 제품 코드 맵을 기반으로 한 고품질의 가상 MRP 데이터(plan_df, bom_df)를 생성합니다.
    """
    import pandas as pd
    
    # 1. 가상 제품 마스터 구축 및 카테고리 분류
    product_records = []
    def classify_category(p_name):
        if p_name.startswith("[증정]") or p_name.startswith("[샘플]"):
            return "증정"
        elif "디퓨저" in p_name:
            return "디퓨저"
        elif "오일" in p_name:
            return "오일 퍼퓸"
        elif "캔들" in p_name:
            return "캔들"
        elif "[액세서리]" in p_name or "스틱" in p_name or "캡" in p_name:
            return "액세서리"
        else:
            return "기타"
            
    for prod_name, prod_code in prod_code_map.items():
        cat = classify_category(prod_name)
        product_records.append({
            'product_code': prod_code,
            'product_name': prod_name,
            'category': cat
        })
    product_df = pd.DataFrame(product_records)
    
    # 2. 동적 BOM 정의 (카테고리별 자재 소요량 전개)
    bom_records = []
    bom_idx = 1
    for _, row in product_df.iterrows():
        p_code = row['product_code']
        p_name = row['product_name']
        p_cat = row['category']
        
        if p_cat == '디퓨저':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_BASE', 'child_name': '디퓨저 베이스 에탄올 (ml)', 'unit_qty': 150.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_FRAG', 'child_name': '디퓨저 조합 향료 오일 (ml)', 'unit_qty': 50.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_GLASS', 'child_name': '프리미엄 디퓨저 유리 용기 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '오일 퍼퓸':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_OL_FRAG', 'child_name': '천연 에센셜 향료 원액 (ml)', 'unit_qty': 8.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_OL_BOTTLE', 'child_name': '롤온 고급 초자 유리병 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '캔들':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_WAX', 'child_name': '천연 골든 소이 왁스 (g)', 'unit_qty': 180.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_FRAG', 'child_name': '캔들 가열용 향료 오일 (ml)', 'unit_qty': 20.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_GLASS', 'child_name': '내열성 캔들 유리 용기 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '액세서리':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_AC_BOX', 'child_name': '수작업 패키징 상자 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '증정':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_PR_MINI', 'child_name': '미니어처 전용 샘플 공병 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
            
    bom_df = pd.DataFrame(bom_records)
    
    # 3. 동적 생산 계획 정의 (2025년 가상 생산량)
    plan_records = []
    plan_idx = 1
    for _, row in product_df.iterrows():
        p_code = row['product_code']
        p_name = row['product_name']
        p_cat = row['category']
        
        if p_cat == '디퓨저':
            planned_qty = 1500
        elif p_cat == '오일 퍼퓸':
            planned_qty = 2500
        elif p_cat == '캔들':
            planned_qty = 1200
        elif p_cat == '액세서리':
            planned_qty = 800
        elif p_cat == '증정':
            planned_qty = 5000
        else:
            planned_qty = 500
            
        plan_records.append({
            'plan_id': f"PLN{plan_idx:04d}",
            'plan_month': '2025-05',
            'product_code': p_code,
            'product_name': p_name,
            'planned_qty': planned_qty
        })
        plan_idx += 1
        
    plan_df = pd.DataFrame(plan_records)
    
    return plan_df, bom_df, prod_code_map

@st.cache_data(ttl=300)
def load_google_mrp_data(prod_code_map):
    """
    구글 스프레드시트 1차 시트에서 '생산 계획' 및 'BOM 데이터' 탭을 가져와
    정밀 전처리 및 품명 매칭을 수행하고, 대시보드 스타 스키마에 호환되는
    plan_df와 bom_df 데이터프레임을 생성하여 반환합니다.
    """
    import re
    # 1. 생산 계획 탭 파싱
    try:
        encoded_name = urllib.parse.quote("생산 계획")
        url_plan = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_1_ID}/gviz/tq?tqx=out:csv&sheet={encoded_name}"
        df_plan_raw = pd.read_csv(url_plan)
        if df_plan_raw.empty or len(df_plan_raw.columns) == 0:
            raise ValueError("생산 계획 데이터가 비어 있습니다.")
        if df_plan_raw.columns[0] == 'Unnamed: 0' or pd.isna(df_plan_raw.columns[0]):
            df_plan_raw = df_plan_raw.iloc[:, 1:]
        df_plan_raw.columns = [str(c).strip() for c in df_plan_raw.columns]
        df_plan_raw = df_plan_raw.dropna(subset=['품명'])
        df_plan_raw['품명'] = df_plan_raw['품명'].astype(str).str.strip()
    except Exception as e:
        st.warning(f"⚠️ 구글 시트 '생산 계획' 탭 로드 실패 (가상 데이터 폴백 작동): {e}")
        return generate_fallback_mrp_data(prod_code_map)
        
    # 2. BOM 데이터 탭 파싱
    try:
        encoded_bom = urllib.parse.quote("BOM 데이터")
        url_bom = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_1_ID}/gviz/tq?tqx=out:csv&sheet={encoded_bom}"
        df_bom_raw = pd.read_csv(url_bom)
        if df_bom_raw.empty or len(df_bom_raw.columns) == 0:
            raise ValueError("BOM 데이터가 비어 있습니다.")
            
        # 1차 컬럼 이름 정규화 ('구분' 포함 컬럼을 '구분'으로 리네임)
        df_bom_raw.columns = [str(c).strip() for c in df_bom_raw.columns]
        for col in df_bom_raw.columns:
            if '구분' in col:
                df_bom_raw = df_bom_raw.rename(columns={col: '구분'})
                
        if df_bom_raw.columns[0] == 'Unnamed: 0' or pd.isna(df_bom_raw.columns[0]):
            df_bom_raw = df_bom_raw.iloc[:, 1:]
            
        # Unnamed 열 제거 후 2차 컬럼 이름 정규화
        df_bom_raw.columns = [str(c).strip() for c in df_bom_raw.columns]
        for col in df_bom_raw.columns:
            if '구분' in col:
                df_bom_raw = df_bom_raw.rename(columns={col: '구분'})
                
        if '구분' not in df_bom_raw.columns:
            raise KeyError("BOM 데이터 탭에 '구분' 관련 컬럼이 존재하지 않습니다.")
            
        df_bom_raw = df_bom_raw.dropna(subset=['구분'])
        df_bom_raw['구분'] = df_bom_raw['구분'].astype(str).str.strip()
    except Exception as e:
        st.warning(f"⚠️ 구글 시트 'BOM 데이터' 탭 로드 실패 (가상 데이터 폴백 작동): {e}")
        return generate_fallback_mrp_data(prod_code_map)

    # 3. 완제품 코드 맵(prod_code_map) 갱신 및 새로운 품목에 대한 코드 발급
    all_plan_products = df_plan_raw['품명'].unique()
    updated_prod_code_map = prod_code_map.copy()
    
    current_max_id = 0
    for code in updated_prod_code_map.values():
        if code.startswith("PROD_"):
            try:
                num = int(code.split("_")[1])
                if num > current_max_id:
                    current_max_id = num
            except ValueError:
                pass
                
    for prod_name in all_plan_products:
        if prod_name not in updated_prod_code_map:
            current_max_id += 1
            updated_prod_code_map[prod_name] = f"PROD_{current_max_id:03d}"

    # 4. plan_df 구축
    month_col = '5월'
    if month_col not in df_plan_raw.columns:
        num_cols = [c for c in df_plan_raw.columns if '월' in c]
        if num_cols:
            month_col = num_cols[0]
            
    plan_records = []
    for idx, row in df_plan_raw.iterrows():
        p_name = row['품명']
        p_code = updated_prod_code_map[p_name]
        
        val = row.get(month_col, 0)
        try:
            planned_qty = int(pd.to_numeric(val, errors='coerce'))
            if pd.isna(planned_qty) or planned_qty < 0:
                planned_qty = 0
        except:
            planned_qty = 0
            
        plan_records.append({
            'plan_id': f"PLN{idx+1:04d}",
            'plan_month': '2025-05',
            'product_code': p_code,
            'product_name': p_name,
            'planned_qty': planned_qty
        })
        
    plan_df = pd.DataFrame(plan_records)

    # 5. 정밀 품명 매칭을 활용한 bom_df 구축
    bom_records = []
    bom_idx = 1
    
    unique_materials = [str(name).strip() for name in df_bom_raw['품명'].dropna().unique()]
    unique_materials = list(dict.fromkeys(unique_materials))
    mat_code_map = {name: f"RAW_{i+1:03d}" for i, name in enumerate(unique_materials)}
    bom_parent_names = df_bom_raw['구분'].unique()
    
    for _, p_row in plan_df.iterrows():
        p_name = p_row['product_name']
        p_code = p_row['product_code']
        
        p_clean = p_name.replace(" ", "").lower()
        
        matched_bom_parent = None
        for b_name in bom_parent_names:
            b_clean = b_name.replace(" ", "").lower()
            
            if p_clean == b_clean or p_clean in b_clean or b_clean in p_clean:
                matched_bom_parent = b_name
                break
                
            p_pure = re.sub(r'\[.*?\]|\(.*?\)', '', p_clean)
            b_pure = re.sub(r'\[.*?\]|\(.*?\)', '', b_clean)
            if p_pure == b_pure and len(p_pure) > 2:
                matched_bom_parent = b_name
                break
                
        if matched_bom_parent:
            sub_bom = df_bom_raw[df_bom_raw['구분'] == matched_bom_parent]
            for _, b_row in sub_bom.iterrows():
                child_name = b_row['품명']
                if pd.isna(child_name):
                    continue
                child_name = str(child_name).strip()
                child_code = mat_code_map.get(child_name, f"RAW_ETC_{bom_idx}")
                
                u_qty_val = b_row.get('BOM', 0)
                try:
                    unit_qty = float(pd.to_numeric(u_qty_val, errors='coerce'))
                    if pd.isna(unit_qty) or unit_qty < 0:
                        unit_qty = 0.0
                except:
                    unit_qty = 0.0
                    
                bom_records.append({
                    'bom_id': f"BOM{bom_idx:04d}",
                    'parent_code': p_code,
                    'parent_name': p_name,
                    'child_code': child_code,
                    'child_name': child_name,
                    'unit_qty': unit_qty
                })
                bom_idx += 1
                
    bom_df = pd.DataFrame(bom_records)
    
    return plan_df, bom_df, updated_prod_code_map

@st.cache_data(ttl=300)
def load_google_sojin_data_v6_DEPRECATED(_quantity_df):
    """
    구글 스프레드시트의 '소진기한' 탭 전체를 스캔하여 
    Excel Q~R열(Index 16, 17)로부터 전사 완제품 실물 보유재고 데이터를 뼈대로 구축하고,
    Excel H, J열(Index 7, 9)로부터 기획 월평균 예상 판매량 데이터를 안전하게 매칭하여 반환합니다.
    (최근 90일 판매이력이 없더라도 재고가 있는 항목을 악성 재고로 포착하기 위해 Q~R열을 뼈대로 사용)
    """
    import urllib.parse
    import streamlit as st
    import numpy as np
    
    spreadsheet_id = "1s46doz2ahU90ygtpBQEn-RHh9Zy0mPZAVitIxFzg5jg"
    encoded_sheet = urllib.parse.quote("소진기한")
    
    exclude_keywords = ['dp', '비품', '증정', '샘플', '키링', '파우치', '시향', '테스터', '리필', 'gift', 'giftset', '부자재', '쇼카드']
    pattern = '|'.join(exclude_keywords)
    
    # 1. 구글 시트 데이터 로드 및 뼈대 구축
    try:
        url_all = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        df_raw = pd.read_csv(url_all, header=None)
        
        if df_raw.empty or df_raw.shape[1] < 18:
            raise ValueError("구글 시트 데이터 형태가 올바르지 않거나 열이 부족합니다.")
            
        # A. Excel S~T 열 (Index 18, 19) - 진짜 보유재고 데이터 파싱 및 본품 뼈대(Skeleton) 구축
        df_stock_master = df_raw.iloc[:, [18, 19]].copy()
        df_stock_master.columns = ['품명', '현재재고_raw']
        df_stock_master = df_stock_master.dropna(subset=['품명'])
        df_stock_master['품명'] = df_stock_master['품명'].astype(str).str.strip()
        df_stock_master = df_stock_master[~df_stock_master['품명'].isin(['품명', 'nan', ''])]
        
        # [데이터 전처리 및 노이즈 제거]: DP, 비품, 증정, 샘플, 시향, 부자재 키워드 제외 필터링 적용
        df_stock_master = df_stock_master[~df_stock_master['품명'].str.contains(pattern, case=False, na=False)]
        
        # 뼈대는 소진기한 탭의 130여개 전 품목을 그대로 유지함 (악성재고 포착을 위해 필터링하지 않음)
        
        df_stock_master['현재재고'] = pd.to_numeric(df_stock_master['현재재고_raw'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        df_stock = df_stock_master[['품명', '현재재고']].copy()
        
        # 조인용 표준화 키 생성
        df_stock['품명_clean'] = df_stock['품명'].apply(clean_product_name)
        
        # B. Excel Q~R 열 (Index 16, 17) - 기획 월평균 예상 판매량 데이터 파싱 (124개 풀 리스트)
        df_estimation_master = df_raw.iloc[:, [16, 17]].copy()
        df_estimation_master.columns = ['품명', '월평균 예상 판매량_raw']
        df_estimation_master = df_estimation_master.dropna(subset=['품명'])
        df_estimation_master['품명'] = df_estimation_master['품명'].astype(str).str.strip()
        df_estimation_master = df_estimation_master[~df_estimation_master['품명'].isin(['품명', 'nan', ''])]
        
        # [데이터 전처리 및 노이즈 제거]: DP, 비품, 증정, 샘플, 시향, 부자재 키워드 제외 필터링 적용
        df_estimation_master = df_estimation_master[~df_estimation_master['품명'].str.contains(pattern, case=False, na=False)]
        
        df_estimation_master['월평균 예상 판매량'] = pd.to_numeric(df_estimation_master['월평균 예상 판매량_raw'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        
        df_est_clean = df_estimation_master[['품명', '월평균 예상 판매량']].copy()
        df_est_clean['품명_clean'] = df_est_clean['품명'].apply(clean_product_name)
        
        # 중복 제거 (예상 판매량)
        est_subset = df_est_clean[['품명_clean', '월평균 예상 판매량']].drop_duplicates(subset=['품명_clean'])
        
        # 뼈대에 예상 판매량 Left Join
        df_stock = df_stock.merge(est_subset, on='품명_clean', how='left')
        df_stock = df_stock.drop(columns=['품명_clean'])
        
        # 결측치 정비
        df_stock['현재재고'] = df_stock['현재재고'].fillna(0.0)
        
    except Exception as e:
        st.error(f"⚠️ '소진기한' 탭 복합 열(Q~R, H~J) 전사 본품 연동 실패: {e}")
        df_stock = pd.DataFrame(columns=['품명', '현재재고', '월평균 예상 판매량'])
        
    # 2. _quantity_df 기반 SCM 과거 판매량 실적 데이터셋 생성
    try:
        df_sales = _quantity_df[['product_name', 'sale_date', 'quantity']].copy()
        df_sales = df_sales.rename(columns={
            'product_name': '품명',
            'sale_date': '날짜',
            'quantity': '수량'
        })
        df_sales['날짜'] = df_sales['날짜'].dt.strftime('%Y.%m.%d')
        df_sales['품명'] = df_sales['품명'].astype(str).str.strip()
        df_sales = df_sales[~df_sales['품명'].str.contains(pattern, case=False, na=False)]
    except Exception as e:
        st.error(f"⚠️ _quantity_df 기반 SCM 판매량 데이터셋 생성 실패: {e}")
        df_sales = pd.DataFrame(columns=['품명', '날짜', '수량'])
        
    return df_stock, df_sales

def clean_product_name(name):
    """
    완제품 품명의 미세한 텍스트 편차(공백, 영어 대소문자, 괄호 부속어 등)를 정밀 정제하고,
    다니엘트루스 고유의 밤쉘 향 표기 편차(밤쉘, 밤쉘루스, 시그니처 밤쉘루스)를 정형화하여
    조인 및 병합 시 매칭률을 극대화하는 표준화 키 생성 함수입니다.
    """
    import re
    if pd.isna(name):
        return ""
    name = str(name).strip().replace(" ", "").lower()
    # 대괄호 [.*?], 소괄호 (.*?) 및 내부 부속어 내용 삭제
    name = re.sub(r'\[.*?\]|\(.*?\)', '', name)
    
    # 다니엘트루스 고유의 밤쉘 향 표기 편차 완벽 보정
    name = name.replace("시그니처밤쉘루스", "밤쉘").replace("밤쉘루스", "밤쉘")
    return name

def calculate_sojin_metrics(df_stock, df_sales):
    """
    정제된 G:M 완제품 본품 마스터(df_stock)와 시계열 판매 실적(df_sales)을 표준화 퍼지 조인으로 결합하여
    과거 가중 판매속도(30일 90% + 60일 10%) 및 예상 소진기한/가중 소진기한 변동치를 산출합니다.
    """
    if df_stock.empty:
        return pd.DataFrame()
        
    max_date = pd.NaT
    
    # 1. 90일 판매 시계열 분석을 통한 가중치 계산
    if not df_sales.empty:
        df_sales = df_sales.copy()
        df_sales['날짜_dt'] = pd.to_datetime(df_sales['날짜'].astype(str).str.replace(' ', ''), format='%Y.%m.%d', errors='coerce')
        max_date = df_sales['날짜_dt'].max()
        
        if pd.isna(max_date):
            max_date = datetime.today()
            
        date_30_ago = max_date - pd.Timedelta(days=30)
        date_90_ago = max_date - pd.Timedelta(days=90)
        
        # 최근 30일 누적 판매량
        sales_30 = df_sales[
            (df_sales['날짜_dt'] > date_30_ago) & (df_sales['날짜_dt'] <= max_date)
        ].groupby('품명')['수량'].sum().reset_index(name='qty_recent_30')
        
        # 과거 31~90일 누적 판매량
        sales_90 = df_sales[
            (df_sales['날짜_dt'] > date_90_ago) & (df_sales['날짜_dt'] <= date_30_ago)
        ].groupby('품명')['수량'].sum().reset_index(name='qty_past_60')
        
        # 품명 표준화 키 기반 퍼지 조인 실행 (정밀 매칭 확보)
        df_stock_copy = df_stock.copy()
        sales_30_copy = sales_30.copy()
        sales_90_copy = sales_90.copy()
        
        df_stock_copy['품명_clean'] = df_stock_copy['품명'].apply(clean_product_name)
        sales_30_copy['품명_clean'] = sales_30_copy['품명'].apply(clean_product_name)
        sales_90_copy['품명_clean'] = sales_90_copy['품명'].apply(clean_product_name)
        
        # [핵심 교정] 중복 품명의 매출 수량을 품명_clean 기준으로 완전 합산(sum)하여 보존!
        s30_subset = sales_30_copy.groupby('품명_clean')['qty_recent_30'].sum().reset_index()
        s90_subset = sales_90_copy.groupby('품명_clean')['qty_past_60'].sum().reset_index()
        
        mrp_summary = df_stock_copy.merge(s30_subset, on='품명_clean', how='left').merge(s90_subset, on='품명_clean', how='left')
        mrp_summary = mrp_summary.drop(columns=['품명_clean'])
        mrp_summary['qty_recent_30'] = mrp_summary['qty_recent_30'].fillna(0.0)
        mrp_summary['qty_past_60'] = mrp_summary['qty_past_60'].fillna(0.0)
    else:
        mrp_summary = df_stock.copy()
        mrp_summary['qty_recent_30'] = 0.0
        mrp_summary['qty_past_60'] = 0.0
        
    # 2. 최근 30일(90%) + 과거 31~90일(10%) 가중 판매 속도(월간 환산) 계산 (사용자 요청 연산 공식 100% 명시적 반영)
    mrp_summary['가중판매속도_월'] = (
        ((mrp_summary['qty_recent_30'] * 0.9) + (mrp_summary['qty_past_60'] * 0.1)) / (3.0 * 30.0) * 30.0
    )
    
    # 0 나누기 오류 방어
    mrp_summary['가중판매속도_월'] = mrp_summary['가중판매속도_월'].apply(lambda x: max(x, 0.1))
    
    # 마케팅 예상치 결측시, '과거 90일 실제 판매 월평균치'를 기본값으로 동적 폴백 적용
    default_monthly_sales = (mrp_summary['qty_recent_30'] + mrp_summary['qty_past_60']) / 3.0
    default_monthly_sales = default_monthly_sales.apply(lambda x: max(x, 10.0))  # 최소 10개로 보정
    
    mrp_summary['월평균 예상 판매량'] = mrp_summary['월평균 예상 판매량'].fillna(default_monthly_sales)
    mrp_summary['월평균 예상 판매량'] = mrp_summary['월평균 예상 판매량'].apply(lambda x: max(x, 0.1))
    
    # 3. 소진 기한 계산 (개월 수)
    mrp_summary['예상소진기한_개월'] = mrp_summary['현재재고'] / mrp_summary['월평균 예상 판매량']
    mrp_summary['가중소진기한_개월'] = mrp_summary['현재재고'] / mrp_summary['가중판매속도_월']
    
    # 4. 예측 괴리도 계산 (절대값을 배제하고 과소/과대예측 구별하도록 부호 보존)
    mrp_summary['예측괴리도_Gap'] = mrp_summary['예상소진기한_개월'] - mrp_summary['가중소진기한_개월']
    
    # [소진 기한 개월수를 날짜로 동적 환산하는 로직]
    base_dt = max_date if not pd.isna(max_date) else datetime.today()
    
    def convert_months_to_date(months):
        if pd.isna(months) or months <= 0:
            return "즉시 소진"
        if months > 120:  # 10년이 넘는 장기 재고품
            return "장기 보관 (10년+)"
        days_to_add = int(months * 30.4375)
        target_dt = base_dt + pd.Timedelta(days=days_to_add)
        return target_dt.strftime('%Y-%m-%d')
        
    mrp_summary['예상소진기한_날짜'] = mrp_summary['예상소진기한_개월'].apply(convert_months_to_date)
    mrp_summary['가중소진기한_날짜'] = mrp_summary['가중소진기한_개월'].apply(convert_months_to_date)
    
    # [시각화용 캡핑 및 제곱근 스케일러 보정 설계]
    mrp_summary['예상소진기한_시각화'] = mrp_summary['예상소진기한_개월'].clip(lower=0.0, upper=24.0)
    mrp_summary['가중소진기한_시각화'] = mrp_summary['가중소진기한_개월'].clip(lower=0.0, upper=24.0)
    mrp_summary['현재재고_시각화'] = np.sqrt(mrp_summary['현재재고'].clip(lower=0.0)) * 1.5 + 6
    
    # 총 판매 수량 계산 (90일 실적 합산)
    mrp_summary['총판매수량_90일'] = mrp_summary['qty_recent_30'] + mrp_summary['qty_past_60']
    
    # 카테고리 컬럼 자동 분류
    def classify_prod_category(name):
        name = str(name)
        if '디퓨저' in name:
            return '🧪 디퓨저'
        elif '퍼퓸' in name or '향수' in name:
            return '🧴 오일 퍼퓸'
        elif '캔들' in name:
            return '🕯️ 캔들'
        elif '스프레이' in name:
            return '💨 룸 스프레이'
        elif '핸드' in name:
            return '🧴 핸드 케어'
        else:
            return '🎁 기타 완제품'
            
    mrp_summary['카테고리'] = mrp_summary['품명'].apply(classify_prod_category)
    
    return mrp_summary

@st.cache_data(ttl=300)
def load_google_sales_data():
    """
    구글 스프레드시트의 매출액 탭('25년 매출', '26년 매출') 및 판매량 탭('25년 판매', '26년 판매') 데이터를
    각각 로드하고 이중 데이터 파이프라인(Dual Data Pipeline) 구조로 전처리하여 반환합니다.
    """
    # ==========================================
    # 1. 실제 매출 데이터 (sales_df) 로드 및 전처리
    # ==========================================
    sales_dfs = []
    
    # 2025년 매출 데이터 로드 (2차 시트)
    try:
        encoded_sheet = urllib.parse.quote("25년 매출")
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_2_ID}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        df_25_sales = pd.read_csv(url)
        sales_dfs.append(df_25_sales)
    except Exception as e:
        st.warning(f"2025년 매출 데이터 로드 경고: {e}")
        
    # 2026년 매출 데이터 로드 (1차 시트)
    try:
        encoded_sheet = urllib.parse.quote("26년 매출")
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_1_ID}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        df_26_sales = pd.read_csv(url)
        sales_dfs.append(df_26_sales)
    except Exception as e:
        st.warning(f"2026년 매출 데이터 로드 경고: {e}")
        
    if not sales_dfs:
        raise ValueError("매출 데이터를 로드하는 데 실패했습니다.")
        
    sales_raw = pd.concat(sales_dfs, ignore_index=True)
    
    # 매출 데이터 표준화 및 날짜 전처리
    sales_raw['sale_date'] = pd.to_datetime(sales_raw['날짜'], errors='coerce')
    sales_raw = sales_raw.dropna(subset=['sale_date'])
    sales_raw['amount'] = pd.to_numeric(sales_raw['매출'], errors='coerce').fillna(0)
    sales_raw['store_name'] = sales_raw['채널/주차'].fillna("기타 매장").astype(str).str.strip()
    sales_raw['channel_type'] = sales_raw['카테고리'].fillna("오프라인").astype(str).str.strip()
    sales_raw['year'] = sales_raw['sale_date'].dt.year.astype(int)
    sales_raw['month'] = sales_raw['sale_date'].dt.month.astype(int)
    
    # 고유 매장 목록 추출 (실제 운영하는 매장만 추출됨)
    unique_stores = sorted(list(sales_raw['store_name'].unique()))
    
    # 고유 매장 코드 맵 생성
    store_code_map = {name: f"ST{i+1:03d}" for i, name in enumerate(unique_stores)}
    sales_raw['store_code'] = sales_raw['store_name'].map(store_code_map)
    sales_raw['sale_id'] = sales_raw.index.map(lambda idx: f"S_SAL_{idx+1:06d}")
    
    # ==========================================
    # 2. 완제품 판매량 데이터 (quantity_df) 로드 및 전처리
    # ==========================================
    qty_dfs = []
    
    # 2025년 판매 데이터 로드 (2차 시트)
    try:
        encoded_sheet = urllib.parse.quote("25년 판매")
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_2_ID}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        df_25_qty = pd.read_csv(url)
        qty_dfs.append(df_25_qty)
    except Exception as e:
        st.warning(f"2025년 판매 데이터 로드 경고: {e}")
        
    # 2026년 판매 데이터 로드 (1차 시트)
    try:
        encoded_sheet = urllib.parse.quote("26년 판매")
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_1_ID}/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
        df_26_qty = pd.read_csv(url)
        qty_dfs.append(df_26_qty)
    except Exception as e:
        st.warning(f"2026년 판매 데이터 로드 경고: {e}")
        
    if not qty_dfs:
        raise ValueError("판매 데이터를 로드하는 데 실패했습니다.")
        
    qty_raw = pd.concat(qty_dfs, ignore_index=True)
    
    # 판매 데이터 날짜 및 수량 전처리
    qty_raw['sale_date'] = pd.to_datetime(qty_raw['날짜'], errors='coerce')
    qty_raw = qty_raw.dropna(subset=['sale_date'])
    qty_raw['quantity'] = pd.to_numeric(qty_raw['수량'], errors='coerce').fillna(0).astype(int)
    qty_raw['year'] = qty_raw['sale_date'].dt.year.astype(int)
    qty_raw['month'] = qty_raw['sale_date'].dt.month.astype(int)
    
    # 품명 전처리 및 [증정], [샘플] 품목 원천 제외
    qty_raw['product_name'] = qty_raw['품명'].fillna("기타 품목").astype(str).str.strip()
    qty_raw = qty_raw[~qty_raw['product_name'].str.startswith("[증정]") & ~qty_raw['product_name'].str.startswith("[샘플]")]
    
    # 정밀 카테고리 분류 규칙 정의 (괄호 안의 코드 파싱 접두사 매칭 고도화)
    def classify_category(p_name):
        import re
        if "[증정]" in p_name or "[샘플]" in p_name:
            return "증정"
            
        # 괄호 안의 제품 코드 파싱
        match = re.search(r'\((.*?)\)', p_name)
        if match:
            code = match.group(1).upper().strip()
            if code.startswith('D'):
                return "디퓨저"
            elif code.startswith('O'):
                return "오일 퍼퓸"
            elif code.startswith('H'):
                return "핸드 크림"
            elif code.startswith('R'):
                return "룸 스프레이"
            elif code.startswith('C'):
                return "캔들"
            elif code.startswith('A'):
                return "액세서리"
                
        # 괄호 코드가 없을 때 예외 텍스트 기반 분류 (Fallback)
        if "디퓨저" in p_name:
            return "디퓨저"
        elif "오일" in p_name:
            return "오일 퍼퓸"
        elif "핸드" in p_name or "크림" in p_name:
            return "핸드 크림"
        elif "룸" in p_name or "스프레이" in p_name:
            return "룸 스프레이"
        elif "캔들" in p_name:
            return "캔들"
        elif "액세서리" in p_name:
            return "액세서리"
            
        return "기타"
            
    qty_raw['category'] = qty_raw['product_name'].apply(classify_category)
    
    # 매출액 계산 (카테고리별 단가 기준 수량 곱연산)
    price_map = {
        '디퓨저': 45000,
        '오일 퍼퓸': 35000,
        '핸드 크림': 25000,
        '룸 스프레이': 32000,
        '캔들': 28000,
        '액세서리': 15000,
        '증정': 0,
        '기타': 10000
    }
    
    # ------------------------------------------
    # 2.A. 일자별 매출 기여도 기반 비례 배분 알고리즘 도입
    # ------------------------------------------
    # 일자별/매장별 매출 기여도(share) 연산
    daily_store_sales = sales_raw.groupby(['sale_date', 'store_name'])['amount'].sum().reset_index()
    daily_total_sales = sales_raw.groupby('sale_date')['amount'].sum().reset_index().rename(columns={'amount': 'total_amount'})
    daily_share = pd.merge(daily_store_sales, daily_total_sales, on='sale_date')
    daily_share['share'] = daily_share.apply(
        lambda r: r['amount'] / r['total_amount'] if r['total_amount'] > 0 else 0.0, axis=1
    )
    
    # 전체 매장 평균 매출 비중 (매출 데이터가 존재하지 않는 일자의 대체 처리용)
    overall_store_sales = sales_raw.groupby('store_name')['amount'].sum().reset_index()
    overall_total = overall_store_sales['amount'].sum()
    if overall_total > 0:
        overall_store_sales['share'] = overall_store_sales['amount'] / overall_total
    else:
        overall_store_sales['share'] = 1.0 / len(unique_stores)
    overall_share_map = overall_store_sales.set_index('store_name')['share'].to_dict()

    # 날짜별 매장 분배 매트릭스 동적 구축
    qty_dates = pd.DataFrame({'sale_date': qty_raw['sale_date'].unique()})
    dist_list = []
    for q_date in qty_dates['sale_date']:
        date_shares = daily_share[daily_share['sale_date'] == q_date]
        if not date_shares.empty:
            for _, row in date_shares.iterrows():
                dist_list.append({
                    'sale_date': q_date,
                    'store_name': row['store_name'],
                    'share': row['share']
                })
        else:
            for s_name, s_share in overall_share_map.items():
                dist_list.append({
                    'sale_date': q_date,
                    'store_name': s_name,
                    'share': s_share
                })
    dist_df = pd.DataFrame(dist_list)
    
    # 수량 데이터와 날짜별 매장 기여도 매트릭스 결합
    qty_allocated = pd.merge(qty_raw, dist_df, on='sale_date')
    
    # 매장별 기여비율에 따라 수량 분배 및 정수화
    qty_allocated['quantity'] = (qty_allocated['quantity'] * qty_allocated['share']).round().astype(int)
    
    # 수량이 0인 실적 데이터 제외 (리포팅 속도 및 데이터 군더더기 배제)
    qty_allocated = qty_allocated[qty_allocated['quantity'] > 0].reset_index(drop=True)
    
    # 최종 매장 정보 매핑
    qty_allocated['store_code'] = qty_allocated['store_name'].map(store_code_map)
    qty_allocated['channel_type'] = qty_allocated['store_name'].apply(
        lambda x: '온라인' if any(kw in x for kw in ['온라인', '스토어', '선물하기']) else '오프라인'
    )
    
    # 정밀 카테고리 단가를 대입한 배분 매출액 재연산
    qty_allocated['amount'] = qty_allocated.apply(
        lambda row: row['quantity'] * price_map.get(row['category'], 10000), axis=1
    )
    
    # 제품 코드 동적 발급 및 고유 매핑
    unique_products = qty_allocated['product_name'].unique()
    prod_code_map = {name: f"PROD_{i+1:03d}" for i, name in enumerate(unique_products)}
    qty_allocated['product_code'] = qty_allocated['product_name'].map(prod_code_map)
    qty_allocated['sale_id'] = qty_allocated.index.map(lambda idx: f"S_{idx+1:06d}")
    
    return sales_raw, qty_allocated, store_code_map, prod_code_map

@st.cache_data
def load_derived_master_data(store_names_list, prod_code_map):
    """
    구글 시트의 매장 및 제품 데이터를 활용하여 다차원 스타 스키마 마스터 데이터와 동적 BOM 관계를 생성합니다.
    """
    # 1. 매장 마스터 생성
    store_records = []
    for i, name in enumerate(store_names_list):
        if '온라인' in name or '스토어' in name or '선물하기' in name:
            region = '전국'
        else:
            region = '수도권'
            
        store_records.append({
            'store_code': f"ST{i+1:03d}",
            'store_name': name,
            'region': region
        })
    store_df = pd.DataFrame(store_records)
    
    # 2. 제품 마스터 생성
    product_records = []
    def classify_category(p_name):
        if p_name.startswith("[증정]") or p_name.startswith("[샘플]"):
            return "증정"
        elif p_name.startswith("디퓨저"):
            return "디퓨저"
        elif p_name.startswith("오일"):
            return "오일 퍼퓸"
        elif p_name.startswith("캔들"):
            return "캔들"
        elif p_name.startswith("[액세서리]"):
            return "액세서리"
        else:
            return "기타"
            
    for prod_name, prod_code in prod_code_map.items():
        cat = classify_category(prod_name)
        product_records.append({
            'product_code': prod_code,
            'product_name': prod_name,
            'category': cat,
            'unit': '개'
        })
    product_df = pd.DataFrame(product_records)
    
    # 3. 달력 마스터
    calendar_df = pd.DataFrame([
        {'date_id': '2025-05-21', 'year': 2025, 'month': 5, 'day': 21, 'day_of_week': 'Wednesday'}
    ])
    
    # 4. 동적 BOM 정의 (카테고리별 자재 소요량 전개)
    bom_records = []
    bom_idx = 1
    for _, row in product_df.iterrows():
        p_code = row['product_code']
        p_name = row['product_name']
        p_cat = row['category']
        
        if p_cat == '디퓨저':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_BASE', 'child_name': '디퓨저 베이스 에탄올 (ml)', 'unit_qty': 150.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_FRAG', 'child_name': '디퓨저 조합 향료 오일 (ml)', 'unit_qty': 50.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_DF_GLASS', 'child_name': '프리미엄 디퓨저 유리 용기 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '오일 퍼퓸':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_OL_FRAG', 'child_name': '천연 에센셜 향료 원액 (ml)', 'unit_qty': 8.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_OL_BOTTLE', 'child_name': '롤온 고급 초자 유리병 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '캔들':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_WAX', 'child_name': '천연 골든 소이 왁스 (g)', 'unit_qty': 180.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_FRAG', 'child_name': '캔들 가열용 향료 오일 (ml)', 'unit_qty': 20.0
            })
            bom_idx += 1
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_CD_GLASS', 'child_name': '내열성 캔들 유리 용기 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '액세서리':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_AC_BOX', 'child_name': '수작업 패키징 상자 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
        elif p_cat == '증정':
            bom_records.append({
                'bom_id': f"BOM{bom_idx:04d}", 'parent_code': p_code, 'parent_name': p_name,
                'child_code': 'RAW_PR_MINI', 'child_name': '미니어처 전용 샘플 공병 (개)', 'unit_qty': 1.0
            })
            bom_idx += 1
            
    bom_df = pd.DataFrame(bom_records)
    
    # 5. 동적 생산 계획 정의 (2025년 가상 생산량)
    plan_records = []
    plan_idx = 1
    for _, row in product_df.iterrows():
        p_code = row['product_code']
        p_name = row['product_name']
        p_cat = row['category']
        
        if p_cat == '디퓨저':
            planned_qty = 1500
        elif p_cat == '오일 퍼퓸':
            planned_qty = 2500
        elif p_cat == '캔들':
            planned_qty = 1200
        elif p_cat == '액세서리':
            planned_qty = 800
        elif p_cat == '증정':
            planned_qty = 5000
        else:
            planned_qty = 500
            
        plan_records.append({
            'plan_id': f"PLN{plan_idx:04d}",
            'plan_month': '2025-05',
            'product_code': p_code,
            'product_name': p_name,
            'planned_qty': planned_qty
        })
        plan_idx += 1
        
    plan_df = pd.DataFrame(plan_records)
    
    return store_df, product_df, calendar_df, bom_df, plan_df

# 3. 실시간 모의 데이터 스트리밍 세션 상태 관리
if 'additional_sales' not in st.session_state:
    st.session_state.additional_sales = []
if 'last_update_time' not in st.session_state:
    st.session_state.last_update_time = datetime.now()
if 'websocket_log' not in st.session_state:
    st.session_state.websocket_log = ["System Initialized - Connected to Anti-Gravity Broker"]

def simulate_realtime_stream(store_names, prod_code_map):
    """실시간 판매 발생 시뮬레이터 (웹소켓 CDC 패치 시뮬레이션)"""
    products = list(prod_code_map.keys())
    prices = {
        '디퓨저': 45000,
        '오일 퍼퓸': 35000,
        '캔들': 28000,
        '액세서리': 15000,
        '증정': 0,
        '기타': 10000
    }
    
    def classify_category(p_name):
        if p_name.startswith("[증정]") or p_name.startswith("[샘플]"):
            return "증정"
        elif p_name.startswith("디퓨저"):
            return "디퓨저"
        elif p_name.startswith("오일"):
            return "오일 퍼퓸"
        elif p_name.startswith("캔들"):
            return "캔들"
        elif p_name.startswith("[액세서리]"):
            return "액세서리"
        else:
            return "기타"
            
    # 1~3건의 신규 주문 랜덤 발생
    new_tx_count = np.random.randint(1, 4)
    now_str = datetime.now().strftime('%Y-%m-%d')
    
    new_records = []
    for _ in range(new_tx_count):
        store = np.random.choice(store_names)
        prod = np.random.choice(products)
        qty = np.random.randint(1, 4)
        cat = classify_category(prod)
        amount = qty * prices.get(cat, 10000)
        tx_id = f"S_STREAM_{np.random.randint(50000, 99999)}"
        
        record = {
            'sale_id': tx_id,
            'sale_date': pd.to_datetime(now_str),
            'store_code': "ST_STREAM",
            'store_name': store,
            'channel_type': '온라인' if any(kw in store for kw in ['온라인', '스토어', '선물하기']) else '오프라인',
            'product_code': prod_code_map[prod],
            'product_name': prod,
            'quantity': qty,
            'amount': amount,
            'category': cat,
            'year': pd.to_datetime(now_str).year,
            'month': pd.to_datetime(now_str).month
        }
        new_records.append(record)
        st.session_state.additional_sales.append(record)
    
    # 웹소켓 디버그 로그 추가
    timestamp = datetime.now().strftime('%H:%M:%S')
    st.session_state.websocket_log.insert(
        0, f"[{timestamp}] [WS:ds_update] Incremental push received: {len(new_records)} items added."
    )
    if len(st.session_state.websocket_log) > 10:
        st.session_state.websocket_log.pop()
        
    st.session_state.last_update_time = datetime.now()

# 데이터 로드
try:
    base_sales_df, base_qty_df, store_code_map, prod_code_map = load_google_sales_data()
    # 구글 시트 생산 계획 및 BOM 데이터 실시간 로드 및 품명 매칭 적용
    plan_df, bom_df, updated_prod_code_map = load_google_mrp_data(prod_code_map)
    # 갱신된 제품 코드 맵을 사용하여 기존 마스터 데이터 구축
    store_df, product_df, calendar_df, _, _ = load_derived_master_data(list(store_code_map.keys()), updated_prod_code_map)
    # prod_code_map을 갱신된 것으로 대체
    prod_code_map = updated_prod_code_map
    
    # ==========================================
    # 구글 시트 BOM 누락 완제품에 대한 동적 가상 자재 자동 충진 세이프가드
    # ==========================================
    mrp_all_products = plan_df[['product_code', 'product_name']].drop_duplicates()
    missing_bom_products = mrp_all_products[~mrp_all_products['product_code'].isin(bom_df['parent_code'].unique())]
    
    if not missing_bom_products.empty:
        filled_records = []
        fb_idx = 1000
        
        for _, p_row in missing_bom_products.iterrows():
            p_code = p_row['product_code']
            p_name = p_row['product_name']
            
            # 카테고리 판별
            p_cat = '기타'
            if "디퓨저" in p_name:
                p_cat = '디퓨저'
            elif "오일" in p_name or "퍼퓸" in p_name:
                p_cat = '오일 퍼퓸'
            elif "캔들" in p_name:
                p_cat = '캔들'
            elif "액세서리" in p_name or "스틱" in p_name or "캡" in p_name:
                p_cat = '액세서리'
            elif "증정" in p_name or "샘플" in p_name:
                p_cat = '증정'
                
            # 디퓨저 관련 부품/액세서리인 경우 (예: 브라운 섬유 리드스틱)
            if p_cat == '액세서리' and "디퓨저" in p_name:
                # 패키징 상자 소량 소요
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}",
                    'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_AC_BOX', 'child_name': '수작업 패키징 상자 (개)',
                    'unit_qty': 1.0
                })
                fb_idx += 1
                # 리드 스틱 소요 (15p일 경우 15개, 아닐 경우 기본 5개)
                stick_qty = 15.0 if "15p" in p_name else 5.0
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}",
                    'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_DF_STICK', 'child_name': '프리미엄 리드 스틱 (개)',
                    'unit_qty': stick_qty
                })
                fb_idx += 1
            elif p_cat == '디퓨저':
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_DF_BASE', 'child_name': '디퓨저 베이스 에탄올 (ml)', 'unit_qty': 150.0
                })
                fb_idx += 1
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_DF_FRAG', 'child_name': '디퓨저 조합 향료 오일 (ml)', 'unit_qty': 50.0
                })
                fb_idx += 1
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_DF_GLASS', 'child_name': '프리미엄 디퓨저 유리 용기 (개)', 'unit_qty': 1.0
                })
                fb_idx += 1
            elif p_cat == '오일 퍼퓸':
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_OL_FRAG', 'child_name': '천연 에센셜 향료 원액 (ml)', 'unit_qty': 8.0
                })
                fb_idx += 1
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_OL_BOTTLE', 'child_name': '롤온 고급 초자 유리병 (개)', 'unit_qty': 1.0
                })
                fb_idx += 1
            elif p_cat == '캔들':
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_CD_WAX', 'child_name': '천연 골든 소이 왁스 (g)', 'unit_qty': 180.0
                })
                fb_idx += 1
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_CD_FRAG', 'child_name': '캔들 가열용 향료 오일 (ml)', 'unit_qty': 20.0
                })
                fb_idx += 1
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_CD_GLASS', 'child_name': '내열성 캔들 유리 용기 (개)', 'unit_qty': 1.0
                })
                fb_idx += 1
            elif p_cat == '증정':
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_PR_MINI', 'child_name': '미니어처 전용 샘플 공병 (개)', 'unit_qty': 1.0
                })
                fb_idx += 1
            else:
                filled_records.append({
                    'bom_id': f"BOM_FILL_{fb_idx}", 'parent_code': p_code, 'parent_name': p_name,
                    'child_code': 'RAW_AC_BOX', 'child_name': '수작업 패키징 상자 (개)', 'unit_qty': 1.0
                })
                fb_idx += 1
                
        if filled_records:
            filled_bom_df = pd.DataFrame(filled_records)
            bom_df = pd.concat([bom_df, filled_bom_df], ignore_index=True)
            
except Exception as e:
    st.error(f"데이터 로드 실패: {e}.")
    st.stop()

# 세션에 있는 실시간 스트리밍 데이터를 기본 매출 및 판매 수량 데이터에 각각 병합
if st.session_state.additional_sales:
    streamed_df = pd.DataFrame(st.session_state.additional_sales)
    sales_df = pd.concat([base_sales_df, streamed_df], ignore_index=True)
    quantity_df = pd.concat([base_qty_df, streamed_df], ignore_index=True)
else:
    sales_df = base_sales_df.copy()
    quantity_df = base_qty_df.copy()

# --- 사이드바: 컨트롤러 & 실시간 시뮬레이션 설정 ---
with st.sidebar:
    # 브랜드 로고 이미지 렌더링 (파일이 존재하는 경우에만 로드, 없으면 고급스러운 텍스트 타이틀로 대체)
    if os.path.exists("daniels_truth_logo.png"):
        st.image("daniels_truth_logo.png", use_container_width=True)
    else:
        st.markdown(
            """
            <div style='text-align: center; padding: 12px; border: 2px solid #C5A880; border-radius: 6px; background-color: #1a1a1a; margin-bottom: 10px;'>
                <h2 style='color: #C5A880; margin: 0; font-family: "Outfit", sans-serif; font-weight: 700; letter-spacing: 2px; font-size: 1.5rem;'>LUXURY SCM</h2>
                <span style='color: #888888; font-size: 0.8rem; letter-spacing: 1px; font-weight: 500;'>PORTFOLIO EDITION</span>
            </div>
            """, 
            unsafe_allow_html=True
        )
    
    st.markdown("<hr style='border: 1px solid #C5A880; margin-top: 10px; margin-bottom: 15px;'/>", unsafe_allow_html=True)
    
    # 실시간 스트리밍 토글
    st.write("### ⚡ 실시간 데이터 동기화 (CDC)")
    auto_refresh = st.checkbox("WebSocket 실시간 갱신 활성화", value=False)
    refresh_rate = st.slider("자동 갱신 주기 (초)", min_value=1, max_value=5, value=3)
    
    if auto_refresh:
        simulate_realtime_stream(list(store_code_map.keys()), prod_code_map)
        st.info("실시간 스트리밍 활성화됨. 데이터를 자동으로 로드하고 집계합니다.")
        time.sleep(refresh_rate)
        st.rerun()
        
    st.write("### 🔄 수동 데이터 제어")
    if st.button("구글 시트 실시간 동기화"):
        st.cache_data.clear()
        st.toast("구글 스프레드시트 실시간 동기화 완료!", icon="🔄")
        timestamp = datetime.now().strftime('%H:%M:%S')
        st.session_state.websocket_log.insert(
            0, f"[{timestamp}] [Manual Sync] Google Sheets cache invalidated. Full reload performed."
        )
        st.rerun()
        
    st.markdown("---")
    
    # Global Filters 구현
    st.write("### 🔍 글로벌 필터 (Global Filters)")
    
    # 1. 날짜 범위 필터
    sales_df['sale_date'] = pd.to_datetime(sales_df['sale_date'])
    min_date = sales_df['sale_date'].min().date()
    max_date = sales_df['sale_date'].max().date()
    
    date_range = st.date_input(
        "조회 기간 선택",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    
    # 2. 매장별 필터
    store_options = ["전체"] + list(store_df['store_name'].unique())
    selected_store = st.selectbox("매장별 필터", store_options)
    
    # 3. 5대 카테고리 필터 (요청하신 순서대로 명시적 정렬 및 '기타' 최하단 배치)
    cats_order = ['디퓨저', '오일 퍼퓸', '캔들', '액세서리', '증정', '기타']
    category_options = ["전체"] + [c for c in cats_order if c in product_df['category'].unique()]
    selected_category = st.selectbox("카테고리 필터", category_options)
    
    # 4. 품목 필터 (완제품 - 카테고리 선택에 연동된 동적 목록 제공)
    if selected_category != "전체":
        filtered_prods = product_df[product_df['category'] == selected_category]['product_name'].unique()
    else:
        filtered_prods = product_df['product_name'].unique()
    product_options = ["전체"] + list(filtered_prods)
    selected_product = st.selectbox("품목 필터 (완제품)", product_options)

# 연도 및 월 컬럼을 미리 계산해 둠
sales_df = sales_df.dropna(subset=['sale_date'])
sales_df['year'] = sales_df['sale_date'].dt.year.astype(int)
sales_df['month'] = sales_df['sale_date'].dt.month.astype(int)
sales_df['year_month_str'] = sales_df.apply(lambda r: f"{int(r['year'])}년 {int(r['month']):02d}월", axis=1)

quantity_df = quantity_df.dropna(subset=['sale_date'])
quantity_df['year'] = quantity_df['sale_date'].dt.year.astype(int)
quantity_df['month'] = quantity_df['sale_date'].dt.month.astype(int)
quantity_df['year_month_str'] = quantity_df.apply(lambda r: f"{int(r['year'])}년 {int(r['month']):02d}월", axis=1)

# 데이터 필터링 적용 (이중 데이터 파이프라인 개별 적용)
filtered_sales = sales_df.copy()
filtered_qty = quantity_df.copy()

# 날짜 필터 적용
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_dt, end_dt = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    filtered_sales = filtered_sales[(filtered_sales['sale_date'] >= start_dt) & (filtered_sales['sale_date'] <= end_dt)]
    filtered_qty = filtered_qty[(filtered_qty['sale_date'] >= start_dt) & (filtered_qty['sale_date'] <= end_dt)]

# 매장 필터 적용
if selected_store != "전체":
    filtered_sales = filtered_sales[filtered_sales['store_name'] == selected_store]
    filtered_qty = filtered_qty[filtered_qty['store_name'] == selected_store]

# 카테고리 필터 적용 (quantity_df 수량 전용, 실제 매출 탭에는 품목 카테고리가 없으므로 미적용 안내)
if selected_category != "전체":
    filtered_qty = filtered_qty[filtered_qty['category'] == selected_category]
    st.sidebar.caption("⚠️ 카테고리 필터는 완제품 수량 및 BOM 예측 지표에만 적용됩니다. (실제 매출 탭에는 품목 정보 없음)")

# 품목 필터 적용 (quantity_df 수량 전용)
if selected_product != "전체":
    filtered_qty = filtered_qty[filtered_qty['product_name'] == selected_product]
    st.sidebar.caption("⚠️ 품목 필터는 완제품 수량 및 BOM 예측 지표에만 적용됩니다.")



# --- 메인 화면 헤더 ---
row_header = st.columns([3, 1])
with row_header[0]:
    st.title("다니엘트루스 대시보드")

with row_header[1]:
    st.markdown("""
    <div style='text-align: right; margin-top: 10px;'>
        <div class='ws-status-connected'>
            <span class='ws-dot'></span>
            WebSocket: Connected (Broker)
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 탭 배치
tab1, tab2, tab3, tab4 = st.tabs(["⚡ 월간 성과 관리 센터", "📊 다개년 채널별 분석", "⚙️ 생산 및 BOM 예측", "📈 소진 기한 및 트렌드 분석"])

# --- 헬퍼 함수: 스파크라인 생성 ---
def make_sparkline(y_data, color='#C5A880'):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        y=y_data,
        mode='lines+markers',
        line=dict(color=color, width=2.5),
        marker=dict(size=4, opacity=0),  # 마우스 오버 시에만 마커 강조
        fill='tozeroy',
        fillcolor=f'rgba{tuple(list(int(color.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)) + [0.08])}', # hex to rgba
        hovertemplate='<b>₩%{y:,.0f}</b><extra></extra>'
    ))
    fig.update_layout(
        xaxis=dict(visible=False, showgrid=False),
        yaxis=dict(visible=False, showgrid=False),
        showlegend=False,
        margin=dict(l=2, r=2, t=2, b=2),
        height=38,
        width=110,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor='#0B132B',
            bordercolor='#C5A880',
            font_size=10,
            font_color='#F8F7F4',
            font_family='Outfit'
        )
    )
    return fig


# --- TAB 1: 성과 관리 센터 ---
with tab1:
    col_date, col_title = st.columns([1.5, 2.5])
    with col_date:
        # 데이터 내의 고유한 날짜(일자) 추출 및 정렬
        unique_dates = pd.to_datetime(sales_df['sale_date']).dt.date.drop_duplicates()
        unique_dates = sorted(unique_dates, reverse=True)  # 최신순
        date_options = [d.strftime('%Y-%m-%d') for d in unique_dates]
        
        # 오늘 날짜를 기준으로 기본값 탐색
        today = datetime.today().date()
        today_str = today.strftime('%Y-%m-%d')
        
        if today_str in date_options:
            default_idx = date_options.index(today_str)
        else:
            default_idx = 0
            
        target_date_str = st.selectbox(
            "📅 조회 기준일 선택", 
            options=date_options, 
            index=default_idx, 
            key="daily_target_date"
        )
        
    target_date = pd.to_datetime(target_date_str)
    
    # 전주 동일 요일(D-7) 연산
    prev_date = target_date - timedelta(days=7)
    prev_date_str = prev_date.strftime('%Y-%m-%d')
    
    with col_title:
        st.markdown(
            f"<div style='text-align: right; padding-top: 15px; color: #4b5563; font-weight: 600;'>"
            f"기준일: {target_date_str} / 전주 동일 요일 대비 비교 대상일: {prev_date_str}</div>", 
            unsafe_allow_html=True
        )
        
    # 당일 및 전주 동일 요일 데이터 필터링 (매출 vs 수량 이중화)
    today_sales_df = sales_df[pd.to_datetime(sales_df['sale_date']).dt.date == target_date.date()]
    prev_sales_df = sales_df[pd.to_datetime(sales_df['sale_date']).dt.date == prev_date.date()]
    
    today_qty_df = quantity_df[pd.to_datetime(quantity_df['sale_date']).dt.date == target_date.date()]
    prev_qty_df = quantity_df[pd.to_datetime(quantity_df['sale_date']).dt.date == prev_date.date()]
    
    # 1. 온/오프라인 구분 매출액 및 전주대비 지표 연산
    today_online_df = today_sales_df[today_sales_df['channel_type'] == '온라인']
    prev_online_df = prev_sales_df[prev_sales_df['channel_type'] == '온라인']
    today_online_sales = today_online_df['amount'].sum()
    prev_online_sales = prev_online_df['amount'].sum()
    online_delta_pct = ((today_online_sales - prev_online_sales) / prev_online_sales * 100) if prev_online_sales > 0 else 0.0
    
    today_offline_df = today_sales_df[today_sales_df['channel_type'] == '오프라인']
    prev_offline_df = prev_sales_df[prev_sales_df['channel_type'] == '오프라인']
    today_offline_sales = today_offline_df['amount'].sum()
    prev_offline_sales = prev_offline_df['amount'].sum()
    offline_delta_pct = ((today_offline_sales - prev_offline_sales) / prev_offline_sales * 100) if prev_offline_sales > 0 else 0.0
    
    # 2. 온/오프라인 각각 7일 일별 스파크라인 트렌드 데이터 수집 (기준일 포함 직전 7일)
    days_7 = []
    for i in range(6, -1, -1):
        days_7.append(target_date - timedelta(days=i))
        
    spark_data_online = []
    spark_data_offline = []
    for d in days_7:
        d_df = sales_df[pd.to_datetime(sales_df['sale_date']).dt.date == d.date()]
        spark_data_online.append(d_df[d_df['channel_type'] == '온라인']['amount'].sum())
        spark_data_offline.append(d_df[d_df['channel_type'] == '오프라인']['amount'].sum())
        
    online_color = '#0B132B' if online_delta_pct >= 0 else '#EF4444'
    offline_color = '#C5A880' if offline_delta_pct >= 0 else '#EF4444'
    
    fig_online_spark = make_sparkline(spark_data_online, online_color)
    fig_offline_spark = make_sparkline(spark_data_offline, offline_color)
    
    # 3. 온/오프라인 매출액 전주대비 KPI 카드 렌더링
    kpi_cols = st.columns(2)
    
    with kpi_cols[0]:
        with st.container(border=True):
            sc1, sc2 = st.columns([1.7, 1.3])
            with sc1:
                st.markdown(f"""
                <span style='font-size: 12px; color: #71717a; font-weight: 600;'>💻 온라인 매출액</span>
                <div style='font-size: 24px; font-weight: 800; color: #1E3A8A; margin: 4px 0;'>₩{today_online_sales:,.0f}</div>
                """, unsafe_allow_html=True)
                if online_delta_pct >= 0:
                    st.markdown(f"<span style='font-size: 11px; color: #10b981; font-weight: 600;'>전주 대비 +{online_delta_pct:.1f}% ▲</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span style='font-size: 11px; color: #ef4444; font-weight: 600;'>전주 대비 {online_delta_pct:.1f}% ▼</span>", unsafe_allow_html=True)
            with sc2:
                st.plotly_chart(fig_online_spark, use_container_width=True, key="online_sales_spark")
                
    with kpi_cols[1]:
        with st.container(border=True):
            sc1, sc2 = st.columns([1.7, 1.3])
            with sc1:
                st.markdown(f"""
                <span style='font-size: 12px; color: #71717a; font-weight: 600;'>🏪 오프라인 매출액</span>
                <div style='font-size: 24px; font-weight: 800; color: #10b981; margin: 4px 0;'>₩{today_offline_sales:,.0f}</div>
                """, unsafe_allow_html=True)
                if offline_delta_pct >= 0:
                    st.markdown(f"<span style='font-size: 11px; color: #10b981; font-weight: 600;'>전주 대비 +{offline_delta_pct:.1f}% ▲</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span style='font-size: 11px; color: #ef4444; font-weight: 600;'>전주 대비 {offline_delta_pct:.1f}% ▼</span>", unsafe_allow_html=True)
            with sc2:
                st.plotly_chart(fig_offline_spark, use_container_width=True, key="offline_sales_spark")
                
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 메인 2분할 레이아웃
    body_cols = st.columns([1.6, 1.4])
    
    with body_cols[0]:
        st.markdown("#### 💳 당일 채널별 매출 비중")
        if not today_sales_df.empty:
            # 매출 비중 집계 (sales_df 기반)
            channel_sales_agg = today_sales_df.groupby('channel_type').agg({'amount': 'sum'}).reset_index()
            
            fig_chan_sales = px.pie(
                channel_sales_agg,
                values='amount',
                names='channel_type',
                hole=0.45,
                color='channel_type',
                color_discrete_map={'온라인': '#0B132B', '오프라인': '#C5A880'},
                labels={'amount': '매출액'}
            )
            fig_chan_sales.update_traces(
                textposition='inside', 
                textinfo='percent+label',
                hovertemplate='<b>%{label} 채널</b><br>당일 매출액: ₩%{value:,.0f}<br>점유율: %{percent}<extra></extra>'
            )
            fig_chan_sales.update_layout(
                margin=dict(l=30, r=30, t=10, b=20),
                height=260,
                legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                hoverlabel=dict(
                    bgcolor='#0B132B',
                    bordercolor='#C5A880',
                    font_size=12,
                    font_color='#F8F7F4',
                    font_family='Outfit'
                )
            )
            st.plotly_chart(fig_chan_sales, use_container_width=True, key="today_chan_sales")
        else:
            st.info("선택한 날짜에 채널 실적 데이터가 존재하지 않습니다.")
            
    with body_cols[1]:
        st.markdown("#### 🏆 구매순위 (당일 판매수량 Top 10)")
        
        if not today_qty_df.empty:
            # 당일 완제품 판매수량 순위
            today_prod_sales = today_qty_df.groupby('product_name')['quantity'].sum().reset_index()
            today_prod_sales = today_prod_sales.sort_values(by='quantity', ascending=False).reset_index(drop=True)
            today_prod_sales['rank'] = today_prod_sales.index + 1
            
            # 전주 동일 요일 완제품 판매수량 순위
            if not prev_qty_df.empty:
                prev_prod_sales = prev_qty_df.groupby('product_name')['quantity'].sum().reset_index()
                prev_prod_sales = prev_prod_sales.sort_values(by='quantity', ascending=False).reset_index(drop=True)
                prev_prod_sales['prev_rank'] = prev_prod_sales.index + 1
            else:
                prev_prod_sales = pd.DataFrame(columns=['product_name', 'quantity', 'prev_rank'])
                
            top10_today = today_prod_sales.head(10).copy()
            
            # 순위 변동 계산 로직
            def calc_rank_delta(row):
                p_name = row['product_name']
                curr_rank = row['rank']
                
                prev_match = prev_prod_sales[prev_prod_sales['product_name'] == p_name]
                if prev_match.empty:
                    return 'NEW'
                    
                prev_rank = prev_match.iloc[0]['prev_rank']
                diff = prev_rank - curr_rank
                
                if diff > 0:
                    return f"▲ {diff}"
                elif diff < 0:
                    return f"▼ {abs(diff)}"
                else:
                    return "-"
                    
            top10_today['delta'] = top10_today.apply(calc_rank_delta, axis=1)
            
            # HTML/CSS 기반 프리미엄 랭킹 테이블 렌더링
            css_style = (
                "<style>"
                ".rank-table {"
                "    width: 100%;"
                "    border-collapse: collapse;"
                "    font-size: 13px;"
                "    margin-top: 5px;"
                "}"
                ".rank-table th {"
                "    border-bottom: 2px solid #e2e8f0;"
                "    color: #4b5563;"
                "    font-weight: 600;"
                "    text-align: left;"
                "    padding: 6px 10px;"
                "}"
                ".rank-table td {"
                "    padding: 8px 10px;"
                "    border-bottom: 1px solid #f1f5f9;"
                "    vertical-align: middle;"
                "}"
                ".rank-num {"
                "    font-weight: 800;"
                "    color: #1e3a8a;"
                "    font-size: 13px;"
                "}"
                ".prod-name {"
                "    font-weight: 600;"
                "    color: #334155;"
                "}"
                ".qty-val {"
                "    font-weight: 700;"
                "    color: #0f172a;"
                "}"
                ".delta-new {"
                "    color: #f59e0b;"
                "    font-weight: bold;"
                "    background-color: #fef3c7;"
                "    padding: 2px 6px;"
                "    border-radius: 4px;"
                "    font-size: 10px;"
                "}"
                ".delta-up {"
                "    color: #10b981;"
                "    font-weight: bold;"
                "}"
                ".delta-down {"
                "    color: #ef4444;"
                "    font-weight: bold;"
                "}"
                ".delta-flat {"
                "    color: #94a3b8;"
                "}"
                "</style>"
            )
            st.markdown(css_style, unsafe_allow_html=True)
            
            rank_rows = []
            for idx, row in top10_today.iterrows():
                delta_str = row['delta']
                if '▲' in delta_str:
                    delta_html = f"<span class='delta-up'>{delta_str}</span>"
                elif '▼' in delta_str:
                    delta_html = f"<span class='delta-down'>{delta_str}</span>"
                elif 'NEW' in delta_str:
                    delta_html = f"<span class='delta-new'>NEW</span>"
                else:
                    delta_html = f"<span class='delta-flat'>-</span>"
                    
                rank_rows.append(
                    f"<tr>"
                    f"<td class='rank-num'>{row['rank']}위</td>"
                    f"<td class='prod-name'>{row['product_name']}</td>"
                    f"<td class='qty-val'>{row['quantity']:,} 개</td>"
                    f"<td style='text-align: right;'>{delta_html}</td>"
                    f"</tr>"
                )
            
            rank_rows_html = "".join(rank_rows)
            
            rank_table_html = (
                f"<table class='rank-table'>"
                f"<thead>"
                f"<tr>"
                f"<th style='width: 15%;'>순위</th>"
                f"<th style='width: 55%;'>상품명</th>"
                f"<th style='width: 15%;'>판매량</th>"
                f"<th style='width: 15%; text-align: right;'>변동 (전주 대비)</th>"
                f"</tr>"
                f"</thead>"
                f"<tbody>"
                f"{rank_rows_html}"
                f"</tbody>"
                f"</table>"
            )
            st.markdown(rank_table_html, unsafe_allow_html=True)
        else:
            st.info("선택한 날짜에 판매 실적이 존재하지 않아 순위를 집계할 수 없습니다.")
            
    st.markdown("---")
    
    # 오프라인 매장별 매출 분석 (조회 기준월 기준 그룹 막대 차트 개편)
    st.markdown("#### 🏪 오프라인 매장별 매출 비교 분석 (조회 기준월 기준)")
    st.caption("선택한 조회 기준월의 오프라인 매출액 기준으로 상위/하위 10개 매장을 내림차순 정렬하고, 전월(M-1) 실적과 직관적으로 비교 대조합니다.")
    
    # 하단 전용 조회 기준월 독립 선택기
    col_store_date, col_store_title = st.columns([1.5, 2.5])
    with col_store_date:
        unique_ym_store = sales_df[['year', 'month', 'year_month_str']].drop_duplicates()
        unique_ym_store['sort_val'] = unique_ym_store['year'] * 100 + unique_ym_store['month']
        unique_ym_store = unique_ym_store.sort_values(by='sort_val', ascending=False)
        ym_store_options = unique_ym_store['year_month_str'].tolist()
        
        # 기본값: 상단 선택일의 년월 매칭
        t_year = target_date.year
        t_month = target_date.month
        t_ym_str = f"{t_year}년 {t_month:02d}월"
        
        if t_ym_str in ym_store_options:
            default_store_ym_idx = ym_store_options.index(t_ym_str)
        else:
            default_store_ym_idx = 0
            
        store_target_ym_str = st.selectbox(
            "📅 오프라인 조회 기준월 선택", 
            options=ym_store_options, 
            index=default_store_ym_idx, 
            key="store_monthly_target_ym"
        )
        
    store_target_year = int(store_target_ym_str.split("년 ")[0])
    store_target_month = int(store_target_ym_str.split("년 ")[1].split("월")[0])
    
    if store_target_month == 1:
        store_prev_year = store_target_year - 1
        store_prev_month = 12
    else:
        store_prev_year = store_target_year
        store_prev_month = store_target_month - 1
        
    store_prev_ym_str = f"{store_prev_year}년 {store_prev_month:02d}월"
    
    with col_store_title:
        st.markdown(
            f"<div style='text-align: right; padding-top: 15px; color: #4b5563; font-weight: 600;'>"
            f"기준월: {store_target_ym_str} / 전월 대비 비교 대상월: {store_prev_ym_str}</div>", 
            unsafe_allow_html=True
        )
        
    offline_sales = sales_df[sales_df['channel_type'] == '오프라인'].copy()
    
    if not offline_sales.empty:
        # 당월 매장 매출 집계
        curr_month_sales = offline_sales[
            (offline_sales['year'] == store_target_year) & 
            (offline_sales['month'] == store_target_month)
        ].groupby('store_name')['amount'].sum().reset_index()
        curr_month_sales.rename(columns={'amount': '당월 매출'}, inplace=True)
        
        # 전월 매장 매출 집계
        prev_month_sales = offline_sales[
            (offline_sales['year'] == store_prev_year) & 
            (offline_sales['month'] == store_prev_month)
        ].groupby('store_name')['amount'].sum().reset_index()
        prev_month_sales.rename(columns={'amount': '전월 매출'}, inplace=True)
        
        # 두 데이터를 병합
        store_compare = pd.merge(curr_month_sales, prev_month_sales, on='store_name', how='outer').fillna(0)
        
        # 당월 매출액 기준으로 내림차순 정렬
        store_compare = store_compare.sort_values(by='당월 매출', ascending=False).reset_index(drop=True)
        
        # 상위 10개 매장 및 하위 10개 매장 선별
        top_compare = store_compare.head(10).copy()
        bottom_compare = store_compare.tail(10).copy()
        
        # Melt하여 '구분' (당월 매출 / 전월 매출)과 '매출액' 컬럼 생성
        top_melt = pd.melt(
            top_compare, 
            id_vars=['store_name'], 
            value_vars=['당월 매출', '전월 매출'],
            var_name='기간 구분', 
            value_name='매출액'
        )
        
        bottom_melt = pd.melt(
            bottom_compare, 
            id_vars=['store_name'], 
            value_vars=['당월 매출', '전월 매출'],
            var_name='기간 구분', 
            value_name='매출액'
        )
        
        hist_cols = st.columns(2)
        
        with hist_cols[0]:
            st.markdown(f"##### 📈 상위 10개 오프라인 매장 매출액 대조 ({store_target_ym_str} 기준)")
            if not top_melt.empty:
                top_store_order = top_compare['store_name'].tolist()
                
                fig_top_bar = px.bar(
                    top_melt,
                    x='store_name',
                    y='매출액',
                    color='기간 구분',
                    barmode='group',
                    labels={'매출액': '매출 금액 (₩)', 'store_name': '오프라인 매장명', '기간 구분': '기간'},
                    color_discrete_map={'당월 매출': '#0B132B', '전월 매출': '#94A3B8'}
                )
                
                fig_top_bar.update_traces(
                    hovertemplate='₩%{y:,.0f}<extra></extra>'
                )
                
                fig_top_bar.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=20, r=20, t=10, b=20),
                    xaxis=dict(
                        showgrid=True, 
                        gridcolor='#E5E7EB', 
                        categoryorder='array', 
                        categoryarray=top_store_order,
                        showspikes=True,
                        spikethickness=1,
                        spikedash='dash',
                        spikemode='across',
                        spikecolor='#C5A880'
                    ),
                    yaxis=dict(showgrid=True, gridcolor='#E5E7EB'),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    hovermode='x unified',
                    hoverlabel=dict(
                        bgcolor='#0B132B',
                        bordercolor='#C5A880',
                        font_size=12,
                        font_color='#F8F7F4',
                        font_family='Outfit'
                    )
                )
                st.plotly_chart(fig_top_bar, use_container_width=True, key="monthly_top10_store_bar")
            else:
                st.info("상위 매장 실적 데이터가 존재하지 않습니다.")
                
        with hist_cols[1]:
            st.markdown(f"##### 📉 하위 10개 오프라인 매장 매출액 대조 ({store_target_ym_str} 기준)")
            if not bottom_melt.empty:
                bottom_store_order = bottom_compare['store_name'].tolist()
                
                fig_bottom_bar = px.bar(
                    bottom_melt,
                    x='store_name',
                    y='매출액',
                    color='기간 구분',
                    barmode='group',
                    labels={'매출액': '매출 금액 (₩)', 'store_name': '오프라인 매장명', '기간 구분': '기간'},
                    color_discrete_map={'당월 매출': '#C5A880', '전월 매출': '#94A3B8'}
                )
                
                fig_bottom_bar.update_traces(
                    hovertemplate='₩%{y:,.0f}<extra></extra>'
                )
                
                fig_bottom_bar.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=20, r=20, t=10, b=20),
                    xaxis=dict(
                        showgrid=True, 
                        gridcolor='#E5E7EB', 
                        categoryorder='array', 
                        categoryarray=bottom_store_order,
                        showspikes=True,
                        spikethickness=1,
                        spikedash='dash',
                        spikemode='across',
                        spikecolor='#C5A880'
                    ),
                    yaxis=dict(showgrid=True, gridcolor='#E5E7EB'),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    hovermode='x unified',
                    hoverlabel=dict(
                        bgcolor='#0B132B',
                        bordercolor='#C5A880',
                        font_size=12,
                        font_color='#F8F7F4',
                        font_family='Outfit'
                    )
                )
                st.plotly_chart(fig_bottom_bar, use_container_width=True, key="monthly_bottom10_store_bar")
            else:
                st.info("하위 매장 실적 데이터가 존재하지 않습니다.")
    else:
        st.warning("조건에 맞는 오프라인 매출 데이터가 없어 분석을 수행할 수 없습니다.")


# --- TAB 2: 다개년 채널별 분석 ---
with tab2:
    st.markdown("### 📊 다개년 채널 및 트렌드 입체 분석")
    st.caption("2025년과 2026년의 월별 실적을 온라인과 오프라인 채널로 나누어 정밀하게 대조 비교합니다.")
    
    # 25년 vs 26년 월별 매출 채널 집계 (매출 데이터 filtered_sales 기반)
    monthly_sales_data = filtered_sales.groupby(['year', 'month', 'channel_type']).agg({
        'amount': 'sum'
    }).reset_index()
    
    # 월 정렬을 위해 01월, 02월 등 문자열 생성
    monthly_sales_data = monthly_sales_data.dropna(subset=['year', 'month'])
    monthly_sales_data['year'] = monthly_sales_data['year'].astype(int)
    monthly_sales_data['month'] = monthly_sales_data['month'].astype(int)
    monthly_sales_data['month_str'] = monthly_sales_data['month'].apply(lambda x: f"{x:02d}월")
    monthly_sales_data = monthly_sales_data.sort_values(by=['year', 'month'])
    
    compare_chart_cols = st.columns(1)
    
    with compare_chart_cols[0]:
        st.markdown("#### 💳 2025년 vs 2026년 월별 온라인/오프라인 매출 비교")
        if not monthly_sales_data.empty:
            fig_compare_sales = px.bar(
                monthly_sales_data,
                x='month_str',
                y='amount',
                color='channel_type',
                facet_col='year',
                barmode='group',
                labels={'amount': '매출 금액 (₩)', 'month_str': '월', 'channel_type': '채널 구분'},
                color_discrete_map={'온라인': '#0B132B', '오프라인': '#C5A880'}
            )
            
            fig_compare_sales.update_traces(
                hovertemplate='₩%{y:,.0f}<extra></extra>'
            )
            
            fig_compare_sales.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                margin=dict(l=20, r=20, t=30, b=20),
                yaxis=dict(showgrid=True, gridcolor='#E5E7EB'),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                hovermode='x unified',
                hoverlabel=dict(
                    bgcolor='#0B132B',
                    bordercolor='#C5A880',
                    font_size=12,
                    font_color='#F8F7F4',
                    font_family='Outfit'
                )
            )
            
            # 서브플롯 X축(Spikeline 포함) 설정
            fig_compare_sales.update_xaxes(
                showgrid=True,
                gridcolor='#E5E7EB',
                showspikes=True,
                spikethickness=1,
                spikedash='dash',
                spikemode='across',
                spikecolor='#C5A880'
            )
            
            # 서브플롯 타이틀 정제
            fig_compare_sales.for_each_annotation(lambda a: a.update(text=f"📊 {a.text.split('=')[-1]}년 실적"))
            st.plotly_chart(fig_compare_sales, use_container_width=True, key="compare_sales_plotly")
        else:
            st.warning("비교 대상 매출 데이터가 없습니다.")
            
    # 기존 매장 주차별 매출 분석 제거됨 (Tab 1 월간 성과 관리 센터 하단으로 고도화 이식 완료)

# --- TAB 3: 생산 및 BOM 예측 ---
with tab3:
    st.markdown("### ⚙️ 다니엘트루스 실시간 제조 및 BOM 통제 센터")
    st.caption("생산 계획에 따른 품목별 BOM 계층 구조와 실시간 자재 소요량(MRP)을 직관적으로 확인하고 분석합니다.")
    
    # 1. 자재 카테고리 동적 분류 함수 정의
    def classify_material_category(child_code, child_name):
        child_code = str(child_code).upper()
        child_name = str(child_name)
        if "FRAG" in child_code or "향료" in child_name:
            return "반제품(조향 오일)"
        elif "BASE" in child_code or "WAX" in child_code or "왁스" in child_name or "에탄올" in child_name:
            return "반제품(베이스 원료)"
        elif "GLASS" in child_code or "BOTTLE" in child_code or "유리" in child_name or "병" in child_name or "공병" in child_name:
            return "원자재(용기/부자재)"
        elif "BOX" in child_code or "상자" in child_name or "패키지" in child_name:
            return "원자재(포장재)"
        else:
            return "기타 자재"

    # 2. 가상 보유 재고 마스터 맵 정의 (MRP 비교용)
    current_stock_map = {
        'RAW_DF_BASE': 180000.0,    # 디퓨저 베이스 에탄올 (ml)
        'RAW_DF_FRAG': 45000.0,     # 디퓨저 조합 향료 오일 (ml) - 소요량에 따라 부족 가능
        'RAW_DF_GLASS': 450.0,      # 프리미엄 디퓨저 유리 용기 (개) - 부족 경고 타겟
        'RAW_OL_FRAG': 18000.0,     # 천연 에센셜 향료 원액 (ml)
        'RAW_OL_BOTTLE': 2200.0,    # 롤온 고급 초자 유리병 (개)
        'RAW_CD_WAX': 150000.0,     # 천연 골든 소이 왁스 (g) - 부족 경고 타겟
        'RAW_CD_FRAG': 22000.0,     # 캔들 가열용 향료 오일 (ml)
        'RAW_CD_GLASS': 850.0,      # 내열성 캔들 유리 용기 (개) - 부족 경고 타겟
        'RAW_AC_BOX': 1200.0,       # 수작업 패키징 상자 (개)
        'RAW_PR_MINI': 4500.0       # 미니어처 전용 샘플 공병 (개)
    }

    # 3. 실시간 생산 계획 수량 입력 및 조정 인터랙티브 그리드
    st.markdown("#### 📝 실시간 확정 생산 계획 조정 (MRP 통제판)")
    st.caption("아래 완제품 목록의 **'생산 예정 수량 (개)'** 열을 직접 변경하면, 하단의 생키 흐름도, 트리맵, MRP 소요량 분석 결과가 실시간으로 재연산됩니다.")
    
    # 2025년 5월 기본 생산 계획 데이터 가공
    if 'mrp_plan_df' not in st.session_state:
        st.session_state.mrp_plan_df = plan_df[plan_df['plan_month'] == '2025-05'][['product_code', 'product_name', 'planned_qty']].copy()
        
    # st.data_editor를 통한 프리미엄 인터랙티브 수량 입력기 구현
    edited_plan_df = st.data_editor(
        st.session_state.mrp_plan_df,
        column_config={
            "product_code": st.column_config.TextColumn("완제품 코드", disabled=True),
            "product_name": st.column_config.TextColumn("완제품명", disabled=True, width="medium"),
            "planned_qty": st.column_config.NumberColumn("생산 예정 수량 (개)", min_value=0, max_value=50000, step=10, format="%d개")
        },
        hide_index=True,
        use_container_width=True,
        key="plan_data_editor"
    )
    # 편집 수량 상태 동기화
    st.session_state.mrp_plan_df = edited_plan_df
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 4. 품목별 BOM 구조 계층형 트리맵 (Treemap) 단독 렌더링
    st.markdown("#### 🌳 품목별 BOM 구조 계층형 트리맵 (Treemap)")
    st.caption("완제품을 드롭다운으로 선택하면 하위 원자재 카테고리 및 자품목 단량 구성비가 면적으로 표시됩니다.")
    
    # 완제품 트리맵 선택용 드롭다운
    treemap_options = list(edited_plan_df['product_name'].unique())
    selected_tree_prod = st.selectbox("BOM 구조 조회 품목 선택", treemap_options, key="mrp_treemap_selector")
    
    selected_p_code = edited_plan_df[edited_plan_df['product_name'] == selected_tree_prod]['product_code'].values[0]
    p_bom_df = bom_df[bom_df['parent_code'] == selected_p_code].copy()
    
    if not p_bom_df.empty:
        # 가상 충진 품목 안내 배너 노출
        if str(p_bom_df['bom_id'].iloc[0]).startswith('BOM_FILL_'):
            st.info("💡 본 품목은 구글 시트에 원본 BOM 정보가 등록되지 않아, 대시보드 사용성 향상을 위해 표준 가상 자재 규격(부자재 리드스틱/패키징)으로 자동 대체 연동되었습니다.")
            
        p_bom_df['category'] = p_bom_df.apply(lambda r: classify_material_category(r['child_code'], r['child_name']), axis=1)
        
        fig_treemap = px.treemap(
            p_bom_df,
            path=['category', 'child_name'],
            values='unit_qty',
            color='category',
            color_discrete_map={
                "반제품(조향 오일)": "#0B132B",
                "반제품(베이스 원료)": "#1E3A8A",
                "원자재(용기/부자재)": "#C5A880",
                "원자재(포장재)": "#D4AF37",
                "기타 자재": "#64748B"
            },
            labels={'unit_qty': '소요단량'},
            title=f"⚜️ {selected_tree_prod[:35]}... BOM 단량 비율"
        )
        
        fig_treemap.update_traces(
            textinfo="label+value",
            hovertemplate='<b>%{label}</b><br>카테고리: %{parent}<br>단품 소요단량: %{value:,.1f}<extra></extra>'
        )
        
        fig_treemap.update_layout(
            margin=dict(l=5, r=5, t=30, b=5),
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            title_font=dict(size=13, family="Cinzel", color="#0B132B"),
            hoverlabel=dict(
                bgcolor='#0B132B',
                bordercolor='#C5A880',
                font_size=11,
                font_color='#F8F7F4',
                font_family='Outfit'
            )
        )
        st.plotly_chart(fig_treemap, use_container_width=True, key="mrp_treemap_chart")
    else:
        st.warning("⚠️ 구글 시트에 이 완제품에 대한 BOM 정보가 등록되지 않았습니다.")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # 5. 실시간 자재 소요량(MRP) 계산 및 재고 부족 테이블
    st.markdown("#### ⚠️ 실시간 자재 소요량(MRP) 및 재고 부족 알림판")
    st.caption("확정 생산 계획 수량과 BOM 소요단량을 조합하여 총 필요 자재량을 산정하고, 보유 재고와 실시간 대조하여 부족 자재를 감지합니다.")
    
    # MRP 계산 프로세스
    mrp_joined = pd.merge(
        bom_df,
        edited_plan_df[['product_code', 'planned_qty']],
        left_on='parent_code',
        right_on='product_code',
        how='inner'
    )
    mrp_joined['required_qty'] = mrp_joined['planned_qty'] * mrp_joined['unit_qty']
    
    # 자품목별 총 소요량 집계
    mrp_agg = mrp_joined.groupby(['child_code', 'child_name']).agg({
        'required_qty': 'sum'
    }).reset_index()
    
    # 속성 가공 및 대조
    mrp_agg['category'] = mrp_agg.apply(lambda r: classify_material_category(r['child_code'], r['child_name']), axis=1)
    mrp_agg['current_stock'] = mrp_agg['child_code'].map(current_stock_map).fillna(0.0)
    mrp_agg['shortage_qty'] = mrp_agg.apply(lambda r: max(0.0, r['required_qty'] - r['current_stock']), axis=1)
    mrp_agg['status'] = mrp_agg.apply(lambda r: "⚠️ 재고 부족" if r['shortage_qty'] > 0 else "✅ 정상", axis=1)
    
    # 출력용 컬럼 재배치 및 헤더 정의
    mrp_display = mrp_agg[[
        'category', 'child_code', 'child_name', 'required_qty', 'current_stock', 'shortage_qty', 'status'
    ]].rename(columns={
        'category': '자재 카테고리',
        'child_code': '자재 코드',
        'child_name': '자재명',
        'required_qty': '총 소요량',
        'current_stock': '현재 보유 재고',
        'shortage_qty': '부족 수량',
        'status': '재고 상태'
    })
    
    # Pandas DataFrame Styler를 통한 재고 부족 경고 하이라이트 함수 정의
    def style_mrp_table(df):
        def highlight_rows(row):
            if row['재고 상태'] == "⚠️ 재고 부족":
                # 부드럽고 가독성 높은 연한 빨간색 파스텔톤 배경색 적용
                return ['background-color: #FEE2E2; color: #991B1B; font-weight: bold;'] * len(row)
            return [''] * len(row)
            
        # 원본 명칭 손상 없이 부족한 자재명 앞에 경고 아이콘 추가
        styled_df = df.copy()
        styled_df['자재명'] = styled_df.apply(
            lambda r: f"⚠️ {r['자재명']}" if r['재고 상태'] == "⚠️ 재고 부족" else r['자재명'], axis=1
        )
        
        return styled_df.style.apply(highlight_rows, axis=1).format({
            '총 소요량': '{:,.1f}',
            '현재 보유 재고': '{:,.1f}',
            '부족 수량': '{:,.1f}'
        })
        
    st.dataframe(
        style_mrp_table(mrp_display),
        use_container_width=True,
        hide_index=True,
        key="mrp_styled_dataframe"
    )

    # 6. 생산 및 BOM 소요량 트리 그리드 현황판 (100% Full Width 및 필터 기반 동적 전개)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 🌲 완제품별 생산 및 BOM 소요량 트리 그리드")
    st.caption("🔍 원하시는 완제품 품목을 다중 선택하거나 직접 검색하여 하위 자품목(BOM)의 실시간 소요량을 전개하고 조회할 수 있습니다.")
    
    # 실시간 편집된 생산 계획 데이터를 기준으로 품목 목록 구성
    active_plan = edited_plan_df.copy()
    
    if not active_plan.empty:
        # 계획량 내림차순 정렬하여 옵션 표시
        sorted_plan = active_plan.sort_values(by='planned_qty', ascending=False)
        all_options = sorted_plan['product_name'].tolist()
        
        # 기본 선택값: 계획량이 0보다 큰 상위 3개 완제품
        default_options = sorted_plan[sorted_plan['planned_qty'] > 0].head(3)['product_name'].tolist()
        if not default_options and all_options:
            default_options = all_options[:3]
            
        selected_products = st.multiselect(
            "🔍 조회할 완제품 품목을 선택해 주세요 (다중 선택 및 검색 가능)",
            options=all_options,
            default=default_options,
            key="mrp_tree_grid_multiselect"
        )
        
        if selected_products:
            st.markdown("<br>", unsafe_allow_html=True)
            for prod_name in selected_products:
                row = active_plan[active_plan['product_name'] == prod_name].iloc[0]
                prod_code = row['product_code']
                planned_qty = row['planned_qty']
                
                # 개별 완제품 Expander 카드로 Full Width 구성
                with st.expander(f"📦 {prod_name} (실시간 계획량: {planned_qty:,.0f} 개)", expanded=True):
                    # 해당 완제품의 자품목 BOM 매핑
                    child_bom = bom_df[bom_df['parent_code'] == prod_code].copy()
                    
                    if not child_bom.empty:
                        child_bom['final_required_qty'] = planned_qty * child_bom['unit_qty']
                        
                        # 시각적 통일성을 위해 자재 카테고리 정보도 동적으로 추가 표시
                        child_bom['category'] = child_bom.apply(
                            lambda r: classify_material_category(r['child_code'], r['child_name']), axis=1
                        )
                        
                        display_bom = child_bom[['category', 'child_name', 'unit_qty', 'final_required_qty']].rename(columns={
                            'category': '자재 카테고리',
                            'child_name': '자품목(원자재) 명칭',
                            'unit_qty': '소요단량 (Unit Qty)',
                            'final_required_qty': '실시간 필요 수량'
                        })
                        
                        st.dataframe(
                            display_bom.style.format({
                                '소요단량 (Unit Qty)': '{:,.1f}',
                                '실시간 필요 수량': '{:,.0f}'
                            }),
                            use_container_width=True,
                            hide_index=True,
                            key=f"grid_{prod_code}"
                        )
                    else:
                        st.info("구글 시트에 이 완제품의 BOM 정보가 누락되어 있습니다.")
        else:
            st.info("조회할 완제품 품목을 필터에서 1개 이상 선택해 주세요.")
    else:
        st.warning("등록된 완제품 계획 데이터가 존재하지 않습니다.")

# --- TAB 4: 소진 기한 및 트렌드 분석 ---
with tab4:
    col_title, col_refresh = st.columns([3, 1])
    with col_title:
        st.markdown("### 📈 완제품 소진 기한 및 트렌드 분석 (SCM Action-Oriented)")
        st.caption("🔍 전사 완제품 81개 뼈대와 실시간 구글 시트를 연동하여, 과거 90일 판매 실적 속도와 마케팅 예측치의 괴리(Gap)를 입체적으로 비교/예측합니다.")
    with col_refresh:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 실시간 구글 시트 캐시 강제 새로고침", key="btn_clear_cache_sojin", use_container_width=True):
            st.cache_data.clear()
            st.success("🔄 캐시가 성공적으로 초기화되었습니다! 최신 데이터를 다시 읽어옵니다.")
            st.rerun()
            
    # 데이터 로드 (quantity_df 아규먼트를 안전하게 바인딩하여 81개 본품 뼈대 복원)
    df_stock_sojin, df_sales_sojin = load_google_sojin_data_v6(quantity_df)
    
    if not df_stock_sojin.empty:
        sojin_df = calculate_sojin_metrics(df_stock_sojin, df_sales_sojin)
        if not sojin_df.empty:
            # SCM 정렬 및 괴리 알림용 절대값 컬럼 선제 선언 (KeyError 방지)
            sojin_df['Gap_Abs'] = sojin_df['예측괴리도_Gap'].abs()
            
            # ----------------------------------------------------
            # 🚀 SCM 격리 분기 (Data Cleansing & Partitioning)
            # ----------------------------------------------------
            # 분류 [1]: 메인 분석 대상 (최근 90일간 누적 판매량이 0보다 큰 활성 품목)
            active_sojin_master = sojin_df[sojin_df['총판매수량_90일'] > 0].copy()
            
            # 분류 [2]: 장기 미판매 악성 재고 (90일 판매량이 0이며, 현재 재고가 남아있는 품목)
            dead_sojin_master = sojin_df[(sojin_df['총판매수량_90일'] == 0) & (sojin_df['현재재고'] > 0)].copy()
            
            # 🎯 세션 상태 기반 요약 필터 변수 사전 정의 🎯
            if 'sojin_kpi_filter' not in st.session_state:
                st.session_state.sojin_kpi_filter = "전체"
                
            # 1. 상단 KPI 요약 카드 배치 (노이즈 방지를 위해 active_sojin_master 기준 연산)
            risk_cnt = len(active_sojin_master[active_sojin_master['가중소진기한_개월'] < 1.5])
            warn_cnt = len(active_sojin_master[(active_sojin_master['가중소진기한_개월'] >= 1.5) & (active_sojin_master['가중소진기한_개월'] < 3.0)])
            safe_cnt = len(active_sojin_master[active_sojin_master['가중소진기한_개월'] >= 3.0])
            avg_gap = active_sojin_master['예측괴리도_Gap'].mean() if not active_sojin_master.empty else 0.0
            
            # 하이라이트 동적 CSS 클래스 할당 (골드/네온 하이라이팅 테두리 적용)
            style_risk = "border: 2.5px solid rgba(255, 75, 75, 0.85); box-shadow: 0 0 15px rgba(255, 75, 75, 0.35); transform: translateY(-2px);" if st.session_state.sojin_kpi_filter == "위험" else "border: 1px solid rgba(255,255,255,0.05);"
            style_warn = "border: 2.5px solid rgba(255, 165, 0, 0.85); box-shadow: 0 0 15px rgba(255, 165, 0, 0.35); transform: translateY(-2px);" if st.session_state.sojin_kpi_filter == "주의" else "border: 1px solid rgba(255,255,255,0.05);"
            style_safe = "border: 2.5px solid rgba(0, 200, 81, 0.85); box-shadow: 0 0 15px rgba(0, 200, 81, 0.35); transform: translateY(-2px);" if st.session_state.sojin_kpi_filter == "안정" else "border: 1px solid rgba(255,255,255,0.05);"
            style_gap = "border: 1px solid rgba(255,255,255,0.05);"
            
            kpi_cols = st.columns(4)
            with kpi_cols[0]:
                st.markdown(
                    f"<div class='kpi-card' style='background-color: #1A1A2E !important; color: #FFFFFF !important; border-left: 5px solid #FF4B4B; {style_risk}'>"
                    f"<div class='kpi-label'>🔴 품절 임박 위험품목 (1.5개월 미만)</div>"
                    f"<div class='kpi-value' style='color:#FF4B4B;'>{risk_cnt} <span style='font-size:18px;'>개</span></div>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                # 🔴 위험 품목 리스트 expander 동적 연동 (클릭 시 세부 품목 노출 입체적 UX)
                with st.expander("🔍 위험품목 목록 보기", expanded=(st.session_state.sojin_kpi_filter == "위험")):
                    risk_items = active_sojin_master[active_sojin_master['가중소진기한_개월'] < 1.5]
                    if not risk_items.empty:
                        for idx, row in risk_items.iterrows():
                            st.markdown(
                                f"<div style='font-size: 11.5px; border-bottom: 1px solid rgba(255,255,255,0.05); padding: 5px 0; color: #FF4B4B; font-weight: 500;'>"
                                f"• {row['품명']}<br>"
                                f"<span style='color: #A0A0A5; font-size: 10.5px;'>재고: {row['현재재고']:,.0f}개 | 소진일: {row['예상소진기한_날짜']}</span>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("해당 품목이 없습니다.")
                        
            with kpi_cols[1]:
                st.markdown(
                    f"<div class='kpi-card' style='background-color: #1A1A2E !important; color: #FFFFFF !important; border-left: 5px solid #FFA500; {style_warn}'>"
                    f"<div class='kpi-label'>🟡 모니터링 주의품목 (1.5~3개월)</div>"
                    f"<div class='kpi-value' style='color:#FFA500;'>{warn_cnt} <span style='font-size:18px;'>개</span></div>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                # 🟡 주의 품목 리스트 expander 동적 연동
                with st.expander("🔍 주의품목 목록 보기", expanded=(st.session_state.sojin_kpi_filter == "주의")):
                    warn_items = active_sojin_master[(active_sojin_master['가중소진기한_개월'] >= 1.5) & (active_sojin_master['가중소진기한_개월'] < 3.0)]
                    if not warn_items.empty:
                        for idx, row in warn_items.iterrows():
                            st.markdown(
                                f"<div style='font-size: 11.5px; border-bottom: 1px solid rgba(255,255,255,0.05); padding: 5px 0; color: #FFA500; font-weight: 500;'>"
                                f"• {row['품명']}<br>"
                                f"<span style='color: #A0A0A5; font-size: 10.5px;'>재고: {row['현재재고']:,.0f}개 | 소진일: {row['예상소진기한_날짜']}</span>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("해당 품목이 없습니다.")
                        
            with kpi_cols[2]:
                st.markdown(
                    f"<div class='kpi-card' style='background-color: #1A1A2E !important; color: #FFFFFF !important; border-left: 5px solid #00C851; {style_safe}'>"
                    f"<div class='kpi-label'>🟢 공급 안정 품목 (3개월 이상)</div>"
                    f"<div class='kpi-value' style='color:#00C851;'>{safe_cnt} <span style='font-size:18px;'>개</span></div>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                # 🟢 안정 품목 리스트 expander 동적 연동
                with st.expander("🔍 안정품목 목록 보기", expanded=(st.session_state.sojin_kpi_filter == "안정")):
                    safe_items = active_sojin_master[active_sojin_master['가중소진기한_개월'] >= 3.0]
                    if not safe_items.empty:
                        for idx, row in safe_items.iterrows():
                            st.markdown(
                                f"<div style='font-size: 11.5px; border-bottom: 1px solid rgba(255,255,255,0.05); padding: 5px 0; color: #00C851; font-weight: 500;'>"
                                f"• {row['품명']}<br>"
                                f"<span style='color: #A0A0A5; font-size: 10.5px;'>재고: {row['현재재고']:,.0f}개 | 소진일: {row['예상소진기한_날짜']}</span>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("해당 품목이 없습니다.")
                        
            with kpi_cols[3]:
                st.markdown(
                    f"<div class='kpi-card' style='background-color: #1A1A2E !important; color: #FFFFFF !important; border-left: 5px solid #C5A880; {style_gap}'>"
                    f"<div class='kpi-label'>⚖️ 기획 vs 실제 평균 예측 괴리</div>"
                    f"<div class='kpi-value' style='color:#C5A880;'>{avg_gap:.1f} <span style='font-size:18px;'>개월</span></div>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                with st.expander("⚖️ 괴리 분석 안내", expanded=False):
                    st.caption("마케팅 예상 판매량 기반의 보존 기간과, 과거 90일 가중 판매 속도 기반 보존 기간의 평균 오차 지표입니다. 오차가 작을수록 정확한 예측을 뜻합니다.")
                    
            st.markdown("<br>", unsafe_allow_html=True)
            
            # ----------------------------------------------------
            # 🧪 분석 분기 서브 탭 신설 (Bento Sub-Tabs Selector)
            # ----------------------------------------------------
            tab_active, tab_dead = st.tabs(["📊 3대 SCM 액션 센터 (활성 품목)", "💀 악성 재고 모니터 (Dead Stock)"])
            
            # ====================================================
            # 서브 탭 [1]: 3대 SCM 액션 센터 (활성 품목)
            # ====================================================
            with tab_active:
                st.markdown("<div style='font-size: 13.5px; font-weight: 600; color: #C5A880; margin-bottom: 8px;'>🎯 리스크 분류별 세부 품목 퀵 필터링 (클릭 시 아래 차트/그리드 연계 필터링)</div>", unsafe_allow_html=True)
                filter_cols = st.columns(4)
                with filter_cols[0]:
                    if st.button(f"🔴 위험품목 리스트만 보기 ({risk_cnt}개)", use_container_width=True, key="btn_filter_risk"):
                        st.session_state.sojin_kpi_filter = "위험"
                with filter_cols[1]:
                    if st.button(f"🟡 주의품목 리스트만 보기 ({warn_cnt}개)", use_container_width=True, key="btn_filter_warn"):
                        st.session_state.sojin_kpi_filter = "주의"
                with filter_cols[2]:
                    if st.button(f"🟢 안정품목 리스트만 보기 ({safe_cnt}개)", use_container_width=True, key="btn_filter_safe"):
                        st.session_state.sojin_kpi_filter = "안정"
                with filter_cols[3]:
                    if st.button(f"🔄 전체 완제품 리스트 복구 ({len(active_sojin_master)}개)", use_container_width=True, key="btn_filter_all"):
                        st.session_state.sojin_kpi_filter = "전체"
                        
                if st.session_state.sojin_kpi_filter != "전체":
                    st.info(f"💡 현재 **[{st.session_state.sojin_kpi_filter} 품목]** 기준 필터링 모드가 활성화되어 대시보드가 슬라이싱되었습니다.")
                    
                st.markdown("<br>", unsafe_allow_html=True)
                
                # 퀵 필터링 데이터셋 세그먼트 생성
                filtered_sojin = active_sojin_master.copy()
                if st.session_state.sojin_kpi_filter == "위험":
                    filtered_sojin = filtered_sojin[filtered_sojin['가중소진기한_개월'] < 1.5]
                elif st.session_state.sojin_kpi_filter == "주의":
                    filtered_sojin = filtered_sojin[(filtered_sojin['가중소진기한_개월'] >= 1.5) & (filtered_sojin['가중소진기한_개월'] < 3.0)]
                elif st.session_state.sojin_kpi_filter == "안정":
                    filtered_sojin = filtered_sojin[filtered_sojin['가중소진기한_개월'] >= 3.0]
                
                st.markdown("### 🏆 SCM Decision-Making Action Center")
                
                sec1_cols = st.columns([1.6, 1.4])
                
                # 차트 1: 의사결정용 재고 위험도 4분면 매트릭스
                with sec1_cols[0]:
                    st.markdown("#### 🎯 [의사결정용] 재고 위험도 4분면 매트릭스 (Quadrant Map)")
                    st.caption("과거 90일 실제 실적 가중 소진기한(Y축)과 마케팅 기획 예상 소진기한(X축)을 45일 안전재고선을 경계로 4대 SCM 지대 영역에 동적 매핑하여 비즈니스 불확실성을 최소화합니다.")
                    
                    if not filtered_sojin.empty:
                        import plotly.graph_objects as go
                        fig_quad = go.Figure()

                        # 4대 SCM 지대 배경 영역 셰이프 추가
                        # 🚨 초비상 (즉시 생산/출고): X < 1.5, Y < 1.5 (매우 투명한 붉은색)
                        fig_quad.add_shape(
                            type="rect", x0=0, x1=1.5, y0=0, y1=1.5,
                            fillcolor="rgba(255, 75, 75, 0.08)", line=dict(width=0),
                            layer="below"
                        )
                        # 🟡 품절 위험 (긴급 발주): X >= 1.5, Y < 1.5 (매우 투명한 주황색)
                        fig_quad.add_shape(
                            type="rect", x0=1.5, x1=24.0, y0=0, y1=1.5,
                            fillcolor="rgba(255, 165, 0, 0.06)", line=dict(width=0),
                            layer="below"
                        )
                        # 📦 재고 과다 (프로모션): X < 1.5, Y >= 1.5 (매우 투명한 파란색)
                        fig_quad.add_shape(
                            type="rect", x0=0, x1=1.5, y0=1.5, y1=24.0,
                            fillcolor="rgba(0, 123, 255, 0.06)", line=dict(width=0),
                            layer="below"
                        )
                        # 🟢 공급 안정 (양호 지대): X >= 1.5, Y >= 1.5 (매우 투명한 녹색)
                        fig_quad.add_shape(
                            type="rect", x0=1.5, x1=24.0, y0=1.5, y1=24.0,
                            fillcolor="rgba(0, 200, 81, 0.06)", line=dict(width=0),
                            layer="below"
                        )

                        # 45일(1.5개월) 안전재고 기준 십자선 가이드 라인 추가
                        fig_quad.add_shape(
                            type="line", x0=1.5, x1=1.5, y0=0, y1=24.0,
                            line=dict(color="rgba(255, 255, 255, 0.4)", width=2, dash="dash")
                        )
                        fig_quad.add_shape(
                            type="line", x0=0, x1=24.0, y0=1.5, y1=1.5,
                            line=dict(color="rgba(255, 255, 255, 0.4)", width=2, dash="dash")
                        )

                        # 각 지대 텍스트 라벨링 추가
                        fig_quad.add_annotation(
                            x=0.75, y=0.75, text="🚨 초비상<br>(즉시 생산/출고)", showarrow=False,
                            font=dict(size=12, color="#FF4B4B", weight="bold"), align="center"
                        )
                        fig_quad.add_annotation(
                            x=12.75, y=0.75, text="🟡 품절 위험<br>(긴급 발주)", showarrow=False,
                            font=dict(size=12, color="#FFA500", weight="bold"), align="center"
                        )
                        fig_quad.add_annotation(
                            x=0.75, y=12.75, text="📦 재고 과다<br>(프로모션/소진 유도)", showarrow=False,
                            font=dict(size=12, color="#007BFF", weight="bold"), align="center"
                        )
                        fig_quad.add_annotation(
                            x=12.75, y=12.75, text="🟢 공급 안정<br>(양호 지대)", showarrow=False,
                            font=dict(size=12, color="#00C851", weight="bold"), align="center"
                        )

                        # 버블 차트 호버 텍스트 및 데이터셋 결합
                        hover_text = []
                        for idx, r in filtered_sojin.iterrows():
                            txt = (
                                f"<b>{r['품명']}</b><br>"
                                f"• 대분류: {r['카테고리']}<br>"
                                f"• 현재 재고: {r['현재재고']:,.0f} 개<br>"
                                f"• 기획 소진: {r['예상소진기한_개월']:.1f}개월 ({r['예상소진기한_날짜']})<br>"
                                f"• 실제 가중소진: {r['가중소진기한_개월']:.1f}개월 ({r['가중소진기한_날짜']})<br>"
                                f"• 예측 괴리도: {r['예측괴리도_Gap']:.1f}개월"
                            )
                            hover_text.append(txt)

                        # 네온 컬러 맵 구성
                        colors_bubble = []
                        for term in filtered_sojin['가중소진기한_개월']:
                            if term < 1.5:
                                colors_bubble.append('#FF4B4B')  # 🔴 위험 품절 임박
                            elif term < 3.0:
                                colors_bubble.append('#FFA500')  # 🟡 주의 관찰
                            else:
                                colors_bubble.append('#00C851')  # 🟢 안정 양호

                        fig_quad.add_trace(go.Scatter(
                            x=filtered_sojin['예상소진기한_시각화'],
                            y=filtered_sojin['가중소진기한_시각화'],
                            mode='markers',
                            marker=dict(
                                size=filtered_sojin['현재재고_시각화'],
                                color=colors_bubble,
                                line=dict(width=1, color='rgba(255,255,255,0.4)'),
                                sizemode='diameter',
                                sizemin=8
                            ),
                            text=hover_text,
                            hoverinfo='text',
                            showlegend=False
                        ))

                        fig_quad.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(26,26,46,0.3)',
                            font=dict(family="Outfit, Nanum Gothic, sans-serif", color="#F5F5F7"),
                            xaxis=dict(
                                title="기획 예상 소진기한 (개월, 최대 24m 캡핑)",
                                showgrid=True,
                                gridcolor='rgba(255,255,255,0.05)',
                                range=[0, 25]
                            ),
                            yaxis=dict(
                                title="실제 가중 소진기한 (개월, 최대 24m 캡핑)",
                                showgrid=True,
                                gridcolor='rgba(255,255,255,0.05)',
                                range=[0, 25]
                            ),
                            margin=dict(l=50, r=30, t=20, b=50),
                            height=480
                        )
                        st.plotly_chart(fig_quad, use_container_width=True, key="scm_quadrant_matrix_plot")
                    else:
                        st.info("현재 필터링된 조건에 해당하는 매트릭스 시각화 품목이 없습니다.")

                # 차트 2: 실무자용 발주 카운트다운 게이지
                with sec1_cols[1]:
                    st.markdown("#### ⏱️ [실무자용] 발주 카운트다운 게이지 (D-Day Countdown Bar)")
                    st.caption("안전재고(45일) 소진 시점 대비 초과/미달한 실질적 D-Day 남은 일수를 정밀 계산하여 리스크 수준에 따라 동적으로 색상이 네온 변환되는 긴급 오더 바입니다. (상위 최대 20개 품목 표출)")
                    
                    # 안전재고 기준: 45일
                    safety_stock_days = 45.0
                    safety_stock_months = safety_stock_days / 30.4375
                    
                    # 안전재고 미만 조건 필터링을 제거하고, 급한 순서(소진기한 짧은 순)대로 정렬 후 상위 20개 품목 추출
                    df_low = filtered_sojin.sort_values(by='가중소진기한_개월', ascending=True).copy()
                    df_low = df_low.head(20)
                    
                    if not df_low.empty:
                        df_low['D-Day_일수'] = (df_low['가중소진기한_개월'] - safety_stock_months) * 30.4375
                        
                        # 안전재고 등급별 색상 부여 함수 정의
                        def get_neon_color(months):
                            if months < (safety_stock_months / 2.0):
                                return 'rgba(255, 75, 75, 0.95)' # 22.5일 미만 (초긴급: Red)
                            elif months < safety_stock_months:
                                return 'rgba(255, 165, 0, 0.95)' # 45일 미만 (긴급: Orange)
                            else:
                                return 'rgba(0, 200, 81, 0.95)'  # 45일 이상 (공급 양호: Green)
                                
                        colors_gauge = df_low['가중소진기한_개월'].apply(get_neon_color)
                        
                        fig_gauge = go.Figure()
                        fig_gauge.add_trace(go.Bar(
                            y=df_low['품명'],
                            x=df_low['D-Day_일수'],
                            orientation='h',
                            marker=dict(
                                color=colors_gauge,
                                line=dict(color='rgba(255,255,255,0.2)', width=0.8)
                            ),
                            # D-Day가 0 이상인 경우 '+' 부호를 붙여 직관성 극대화
                            text=df_low['D-Day_일수'].apply(lambda x: f"D{x:.0f}일" if x < 0 else f"D+{x:.0f}일"),
                            textposition='outside',
                            hovertemplate='<b>%{y}</b><br>• 안전재고 대비 D-Day: %{x:.1f}일<br>• 실제 가중 소진기한: %{customdata:.1f}개월<extra></extra>',
                            customdata=df_low['가중소진기한_개월']
                        ))
                        
                        fig_gauge.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(26,26,46,0.3)',
                            font=dict(family="Outfit, Nanum Gothic, sans-serif", color="#F5F5F7"),
                            xaxis=dict(
                                title="안전재고(45일) 기준 D-Day (일수)",
                                showgrid=True,
                                gridcolor='rgba(255,255,255,0.05)'
                            ),
                            yaxis=dict(
                                title="",
                                showgrid=False,
                                categoryorder='total descending'
                            ),
                            # 긴 품목명이 잘리지 않도록 좌측 여백을 260으로 충분히 확보
                            margin=dict(l=260, r=50, t=20, b=50),
                            # 20개 품목이 쾌적하게 렌더링되도록 높이를 600으로 증대
                            height=600
                        )
                        st.plotly_chart(fig_gauge, use_container_width=True, key="scm_countdown_gauge_plot")
                    else:
                        st.success("🎉 축하합니다! 분석 가능한 품목이 존재하지 않습니다.")

                st.markdown("<br>", unsafe_allow_html=True)
                
                sec2_cols = st.columns([1.5, 1.5])
                
                # 차트 3: 기획/마케팅용 예측 괴리도(Gap) 알림 Diverging 리스트
                with sec2_cols[0]:
                    st.markdown("#### ⚖️ [기획/마케팅용] 예측 괴리도(Gap) 알림 리스트 (Diverging Gap Cards)")
                    st.caption("기획 예상 소진기한과 실제 판매 가중 소진기한의 편차(Gap)를 기준으로 과소예측(🩵진파랑)과 과대예측(🧡진주황) 카드를 정렬하여 마케팅 판매 계획을 전략 보정합니다.")
                    
                    df_gap_list = filtered_sojin.sort_values(by='Gap_Abs', ascending=False).head(6)
                    
                    if not df_gap_list.empty:
                        for i, (_, row) in enumerate(df_gap_list.iterrows()):
                            gap_val = row['예측괴리도_Gap']
                            p_name = row['품명']
                            est_date = row['예상소진기한_날짜']
                            act_months = row['가중소진기한_개월']
                            
                            if gap_val > 0:
                                badge_color = "#007BFF"  # 🩵 진파랑
                                badge_text = "과소 예측"
                                desc_text = f"실제 판매 속도가 빨라 예상보다 <b>{gap_val:.1f}개월 일찍</b> 소진 중 (품절 임박!)"
                                border_style = "border-left: 5px solid #007BFF;"
                            else:
                                badge_color = "#FFA500"  # 🧡 진주황
                                badge_text = "과대 예측"
                                desc_text = f"실제 판매 속도가 저하되어 예상보다 <b>{abs(gap_val):.1f}개월 더</b> 재고 적체 중"
                                border_style = "border-left: 5px solid #FFA500;"
                                
                            st.markdown(
                                f"<div class='kpi-card' style='background-color: #1A1A2E !important; margin-bottom: 8px; padding: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.05); color: #FFFFFF !important; {border_style}'>"
                                f"  <div style='display: flex; justify-content: space-between; align-items: center;'>"
                                f"    <span style='font-size: 13.5px; font-weight: 700; color: #FFFFFF !important;'>{i+1}. {p_name}</span>"
                                f"    <span style='background-color: {badge_color}; color: #FFFFFF; font-size: 10px; font-weight: bold; padding: 2px 6px; border-radius: 4px;'>{badge_text}</span>"
                                f"  </div>"
                                f"  <div style='font-size: 11.5px; color: #E0E0E0; margin-top: 4px;'>{desc_text}</div>"
                                f"  <div style='font-size: 10.5px; color: #A0A0A5; margin-top: 2px;'>기획 소진: {est_date} | 실제 가중: {act_months:.1f}개월</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("예측 괴리 분석 데이터가 없습니다.")

                # 오른쪽: ⚠️ 예측 오차 TOP 5 급변 경보판
                with sec2_cols[1]:
                    st.markdown("#### ⚠️ 예측 오차 TOP 5 급변 경보판")
                    st.caption("기획 부서의 판매 예상 속도와 최근 실제 수요 격차가 큰 품목들입니다. 생산 및 원자재 발주 수급을 긴급 조율하십시오.")
                    
                    top_gap = filtered_sojin.sort_values(by='Gap_Abs', ascending=False).head(5)
                    
                    for i, (_, row) in enumerate(top_gap.iterrows()):
                        p_name = row['품명']
                        gap_val = row['예측괴리도_Gap']
                        est_val_str = row['예상소진기한_날짜']
                        act_val = row['가중소진기한_개월']
                        
                        if gap_val > 0:
                            state_msg = f"🚨 실제 소진속도가 훨씬 빠름 (품절 임박, 괴리: {gap_val:.1f}개월)"
                            card_border = "border-left: 4px solid #FF4B4B;"
                        else:
                            state_msg = f"📦 실제 판매 속도 저하 (재고 과다, 괴리: {abs(gap_val):.1f}개월)"
                            card_border = "border-left: 4px solid #FFA500;"
                            
                        st.markdown(
                            f"<div class='kpi-card' style='background-color: #1A1A2E !important; color: #FFFFFF !important; margin-bottom: 8px; padding: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.05); {card_border}'>"
                            f"<div style='font-size: 13.5px; font-weight: 700; color: #FFFFFF !important;'>{i+1}. {p_name}</div>"
                            f"<div style='font-size: 11.5px; color: #C5A880; margin-top: 4px; font-weight: 600;'>{state_msg}</div>"
                            f"<div style='font-size: 11px; color: #A0A0A5;'>기획 소진: {est_val_str} | 가중 기한: {act_val:.1f}개월</div>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                
                st.markdown("---")
                
                # 3. 상세 데이터 테이블 및 정렬/검색 필터
                st.markdown("#### 🌲 완제품 품목별 소진기한 및 과거 판매 스크롤러")
                
                search_cols = st.columns([1.5, 1.2, 1.2])
                with search_cols[0]:
                    search_q = st.text_input("🔍 품명 실시간 검색", value="", key="sojin_search_input")
                with search_cols[1]:
                    cat_filter = st.selectbox("🧪 카테고리 필터", options=["전체", "🧪 디퓨저", "🧴 오일 퍼퓸", "🕯️ 캔들", "💨 룸 스프레이", "🧴 핸드 케어", "🎁 기타 완제품"], key="sojin_cat_selectbox")
                with search_cols[2]:
                    sort_order = st.selectbox("⚖️ 정렬 기준", options=["예측 괴리도(Gap) 높은 순", "가중 소진기한 짧은 순", "현재 재고 많은 순"], key="sojin_sort_selectbox")
                    
                # 데이터 필터링 적용 (상단 퀵 필터도 누적 결합 적용)
                if search_q:
                    filtered_sojin = filtered_sojin[filtered_sojin['품명'].str.contains(search_q, case=False, na=False)]
                if cat_filter != "전체":
                    filtered_sojin = filtered_sojin[filtered_sojin['카테고리'] == cat_filter]
                    
                # 데이터 정렬 적용
                if sort_order == "예측 괴리도(Gap) 높은 순":
                    filtered_sojin = filtered_sojin.sort_values(by='Gap_Abs', ascending=False)
                elif sort_order == "가중 소진기한 짧은 순":
                    filtered_sojin = filtered_sojin.sort_values(by='가중소진기한_개월', ascending=True)
                elif sort_order == "현재 재고 많은 순":
                    filtered_sojin = filtered_sojin.sort_values(by='현재재고', ascending=False)
                    
                # 테이블 컬럼 예쁘게 매핑
                display_sojin = filtered_sojin[['카테고리', '품명', '현재재고', '총판매수량_90일', '월평균 예상 판매량', '가중판매속도_월', '예상소진기한_날짜', '가중소진기한_개월', '예측괴리도_Gap']].copy()
                
                display_sojin = display_sojin.rename(columns={
                    '카테고리': '대분류',
                    '현재재고': '현재 보유 재고 (개)',
                    '총판매수량_90일': '과거 90일 판매 (개)',
                    '월평균 예상 판매량': '월평균 예상 판매 (개)',
                    '가중판매속도_월': '가중 판매 속도 (개/월)',
                    '예상소진기한_날짜': '예상 소진 기한 (날짜)',
                    '가중소진기한_개월': '가중 소진 기한 (개월)',
                    '예측괴리도_Gap': '예측 괴리 (개월)'
                })
                
                styled_table = display_sojin.style.format({
                    '현재 보유 재고 (개)': '{:,.0f}',
                    '과거 90일 판매 (개)': '{:,.0f}',
                    '월평균 예상 판매 (개)': '{:,.1f}',
                    '가중 판매 속도 (개/월)': '{:,.1f}',
                    '가중 소진 기한 (개월)': '{:,.1f}',
                    '예측 괴리 (개월)': '{:,.1f}'
                })
                
                def color_sojin_term(val):
                    try:
                        v = float(val)
                        if v < 1.5:
                            return 'background-color: rgba(255, 75, 75, 0.15); color: #FF4B4B; font-weight: bold;'
                        elif v < 3.0:
                            return 'background-color: rgba(255, 165, 0, 0.15); color: #FFA500;'
                        else:
                            return 'background-color: rgba(0, 200, 81, 0.1); color: #00C851;'
                    except:
                        return ''
                        
                styled_table = styled_table.map(color_sojin_term, subset=['가중 소진 기한 (개월)'])
                
                st.dataframe(
                    styled_table,
                    use_container_width=True,
                    hide_index=True,
                    key="sojin_final_dataframe"
                )
                
            # ====================================================
            # 서브 탭 [2]: 💀 악성 재고 모니터 (Dead Stock Monitor)
            # ====================================================
            with tab_dead:
                st.markdown("#### 💀 SCM 완제품 악성 재고 모니터 (Dead Stock Monitor)")
                st.warning("⚠️ **[SCM 긴급 경고]** 아래 품목들은 최근 90일(3개월) 동안 단 1건의 판매 실적도 발생하지 않았으나, 현재 창고에 실물 재고가 남아있어 기업의 자산 유동성을 침해하는 악성 적체 재고(Dead Stock)입니다. 즉각 특별 프로모션, B2B 대량 사은품 연계 또는 폐기 EOL 절차 등의 신속한 처분을 시행하십시오.")
                
                if not dead_sojin_master.empty:
                    # 카테고리별 가상의 SCM 평가 단가 사전 설정 (추정 재고자산 가치 평가용)
                    price_map = {
                        '🧪 디퓨저': 30000.0,
                        '🧴 오일 퍼퓸': 35000.0,
                        '🕯️ 캔들': 25000.0,
                        '💨 룸 스프레이': 20000.0,
                        '🧴 핸드 케어': 15000.0,
                        '🎁 기타 완제품': 10000.0
                    }
                    
                    # 묶인 추정 재고 자산 가치 계산
                    dead_sojin_master['평가단가'] = dead_sojin_master['카테고리'].map(price_map).fillna(10000.0)
                    dead_sojin_master['묶인_추정자산가치'] = dead_sojin_master['현재재고'] * dead_sojin_master['평가단가']
                    
                    # 묶여 있는 자산가치 규모가 가장 높은 순으로 정렬
                    dead_sojin_master = dead_sojin_master.sort_values(by='묶인_추정자산가치', ascending=False)
                    
                    total_dead_assets = dead_sojin_master['묶인_추정자산가치'].sum()
                    
                    # 자산 가치 규모 메인 KPI 리포트
                    asset_cols = st.columns(3)
                    with asset_cols[0]:
                        st.metric("💀 악성 적체 품목 수", f"{len(dead_sojin_master)} 개", delta="격리 관리 중", delta_color="inverse")
                    with asset_cols[1]:
                        st.metric("📦 악성 적체 총 수량", f"{dead_sojin_master['현재재고'].sum():,.0f} 개", delta="판매속도: 0개/월", delta_color="inverse")
                    with asset_cols[2]:
                        st.metric("💰 묶여 있는 추정 자산 규모", f"{total_dead_assets:,.0f} 원", delta="자산 유동성 침해", delta_color="inverse")
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # 상세 리스트 렌더링
                    st.markdown("##### 💀 악성 적체 긴급 처분 대상 목록")
                    
                    display_dead = dead_sojin_master[['카테고리', '품명', '현재재고', '묶인_추정자산가치']].copy()
                    display_dead = display_dead.rename(columns={
                        '카테고리': '대분류',
                        '현재재고': '현재 보유 재고 (개)',
                        '묶인_추정자산가치': '묶여 있는 추정 자산가치 (원)'
                    })
                    
                    styled_dead = display_dead.style.format({
                        '현재 보유 재고 (개)': '{:,.0f}',
                        '묶여 있는 추정 자산가치 (원)': '{:,.0f} 원'
                    })
                    
                    st.dataframe(
                        styled_dead,
                        use_container_width=True,
                        hide_index=True,
                        key="sojin_deadstock_dataframe"
                    )
                    
                    # SCM 자산 유동화 액션 추천 패널 배치
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("##### 💡 Anti-Gravity SCM 자산 유동화 추천 액션 가이드")
                    
                    act_cols = st.columns(3)
                    with act_cols[0]:
                        st.markdown(
                            "<div style='background-color: #1A1A2E; border-left: 4px solid #007BFF; padding: 12px; border-radius: 6px; height: 140px;'>"
                            "<div style='font-size: 13px; font-weight: 700; color: #FFFFFF;'>🎁 B2B / B2C 마케팅 증정 연계</div>"
                            "<div style='font-size: 11px; color: #A0A0A5; margin-top: 6px; line-height: 1.4;'>"
                            "기획/마케팅팀과 연계하여 고회전 제품(디퓨저/퍼퓸) 구매 시 사은품으로 번들 패키징을 기획하고 패키지 프로모션 소진을 추진하십시오.</div>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                    with act_cols[1]:
                        st.markdown(
                            "<div style='background-color: #1A1A2E; border-left: 4px solid #FFA500; padding: 12px; border-radius: 6px; height: 140px;'>"
                            "<div style='font-size: 13px; font-weight: 700; color: #FFFFFF;'>🏷️ 80% Clearance 특별 덤핑</div>"
                            "<div style='font-size: 11px; color: #A0A0A5; margin-top: 6px; line-height: 1.4;'>"
                            "추정 자산가치 비중이 큰 품목(1위~3위)을 중심으로 패밀리 세일 또는 온라인 한정 80% 클리어런스 긴급 균일가 처분 행사를 개설하여 회수를 촉진하십시오.</div>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                    with act_cols[2]:
                        st.markdown(
                            "<div style='background-color: #1A1A2E; border-left: 4px solid #FF4B4B; padding: 12px; border-radius: 6px; height: 140px;'>"
                            "<div style='font-size: 13px; font-weight: 700; color: #FFFFFF;'>🛑 EOL (단종/폐기) 및 자산상각</div>"
                            "<div style='font-size: 11px; color: #A0A0A5; margin-top: 6px; line-height: 1.4;'>"
                            "장기 보관 수수료 및 물류센터 공간 적체 비용이 재고 자산 가치를 상회할 것으로 판단 시, 단종(End of Life)을 확정하고 세무상 자산 손실상각 절차를 밟으십시오.</div>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                else:
                    st.success("🎉 경축! 현재 창고에 최근 90일 미판매 악성 적체 재고가 단 1개도 존재하지 않아 완벽한 자산 선순환 상태를 유지하고 있습니다.")
        else:
            st.info("시계열 데이터 가공 중 에러가 발생하여 분석 지표를 노출할 수 없습니다.")
    else:
        st.warning("구글 스프레드시트에서 소진기한 재고 마스터를 로드하지 못했습니다.")

st.markdown("---")

# --- 실시간 CDC 및 웹소켓 데이터 동기화 디버그 패널 ---
st.subheader("🔌 Anti-Gravity Broker Console (실시간 동기화 상태)")

console_cols = st.columns([2, 1])

with console_cols[0]:
    st.markdown("##### 🖥️ 실시간 증분 업데이트 및 캐시 로그 (Delta Sync Log)")
    log_content = "\n".join(st.session_state.websocket_log)
    st.text_area("WebSocket & Cache Event Stream", value=log_content, height=180, disabled=True, key="log_textarea")

with console_cols[1]:
    st.markdown("##### ⚙️ 구글 시트 연동 설정 정보")
    st.markdown(f"""
    * **동기화 방식**: `Incremental Delta Sync (Multi-Sheets Concat)`
    * **연동 시트 탭**: `25년 판매 (2차) + 26년 판매 (1차) 통합`
    * **총 고유 품목 수**: `{quantity_df['product_name'].nunique()} 개`
    * **최종 데이터 갱신 시점**: `{st.session_state.last_update_time.strftime('%Y-%m-%d %H:%M:%S')}`
    * **메모리 캐싱 상태**: `@st.cache_data` 활성화됨 (5분 캐시 보호)
    """)
