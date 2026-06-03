import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import requests
import time
import re
import io
import os

# 가상 서버 루트 작업 영역 직통 세팅
SAVED_EXCEL_PATH = "permanent_production_schedule.xlsx"
NOTES_DB_PATH = "production_notes.txt"
MASTER_PASSWORD = "Fineformulation"

# [오너 지시 정규식 핵심 축]: 띄어쓰기, 언더바 다 무시하고 오직 앞자리 순수 6자리 코드만 정밀 추출
def extract_pure_6_code(text):
    if not text:
        return ""
    cleaned = str(text).replace(" ", "").replace("_", "").replace("\r", "").replace("\n", "").replace("\t", "").strip().upper()
    match = re.search(r'(\d{5}[A-Z])', cleaned)
    return match.group(1) if match else ""

# [대표님 명세 1순위 조항]: 외부 서버 차단막을 원천 파괴하기 위해 로컬 저장 파일을 1순위로 호출
def get_saved_local_image_bytes(pure_code):
    pure_code_clean = str(pure_code).strip().upper()
    target_path = f"{pure_code_clean}.png"
    if os.path.exists(target_path):
        try:
            with open(target_path, "rb") as f:
                return f.read()
        except Exception:
            return None
    return None

# =========================================================================
# 2. 스트림릿 웹 대시보드 UI 레이아웃 구성
# =========================================================================
st.set_page_config(layout="wide", page_title="생산 스케줄 비주얼 대시보드")
st.title("🏭 생산 스케줄 마스터 시스템")

has_saved_file = os.path.exists(SAVED_EXCEL_PATH)
final_file_target = SAVED_EXCEL_PATH if os.path.exists(SAVED_EXCEL_PATH) else None

# 사이드바 통합 제어 센터 배치
with st.sidebar:
    st.markdown('<div style="font-size:20px; font-weight:bold; color:#38bdf8; margin-bottom:15px; border-bottom:2px solid #38bdf8; padding-bottom:5px;">⚙️ 마스터 데이터 제어 센터</div>', unsafe_allow_html=True)
    
    if has_saved_file:
        st.markdown('<div style="color:#4ade80; font-size:14px; font-weight:bold; background-color:#064e3b; padding:10px; border-radius:8px; margin-bottom:15px;">🟢 시스템 가동중.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="color:#f87171; font-size:14px; font-weight:bold; background-color:#7f1d1d; padding:10px; border-radius:8px; margin-bottom:15px;">💡 마스터 엑셀 파일 업로드가 필요합니다.</div>', unsafe_allow_html=True)
    
    input_password = st.text_input("🔓 관리자 승인 인증 (Password)", type="password", key="auth_pwd_input")
    is_authenticated = (input_password == MASTER_PASSWORD)

    st.markdown("---")
    st.write("📂 **새로운 스케줄 파일 업로드 / 교체**")
    uploaded_file = st.file_uploader("여기에 엑셀 파일을 드래그 앤 드롭 하세요.", type=["xlsx", "xls"], label_visibility="collapsed")
    
    if uploaded_file and is_authenticated:
        with open(SAVED_EXCEL_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("🚀 마스터 스케줄 파일 교체 성공!")
        time.sleep(1)
        st.rerun()

# ---------------------------------------------------------------------
# [트렐로 카드 공식 커버 대표 썸네일 수집 파이프라인]
# ---------------------------------------------------------------------
if final_file_target:
    raw_df = pd.read_excel(final_file_target, usecols="A,C,F,K,L,O,P,U", header=None)
    if raw_df.iloc[0].astype(str).str.contains('일정|코드|카테고리|Date|Item').any():
        raw_df = raw_df.iloc[1:]
        
    df = pd.DataFrame()
    df['item_code'] = raw_df.iloc[:, 0].fillna('-').astype(str).str.strip()
    df['category'] = raw_df.iloc[:, 1].fillna('기타 카테고리').astype(str).str.strip()
    df['price_tag'] = raw_df.iloc[:, 2].fillna('-').astype(str).str.strip().replace(['nan', 'NAN', 'NaN', 'None', ''], '-')
    df['po_number'] = raw_df.iloc[:, 3].fillna('-').astype(str).str.strip().replace(['nan', 'NAN', 'NaN', 'None', ''], '-')
    df['bag_number'] = raw_df.iloc[:, 4].fillna('-').astype(str).str.strip().replace(['nan', 'NAN', 'NaN', 'None', ''], '-')
    df['product_name'] = raw_df.iloc[:, 5].fillna('-').astype(str).str.strip()
    df['quantity'] = pd.to_numeric(raw_df.iloc[:, 6], errors='coerce').fillna(0).astype(int)
    df['production_date'] = pd.to_datetime(raw_df.iloc[:, 7], errors='coerce')
    
    df['volume'] = df['product_name'].apply(lambda x: re.search(r'(\d+ml|\d+oz|\d+g)', x, re.IGNORECASE).group(1) if re.search(r'(\d+ml|\d+oz|\d+g)', x, re.IGNORECASE) else "500ml")
    df = df.dropna(subset=['production_date'])
    df = df.sort_values(by=['category', 'item_code'], ascending=[True, True])
    
    today_dt = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    current_weekday = today_dt.weekday() 
    
    next_monday_dist = (7 - current_weekday) % 7
    if next_monday_dist == 0:
        next_monday_dist = 7
    
    target_next_monday = (today_dt + timedelta(days=next_monday_dist)).replace(hour=23, minute=59, second=59)
    second_monday_start = target_next_monday + timedelta(seconds=1)
    target_second_monday = (second_monday_start + timedelta(days=6)).replace(hour=23, minute=59, second=59)
    
    df_1week = df[(df['production_date'] >= today_dt) & (df['production_date'] <= target_next_monday)]
    df_2weeks = df[(df['production_date'] >= second_monday_start) & (df['production_date'] <= target_second_monday)]
    df_filtered_total = pd.concat([df_1week, df_2weeks]).copy()

    with st.sidebar:
        st.markdown("---")
        st.markdown('<div style="font-size:16px; font-weight:bold; color:#fbbf24;">⚡ 트렐로 이미지 서버 백업</div>', unsafe_allow_html=True)
        
        if st.button("🔄 현재 스케줄 이미지 서버에 저장", use_container_width=True):
            if is_authenticated:
                sync_success_count = 0
                
                status_placeholder = st.empty()
                status_placeholder.info("🔄 마스터 엑셀에서 순수 6자리 코드 스캔 중...")
                
                raw_excel_data = pd.read_excel(SAVED_EXCEL_PATH, header=None)
                target_pure_codes = []
                for col_idx in raw_excel_data.columns:
                    cell_values = raw_excel_data[col_idx].dropna().astype(str)
                    for val in cell_values:
                        p_code = extract_pure_6_code(val)
                        if p_code:
                            target_pure_codes.append(p_code)
                target_pure_codes = list(set(target_pure_codes))
                
                if len(target_pure_codes) > 0:
                    TRELLO_API_KEY = st.secrets["TRELLO_API_KEY"]
                    TRELLO_TOKEN = st.secrets["TRELLO_TOKEN"]
                    TRELLO_BOARD_ID = st.secrets["TRELLO_BOARD_ID"]
                    
                    secured_headers = {
                        "Authorization": f'OAuth oauth_consumer_key="{TRELLO_API_KEY}", oauth_token="{TRELLO_TOKEN}"',
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                    }
                    
                    status_placeholder.info(f"🛰️ 오리지널 썸네일 수집 허브 연동 (총 {len(target_pure_codes)}개 코드 대조)...")
                    # 커버 데이터를 완벽히 확보하기 위해 카드 목록 호출 단계에 옵션을 강제 주입합니다.
                    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/cards"
                    params = {'key': TRELLO_API_KEY, 'token': TRELLO_TOKEN, 'attachments': 'true', 'attachment_fields': 'all', 'limit': '1000'}
                    card_res = requests.get(url, headers=secured_headers, params=params, timeout=25)
                    
                    if card_res.status_code == 200:
                        all_cards = card_res.json()
                        progress_bar = st.progress(0)
                        total_items = len(target_pure_codes)
                        
                        for i, code_key in enumerate(target_pure_codes):
                            code_key_clean = str(code_key).strip().upper()
                            trello_url = None
                            
                            status_placeholder.info(f"⏳ 썸네일 정밀 매칭 중: [{code_key_clean}] ({i+1}/{total_items}) | 현재까지 {sync_success_count}개 박제 성공")
                            
                            for card in all_cards:
                                card_name_clean = card.get('name', '').replace(" ", "").replace("_", "").upper()
                                if code_key_clean in card_name_clean:
                                    cover = card.get('cover', {})
                                    
                                    # [🚨 도면 파일 간섭 배제 수술 축]: 도면 첨부파일이 아니라 실무진이 지정한 공식 카드 "대표 커버 이미지" 주소를 타겟팅
                                    if cover and cover.get('scaled'):
                                        scaled_images = cover.get('scaled', [])
                                        if scaled_images:
                                            # 가장 화질이 높은 대형 대표 썸네일 규격 주소 획득
                                            trello_url = scaled_images[-1].get('url')
                                            
                                    # 만약 공식 커버가 지정되지 않은 예외적인 경우에만 첨부파일 백업본 작동
                                    if not trello_url:
                                        attachments = card.get('attachments', [])
                                        if attachments:
                                            for att in attachments:
                                                a_url = att.get('url', '')
                                                # 도면(label, cap, size) 텍스트 찌꺼기가 파일명에 들어있으면 강제 스킵 필터링
                                                a_url_lower = a_url.lower()
                                                if any(skip in a_url_lower for skip in ['label', 'cap', 'size', 'spec', '도면']):
                                                    continue
                                                if any(ext in a_url_lower for ext in ['.jpg', '.jpeg', '.png', '.webp']):
                                                    trello_url = a_url
                                                    break
                                if trello_url:
                                    break
                            
                            if trello_url:
                                try:
                                    img_res = requests.get(trello_url, headers=secured_headers, timeout=12)
                                    if img_res.status_code == 200:
                                        target_save_file = f"{code_key_clean}.png"
                                        with open(target_save_file, "wb") as img_f:
                                            img_f.write(img_res.content)
                                        sync_success_count += 1
                                except Exception:
                                    pass
                            progress_bar.progress(int((i + 1) / total_items * 100))
                        
                        status_placeholder.empty()
                        st.markdown(f'<div style="color:#4ade80; font-size:16px; font-weight:bold; background-color:#064e3b; padding:12px; border-radius:8px; margin-top:10px;">🎯 백업 마감 결과: 총 {sync_success_count}개 품목의 공식 정형 썸네일 다른 이름 저장 성공! 바로 F5를 눌러 확인하십시오.</div>', unsafe_allow_html=True)
                    else:
                        status_placeholder.empty()
                        st.error("❌ 트렐로 API 통신 세션 인증 실패.")
                else:
                    status_placeholder.empty()
                    st.error("❌ 마스터 엑셀에서 코드를 식별하지 못했습니다.")
            else:
                st.error("❌ 패스워드 승인이 필요합니다.")

    # ---------------------------------------------------------------------
    # 5. 마스터 대시보드 메인 레이아웃 카드 스펙 마감 구역
    # ---------------------------------------------------------------------
    box_style = "width:100% !important; height:180px !important; background-color:#1e293b !important; border-radius:10px !important; display:flex !important; align-items:center !important; justify-content:center !important; overflow:hidden !important; margin-bottom:12px !important;"
    card_container_style = "background-color:#1e2530 !important; border:1px solid #2d3748 !important; border-radius:14px !important; padding:18px !important; box-shadow:0 10px 15px -3px rgba(0,0,0,0.4) !important;"
    text_base = "margin:0px !important; padding:0px !important; text-align:left !important; line-height:1.4 !important;"

    st.markdown("""
        <style>
            div[data-testid="stTextInput"] { margin-top: -15px !important; padding: 0px 5px !important; }
            div[data-testid="stTextInput"] input { background-color: #111622 !important; color: #ffffff !important; border: 1px solid #2d3748 !important; border-radius: 6px !important; font-size: 13px !important; height: 32px !important; }
            /* 고해상도 화장품 디자인 원본 정밀 보존 고정 비율 CSS 프로토콜 */
            div[data-testid="stImage"] { display: flex !important; justify-content: center !important; background-color: #1e293b !important; border-radius: 12px !important; padding: 10px !important; margin-bottom: 8px !important; height: 180px !important; align-items: center !important; }
            div[data-testid="stImage"] img { max-height: 160px !important; width: auto !important; object-fit: contain !important; }
        </style>
    """, unsafe_allow_html=True)

    # [섹션 1] 1주 차 생산 라인업
    st.markdown("---")
    st.subheader(f"📅 1주 차 생산 스케줄 대쉬보드 ({today_dt.strftime('%m/%d')} ~ {target_next_monday.strftime('%m/%d')})")
    
    if not df_1week.empty:
        for category_name, group_df in df_1week.groupby('category', sort=False):
            st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#38bdf8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #38bdf8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 {category_name} care</div>', unsafe_allow_html=True)
            cols = st.columns(6)
            for idx, row in group_df.reset_index().iterrows():
                excel_code = row['item_code']
                pure_excel_code = extract_pure_6_code(excel_code)
                
                with cols[idx % 6]:
                    local_saved_bytes = get_saved_local_image_bytes(pure_excel_code)
                    
                    if local_saved_bytes:
                        st.image(local_saved_bytes, use_container_width=False)
                    else:
                        st.html(f'<div style="{box_style}"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center; padding:10px;">{excel_code}<br>[백업 단추 클릭 필요]</div></div>')
                    
                    st.html(f"""
                        <div style="{card_container_style}">
                            <div style="{text_base} font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important; letter-spacing:0.5px !important;">{excel_code}</div>
                            <div style="{text_base} font-size:14px !important; color:#a0aec0 !important; font-weight:500 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                            <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                            <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                <span style="{text_base} font-size:14px !important; color:#718096 !important;">가격표: <span style="color:#63b3ed !important; font-weight:bold !important;">{row['price_tag']}</span></span>
                                <span style="{text_base} font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                            </div>
                            <div style="{text_base} font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                            <div style="{text_base} font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                            <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !items:center !important; margin-bottom:15px !important;">
                                <span style="{text_base} font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                <span style="{text_base} font-size:13px !important; color:#a0aec0 !important; font-weight:500 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                            </div>
                        </div>
                    """)
                    st.markdown('<div style="margin-bottom:25px;"></div>', unsafe_allow_html=True)

    # [섹션 2] 2주 차 생산 라인업
    st.markdown("---")
    st.subheader(f"📅 2주 차 생산 스케줄 대쉬보드 ({second_monday_start.strftime('%m/%d')} ~ {target_second_monday.strftime('%m/%d')})")
    
    if not df_2weeks.empty:
        for category_name, group_df in df_2weeks.groupby('category', sort=False):
            st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#38bdf8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #38bdf8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 {category_name} care</div>', unsafe_allow_html=True)
            cols = st.columns(6)
            for idx, row in group_df.reset_index().iterrows():
                excel_code = row['item_code']
                pure_excel_code = extract_pure_6_code(excel_code)
                
                with cols[idx % 6]:
                    local_saved_bytes = get_saved_local_image_bytes(pure_excel_code)
                    
                    if local_saved_bytes:
                        st.image(local_saved_bytes, use_container_width=False)
                    else:
                        st.html(f'<div style="{box_style}"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center; padding:10px;">{excel_code}<br>[백업 단추 클릭 필요]</div></div>')
                    
                    st.html(f"""
                        <div style="{card_container_style}">
                            <div style="{text_base} font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important; letter-spacing:0.5px !important;">{excel_code}</div>
                            <div style="{text_base} font-size:14px !important; color:#a0aec0 !important; font-weight:500 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                            <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                            <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                <span style="{text_base} font-size:14px !important; color:#718096 !important;">가격표: <span style="color:#63b3ed !important; font-weight:bold !important;">{row['price_tag']}</span></span>
                                <span style="{text_base} font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                            </div>
                            <div style="{text_base} font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                            <div style="{text_base} font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                            <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !items:center !important; margin-bottom:15px !important;">
                                <span style="{text_base} font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                <span style="{text_base} font-size:13px !important; color:#a0aec0 !important; font-weight:500 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                            </div>
                        </div>
                    """)
                    st.markdown('<div style="margin-bottom:25px;"></div>', unsafe_allow_html=True)
else:
    st.info("💡 스케줄 마스터 엑셀 파일 로드 대기중")