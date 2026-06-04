import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import requests
import time
import re
import io
import os
import base64
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image as PILImage

# 가상 서버 루트 작업 영역 직통 세팅
SAVED_EXCEL_PATH = "permanent_production_schedule.xlsx"
NOTES_DB_PATH = "production_notes.txt"
MASTER_PASSWORD = "Fineformulation"
ENTRY_SECURITY_CODE = "1234"      # [대표님 지정 핵심 보안 코드]
SESSION_TIMEOUT_SEC = 300          # [대표님 지정: 열람 유효시간 5분 (300초)]

# [오너 지시 정규식 핵심 축]: 띄어쓰기, 언더바 다 무시하고 오직 앞자리 순수 6자리 코드만 정밀 추출
def extract_pure_6_code(text):
    if not text:
        return ""
    cleaned = str(text).replace(" ", "").replace("_", "").replace("\r", "").replace("\n", "").replace("\t", "").strip().upper()
    match = re.search(r'(\d{5}[A-Z])', cleaned)
    return match.group(1) if match else ""

# [대표님 명세 1순위 조항]: 로컬 저장 파일을 1순위로 호출
def get_saved_local_image_base64(pure_code):
    pure_code_clean = str(pure_code).strip().upper()
    target_path = f"{pure_code_clean}.png"
    if os.path.exists(target_path):
        try:
            with open(target_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode()
                return f"data:image/png;base64,{encoded}"
        except Exception:
            return None
    return None

# ---------------------------------------------------------------------
# [📝 특기사항 1, 2 멀티 메모리 영구 저장 엔진]
# ---------------------------------------------------------------------
def load_production_notes():
    notes = {}
    if os.path.exists(NOTES_DB_PATH):
        try:
            with open(NOTES_DB_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    if "::" in line:
                        parts = line.split("::", 2)
                        code = parts[0].strip()
                        memo1 = parts[1].strip() if len(parts) > 1 else ""
                        memo2 = parts[2].strip() if len(parts) > 2 else ""
                        notes[code] = (memo1, memo2)
        except Exception:
            pass
    return notes

def save_production_note(pure_code, memo1, memo2):
    notes = load_production_notes()
    notes[pure_code] = (memo1.strip(), memo2.strip())
    try:
        with open(NOTES_DB_PATH, "w", encoding="utf-8") as f:
            for code, values in notes.items():
                if values[0] or values[1]:
                    f.write(f"{code}::{values[0]}::{values[1]}\n")
    except Exception:
        pass

# =========================================================================
# 2. 스트림릿 웹 대시보드 UI 레이아웃 구성
# =========================================================================
st.set_page_config(layout="wide", page_title="생산 스케줄 마스터 데이터 경영 대시보드")

# ---------------------------------------------------------------------
# [🚨 실시간 사용 감지 5분 연장 엔진]
# ---------------------------------------------------------------------
if "app_unlocked" not in st.session_state:
    st.session_state["app_unlocked"] = False
if "unlock_time" not in st.session_state:
    st.session_state["unlock_time"] = None

if st.session_state["app_unlocked"] and st.session_state["unlock_time"] is not None:
    elapsed_time = time.time() - st.session_state["unlock_time"]
    if elapsed_time > SESSION_TIMEOUT_SEC:
        st.session_state["app_unlocked"] = False
        st.session_state["unlock_time"] = None
        st.toast("⚠️ 보안 유지를 위해 자리를 비우신 지 5분이 경과되어 자동 잠금되었습니다.")
        time.sleep(1)
        st.rerun()
    else:
        st.session_state["unlock_time"] = time.time()

# ---------------------------------------------------------------------
# [🔒 게이트웨이 정문 차단막 인터페이스]
# ---------------------------------------------------------------------
if not st.session_state["app_unlocked"]:
    st.markdown("""
        <style>
            .stApp { background-color: #0f172a !important; }
            .security-gate {
                text-align: center;
                margin-top: 15vh;
                padding: 40px;
                background-color: #1e2530;
                border: 2px solid #38bdf8;
                border-radius: 16px;
                box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.7);
                max-width: 500px;
                margin-left: auto;
                margin-right: auto;
            }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div class="security-gate">
            <h1 style="color: #38bdf8; font-size: 28px; font-weight: bold; margin-bottom: 10px;">🔒 FINE FORMULATION</h1>
            <p style="color: #94a3b8; font-size: 15px; margin-bottom: 25px;">본 시스템은 기업 기밀 자산 보호 구역입니다.<br>열람 유효시간은 5분이며, <b>사용 중일 경우 실시간으로 자동 연장</b>됩니다.</p>
        </div>
    """, unsafe_allow_html=True)
    
    cols = st.columns([1, 2, 1])
    with cols[1]:
        input_gate_code = st.text_input("🔑 보안 코드 입력 (Security Code)", type="password", key="gate_code_input")
        
        if input_gate_code == ENTRY_SECURITY_CODE:
            st.session_state["app_unlocked"] = True
            st.session_state["unlock_time"] = time.time()
            st.success("🔓 자격 증명이 확인되었습니다. 시스템을 개방합니다.")
            time.sleep(0.5)
            st.rerun()
        elif input_gate_code != "":
            st.error("❌ 보안 코드가 올바르지 않습니다. 접근이 거부되었습니다.")
            
    st.stop()

# ---------------------------------------------------------------------
# [🔓 1234 통과 시 오픈되는 마스터 대시보드 코어]
# ---------------------------------------------------------------------
has_saved_file = os.path.exists(SAVED_EXCEL_PATH)
final_file_target = SAVED_EXCEL_PATH if os.path.exists(SAVED_EXCEL_PATH) else None

with st.sidebar:
    st.markdown(f'<div style="color:#ffffff; font-size:15px; font-weight:bold; background-color:#0284c7; padding:10px; border-radius:8px; margin-bottom:15px; text-align:center;">🟢 시스템 가동 중 (활동 중 자동 연장)</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:20px; font-weight:bold; color:#38bdf8; margin-bottom:15px; border-bottom:2px solid #38bdf8; padding-bottom:5px;">⚙️ 마스터 데이터 제어 센터</div>', unsafe_allow_html=True)
    
    if has_saved_file:
        st.markdown('<div style="color:#4ade80; font-size:14px; font-weight:bold; background-color:#064e3b; padding:10px; border-radius:8px; margin-bottom:15px;">🟢 스케줄 파일 연동 완료</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="color:#f87171; font-size:14px; font-weight:bold; background-color:#7f1d1d; padding:10px; border-radius:8px; margin-bottom:15px;">💡 마스터 엑셀 파일 업로드가 필요합니다.</div>', unsafe_allow_html=True)
    
    input_password = st.text_input("🔓 데이터 제어 승인 암호", type="password", key="auth_pwd_input")
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
        
    st.markdown("---")
    if st.button("🔒 대시보드 즉시 잠금 (로그아웃)", use_container_width=True):
        st.session_state["app_unlocked"] = False
        st.session_state["unlock_time"] = None
        st.rerun()

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
    
    # 주차와 카테고리 안에서 동일 코드 밀착 정렬 알고리즘
    df = df.sort_values(by=['category', 'item_code', 'production_date'], ascending=[True, True, True])
    
    today_dt = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    current_weekday = today_dt.weekday() 
    
    next_monday_dist = (7 - current_weekday) % 7
    if next_monday_dist == 0:
        next_monday_dist = 7
    
    target_next_monday = (today_dt + timedelta(days=next_monday_dist)).replace(hour=23, minute=59, second=59)
    second_monday_start = target_next_monday + timedelta(seconds=1)
    target_second_monday = (second_monday_start + timedelta(days=6)).replace(hour=23, minute=59, second=59)
    
    df_1week = df[(df['production_date'] >= today_dt) & (df['production_date'] <= target_next_monday)].copy()
    df_2weeks = df[(df['production_date'] >= second_monday_start) & (df['production_date'] <= target_second_monday)].copy()
    
    saved_notes = load_production_notes()

    # ---------------------------------------------------------------------
    # [🚨 오너 지시 핵심 패치: 특기사항 1, 2 칼럼이 추가된 엑셀 컴파일러]
    # ---------------------------------------------------------------------
    def generate_premium_split_excel(df_w1, df_w2):
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "주차별_생산라인업"
        ws.views.sheetView[0].showGridLines = True
        
        # 디자인 서식 프로토콜 설정
        font_main_title = Font(name="Malgun Gothic", size=14, bold=True, color="FFFFFF")
        font_header = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Malgun Gothic", size=10)
        font_group = Font(name="Malgun Gothic", size=11, bold=True, color="0f172a")
        
        fill_week_title = PatternFill(start_color="0369a1", end_color="0369a1", fill_type="solid")
        fill_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        fill_group = PatternFill(start_color="f8fafc", end_color="f8fafc", fill_type="solid")
        
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        
        thin_side = Side(border_style="thin", color="cbd5e1")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        # [🚨 특기사항 1, 2 칼럼 전격 끝단 신설 명세 채결]
        headers = ["카테고리 그룹", "아이템 사진", "아이템 코드", "아이템 이름", "용량", "생산 수량", "PO 번호", "가격표 유무", "특기사항 1", "특기사항 2"]
        categories_order = ["skin", "body", "hair", "기타 카테고리"]
        current_row_idx = 1
        
        def write_week_block(ws, target_df, week_label_text, start_row):
            r_idx = start_row
            
            # 주차별 헤드라인 대형 바 마감
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=10)
            title_cell = ws.cell(row=r_idx, column=1)
            title_cell.value = week_label_text
            title_cell.font = font_main_title
            title_cell.fill = fill_week_title
            title_cell.alignment = align_center
            ws.row_dimensions[r_idx].height = 35
            r_idx += 1
            
            # 테이블 컬럼 인덱스 매핑
            for col_num, h_text in enumerate(headers, 1):
                h_cell = ws.cell(row=r_idx, column=col_num, value=h_text)
                h_cell.font = font_header
                h_cell.fill = fill_header
                h_cell.alignment = align_center
                h_cell.border = border_all
            ws.row_dimensions[r_idx].height = 25
            r_idx += 1
            
            # 카테고리 섹션 빌딩
            for cate in categories_order:
                if cate != "기타 카테고리":
                    cate_df = target_df[target_df['category'].str.lower().str.contains(cate)]
                else:
                    cate_df = target_df[~target_df['category'].str.lower().str.contains('skin|body|hair')]
                    
                if not cate_df.empty:
                    ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=10)
                    g_cell = ws.cell(row=r_idx, column=1)
                    g_cell.value = f"🌿 {cate.upper()} CARE LINEUP"
                    g_cell.font = font_group
                    g_cell.fill = fill_group
                    g_cell.alignment = align_left
                    for c_num in range(1, 11):
                        ws.cell(row=r_idx, column=c_num).border = border_all
                    ws.row_dimensions[r_idx].height = 24
                    r_idx += 1
                    
                    for _, r in cate_df.iterrows():
                        p_code = extract_pure_6_code(r['item_code'])
                        memo_vals = saved_notes.get(p_code, ("", ""))
                        
                        ws.cell(row=r_idx, column=1, value=r['category'])
                        ws.cell(row=r_idx, column=3, value=r['item_code'])
                        ws.cell(row=r_idx, column=4, value=r['product_name'])
                        ws.cell(row=r_idx, column=5, value=r['volume'])
                        
                        qty_cell = ws.cell(row=r_idx, column=6, value=r['quantity'])
                        qty_cell.number_format = '#,##0'
                        qty_cell.alignment = align_right
                        
                        ws.cell(row=r_idx, column=7, value=r['po_number'])
                        ws.cell(row=r_idx, column=8, value="유" if str(r['price_tag']).strip() != "-" else "무")
                        
                        # [🚨 실시간 특기사항 데이터 베이스 연동 매핑 축]
                        ws.cell(row=r_idx, column=9, value=memo_vals[0]).alignment = align_left
                        ws.cell(row=r_idx, column=10, value=memo_vals[1]).alignment = align_left
                        
                        for c_idx in range(1, 11):
                            c_cell = ws.cell(row=r_idx, column=c_idx)
                            c_cell.font = font_data
                            c_cell.border = border_all
                            if c_idx != 4 and c_idx != 6 and c_idx != 1 and c_idx != 9 and c_idx != 10:
                                c_cell.alignment = align_center
                            elif c_idx == 1:
                                c_cell.alignment = align_center
                                
                        ws.row_dimensions[r_idx].height = 35
                        img_path = f"{p_code}.png"
                        
                        if os.path.exists(img_path):
                            try:
                                pil_img = PILImage.open(img_path)
                                pil_img.thumbnail((50, 45))
                                
                                img_stream = io.BytesIO()
                                pil_img.save(img_stream, format="PNG")
                                img_stream.seek(0)
                                
                                xl_img = OpenpyxlImage(img_stream)
                                ws.add_image(xl_img, f"B{r_idx}")
                            except Exception:
                                pass
                        r_idx += 1
            return r_idx + 2
            
        next_start_row = write_week_block(ws, df_1week, f"🗓️ 1주 차 생산 라인업 계획 ({today_dt.strftime('%m/%d')} ~ {target_next_monday.strftime('%m/%d')})", current_row_idx)
        write_week_block(ws, df_2weeks, f"🗓️ 2주 차 생산 라인업 계획 ({second_monday_start.strftime('%m/%d')} ~ {target_second_monday.strftime('%m/%d')})", next_start_row)
        
        # 전체 칼럼 레이아웃 너비 밸런싱 최적화
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 38
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 14
        ws.column_dimensions['G'].width = 16
        ws.column_dimensions['H'].width = 14
        ws.column_dimensions['I'].width = 25 # 특기사항 1 폭 마감
        ws.column_dimensions['J'].width = 25 # 특기사항 2 폭 마감
        
        wb.save(output)
        return output.getvalue()

    with st.sidebar:
        st.markdown("---")
        st.markdown('<div style="font-size:16px; font-weight:bold; color:#38bdf8;">📥 오너 기획 데이터 추출 센터</div>', unsafe_allow_html=True)
        
        split_excel_bytes = generate_premium_split_excel(df_1week, df_2weeks)
        st.download_button(
            label="📊 주차별 분리 마스터 엑셀 다운로드",
            data=split_excel_bytes,
            file_name=f"Fine_Formulation_Split_Schedule_{datetime.now().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

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
                    
                    status_placeholder.info(f"🛰️ 오리지널 썸네일 수집 허브 연동 시작...")
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
                                    if cover and cover.get('scaled'):
                                        scaled_images = cover.get('scaled', [])
                                        if scaled_images:
                                            trello_url = scaled_images[-1].get('url')
                                            
                                if not trello_url:
                                    attachments = card.get('attachments', [])
                                    if attachments:
                                        for att in attachments:
                                            a_url = att.get('url', '')
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
                st.error("❌ 데이터 제어 승인 암호가 일치하지 않습니다.")

    # ---------------------------------------------------------------------
    # 5. [🚨 오너 지시 핵심 반영]: 배경 검은색 완벽 제거 및 화면 바탕 연동 CSS
    # ---------------------------------------------------------------------
    st.markdown("""
        <style>
            /* [핵심 수술 축]: 투명 박스 백그라운드로 전환하여 대시보드 바탕인 다크 네이비(#0f172a 계열)와 100% 카멜레온 동기화 */
            .owner-square-frame {
                width: 100% !important;
                aspect-ratio: 1 / 1 !important;
                background-color: transparent !important; /* 기존 검은색 박스 완벽 제거 및 배경색 투명화 마감 */
                border-radius: 0px !important;
                display: flex !important;
                justify-content: center !important;
                align-items: center !important;
                overflow: hidden !important;
                padding: 5px !important;
                box-sizing: border-box !important;
                margin-bottom: 8px !important;
            }
            .owner-square-frame img {
                max-width: 100% !important;
                max-height: 100% !important;
                width: auto !important;
                height: auto !important;
                object-fit: contain !important; /* 원본 패키지 짤림 0% 전체 노출 스펙 유지 */
            }
            .owner-info-card-wrap {
                background-color: #1e2530 !important; 
                border: 1px solid #2d3748 !important; 
                border-radius: 14px !important; 
                padding: 18px !important; 
                box-shadow: 0 10px 15px -3px rgba(0,0,0,0.4) !important;
                margin-bottom: 8px !important;
            }
            .owner-text-row {
                margin: 0px !important; 
                padding: 0px !important; 
                text-align: left !important; 
                line-height: 1.4 !important;
            }
            div[data-testid="stTextInput"] { margin-top: 4px !important; padding: 0px !important; }
            div[data-testid="stTextInput"] input { background-color: #0f172a !important; color: #38bdf8 !important; border: 1px solid #334155 !important; border-radius: 8px !important; font-size: 13px !important; height: 36px !important; }
            div[data-testid="stTextInput"] input:focus { border-color: #38bdf8 !important; }
        </style>
    """, unsafe_allow_html=True)

    def render_schedule_grid(target_df, title_label, section_prefix):
        st.markdown("---")
        st.subheader(title_label)
        
        if not target_df.empty:
            fixed_categories = ["skin", "body", "hair"]
            
            for cate in fixed_categories:
                group_df = target_df[target_df['category'].str.lower().str.contains(cate)]
                if not group_df.empty:
                    st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#38bdf8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #38bdf8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 {cate.upper()} care Lineup</div>', unsafe_allow_html=True)
                    cols = st.columns(6)
                    for idx, row in group_df.reset_index().iteritems() if hasattr(group_df.reset_index(), 'iteritems') else group_df.reset_index().iterrows():
                        excel_code = row['item_code']
                        pure_excel_code = extract_pure_6_code(excel_code)
                        
                        with cols[idx % 6]:
                            local_base64_data = get_saved_local_image_base64(pure_excel_code)
                            
                            if local_base64_data:
                                st.html(f'<div class="owner-square-frame"><img src="{local_base64_data}"></div>')
                            else:
                                st.html(f'<div class="owner-square-frame"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center; padding:10px;">{excel_code}<br>[백업 필요]</div></div>')
                            
                            st.html(f"""
                                <div class="owner-info-card-wrap">
                                    <div class="owner-text-row" style="font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important; letter-spacing:0.5px !important;">{excel_code}</div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#a0aec0 !important; font-weight:500 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                                    <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                                    <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                        <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">가격표: <span style="color:#63b3ed !important; font-weight:bold !important;">{row['price_tag']}</span></span>
                                        <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                                    </div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                                    <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !important; align-items:center !important;">
                                        <span class="owner-text-row" style="font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                        <span class="owner-text-row" style="font-size:13px !important; color:#a0aec0 !important; font-weight:500 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                                    </div>
                                </div>
                            """)
                            
                            memo_tuple = saved_notes.get(pure_excel_code, ("", ""))
                            key_m1 = f"input_m1_{section_prefix}_{pure_excel_code}_{idx}"
                            user_m1 = st.text_input(label=f"T1_{pure_excel_code}", value=memo_tuple[0], key=key_m1, placeholder="📋 특기사항 1 입력 후 Enter", label_visibility="collapsed")
                            key_m2 = f"input_m2_{section_prefix}_{pure_excel_code}_{idx}"
                            user_m2 = st.text_input(label=f"T2_{pure_excel_code}", value=memo_tuple[1], key=key_m2, placeholder="📦 특기사항 2 입력 후 Enter", label_visibility="collapsed")
                            
                            if user_m1 != memo_tuple[0] or user_m2 != memo_tuple[1]:
                                save_production_note(pure_excel_code, user_m1, user_m2)
                                st.rerun()
                                
                            st.markdown('<div style="margin-bottom:30px;"></div>', unsafe_allow_html=True)
            
            other_df = target_df[~target_df['category'].str.lower().str.contains('skin|body|hair')]
            if not other_df.empty:
                st.markdown('<div style="font-size:20px; font-weight:bold; color:#94a3b8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #94a3b8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 기타 카테고리 Lineup</div>', unsafe_allow_html=True)
                cols = st.columns(6)
                for idx, row in other_df.reset_index().iteritems() if hasattr(other_df.reset_index(), 'iteritems') else other_df.reset_index().iterrows():
                    excel_code = row['item_code']
                    pure_excel_code = extract_pure_6_code(excel_code)
                    with cols[idx % 6]:
                        local_base64_data = get_saved_local_image_base64(pure_excel_code)
                        if local_base64_data:
                            st.html(f'<div class="owner-square-frame"><img src="{local_base64_data}"></div>')
                        else:
                            st.html(f'<div class="owner-square-frame"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center; padding:10px;">{excel_code}<br>[백업 필요]</div></div>')
                        
                        st.html(f"""
                            <div class="owner-info-card-wrap">
                                <div class="owner-text-row" style="font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important; letter-spacing:0.5px !important;">{excel_code}</div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#a0aec0 !important; font-weight:500 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                                <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                                <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                    <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">가격표: <span style="color:#63b3ed !important; font-weight:bold !important;">{row['price_tag']}</span></span>
                                    <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                                </div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                                <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !important; align-items:center !important;">
                                    <span class="owner-text-row" style="font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                    <span class="owner-text-row" style="font-size:13px !important; color:#a0aec0 !important; font-weight:500 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                                </div>
                            </div>
                        """)
                        memo_tuple = saved_notes.get(pure_excel_code, ("", ""))
                        key_m1 = f"input_m1_oth_{pure_excel_code}_{idx}"
                        user_m1 = st.text_input(label=f"T1_{pure_excel_code}_oth", value=memo_tuple[0], key=key_m1, placeholder="📋 특기사항 1 입력 후 Enter", label_visibility="collapsed")
                        key_m2 = f"input_m2_oth_{pure_excel_code}_{idx}"
                        user_m2 = st.text_input(label=f"T2_{pure_excel_code}_oth", value=memo_tuple[1], key=key_m2, placeholder="📦 특기사항 2 입력 후 Enter", label_visibility="collapsed")
                        if user_m1 != memo_tuple[0] or user_m2 != memo_tuple[1]:
                            save_production_note(pure_excel_code, user_m1, user_m2)
                            st.rerun()
                        st.markdown('<div style="margin-bottom:30px;"></div>', unsafe_allow_html=True)

    # 1주 차 및 2주 차 통합 그리드 빌드 가동
    render_schedule_grid(df_1week, f"📅 1주 차 생산 스케줄 대쉬보드 ({today_dt.strftime('%m/%d')} ~ {target_next_monday.strftime('%m/%d')})", "w1")
    render_schedule_grid(df_2weeks, f"📅 2주 차 생산 스케줄 대쉬보드 ({second_monday_start.strftime('%m/%d')} ~ {target_second_monday.strftime('%m/%d')})", "w2")

else:
    st.info("💡 마스터 엑셀 파일 로드 대기중")