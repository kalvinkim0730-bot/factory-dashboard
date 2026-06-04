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

# 경로 및 보안 설정
SAVED_EXCEL_PATH = "permanent_production_schedule.xlsx"
NOTES_DB_PATH = "production_notes.txt"
MASTER_PASSWORD = "Fineformulation"
ENTRY_SECURITY_CODE = "1234"
SESSION_TIMEOUT_SEC = 300

def extract_pure_6_code(text):
    if not text: return ""
    cleaned = str(text).replace(" ", "").replace("_", "").strip().upper()
    match = re.search(r'(\d{5}[A-Z])', cleaned)
    return match.group(1) if match else ""

def get_saved_local_image_base64(pure_code):
    target_path = f"{str(pure_code).strip().upper()}.png"
    if os.path.exists(target_path):
        try:
            with open(target_path, "rb") as f:
                return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
        except: return None
    return None

def load_production_notes():
    notes = {}
    if os.path.exists(NOTES_DB_PATH):
        try:
            with open(NOTES_DB_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    if "::" in line:
                        parts = line.split("::", 2)
                        if len(parts) >= 3: notes[parts[0].strip()] = (parts[1].strip(), parts[2].strip())
        except: pass
    return notes

def save_production_note(pure_code, memo1, memo2):
    notes = load_production_notes()
    notes[pure_code] = (memo1.strip(), memo2.strip())
    try:
        with open(NOTES_DB_PATH, "w", encoding="utf-8") as f:
            for code, values in notes.items():
                if values[0] or values[1]: f.write(f"{code}::{values[0]}::{values[1]}\n")
    except: pass

st.set_page_config(layout="wide", page_title="생산 스케줄 마스터")

# 보안 엔진
if "app_unlocked" not in st.session_state: st.session_state["app_unlocked"] = False
if "unlock_time" not in st.session_state: st.session_state["unlock_time"] = None

if st.session_state["app_unlocked"] and st.session_state["unlock_time"]:
    if time.time() - st.session_state["unlock_time"] > SESSION_TIMEOUT_SEC:
        st.session_state["app_unlocked"] = False
        st.rerun()
    else: st.session_state["unlock_time"] = time.time()

if not st.session_state["app_unlocked"]:
    st.markdown('<div style="text-align:center; margin-top:15vh; padding:40px; background-color:#1e2530; border:2px solid #38bdf8; border-radius:16px; max-width:500px; margin-left:auto; margin-right:auto;"><h1 style="color:#38bdf8;">🔒 FINE FORMULATION</h1></div>', unsafe_allow_html=True)
    cols = st.columns([1, 2, 1])
    with cols[1]:
        if st.text_input("보안 코드", type="password") == ENTRY_SECURITY_CODE:
            st.session_state["app_unlocked"] = True
            st.session_state["unlock_time"] = time.time()
            st.rerun()
    st.stop()

has_saved_file = os.path.exists(SAVED_EXCEL_PATH)
with st.sidebar:
    st.title("⚙️ 제어 센터")
    input_password = st.text_input("승인 암호", type="password")
    is_authenticated = (input_password == MASTER_PASSWORD)
    uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xls"])
    if uploaded_file and is_authenticated:
        with open(SAVED_EXCEL_PATH, "wb") as f: f.write(uploaded_file.getbuffer())
        st.success("파일 교체 성공")
        st.rerun()

if has_saved_file:
    # 헤더 꼬임 버그를 차단하기 위해 원본 순수 행렬 구조로 정직하게 판독
    raw_excel = pd.read_excel(SAVED_EXCEL_PATH, header=None)
    clean_data_list = []
    
    for idx in range(len(raw_excel)):
        row = raw_excel.iloc[idx]
        if len(row) < 21: continue
        
        # U열(20) 생산일자 날짜 형식 유효성 확인 검증
        p_date = pd.to_datetime(row[20], errors='coerce')
        if pd.isna(p_date): continue
        
        # [🚨 대표님 지정 오더 알파벳 열 1:1 다이렉트 고정 맵핑]
        # A=0(코드), C=2(카테고리), F=5(가격표), K=10(PO#), L=11(Bag#), M=12(용량), O=14(품목명), Q=16(수량)
        clean_data_list.append({
            'item_code': str(row[0]).strip(),       # A열
            'category': str(row[2]).strip(),        # C열
            'price_tag': str(row[5]).strip(),       # F열
            'po_number': str(row[10]).strip(),      # K열
            'bag_number': str(row[11]).strip(),     # L열
            'volume': str(row[12]).strip(),         # M열
            'product_name': str(row[14]).strip(),   # O열
            'quantity': pd.to_numeric(row[16], errors='coerce') if not pd.isna(row[16]) else 0, # Q열
            'production_date': p_date
        })
    
    df = pd.DataFrame(clean_data_list)
    df['quantity'] = df['quantity'].fillna(0).astype(int)
    
    # 🚨 [공백 데이터 왜곡 차단]: 공백이거나 nan 값일 경우 대표님 지시대로 '-'로 깔끔하게 치환
    for col in ['item_code', 'category', 'price_tag', 'po_number', 'bag_number', 'volume', 'product_name']:
        df[col] = df[col].replace(['nan', 'NAN', 'NaN', 'None', ''], '-')
        
    # 주차 내 카테고리별 정렬 마감
    df = df.sort_values(by=['category', 'item_code', 'production_date'])
    
    # 주차 범위 계산
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    w1_end = (today + timedelta(days=(7-today.weekday())%7 if (7-today.weekday())%7 !=0 else 7)).replace(hour=23, minute=59)
    w2_start = w1_end + timedelta(seconds=1)
    w2_end = (w2_start + timedelta(days=6)).replace(hour=23, minute=59)
    
    df_w1 = df[(df['production_date'] >= today) & (df['production_date'] <= w1_end)]
    df_w2 = df[(df['production_date'] >= w2_start) & (df['production_date'] <= w2_end)]
    
    saved_notes = load_production_notes()

    # [📊 주차별 분리형 마스터 엑셀 컴파일러 다운로드 엔진]
    def generate_premium_split_excel(df_w1, df_w2):
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "주차별_생산라인업"
        ws.views.sheetView[0].showGridLines = True
        
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
        
        headers = ["카테고리 그룹", "아이템 사진", "아이템 코드", "아이템 이름", "용량", "생산 수량", "PO 번호", "Bag#", "가격표 유무", "특기사항 1", "특기사항 2"]
        categories_order = ["skin", "body", "hair", "기타 카테고리"]
        current_row_idx = 1
        
        def write_week_block(ws, target_df, week_label_text, start_row):
            r_idx = start_row
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=11)
            title_cell = ws.cell(row=r_idx, column=1, value=week_label_text)
            title_cell.font = font_main_title; title_cell.fill = fill_week_title; title_cell.alignment = align_center
            ws.row_dimensions[r_idx].height = 35
            r_idx += 1
            
            for col_num, h_text in enumerate(headers, 1):
                h_cell = ws.cell(row=r_idx, column=col_num, value=h_text)
                h_cell.font = font_header; h_cell.fill = fill_header; h_cell.alignment = align_center; h_cell.border = border_all
            ws.row_dimensions[r_idx].height = 25
            r_idx += 1
            
            for cate in categories_order:
                cate_df = target_df[target_df['category'].str.lower().str.contains(cate)] if cate != "기타 카테고리" else target_df[~target_df['category'].str.lower().str.contains('skin|body|hair')]
                if not cate_df.empty:
                    ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=11)
                    g_cell = ws.cell(row=r_idx, column=1, value=f"🌿 {cate.upper()} CARE LINEUP")
                    g_cell.font = font_group; g_cell.fill = fill_group; g_cell.alignment = align_left
                    for c_num in range(1, 12): ws.cell(row=r_idx, column=c_num).border = border_all
                    ws.row_dimensions[r_idx].height = 24
                    r_idx += 1
                    
                    for _, r in cate_df.iterrows():
                        p_code = extract_pure_6_code(r['item_code'])
                        memo_vals = saved_notes.get(p_code, ("", ""))
                        ws.cell(row=r_idx, column=1, value=r['category'])
                        ws.cell(row=r_idx, column=3, value=r['item_code'])
                        ws.cell(row=r_idx, column=4, value=r['product_name'])
                        ws.cell(row=r_idx, column=5, value=r['volume'])
                        qty_c = ws.cell(row=r_idx, column=6, value=r['quantity']); qty_c.number_format = '#,##0'; qty_c.alignment = align_right
                        ws.cell(row=r_idx, column=7, value=r['po_number'])
                        ws.cell(row=r_idx, column=8, value=r['bag_number'])
                        ws.cell(row=r_idx, column=9, value=r['price_tag']) # 엑셀에는 F열 연동 수록
                        ws.cell(row=r_idx, column=10, value=memo_vals[0]).alignment = align_left
                        ws.cell(row=r_idx, column=11, value=memo_vals[1]).alignment = align_left
                        
                        for c_idx in range(1, 12):
                            c_cell = ws.cell(row=r_idx, column=c_idx); c_cell.font = font_data; c_cell.border = border_all
                            if c_idx not in [4, 6, 10, 11]: c_cell.alignment = align_center
                        ws.row_dimensions[r_idx].height = 35
                        img_path = f"{p_code}.png"
                        if os.path.exists(img_path):
                            try:
                                pil_img = PILImage.open(img_path); pil_img.thumbnail((50, 45))
                                img_stream = io.BytesIO(); pil_img.save(img_stream, format="PNG"); img_stream.seek(0)
                                xl_img = OpenpyxlImage(img_stream); ws.add_image(xl_img, f"B{r_idx}")
                            except: pass
                        r_idx += 1
            return r_idx + 2

        next_start = write_week_block(ws, df_w1, f"🗓️ 1주 차 계획 수립 명세서 ({today.strftime('%m/%d')} ~ {w1_end.strftime('%m/%d')})", current_row_idx)
        write_week_block(ws, df_w2, f"🗓️ 2주 차 계획 수립 명세서 ({w2_start.strftime('%m/%d')} ~ {w2_end.strftime('%m/%d')})", next_start)
        
        for col_letter, col_width in [('A', 15), ('B', 12), ('C', 16), ('D', 38), ('E', 12), ('F', 14), ('G', 16), ('H', 14), ('I', 14), ('J', 25), ('K', 25)]:
            ws.column_dimensions[col_letter].width = col_width
        wb.save(output)
        return output.getvalue()

    with st.sidebar:
        st.markdown("---")
        st.download_button(label="📊 주차별 분리 마스터 엑셀 다운로드", data=generate_premium_split_excel(df_w1, df_w2), file_name=f"Fine_Formulation_Fixed_Schedule_{datetime.now().strftime('%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # 프론트엔드 스타일 레이아웃 고정
    st.markdown("""<style>
        .owner-square-frame { width: 100% !important; aspect-ratio: 1 / 1 !important; background-color: transparent !important; display: flex !important; justify-content: center !important; align-items: center !important; overflow: hidden !important; padding: 5px !important; box-sizing: border-box !important; margin-bottom: 8px !important; }
        .owner-square-frame img { max-width: 100% !important; max-height: 100% !important; width: auto !important; height: auto !important; object-fit: contain !important; }
        .owner-info-card-wrap { background-color: #1e2530 !important; border: 1px solid #2d3748 !important; border-radius: 14px !important; padding: 18px !important; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.4) !important; margin-bottom: 8px !important; }
        .owner-text-row { margin: 0px !important; padding: 0px !important; text-align: left !important; line-height: 1.4 !important; }
        div[data-testid="stTextInput"] { margin-top: 4px !important; padding: 0px !important; }
        div[data-testid="stTextInput"] input { background-color: #0f172a !important; color: #38bdf8 !important; border: 1px solid #334155 !important; border-radius: 8px !important; font-size: 13px !important; height: 36px !important; }
    </style>""", unsafe_allow_html=True)

    def render_schedule_grid(target_df, title_label, section_prefix):
        st.markdown("---")
        st.subheader(title_label)
        
        if not target_df.empty:
            for cate in ["skin", "body", "hair"]:
                group_df = target_df[target_df['category'].str.lower().str.contains(cate)]
                if not group_df.empty:
                    st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#38bdf8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #38bdf8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 {cate.upper()} care Lineup</div>', unsafe_allow_html=True)
                    cols = st.columns(6)
                    for idx, row in group_df.reset_index(drop=True).iterrows():
                        excel_code = row['item_code']
                        pure_excel_code = extract_pure_6_code(excel_code)
                        
                        with cols[idx % 6]:
                            local_base64_data = get_saved_local_image_base64(pure_excel_code)
                            st.html(f'<div class="owner-square-frame"><img src="{local_base64_data if local_base64_data else ""}"></div>')
                            
                            # [🚨 대표님 핵심 요구사항 일치화 마감]: 가격표(F열), 용량(M열), PO#(K열), Bag#(L열) 원본 위치 100% 매핑 스펙 구현
                            st.html(f"""
                                <div class="owner-info-card-wrap">
                                    <div class="owner-text-row" style="font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important;">{excel_code}</div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#a0aec0 !important; font-weight:500 !important; min-height:40px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                                    <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">가격표 유무: <span style="color:#63b3ed !important; font-weight:bold !important;">{row['price_tag']}</span></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                                    <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between; align-items:center !important;">
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

    render_schedule_grid(df_1week, "📅 1주 차 생산 스케줄 대쉬보드", "w1")
    render_schedule_grid(df_2weeks, "📅 2주 차 생산 스케줄 대쉬보드", "w2")

else:
    st.info("💡 스케줄 마스터 엑셀 파일 로드 대기중")