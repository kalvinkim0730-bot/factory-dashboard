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
ENTRY_SECURITY_CODE = "1234"      # [보안 코드]
SESSION_TIMEOUT_SEC = 300          # [열람 유효시간 5분]

def extract_pure_6_code(text):
    if not text:
        return ""
    cleaned = str(text).replace(" ", "").replace("_", "").replace("\r", "").replace("\n", "").replace("\t", "").strip().upper()
    match = re.search(r'(\d{5}[A-Z])', cleaned)
    return match.group(1) if match else ""

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

# 특기사항 데이터 로드/저장
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

# 웹 레이아웃 설정
st.set_page_config(layout="wide", page_title="생산 스케줄 마스터 데이터 경영 대시보드")

# 5분 제한 타이머 세션
if "app_unlocked" not in st.session_state:
    st.session_state["app_unlocked"] = False
if "unlock_time" not in st.session_state:
    st.session_state["unlock_time"] = None

if st.session_state["app_unlocked"] and st.session_state["unlock_time"] is not None:
    elapsed_time = time.time() - st.session_state["unlock_time"]
    if elapsed_time > SESSION_TIMEOUT_SEC:
        st.session_state["app_unlocked"] = False
        st.session_state["unlock_time"] = None
        st.toast("⚠️ 보안을 위해 자동 잠금되었습니다.")
        time.sleep(1)
        st.rerun()
    else:
        st.session_state["unlock_time"] = time.time()

# 게이트웨이 화면
if not st.session_state["app_unlocked"]:
    st.markdown("""
        <style>
            .stApp { background-color: #0f172a !important; }
            .security-gate {
                text-align: center; margin-top: 15vh; padding: 40px;
                background-color: #1e2530; border: 2px solid #38bdf8; border-radius: 16px;
                box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.7); max-width: 500px; margin-left: auto; margin-right: auto;
            }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("""
        <div class="security-gate">
            <h1 style="color: #38bdf8; font-size: 28px; font-weight: bold; margin-bottom: 10px;">🔒 FINE FORMULATION</h1>
            <p style="color: #94a3b8; font-size: 15px; margin-bottom: 25px;">생산 스케줄 제어 구역</p>
        </div>
    """, unsafe_allow_html=True)
    
    cols = st.columns([1, 2, 1])
    with cols[1]:
        input_gate_code = st.text_input("🔑 보안 코드 입력", type="password", key="gate_code_input")
        if input_gate_code == ENTRY_SECURITY_CODE:
            st.session_state["app_unlocked"] = True
            st.session_state["unlock_time"] = time.time()
            st.success("🔓 시스템 개방")
            time.sleep(0.5)
            st.rerun()
    st.stop()

has_saved_file = os.path.exists(SAVED_EXCEL_PATH)
final_file_target = SAVED_EXCEL_PATH if os.path.exists(SAVED_EXCEL_PATH) else None

with st.sidebar:
    st.markdown('<div style="font-size:20px; font-weight:bold; color:#38bdf8; margin-bottom:15px; border-bottom:2px solid #38bdf8; padding-bottom:5px;">⚙️ 데이터 제어 센터</div>', unsafe_allow_html=True)
    input_password = st.text_input("🔓 제어 승인 암호", type="password", key="auth_pwd_input")
    is_authenticated = (input_password == MASTER_PASSWORD)

    st.markdown("---")
    uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"], label_visibility="collapsed")
    if uploaded_file and is_authenticated:
        with open(SAVED_EXCEL_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("🚀 스케줄 파일 교체 성공!")
        time.sleep(1)
        st.rerun()

if final_file_target:
    # 데이터 강제 스캔: 건너뛰기 없이 전량 로드
    raw_excel = pd.read_excel(final_file_target, header=None)
    
    clean_data_list = []
    for idx in range(len(raw_excel)):
        row_cells = raw_excel.iloc[idx]
        if len(row_cells) < 21:
            continue
            
        # 열 매핑 고정 (A=0, C=2, K=10, L=11, M=12, O=14, Q=16, U=20)
        item_code_val = str(row_cells[0]).strip()
        category_val = str(row_cells[2]).strip()
        
        # 품목 코드가 비어있거나 제목 형식이면 데이터 행이 아니므로 제외
        if item_code_val in ['nan', 'NAN', 'NaN', 'None', '', '코드', 'Item Code']:
            continue
            
        # 날짜 파싱이 안 되는 행이라도 에러로 튕구지 않고 오늘 날짜 기준 기본값 처리하여 유실 원천 차단
        p_date = pd.to_datetime(row_cells[20], errors='coerce')
        if pd.isna(p_date):
            p_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            
        clean_data_list.append({
            'item_code': item_code_val,
            'category': category_val,
            'po_number': str(row_cells[10]).strip(),
            'bag_number': str(row_cells[11]).strip(),
            'volume': str(row_cells[12]).strip(),
            'product_name': str(row_cells[14]).strip(),
            'quantity': pd.to_numeric(row_cells[16], errors='coerce') if not pd.isna(row_cells[16]) else 0,
            'production_date': p_date
        })
        
    df = pd.DataFrame(clean_data_list)
    df['quantity'] = df['quantity'].fillna(0).astype(int)
    
    for col in ['item_code', 'category', 'po_number', 'bag_number', 'volume', 'product_name']:
        df[col] = df[col].replace(['nan', 'NAN', 'NaN', 'None', ''], '-')

    df = df.sort_values(by=['category', 'item_code', 'production_date'], ascending=[True, True, True])
    
    # 주차 분할 기준 계산
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
    
    # 주차 구분에 걸리지 않은 나머지 전체 데이터가 존재할 경우 1주 차에 기본 포함하여 데이터 노출 보장
    leftover_df = df[~df.index.isin(df_1week.index) & ~df.index.isin(df_2weeks.index)].copy()
    if not leftover_df.empty:
        df_1week = pd.concat([df_1week, leftover_df]).sort_values(by=['category', 'item_code'])

    saved_notes = load_production_notes()

    # 엑셀 다운로드 파일 빌더
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
        
        headers = ["카테고리 그룹", "아이템 사진", "아이템 코드", "아이템 이름", "용량", "생산 수량", "PO 번호", "Bag#", "특기사항 1", "특기사항 2"]
        categories_order = ["skin", "body", "hair", "기타 카테고리"]
        current_row_idx = 1
        
        def write_week_block(ws, target_df, week_label_text, start_row):
            r_idx = start_row
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=10)
            title_cell = ws.cell(row=r_idx, column=1)
            title_cell.value = week_label_text
            title_cell.font = font_main_title
            title_cell.fill = fill_week_title
            title_cell.alignment = align_center
            ws.row_dimensions[r_idx].height = 35
            r_idx += 1
            
            for col_num, h_text in enumerate(headers, 1):
                h_cell = ws.cell(row=r_idx, column=col_num, value=h_text)
                h_cell.font = font_header
                h_cell.fill = fill_header
                h_cell.alignment = align_center
                h_cell.border = border_all
            ws.row_dimensions[r_idx].height = 25
            r_idx += 1
            
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
                        ws.cell(row=r_idx, column=8, value=r['bag_number'])
                        
                        ws.cell(row=r_idx, column=9, value=memo_vals[0]).alignment = align_left
                        ws.cell(row=r_idx, column=10, value=memo_vals[1]).alignment = align_left
                        
                        for c_idx in range(1, 11):
                            c_cell = ws.cell(row=r_idx, column=c_idx)
                            c_cell.font = font_data
                            c_cell.border = border_all
                            if c_idx != 4 and c_idx != 6 and c_idx != 1 and c_idx != 9 and c_idx != 10:
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
            
        next_start_row = write_week_block(ws, df_1week, "🗓️ 1주 차 생산 라인업 계획", current_row_idx)
        write_week_block(ws, df_2weeks, "🗓️ 2주 차 생산 라인업 계획", next_start_row)
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 38
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 14
        ws.column_dimensions['G'].width = 16
        ws.column_dimensions['H'].width = 14
        ws.column_dimensions['I'].width = 25
        ws.column_dimensions['J'].width = 25
        
        wb.save(output)
        return output.getvalue()

    with st.sidebar:
        st.markdown("---")
        split_excel_bytes = generate_premium_split_excel(df_1week, df_2weeks)
        st.download_button(
            label="📊 주차별 분리 마스터 엑셀 다운로드",
            data=split_excel_bytes,
            file_name=f"Fine_Formulation_Split_Schedule_{datetime.now().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # 대시보드 화면 스타일 및 그리드 구성
    st.markdown("""
        <style>
            .owner-square-frame {
                width: 100% !important; aspect-ratio: 1 / 1 !important;
                display: flex !important; justify-content: center !important; align-items: center !important;
                overflow: hidden !important; padding: 5px !important; box-sizing: border-box !important; margin-bottom: 8px !important;
            }
            .owner-square-frame img { max-width: 100% !important; max-height: 100% !important; object-fit: contain !important; }
            .owner-info-card-wrap {
                background-color: #1e2530 !important; border: 1px solid #2d3748 !important; border-radius: 14px !important; 
                padding: 18px !important; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.4) !important; margin-bottom: 8px !important;
            }
            .owner-text-row { margin: 0px !important; padding: 0px !important; text-align: left !important; line-height: 1.4 !important; }
            div[data-testid="stTextInput"] { margin-top: 4px !important; padding: 0px !important; }
            div[data-testid="stTextInput"] input { background-color: #0f172a !important; color: #38bdf8 !important; border: 1px solid #334155 !important; border-radius: 8px !important; font-size: 13px !important; height: 36px !important; }
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
                    for idx, row in group_df.reset_index().iterrows():
                        excel_code = row['item_code']
                        pure_excel_code = extract_pure_6_code(excel_code)
                        
                        with cols[idx % 6]:
                            local_base64_data = get_saved_local_image_base64(pure_excel_code)
                            if local_base64_data:
                                st.html(f'<div class="owner-square-frame"><img src="{local_base64_data}"></div>')
                            else:
                                st.html(f'<div class="owner-square-frame"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center;">{excel_code}</div></div>')
                            
                            st.html(f"""
                                <div class="owner-info-card-wrap">
                                    <div class="owner-text-row" style="font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important;">{excel_code}</div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#a0aec0 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                                    <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                                    <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                        <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                                    </div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                                    <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                                    <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !important; align-items:center !important;">
                                        <span class="owner-text-row" style="font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                        <span class="owner-text-row" style="font-size:13px !important; color:#a0aec0 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                                    </div>
                                </div>
                            """)
                            
                            memo_tuple = saved_notes.get(pure_excel_code, ("", ""))
                            key_m1 = f"input_m1_{section_prefix}_{pure_excel_code}_{idx}"
                            user_m1 = st.text_input(label=f"T1_{pure_excel_code}", value=memo_tuple[0], key=key_m1, placeholder="📋 특기사항 1", label_visibility="collapsed")
                            key_m2 = f"input_m2_{section_prefix}_{pure_excel_code}_{idx}"
                            user_m2 = st.text_input(label=f"T2_{pure_excel_code}", value=memo_tuple[1], key=key_m2, placeholder="📦 특기사항 2", label_visibility="collapsed")
                            
                            if user_m1 != memo_tuple[0] or user_m2 != memo_tuple[1]:
                                save_production_note(pure_excel_code, user_m1, user_m2)
                                st.rerun()
                            st.markdown('<div style="margin-bottom:30px;"></div>', unsafe_allow_html=True)
            
            other_df = target_df[~target_df['category'].str.lower().str.contains('skin|body|hair')]
            if not other_df.empty:
                st.markdown('<div style="font-size:20px; font-weight:bold; color:#94a3b8; padding:6px 12px; background-color:#0f172a; border-left:5px solid #94a3b8; border-radius:4px; margin-top:25px; margin-bottom:15px;">📦 기타 카테고리 Lineup</div>', unsafe_allow_html=True)
                cols = st.columns(6)
                for idx, row in other_df.reset_index().iterrows():
                    excel_code = row['item_code']
                    pure_excel_code = extract_pure_6_code(excel_code)
                    with cols[idx % 6]:
                        local_base64_data = get_saved_local_image_base64(pure_excel_code)
                        if local_base64_data:
                            st.html(f'<div class="owner-square-frame"><img src="{local_base64_data}"></div>')
                        else:
                            st.html(f'<div class="owner-square-frame"><div style="color:#f87171; font-size:13px; font-weight:bold; text-align:center;">{excel_code}</div></div>')
                        
                        st.html(f"""
                            <div class="owner-info-card-wrap">
                                <div class="owner-text-row" style="font-size:30px !important; font-weight:900 !important; color:#ffffff !important; margin-bottom:6px !important;">{excel_code}</div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#a0aec0 !important; min-height:40px !important; margin-bottom:14px !important; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">{row['product_name']}</div>
                                <div style="border-bottom:1px solid #2d3748 !important; margin-bottom:12px !important;"></div>
                                <div style="display:flex !important; justify-content:space-between !important; margin-bottom:5px !important;">
                                    <span class="owner-text-row" style="font-size:14px !important; color:#718096 !important;">용량: <span style="color:#ffffff !important; font-weight:bold !important;">{row['volume']}</span></span>
                                </div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:3px !important;">PO#: <span style="color:#ecc94b !important; font-weight:bold !important;">{row['po_number']}</span></div>
                                <div class="owner-text-row" style="font-size:14px !important; color:#718096 !important; margin-bottom:16px !important;">Bag#: <span style="color:#e53e3e !important; font-weight:bold !important;">{row['bag_number']}</span></div>
                                <div style="background-color:#111622 !important; border-radius:8px !important; padding:8px 12px !important; display:flex !important; justify-content:space-between !important; align-items:center !important;">
                                    <span class="owner-text-row" style="font-size:16px !important; color:#48bb78 !important; font-weight:bold !important;">📦 {row['quantity']:,}개</span>
                                    <span class="owner-text-row" style="font-size:13px !important; color:#a0aec0 !important;">📅 {row['production_date'].strftime('%m-%d')}</span>
                                </div>
                            </div>
                        """)
                        memo_tuple = saved_notes.get(pure_excel_code, ("", ""))
                        key_m1 = f"input_m1_oth_{pure_excel_code}_{idx}"
                        user_m1 = st.text_input(label=f"T1_{pure_excel_code}_oth", value=memo_tuple[0], key=key_m1, placeholder="📋 특기사항 1", label_visibility="collapsed")
                        key_m2 = f"input_m2_oth_{pure_excel_code}_{idx}"
                        user_m2 = st.text_input(label=f"T2_{pure_excel_code}_oth", value=memo_tuple[1], key=key_m2, placeholder="📦 특기사항 2", label_visibility="collapsed")
                        if user_m1 != memo_tuple[0] or user_m2 != memo_tuple[1]:
                            save_production_note(pure_excel_code, user_m1, user_m2)
                            st.rerun()
                        st.markdown('<div style="margin-bottom:30px;"></div>', unsafe_allow_html=True)

    render_schedule_grid(df_1week, "📅 1주 차 생산 스케줄 대쉬보드", "w1")
    render_schedule_grid(df_2weeks, "📅 2주 차 생산 스케줄 대쉬보드", "w2")
else:
    st.info("💡 마스터 엑셀 파일 로드 대기중")