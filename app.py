import os
import re
import io
import pandas as pd
import pdfplumber
import streamlit as st
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# --- 추출 로직 (기존과 동일) ---
def calculate_validity_period(cert_date_str):
    try:
        clean_str = cert_date_str.replace(" ", "")
        if clean_str.endswith("."): clean_str = clean_str[:-1]
        parts = clean_str.split(".")
        if len(parts) != 3: return cert_date_str, ""
        
        start_date = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
        end_date = start_date + relativedelta(years=5) - timedelta(days=1)
        return f"{start_date.year}. {start_date.month}. {start_date.day}." , f"{end_date.year}. {end_date.month}. {end_date.day}."
    except: 
        return cert_date_str, ""

def extract_data_from_pdf(pdf_file_buffer):
    extracted = {}
    try:
        with pdfplumber.open(pdf_file_buffer) as pdf:
            page = pdf.pages[0]
            text_full = page.extract_text()
            patterns = {
                "제조자": r"제조[자사]\s*[:]?\s*([^\n]+)",
                "제조국가": r"제조국가\s*[:]?\s*([^\n]+)",
                "인증연월일": r"인증\s*연월일\s*[:]?\s*([^\n]+)",
                "인증번호": r"인증\s*번호\s*[:]?\s*([A-Za-z0-9\-]+)", 
                "인증범위": r"인증[ ]?범위\s*[:]?\s*([^\n]+)" 
            }
            for key, pattern in patterns.items():
                match = re.search(pattern, text_full if text_full else "")
                if match: extracted[key] = re.sub(r'\s{2,}.*', '', match.group(1)).strip()

            width, height = page.width, page.height
            right_bbox = (width * 0.32, 0, width, height) 
            right_page = page.within_bbox(right_bbox)
            right_text = right_page.extract_text(layout=True) 
            
            if right_text:
                match = re.search(r'Test\s*Highlights[^\n]*\n(.*)', right_text, re.DOTALL | re.IGNORECASE)
                if match:
                    raw_highlights = match.group(1)
                    lines = raw_highlights.split('\n')
                    valid_lines = []
                    found_bullets = False
                    empty_line_buffer = False
                    
                    for line in lines:
                        stripped = line.strip()
                        if re.match(r'^(CONTENTS|ABOUT TTA|TTA)', stripped, re.IGNORECASE): continue
                        if not stripped:
                            if found_bullets: empty_line_buffer = True
                            continue
                            
                        is_symbol = re.match(r'^([⚫●\-∎■▪\u2022\u25E6\u2023\u2043\u2219])', stripped)
                        is_char = re.match(r'^[lOㅇo]\s', stripped) 
                        
                        if is_symbol or is_char:
                            found_bullets = True
                            empty_line_buffer = False
                            if is_char and stripped.startswith('l '): stripped = '● ' + stripped[2:]
                            valid_lines.append(stripped)
                        elif found_bullets:
                            if empty_line_buffer: break
                            stripped = re.sub(r'\s+(개\s*요|개요)$', '', stripped)
                            valid_lines[-1] += f" {stripped}"
                            
                    extracted['Test Highlights'] = '\n'.join(valid_lines)
                else: extracted['Test Highlights'] = ""
    except Exception as e:
        st.error(f"PDF 읽기 오류: {e}")
    return extracted

# --- 웹 화면(UI) 구성 ---
st.set_page_config(page_title="인증서 추출기", page_icon="📝", layout="centered")

st.title("📄 인증서 데이터 자동 추출기")
st.markdown("엑셀 목록과 PDF 요약서를 올리면 자동으로 매칭하여 최종 엑셀을 만들어줍니다.")

# 파일 업로드 구역
st.subheader("1. 파일 업로드")
excel_file = st.file_uploader("엑셀 파일 (인증 제품 관리 목록) 업로드", type=['xlsx', 'xls'])
pdf_files = st.file_uploader("시험결과요약서 PDF 파일 업로드 (여러 개 동시 선택 가능)", type=['pdf'], accept_multiple_files=True)

# 실행 버튼 구역
if st.button("데이터 매칭 및 추출 시작 🚀", type="primary"):
    if not excel_file:
        st.warning("엑셀 파일을 먼저 업로드해주세요.")
    elif not pdf_files:
        st.warning("PDF 파일을 한 개 이상 업로드해주세요.")
    else:
        with st.spinner('데이터를 처리하는 중입니다. 잠시만 기다려주세요...'):
            # 1. 엑셀 데이터 로드
            all_sheets = pd.read_excel(excel_file, sheet_name=None, header=1) 
            df_list = [df for sheet_name, df in all_sheets.items() if "TTA" in sheet_name]
            excel_data = pd.concat(df_list, ignore_index=True).fillna("") if df_list else pd.DataFrame()
            
            if excel_data.empty:
                st.error("엑셀 파일에서 'TTA'가 포함된 시트를 찾지 못했습니다.")
                st.stop()

            final_results = []
            
            # 2. PDF 순회 및 추출
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, pdf_file in enumerate(pdf_files):
                status_text.text(f"처리 중: {pdf_file.name} ({i+1}/{len(pdf_files)})")
                
                pdf_info = extract_data_from_pdf(pdf_file)
                pdf_cert_num = pdf_info.get("인증번호", "").replace(" ", "")
                
                if not pdf_cert_num:
                    st.warning(f"[스킵] {pdf_file.name}: 인증번호를 찾을 수 없습니다.")
                    continue
                    
                excel_cert_series = excel_data['인증번호'].astype(str).str.replace(" ", "")
                matched_row = excel_data[excel_cert_series == pdf_cert_num]
                
                if not matched_row.empty:
                    row_data = matched_row.iloc[0]
                    cert_date = pdf_info.get("인증연월일", "")
                    valid_start, valid_end = calculate_validity_period(cert_date) 
                    
                    manu = pdf_info.get("제조자", "").strip()
                    country = pdf_info.get("제조국가", "").strip()
                    if not country: country = "대한민국"
                    manu_and_country = f"{manu} / {country}" if manu else country
                    
                    final_results.append({
                        "업체명": row_data.get("업체명", ""),
                        "영문명": row_data.get("영문명", ""),
                        "시험번호": row_data.get("시험번호", ""),
                        "인증번호": row_data.get("인증번호", ""), 
                        "제품명": row_data.get("제품명", ""),
                        "모델명": row_data.get("모델명", ""),
                        "제조자 및 제조국가": manu_and_country,
                        "인증연월일": cert_date,
                        "유효기간(시작)": valid_start,
                        "유효기간(종료)": valid_end,
                        "인증기준": row_data.get("인증기준", ""),
                        "인증범위": pdf_info.get("인증범위", ""), 
                        "Test Highlights": pdf_info.get("Test Highlights", "")
                    })
                
                progress_bar.progress((i + 1) / len(pdf_files))
            
            status_text.text("모든 처리 완료!")
            
            # 3. 결과 다운로드 버튼 생성
            if final_results:
                result_df = pd.DataFrame(final_results)
                result_df = result_df.sort_values(by="인증번호", ascending=True)
                
                # 메모리에 엑셀 파일 만들기
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                st.success(f"🎉 총 {len(final_results)}건의 데이터 매칭 및 추출에 성공했습니다!")
                
                st.download_button(
                    label="📥 최종 결과 엑셀 다운로드",
                    data=excel_data,
                    file_name=f"최종_인증정보_추출결과_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("매칭된 데이터가 한 건도 없습니다.")