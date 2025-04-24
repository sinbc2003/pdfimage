import streamlit as st
import requests
import base64
import os
import tempfile
import fitz  # PyMuPDF
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image
import io
import time
import pytesseract
import cv2
import numpy as np

# 페이지 설정
st.set_page_config(page_title="PDF/이미지 텍스트 추출 도구", page_icon="📄", layout="wide")

# 탭 설정
tab1, tab2 = st.tabs(["PDF 텍스트 추출", "이미지 OCR"])

# 사이드바 - 공통 설정
with st.sidebar:
    st.header("설정")
    api_key = st.text_input("OpenAI API 키", type="password")
    google_sheet_id = st.text_input("Google 스프레드시트 ID (선택사항)", 
                                   help="결과를 저장할 Google 스프레드시트 ID")
    use_google_sheets = st.checkbox("Google 스프레드시트에 결과 저장", value=False)
    
    st.markdown("---")
    st.markdown("### OCR 설정")
    ocr_method = st.radio(
        "OCR 방법 선택",
        ["OpenAI (o4-mini)", "Tesseract OCR", "둘 다 사용"]
    )
    
    st.markdown("---")
    st.markdown("### 참고")
    st.info("업로드한 파일은 서버에 저장되지 않으며, 처리 후 자동으로 삭제됩니다.")

# Google 스프레드시트 연결 함수
def connect_to_google_sheets(sheet_id):
    try:
        # 서비스 계정 정보 (환경 변수 또는 secrets에서 가져옴)
        # 실제 배포 시에는 Streamlit의 secrets 관리 기능을 사용하는 것이 좋습니다
        if os.path.exists('service_account.json'):
            # 로컬 개발 환경
            creds = Credentials.from_service_account_file(
                'service_account.json',
                scopes=['https://www.googleapis.com/auth/spreadsheets',
                       'https://www.googleapis.com/auth/drive']
            )
        else:
            # Streamlit Cloud 환경
            service_account_info = st.secrets["gcp_service_account"]
            creds = Credentials.from_service_account_info(
                service_account_info,
                scopes=['https://www.googleapis.com/auth/spreadsheets',
                       'https://www.googleapis.com/auth/drive']
            )
        
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id).sheet1
        return sheet
    except Exception as e:
        st.error(f"Google 스프레드시트 연결 오류: {str(e)}")
        return None

# PDF 페이지를 이미지로 변환하는 함수
def convert_pdf_page_to_image(pdf_file, page_num):
    try:
        # PyMuPDF를 사용하여 PDF 페이지를 이미지로 변환
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        pdf_file.seek(0)  # 파일 포인터 리셋
        
        if page_num >= doc.page_count:
            return None
        
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))  # 300 DPI로 렌더링
        
        img_bytes = pix.tobytes("png")
        return img_bytes
    except Exception as e:
        st.error(f"PDF 페이지 변환 오류: {str(e)}")
        return None

# OpenAI API를 사용하여 이미지에서 텍스트 추출 함수
def extract_text_with_openai(image_bytes, page_num, api_key):
    try:
        # 이미지를 Base64로 인코딩
        base64_image = base64.b64encode(image_bytes).decode('utf-8')
        
        # OpenAI API 요청 설정
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        # 요청 페이로드 구성 - Chat Completions API 사용
        payload = {
            "model": "o4-mini",
            "messages": [
                {
                    "role": "system",
                    "content": "수식은 LaTeX형식으로 제공해줘. 이미지의 모든 텍스트 내용을 추출해서 원본 서식을 최대한 유지하며 보여줘."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": f"이 이미지에서 모든 텍스트를 추출해줘."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}",
                                "detail": "high"
                            }
                        }
                    ]
                }
            ]
        }
        
        # API 호출
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload
        )
        
        # 응답 확인
        if response.status_code == 200:
            result = response.json()
            if "choices" in result and len(result["choices"]) > 0:
                extracted_text = result["choices"][0]["message"]["content"]
                return extracted_text
            else:
                return "OpenAI로 텍스트를 추출하지 못했습니다."
        else:
            st.error(f"API 오류 ({response.status_code}): {response.text}")
            return f"OpenAI API 오류가 발생했습니다."
    except Exception as e:
        st.error(f"OpenAI 텍스트 추출 오류: {str(e)}")
        return f"OpenAI 텍스트 추출 중 오류 발생: {str(e)}"

# Tesseract OCR을 사용하여 이미지에서 텍스트 추출
def extract_text_with_tesseract(image_bytes):
    try:
        # 이미지 바이트를 NumPy 배열로 변환
        nparr = np.frombuffer(image_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        
        # OpenCV로 이미지 전처리
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # 노이즈 제거
        gray = cv2.medianBlur(gray, 3)
        # 이진화
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Tesseract OCR로 텍스트 추출
        text = pytesseract.image_to_string(thresh, lang='kor+eng')
        
        return text if text.strip() else "Tesseract OCR로 텍스트를 추출하지 못했습니다."
    except Exception as e:
        st.error(f"Tesseract OCR 오류: {str(e)}")
        return f"Tesseract OCR 중 오류 발생: {str(e)}"

# PDF 처리 함수
def process_pdf(pdf_file, api_key, sheet=None):
    try:
        # PDF 파일 정보 확인
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        pdf_file.seek(0)  # 파일 포인터 리셋
        page_count = doc.page_count
        
        st.info(f"PDF 파일: {pdf_file.name}, 총 {page_count} 페이지")
        
        # 진행 상황 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 추출 결과를 저장할 리스트
        extracted_texts = []
        
        # 각 페이지 처리
        for i in range(page_count):
            # 상태 업데이트
            status_text.text(f"페이지 {i+1}/{page_count} 처리 중...")
            progress_bar.progress((i) / page_count)
            
            # 페이지를 이미지로 변환
            img_bytes = convert_pdf_page_to_image(pdf_file, i)
            if img_bytes:
                # 선택한 OCR 방법에 따라 텍스트 추출
                if ocr_method == "OpenAI (o4-mini)":
                    extracted_text = extract_text_with_openai(img_bytes, i, api_key)
                    extracted_texts.append({
                        "페이지": i + 1,
                        "OpenAI 추출 텍스트": extracted_text
                    })
                elif ocr_method == "Tesseract OCR":
                    extracted_text = extract_text_with_tesseract(img_bytes)
                    extracted_texts.append({
                        "페이지": i + 1,
                        "Tesseract 추출 텍스트": extracted_text
                    })
                else:  # 둘 다 사용
                    openai_text = extract_text_with_openai(img_bytes, i, api_key)
                    tesseract_text = extract_text_with_tesseract(img_bytes)
                    extracted_texts.append({
                        "페이지": i + 1,
                        "OpenAI 추출 텍스트": openai_text,
                        "Tesseract 추출 텍스트": tesseract_text
                    })
                
                # API 호출 간 짧은 대기 시간
                time.sleep(1)
            else:
                extracted_texts.append({
                    "페이지": i + 1,
                    "OpenAI 추출 텍스트": "페이지 변환 실패",
                    "Tesseract 추출 텍스트": "페이지 변환 실패"
                })
        
        # 진행 상황 완료
        progress_bar.progress(1.0)
        status_text.text("처리 완료!")
        
        # 결과를 데이터프레임으로 변환
        df = pd.DataFrame(extracted_texts)
        
        # Google 스프레드시트에 결과 저장 (선택 사항)
        if sheet:
            try:
                # 기존 내용 지우기
                sheet.clear()
                
                # 헤더 추가
                if ocr_method == "OpenAI (o4-mini)":
                    sheet.update("A1:B1", [["페이지", "OpenAI 추출 텍스트"]])
                    # 각 페이지 텍스트 추가
                    data_to_insert = [[row["페이지"], row["OpenAI 추출 텍스트"]] for row in extracted_texts]
                elif ocr_method == "Tesseract OCR":
                    sheet.update("A1:B1", [["페이지", "Tesseract 추출 텍스트"]])
                    # 각 페이지 텍스트 추가
                    data_to_insert = [[row["페이지"], row["Tesseract 추출 텍스트"]] for row in extracted_texts]
                else:
                    sheet.update("A1:C1", [["페이지", "OpenAI 추출 텍스트", "Tesseract 추출 텍스트"]])
                    # 각 페이지 텍스트 추가
                    data_to_insert = [[row["페이지"], row["OpenAI 추출 텍스트"], row["Tesseract 추출 텍스트"]] for row in extracted_texts]
                
                if data_to_insert:
                    if ocr_method == "둘 다 사용":
                        sheet.update(f"A2:C{len(data_to_insert)+1}", data_to_insert)
                    else:
                        sheet.update(f"A2:B{len(data_to_insert)+1}", data_to_insert)
                
                # A1 셀에 바로 첫 번째 페이지 내용 넣기
                if len(extracted_texts) > 0:
                    if ocr_method == "OpenAI (o4-mini)":
                        sheet.update("A1", extracted_texts[0]["OpenAI 추출 텍스트"])
                    elif ocr_method == "Tesseract OCR":
                        sheet.update("A1", extracted_texts[0]["Tesseract 추출 텍스트"])
                    else:
                        sheet.update("A1", extracted_texts[0]["OpenAI 추출 텍스트"])
                
                st.success(f"결과가 Google 스프레드시트에 저장되었습니다. ID: {google_sheet_id}")
            except Exception as e:
                st.error(f"스프레드시트 저장 오류: {str(e)}")
        
        return df
    except Exception as e:
        st.error(f"PDF 처리 오류: {str(e)}")
        return None

# 이미지 파일 처리 함수
def process_image(image_file, api_key, sheet=None):
    try:
        # 이미지 파일 읽기
        image_bytes = image_file.getvalue()
        
        # 이미지 표시
        img = Image.open(io.BytesIO(image_bytes))
        st.image(img, caption="업로드된 이미지", use_column_width=True)
        
        # 진행 상황 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("이미지 처리 중...")
        progress_bar.progress(0.3)
        
        # 추출 결과를 저장할 딕셔너리
        results = {}
        
        # 선택한 OCR 방법에 따라 텍스트 추출
        if ocr_method == "OpenAI (o4-mini)" or ocr_method == "둘 다 사용":
            progress_bar.progress(0.4)
            status_text.text("OpenAI로 텍스트 추출 중...")
            openai_text = extract_text_with_openai(image_bytes, 0, api_key)
            results["OpenAI 추출 텍스트"] = openai_text
            progress_bar.progress(0.7)
        
        if ocr_method == "Tesseract OCR" or ocr_method == "둘 다 사용":
            progress_bar.progress(0.8)
            status_text.text("Tesseract OCR로 텍스트 추출 중...")
            tesseract_text = extract_text_with_tesseract(image_bytes)
            results["Tesseract 추출 텍스트"] = tesseract_text
        
        # 진행 상황 완료
        progress_bar.progress(1.0)
        status_text.text("처리 완료!")
        
        # Google 스프레드시트에 결과 저장 (선택 사항)
        if sheet:
            try:
                # 기존 내용 지우기
                sheet.clear()
                
                # 결과 저장
                if ocr_method == "OpenAI (o4-mini)":
                    sheet.update("A1", "OpenAI 추출 텍스트")
                    sheet.update("A2", results["OpenAI 추출 텍스트"])
                elif ocr_method == "Tesseract OCR":
                    sheet.update("A1", "Tesseract 추출 텍스트")
                    sheet.update("A2", results["Tesseract 추출 텍스트"])
                else:
                    sheet.update("A1:B1", [["OpenAI 추출 텍스트", "Tesseract 추출 텍스트"]])
                    sheet.update("A2:B2", [[results["OpenAI 추출 텍스트"], results["Tesseract 추출 텍스트"]]])
                
                # A1 셀에 OpenAI 결과 또는 Tesseract 결과 중 하나 저장
                if "OpenAI 추출 텍스트" in results:
                    sheet.update("A1", results["OpenAI 추출 텍스트"])
                elif "Tesseract 추출 텍스트" in results:
                    sheet.update("A1", results["Tesseract 추출 텍스트"])
                
                st.success(f"결과가 Google 스프레드시트에 저장되었습니다. ID: {google_sheet_id}")
            except Exception as e:
                st.error(f"스프레드시트 저장 오류: {str(e)}")
        
        return results
    except Exception as e:
        st.error(f"이미지 처리 오류: {str(e)}")
        return None

# PDF 텍스트 추출 탭
with tab1:
    st.title("📄 PDF 텍스트 추출 도구")
    st.markdown("PDF 파일을 업로드하면 o4-mini 모델을 사용하여 텍스트를 추출합니다.")
    
    uploaded_pdf = st.file_uploader("PDF 파일 업로드", type=["pdf"], key="pdf_uploader")

    if uploaded_pdf is not None:
        if not api_key and ocr_method != "Tesseract OCR":
            st.warning("OpenAI API 키를 입력해주세요.")
        else:
            # 처리 버튼
            if st.button("PDF 텍스트 추출 시작", key="pdf_button"):
                sheet = None
                # Google 스프레드시트 연결 (선택 사항)
                if use_google_sheets and google_sheet_id:
                    sheet = connect_to_google_sheets(google_sheet_id)
                
                # PDF 처리
                with st.spinner("PDF 처리 중..."):
                    result_df = process_pdf(uploaded_pdf, api_key, sheet)
                    
                    if result_df is not None:
                        # 결과 표시
                        st.subheader("추출된 텍스트")
                        
                        # 탭 생성
                        result_tab1, result_tab2 = st.tabs(["텍스트 보기", "표 형식으로 보기"])
                        
                        with result_tab1:
                            for _, row in result_df.iterrows():
                                st.markdown(f"### 페이지 {row['페이지']}")
                                
                                # OCR 방법에 따라 다른 결과 표시
                                if ocr_method == "OpenAI (o4-mini)":
                                    st.text_area(f"OpenAI 추출 텍스트", row['OpenAI 추출 텍스트'], height=200)
                                elif ocr_method == "Tesseract OCR":
                                    st.text_area(f"Tesseract 추출 텍스트", row['Tesseract 추출 텍스트'], height=200)
                                else:
                                    st.text_area(f"OpenAI 추출 텍스트", row['OpenAI 추출 텍스트'], height=200)
                                    st.text_area(f"Tesseract 추출 텍스트", row['Tesseract 추출 텍스트'], height=200)
                        
                        with result_tab2:
                            st.dataframe(result_df)
                        
                        # CSV 다운로드 버튼
                        csv = result_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="CSV로 다운로드",
                            data=csv,
                            file_name=f"{uploaded_pdf.name}_extracted_text.csv",
                            mime="text/csv",
                        )

# 이미지 OCR 탭
with tab2:
    st.title("🖼️ 이미지 OCR 도구")
    st.markdown("이미지를 업로드하면 텍스트를 추출합니다.")
    
    uploaded_image = st.file_uploader("이미지 파일 업로드", type=["jpg", "jpeg", "png", "bmp", "tiff"], key="image_uploader")

    if uploaded_image is not None:
        if not api_key and ocr_method != "Tesseract OCR":
            st.warning("OpenAI API 키를 입력해주세요.")
        else:
            # 처리 버튼
            if st.button("이미지 텍스트 추출 시작", key="image_button"):
                sheet = None
                # Google 스프레드시트 연결 (선택 사항)
                if use_google_sheets and google_sheet_id:
                    sheet = connect_to_google_sheets(google_sheet_id)
                
                # 이미지 처리
                with st.spinner("이미지 처리 중..."):
                    results = process_image(uploaded_image, api_key, sheet)
                    
                    if results:
                        # 결과 표시
                        st.subheader("추출된 텍스트")
                        
                        # OCR 방법에 따라 다른 결과 표시
                        if "OpenAI 추출 텍스트" in results:
                            st.markdown("### OpenAI (o4-mini) 결과")
                            st.text_area("OpenAI 추출 텍스트", results["OpenAI 추출 텍스트"], height=300)
                        
                        if "Tesseract 추출 텍스트" in results:
                            st.markdown("### Tesseract OCR 결과")
                            st.text_area("Tesseract 추출 텍스트", results["Tesseract 추출 텍스트"], height=300)
                        
                        # CSV 다운로드 버튼
                        df = pd.DataFrame([results])
                        csv = df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="CSV로 다운로드",
                            data=csv,
                            file_name=f"{uploaded_image.name}_extracted_text.csv",
                            mime="text/csv",
                        )
