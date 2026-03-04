import streamlit as st
import os
import io
import zipfile
from datetime import datetime
from docxtpl import DocxTemplate

st.set_page_config(page_title="양식 자동 변환기", layout="centered")

st.title("문서 양식 자동 변환 프로그램")
st.write("고객사와 제품명을 입력하면 `templates` 폴더 내의 모든 양식이 자동으로 변환됩니다.")

# 1. UI: 사용자 입력 받기
customer_name = st.text_input("고객사명 (CUSTOMER)")
product_name = st.text_input("제품명 (PRODUCT)")

if st.button("문서 일괄 변환"):
    # 입력값 검증
    if not customer_name or not product_name:
        st.warning("고객사명과 제품명을 모두 입력해주세요.")
    else:
        # 2. 날짜 포맷팅: 현재 날짜를 "Month DD, YYYY" 형식으로 변환 (첫 글자 대문자)
        # %B는 로케일에 맞춰 영어의 경우 'December'와 같이 첫 글자가 대문자인 전체 월 이름을 반환합니다.
        current_date = datetime.now().strftime("%B %d, %Y")
        
        template_dir = "templates"
        
        # 폴더 존재 여부 확인
        if not os.path.exists(template_dir):
            st.error(f"'{template_dir}' 폴더를 찾을 수 없습니다. 깃허브 최상단 경로에 폴더가 있는지 확인해주세요.")
        else:
            # 3. 여러 파일을 하나로 묶기 위한 ZIP 버퍼 생성 (메모리 상에서 처리)
            zip_buffer = io.BytesIO()
            
            # 압축 파일 생성 시작
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                # 폴더 내 모든 파일 순회
                for filename in os.listdir(template_dir):
                    # .docx 파일만 처리 (임시 파일인 ~로 시작하는 파일 제외)
                    if filename.endswith(".docx") and not filename.startswith("~"):
                        file_path = os.path.join(template_dir, filename)
                        
                        try:
                            # 템플릿 열기 및 변환 로직
                            doc = DocxTemplate(file_path)
                            context = {
                                'DATE': current_date,
                                'CUSTOMER': customer_name,
                                'PRODUCT': product_name
                            }
                            # 문서 내 태그 렌더링 (치환)
                            doc.render(context)
                            
                            # 4. 파일명 변경: "STH"를 입력받은 제품명으로 치환
                            new_filename = filename.replace("STH", product_name)
                            
                            # 변환된 문서를 메모리에 임시 저장하여 ZIP 파일에 추가
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)
                            
                            zip_file.writestr(new_filename, doc_io.read())
                            
                        except Exception as e:
                            st.error(f"'{filename}' 변환 중 오류가 발생했습니다: {e}")
            
            # ZIP 버퍼 포인터를 처음으로 되돌림
            zip_buffer.seek(0)
            
            st.success("모든 파일이 성공적으로 변환되었습니다! 아래 버튼을 눌러 다운로드하세요.")
            
            # 다운로드 버튼 활성화
            st.download_button(
                label="📁 변환된 문서 전체 다운로드 (ZIP)",
                data=zip_buffer,
                file_name=f"Documents_{customer_name}_{product_name}.zip",
                mime="application/zip"
            )
