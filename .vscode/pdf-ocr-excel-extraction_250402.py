import os
import re
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import pandas as pd

# Poppler 경로 설정
POPPLER_PATH = r"C:\poppler\Library\bin"
if POPPLER_PATH not in os.environ["PATH"]:
    os.environ["PATH"] += os.pathsep + POPPLER_PATH

# Tesseract OCR 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# PDF에서 이미지 추출 함수
def extract_images_from_pdf(pdf_path, dpi=300):
    try:
        return convert_from_path(pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)
    except Exception as e:
        print(f"❌ PDF 변환 오류: {e}")
        return []

# OCR 및 데이터 추출 함수
def extract_text_from_image(img):
    width, height = img.size
    section_height = height // 4  # 4등분한 높이 계산
    crop_area = (0, 2 * section_height, width, 3 * section_height)  # 위에서 세 번째 영역
    cropped_img = img.crop(crop_area).convert("L")
    return pytesseract.image_to_string(cropped_img, lang='kor', config='--psm 6')

# 데이터 정규식 추출 함수
def extract_data_from_text(text):
    # "이해관계 사실확인서"로 시작하는지 확인
    if not text.startswith("이해관계 사실확인서"):
        return None
    
    name = "발급"  # 성명은 항상 "발급"
    id_number = re.search(r'(\d{6})[- ]?\d{1,7}', text)  # 주민등록번호 앞 6자리 추출
    id_number_only = id_number.group(1) if id_number else ""
    address = re.search(r'([가-힣]+시 [가-힣]+구 [가-힣0-9 ]+)', text)
    debt_amount = re.search(r'([0-9,]+원)', text)

    return {
        "성명": name,
        "주민번호 앞 6자리": id_number_only,
        "주소": address.group(1) if address else "",
        "채무금액": debt_amount.group(1) if debt_amount else ""
    }

# 실행 함수
def main():
    pdf_path = "scanned.pdf"
    images = extract_images_from_pdf(pdf_path)
    if not images:
        print("❌ PDF에서 이미지를 추출할 수 없음.")
        return

    extracted_data = []
    for i, img in enumerate(images):
        text = extract_text_from_image(img)
        data = extract_data_from_text(text)
        if data:
            extracted_data.append(data)
        else:
            print(f"페이지 {i+1}: '이해관계 사실확인서' 제목을 찾지 못했거나 추출 실패!")

    if extracted_data:
        df = pd.DataFrame(extracted_data)
        df.to_excel('extracted_data.xlsx', index=False)
        print("✅ OCR 완료! 필요한 데이터만 엑셀로 저장됨.")
    else:
        print("❌ 추출된 데이터가 없습니다.")

if __name__ == "__main__":
    main()
