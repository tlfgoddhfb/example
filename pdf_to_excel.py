import re
import pdfplumber
import openpyxl

def extract_text_from_pdf(pdf_path):
    """PDF에서 텍스트를 추출합니다 (4등분 중 세 번째 부분만)."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # 페이지 높이를 기준으로 4등분
            top = page.height * 0.5  # 세 번째 부분 시작 (50%)
            bottom = page.height * 0.75  # 세 번째 부분 끝 (75%)
            cropped_page = page.within_bbox((0, top, page.width, bottom))
            text += cropped_page.extract_text()
    return text

def extract_data(text):
    """텍스트에서 필요한 데이터를 추출합니다."""
    # "이해관계 사실확인서"로 시작하는 부분만 추출
    relevant_section = re.search(r'이해관계 사실확인서.*', text, re.DOTALL)
    if not relevant_section:
        return []

    section_text = relevant_section.group(0)

    # 성명 추출
    name_match = re.search(r'성명\s*([가-힣]+)', section_text)
    name = name_match.group(1) if name_match else ""

    # 주민등록번호 추출 (######-$ 형태에서 ######만 추출)
    id_match = re.search(r'주민등록번호\s*(\d{6})-\d', section_text)
    id_number = id_match.group(1) if id_match else ""

    # 주소 추출 및 분리
    address_match = re.search(r'주소\s*([가-힣]+시\s[가-힣]+구\s[가-힣0-9\s]+로\s\d+\s\d+호\s*\(.*?\))', section_text)
    address = address_match.group(1) if address_match else ""
    address_parts = split_address(address)

    # 채무금액 추출 (숫자만 추출)
    debt_match = re.search(r'채무금액\s*([\d,]+)원', section_text)
    debt_amount = debt_match.group(1).replace(",", "") if debt_match else ""

    return [name, id_number] + address_parts + [debt_amount]

def split_address(address):
    """주소를 광역, 기초, 도로명, 건물번호, 상세주소로 분리합니다."""
    match = re.match(r'([가-힣]+시)\s([가-힣]+구)\s([가-힣0-9\s]+로)\s(\d+)\s(\d+호)\s*\((.*?)\)', address)
    if match:
        return list(match.groups())
    return ["", "", "", "", "", ""]  # 주소 형식이 맞지 않을 경우 빈 값 반환

def save_to_excel(data, excel_path):
    """데이터를 엑셀 파일로 저장합니다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    # 헤더 작성
    headers = ["성명", "주민등록번호", "광역", "기초", "도로명", "건물번호", "상세주소", "채무금액"]
    ws.append(headers)

    # 데이터 쓰기
    for row in data:
        ws.append(row)

    wb.save(excel_path)

if __name__ == "__main__":
    pdf_path = "example.pdf"  # 처리할 PDF 파일 경로
    excel_path = "output.xlsx"  # 저장할 엑셀 파일 경로

    # PDF에서 텍스트 추출
    text = extract_text_from_pdf(pdf_path)

    # 데이터 추출
    extracted_data = extract_data(text)

    # 엑셀로 저장
    save_to_excel([extracted_data], excel_path)

    print(f"데이터가 '{excel_path}'에 저장되었습니다.")
