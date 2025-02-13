import os
import csv
import fitz  # PyMuPDF
import re

def parse_filename(filename):
    """ 파일명에서 연도, 위원회명, 분과명, 차수, 소위원회 여부 추출 """
    parts = os.path.splitext(filename)[0].split()
    year = parts[0].replace('년도', '')
    committee = parts[1]  # 문화재위원회
    subcommittee = parts[2]  # 사적분과
    session = parts[3]  # 제1차
    subcommittee_flag = '소위원회' if '소위원회' in parts else '위원회'
    return year, committee, subcommittee, session, subcommittee_flag

def parse_table_item(item):
    # 예제: "심의사항 목차번호 안건명"
    parts = item.split()
    category = parts[0].replace('【', '').replace('】', '')
    toc_number = parts[1]
    agenda_name = ' '.join(parts[2:])
    return category, toc_number, agenda_name

def extract_table_of_contents(pdf_path):
    """ PDF 2~3페이지에서 목차번호와 안건명을 블록 단위로 추출 """
    doc = fitz.open(pdf_path)
    toc_items = []
    found_section = None

    for page_num in range(2):  # 2~3페이지(0-indexed)
        if page_num >= len(doc):
            break

        # 블록 단위로 텍스트 가져오기
        blocks = doc[page_num].get_text("blocks")

        for block in blocks:
            text = block[4].strip()  # 블록 내 텍스트 가져오기

            # '심의사항', '검토사항', '보고사항'이 나오면 해당 구분으로 설정
            if any(section in text for section in ["심의사항", "검토사항", "보고사항"]):
                found_section = re.sub(r'\(.*?\)', '', text).strip()  # 괄호와 그 안의 내용 제거
                found_section = found_section.replace('【', '').replace('】', '')  # 【 】 제거
                continue  # 다음 블록부터 목차 데이터라고 간주

            # 정규식으로 목차 번호와 안건명 추출
            match = re.match(r"(\d+)\s+(.+)", text)  # 예: "1 경산 병영유적 주변 도시계획도로 개설"
            if match:
                toc_items.append({
                    "사항구분": found_section if found_section else "심의사항",  # 기본값은 '심의사항'
                    "목차번호": match.group(1),
                    "안건명": match.group(2).strip()
                })

    print(f"📄 {pdf_path}에서 추출된 목차 항목: {toc_items}")  # 디버깅 출력
    return toc_items

def extract_decisions(pdf_path, toc_items):
    """ PDF 본문에서 의결사항 추출 """
    doc = fitz.open(pdf_path)
    decision_keywords = ["원안가결", "원안의결", "조건부가결", "보류", "부결"]
    
    for item in toc_items:
        found_decision = False
        for page_num in range(len(doc)):
            page = doc[page_num]
            blocks = page.get_text("blocks")
            for block in blocks:
                text = block[4].strip()
                if item["목차번호"] in text:
                    for keyword in decision_keywords:
                        if keyword in text:
                            item["의결사항"] = keyword
                            found_decision = True
                            break
                    if found_decision:
                        break
            if found_decision:
                break
        if not found_decision:
            item["의결사항"] = "의결사항 없음"

    print(f"📄 {pdf_path}에서 추출된 의결사항: {toc_items}")  # 디버깅 출력

def process_pdfs_in_folder(folder_path, output_csv):
    """ 폴더 내 모든 PDF 파일을 처리하고 CSV로 저장 """
    all_data = []

    for file in os.listdir(folder_path):
        if file.endswith(".pdf") or file.endswith(".PDF"):  # 대소문자 구분 없이 PDF 파일 찾기
            pdf_path = os.path.join(folder_path, file)
            print(f"📂 처리 중: {pdf_path}")  # 진행 상황 표시
            
            year, committee, subcommittee, session, subcommittee_flag = parse_filename(file)

            # 목차(안건) 정보 추출
            toc_items = extract_table_of_contents(pdf_path)

            # 의결사항 정보 추가
            extract_decisions(pdf_path, toc_items)

            # 결과 저장
            for item in toc_items:
                all_data.append([
                    file,
                    year,
                    committee,
                    subcommittee,
                    session,
                    subcommittee_flag,
                    item["사항구분"],
                    item["목차번호"],
                    item["안건명"],
                    item.get("의결사항", "의결사항 없음")
                ])

    print(f"📊 최종 데이터: {all_data}")  # 디버깅 출력

    # CSV 파일 저장
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['파일명', '년도', '위원회명', '분과명', '차수', '소위원회여부', '사항구분', '목차번호', '안건명', '의결사항'])
        csvwriter.writerows(all_data)

if __name__ == "__main__":
    folder_path = r"C:\Users\A\Documents\GitHub\pdf_classification\회의록"  # 회의록 폴더 경로 입력
    output_csv = os.path.join(os.path.dirname(__file__), r'output.csv')
    
    process_pdfs_in_folder(folder_path, output_csv)

    print("✅ 모든 PDF 처리 완료! 결과 파일:", output_csv)
