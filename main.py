import os
import csv
import fitz  # PyMuPDF
import re

# 위원회 명칭 매핑 (동일한 위원회를 하나로 통합)
COMMITTEE_MAPPING = {
    "문화유산위원회": "문화재위원회",  # 문화유산위원회 → 문화재위원회로 통합
    "문화재위원회": "문화재위원회",
    "자연문화재위원회": "자연문화재위원회"
}

def parse_filename(filename):
    """파일명에서 연도, 위원회명, 차수 추출"""
    parts = os.path.splitext(filename)[0].split()
    year = parts[0].replace('년도', '')
    committee = COMMITTEE_MAPPING.get(parts[1], parts[1])  # 명칭 매핑 적용
    session = re.sub(r'\D', '', parts[2])  # 숫자만 추출
    return year, committee, session

def extract_table_of_contents(pdf_path):
    """PDF에서 목차번호와 안건명 추출"""
    doc = fitz.open(pdf_path)
    toc_items = []
    found_section = None

    for page_num in range(min(3, len(doc))):  # 첫 3페이지만 탐색
        blocks = doc[page_num].get_text("blocks")
        for block in blocks:
            text = block[4].strip()

            if any(section in text for section in ["심의사항", "검토사항", "보고사항"]):
                found_section = re.sub(r'\(.*?\)', '', text).strip().replace('【', '').replace('】', '')
                continue  

            match = re.match(r"(\d+)\s+(.+)", text)  
            if match:
                toc_items.append({
                    "사항구분": found_section if found_section else "심의사항",
                    "목차번호": match.group(1),
                    "안건명": match.group(2).strip()
                })

    return toc_items

def extract_decisions(pdf_path, toc_items):
    """PDF 본문에서 의결사항 추출"""
    doc = fitz.open(pdf_path)
    decision_keywords = ["원안가결", "원안의결", "조건부가결", "보류", "부결"]

    for item in toc_items:
        found_decision = False
        for page_num in range(len(doc)):
            blocks = doc[page_num].get_text("blocks")
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

def load_existing_data(csv_path):
    """기존 CSV 파일에서 저장된 PDF 목록을 불러오기"""
    existing_files = set()
    if os.path.exists(csv_path):
        with open(csv_path, 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            next(reader, None)  # 헤더 스킵
            for row in reader:
                if row:  
                    existing_files.add(row[0])  # 파일명만 저장
    return existing_files

def process_pdfs_in_folder(folder_path, output_folder, error_log):
    """위원회별 CSV 파일을 생성하며, 중복 데이터는 제외하고 추가"""
    error_files = []

    for committee_folder in os.listdir(folder_path):
        committee_path = os.path.join(folder_path, committee_folder)
        if not os.path.isdir(committee_path):
            continue

        committee_name = COMMITTEE_MAPPING.get(committee_folder, committee_folder)
        output_csv = os.path.join(output_folder, f"{committee_name}.csv")

        existing_files = load_existing_data(output_csv)  # 기존 CSV 데이터 불러오기
        new_data = []

        print(f"📂 처리 중: {committee_folder} → 저장: {committee_name}.csv")

        for year_folder in os.listdir(committee_path):
            year_path = os.path.join(committee_path, year_folder)
            if not os.path.isdir(year_path):
                continue

            for file in os.listdir(year_path):
                if "소위원회" in file or "별첨" in file:
                    print(f"  ⏩ 소위원회 또는 별첨 파일 건너뜀: {file}")
                    continue

                if file.lower().endswith(".pdf") and file not in existing_files:
                    pdf_path = os.path.join(year_path, file)
                    print(f"  📄 새로운 파일 추가: {pdf_path}")

                    try:
                        year, committee, session = parse_filename(file)
                        toc_items = extract_table_of_contents(pdf_path)
                        extract_decisions(pdf_path, toc_items)

                        for item in toc_items:
                            new_data.append([
                                file, year, committee, session,
                                item["사항구분"], item["목차번호"], item["안건명"], item.get("의결사항", "의결사항 없음")
                            ])
                    except Exception as e:
                        print(f"❌ 오류 발생: {file} - {str(e)}")
                        error_files.append(f"{file}: {str(e)}")

        # 🚀 새로운 데이터 추가 저장
        if new_data:
            with open(output_csv, 'a', newline='', encoding='utf-8-sig') as csvfile:
                csvwriter = csv.writer(csvfile)

                # 파일이 없거나 비어 있으면 헤더 추가
                if not os.path.exists(output_csv) or os.stat(output_csv).st_size == 0:
                    csvwriter.writerow(['파일명', '년도', '위원회명', '차수', '사항구분', '목차번호', '안건명', '의결사항'])

                csvwriter.writerows(new_data)
            print(f"✅ 새로운 데이터 저장 완료: {output_csv}")

    # 오류 파일 저장
    if error_files:
        with open(error_log, 'w', encoding='utf-8') as error_file:
            error_file.write("\n".join(error_files))
        print(f"⚠️ 오류 파일이 있습니다! 오류 목록: {error_log}")

if __name__ == "__main__":
    folder_path = r"C:\Users\A\Documents\GitHub\pdf_classification\회의록"
    output_folder = r"C:\Users\A\Documents\GitHub\pdf_classification\csv_output"
    error_log = os.path.join(os.path.dirname(__file__), 'error_files.txt')

    os.makedirs(output_folder, exist_ok=True)  # 📁 CSV 저장 폴더 생성

    try:
        output_csv_files = process_pdfs_in_folder(folder_path, output_folder, error_log)
        print("✅ 모든 PDF 처리 완료! 결과 파일:", output_csv_files)
    except Exception as e:
        print(f"🚨 오류 발생: {str(e)}")
    
    if os.path.exists(error_log):
        print("🚨 오류 발생한 파일 목록:", error_log)
