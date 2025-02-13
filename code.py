import os
import csv
import fitz  # PyMuPDF

def extract_toc_from_pdf(pdf_path):
    toc = []
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        toc.append(text)
    return toc

def parse_filename(filename):
    # 예제 파일명: "2025 위원회명 분과명 차수.pdf"
    parts = os.path.splitext(filename)[0].split()
    year = parts[0]
    committee = parts[1]
    subcommittee = parts[2]
    session = parts[3]
    return year, committee, subcommittee, session

def parse_toc_item(item):
    # 예제: "목차번호 의결사항"
    parts = item.split()
    toc_number = parts[0]
    return toc_number

def find_resolution(doc, toc_number):
    resolution_keywords = ["원안의결", "부결", "조건부가결"]
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        if toc_number in text:
            for keyword in resolution_keywords:
                if keyword in text:
                    return keyword
    return "의결사항 없음"

def main(folder_path, output_csv):
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['파일명', '년도', '위원회명', '분과명', '차수', '목차번호', '의결사항'])
        for pdf_file in pdf_files:
            year, committee, subcommittee, session = parse_filename(os.path.basename(pdf_file))
            toc = extract_toc_from_pdf(pdf_file)
            doc = fitz.open(pdf_file)
            for item in toc:
                toc_number = parse_toc_item(item)
                resolution = find_resolution(doc, toc_number)
                csvwriter.writerow([os.path.basename(pdf_file), year, committee, subcommittee, session, toc_number, resolution])

if __name__ == "__main__":
    folder_path = r'C:\Users\A\Documents\GitHub\pdf_classification\회의록'
    output_csv = os.path.join(os.path.dirname(__file__), 'output.csv')
    main(folder_path, output_csv)