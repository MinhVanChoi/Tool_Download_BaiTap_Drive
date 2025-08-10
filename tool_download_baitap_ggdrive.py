import pandas as pd
import os
import requests
import re

def get_confirm_token(response):
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            return value
    return None

def download_file_from_google_drive(file_id, dest_folder):
    session = requests.Session()
    URL = "https://drive.google.com/uc?export=download"

    response = session.get(URL, params={'id': file_id}, stream=True)
    token = get_confirm_token(response)

    if token:
        response = session.get(URL, params={'id': file_id, 'confirm': token}, stream=True)

    if "Content-Disposition" not in response.headers:
        raise Exception("Không thể lấy tên file từ header. Có thể file là Google Docs hoặc chưa chia sẻ đúng.")

    # Lấy tên file từ header
    content_disp = response.headers.get("Content-Disposition")
    filename_match = re.findall('filename="(.+)"', content_disp)
    if filename_match:
        filename = filename_match[0]
    else:
        filename = f"{file_id}.download"

    file_path = os.path.join(dest_folder, filename)

    with open(file_path, "wb") as f:
        for chunk in response.iter_content(32768):
            if chunk:
                f.write(chunk)

    return file_path

# Đọc file Excel
df = pd.read_excel('excel.xlsx')
base_dir = 'excel'
os.makedirs(base_dir, exist_ok=True)

# Check các cột - thay đổi nếu muốn xài theo nhiều cách - giữ cách chia link lại
with open('log.txt', 'w', encoding='utf-8') as log_file:
    for index, row in df.head(90).iterrows():
        name = str(row['Họ và tên:']).strip()
        mssv = str(row['MSSV:']).strip()
        link = str(row['Nộp Tập']).strip()

        if pd.isna(link) or link == '' or link == 'nan':
            log_file.write(f"[BỎ QUA] {mssv} - {name} không có link\n")
            continue

        # Xử lý link Google Drive
        if 'id=' in link:
            file_id = link.split('id=')[1]
        elif 'file/d/' in link:
            file_id = link.split('/file/d/')[1].split('/')[0]
        else:
            log_file.write(f"[LINK SAI] {mssv} - {name}: {link}\n")
            continue

        # Thư mục sinh viên
        student_dir = os.path.join(base_dir, f"{mssv}_{name.replace(' ', '_')}")
        os.makedirs(student_dir, exist_ok=True)

        # Bỏ qua nếu đã có file
        if any(os.path.isfile(os.path.join(student_dir, f)) for f in os.listdir(student_dir)):
            log_file.write(f"[ĐÃ TẢI] {mssv} - {name} đã có file\n")
            continue

        try:
            file_path = download_file_from_google_drive(file_id, student_dir)
            log_file.write(f"[THÀNH CÔNG] {mssv} - {name}: {os.path.basename(file_path)}\n")
        except Exception as e:
            log_file.write(f"[LỖI] {mssv} - {name}: {str(e)}\n👉 Kiểm tra thủ công: https://drive.google.com/uc?id={file_id}\n")

print("Đã hoàn tất! Kiểm tra thư mục 'tuan03' và file 'log.txt'.")
