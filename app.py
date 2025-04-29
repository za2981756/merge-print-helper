from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
import pandas as pd
import re
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 載入郵遞區號對照表
zipcode_mapping = {}
with open("zipcode_mapping.txt", 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        if line:
            parts = line.split()
            if len(parts) == 2:
                zipcode_mapping[parts[0]] = parts[1]

def extract_info(text):
    doorplate = ""
    owner = ""
    address = ""
    lines = text.splitlines()
    for line in lines:
        line = line.strip()
        if "建物門牌" in line and not doorplate:
            m = re.search(r"建物門牌\s*[:：]?\s*(.+)", line)
            if m:
                doorplate = re.sub(r"(債權加總|坪數加總).*", "", m.group(1).strip())
        if "所有權人" in line and not owner:
            m = re.search(r"所有權人\s*[:：]?\s*([^\s統一編號]*)", line)
            if m:
                owner = m.group(1).strip()
        if ("地 址" in line or "住 址" in line or "地址" in line) and not address:
            m = re.search(r"(地 址|住 址|地址)\s*[:：]?\s*(.+)", line)
            if m:
                address = m.group(2).strip()

    zipcode = ""
    for area, code in zipcode_mapping.items():
        if area in doorplate or area in address:
            zipcode = code
            break

    return {
        "郵遞區號": zipcode,
        "建物門牌": doorplate,
        "姓名": owner,
        "地址": address
    }

def ocr_pdf(filepath):
    images = convert_from_path(filepath, dpi=300)
    full_text = ""
    for img in images:
        gray = img.convert('L')
        bw = gray.point(lambda x: 0 if x < 200 else 255, '1')
        text = pytesseract.image_to_string(bw, lang='chi_tra+eng', config='--psm 6')
        full_text += text + "\n"
    return full_text

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        files = request.files.getlist('pdfs')
        results = []
        for file in files:
            filename = secure_filename(file.filename)
            save_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(save_path)
            try:
                with pdfplumber.open(save_path) as pdf:
                    full_text = ""
                    for page in pdf.pages:
                        full_text += page.extract_text() or ""
                if len(full_text.strip()) < 20:
                    full_text = ocr_pdf(save_path)
                info = extract_info(full_text)
                results.append(info)
            except Exception as e:
                results.append({
                    "郵遞區號": "",
                    "建物門牌": f"{filename}（錯誤：{str(e)}）",
                    "姓名": "",
                    "地址": ""
                })
        df = pd.DataFrame(results)
        df.to_excel(os.path.join(OUTPUT_FOLDER, '批次擷取結果.xlsx'), index=False)
        return render_template('index.html', results=results)
    return render_template('index.html')

@app.route('/download')
def download_file():
    return send_file(os.path.join(OUTPUT_FOLDER, '批次擷取結果.xlsx'), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)
