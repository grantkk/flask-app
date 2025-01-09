from flask import Flask, request, render_template, send_file
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"

# 確保目錄存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# 格式修正函數
def format_option(text):
    import re
    text = re.sub(r'[ＡＢＣＤ]', lambda m: chr(ord(m.group(0)) - 0xFEE0), text)
    text = re.sub(r'[（）]', lambda m: '(' if m.group(0) == '（' else ')', text)
    text = re.sub(r'\(?([A-D])\)?', r'(\1)', text)
    text = re.sub(r'\(\w\)\s*', lambda m: f"{m.group(0).strip()} ", text)
    return text

# 設定段落字型函數
def set_font_to_mingti(paragraph):
    for run in paragraph.runs:
        run.font.name = "新細明體"
        r = run._element
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:eastAsia"), "新細明體")
        r.insert(0, rFonts)

# 修改 Word 文件函數
def process_word_file(input_file, output_file):
    doc = Document(input_file)
    for paragraph in doc.paragraphs:
        if any(option in paragraph.text for option in ["A", "B", "C", "D", "Ａ", "Ｂ", "Ｃ", "Ｄ"]):
            paragraph.text = format_option(paragraph.text)
            set_font_to_mingti(paragraph)
    doc.save(output_file)

# 路由：首頁
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # 處理上傳的文件
        uploaded_file = request.files["file"]
        if uploaded_file.filename.endswith(".docx"):
            input_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            output_path = os.path.join(PROCESSED_FOLDER, f"processed_{uploaded_file.filename}")
            uploaded_file.save(input_path)
            
            # 修改文件
            process_word_file(input_path, output_path)
            
            # 提供下載連結
            return f'''
            <p>文件處理完成！點擊下載：</p>
            <a href="/download/{os.path.basename(output_path)}">下載修正版文件</a>
            '''
        else:
            return "<p>請上傳 .docx 文件！</p>"
    return '''
    <h1>上傳 Word 文件並修正格式</h1>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="file" />
        <button type="submit">上傳並處理</button>
    </form>
    '''

# 路由：下載文件
@app.route("/download/<filename>")
def download_file(filename):
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    # Render 需要綁定到 0.0.0.0 並從 PORT 環境變數獲取端口號
    port = int(os.environ.get("PORT", 5000))  # 預設為 5000，適用於本地測試
    app.run(host="0.0.0.0", port=port, debug=True)
