from flask import Flask, request, send_file, render_template_string
from flask_cors import CORS
from PIL import Image
import pytesseract
import io
from openpyxl import Workbook
from openpyxl.styles import Font

app = Flask(__name__)
CORS(app)

HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OCR Studio - Fixed</title>
<style>
*{box-sizing:border-box;font-family:Inter,sans-serif}
body{margin:0;background:#f8f9ff;padding:20px}
.card{max-width:700px;margin:auto;background:white;border-radius:16px;padding:25px;box-shadow:0 10px 30px rgba(0,0,0,.06)}
.drop{border:2px dashed #7c5cfc;border-radius:12px;padding:40px;text-align:center;background:#faf8ff;cursor:pointer}
.btn{width:100%;padding:14px;background:#111;color:#fff;border:none;border-radius:10px;margin-top:12px;cursor:pointer;font-weight:600}
.log{background:#111;color:#0f0;padding:10px;border-radius:8px;margin-top:12px;font-family:monospace;font-size:13px}
</style>
</head>
<body>
<div class="card">
<h2>OCR Studio - File Name in Row 1</h2>
<div class="drop" onclick="document.getElementById('images').click()">
Click to select images<br><small>Multiple images allowed</small>
<input type="file" id="images" multiple hidden accept="image/*">
</div>
<button class="btn" onclick="upload()">Process & Download Excel</button>
<div id="log" class="log">Ready - Select images</div>
</div>
<script>
async function upload(){
 const files=document.getElementById('images').files;
 if(!files.length) return alert('Select images first');
 const log=document.getElementById('log');
 log.textContent='Uploading '+files.length+' files to server... Please wait';
 const fd=new FormData();
 for(let f of files) fd.append('images', f);
 try{
  const res=await fetch('/ocr',{method:'POST',body:fd});
  if(!res.ok) throw new Error('Server error');
  const blob=await res.blob();
  const url=window.URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='extracted_sheets.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  log.textContent='Success! Excel downloaded - Open it in Excel/Sheets';
 }catch(e){
  log.textContent='Error: '+e.message;
 }
}
</script>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_PAGE)

@app.route('/ocr', methods=['POST'])
def ocr():
    files = request.files.getlist('images')

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    for f in files:
        # Clean sheet name (Excel sheet name max 31 chars, no special chars)
        raw_name = f.filename.rsplit('.', 1)[0]
        safe_sheet = "".join(c if c.isalnum() or c in (" ", "_") else "_" for c in raw_name)[:31]
        if not safe_sheet:
            safe_sheet = "Sheet1"

        ws = wb.create_sheet(title=safe_sheet)

        try:
            # Image preprocessing for better OCR
            img = Image.open(f.stream).convert("L") # grayscale
            # Optional: increase contrast
            # img = img.point(lambda x: 0 if x < 140 else 255)

            text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
        except Exception as e:
            text = f"OCR Error: {str(e)}"

        # --- YOUR REQUESTED LAYOUT ---
        # Row 1, Col A = File Name
        ws['A1'] = f.filename
        ws['A1'].font = Font(bold=True, size=12, color="0000FF")

        # Row 2 = empty for spacing
        # Row 3 onwards = extracted text, one line per row, same as in image
        row_num = 3
        for line in text.splitlines():
            cleaned = line.strip()
            if cleaned == "":
                row_num += 1 # keep empty lines to preserve layout
                continue
            ws.cell(row=row_num, column=1, value=cleaned)
            row_num += 1

        # Make column A wide to see text well
        ws.column_dimensions['A'].width = 80

    # Save to memory and send to user
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="extracted_sheets.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
