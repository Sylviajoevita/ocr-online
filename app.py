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
<title>OCR Studio - With Preview</title>
<style>
*{box-sizing:border-box;font-family:Inter,sans-serif}
body{margin:0;background:#f8f9ff;padding:15px}
.card{max-width:750px;margin:auto;background:white;border-radius:16px;padding:22px;box-shadow:0 10px 30px rgba(0,0,0,.07)}
h2{margin:0 0 15px}
.drop{border:2px dashed #7c5cfc;border-radius:12px;padding:35px;text-align:center;background:#faf8ff;cursor:pointer;transition:.2s}
.drop:hover{background:#f3f0ff}
.preview{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:12px;margin-top:18px}
.item{border:1px solid #e9e9e9;border-radius:10px;overflow:hidden;background:#fff;position:relative}
.item img{width:100%;height:100px;object-fit:cover;display:block}
.item span{display:block;padding:6px;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;background:#f9f9f9}
.item button{position:absolute;top:4px;right:4px;background:rgba(0,0,0,.7);color:#fff;border:none;border-radius:50%;width:22px;height:22px;cursor:pointer;font-size:12px}
.btn{width:100%;padding:14px;background:#111;color:#fff;border:none;border-radius:10px;margin-top:14px;cursor:pointer;font-weight:700;font-size:15px;display:none}
.btn.show{display:block}
.btn:disabled{background:#999;cursor:not-allowed}
.log{background:#111;color:#0f0;padding:10px;border-radius:8px;margin-top:12px;font-family:monospace;font-size:13px;white-space:pre-wrap}
</style>
</head>
<body>
<div class="card">
<h2>OCR Studio</h2>
<div class="drop" id="dropZone" onclick="document.getElementById('images').click()">
<strong>Click to select images</strong><br><small>or drag & drop here - Multiple allowed</small>
<input type="file" id="images" multiple hidden accept="image/*">
</div>

<div id="preview" class="preview"></div>

<button class="btn" id="processBtn" onclick="upload()">Process & Download Excel</button>
<div id="log" class="log">Ready - Select images to preview</div>
</div>

<script>
const input = document.getElementById('images');
const preview = document.getElementById('preview');
const btn = document.getElementById('processBtn');
const log = document.getElementById('log');
const dropZone = document.getElementById('dropZone');
let selectedFiles = [];

input.addEventListener('change', handleFiles);
dropZone.addEventListener('dragover', e=>{ e.preventDefault(); dropZone.style.background='#ece6ff'; });
dropZone.addEventListener('dragleave', ()=>{ dropZone.style.background='#faf8ff'; });
dropZone.addEventListener('drop', e=>{
 e.preventDefault();
 dropZone.style.background='#faf8ff';
 handleFiles({target:{files:e.dataTransfer.files}});
});

function handleFiles(e){
 const newFiles = Array.from(e.target.files);
 selectedFiles = selectedFiles.concat(newFiles);
 renderPreview();
}

function renderPreview(){
 preview.innerHTML='';
 if(selectedFiles.length===0){
   btn.classList.remove('show');
   log.textContent='Ready - Select images to preview';
   return;
 }
 selectedFiles.forEach((file, index)=>{
   const div=document.createElement('div');
   div.className='item';
   const img=document.createElement('img');
   img.src=URL.createObjectURL(file);
   const name=document.createElement('span');
   name.textContent=file.name;
   const rm=document.createElement('button');
   rm.textContent='x';
   rm.onclick=()=>{ selectedFiles.splice(index,1); renderPreview(); };
   div.appendChild(img); div.appendChild(name); div.appendChild(rm);
   preview.appendChild(div);
 });
 btn.classList.add('show');
 btn.textContent=`Process ${selectedFiles.length} Image(s) & Download Excel`;
 log.textContent=`${selectedFiles.length} image(s) selected - Preview above - Click Process to download`;
}

async function upload(){
 if(selectedFiles.length===0) return;
 btn.disabled=true;
 btn.textContent='Processing... Please wait';
 log.textContent='Uploading '+selectedFiles.length+' files to server... OCR running...';

 const fd=new FormData();
 selectedFiles.forEach(f=>fd.append('images', f));

 try{
  const res=await fetch('/ocr',{method:'POST',body:fd});
  if(!res.ok) throw new Error('Server error '+res.status);
  const blob=await res.blob();
  const url=window.URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='extracted_sheets.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  log.textContent='Success! '+selectedFiles.length+' images processed - Excel downloaded';
  btn.textContent='Downloaded! Process again?';
 }catch(e){
  log.textContent='Error: '+e.message;
  btn.textContent='Try Again';
 }finally{
  btn.disabled=false;
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
