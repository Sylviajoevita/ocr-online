from flask import Flask, request, send_file, render_template_string, send_from_directory
from flask_cors import CORS
from PIL import Image
import pytesseract
import io
from openpyxl import Workbook
from openpyxl.styles import Font
import os

app = Flask(__name__)
CORS(app)

HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OCR Studio</title>
<style>
*{box-sizing:border-box;font-family:Inter,system-ui,sans-serif}
body{
  margin:0;
  min-height:100vh;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:20px;
  /* Background image from your upload */
  background-image: url('/background.jpg');
  background-size: cover;
  background-position: center;
  background-attachment: fixed;
  background-color: #6a7cff;
}
.card{
  width:100%;
  max-width:680px;
  background: rgba(255,255,255,0.92);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border-radius:20px;
  padding:28px;
  box-shadow:0 20px 60px rgba(0,0,0,.25);
  border:1px solid rgba(255,255,255,.6);
}
h2{margin:0 0 8px;text-align:center}
.sub{text-align:center;color:#666;font-size:13px;margin-bottom:18px}
.drop{border:2px dashed #7c5cfc;border-radius:14px;padding:36px;text-align:center;background:rgba(250,248,255,.9);cursor:pointer;transition:.2s}
.drop:hover{background:#fff;border-color:#5a3de8}
.preview{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:12px;margin-top:18px}
.item{border:1px solid #e9e9e9;border-radius:10px;overflow:hidden;background:#fff;position:relative;box-shadow:0 2px 8px rgba(0,0,0,.05)}
.item img{width:100%;height:100px;object-fit:cover;display:block}
.item span{display:block;padding:6px;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;background:#f9f9f9}
.item button{position:absolute;top:4px;right:4px;background:rgba(0,0,0,.75);color:#fff;border:none;border-radius:50%;width:22px;height:22px;cursor:pointer}
.btn{width:100%;padding:15px;background:#111;color:#fff;border:none;border-radius:12px;margin-top:16px;cursor:pointer;font-weight:700;font-size:15px;display:none}
.btn.show{display:block}
.btn:disabled{background:#888}
.log{background:#111;color:#0f0;padding:10px;border-radius:8px;margin-top:12px;font-family:monospace;font-size:12px;white-space:pre-wrap}
</style>
</head>
<body>
<div class="card">
<h2>OCR Studio</h2>
<div class="sub">File name will be in Row 1 • Text extracted exactly as in image</div>
<div class="drop" id="dropZone" onclick="document.getElementById('images').click()">
<strong>Click to select images</strong><br><small>or drag & drop here</small>
<input type="file" id="images" multiple hidden accept="image/*">
</div>
<div id="preview" class="preview"></div>
<button class="btn" id="processBtn" onclick="upload()">Process & Download Excel</button>
<div id="log" class="log">Ready - Select images to preview</div>
</div>
<script>
const input=document.getElementById('images'), preview=document.getElementById('preview'), btn=document.getElementById('processBtn'), log=document.getElementById('log'), dropZone=document.getElementById('dropZone');
let selectedFiles=[];
input.addEventListener('change', e=>{ handleFiles(e.target.files); });
dropZone.addEventListener('dragover', e=>{ e.preventDefault(); dropZone.style.background='#fff'; });
dropZone.addEventListener('dragleave', ()=>{ dropZone.style.background='rgba(250,248,255,.9)'; });
dropZone.addEventListener('drop', e=>{ e.preventDefault(); dropZone.style.background='rgba(250,248,255,.9)'; handleFiles(e.dataTransfer.files); });
function handleFiles(files){ selectedFiles=selectedFiles.concat(Array.from(files)); renderPreview(); }
function renderPreview(){
 preview.innerHTML='';
 if(selectedFiles.length===0){ btn.classList.remove('show'); log.textContent='Ready - Select images to preview'; return; }
 selectedFiles.forEach((file,i)=>{
   const div=document.createElement('div'); div.className='item';
   const img=document.createElement('img'); img.src=URL.createObjectURL(file);
   const name=document.createElement('span'); name.textContent=file.name;
   const rm=document.createElement('button'); rm.textContent='x'; rm.onclick=()=>{ selectedFiles.splice(i,1); renderPreview(); };
   div.appendChild(img); div.appendChild(name); div.appendChild(rm); preview.appendChild(div);
 });
 btn.classList.add('show'); btn.textContent=`Process ${selectedFiles.length} Image(s) & Download Excel`;
 log.textContent=`${selectedFiles.length} image(s) ready - Preview above - Click Process`;
}
async function upload(){
 if(selectedFiles.length===0) return; btn.disabled=true; btn.textContent='Processing...';
 log.textContent='OCR running on server...';
 const fd=new FormData(); selectedFiles.forEach(f=>fd.append('images', f));
 try{
  const res=await fetch('/ocr',{method:'POST',body:fd});
  if(!res.ok) throw new Error('Server error');
  const blob=await res.blob(); const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download='extracted_sheets.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  log.textContent='Success! Excel downloaded';
  btn.textContent='Downloaded! Process again?';
 }catch(e){ log.textContent='Error: '+e.message; btn.textContent='Try Again'; }
 finally{ btn.disabled=false; }
}
</script>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_PAGE)

@app.route('/background.jpg')
def background():
    # Serves the background image you uploaded to GitHub
    if os.path.exists("background.jpg"):
        return send_file("background.jpg", mimetype="image/jpeg")
    # fallback if not uploaded yet - return 404 so CSS shows solid color
    return "", 404

@app.route('/ocr', methods=['POST'])
def ocr():
    files = request.files.getlist('images')
    wb = Workbook()
    wb.remove(wb.active)
    for f in files:
        raw_name = f.filename.rsplit('.', 1)[0]
        safe_sheet = "".join(c if c.isalnum() or c in (" ", "_") else "_" for c in raw_name)[:31] or "Sheet1"
        ws = wb.create_sheet(title=safe_sheet)
        try:
            img = Image.open(f.stream).convert("L")
            text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
        except Exception as e:
            text = f"OCR Error: {str(e)}"
        # Row 1 = File Name
        ws['A1'] = f.filename
        ws['A1'].font = Font(bold=True, size=12, color="0000FF")
        # Row 3 onwards = text as in image
        row_num = 3
        for line in text.splitlines():
            cleaned = line.strip()
            if cleaned == "":
                row_num += 1
                continue
            ws.cell(row=row_num, column=1, value=cleaned)
            row_num += 1
        ws.column_dimensions['A'].width = 80
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="extracted_sheets.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
