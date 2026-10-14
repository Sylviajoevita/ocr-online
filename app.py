from flask import Flask, request, send_file, render_template_string
from flask_cors import CORS
from PIL import Image
import pytesseract
import io
import json
from openpyxl import Workbook
from openpyxl.styles import Font
import os
import traceback

app = Flask(__name__)
CORS(app)

HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OCR Studio </title>
<link rel="icon" href="/favicon.ico">
<style>
*{box-sizing:border-box;font-family:Inter,system-ui,sans-serif}
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;background-image:url('/background.jpg');background-size:cover;background-position:center;background-color:#7a8cff}
.card{width:100%;max-width:700px;background:rgba(255,255,255,.28);backdrop-filter:blur(20px);border-radius:24px;padding:32px;box-shadow:0 8px 32px rgba(31,38,135,.2);border:1px solid rgba(255,255,255,.5)}
h2{margin:0 0 8px;text-align:center;color:#111;font-size:26px}
.sub{text-align:center;color:rgba(0,0,0,.7);font-size:13px;margin-bottom:20px}
.drop{border:1.5px dashed rgba(255,255,255,.9);border-radius:16px;padding:40px;text-align:center;background:rgba(255,255,255,.35);cursor:pointer;font-weight:600}
.preview{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:12px;margin-top:18px}
.item{border-radius:12px;overflow:hidden;background:rgba(255,255,255,.9);position:relative;padding:10px;text-align:center;font-size:12px}
.item img{width:100%;height:80px;object-fit:cover;border-radius:8px}
.item button{position:absolute;top:5px;right:5px;background:rgba(0,0,0,.7);color:#fff;border:none;border-radius:50%;width:24px;height:24px;cursor:pointer}
.btn{width:100%;padding:16px;background:#111;color:#fff;border:none;border-radius:14px;margin-top:18px;cursor:pointer;font-weight:700;display:none}
.btn.show{display:block}
.log{background:rgba(17,17,17,.9);color:#7CFF7C;padding:12px;border-radius:10px;margin-top:14px;font-family:monospace;font-size:12px;white-space:pre-wrap}
</style>
</head>
<body>
<div class="card">
<h2>OCR Studio</h2>
<div class="sub">Upload Images OR.json files → Excel tables</div>
<div class="drop" onclick="document.getElementById('files').click()">Click to select.json or images<input type="file" id="files" multiple hidden accept="image/*,.json,application/json"></div>
<div id="preview" class="preview"></div>
<button class="btn" id="processBtn" onclick="upload()">Process & Download Excel</button>
<div id="log" class="log">Ready - select.json or images</div>
</div>
<script>
const input=document.getElementById('files'), preview=document.getElementById('preview'), btn=document.getElementById('processBtn'), log=document.getElementById('log');
let selectedFiles=[];
input.addEventListener('change', e=>{ selectedFiles=selectedFiles.concat(Array.from(e.target.files)); renderPreview(); });
function renderPreview(){
 preview.innerHTML=''; if(selectedFiles.length==0){ btn.classList.remove('show'); return; }
 selectedFiles.forEach((file,i)=>{
   const div=document.createElement('div'); div.className='item';
   if(file.type.startsWith('image/')){
     const img=document.createElement('img'); img.src=URL.createObjectURL(file); div.appendChild(img);
   } else {
     div.innerHTML='<div style="font-size:30px">📄</div>';
   }
   const name=document.createElement('span'); name.textContent=file.name;
   const rm=document.createElement('button'); rm.textContent='×'; rm.onclick=()=>{ selectedFiles.splice(i,1); renderPreview(); };
   div.appendChild(name); div.appendChild(rm); preview.appendChild(div);
 });
 btn.classList.add('show'); btn.textContent=`Process ${selectedFiles.length} file(s)`; log.textContent=`${selectedFiles.length} ready`;
}
async function upload(){
 if(selectedFiles.length==0) return; btn.disabled=true; btn.textContent='Processing...'; log.textContent='Reading files...';
 const fd=new FormData(); selectedFiles.forEach(f=>fd.append('files', f));
 try{
  const res=await fetch('/ocr',{method:'POST',body:fd});
  if(!res.ok){ const t=await res.text(); log.textContent=t.slice(0,3000); throw new Error('Server error'); }
  const blob=await res.blob(); const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download='from_json.xlsx'; document.body.appendChild(a); a.click(); a.remove();
  log.textContent='✅ Excel downloaded from JSON'; btn.textContent='✅ Done';
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
def bg():
    if os.path.exists("background.jpg"):
        return send_file("background.jpg", mimetype="image/jpeg")
    return "", 404

@app.route('/favicon.ico')
def fav():
    if os.path.exists("favicon.ico"):
        return send_file("favicon.ico", mimetype="image/x-icon")
    return "", 404

def write_json_to_sheet(ws, data):
    """Smart JSON to Excel table writer"""
    ws.column_dimensions['A'].width = 25
    # Case 1: List of dicts -> table with headers
    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
        headers = list(data[0].keys())
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True)
        for row_idx, row_data in enumerate(data, start=2):
            for col_idx, h in enumerate(headers, start=1):
                ws.cell(row=row_idx, column=col_idx, value=str(row_data.get(h, "")))
        return

    # Case 2: List of lists -> already table
    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], (list, tuple)):
        for r_idx, row in enumerate(data, start=1):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=str(val))
        return

    # Case 3: Dict -> keys in A, values in B, or nested
    if isinstance(data, dict):
        row_num = 1
        for k, v in data.items():
            ws.cell(row=row_num, column=1, value=k).font = Font(bold=True)
            if isinstance(v, (dict, list)):
                # Write nested as JSON string or flatten
                if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict):
                    # Sub-table
                    ws.cell(row=row_num, column=2, value=f"{len(v)} records -> see next rows")
                    row_num += 1
                    sub_headers = list(v[0].keys())
                    for col_idx, h in enumerate(sub_headers, start=1):
                        ws.cell(row=row_num, column=col_idx, value=h).font = Font(bold=True)
                    row_num += 1
                    for row_data in v:
                        for col_idx, h in enumerate(sub_headers, start=1):
                            ws.cell(row=row_num, column=col_idx, value=str(row_data.get(h, "")))
                        row_num += 1
                    continue
                else:
                    ws.cell(row=row_num, column=2, value=json.dumps(v, indent=2)[:30000])
            else:
                ws.cell(row=row_num, column=2, value=str(v))
            row_num += 1
        return

    # Fallback: single value
    ws.cell(row=1, column=1, value=str(data))

@app.route('/ocr', methods=['POST'])
def ocr():
    try:
        files = request.files.getlist('files') or request.files.getlist('images')
        if not files:
            return "No files", 400

        wb = Workbook()
        wb.remove(wb.active)

        for f in files:
            raw_name = f.filename.rsplit('.', 1)[0]
            safe = "".join(c if c.isalnum() or c in (" ", "_") else "_" for c in raw_name)[:31] or "Sheet1"
            ws = wb.create_sheet(title=safe)
            ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ""

            if ext == 'json':
                try:
                    content = f.read().decode('utf-8')
                    data = json.loads(content)
                    ws['A1'].value = f.filename
                    ws['A1'].font = Font(bold=True, color="0000FF")
                    # Create second sheet to hold data clean (keep A1 filename in first row, data from row 3 or row 1)
                    # We'll write from row 2
                    wb.remove(ws)
                    ws = wb.create_sheet(title=safe)
                    write_json_to_sheet(ws, data)
                    # Add filename on top
                    ws.insert_rows(1)
                    ws['A1'] = f.filename
                    ws['A1'].font = Font(bold=True, color="0000FF")
                except Exception as e:
                    ws['A1'] = f"JSON Error in {f.filename}: {str(e)}\n{traceback.format_exc()}"
            else:
                # Image OCR with table logic
                try:
                    img = Image.open(f.stream)
                    if max(img.size) > 2500:
                        img.thumbnail((2500,2500))
                    gray = img.convert("L")
                    data = pytesseract.image_to_data(gray, lang='eng', config='--oem 3 --psm 6', output_type=pytesseract.Output.DICT)
                    lines = {}
                    n = len(data['text'])
                    for i in range(n):
                        txt = data['text'][i].strip()
                        conf = int(data['conf'][i]) if data['conf'][i]!='-1' else 0
                        if txt=="" or conf<30: continue
                        key = (data['block_num'][i], data['par_num'][i], data['line_num'][i], data['top'][i]//10)
                        if key not in lines: lines[key]=[]
                        lines[key].append({'text':txt,'left':data['left'][i],'width':data['width'][i]})
                    row_num=2
                    ws['A1']=f.filename
                    ws['A1'].font=Font(bold=True,color="0000FF")
                    for key in sorted(lines.keys(), key=lambda k:(k[0],k[1],k[2],k[3])):
                        words=sorted(lines[key], key=lambda w:w['left'])
                        if not words: continue
                        avg_w=sum(w['width'] for w in words)/len(words)
                        cols=[]; cur=words[0]['text']; last_end=words[0]['left']+words[0]['width']
                        for w in words[1:]:
                            gap=w['left']-last_end
                            if gap>avg_w*0.8: cols.append(cur); cur=w['text']
                            else: cur+=" "+w['text']
                            last_end=w['left']+w['width']
                        cols.append(cur)
                        for col_idx,val in enumerate(cols,start=1):
                            ws.cell(row=row_num,column=col_idx,value=val)
                        row_num+=1
                except Exception as e:
                    ws.cell(row=2,column=1,value=f"OCR Error: {e}")

            for c in ['A','B','C','D','E','F','G','H','I','J']:
                ws.column_dimensions[c].width = 25

        out=io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, as_attachment=True, download_name="from_json_tables.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return f"Server Error: {str(e)}\n{traceback.format_exc()}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT",10000)))
