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

try:
    import fitz
    HAS_FITZ = True
except:
    HAS_FITZ = False

HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OCR Studio - Fast</title>
<link rel="icon" href="/favicon.ico">
<style>
*{box-sizing:border-box;font-family:Inter,system-ui,sans-serif}
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;background-image:url('/background.jpg');background-size:cover;background-position:center;background-color:#7a8cff}
.card{width:100%;max-width:720px;background:rgba(255,255,255,.28);backdrop-filter:blur(20px);border-radius:24px;padding:32px;box-shadow:0 8px 32px rgba(31,38,135,.2);border:1px solid rgba(255,255,255,.5)}
h2{margin:0 0 8px;text-align:center;color:#111;font-size:26px}
.sub{text-align:center;color:rgba(0,0,0,.7);font-size:13px;margin-bottom:20px}
.drop{border:1.5px dashed rgba(255,255,255,.9);border-radius:16px;padding:40px;text-align:center;background:rgba(255,255,255,.35);cursor:pointer;font-weight:600}
.progress{width:100%;height:8px;background:rgba(0,0,0,.1);border-radius:10px;margin-top:15px;overflow:hidden;display:none}
.bar{height:100%;width:0%;background:#111;transition:width.3s}
.preview{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:12px;margin-top:18px}
.item{border-radius:12px;overflow:hidden;background:rgba(255,255,255,.9);position:relative;padding:10px;text-align:center;font-size:11px}
.item img{width:100%;height:70px;object-fit:cover;border-radius:8px}
.item button{position:absolute;top:5px;right:5px;background:rgba(0,0,0,.7);color:#fff;border:none;border-radius:50%;width:22px;height:22px;cursor:pointer}
.btn{width:100%;padding:16px;background:#111;color:#fff;border:none;border-radius:14px;margin-top:18px;cursor:pointer;font-weight:700;display:none}
.btn.show{display:block}
.log{background:rgba(17,17,17,.9);color:#7CFF7C;padding:12px;border-radius:10px;margin-top:14px;font-family:monospace;font-size:12px;white-space:pre-wrap}
</style>
</head>
<body>
<div class="card">
<h2>OCR Studio - Fast</h2>
<div class="sub">PDF, JSON, Images - Optimized for speed</div>
<div class="drop" onclick="document.getElementById('files').click()">Click to select files<br><small>PDF, JSON, JPG, PNG</small><input type="file" id="files" multiple hidden accept=".pdf,.json,application/json,image/*"></div>
<div class="progress" id="progress"><div class="bar" id="bar"></div></div>
<div id="preview" class="preview"></div>
<button class="btn" id="processBtn" onclick="upload()">Process & Download Excel</button>
<div id="log" class="log">Ready - Fast mode enabled</div>
</div>
<script>
const input=document.getElementById('files'), preview=document.getElementById('preview'), btn=document.getElementById('processBtn'), log=document.getElementById('log'), bar=document.getElementById('bar'), prog=document.getElementById('progress');
let selectedFiles=[];
input.addEventListener('change', e=>{ selectedFiles=selectedFiles.concat(Array.from(e.target.files)); renderPreview(); });
function renderPreview(){
 preview.innerHTML=''; if(selectedFiles.length==0){ btn.classList.remove('show'); log.textContent='Ready'; return; }
 selectedFiles.forEach((file,i)=>{
   const div=document.createElement('div'); div.className='item';
   if(file.type.startsWith('image/')){ const img=document.createElement('img'); img.src=URL.createObjectURL(file); div.appendChild(img); }
   else if(file.name.endsWith('.pdf')){ div.innerHTML='<div style="font-size:28px">📕</div>PDF'; }
   else if(file.name.endsWith('.json')){ div.innerHTML='<div style="font-size:28px">⚡</div>JSON'; }
   const name=document.createElement('div'); name.textContent=file.name; name.style.marginTop='6px'; name.style.fontSize='10px';
   const rm=document.createElement('button'); rm.textContent='×'; rm.onclick=()=>{ selectedFiles.splice(i,1); renderPreview(); };
   div.appendChild(name); div.appendChild(rm); preview.appendChild(div);
 });
 btn.classList.add('show'); btn.textContent=`Process ${selectedFiles.length} file(s)`; log.textContent=`${selectedFiles.length} ready - Fast mode`;
}
async function upload(){
 if(selectedFiles.length==0) return; btn.disabled=true; prog.style.display='block'; bar.style.width='20%'; btn.textContent='Processing...'; log.textContent='Fast OCR running...';
 const fd=new FormData(); selectedFiles.forEach(f=>fd.append('files', f));
 try{
  bar.style.width='60%';
  const res=await fetch('/ocr',{method:'POST',body:fd});
  bar.style.width='90%';
  if(!res.ok){ const t=await res.text(); log.textContent=t.slice(0,4000); throw new Error('Server error'); }
  const blob=await res.blob(); bar.style.width='100%'; const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download='fast_extract.xlsx'; document.body.appendChild(a); a.click(); a.remove();
  log.textContent='✅ Done! Fast mode - Excel ready'; btn.textContent='✅ Downloaded';
  setTimeout(()=>{prog.style.display='none'; bar.style.width='0%'}, 1500);
 }catch(e){ log.textContent='Error: '+e.message; btn.textContent='Try Again'; prog.style.display='none'; }
 finally{ btn.disabled=false; }
}
</script>
</body>
</html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)
@app.route('/background.jpg')
def bg():
    if os.path.exists("background.jpg"): return send_file("background.jpg", mimetype="image/jpeg")
    return "",404
@app.route('/favicon.ico')
def fav():
    if os.path.exists("favicon.ico"): return send_file("favicon.ico", mimetype="image/x-icon")
    return "",404

def write_json_to_sheet(ws, data, start_row=1):
    if isinstance(data, list) and len(data)>0 and isinstance(data[0], dict):
        headers=list(data[0].keys())
        for c_idx,h in enumerate(headers, start=1):
            ws.cell(row=start_row, column=c_idx, value=h).font=Font(bold=True)
        for r_idx,row_data in enumerate(data, start=start_row+1):
            for c_idx,h in enumerate(headers, start=1):
                ws.cell(row=r_idx, column=c_idx, value=str(row_data.get(h,"")))
        return start_row+len(data)+1
    if isinstance(data, list) and len(data)>0 and isinstance(data[0], (list,tuple)):
        for r_idx,row in enumerate(data, start=start_row):
            for c_idx,val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=str(val))
        return start_row+len(data)+1
    if isinstance(data, dict):
        r=start_row
        for k,v in data.items():
            ws.cell(row=r, column=1, value=k).font=Font(bold=True)
            if isinstance(v, list) and len(v)>0 and isinstance(v[0], dict):
                r+=1
                r=write_json_to_sheet(ws, v, r)
                continue
            else:
                ws.cell(row=r, column=2, value=str(v) if not isinstance(v,(dict,list)) else json.dumps(v)[:20000])
            r+=1
        return r
    ws.cell(row=start_row, column=1, value=str(data))
    return start_row+1

def fast_ocr(pil_image):
    # FASTEST OCR
    if max(pil_image.size) > 1600:
        pil_image.thumbnail((1600, 1600))
    gray = pil_image.convert("L")
    # Speed config
    config = '--oem 1 --psm 6 -c tessedit_do_invert=0'
    data = pytesseract.image_to_data(gray, lang='eng', config=config, output_type=pytesseract.Output.DICT)
    lines={}
    n=len(data['text'])
    for i in range(n):
        txt=data['text'][i].strip()
        if not txt: continue
        try:
            conf=int(data['conf'][i])
        except: conf=0
        if conf < 35: continue
        key=data['line_num'][i]
        if key not in lines: lines[key]=[]
        lines[key].append({'text':txt,'left':data['left'][i],'width':data['width'][i]})
    rows=[]
    for k in sorted(lines.keys()):
        words=sorted(lines[k], key=lambda w:w['left'])
        if not words: continue
        avg_w=sum(w['width'] for w in words)/len(words)
        cols=[]; cur=words[0]['text']; last_end=words[0]['left']+words[0]['width']
        for w in words[1:]:
            gap=w['left']-last_end
            if gap > avg_w*1.0:
                cols.append(cur); cur=w['text']
            else: cur+=" "+w['text']
            last_end=w['left']+w['width']
        cols.append(cur)
        rows.append(cols)
    if not rows:
        text=pytesseract.image_to_string(gray, lang='eng', config='--oem 1 --psm 6')
        rows=[[l] for l in text.splitlines() if l.strip()]
    return rows

@app.route('/ocr', methods=['POST'])
def ocr():
    try:
        files=request.files.getlist('files') or request.files.getlist('images')
        if not files: return "No files",400
        wb=Workbook(); wb.remove(wb.active)

        for f in files:
            filename=f.filename
            ext=filename.rsplit('.',1)[-1].lower() if '.' in filename else ""
            raw=filename.rsplit('.',1)[0]
            safe="".join(c if c.isalnum() or c in (" ","_") else "_" for c in raw)[:31] or "Sheet1"
            ws=wb.create_sheet(title=safe)

            if ext=='json':
                try:
                    data=json.loads(f.read().decode('utf-8'))
                    ws['A1']=filename; ws['A1'].font=Font(bold=True,color="0000FF")
                    write_json_to_sheet(ws, data, start_row=3)
                except Exception as e:
                    ws['A1']=f"JSON Error: {e}"

            elif ext=='pdf':
                try:
                    file_bytes=f.read()
                    ws['A1']=filename; ws['A1'].font=Font(bold=True,color="0000FF")
                    row_num=3
                    if HAS_FITZ:
                        doc=fitz.open(stream=file_bytes, filetype="pdf")
                        for page_idx in range(min(len(doc), 20)): # Limit to 20 pages for speed
                            page=doc[page_idx]
                            text=page.get_text("text").strip()
                            if text and len(text) > 20: # Digital PDF - instant
                                ws.cell(row=row_num, column=1, value=f"Page {page_idx+1}").font=Font(bold=True)
                                row_num+=1
                                for line in text.splitlines():
                                    if line.strip():
                                        # Split table if many spaces
                                        if " " in line:
                                            parts=[p for p in line.split(" ") if p.strip()]
                                            if len(parts)>1:
                                                for c_idx,p in enumerate(parts, start=1):
                                                    ws.cell(row=row_num, column=c_idx, value=p.strip())
                                            else:
                                                ws.cell(row=row_num, column=1, value=line.strip())
                                        else:
                                            ws.cell(row=row_num, column=1, value=line.strip())
                                        row_num+=1
                            else: # Scanned - OCR (slower)
                                ws.cell(row=row_num, column=1, value=f"Page {page_idx+1} - OCR").font=Font(bold=True)
                                row_num+=1
                                pix=page.get_pixmap(dpi=120) # 120 dpi = FAST
                                img=Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                rows=fast_ocr(img)
                                for r in rows:
                                    for c_idx,val in enumerate(r, start=1):
                                        ws.cell(row=row_num, column=c_idx, value=val)
                                    row_num+=1
                            row_num+=1
                        doc.close()
                    else:
                        ws.cell(row=row_num, column=1, value="Add PyMuPDF to requirements.txt for PDF support")
                except Exception as e:
                    ws.cell(row=3, column=1, value=f"PDF Error: {e}")

            else:
                try:
                    img=Image.open(f.stream)
                    rows=fast_ocr(img)
                    ws['A1']=filename; ws['A1'].font=Font(bold=True,color="0000FF")
                    r=3
                    for row_data in rows:
                        for c_idx,val in enumerate(row_data, start=1):
                            ws.cell(row=r, column=c_idx, value=val)
                        r+=1
                except Exception as e:
                    ws.cell(row=3, column=1, value=f"Image Error: {e}")

            for c in ['A','B','C','D','E','F','G','H','I','J']:
                ws.column_dimensions[c].width=25

        out=io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, as_attachment=True, download_name="fast_tables.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return f"Server Error: {e}\n{traceback.format_exc()}",500

if __name__=='__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT",10000)))
