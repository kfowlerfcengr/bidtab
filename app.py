import os, io, json, re, tempfile
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string

import anthropic
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

app = Flask(__name__)
OUTPUT_DIR = tempfile.mkdtemp(prefix="bidtab_")

YELLOW = PatternFill("solid", fgColor="FFFF00")

# These field keys will always be treated as numbers when found
NUMERIC_KEYWORDS = [
    "price", "cost", "qty", "quantity", "total", "amount",
    "adder", "witnessed", "test_cert", "fee"
]

def is_numeric_field(key):
    return any(kw in key for kw in NUMERIC_KEYWORDS)

def build_row_map(ws):
    """
    Dynamically build {row_number: field_key} by reading column A labels
    from the template. Works for ANY bid tab template regardless of equipment type.
    Also detects which rows are vendor data rows (have editable cells in cols C/D/E).
    Returns: (row_map, vendor_col_start)
    """
    from openpyxl.cell.cell import MergedCell
    row_map = {}

    for row in ws.iter_rows():
        # Get label from col A
        label_cell = row[0] if row else None
        if not label_cell or isinstance(label_cell, MergedCell):
            continue
        label = str(label_cell.value or "").strip()
        if not label or label.lower() in ("none", ""):
            continue

        # Convert label to a clean field key
        key = label.lower()
        key = re.sub(r'[^a-z0-9\s]', '', key)   # remove special chars
        key = re.sub(r'\s+', '_', key.strip())   # spaces to underscores
        key = re.sub(r'^_+|_+$', '', key)        # trim underscores
        key = re.sub(r'_+', '_', key)            # collapse doubles

        if key:
            row_map[label_cell.row] = key

    return row_map

def build_system_prompt(row_map):
    """Build a Claude system prompt from the actual template field keys."""
    field_list = ", ".join(sorted(set(row_map.values())))
    return f"""You are a technical bid tab assistant for an engineering procurement team.

You receive a DATA SHEET and up to 3 VENDOR QUOTES.
Extract all available data and return ONLY valid JSON — no markdown, no explanation.

Return this exact structure:
{{
  "datasheet": {{ ...fields from the data sheet... }},
  "vendor1": {{ ...fields from vendor 1... }},
  "vendor2": {{ ...fields from vendor 2... }},
  "vendor3": {{ ...fields from vendor 3... }}
}}

The template has these field keys — extract values for as many as you can find:
{field_list}

Rules:
- datasheet key = values from the DATA SHEET only
- vendor1/2/3 = values from each vendor quote
- Any field containing "price", "cost", "qty", "quantity", "total", "amount", or "adder" must be a number (no $ or commas) or null
- null for any field not found
- Be thorough — extract every spec, price, note, and commercial term you can find
- Return ONLY the JSON object"""


def extract_text(file_storage) -> str:
    name = (file_storage.filename or "").lower()
    data = file_storage.read()
    if name.endswith(".pdf"):
        if not HAS_PDF:
            return "[PDF unavailable]"
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n\n".join(p.extract_text() or "" for p in pdf.pages)
    if name.endswith((".xlsx", ".xls")):
        wb = load_workbook(io.BytesIO(data), data_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"[Sheet: {ws.title}]")
            for row in ws.iter_rows(values_only=True):
                parts = [str(c) if c is not None else "" for c in row]
                if any(p.strip() for p in parts):
                    lines.append("\t".join(parts))
        return "\n".join(lines)
    return data.decode("utf-8", errors="replace")


def coerce(val, field):
    if val is None: return None
    s = str(val).strip()
    if not s or s.lower() == "null": return None
    if is_numeric_field(field):
        try:
            f = float(s.replace(",","").replace("$","").replace(" ",""))
            return int(f) if f == int(f) else f
        except: return s
    return s


def fill_excel(template_bytes, data, pi):
    from openpyxl.cell.cell import MergedCell
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Dynamically build row map from this template's column A labels
    row_map = build_row_map(ws)

    # Detect vendor columns — find row 7 (or first row with 3+ filled cells after col B)
    # Default: B=datasheet, C=vendor1, D=vendor2, E=vendor3
    col_map = {"datasheet":"B", "vendor1":"C", "vendor2":"D", "vendor3":"E"}

    # Detect header rows by looking for Client/Project labels
    # Try to write project info into common header positions
    header_written = False
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if isinstance(cell, MergedCell): continue
            lbl = str(cell.value or "").strip().lower()
            # Find vendor name row (row with "vendor 1" or similar in C/D/E)
            if lbl in ("client:", "client"):
                # Write project info relative to this position
                r = cell.row
                for addr, val in {
                    f"C{r}": pi.get("client",""),
                    f"C{r+1}": pi.get("project_no",""),
                    f"C{r+2}": pi.get("project_name",""),
                    f"E{r+2}": datetime.today().strftime("%m/%d/%Y"),
                    f"C{r+3}": pi.get("location",""),
                    f"F{r+3}": pi.get("author",""),
                }.items():
                    c = ws[addr]
                    if val and not isinstance(c, MergedCell): c.value = val
                header_written = True
                break
        if header_written: break

    # Find the vendor header row (where vendor names go)
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if isinstance(cell, MergedCell): continue
            lbl = str(cell.value or "").strip().lower()
            if lbl in ("equipment/item:", "equipment/item", "equipment:", "item:"):
                r = cell.row
                # Write equipment and vendor names in cols C, D, E of this row
                for col_letter, key in [("B", "equipment"), ("C","vendor1_name"), ("D","vendor2_name"), ("E","vendor3_name")]:
                    c = ws[f"{col_letter}{r}"]
                    val = pi.get(key,"")
                    if val and not isinstance(c, MergedCell): c.value = val
                break

    # Fill data rows using the dynamic row map
    filled = 0
    for row_num, field in row_map.items():
        for src, col in col_map.items():
            val = coerce((data.get(src) or {}).get(field), field)
            cell = ws[f"{col}{row_num}"]
            if isinstance(cell, MergedCell):
                continue
            if val is not None:
                cell.value = val
                filled += 1
            elif col in ("C","D","E"):
                curr = cell.value
                if curr is None or str(curr).strip() in ("","-","By Vendor","N/A","by vendor","n/a"):
                    cell.fill = YELLOW

    # Try to find and restore TOTAL formula rows
    # Look for a row labeled "TOTAL" and write SUM formulas in vendor cols
    for row_num, field in row_map.items():
        if "total" in field:
            for col in ("C","D","E"):
                cell = ws[f"{col}{row_num}"]
                if not isinstance(cell, MergedCell):
                    # Find price rows above to sum
                    cell.value = f"=SUM({col}{row_num-10}:{col}{row_num-1})"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read(), row_map


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Bid Tab Agent</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
:root{
  --ink:#0f1117;--ink2:#3a3d4a;--ink3:#6b7080;
  --line:#e2e4ea;--line2:#c8cad4;--bg:#f7f8fa;--bg2:#fff;
  --accent:#1a56db;--accent2:#ddeaff;
  --green:#0d6e3b;--gbg:#e6f4ed;
  --red:#9b1c1c;--rbg:#fde8e8;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:var(--sans);background:var(--bg);color:var(--ink);font-size:14px;line-height:1.5;}
.topbar{background:var(--ink);color:#fff;height:54px;padding:0 36px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;}
.brand{font-family:var(--mono);font-size:13px;letter-spacing:.06em;}
.brand em{color:#6b9fff;font-style:normal;}
.badge{font-family:var(--mono);font-size:11px;padding:4px 12px;border-radius:20px;background:rgba(255,255,255,.1);color:rgba(255,255,255,.6);}
.page{max-width:780px;margin:0 auto;padding:36px 20px 80px;}
.card{background:var(--bg2);border:1px solid var(--line);border-radius:12px;margin-bottom:18px;overflow:hidden;}
.ch{padding:15px 22px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:12px;}
.ch-ico{width:28px;height:28px;border-radius:50%;background:var(--accent);color:#fff;font-family:var(--mono);font-size:11px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-weight:500;}
.ch h2{font-size:14px;font-weight:600;margin-bottom:2px;}
.ch p{font-size:11px;color:var(--ink3);}
.cb{padding:18px 22px;}
.igrid{display:grid;grid-template-columns:1fr 1fr;gap:13px;}
.fld label{display:block;font-size:10px;font-family:var(--mono);color:var(--ink3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;}
.fld input{width:100%;border:1px solid var(--line2);border-radius:7px;padding:8px 11px;font-family:var(--sans);font-size:13px;color:var(--ink);outline:none;}
.fld input:focus{border-color:var(--accent);}
.uz{border:2px dashed var(--line2);border-radius:10px;padding:24px 18px;text-align:center;cursor:pointer;transition:all .2s;background:var(--bg);}
.uz:hover,.uz.drag{border-color:var(--accent);background:var(--accent2);}
.uz input{display:none;}
.uz .uico{font-size:24px;margin-bottom:6px;}
.uz h4{font-size:13px;font-weight:500;margin-bottom:2px;}
.uz p{font-size:11px;color:var(--ink3);}
.fchip{display:inline-flex;align-items:center;gap:6px;background:var(--gbg);border:1px solid #a3d9b8;color:var(--green);border-radius:6px;padding:4px 10px;font-family:var(--mono);font-size:11px;margin-top:6px;}
.vgrid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}
.vcard{border:1px solid var(--line2);border-radius:10px;overflow:hidden;}
.vcard.loaded{border-color:#a3d9b8;}
.vhead{background:var(--bg);padding:8px 12px;border-bottom:1px solid var(--line);font-family:var(--mono);font-size:10px;color:var(--ink3);text-transform:uppercase;letter-spacing:.05em;}
.vhead.loaded{background:var(--gbg);color:var(--green);border-bottom-color:#a3d9b8;}
.vbody{padding:11px 12px;background:var(--bg2);}
.vbody input{width:100%;border:1px solid var(--line2);border-radius:6px;padding:7px 10px;font-size:13px;font-family:var(--sans);color:var(--ink);outline:none;margin-bottom:8px;}
.vbody input:focus{border-color:var(--accent);}
.vup{width:100%;padding:7px;border:1px dashed var(--line2);border-radius:6px;background:var(--bg);font-size:11px;color:var(--ink3);cursor:pointer;text-align:center;font-family:var(--sans);transition:all .15s;}
.vup:hover{border-color:var(--accent);color:var(--accent);background:var(--accent2);}
.vfname{font-family:var(--mono);font-size:10px;color:var(--green);margin-top:5px;min-height:13px;}
.sub-area{margin-top:24px;text-align:center;}
.sub-btn{background:var(--accent);color:#fff;border:none;padding:13px 42px;border-radius:8px;font-size:15px;font-weight:500;font-family:var(--sans);cursor:pointer;transition:background .15s;}
.sub-btn:hover{background:#1448c0;}
.sub-btn:disabled{background:#aaa;cursor:not-allowed;}
.sub-note{font-size:11px;color:var(--ink3);margin-top:8px;}
.rbox{background:var(--gbg);border:1px solid #a3d9b8;border-radius:10px;padding:22px;margin-top:22px;display:none;text-align:center;}
.rbox.show{display:block;}
.rbox h3{font-size:15px;font-weight:600;color:var(--green);margin-bottom:6px;}
.rbox p{font-size:12px;color:#1a5c35;margin-bottom:14px;}
.dl-btn{display:inline-flex;align-items:center;gap:7px;background:var(--green);color:#fff;padding:10px 26px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:500;}
.dl-btn:hover{background:#0a5a30;}
.ebox{background:var(--rbg);border:1px solid #f5c2c2;border-radius:10px;padding:14px 18px;margin-top:18px;display:none;}
.ebox.show{display:block;}
.ebox p{color:var(--red);font-size:13px;}
.overlay{position:fixed;inset:0;background:rgba(15,17,23,.72);display:none;align-items:center;justify-content:center;z-index:200;}
.overlay.show{display:flex;}
.pbox{background:#fff;border-radius:14px;padding:38px 46px;text-align:center;max-width:360px;width:90%;}
.spin{width:42px;height:42px;border:3px solid var(--line);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 18px;}
@keyframes spin{to{transform:rotate(360deg);}}
.pbox h3{font-size:15px;font-weight:600;margin-bottom:5px;}
.pbox p{font-size:12px;color:var(--ink3);}
.pstep{font-family:var(--mono);font-size:11px;color:var(--accent);margin-top:10px;}
</style>
</head>
<body>
<div class="topbar">
  <div class="brand">BID<em>/</em>TAB <span style="opacity:.4;font-size:11px;font-weight:300;">Agent</span></div>
  <div class="badge" id="badge">Ready</div>
</div>
<div class="page">

  <div class="card">
    <div class="ch"><div class="ch-ico">📁</div><div><h2>Project Info</h2><p>Fills the header of your bid tab</p></div></div>
    <div class="cb">
      <div class="igrid">
        <div class="fld"><label>Client</label><input id="pi-client" placeholder="e.g. FTAI"></div>
        <div class="fld"><label>Project No.</label><input id="pi-projno" placeholder="e.g. FTA260107"></div>
        <div class="fld"><label>Project Name</label><input id="pi-name" placeholder="e.g. FTAI CFM56 Package"></div>
        <div class="fld"><label>Project Location</label><input id="pi-loc" placeholder="e.g. Mobile"></div>
        <div class="fld"><label>Equipment / Item</label><input id="pi-equip" placeholder="e.g. Pumps"></div>
        <div class="fld"><label>Author</label><input id="pi-author" placeholder="Your name"></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="ch"><div class="ch-ico">1</div><div><h2>Bid Tab Template</h2><p>Your blank .xlsx — formatting, colors and merges will be preserved exactly</p></div></div>
    <div class="cb">
      <div class="uz" id="z-tmpl" onclick="document.getElementById('f-tmpl').click()"
           ondragover="zo(event,'z-tmpl')" ondragleave="zl('z-tmpl')" ondrop="zd(event,'z-tmpl','f-tmpl')">
        <input type="file" id="f-tmpl" accept=".xlsx" onchange="chip(this,'chip-tmpl')">
        <div class="uico">📋</div><h4>Upload your blank bid tab template (.xlsx)</h4>
        <p>Required — all original formatting will be preserved</p>
        <div id="chip-tmpl"></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="ch"><div class="ch-ico">2</div><div><h2>Data Sheet</h2><p>Engineering spec — fills F&C standards column</p></div></div>
    <div class="cb">
      <div class="uz" id="z-ds" onclick="document.getElementById('f-ds').click()"
           ondragover="zo(event,'z-ds')" ondragleave="zl('z-ds')" ondrop="zd(event,'z-ds','f-ds')">
        <input type="file" id="f-ds" accept=".xlsx,.xls,.csv,.txt,.pdf" onchange="chip(this,'chip-ds')">
        <div class="uico">📄</div><h4>Upload data sheet</h4><p>Excel, PDF, CSV or text</p>
        <div id="chip-ds"></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="ch"><div class="ch-ico">3</div><div><h2>Vendor Quotes</h2><p>Any format, any layout — AI figures it out</p></div></div>
    <div class="cb">
      <div class="vgrid">
        <div class="vcard" id="vc0"><div class="vhead" id="vh0">Vendor 1</div>
          <div class="vbody">
            <input id="vn0" placeholder="Vendor name" value="Vendor 1" oninput="document.getElementById('vh0').textContent=this.value||'Vendor 1'">
            <div class="vup" onclick="document.getElementById('vf0').click()">+ Attach quote</div>
            <input type="file" id="vf0" style="display:none" accept=".xlsx,.xls,.csv,.txt,.pdf" onchange="vfile(this,0)">
            <div class="vfname" id="vfn0"></div>
          </div>
        </div>
        <div class="vcard" id="vc1"><div class="vhead" id="vh1">Vendor 2</div>
          <div class="vbody">
            <input id="vn1" placeholder="Vendor name" value="Vendor 2" oninput="document.getElementById('vh1').textContent=this.value||'Vendor 2'">
            <div class="vup" onclick="document.getElementById('vf1').click()">+ Attach quote</div>
            <input type="file" id="vf1" style="display:none" accept=".xlsx,.xls,.csv,.txt,.pdf" onchange="vfile(this,1)">
            <div class="vfname" id="vfn1"></div>
          </div>
        </div>
        <div class="vcard" id="vc2"><div class="vhead" id="vh2">Vendor 3</div>
          <div class="vbody">
            <input id="vn2" placeholder="Vendor name" value="Vendor 3" oninput="document.getElementById('vh2').textContent=this.value||'Vendor 3'">
            <div class="vup" onclick="document.getElementById('vf2').click()">+ Attach quote</div>
            <input type="file" id="vf2" style="display:none" accept=".xlsx,.xls,.csv,.txt,.pdf" onchange="vfile(this,2)">
            <div class="vfname" id="vfn2"></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="sub-area">
    <button class="sub-btn" id="sub-btn" onclick="run()">Generate Bid Tab</button>
    <div class="sub-note">Server reads your files with AI → returns a perfectly formatted .xlsx</div>
  </div>

  <div class="rbox" id="rbox">
    <h3>✓ Bid tab ready</h3>
    <p id="rmsg"></p>
    <a class="dl-btn" id="dl-link" href="#">⬇ Download Filled Bid Tab (.xlsx)</a>
  </div>
  <div class="ebox" id="ebox"><p id="emsg"></p></div>
</div>

<div class="overlay" id="overlay">
  <div class="pbox">
    <div class="spin"></div>
    <h3>Working on it…</h3>
    <p id="pmsg">Reading your files</p>
    <div class="pstep" id="pstep"></div>
  </div>
</div>

<script>
function zo(e,id){e.preventDefault();document.getElementById(id).classList.add('drag');}
function zl(id){document.getElementById(id).classList.remove('drag');}
function zd(e,zid,fid){e.preventDefault();zl(zid);const f=e.dataTransfer.files[0];if(!f)return;const inp=document.getElementById(fid);const dt=new DataTransfer();dt.items.add(f);inp.files=dt.files;chip(inp,fid==='f-ds'?'chip-ds':fid==='f-tmpl'?'chip-tmpl':null);}
function chip(inp,chipId){const f=inp.files[0];if(!f||!chipId)return;document.getElementById(chipId).innerHTML=`<div class="fchip">✓ ${f.name}</div>`;}
function vfile(inp,idx){const f=inp.files[0];if(!f)return;document.getElementById('vfn'+idx).textContent='✓ '+f.name;document.getElementById('vc'+idx).classList.add('loaded');document.getElementById('vh'+idx).classList.add('loaded');}

async function run(){
  const tmplFile=document.getElementById('f-tmpl').files[0];
  if(!tmplFile){alert('Please upload your bid tab template (.xlsx).');return;}
  const dsFile=document.getElementById('f-ds').files[0];
  const vFiles=[0,1,2].map(i=>document.getElementById('vf'+i).files[0]||null);
  if(!dsFile&&!vFiles.some(Boolean)){alert('Please upload at least a data sheet or one vendor quote.');return;}

  document.getElementById('sub-btn').disabled=true;
  document.getElementById('rbox').classList.remove('show');
  document.getElementById('ebox').classList.remove('show');
  document.getElementById('overlay').classList.add('show');
  document.getElementById('badge').textContent='Processing…';

  const msgs=['Reading files…','AI extracting specs…','Mapping vendor data…','Writing Excel…'];
  let mi=0;
  const ticker=setInterval(()=>{document.getElementById('pmsg').textContent=msgs[Math.min(mi,msgs.length-1)];document.getElementById('pstep').textContent=`Step ${Math.min(mi+1,4)} of 4`;mi++;},5000);

  try{
    const fd=new FormData();
    fd.append('template',tmplFile);
    if(dsFile)fd.append('datasheet',dsFile);
    [0,1,2].forEach(i=>{
      fd.append(`vendor_name_${i+1}`,document.getElementById('vn'+i).value||`Vendor ${i+1}`);
      if(vFiles[i])fd.append(`vendor_${i+1}`,vFiles[i]);
    });
    fd.append('client',document.getElementById('pi-client').value);
    fd.append('project_no',document.getElementById('pi-projno').value);
    fd.append('project_name',document.getElementById('pi-name').value);
    fd.append('location',document.getElementById('pi-loc').value);
    fd.append('equipment',document.getElementById('pi-equip').value);
    fd.append('author',document.getElementById('pi-author').value);

    const resp=await fetch('/generate',{method:'POST',body:fd});
    const result=await resp.json();
    clearInterval(ticker);
    document.getElementById('overlay').classList.remove('show');
    document.getElementById('sub-btn').disabled=false;
    document.getElementById('badge').textContent='Ready';

    if(result.error){
      document.getElementById('emsg').textContent='Error: '+result.error;
      document.getElementById('ebox').classList.add('show');
    } else {
      document.getElementById('rmsg').textContent=result.message;
      document.getElementById('dl-link').href='/download/'+result.filename;
      document.getElementById('dl-link').download=result.filename;
      document.getElementById('rbox').classList.add('show');
    }
  }catch(err){
    clearInterval(ticker);
    document.getElementById('overlay').classList.remove('show');
    document.getElementById('sub-btn').disabled=false;
    document.getElementById('badge').textContent='Error';
    document.getElementById('emsg').textContent='Error: '+err.message;
    document.getElementById('ebox').classList.add('show');
  }
}
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY not set on server."})

    template_file = request.files.get("template")
    if not template_file or not template_file.filename:
        return jsonify({"error": "No template uploaded."})

    template_bytes = template_file.read()

    pi = {
        "client":       request.form.get("client", ""),
        "project_no":   request.form.get("project_no", ""),
        "project_name": request.form.get("project_name", ""),
        "location":     request.form.get("location", ""),
        "equipment":    request.form.get("equipment", ""),
        "author":       request.form.get("author", ""),
        "vendor1_name": request.form.get("vendor_name_1", "Vendor 1"),
        "vendor2_name": request.form.get("vendor_name_2", "Vendor 2"),
        "vendor3_name": request.form.get("vendor_name_3", "Vendor 3"),
    }

    # Read template to build dynamic row map and system prompt
    try:
        wb_tmp = load_workbook(io.BytesIO(template_bytes))
        ws_tmp = wb_tmp.active
        row_map = build_row_map(ws_tmp)
        system_prompt = build_system_prompt(row_map)
    except Exception as e:
        return jsonify({"error": f"Could not read template: {e}"})

    # Build content for Claude
    content = ""
    ds = request.files.get("datasheet")
    if ds and ds.filename:
        content += f"=== DATA SHEET ({ds.filename}) ===\n{extract_text(ds)}\n\n"

    for i in (1, 2, 3):
        vf = request.files.get(f"vendor_{i}")
        vn = request.form.get(f"vendor_name_{i}", f"Vendor {i}")
        if vf and vf.filename:
            content += f"=== VENDOR {i} — {vn} ({vf.filename}) ===\n{extract_text(vf)}\n\n"

    if not content.strip():
        return jsonify({"error": "No data sheet or vendor quotes uploaded."})

    # Call Claude with template-specific prompt
    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": content}]
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```json\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"AI returned malformed JSON: {e}"})
    except Exception as e:
        return jsonify({"error": f"AI error: {e}"})

    # Fill Excel using the dynamic row map
    try:
        xlsx_bytes, _ = fill_excel(template_bytes, data, pi)
    except Exception as e:
        return jsonify({"error": f"Excel error: {e}"})

    # Save output
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    proj = (pi.get("project_no") or "BidTab").replace(" ", "_")
    fname = f"Bid_Tab_{proj}_{stamp}.xlsx"
    fpath = os.path.join(OUTPUT_DIR, fname)
    with open(fpath, "wb") as f:
        f.write(xlsx_bytes)

    file_count = len([x for x in [ds] + [request.files.get(f"vendor_{i}") for i in (1,2,3)] if x and x.filename])
    msg = f"Filled {file_count} files across {len(row_map)} template fields. Yellow cells = not found in vendor quote."
    return jsonify({"filename": fname, "message": msg})


@app.route("/download/<fname>")
def download(fname):
    fpath = os.path.join(OUTPUT_DIR, fname)
    if not os.path.exists(fpath):
        return "File not found", 404
    return send_file(fpath, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  Bid Tab Agent → http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)