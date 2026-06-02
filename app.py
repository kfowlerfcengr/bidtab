import os, io, json, re, tempfile
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string

import anthropic
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string

try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

app = Flask(__name__)
OUTPUT_DIR = tempfile.mkdtemp(prefix="bidtab_")

YELLOW = PatternFill("solid", fgColor="FFFF00")

ROW_MAP = {
    8:"supplier", 9:"quotation_number", 10:"quotation_date", 11:"quotation_validity",
    13:"tag_number", 14:"service", 15:"datasheet_doc_no", 16:"datasheet_revision",
    17:"pump_type", 18:"pump_size_qty", 19:"pump_manufacturer", 20:"pump_model",
    21:"flow_capacity", 22:"suction_pressure", 23:"discharge_pressure",
    24:"npsh_available", 25:"operating_temp", 26:"specific_gravity",
    27:"viscosity", 28:"vapor_pressure", 29:"corrosion_erosion_causes",
    31:"proposal_curve_no", 32:"speed_rpm", 33:"npsh_required", 34:"rated_bhp",
    35:"max_bhp_rated_impeller", 36:"max_head_rated_impeller",
    37:"suction_specific_speed", 38:"efficiency",
    40:"max_working_pressure", 41:"hydrostatic_test_pressure",
    42:"impeller_diameter_criteria", 43:"mechanical_seal_type",
    44:"suction_nozzle_size", 45:"suction_vents_drains",
    46:"suction_pressure_gauge", 47:"discharge_nozzle_size",
    49:"nace_mr0175", 50:"lethal_service", 51:"vibration_sensor_type",
    53:"seal_flush_plans", 54:"external_seal_flush_flow",
    56:"materials_class", 57:"barrel_case", 58:"impeller",
    59:"case_wear_rings", 60:"shaft_sleeve",
    61:"motor_type", 63:"motor_manufacturer", 64:"motor_model",
    65:"horsepower", 66:"motor_rpm", 67:"motor_duty",
    68:"motor_efficiency_type", 69:"enclosure", 70:"volts_phase_hertz",
    71:"safety_factor", 72:"variable_frequency",
    74:"performance_test", 75:"hydrostatic_test", 76:"options",
    78:"location", 79:"max_min_site_temp", 80:"area_classification", 81:"max_sound",
    85:"technical_recommendation",
    90:"adder_hydrostatic_cert", 91:"adder_witnessed_hydrostatic",
    92:"adder_npshr_test", 93:"adder_witnessed_npshr",
    94:"adder_api674_test", 95:"adder_witnessed_api674",
    100:"pump_motor_price", 101:"vfd_price", 103:"quantity",
    108:"lead_time", 113:"commercial_recommendation",
    118:"payment_terms", 119:"payment_schedule",
    122:"manufacturing_location", 123:"delivery_terms",
    124:"delivery_location", 125:"additional_freight",
    128:"warranty", 130:"comments",
}

NUMERIC = {
    "pump_motor_price","vfd_price","quantity",
    "adder_hydrostatic_cert","adder_witnessed_hydrostatic",
    "adder_npshr_test","adder_witnessed_npshr",
    "adder_api674_test","adder_witnessed_api674",
}

SYSTEM = """You are a technical bid tab assistant for an engineering procurement team.

You receive a DATA SHEET and up to 3 VENDOR QUOTES (marked === VENDOR 1 ===, === VENDOR 2 ===, === VENDOR 3 ===).
Extract ALL available data from EVERY section and return ONLY valid JSON — no markdown, no explanation, no extra text.

Return this exact top-level structure:
{
  "datasheet": { ...all fields found in the data sheet... },
  "vendor1":   { ...all fields found in vendor 1 quote... },
  "vendor2":   { ...all fields found in vendor 2 quote... },
  "vendor3":   { ...all fields found in vendor 3 quote... }
}

Use EXACTLY these field keys (descriptions in parentheses are for your reference only — do not include them):
  supplier                  (vendor/manufacturer company name)
  quotation_number          (quote or proposal number/reference)
  quotation_date            (date of quotation)
  quotation_validity        (validity period of quote)
  tag_number                (equipment tag number or item number)
  service                   (pump service description)
  datasheet_doc_no          (data sheet document number)
  datasheet_revision        (data sheet revision letter/number)
  pump_type                 (e.g. centrifugal, API 610, between bearings, etc.)
  pump_size_qty             (pump size and quantity)
  pump_manufacturer         (pump maker name)
  pump_model                (pump model or series designation)
  flow_capacity             (rated flow with units, e.g. "500 GPM")
  suction_pressure          (suction pressure with units)
  discharge_pressure        (discharge pressure with units)
  npsh_available            (NPSHa with units)
  operating_temp            (operating temperature with units)
  specific_gravity          (fluid specific gravity)
  viscosity                 (fluid viscosity with units)
  vapor_pressure            (vapor pressure with units)
  corrosion_erosion_causes  (corrosive or erosive agents in fluid)
  proposal_curve_no         (performance curve number)
  speed_rpm                 (rated pump speed in RPM)
  npsh_required             (NPSHr with units)
  rated_bhp                 (brake horsepower at rated point)
  max_bhp_rated_impeller    (max BHP at rated impeller diameter)
  max_head_rated_impeller   (max head at rated impeller)
  suction_specific_speed    (suction specific speed value)
  efficiency                (pump efficiency at rated point, %)
  max_working_pressure      (maximum allowable working pressure)
  hydrostatic_test_pressure (hydrostatic test pressure)
  impeller_diameter_criteria (impeller diameter / trim criteria)
  mechanical_seal_type      (seal type, API plan, arrangement)
  suction_nozzle_size       (suction flange size and rating)
  suction_vents_drains      (vent/drain provisions on suction)
  suction_pressure_gauge    (suction pressure gauge provided?)
  discharge_nozzle_size     (discharge flange size and rating)
  nace_mr0175               (NACE MR0175 compliance: Yes/No)
  lethal_service            (lethal service designation: Yes/No)
  vibration_sensor_type     (vibration monitoring sensor type)
  seal_flush_plans          (API seal flush plan numbers)
  external_seal_flush_flow  (external flush flow rate/source)
  materials_class           (materials class designation)
  barrel_case               (casing/barrel material)
  impeller                  (impeller material)
  case_wear_rings           (wear ring material)
  shaft_sleeve              (shaft sleeve material)
  motor_type                (motor type, e.g. TEFC induction)
  motor_manufacturer        (motor maker name)
  motor_model               (motor model)
  horsepower                (motor rated horsepower)
  motor_rpm                 (motor synchronous or nameplate RPM)
  motor_duty                (motor duty rating, e.g. continuous)
  motor_efficiency_type     (efficiency class, e.g. NEMA Premium)
  enclosure                 (motor enclosure type, e.g. TEFC, TEAC)
  volts_phase_hertz         (electrical supply, e.g. "460V/3Ph/60Hz")
  safety_factor             (motor service factor)
  variable_frequency        (VFD provided or compatible: Yes/No)
  performance_test          (performance test type included)
  hydrostatic_test          (hydrostatic test type included)
  options                   (optional items or extras quoted)
  location                  (installation site/location)
  max_min_site_temp         (maximum and minimum ambient temperature at site)
  area_classification       (electrical area classification, e.g. Class I Div 2)
  max_sound                 (maximum allowable sound level)
  technical_recommendation  (engineer technical notes or recommendation)
  commercial_recommendation (commercial evaluation notes or recommendation)
  pump_motor_price          (total pump+motor unit price — NUMBER only, no $ or commas)
  vfd_price                 (VFD / variable frequency drive price — NUMBER only)
  quantity                  (quantity of units — NUMBER only)
  lead_time                 (delivery lead time)
  payment_terms             (payment terms, e.g. Net 30)
  payment_schedule          (milestone payment schedule)
  manufacturing_location    (where equipment is manufactured)
  delivery_terms            (delivery/shipping terms, e.g. FOB)
  delivery_location         (delivery destination)
  additional_freight        (extra freight or shipping charges)
  warranty                  (warranty period and terms)
  comments                  (any other notable comments or clarifications)
  adder_hydrostatic_cert    (adder price for hydrostatic certification — NUMBER only)
  adder_witnessed_hydrostatic (adder price for witnessed hydrostatic — NUMBER only)
  adder_npshr_test          (adder price for NPSHr test — NUMBER only)
  adder_witnessed_npshr     (adder price for witnessed NPSHr — NUMBER only)
  adder_api674_test         (adder price for API 674 test — NUMBER only)
  adder_witnessed_api674    (adder price for witnessed API 674 test — NUMBER only)

Critical rules:
- Use null for any field not found in the source document — never omit a key.
- For NUMBER-only fields: strip all $ signs, commas, and units; return a bare number or null.
- Be thorough: scan every table, section header, and footnote. Do not skip data.
- vendor2 and vendor3 should be populated only from their respective === VENDOR N === sections; use null for all fields if that vendor section is absent.
- Return ONLY the JSON object. No markdown fences, no commentary."""


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
    if field in NUMERIC:
        try:
            f = float(s.replace(",","").replace("$","").replace(" ",""))
            return int(f) if f == int(f) else f
        except: return s
    return s


def _repair_json(s):
    """Best-effort repair of truncated JSON by closing open structures."""
    # Track nesting so we can close what was opened
    stack = []
    in_string = False
    escape_next = False
    for ch in s:
        if escape_next:
            escape_next = False
            continue
        if ch == "\\" and in_string:
            escape_next = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch in "{[":
            stack.append("}" if ch == "{" else "]")
        elif ch in "}]":
            if stack and stack[-1] == ch:
                stack.pop()
    # If we're mid-string, close it
    if in_string:
        s += '"'
    # Close any trailing comma before closing braces
    s = re.sub(r",\s*$", "", s.rstrip())
    # Close all open structures
    s += "".join(reversed(stack))
    return s


def _master_addr(ws, addr):
    """Return the address of the master cell for a merged cell (for logging)."""
    for merge_range in ws.merged_cells.ranges:
        if addr in merge_range:
            from openpyxl.utils import get_column_letter
            return f"{get_column_letter(merge_range.min_col)}{merge_range.min_row}"
    return addr


def resolve_cell(ws, addr):
    """Return the writable master cell for addr, or None if it's a cross-column merge.

    - Cell merged across rows (same column): return the master cell so we can write to it.
    - Cell merged across columns (different column): return None — writing to the master
      cell of another column would corrupt the layout, so the caller should skip it.
    """
    cell = ws[addr]
    if not isinstance(cell, MergedCell):
        return cell
    col_letter = re.match(r'([A-Za-z]+)', addr).group(1).upper()
    target_col = column_index_from_string(col_letter)
    for merge_range in ws.merged_cells.ranges:
        if addr in merge_range:
            if merge_range.min_col == target_col:
                # Merged vertically (same column) — write to master row
                return ws.cell(merge_range.min_row, merge_range.min_col)
            else:
                # Merged horizontally into another column — skip
                return None
    return cell  # fallback: unrecognised merge, try as-is


def fill_excel(template_bytes, data, pi, vendor_count=3):
    from openpyxl.utils import get_column_letter
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Dump template row labels so we can verify ROW_MAP matches the actual file
    print("\n=== TEMPLATE ROW STRUCTURE (cols A-B, rows 1-135) ===")
    for r in range(1, 136):
        a = ws[f"A{r}"]
        b = ws[f"B{r}"]
        av = a.value if not isinstance(a, MergedCell) else f"[merged→{_master_addr(ws,'A'+str(r))}]"
        bv = b.value if not isinstance(b, MergedCell) else f"[merged→{_master_addr(ws,'B'+str(r))}]"
        if av or bv:
            print(f"  row {r:3d} | A={repr(av)!s:45s} | B={repr(bv)}")
    print("=====================================================\n")

    # Build dynamic column map: datasheet→B, vendor1→C, vendor2→D, vendor3→E, vendor4→F …
    # Column index: datasheet=2(B), vendor1=3(C), vendor2=4(D), …
    col_map = {"datasheet": "B"}
    vendor_cols = []  # column letters for all vendor columns
    for i in range(1, vendor_count + 1):
        col_letter = get_column_letter(2 + i)   # i=1→C(3), i=2→D(4), …
        col_map[f"vendor{i}"] = col_letter
        vendor_cols.append(col_letter)

    # Header — fixed project info fields
    for addr, val in {
        "C2": pi.get("client",""), "C3": pi.get("project_no",""),
        "C4": pi.get("project_name",""),
        "E4": datetime.today().strftime("%m/%d/%Y"),
        "C5": pi.get("location",""), "F5": pi.get("author",""),
        "B7": pi.get("equipment",""),
    }.items():
        if val:
            c = resolve_cell(ws, addr)
            if c is not None:
                c.value = val

    # Write vendor names into row 7 (C7, D7, E7, F7 …)
    for i in range(1, vendor_count + 1):
        col_letter = get_column_letter(2 + i)
        addr = f"{col_letter}7"
        name = pi.get(f"vendor{i}_name", f"Vendor {i}")
        if name:
            c = resolve_cell(ws, addr)
            if c is not None:
                c.value = name

    # Data rows
    for row, field in ROW_MAP.items():
        for src, col in col_map.items():
            val = coerce((data.get(src) or {}).get(field), field)
            addr = f"{col}{row}"
            cell = resolve_cell(ws, addr)
            if cell is None:
                continue  # horizontally-merged cell — do not corrupt another column
            if val is not None:
                cell.value = val
            elif col in vendor_cols:
                curr = cell.value
                if curr is None or str(curr).strip() in ("","-","By Vendor","N/A"):
                    cell.fill = YELLOW

    # TOTAL formulas for all vendor columns
    for col in vendor_cols:
        c = resolve_cell(ws, f"{col}105")
        if c is not None:
            c.value = f"=SUM({col}100:{col}101)*{col}103"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


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
  --accent:#e8286a;--accent2:#fdedf4;
  --accent-warm:#f5a623;
  --green:#0d6e3b;--gbg:#e6f4ed;
  --red:#9b1c1c;--rbg:#fde8e8;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:var(--sans);background:var(--bg);color:var(--ink);font-size:14px;line-height:1.5;}
.topbar{background:var(--ink);color:#fff;height:58px;padding:0 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;border-bottom:1px solid rgba(255,255,255,.06);}
.brand{display:flex;align-items:center;gap:10px;}
.brand img{height:34px;width:auto;display:block;mix-blend-mode:screen;}
.brand-name{font-family:var(--sans);font-size:15px;font-weight:500;color:#fff;letter-spacing:.01em;}
.brand-name span{color:rgba(255,255,255,.38);font-weight:300;font-size:12px;margin-left:6px;letter-spacing:.04em;}
.badge{font-family:var(--mono);font-size:11px;padding:4px 12px;border-radius:20px;background:rgba(255,255,255,.08);color:rgba(255,255,255,.5);}
.page{max-width:780px;margin:0 auto;padding:36px 20px 80px;}
.card{background:var(--bg2);border:1px solid var(--line);border-radius:12px;margin-bottom:18px;overflow:hidden;}
.ch{padding:15px 22px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:12px;}
.ch-ico{width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,var(--accent-warm),var(--accent));color:#fff;font-family:var(--mono);font-size:11px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-weight:500;}
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
.vgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:12px;}
.vcard{border:1px solid var(--line2);border-radius:10px;overflow:hidden;}
.vcard.loaded{border-color:#f5a19d;}
.vhead{background:var(--bg);padding:8px 12px;border-bottom:1px solid var(--line);font-family:var(--mono);font-size:10px;color:var(--ink3);text-transform:uppercase;letter-spacing:.05em;display:flex;align-items:center;justify-content:space-between;}
.vhead.loaded{background:#fff5f8;color:var(--accent);border-bottom-color:#f5a19d;}
.vremove{background:none;border:none;color:var(--ink3);cursor:pointer;font-size:15px;line-height:1;padding:0 2px;border-radius:3px;font-family:var(--sans);}
.vremove:hover{color:var(--red);background:var(--rbg);}
.vbody{padding:11px 12px;background:var(--bg2);}
.vbody input{width:100%;border:1px solid var(--line2);border-radius:6px;padding:7px 10px;font-size:13px;font-family:var(--sans);color:var(--ink);outline:none;margin-bottom:8px;}
.vbody input:focus{border-color:var(--accent);}
.vup{width:100%;padding:7px;border:1px dashed var(--line2);border-radius:6px;background:var(--bg);font-size:11px;color:var(--ink3);cursor:pointer;text-align:center;font-family:var(--sans);transition:all .15s;}
.vup:hover{border-color:var(--accent);color:var(--accent);background:var(--accent2);}
.vfname{font-family:var(--mono);font-size:10px;color:var(--green);margin-top:5px;min-height:13px;}
.add-vbtn{display:block;width:100%;margin-top:10px;padding:8px;background:var(--bg2);border:1px dashed var(--line2);border-radius:8px;color:var(--ink3);font-size:12px;cursor:pointer;font-family:var(--sans);transition:all .15s;}
.add-vbtn:hover{border-color:var(--accent);color:var(--accent);background:var(--accent2);}
.sub-area{margin-top:24px;text-align:center;}
.sub-btn{background:linear-gradient(135deg,var(--accent-warm),var(--accent));color:#fff;border:none;padding:13px 42px;border-radius:8px;font-size:15px;font-weight:500;font-family:var(--sans);cursor:pointer;transition:opacity .15s;}
.sub-btn:hover{opacity:.88;}
.sub-btn:disabled{background:linear-gradient(135deg,#ccc,#aaa);cursor:not-allowed;opacity:1;}
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
  <div class="brand">
    <span class="brand-name">Bid Tab Agent<span>AI</span></span>
  </div>
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
        <div class="fld"><label>Project Location</label><input id="pi-loc" placeholder="e.g. Mobile, AL"></div>
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
    <div class="ch"><div class="ch-ico">2</div><div><h2>Data Sheet</h2><p>Engineering spec — fills col B (standards column)</p></div></div>
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
      <div class="vgrid" id="vgrid"></div>
      <button type="button" class="add-vbtn" onclick="addVendor()">+ Add Vendor</button>
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
    <h3>AI is working…</h3>
    <p id="pmsg">Reading your files</p>
    <div class="pstep" id="pstep"></div>
  </div>
</div>

<script>
/* ── Drag-drop / chip helpers ── */
function zo(e,id){e.preventDefault();document.getElementById(id).classList.add('drag');}
function zl(id){document.getElementById(id).classList.remove('drag');}
function zd(e,zid,fid){e.preventDefault();zl(zid);const f=e.dataTransfer.files[0];if(!f)return;const inp=document.getElementById(fid);const dt=new DataTransfer();dt.items.add(f);inp.files=dt.files;chip(inp,fid==='f-ds'?'chip-ds':fid==='f-tmpl'?'chip-tmpl':null);}
function chip(inp,chipId){const f=inp.files[0];if(!f||!chipId)return;document.getElementById(chipId).innerHTML=`<div class="fchip">✓ ${f.name}</div>`;}

/* ── Dynamic vendor slots ── */
let vendorSlots=[];   // ordered array of slot indices currently shown
let nextVIdx=0;       // ever-increasing unique id for new slots

function makeVendorCard(idx,displayNum){
  const div=document.createElement('div');
  div.className='vcard';div.id='vc'+idx;
  div.innerHTML=`
    <div class="vhead" id="vh${idx}">
      <span id="vhl${idx}">Vendor ${displayNum}</span>
      <button type="button" class="vremove" id="vrm${idx}" onclick="removeVendor(${idx})" title="Remove vendor">×</button>
    </div>
    <div class="vbody">
      <input id="vn${idx}" placeholder="Vendor name" value="Vendor ${displayNum}"
             oninput="document.getElementById('vhl${idx}').textContent=this.value||'Vendor ${displayNum}'">
      <div class="vup" onclick="document.getElementById('vf${idx}').click()">+ Attach quote</div>
      <input type="file" id="vf${idx}" style="display:none" accept=".xlsx,.xls,.csv,.txt,.pdf"
             onchange="vfile(this,${idx})">
      <div class="vfname" id="vfn${idx}"></div>
    </div>`;
  return div;
}

function addVendor(){
  const idx=nextVIdx++;
  vendorSlots.push(idx);
  document.getElementById('vgrid').appendChild(makeVendorCard(idx,vendorSlots.length));
  syncRemoveButtons();
}

function removeVendor(idx){
  if(vendorSlots.length<=1)return;
  vendorSlots=vendorSlots.filter(i=>i!==idx);
  const card=document.getElementById('vc'+idx);
  if(card)card.remove();
  syncRemoveButtons();
}

function syncRemoveButtons(){
  const show=vendorSlots.length>1;
  vendorSlots.forEach(i=>{
    const btn=document.getElementById('vrm'+i);
    if(btn)btn.style.visibility=show?'visible':'hidden';
  });
}

function vfile(inp,idx){
  const f=inp.files[0];if(!f)return;
  document.getElementById('vfn'+idx).textContent='✓ '+f.name;
  document.getElementById('vc'+idx).classList.add('loaded');
  document.getElementById('vh'+idx).classList.add('loaded');
}

/* Start with 3 vendor slots */
(function initVendors(){for(let i=0;i<3;i++)addVendor();}());

/* ── Main generate ── */
async function run(){
  const tmplFile=document.getElementById('f-tmpl').files[0];
  if(!tmplFile){alert('Please upload your bid tab template (.xlsx).');return;}
  const dsFile=document.getElementById('f-ds').files[0];
  const hasVendor=vendorSlots.some(i=>document.getElementById('vf'+i)?.files[0]);
  if(!dsFile&&!hasVendor){alert('Please upload at least a data sheet or one vendor quote.');return;}

  document.getElementById('sub-btn').disabled=true;
  document.getElementById('rbox').classList.remove('show');
  document.getElementById('ebox').classList.remove('show');
  document.getElementById('overlay').classList.add('show');
  document.getElementById('badge').textContent='Processing…';

  const msgs=['Reading files…','AI extracting specs…','Mapping vendor data…','Writing Excel…'];
  let mi=0;
  const ticker=setInterval(()=>{
    document.getElementById('pmsg').textContent=msgs[Math.min(mi,msgs.length-1)];
    document.getElementById('pstep').textContent=`Step ${Math.min(mi+1,4)} of 4`;
    mi++;
  },5000);

  try{
    const fd=new FormData();
    fd.append('template',tmplFile);
    if(dsFile)fd.append('datasheet',dsFile);

    /* Submit all vendor slots in display order */
    vendorSlots.forEach((slotIdx,pos)=>{
      const name=document.getElementById('vn'+slotIdx)?.value||`Vendor ${pos+1}`;
      const file=document.getElementById('vf'+slotIdx)?.files[0];
      fd.append(`vendor_name_${pos+1}`,name);
      if(file)fd.append(`vendor_${pos+1}`,file);
    });
    fd.append('vendor_count',vendorSlots.length);

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
    }else{
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
        return jsonify({"error": "ANTHROPIC_API_KEY not set on server. Add it in Railway environment variables."})

    template_file = request.files.get("template")
    if not template_file or not template_file.filename:
        return jsonify({"error": "No template uploaded."})

    template_bytes = template_file.read()

    # Detect how many vendor slots were submitted
    vendor_count = int(request.form.get("vendor_count", 3))
    vendor_count = max(1, min(vendor_count, 10))  # clamp 1-10

    pi = {
        "client":       request.form.get("client", ""),
        "project_no":   request.form.get("project_no", ""),
        "project_name": request.form.get("project_name", ""),
        "location":     request.form.get("location", ""),
        "equipment":    request.form.get("equipment", ""),
        "author":       request.form.get("author", ""),
    }
    for i in range(1, vendor_count + 1):
        pi[f"vendor{i}_name"] = request.form.get(f"vendor_name_{i}", f"Vendor {i}")

    # Build content for Claude
    content = ""
    ds = request.files.get("datasheet")
    if ds and ds.filename:
        content += f"=== DATA SHEET ({ds.filename}) ===\n{extract_text(ds)}\n\n"

    vendors_present = []
    for i in range(1, vendor_count + 1):
        vf = request.files.get(f"vendor_{i}")
        vn = pi.get(f"vendor{i}_name", f"Vendor {i}")
        if vf and vf.filename:
            content += f"=== VENDOR {i} — {vn} ({vf.filename}) ===\n{extract_text(vf)}\n\n"
            vendors_present.append(i)

    if not content.strip():
        return jsonify({"error": "No data sheet or vendor quotes uploaded."})

    # Call Claude
    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=8096,
            system=SYSTEM,
            messages=[{"role": "user", "content": content}]
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```json\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            data = json.loads(_repair_json(raw))
        # Log extraction summary for debugging
        print("\n=== AI EXTRACTION SUMMARY ===")
        all_srcs = ["datasheet"] + [f"vendor{i}" for i in range(1, vendor_count + 1)]
        for src in all_srcs:
            src_data = data.get(src) or {}
            found = {k: v for k, v in src_data.items() if v is not None}
            print(f"  [{src}] {len(found)} fields filled: {list(found.keys())}")
        print("==============================\n")
    except json.JSONDecodeError as e:
        return jsonify({"error": f"AI returned malformed JSON: {e}"})
    except Exception as e:
        return jsonify({"error": f"AI error: {e}"})

    # Fill Excel
    try:
        xlsx_bytes = fill_excel(template_bytes, data, pi, vendor_count=vendor_count)
    except Exception as e:
        return jsonify({"error": f"Excel error: {e}"})

    # Save
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    proj = (pi.get("project_no") or "BidTab").replace(" ", "_")
    fname = f"Bid_Tab_{proj}_{stamp}.xlsx"
    fpath = os.path.join(OUTPUT_DIR, fname)
    with open(fpath, "wb") as f:
        f.write(xlsx_bytes)

    file_count = sum(1 for f in [ds] + [request.files.get(f"vendor_{i}") for i in range(1, vendor_count + 1)] if f and f.filename)
    vendor_names = [pi.get(f"vendor{i}_name", f"Vendor {i}") for i in vendors_present]
    msg = f"Filled from {file_count} file(s) across {len(vendors_present)} vendor(s). Yellow cells = fields not found in quote."

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
