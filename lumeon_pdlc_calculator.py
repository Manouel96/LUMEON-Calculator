"""LUMEON PDLC calculator. Run: py -m streamlit run lumeon_pdlc_calculator.py"""
import json, math, os, copy, platform, subprocess
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
import streamlit as st
import docx
from io import BytesIO
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
from google.oauth2 import service_account

# Google Drive IDs
GDRIVE_PROJECTS_FILE_ID = "1PF_NU3GHoVdJpz0iAtLxTjNb9_G0y05D"
GDRIVE_QUOTES_FOLDER_ID = "1gNfIX4RjSXuKB73r2j8o5VOHbeSRnRwS"

def get_drive_service():
    """Authenticates using Streamlit secrets."""
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        print(f"Drive connection error: {e}")
        return None

import base64
import requests

def upload_file_to_drive(local_path, project_name="Manual entry"):
    """Uploads quotes to Google Drive organized into project subfolders."""
    webhook_url = st.secrets.get("gdrive_webhook_url") or st.secrets.get("gcp_service_account", {}).get("gdrive_webhook_url")
    if not webhook_url:
        return None
    try:
        with open(local_path, "rb") as f:
            b64_content = base64.b64encode(f.read()).decode("utf-8")

        mime_type = "application/pdf" if local_path.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        payload = {
            "fileName": os.path.basename(local_path),
            "base64": b64_content,
            "mimeType": mime_type,
            "projectName": project_name
        }
        response = requests.post(webhook_url, json=payload, timeout=60)
        res_data = response.json()
        if res_data.get("status") == "success":
            return res_data.get("url")
        return None
    except Exception as e:
        print(f"Drive upload failed: {e}")
        return None

def delete_quote_from_drive(project_name, file_prefix):
    """Deletes matching quote files directly from Google Drive."""
    webhook_url = st.secrets.get("gdrive_webhook_url") or st.secrets.get("gcp_service_account", {}).get("gdrive_webhook_url")
    if not webhook_url:
        return
    try:
        payload = {
            "action": "delete_quote",
            "projectName": project_name,
            "prefix": file_prefix
        }
        requests.post(webhook_url, json=payload, timeout=30)
    except Exception as e:
        print(f"Drive deletion failed: {e}")

def get_live_drive_files(project_name):
    """Gets the list of files currently existing in Google Drive for this project."""
    webhook_url = st.secrets.get("gdrive_webhook_url") or st.secrets.get("gcp_service_account", {}).get("gdrive_webhook_url")
    if not webhook_url:
        return None
    try:
        payload = {
            "action": "list_quotes",
            "projectName": project_name
        }
        response = requests.post(webhook_url, json=payload, timeout=30)
        res_data = response.json()
        if res_data.get("status") == "success":
            return res_data.get("files", [])
        return None
    except Exception as e:
        print(f"Drive file list fetch failed: {e}")
        return None

HISTORY = Path(__file__).with_name('lumeon_calculator_history.json')
PROJECTS_FILE = Path(__file__).with_name('lumeon_projects.json')
DEFAULTS_FILE = Path(__file__).with_name('lumeon_defaults.json')

BASE_DEFAULTS = {
    'film_price_88': 31.0,
    'film_price_92': 38.0,
    'film_price_96': 70.0,
    'film': '92%', 'film_cost': 38., 'film_cost_cur': 'USD',
    'manual_sell': 52925., 'manual_sell_cur': 'AMD',
    'enforce_500mm': True, 'target_margin': 25.,
    'waste': 0., 'waste_pkg_mode': 'Auto-calculate',
    'method': 'Sheet box', 'use_volumetric_88_92': True,
    'ship_kg': 4000., 'ship_kg_cur': 'AMD',
    'fixed_ship': 0., 'fixed_ship_cur': 'AMD',
    'clearance': 35000., 'clearance_cur': 'AMD',
    'import_tax': 32., 'alibaba_fee_pct': 3.0,
    'input_tax_pct': 10., 'output_vat': 0.,
    'labor_mode': 'Per sqm', 'labor_rate_sqm': 10000., 'labor_cur': 'AMD',
    'labor_rate_hour': 10000., 'labor_hour_cur': 'AMD',
    'other': 0., 'other_cur': 'AMD',
    'enable_turnover_tax': True,
    'include_labor_in_turnover_tax': True,
    'turnover_tax_base': 10.0, 'turnover_tax_deduct': 6.0, 
    'turnover_tax_floor': 4.5, 'doc_expense_pct': 100.0,
    'profit_tax': 0., 'dividend_tax': 0.,
    'rate': 365., 'manual_rate': 365.,
    'exchange_mode': 'Automatic CBA', 'rate_date': '', 'rate_source': 'Manual/default',
    'remote_cost': 10., 'remote_cur': 'USD',
    'silicone_cost': 8., 'silicone_cur': 'USD',
    'ps_cost_30W': 22.0, 'ps_cur_30W': 'USD',
    'ps_cost_50W': 24.0, 'ps_cur_50W': 'USD',
    'ps_cost_100W': 32.0, 'ps_cur_100W': 'USD',
    'ps_cost_200W': 37.0, 'ps_cur_200W': 'USD',
    'ps_cost_300W': 45.0, 'ps_cur_300W': 'USD',
    'ps_cost_500W': 60.0, 'ps_cur_500W': 'USD'
}

def load_defaults():
    try:
        if DEFAULTS_FILE.exists():
            saved_defs = json.loads(DEFAULTS_FILE.read_text('utf-8'))
            return {**BASE_DEFAULTS, **saved_defs}
    except: pass
    return copy.deepcopy(BASE_DEFAULTS)

def save_defaults(d):
    DEFAULTS_FILE.write_text(json.dumps(d, indent=2), 'utf-8')

DEFAULT = load_defaults()

FILMS = {
    '88%': (DEFAULT['film_price_88'], 120), '88% Coated': (DEFAULT['film_price_88'], 500),
    '92%': (DEFAULT['film_price_92'], 145), '92% Coated': (DEFAULT['film_price_92'], 500), '92% Non-Adhesive': (DEFAULT['film_price_92'], 145),
    '96%': (DEFAULT['film_price_96'], 175), '96% Coated': (DEFAULT['film_price_96'], 500), '96% Non-Adhesive': (DEFAULT['film_price_96'], 175)
}

SHEET = [(0, 930, '1110 x 360 x 360', 6.6, 999, 33), (940, 1130, '1310 x 360 x 360', 7.4, 999, 33), (1140, 1330, '1510 x 360 x 360', 8.34, 999, 33), (1340, 1530, '1710 x 360 x 360', 9.16, 999, 33), (1540, 1730, '1910 x 360 x 360', 9.96, 999, 30), (1740, None, '2110 x 360 x 360', 12, 999, 30)]
TUBE = [(0, 950, '1090 x 160 x 160', 1.5, 2, None), (960, 1150, '1290 x 160 x 160', 1.64, 2, None), (1160, 1350, '1490 x 160 x 160', 1.92, 2, None), (1360, 1550, '1690 x 160 x 160', 2.18, 2, None), (1560, 1750, '1890 x 160 x 160', 2.4, 1, None), (1760, None, '2090 x 160 x 160', 3, 1, None)]
PS = {'30W': DEFAULT['ps_cost_30W'], '50W': DEFAULT['ps_cost_50W'], '100W': DEFAULT['ps_cost_100W'], '200W': DEFAULT['ps_cost_200W'], '300W': DEFAULT['ps_cost_300W'], '500W': DEFAULT['ps_cost_500W']}
SMALL_96 = [(90, 27, 28, 7.0), (100, 27, 28, 7.3), (118, 27, 28, 7.7), (126, 27, 28, 8.0), (148, 27, 28, 8.5)]
LARGE_96 = [(81, 34, 34, 7.0), (118, 34, 34, 8.5), (126, 34, 34, 8.8), (136, 34, 34, 9.2), (146, 34, 34, 9.7), (161, 34, 34, 10.7), (172, 34, 34, 10.85), (181, 34, 34, 11.0), (207, 34, 34, 11.5)]
VOLUMETRIC_DIVISOR = 5000.0

DEFAULT_PANELS = [{'Quantity': 1, 'Width (mm)': 700.0, 'Length (mm)': 2600.0}]
DEFAULT_PKGS = [{'Quantity': 1, 'Length (cm)': 0.0, 'Width (cm)': 0.0, 'Height (cm)': 0.0, 'Gross weight (kg)': 0.0}]

def safe_qty(val):
    try:
        v = float(val)
        return 0 if math.isnan(v) else int(v)
    except: return 0

def safe_float(val):
    try:
        v = float(val)
        return 0.0 if math.isnan(v) else v
    except: return 0.0

def to_amd(val, cur, rate): return val * rate if cur == 'USD' else val
def to_usd(val, cur, rate): return val if cur == 'USD' else val / rate

def sanitize_config(cfg):
    clean = copy.deepcopy(DEFAULT)
    if isinstance(cfg, dict):
        for k in DEFAULT.keys():
            if k in cfg: clean[k] = copy.deepcopy(cfg[k])
        if 'manual_sell_amd' in cfg and 'manual_sell' not in cfg:
            clean['manual_sell'] = float(cfg['manual_sell_amd'])
            clean['manual_sell_cur'] = 'AMD'
    return clean

def upgrade_panels(panels):
    out = []
    for p in (panels or []):
        new_p = dict(p)
        if 'Width (m)' in new_p: new_p['Width (mm)'] = new_p.pop('Width (m)') * 1000.0
        if 'Length (m)' in new_p: new_p['Length (mm)'] = new_p.pop('Length (m)') * 1000.0
        if 'Quantity' not in new_p: new_p['Quantity'] = 0
        if 'Width (mm)' not in new_p: new_p['Width (mm)'] = 0.0
        if 'Length (mm)' not in new_p: new_p['Length (mm)'] = 0.0
        out.append(new_p)
    return out

def calc_area(panels):
    area = 0.0
    for p in panels:
        q = safe_qty(p.get('Quantity', 0))
        w = safe_float(p.get('Width (mm)', 0)) / 1000.0
        l = safe_float(p.get('Length (mm)', 0)) / 1000.0
        area += q * w * l
    return area

def box_mm_volumetric(size_str):
    try:
        dims = [float(x.strip()) for x in size_str.split('x')]
        return (dims[0] / 10.0) * (dims[1] / 10.0) * (dims[2] / 10.0) / VOLUMETRIC_DIVISOR
    except: return 0.0

def fetch_cba_usd_rate():
    url = 'http://api.cba.am/exchangerates.asmx'
    envelope = '''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
 xmlns:xsd="http://www.w3.org/2001/XMLSchema"
 xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body><ExchangeRatesLatestByISO xmlns="http://www.cba.am/">
  <iso>USD</iso>
 </ExchangeRatesLatestByISO></soap:Body>
</soap:Envelope>'''.encode('utf-8')
    request = urllib.request.Request(url, data=envelope, headers={'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '"http://www.cba.am/ExchangeRatesLatestByISO"', 'User-Agent': 'LUMEON-PDLC-Calculator/1.0'})
    with urllib.request.urlopen(request, timeout=12) as response:
        root = ET.fromstring(response.read())
    def find_text(*names):
        wanted = {x.lower() for x in names}
        for element in root.iter():
            local = element.tag.split('}')[-1].lower()
            if local in wanted and element.text and element.text.strip(): return element.text.strip()
        return ''
    rate_text = find_text('Rate', 'ExchangeRate')
    amount_text = find_text('Amount') or '1'
    date_text = find_text('CurrentDate', 'Date')
    if not rate_text: raise ValueError('The CBA response did not contain a USD exchange rate.')
    rate = float(rate_text.replace(',', '.')) / float(amount_text.replace(',', '.'))
    return rate, date_text or datetime.now().strftime('%Y-%m-%d')

def load_projects():
    """Loads projects from Google Drive, falling back to local file."""
    service = get_drive_service()
    if service:
        try:
            request = service.files().get_media(fileId=GDRIVE_PROJECTS_FILE_ID)
            content = request.execute().decode('utf-8')
            data = json.loads(content)
            if isinstance(data, dict):
                PROJECTS_FILE.write_text(content, encoding='utf-8')
                return data
        except Exception as e:
            print(f"Drive load failed, falling back to local: {e}")

    try:
        data = json.loads(PROJECTS_FILE.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_projects(projects):
    """Saves projects to local file AND syncs directly to Google Drive."""
    json_str = json.dumps(projects, ensure_ascii=False, indent=2)
    PROJECTS_FILE.write_text(json_str, encoding='utf-8')

    service = get_drive_service()
    if service:
        try:
            media = MediaIoBaseUpload(BytesIO(json_str.encode('utf-8')), mimetype='application/json', resumable=True)
            service.files().update(fileId=GDRIVE_PROJECTS_FILE_ID, media_body=media).execute()
        except Exception as e:
            print(f"Drive save failed: {e}")

def hist():
    try: return json.loads(HISTORY.read_text('utf-8'))
    except: return []

def save_hist(x): HISTORY.write_text(json.dumps(x, indent=2), 'utf-8')

def pkg_spec(width, method):
    table = SHEET if method == 'Sheet box' else TUBE
    for r in table:
        if r[0] <= width <= (r[1] if r[1] is not None else 10**9): return r, False
    for r in table:
        if width < r[0]: return r, True
    return table[-1], False

def next_carton(cartons, required_length_cm):
    for carton in cartons:
        if carton[0] >= required_length_cm: return carton, False
    return cartons[-1], True

def packaging_manual(manual_packages, label_prefix=""):
    rows = manual_packages or []
    total_packages = 0; total_gross = 0.0; total_volumetric = 0.0; descriptions = []; package_list = []
    for row in rows:
        qty = max(safe_qty(row.get('Quantity', 0)), 0)
        length = max(safe_float(row.get('Length (cm)', 0)), 0.0)
        width = max(safe_float(row.get('Width (cm)', 0)), 0.0)
        height = max(safe_float(row.get('Height (cm)', 0)), 0.0)
        gross_each = max(safe_float(row.get('Gross weight (kg)', 0)), 0.0)
        if qty <= 0 or length <= 0 or width <= 0 or height <= 0: continue
        volumetric_each = length * width * height / VOLUMETRIC_DIVISOR
        total_packages += qty
        total_gross += qty * gross_each
        total_volumetric += qty * volumetric_each
        prefix_str = f"{label_prefix}: " if label_prefix else ""
        descriptions.append(f'{prefix_str}{qty} x {length:g}x{width:g}x{height:g} cm (gross {gross_each:g} kg, vol {volumetric_each:.2f} kg)')
        package_list.append({'Quantity': qty, 'Length (cm)': length, 'Width (cm)': width, 'Height (cm)': height, 'Gross weight (kg)': gross_each})
    return {'packages': total_packages, 'tare': 0.0, 'filmkg': 0.0, 'accessories': 0.0, 'gross': total_gross, 'volumetric': total_volumetric, 'chargeable': max(total_gross, total_volumetric), 'description': '; '.join(descriptions) or (f'No valid {label_prefix.lower()} package rows entered.' if label_prefix else 'No valid manual package rows entered.'), 'warning': '' if total_packages else 'Enter at least one complete manual package row.', 'package_list': package_list}

def calc_waste_package(waste_area, max_width_m, film_type, method):
    width_m = max_width_m if max_width_m > 0 else 1.5
    width_mm = width_m * 1000
    length_m = waste_area / width_m
    filmkg = waste_area * 0.78
    if film_type.startswith('96%'):
        width_cm = width_m * 100
        use_small = (length_m <= 1.5 and width_cm <= 120)
        cartons = SMALL_96 if use_small else LARGE_96
        required = max(width_cm + 20, 90 if use_small else width_cm + 18)
        carton, oversize = next_carton(cartons, required)
        L, W, H, tare_each = carton
        vol = (L * W * H) / VOLUMETRIC_DIVISOR
        gross = tare_each + filmkg
        desc = f"1 x Separated Waste Box ({L}x{W}x{H} cm; gross {gross:.2f} kg; vol {vol:.2f} kg) for {waste_area:.2f} sqm"
        return {'packages': 1, 'tare': tare_each, 'filmkg': filmkg, 'gross': gross, 'volumetric': vol, 'description': desc, 'package_list': [{'Quantity': 1, 'Length (cm)': L, 'Width (cm)': W, 'Height (cm)': H, 'Gross weight (kg)': gross}]}
    else:
        calc_method = method if method != 'Manual packages' else 'Sheet box'
        spec, gap = pkg_spec(width_mm, calc_method)
        lo, hi, size, kg, cap, maxlen = spec
        vol = box_mm_volumetric(size)
        gross = kg + filmkg
        desc = f"1 x Separated Waste Pkg ({size} mm; gross {gross:.2f} kg; vol {vol:.2f} kg) for {waste_area:.2f} sqm"
        dims = [float(x.strip())/10.0 for x in size.split('x')] if 'x' in size else [0.0, 0.0, 0.0]
        return {'packages': 1, 'tare': kg, 'filmkg': filmkg, 'gross': gross, 'volumetric': vol, 'description': desc, 'package_list': [{'Quantity': 1, 'Length (cm)': dims[0] if len(dims)>0 else 0, 'Width (cm)': dims[1] if len(dims)>1 else 0, 'Height (cm)': dims[2] if len(dims)>2 else 0, 'Gross weight (kg)': gross}]}

def packaging_88_92(df, c, ps_list, silicone_qty):
    active = df[(df.Quantity > 0) & (df['Supplier Width (m)'] > 0) & (df['Length (m)'] > 0)].copy()
    if active.empty:
        return {'packages': 0, 'tare': 0., 'filmkg': 0., 'accessories': 0., 'gross': 0., 'volumetric': 0., 'chargeable': 0., 'description': '', 'warning': '', 'package_list': []}
    
    active['spec_tuple'] = active['Supplier Width (m)'].apply(lambda w: pkg_spec(w * 1000, c['method'])[0])
    active['gap'] = active['Supplier Width (m)'].apply(lambda w: pkg_spec(w * 1000, c['method'])[1])
    
    total_n = 0; tare = 0.; vol = 0.; desc = []; gaps = []; package_list = []
    
    for spec_tuple, g in active.groupby('spec_tuple'):
        lo, hi, size, kg, cap, maxlen = spec_tuple
        pieces = int(g.Quantity.sum())
        length = float((g.Quantity * g['Length (m)']).sum())
        n = math.ceil(pieces / cap)
        if maxlen: n = max(n, math.ceil(length / maxlen))
        total_n += n; tare += n * kg; vol += n * box_mm_volumetric(size)
        desc.append(f'{n} x {size} mm package')
        
        grp_area = float((g.Quantity * g['Supplier Width (m)'] * g['Length (m)']).sum())
        grp_filmkg = grp_area * 0.78
        grp_gross = n * kg + grp_filmkg
        dims = [float(x.strip())/10.0 for x in size.split('x')] if 'x' in size else [0.0, 0.0, 0.0]
        package_list.append({'Quantity': n, 'Length (cm)': dims[0] if len(dims)>0 else 0, 'Width (cm)': dims[1] if len(dims)>1 else 0, 'Height (cm)': dims[2] if len(dims)>2 else 0, 'Gross weight (kg)': grp_gross / max(n, 1)})
        
        if g['gap'].any(): gaps.append(f'Some panels fell in a 10 mm supplier range gap; next larger package selected.')
            
    area = float((active.Quantity * active['Supplier Width (m)'] * active['Length (m)']).sum())
    filmkg = area * .78
    ps_w = sum(PS.get(ps.get('Type', '30W'), 1.16) * safe_qty(ps.get('Quantity', 0)) for ps in ps_list)
    accessories = ps_w + safe_float(silicone_qty) * 0.3
    
    gross = tare + filmkg + accessories
    
    if package_list and accessories > 0:
        package_list[0]['Gross weight (kg)'] += accessories / package_list[0]['Quantity']
        
    chargeable = max(gross, vol) if c.get('use_volumetric_88_92', True) else gross
    return {'packages': total_n, 'tare': tare, 'filmkg': filmkg, 'accessories': accessories, 'gross': gross, 'volumetric': vol, 'chargeable': max(gross, vol), 'description': '; '.join(desc), 'warning': ' '.join(set(gaps)), 'package_list': package_list}

def packaging_96(df, c, ps_list, silicone_qty):
    active = df[(df.Quantity > 0) & (df['Supplier Width (m)'] > 0) & (df['Length (m)'] > 0)].copy()
    if active.empty: return {'packages': 0, 'tare': 0., 'filmkg': 0., 'accessories': 0., 'gross': 0., 'volumetric': 0., 'chargeable': 0., 'description': '', 'warning': '', 'package_list': []}
    
    total_n = 0; tare = 0.; gross = 0.; vol = 0.; desc = []; warnings = []; package_list = []
    for _, r in active.iterrows():
        qty = int(r['Quantity']); width_cm = float(r['Supplier Width (m)']) * 100; height = float(r['Length (m)'])
        use_small = qty <= 2 and height <= 1.5 and width_cm <= 120
        cartons = SMALL_96 if use_small else LARGE_96
        required = max(width_cm + 20, 90 if use_small else width_cm + 18)
        carton, oversize = next_carton(cartons, required)
        L, W, H, tare_each = carton
        if use_small: capacity = 2
        elif height < 1.0: capacity = 20
        elif height < 2.0: capacity = 15
        else: capacity = 10
        n = math.ceil(qty / capacity)
        film_area = qty * float(r['Supplier Width (m)']) * height
        filmkg = film_area * .78
        row_tare = n * tare_each; row_gross = filmkg + row_tare; row_vol = n * (L * W * H / VOLUMETRIC_DIVISOR)
        total_n += n; tare += row_tare; gross += row_gross; vol += row_vol
        desc.append(f'{n} x {L}x{W}x{H} cm, capacity {capacity}, for {width_cm:.0f} cm width x {height:.2f} m height')
        package_list.append({'Quantity': n, 'Length (cm)': L, 'Width (cm)': W, 'Height (cm)': H, 'Gross weight (kg)': row_gross / max(n, 1)})
        if oversize: warnings.append(f'{width_cm:.0f} cm film exceeds the largest listed carton; {L} cm carton used provisionally.')
    ps_w = sum(PS.get(ps.get('Type', '30W'), 1.16) * safe_qty(ps.get('Quantity', 0)) for ps in ps_list)
    accessory = ps_w + safe_float(silicone_qty) * 0.3
    gross += accessory
    if package_list and accessory > 0:
        package_list[0]['Gross weight (kg)'] += accessory / package_list[0]['Quantity']
    chargeable = max(gross, vol)
    return {'packages': total_n, 'tare': tare, 'filmkg': float((active.Quantity * active['Supplier Width (m)'] * active['Length (m)']).sum()) * .78, 'accessories': accessory, 'gross': gross, 'volumetric': vol, 'chargeable': chargeable, 'description': '; '.join(desc), 'warning': ' '.join(warnings), 'package_list': package_list}

def packaging(df, c, ps_list, silicone_qty, manual_packages, waste_manual_packages):
    if c.get('method') == 'Manual packages': p = packaging_manual(manual_packages)
    elif c['film'].startswith('96%'): p = packaging_96(df, c, ps_list, silicone_qty)
    else: p = packaging_88_92(df, c, ps_list, silicone_qty)
    
    waste_pct = float(c.get('waste', 0.0))
    if waste_pct > 0:
        active = df[(df.Quantity > 0) & (df['Supplier Width (m)'] > 0) & (df['Length (m)'] > 0)]
        chargeable_area = float((active.Quantity * active['Supplier Width (m)'] * active['Length (m)']).sum())
        waste_area = chargeable_area * (waste_pct / 100.0)
        if waste_area > 0:
            waste_mode = c.get('waste_pkg_mode', 'Auto-calculate')
            if c.get('method') == 'Manual packages' and waste_mode == 'Enter manually':
                wp = packaging_manual(waste_manual_packages, label_prefix="Separated Waste Pkg")
            else:
                max_w = float(active['Supplier Width (m)'].max()) if not active.empty else 1.5
                if math.isnan(max_w) or max_w <= 0: max_w = 1.5
                wp = calc_waste_package(waste_area, max_w, c['film'], c.get('method', 'Sheet box'))
            
            p['packages'] += wp['packages']
            p['tare'] += wp['tare']
            p['filmkg'] += wp['filmkg']
            p['gross'] += wp['gross']
            p['volumetric'] += wp['volumetric']
            p['package_list'].extend(wp['package_list'])
            
            sep = " | " if p['description'] else ""
            p['description'] += f"{sep}+ {wp['description']}"
            if wp.get('warning'): p['warning'] = (p['warning'] + " " + wp['warning']).strip()
            
    p['chargeable'] = max(p['gross'], p['volumetric'])
    return p

def costs(df, c, ps_list, silicone_qty, remote_qty, manual_packages, waste_manual_packages, override_clearance=None):
    rate = c['rate']
    d = df.copy()
    for x in ['Quantity', 'Width (mm)', 'Length (mm)']: d[x] = pd.to_numeric(d.get(x, 0), errors='coerce').fillna(0).clip(lower=0)
    d['Width (m)'] = d['Width (mm)'] / 1000.0
    d['Length (m)'] = d['Length (mm)'] / 1000.0
    d['Supplier Width (m)'] = d['Width (m)'].apply(lambda x: max(x, 0.5) if x > 0 else 0)
    if c.get('enforce_500mm', True): d['Customer Width (m)'] = d['Supplier Width (m)']
    else: d['Customer Width (m)'] = d['Width (m)']
    d['Customer area (sqm)'] = d.Quantity * d['Customer Width (m)'] * d['Length (m)']
    d['Chargeable area (sqm)'] = d.Quantity * d['Supplier Width (m)'] * d['Length (m)']
    d['Panel class'] = d.apply(lambda r: 'Large' if r['Length (m)'] > 1.5 and r['Width (m)'] > 0.6 else 'Small', axis=1)
    d['Labor hours / panel'] = d['Panel class'].map({'Large': 2.0, 'Small': 1.25})
    d['Total labor hours'] = d.Quantity * d['Labor hours / panel']
    exact_area = float(d['Customer area (sqm)'].sum())
    chargeable_area = float(d['Chargeable area (sqm)'].sum())
    purchased = chargeable_area * (1 + c['waste'] / 100)
    total_labor_hours = float(d['Total labor hours'].sum())
    if c['labor_mode'] == 'Per sqm': labor_amd = to_amd(exact_area * float(c.get('labor_rate_sqm', 10000.)), c.get('labor_cur', 'AMD'), rate)
    else: labor_amd = to_amd(total_labor_hours * float(c.get('labor_rate_hour', 10000.)), c.get('labor_hour_cur', 'AMD'), rate)
    p = packaging(d, c, ps_list, silicone_qty, manual_packages, waste_manual_packages)
    filmusd = to_usd(purchased * float(c['film_cost']), c.get('film_cost_cur', 'USD'), rate)
    filmamd = to_amd(purchased * float(c['film_cost']), c.get('film_cost_cur', 'USD'), rate)
    total_ps_qty = sum(safe_qty(ps.get('Quantity', 0)) for ps in ps_list)
    remotes_per_ps = 1 + safe_qty(remote_qty)
    total_remotes = total_ps_qty * remotes_per_ps
    ps_cost_amd = sum(to_amd(float(c.get(f"ps_cost_{ps.get('Type', '30W')}", 0)), c.get(f"ps_cur_{ps.get('Type', '30W')}", 'AMD'), rate) * safe_qty(ps.get('Quantity', 0)) for ps in ps_list)
    silicone_cost_amd = safe_float(silicone_qty) * to_amd(float(c.get('silicone_cost', 0)), c.get('silicone_cur', 'AMD'), rate)
    remote_cost_amd = float(total_remotes) * to_amd(float(c.get('remote_cost', 0)), c.get('remote_cur', 'AMD'), rate)
    accessories_cost = ps_cost_amd + silicone_cost_amd + remote_cost_amd
    is_alibaba_film = c['film'].startswith('88%') or c['film'].startswith('92%')
    alibaba_fee_pct = float(c.get('alibaba_fee_pct', 3.0)) if is_alibaba_film else 0.0
    goods = filmamd + accessories_cost
    alibaba_fee_amd = goods * (alibaba_fee_pct / 100.0)
    ship_kg_amd = to_amd(float(c.get('ship_kg', 0)), c.get('ship_kg_cur', 'AMD'), rate)
    fixed_ship_amd = to_amd(float(c.get('fixed_ship', 0)), c.get('fixed_ship_cur', 'AMD'), rate)
    shipping = p['chargeable'] * ship_kg_amd + fixed_ship_amd
    duty = (goods + alibaba_fee_amd + shipping) * float(c.get('import_tax', 0)) / 100
    
    if override_clearance is not None: clearance_amd = override_clearance
    else: clearance_amd = to_amd(float(c.get('clearance', 0)), c.get('clearance_cur', 'AMD'), rate)
        
    other_amd = to_amd(float(c.get('other', 0)), c.get('other_cur', 'AMD'), rate)
    base_total = goods + alibaba_fee_amd + shipping + clearance_amd + duty + labor_amd + other_amd
    return d, {'area': exact_area, 'chargeable_area': chargeable_area, 'purchased': purchased, 'film_usd': filmusd, 'film_amd': filmamd, 'accessories_cost': accessories_cost, 'alibaba_fee': alibaba_fee_amd, 'shipping': shipping, 'clearance': clearance_amd, 'duty': duty, 'labor': labor_amd, 'labor_hours': total_labor_hours, 'base_total': base_total, **p}

def profit(rev, k, c):
    input_tax = rev * c.get('input_tax_pct', 10.) / 100
    output_vat = rev * c.get('output_vat', 0.) / 100
    total_cost = k['base_total'] + input_tax
    gross = rev - total_cost
    opt_status = 'none'
    opt_amount = 0.0
    
    if c.get('enable_turnover_tax', True):
        eligible_expense_base = k['base_total'] if c.get('include_labor_in_turnover_tax', True) else (k['base_total'] - k.get('labor', 0.0))
        doc_expenses = eligible_expense_base * (c.get('doc_expense_pct', 100.0) / 100.0)
        t_base_rate = c.get('turnover_tax_base', 10.0)
        t_deduct_rate = c.get('turnover_tax_deduct', 6.0)
        t_floor_rate = c.get('turnover_tax_floor', 4.5)
        
        base_tax = rev * (t_base_rate / 100.0)
        deduction = doc_expenses * (t_deduct_rate / 100.0)
        tentative_tax = base_tax - deduction
        min_tax = rev * (t_floor_rate / 100.0)
        tt = max(tentative_tax, min_tax)
        
        if t_deduct_rate > 0: optimal_expenses = rev * ((t_base_rate - t_floor_rate) / t_deduct_rate)
        else: optimal_expenses = 0.0
            
        if doc_expenses < optimal_expenses:
            opt_status = 'room'
            opt_amount = optimal_expenses - doc_expenses
        else:
            opt_status = 'wasted'
            opt_amount = doc_expenses - optimal_expenses
    else:
        tt = 0.0; tentative_tax = 0.0; min_tax = 0.0
        
    pre = gross - output_vat
    after = pre - tt
    dt = max(after, 0) * c.get('dividend_tax', 0.) / 100
    net = after - dt
    
    return {'revenue': rev, 'input_tax': input_tax, 'output_vat': output_vat, 'total_cost': total_cost, 'gross_profit': gross, 'net_profit': net, 'margin': net / rev * 100 if rev else 0, 'turnover_tax': tt, 'tentative_tax': tentative_tax, 'min_tax': min_tax, 'dividend_tax': dt, 'opt_status': opt_status, 'opt_amount': opt_amount}

def solve(k, c):
    target = c['target_margin'] / 100
    lo = max(k['base_total'], 1)
    hi = lo * 2
    def m(x): return profit(x, k, c)['margin'] / 100
    for _ in range(100):
        if m(hi) >= target: break
        hi *= 2
    for _ in range(100):
        mid = (lo + hi) / 2
        if m(mid) < target: lo = mid
        else: hi = mid
    return hi

def fill_bookmark(doc, bookmark_name, data):
    for paragraph in doc.paragraphs:
        if paragraph.runs:
            for child in paragraph._p:
                if child.tag.endswith('bookmarkStart'):
                    if child.get(f"{{{child.nsmap['w']}}}name") == bookmark_name:
                        run = paragraph.add_run(str(data))
                        child.addnext(run._r)
                        return
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.runs:
                        for child in paragraph._p:
                            if child.tag.endswith('bookmarkStart'):
                                if child.get(f"{{{child.nsmap['w']}}}name") == bookmark_name:
                                    run = paragraph.add_run(str(data))
                                    child.addnext(run._r)
                                    return
import cloudconvert

def convert_with_cloudconvert(docx_path, pdf_path):
    """Converts DOCX to PDF using genuine Microsoft Office engine on CloudConvert."""
    api_key = st.secrets.get("cloudconvert_api_key") or st.secrets.get("gcp_service_account", {}).get("cloudconvert_api_key")
    if not api_key:
        return False
    try:
        cloudconvert.configure(api_key=api_key)
        job = cloudconvert.Job.create(payload={
            "tasks": {
                "upload-docx": {
                    "operation": "import/upload"
                },
                "convert-to-pdf": {
                    "operation": "convert",
                    "input": "upload-docx",
                    "output_format": "pdf",
                    "engine": "office"
                },
                "export-pdf": {
                    "operation": "export/url",
                    "input": "convert-to-pdf"
                }
            }
        })
        upload_task_id = [t['id'] for t in job['tasks'] if t['name'] == 'upload-docx'][0]
        upload_task = cloudconvert.Task.find(id=upload_task_id)
        cloudconvert.Task.upload(file_name=docx_path, task=upload_task)

        # Wait for conversion to finish
        res_job = cloudconvert.Job.wait(id=job['id'])
        for task in res_job.get('tasks', []):
            if task.get('name') == 'export-pdf' and task.get('status') == 'finished':
                files = task.get('result', {}).get('files', [])
                if files:
                    cloudconvert.download(filename=pdf_path, url=files[0]['url'])
                    return os.path.exists(pdf_path)
        return False
    except Exception as e:
        print(f"CloudConvert conversion error: {e}")
        return False

def generate_quote(client_type, warranty_yes_no, measurement, client_name, client_company, client_address, client_phone, warranty_text, transformer_text, remote_text, manual_date, unit_price, total_amd, panels_df, transparency, project_name="Manual_Entry", enforce_500mm=True, compile_pdf=True, **kwargs):
    base_path = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(base_path, "Quote Templates")
    safe_proj = "".join(c for c in project_name if c.isalnum() or c in (' ', '_', '-')).strip()
    if not safe_proj:
        safe_proj = "Manual_Entry"
    output_dir = os.path.join(base_path, "Generated Quotes", safe_proj)
    os.makedirs(templates_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    template_name = ""
    if client_type == "Company":
        if warranty_yes_no == "Yes":
            template_name = "LUMEON_Quote_Template_Company_Warranty_Provided.docx" if measurement == "Provided" else "LUMEON_Quote_Template_Company_Warranty_Measured.docx"
        else:
            template_name = "LUMEON_Quote_Template_Company_NoWarranty_Provided.docx" if measurement == "Provided" else "LUMEON_Quote_Template_Company_NoWarranty_Measured.docx"
    elif client_type == "Personal":
        if warranty_yes_no == "Yes":
            template_name = "LUMEON_Quote_Template_Personal_Warranty_Provided.docx" if measurement == "Provided" else "LUMEON_Quote_Template_Personal_Warranty_Measured.docx"
        else:
            template_name = "LUMEON_Quote_Template_Personal_NoWarranty_Provided.docx" if measurement == "Provided" else "LUMEON_Quote_Template_Personal_NoWarranty_Measured.docx"
    else:
        return None, "Please select 'Personal' or 'Company' for Client Type."
        
    template_path = os.path.join(templates_dir, template_name)
    if not os.path.exists(template_path):
        return None, f"Template not found at: {template_path}"
        
    base_date = manual_date if manual_date else datetime.now()
    issue_date = base_date.strftime("%d/%m/%Y")
    valid_date = (base_date + timedelta(days=14)).strftime("%d/%m/%Y")
    quote_num = datetime.now().strftime("%y%m%d%H%M")
    
    doc = docx.Document(template_path)
    fill_bookmark(doc, "bmCustomerName", client_name)
    fill_bookmark(doc, "bmCompany", client_company)
    fill_bookmark(doc, "bmAddress", client_address)
    fill_bookmark(doc, "bmPhone", client_phone)
    fill_bookmark(doc, "bmWarranty", warranty_text)
    fill_bookmark(doc, "bmTransformer", transformer_text)
    fill_bookmark(doc, "bmRemote", remote_text)
    fill_bookmark(doc, "bmDate", issue_date)
    fill_bookmark(doc, "bmValidDate", valid_date)
    fill_bookmark(doc, "bmQuoteNum", quote_num)
    fill_bookmark(doc, "bmGrandTotal", f"{total_amd:,.0f}")
    
    # Target widths: 3420, 1350, 1170, 1620, 1440 dxa (Sum = 9000 dxa / 450 pt)
    TARGET_WIDTHS_DXA = [3420, 1350, 1170, 1620, 1440]
    TARGET_WIDTHS_EMU = [docx.shared.Emu(int(w * 914400 / 1440)) for w in TARGET_WIDTHS_DXA]
    
    def align_table(tbl):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        
        tblPr = tbl._tbl.tblPr
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), "9000")
        tblW.set(qn("w:type"), "dxa")
        
        tblInd = tblPr.find(qn("w:tblInd"))
        if tblInd is None:
            tblInd = OxmlElement("w:tblInd")
            tblPr.append(tblInd)
        tblInd.set(qn("w:w"), "108")
        tblInd.set(qn("w:type"), "dxa")
        
        tblLayout = tblPr.find(qn("w:tblLayout"))
        if tblLayout is None:
            tblLayout = OxmlElement("w:tblLayout")
            tblPr.append(tblLayout)
        tblLayout.set(qn("w:type"), "fixed")
        
        tblGrid = tbl._tbl.tblGrid
        if tblGrid is not None:
            tbl._tbl.remove(tblGrid)
        new_grid = OxmlElement("w:tblGrid")
        for w_dxa in TARGET_WIDTHS_DXA:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(w_dxa))
            new_grid.append(col)
        tblPr.addnext(new_grid)
        
        for row in tbl.rows:
            for idx, cell in enumerate(row.cells):
                w_dxa = TARGET_WIDTHS_DXA[idx]
                w_emu = TARGET_WIDTHS_EMU[idx]
                cell.width = w_emu
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(qn("w:w"), str(w_dxa))
                tcW.set(qn("w:type"), "dxa")

    if len(doc.tables) >= 2:
        tbl = doc.tables[1]
        active_panels = panels_df[(panels_df['Quantity'] > 0) & (panels_df['Width (mm)'] > 0) & (panels_df['Length (mm)'] > 0)]
        insert_idx = 1
        for i, (_, row) in enumerate(active_panels.iterrows()):
            qty = int(row['Quantity'])
            exact_width_m = float(row['Width (mm)']) / 1000.0
            length_m = float(row['Length (mm)']) / 1000.0
            print_width_m = max(exact_width_m, 0.5) if enforce_500mm and exact_width_m < 0.5 else exact_width_m
            line_total = qty * print_width_m * length_m * unit_price
            
            if i == 0:
                cells = tbl.rows[insert_idx].cells
            else:
                prev_row = tbl.rows[insert_idx - 1]
                new_tr = copy.deepcopy(prev_row._tr)
                prev_row._tr.addnext(new_tr)
                cells = tbl.rows[insert_idx].cells
                for cell in cells:
                    cell.text = ""
                    
            cells[0].text = f"LUMEON Smart Film (Self-Adhesive) {transparency} Color: White"
            cells[1].text = f"{print_width_m * 100:g} cm x {length_m * 100:g} cm"
            cells[2].text = str(qty)
            cells[3].text = f"{unit_price:,.0f}"
            cells[4].text = f"{line_total:,.0f}"
            insert_idx += 1
            
        align_table(doc.tables[1])
        if len(doc.tables) >= 3:
            align_table(doc.tables[2])
            
    safe_company = client_company.replace(" ", "_") if client_company else "NoCompany"
    safe_name = client_name.replace(" ", "_") if client_name else "NoName"
    file_prefix = f"LUMEON_Quote#{quote_num}_{safe_name}_{safe_company}_{transparency}"
    docx_path = os.path.join(output_dir, f"{file_prefix}.docx")
    pdf_path = os.path.join(output_dir, f"{file_prefix}.pdf")
    doc.save(docx_path)
    
    pdf_success = False
    if compile_pdf:
        cc_api_key = st.secrets.get("CLOUDCONVERT_API_KEY", "")
        if cc_api_key:
            try:
                import requests, time
                headers = {"Authorization": f"Bearer {cc_api_key}", "Content-Type": "application/json"}
                payload = {
                    "tasks": {
                        "import-file": {"operation": "import/upload"},
                        "convert-to-pdf": {
                            "operation": "convert",
                            "input": "import-file",
                            "output_format": "pdf",
                            "engine": "office"
                        },
                        "export-file": {
                            "operation": "export/url",
                            "input": "convert-to-pdf"
                        }
                    }
                }
                res = requests.post("https://api.cloudconvert.com/v2/jobs", headers=headers, json=payload, timeout=30)
                if res.status_code == 201:
                    job_data = res.json()["data"]
                    upload_task = next(t for t in job_data["tasks"] if t["name"] == "import-file")
                    upload_url = upload_task["result"]["form"]["url"]
                    upload_params = upload_task["result"]["form"]["parameters"]
                    
                    with open(docx_path, "rb") as f:
                        files = {"file": f}
                        requests.post(upload_url, data=upload_params, files=files, timeout=60)
                        
                    job_id = job_data["id"]
                    for _ in range(30):
                        time.sleep(1)
                        poll_res = requests.get(f"https://api.cloudconvert.com/v2/jobs/{job_id}", headers=headers, timeout=15)
                        if poll_res.status_code == 200:
                            status = poll_res.json()["data"]["status"]
                            if status == "finished":
                                export_task = next(t for t in poll_res.json()["data"]["tasks"] if t["name"] == "export-file")
                                file_url = export_task["result"]["files"][0]["url"]
                                pdf_bytes = requests.get(file_url, timeout=30).content
                                with open(pdf_path, "wb") as pf:
                                    pf.write(pdf_bytes)
                                pdf_success = True
                                break
                            elif status == "error":
                                break
            except Exception as e:
                print(f"CloudConvert Conversion Error: {e}")

        # Local conversion fallback
        if not pdf_success:
            try:
                import pythoncom
                from docx2pdf import convert
                pythoncom.CoInitialize()
                convert(docx_path, pdf_path)
                pythoncom.CoUninitialize()
                pdf_success = True
            except Exception as e:
                print(f"Local PDF Conversion Failed: {e}")
                
    # Auto-upload to Google Drive via Google Apps Script Webhook
    final_output_path = (pdf_path if (compile_pdf and pdf_success) else docx_path)
    gas_webhook_url = st.secrets.get("GOOGLE_DRIVE_WEBHOOK_URL", "")
    if gas_webhook_url and os.path.exists(final_output_path):
        try:
            import requests, base64
            with open(final_output_path, "rb") as f:
                b64_content = base64.b64encode(f.read()).decode("utf-8")
            payload = {
                "filename": os.path.basename(final_output_path),
                "folderId": "1gNfIX4RjSXuKB73r2j8o5VOHbeSRnRwS",
                "fileData": b64_content,
                "mimeType": "application/pdf" if (compile_pdf and pdf_success) else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            requests.post(gas_webhook_url, json=payload, timeout=20)
        except Exception as e:
            print(f"Drive Webhook upload error: {e}")
            
    return final_output_path, "Success"

st.set_page_config(page_title='LUMEON Calculator', layout='wide')

if 'q_client_type' not in st.session_state:
    st.session_state.q_client_type = 'Company'
    st.session_state.q_warranty_yes_no = 'Yes'
    st.session_state.q_measurement = 'Provided'
    st.session_state.q_client_name = ''
    st.session_state.q_client_company = ''
    st.session_state.q_client_address = ''
    st.session_state.q_client_phone = ''
    st.session_state.q_warranty_text = ''
    st.session_state.q_transformer_text = ''
    st.session_state.q_remote_text = ''
    
if 'c' not in st.session_state: st.session_state.c = sanitize_config(DEFAULT)
if 'panels' not in st.session_state: st.session_state.panels = copy.deepcopy(DEFAULT_PANELS)
else: st.session_state.panels = upgrade_panels(st.session_state.panels)
if 'manual_packages' not in st.session_state: st.session_state.manual_packages = copy.deepcopy(DEFAULT_PKGS)
if 'waste_manual_packages' not in st.session_state: st.session_state.waste_manual_packages = copy.deepcopy(DEFAULT_PKGS)
if 'power_supplies' not in st.session_state: st.session_state.power_supplies = [{'Type': '30W', 'Quantity': 0}]
if 'remote_qty' not in st.session_state: st.session_state.remote_qty = 0
if 'silicone_qty' not in st.session_state: st.session_state.silicone_qty = 0
if 'clearance_w' not in st.session_state: st.session_state.clearance_w = 2.0
if 'clearance_h' not in st.session_state: st.session_state.clearance_h = 2.0
if 'auto' not in st.session_state: st.session_state.auto = False
if 'panel_revision' not in st.session_state: st.session_state.panel_revision = 0
if 'config_revision' not in st.session_state: st.session_state.config_revision = 0

if st.session_state.pop('saved_notice', False): st.success('Configuration version saved.')
if st.session_state.pop('deleted_notice', False): st.success('Selected configuration version was deleted.')
if st.session_state.pop('inputs_applied', False): st.success('Global configuration changes applied.')
if st.session_state.pop('manual_notice', False): st.success('Manual selling price applied.')

c = st.session_state.c
if c.get('exchange_mode', 'Automatic CBA') == 'Automatic CBA' and not c.get('rate_date'):
    try:
        c['rate'], c['rate_date'] = fetch_cba_usd_rate()
        c['rate_source'] = 'Central Bank of Armenia'
    except Exception as exc: st.session_state.cba_error = str(exc)

page = st.sidebar.radio('Navigation', ['Calculator', 'Projects', 'Batch & Quarter Close', 'Settings'], key='app_page')
projects = load_projects()

# SETTINGS PAGE
if page == 'Settings':
    st.title('LUMEON Global Defaults Configuration')
    st.caption('Change the permanent default values that automatically load when the calculator starts or resets.')
    
    if st.session_state.get('settings_auth') != True:
        pwd = st.text_input('Enter Administrator Password', type='password')
        if st.button('Login'):
            if pwd == '199672':
                st.session_state.settings_auth = True
                st.rerun()
            else:
                st.error('Incorrect password')
        st.stop()
        
    if st.button('Logout'):
        st.session_state.settings_auth = False
        st.rerun()
        
    st.subheader('Permanent Base Values')
    with st.form('defaults_form'):
        d = copy.deepcopy(DEFAULT)
        
        st.markdown("#### Supplier Film Cost Defaults")
        col1, col2, col3 = st.columns(3)
        d['film_price_88'] = col1.number_input('88% Film ($/sqm)', 0.0, value=float(d.get('film_price_88', 31.0)))
        d['film_price_92'] = col2.number_input('92% Film ($/sqm)', 0.0, value=float(d.get('film_price_92', 38.0)))
        d['film_price_96'] = col3.number_input('96% Film ($/sqm)', 0.0, value=float(d.get('film_price_96', 70.0)))
        
        st.markdown("#### Margins & Shipping")
        col1, col2, col3 = st.columns(3)
        d['target_margin'] = col1.number_input('Target Net Margin (%)', 0.0, 99.0, float(d['target_margin']))
        d['ship_kg'] = col2.number_input('Shipping cost (/kg AMD)', 0.0, value=float(d['ship_kg']))
        d['clearance'] = col3.number_input('Clearance (AMD)', 0.0, value=float(d['clearance']))
        
        st.markdown("#### Taxes & Fees")
        col1, col2, col3 = st.columns(3)
        d['import_tax'] = col1.number_input('Import Tax (%)', 0.0, 100.0, float(d['import_tax']))
        d['alibaba_fee_pct'] = col2.number_input('Alibaba Fee (%)', 0.0, 100.0, float(d['alibaba_fee_pct']))
        d['input_tax_pct'] = col3.number_input('Input VAT / Tax (%)', 0.0, 100.0, float(d['input_tax_pct']))
        
        st.markdown("#### Turnover Tax (Services)")
        col1, col2 = st.columns(2)
        d['turnover_tax_base'] = col1.number_input('Base Tax Rate (%)', 0.0, 100.0, float(d.get('turnover_tax_base', 10.0)))
        d['turnover_tax_deduct'] = col2.number_input('Expense Deduction Rate (%)', 0.0, 100.0, float(d.get('turnover_tax_deduct', 6.0)))
        col3, col4 = st.columns(2)
        d['turnover_tax_floor'] = col3.number_input('Minimum Tax Floor (%)', 0.0, 100.0, float(d.get('turnover_tax_floor', 4.5)))
        d['doc_expense_pct'] = col4.number_input('Documented Expenses (% of Base Cost)', 0.0, 100.0, float(d.get('doc_expense_pct', 100.0)), help="Percentage of base cost officially documented for the deduction.")
        d['enable_turnover_tax'] = st.checkbox('Enable Turnover Tax by Default', value=bool(d.get('enable_turnover_tax', True)))
        d['include_labor_in_turnover_tax'] = st.checkbox('Include Labor Cost in Tax Deduction by Default', value=bool(d.get('include_labor_in_turnover_tax', True)))

        st.markdown("#### Other Taxes")
        col1, col2 = st.columns(2)
        d['output_vat'] = col1.number_input('Output VAT (%)', 0.0, 100.0, float(d['output_vat']))
        d['dividend_tax'] = col2.number_input('Dividend Tax (%)', 0.0, 100.0, float(d.get('dividend_tax', 0.0)))
        
        st.markdown("#### Labor & Exchange")
        col1, col2 = st.columns(2)
        d['labor_rate_sqm'] = col1.number_input('Labor rate (/sqm AMD)', 0.0, value=float(d['labor_rate_sqm']))
        d['manual_rate'] = col2.number_input('Manual Exchange Rate', 0.0, value=float(d['manual_rate']))
        
        st.markdown("#### Accessory Base Costs (USD)")
        col1, col2, col3, col4 = st.columns(4)
        d['remote_cost'] = col1.number_input('Remote ($)', 0.0, value=float(d['remote_cost']))
        d['silicone_cost'] = col2.number_input('Silicone ($)', 0.0, value=float(d['silicone_cost']))
        d['ps_cost_30W'] = col3.number_input('30W PS ($)', 0.0, value=float(d['ps_cost_30W']))
        d['ps_cost_50W'] = col4.number_input('50W PS ($)', 0.0, value=float(d['ps_cost_50W']))
        
        col1, col2, col3, col4 = st.columns(4)
        d['ps_cost_100W'] = col1.number_input('100W PS ($)', 0.0, value=float(d['ps_cost_100W']))
        d['ps_cost_200W'] = col2.number_input('200W PS ($)', 0.0, value=float(d['ps_cost_200W']))
        d['ps_cost_300W'] = col3.number_input('300W PS ($)', 0.0, value=float(d['ps_cost_300W']))
        d['ps_cost_500W'] = col4.number_input('500W PS ($)', 0.0, value=float(d['ps_cost_500W']))
        
        if st.form_submit_button('Save Permanent Defaults', type='primary'):
            save_defaults(d)
            st.session_state.c = sanitize_config(d)
            st.success('Defaults saved! They will be automatically applied to all new sessions and calculations.')
    st.stop()


# SIDEBAR (Active Configuration)
with st.sidebar:
    st.header('Saved Configurations')
    st.caption('Stores global pricing, tax, and cost settings only.')
    h = hist()
    if h:
        st.caption(f'{len(h)} saved configuration(s) available')
        names = [f"{x['time']} | {x['name']}" for x in reversed(h)]
        pick = st.selectbox('Configuration History', names)
        selected_index = names.index(pick)
        selected_item = list(reversed(h))[selected_index]

        if st.button('Load selected config', type='primary', width='stretch'):
            st.session_state.c = sanitize_config(selected_item.get('config', {}))
            st.session_state.auto = selected_item.get('auto', False)
            st.session_state.config_revision = st.session_state.get('config_revision', 0) + 1
            st.rerun()

        col_ovr, col_del = st.columns(2)
        if col_ovr.button('Update selected', width='stretch'):
            real_idx = len(h) - 1 - selected_index
            h[real_idx] = {'name': h[real_idx]['name'], 'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'config': sanitize_config(c), 'auto': st.session_state.auto}
            save_hist(h)
            st.session_state.saved_notice = True
            st.rerun()

        if col_del.button('Delete selected', width='stretch'):
            h.pop(len(h) - 1 - selected_index)
            save_hist(h)
            st.session_state.deleted_notice = True
            st.rerun()
    
    name = st.text_input('Configuration name')
    if st.button('Save as new config', width='stretch'):
        if name.strip():
            h.append({'name': name.strip(), 'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'config': sanitize_config(c), 'auto': st.session_state.auto})
            save_hist(h); st.session_state.saved_notice = True; st.rerun()
        else: st.error('Enter a configuration name.')

    st.header('Global Configuration')
    cr = st.session_state.config_revision

    def update_film_price():
        new_film = st.session_state[f"f_type_{cr}"]
        st.session_state.c['film'] = new_film
        st.session_state.c['film_cost'] = float(FILMS[new_film][0])
        st.session_state.c['film_cost_cur'] = 'USD'
        st.session_state.config_revision += 1

    st.selectbox('Film type', list(FILMS), index=list(FILMS).index(c['film']), key=f"f_type_{cr}", on_change=update_film_price)

    with st.form('calculator_inputs_form'):
        new_c = sanitize_config(c)
        new_c['film'] = c['film']
        
        col1, col2 = st.columns([3, 1])
        new_c['film_cost'] = col1.number_input('Film buying cost (/sqm)', 0., value=float(c.get('film_cost', FILMS[c['film']][0])), key=f"f_cost_{cr}")
        new_c['film_cost_cur'] = col2.selectbox('Currency', ['USD', 'AMD'], index=['USD', 'AMD'].index(c.get('film_cost_cur', 'USD')), key=f"f_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['manual_sell'] = col1.number_input('Manual selling price (/sqm)', 0., value=float(c.get('manual_sell', 52925.)), step=1000., key=f"m_sell_{cr}")
        new_c['manual_sell_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('manual_sell_cur', 'AMD')), key=f"m_scur_{cr}", label_visibility="hidden")

        new_c['enforce_500mm'] = st.checkbox('Enforce 500mm minimum width on customer billing', value=c.get('enforce_500mm', True), key=f"enf_500_{cr}")

        new_c['target_margin'] = st.number_input('Target net margin (%)', 0., 99., float(c['target_margin']), key=f"t_marg_{cr}")
        new_c['waste'] = st.number_input('Waste (%)', 0., 100., float(c['waste']), key=f"waste_{cr}")

        packaging_methods = ['Sheet box', 'Paper tube', 'Manual packages']
        current_method = c.get('method', 'Sheet box')
        if current_method not in packaging_methods: current_method = 'Sheet box'
        new_c['method'] = st.selectbox('Packaging method', packaging_methods, index=packaging_methods.index(current_method), help='Manual packages lets you enter project carton sizes and gross weights.', key=f"pack_m_{cr}")

        col1, col2 = st.columns([3, 1])
        new_c['ship_kg'] = col1.number_input('Shipping cost (/kg)', 0., value=float(c.get('ship_kg', 4000.)), key=f"sh_kg_{cr}")
        new_c['ship_kg_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('ship_kg_cur', 'AMD')), key=f"sh_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['fixed_ship'] = col1.number_input('Fixed shipping', 0., value=float(c.get('fixed_ship', 0.)), key=f"f_ship_{cr}")
        new_c['fixed_ship_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('fixed_ship_cur', 'AMD')), key=f"fs_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['clearance'] = col1.number_input('Clearance', 0., value=float(c.get('clearance', 35000.)), key=f"clear_{cr}")
        new_c['clearance_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('clearance_cur', 'AMD')), key=f"cl_cur_{cr}", label_visibility="hidden")

        st.markdown("#### Taxes & Fees")
        col1, col2, col3 = st.columns(3)
        new_c['import_tax'] = col1.number_input('Import tax (%)', 0., 100., float(c['import_tax']), key=f"i_tax_{cr}")
        new_c['alibaba_fee_pct'] = col2.number_input('Alibaba Fee (%)', 0., 100., float(c.get('alibaba_fee_pct', 3.0)), key=f"ali_fee_{cr}")
        new_c['input_tax_pct'] = col3.number_input('Input VAT/Tax (%)', 0., 100., float(c.get('input_tax_pct', 10.)), key=f"in_tax_{cr}")
        
        st.markdown("#### Turnover Tax")
        new_c['enable_turnover_tax'] = st.checkbox('Apply Turnover Tax', value=c.get('enable_turnover_tax', True), help="Uses the background rates defined in Settings.", key=f"tt_en_{cr}")
        new_c['include_labor_in_turnover_tax'] = st.checkbox('Include Labor Cost in Tax Deduction', value=c.get('include_labor_in_turnover_tax', True), help="Turn off to see how much your turnover tax increases if you cannot officially document labor expenses.", key=f"tt_lab_{cr}")
        
        new_c['turnover_tax_base'] = c.get('turnover_tax_base', 10.0)
        new_c['turnover_tax_deduct'] = c.get('turnover_tax_deduct', 6.0)
        new_c['turnover_tax_floor'] = c.get('turnover_tax_floor', 4.5)
        new_c['doc_expense_pct'] = c.get('doc_expense_pct', 100.0)

        st.markdown("#### Other Taxes")
        col1, col2 = st.columns(2)
        new_c['output_vat'] = col1.number_input('Output VAT (%)', 0., 100., float(c['output_vat']), key=f"out_tax_{cr}")
        new_c['dividend_tax'] = col2.number_input('Dividend tax (%)', 0., 100., float(c['dividend_tax']), key=f"dtax_{cr}")

        new_c['labor_mode'] = st.selectbox('Labor calculation mode', ['Per sqm', 'Per hour'], index=['Per sqm', 'Per hour'].index(c.get('labor_mode', 'Per sqm')), key=f"lab_m_{cr}")
        
        col1, col2 = st.columns([3, 1])
        new_c['labor_rate_sqm'] = col1.number_input('Labor rate (/sqm)', 0., value=float(c.get('labor_rate_sqm', 10000.)), step=500., key=f"lab_r_s_{cr}")
        new_c['labor_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('labor_cur', 'AMD')), key=f"l_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['labor_rate_hour'] = col1.number_input('Labor rate (/hour)', 0., value=float(c.get('labor_rate_hour', 10000.)), step=500., key=f"lab_r_h_{cr}")
        new_c['labor_hour_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('labor_hour_cur', 'AMD')), key=f"lh_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['other'] = col1.number_input('Other costs', 0., value=float(c.get('other', 0.)), key=f"oth_{cr}")
        new_c['other_cur'] = col2.selectbox('Currency', ['AMD', 'USD'], index=['AMD', 'USD'].index(c.get('other_cur', 'AMD')), key=f"oth_cur_{cr}", label_visibility="hidden")

        st.subheader('Exchange rate')
        modes = ['Automatic CBA', 'Manual']
        new_c['exchange_mode'] = st.radio('Exchange-rate mode', modes, index=modes.index(c.get('exchange_mode', 'Automatic CBA')), key=f"ex_mode_{cr}")
        new_c['manual_rate'] = st.number_input('Manual USD to AMD Rate', 0.01, value=float(c.get('manual_rate', c.get('rate', 365.))), key=f"man_r_{cr}")

        if c.get('exchange_mode') == 'Automatic CBA' and c.get('rate_date'):
            st.caption(f'Current CBA Rate: **{c["rate"]:.2f}** (Date: {c["rate_date"]})')

        st.subheader('Accessories Base Costs')
        col1, col2 = st.columns([3, 1])
        new_c['remote_cost'] = col1.number_input('Remote cost (/pc)', 0., value=float(c.get('remote_cost', 10.)), key=f"rem_c_{cr}")
        new_c['remote_cur'] = col2.selectbox('Currency', ['USD', 'AMD'], index=['USD', 'AMD'].index(c.get('remote_cur', 'USD')), key=f"rem_cur_{cr}", label_visibility="hidden")

        col1, col2 = st.columns([3, 1])
        new_c['silicone_cost'] = col1.number_input('Silicone cost (/pc)', 0., value=float(c.get('silicone_cost', 8.)), key=f"sil_c_{cr}")
        new_c['silicone_cur'] = col2.selectbox('Currency', ['USD', 'AMD'], index=['USD', 'AMD'].index(c.get('silicone_cur', 'USD')), key=f"sil_cur_{cr}", label_visibility="hidden")

        for pst in PS.keys():
            col1, col2 = st.columns([3, 1])
            new_c[f'ps_cost_{pst}'] = col1.number_input(f'{pst} PS cost (/pc)', 0., value=float(c.get(f'ps_cost_{pst}', DEFAULT[f'ps_cost_{pst}'])), key=f"ps_c_{pst}_{cr}")
            new_c[f'ps_cur_{pst}'] = col2.selectbox('Currency', ['USD', 'AMD'], index=['USD', 'AMD'].index(c.get(f'ps_cur_{pst}', 'USD')), key=f"ps_cur_{pst}_{cr}", label_visibility="hidden")

        apply_inputs = st.form_submit_button('Apply config changes', type='primary', width='stretch')

    if apply_inputs:
        if new_c['exchange_mode'] == 'Manual':
            new_c['rate'] = new_c['manual_rate']
            new_c['rate_source'] = 'Manual'
            new_c['rate_date'] = datetime.now().strftime('%Y-%m-%d')
        st.session_state.c = sanitize_config(new_c)
        st.session_state.config_revision += 1
        st.session_state.inputs_applied = True
        st.rerun()

    if c.get('exchange_mode', 'Automatic CBA') == 'Automatic CBA':
        if st.button('Refresh CBA exchange rate', width='stretch'):
            try:
                c['rate'], c['rate_date'] = fetch_cba_usd_rate()
                c['rate_source'] = 'Central Bank of Armenia'
                st.session_state.cba_error = ''
                st.rerun()
            except Exception as exc:
                st.session_state.cba_error = str(exc)
        if st.session_state.get('cba_error'):
            st.warning(f'CBA rate could not be refreshed. Using {c["rate"]:.2f} AMD/USD. Details: {st.session_state.cba_error}')

# BATCH & QUARTER CLOSE PAGE
if page == 'Batch & Quarter Close':
    st.title('Batching & Quarterly Optimizer')
    st.caption('Consolidate small projects under a flagship anchor project to compute shipping margins and real turnover tax.')
    
    if not projects:
        st.info("No projects saved. Create projects first to batch them.")
        st.stop()
        
    all_projects = sorted(projects.keys())
    selected_batch = st.multiselect("Select Projects for this Quarter/Batch", all_projects, default=[])
    
    if len(selected_batch) < 1:
        st.warning("Select at least one project to view calculations.")
        st.stop()
        
    anchor_proj = st.selectbox("Select Flagship/Anchor Project", selected_batch)
    
    st.markdown("### Combined Shipment Packaging")
    batch_pkg_mode = st.radio("Packaging Method for Combined Shipment:", ['Auto-Calculate Combined', 'Manual Combined Packaging'], horizontal=True)
    
    if st.session_state.get('last_anchor_proj') != anchor_proj:
        st.session_state.last_anchor_proj = anchor_proj
        anc_data = projects[anchor_proj]
        anc_c = anc_data.get('config', c)
        
        if anc_c.get('method') == 'Manual packages':
            pkgs = copy.deepcopy(anc_data.get('manual_packages', []))
            wpkgs = copy.deepcopy(anc_data.get('waste_manual_packages', []))
            st.session_state.batch_manual_pkgs = pkgs + wpkgs if (pkgs + wpkgs) else copy.deepcopy(DEFAULT_PKGS)
        else:
            anc_pdf = pd.DataFrame(upgrade_panels(anc_data.get('panels', [])))
            if not anc_pdf.empty:
                _, anc_k = costs(anc_pdf, anc_c, anc_data.get('power_supplies', []), anc_data.get('silicone_qty', 0), anc_data.get('remote_qty', 0), [], [])
                st.session_state.batch_manual_pkgs = anc_k.get('package_list', copy.deepcopy(DEFAULT_PKGS))
            else:
                st.session_state.batch_manual_pkgs = copy.deepcopy(DEFAULT_PKGS)
                
    if 'batch_manual_pkgs' not in st.session_state:
        st.session_state.batch_manual_pkgs = copy.deepcopy(DEFAULT_PKGS)
        
    if batch_pkg_mode == 'Manual Combined Packaging':
        mdf_batch = st.data_editor(pd.DataFrame(st.session_state.batch_manual_pkgs), num_rows='dynamic', hide_index=True, width='stretch', key='batch_manual_editor')
        b_rows = []
        for _, r in mdf_batch.fillna(0).iterrows():
            b_rows.append({'Quantity': safe_qty(r.get('Quantity', 0)), 'Length (cm)': safe_float(r.get('Length (cm)', 0)), 'Width (cm)': safe_float(r.get('Width (cm)', 0)), 'Height (cm)': safe_float(r.get('Height (cm)', 0)), 'Gross weight (kg)': safe_float(r.get('Gross weight (kg)', 0))})
        st.session_state.batch_manual_pkgs = b_rows

    st.divider()

    standalone_metrics = {}
    combined_panels = []
    combined_ps = []
    combined_silicone = 0
    total_rev = 0.0
    
    for p_name in selected_batch:
        p_data = projects[p_name]
        p_c = p_data.get('config', c)
        p_auto = p_data.get('auto', False)
        
        pdf = pd.DataFrame(upgrade_panels(p_data.get('panels', [])))
        if pdf.empty: continue
            
        _, p_k = costs(
            pdf, p_c, p_data.get('power_supplies', []), 
            p_data.get('silicone_qty', 0), p_data.get('remote_qty', 0), 
            p_data.get('manual_packages', []), p_data.get('waste_manual_packages', [])
        )
        
        billing_area = p_k['chargeable_area'] if p_c.get('enforce_500mm', True) else p_k['area']
        manual_sell_price_amd = to_amd(float(p_c.get('manual_sell', 52925.0)), p_c.get('manual_sell_cur', 'AMD'), p_c['rate'])
        p_rev = solve(p_k, p_c) if p_auto else billing_area * manual_sell_price_amd
        total_rev += p_rev
        
        standalone_freight = p_k['shipping'] + p_k['clearance']
        standalone_metrics[p_name] = {'freight': standalone_freight, 'area': p_k['chargeable_area']}
        
        combined_panels.extend(upgrade_panels(p_data.get('panels', [])))
        combined_ps.extend(p_data.get('power_supplies', []))
        combined_silicone += int(p_data.get('silicone_qty', 0))
        
    combined_df = pd.DataFrame(combined_panels)
    
    batch_c = projects[anchor_proj].get('config', c)
    if batch_pkg_mode == 'Manual Combined Packaging':
        batch_c['method'] = 'Manual packages'
        batch_man = st.session_state.batch_manual_pkgs
    else:
        batch_man = []
        
    batch_detail, batch_k = costs(combined_df, batch_c, combined_ps, combined_silicone, 0, batch_man, [])
    total_batched_freight = batch_k['shipping'] + batch_k['clearance']
    
    st.subheader("Batched Logistics & Extra Margin")
    total_charge_area = sum(m['area'] for m in standalone_metrics.values())
    
    margin_table = []
    total_savings = 0.0
    
    for p_name, m in standalone_metrics.items():
        ratio = m['area'] / total_charge_area if total_charge_area > 0 else 0
        allocated_freight = total_batched_freight * ratio
        savings = m['freight'] - allocated_freight
        total_savings += savings
        
        margin_table.append({
            'Project Name': p_name,
            'Role': 'Anchor' if p_name == anchor_proj else 'Small Project',
            'Standalone Freight (AMD)': f"{m['freight']:,.0f}",
            'Pro-Rated Batched Freight (AMD)': f"{allocated_freight:,.0f}",
            'Extra Margin Earned (AMD)': f"{savings:,.0f}"
        })
        
    st.dataframe(pd.DataFrame(margin_table), hide_index=True, width='stretch')
    st.success(f"**Total Extra Profit Generated via Consolidation:** {total_savings:,.0f} AMD")

    st.divider()
    st.subheader("Quarterly Financials & Turnover Tax")
    
    r_batch = profit(total_rev, batch_k, batch_c)
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Quarterly Revenue", f"{total_rev:,.0f} AMD")
    col2.metric("Combined Base Costs", f"{batch_k['base_total']:,.0f} AMD")
    
    if batch_c.get('enable_turnover_tax', True):
        col3.metric("True Aggregate Turnover Tax", f"{r_batch['turnover_tax']:,.0f} AMD")
        col4.metric("Total Net Profit", f"{r_batch['net_profit']:,.0f} AMD")
        
        st.markdown("### 💡 Quarterly Tax Expense Optimizer")
        if r_batch.get('opt_status') == 'room':
            st.info(f"📉 **Tax Optimization Alert:** Across this entire batched pipeline, LUMEON has not hit the tax floor. You can safely spend an additional **{r_batch['opt_amount']:,.0f} AMD** on documented expenses this quarter to aggressively lower your tax bill. Every 1,000 AMD spent reduces the combined tax by 60 AMD.")
        elif r_batch.get('opt_status') == 'wasted':
            st.warning(f"🛑 **Tax Floor Reached:** Across this entire batched pipeline, LUMEON has hit the 4.5% absolute floor. **{r_batch['opt_amount']:,.0f} AMD** of your aggregated expenses are bleeding over and providing NO tax benefit. **Do not document any more expenses.** Have suppliers date all new e-invoices for the 1st of next quarter.")
            
        st.markdown("#### 🧪 Simulate Expense Spending")
        simulated_exp = st.number_input("Simulate Additional Documented Expenses (AMD)", min_value=0.0, step=10000.0, help="Enter an amount (e.g. tools, marketing, inventory) to see exactly how it impacts your tax and net profit.")
        
        if simulated_exp > 0:
            eligible_expense_base = batch_k['base_total'] if batch_c.get('include_labor_in_turnover_tax', True) else (batch_k['base_total'] - batch_k.get('labor', 0.0))
            current_doc_expenses = eligible_expense_base * (batch_c.get('doc_expense_pct', 100.0) / 100.0)
            new_doc_expenses = current_doc_expenses + simulated_exp
            
            t_base_rate = batch_c.get('turnover_tax_base', 10.0)
            t_deduct_rate = batch_c.get('turnover_tax_deduct', 6.0)
            t_floor_rate = batch_c.get('turnover_tax_floor', 4.5)
            
            base_tax = total_rev * (t_base_rate / 100.0)
            min_tax = total_rev * (t_floor_rate / 100.0)
            
            new_deduction = new_doc_expenses * (t_deduct_rate / 100.0)
            new_tentative = base_tax - new_deduction
            new_tt = max(new_tentative, min_tax)
            
            tax_savings = r_batch['turnover_tax'] - new_tt
            new_net_profit = r_batch['net_profit'] - simulated_exp + tax_savings
            
            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("Simulated Turnover Tax", f"{new_tt:,.0f} AMD", f"-{tax_savings:,.0f} AMD saved", delta_color="inverse")
            sc2.metric("Effective Cost of Expense", f"{simulated_exp - tax_savings:,.0f} AMD", help="You spend X out of pocket, but save Y in taxes.")
            sc3.metric("Simulated Net Profit", f"{new_net_profit:,.0f} AMD", f"{new_net_profit - r_batch['net_profit']:,.0f} AMD")
            
            if new_tt == min_tax and r_batch['turnover_tax'] > min_tax:
                st.caption("⚠️ **Note:** This simulated expense just pushed you completely down into the 4.5% tax floor. Any further spending beyond this amount will NOT yield additional tax savings.")
            elif new_tt == min_tax and r_batch['turnover_tax'] == min_tax:
                st.caption("🛑 **Note:** You were already at the 4.5% floor. This expense provides absolutely ZERO tax savings this quarter. Save the invoice for next quarter!")

    else:
        col3.metric("Turnover Tax", "0 AMD (Disabled)")
        col4.metric("Total Net Profit", f"{r_batch['net_profit']:,.0f} AMD")
        st.write("Turnover Tax calculation is disabled in the Anchor project's configuration.")


# PROJECTS PAGE
if page == 'Projects':
    st.title('LUMEON Projects')
    st.caption('Select one project to view and edit its dimensions, packaging, accessories, and quote history.')
    names = sorted(projects)
    mode = st.radio('Action', ['Open existing project', 'Create new project'], horizontal=True)
    if mode == 'Create new project':
        new_name = st.text_input('New project name')
        if st.button('Create project', type='primary', disabled=not new_name.strip()):
            clean = new_name.strip()
            if clean in projects:
                st.error('A project with this name already exists.')
            else:
                projects[clean] = {
                    'panels': [], 'manual_packages': [], 'waste_manual_packages': [], 'quotes': [],
                    'power_supplies': [{'Type': '30W', 'Quantity': 0}],
                    'remote_qty': 0, 'silicone_qty': 0, 'clearance_w': 2.0, 'clearance_h': 2.0,
                    'client_info': {}, 'config': sanitize_config(st.session_state.c), 'auto': st.session_state.auto
                }
                save_projects(projects)
                st.session_state.project_to_open = clean
                st.rerun()
        st.stop()
    if not names:
        st.info('No projects saved.')
        st.stop()
    preferred = st.session_state.pop('project_to_open', None)
    idx = names.index(preferred) if preferred in names else 0
    project_name = st.selectbox('Project', names, index=idx, key='project_page_selector')

    p_data = projects.get(project_name, {})
    if isinstance(p_data, dict):
        dimensions = upgrade_panels(p_data.get('panels', []))
        manual_pkgs = copy.deepcopy(p_data.get('manual_packages', []))
        waste_pkgs = copy.deepcopy(p_data.get('waste_manual_packages', []))
        p_quotes = copy.deepcopy(p_data.get('quotes', []))
        power_supplies = copy.deepcopy(p_data.get('power_supplies', [{'Type': '30W', 'Quantity': 0}]))
        remote_qty = p_data.get('remote_qty', 0)
        silicone_qty = p_data.get('silicone_qty', 0)
        p_clear_w = p_data.get('clearance_w', 2.0)
        p_clear_h = p_data.get('clearance_h', 2.0)
        ci = p_data.get('client_info', {})
    else:
        dimensions = upgrade_panels(p_data)
        manual_pkgs = []; waste_pkgs = []; p_quotes = []; ci = {}
        power_supplies = [{'Type': '30W', 'Quantity': 0}]
        remote_qty = 0; silicone_qty = 0; p_clear_w = 2.0; p_clear_h = 2.0

    if not dimensions: dimensions = copy.deepcopy(DEFAULT_PANELS)
    if not manual_pkgs: manual_pkgs = copy.deepcopy(DEFAULT_PKGS)
    if not waste_pkgs: waste_pkgs = copy.deepcopy(DEFAULT_PKGS)

    total_project_area = calc_area(dimensions)
    st.subheader(f"Editing: {project_name} (Physical Area: {total_project_area:.2f} sqm)")
    
    with st.expander("Client Information"):
        qc1, qc2, qc3 = st.columns(3)
        p_c_type = qc1.selectbox("Client Type", ["Company", "Personal"], index=0 if ci.get('client_type', 'Company') == 'Company' else 1, key=f"p_ctype_{project_name}")
        p_c_warr = qc2.selectbox("Include Warranty?", ["Yes", "No"], index=0 if ci.get('warranty_yes_no', 'Yes') == 'Yes' else 1, key=f"p_warr_{project_name}")
        p_c_meas = qc3.selectbox("Measurements", ["Provided", "Measured"], index=0 if ci.get('measurement', 'Provided') == 'Provided' else 1, key=f"p_meas_{project_name}")
        
        c_col1, c_col2 = st.columns(2)
        p_c_name = c_col1.text_input("Client Name", value=ci.get('client_name', ''), key=f"p_cname_{project_name}")
        p_c_comp = c_col2.text_input("Client Company", value=ci.get('client_company', ''), key=f"p_ccomp_{project_name}")
        p_c_addr = c_col1.text_input("Client Address", value=ci.get('client_address', ''), key=f"p_cadd_{project_name}")
        p_c_phone = c_col2.text_input("Client Phone", value=ci.get('client_phone', ''), key=f"p_cphone_{project_name}")
        
        t_col1, t_col2, t_col3 = st.columns(3)
        p_c_wtext = t_col1.text_input("Warranty Details (Bookmark)", value=ci.get('warranty_text', ''), key=f"p_wtext_{project_name}")
        p_c_ttext = t_col2.text_input("Transformer Details (Bookmark)", value=ci.get('transformer_text', ''), key=f"p_ttext_{project_name}")
        p_c_rtext = t_col3.text_input("Remote Details (Bookmark)", value=ci.get('remote_text', ''), key=f"p_rtext_{project_name}")

    st.write("Panels")
    edited = st.data_editor(pd.DataFrame(dimensions), num_rows='dynamic', hide_index=True, width='stretch', key=f'project_editor_{project_name}')

    with st.expander("Import dimensions from CSV/Excel"):
        st.info("The file must contain exactly these column headers: 'Quantity', 'Width (mm)', 'Length (mm)'.")
        uploaded_file = st.file_uploader("Upload file", type=['csv', 'xlsx', 'xls'], key=f"upload_{project_name}")
        if uploaded_file:
            if st.button("Import Data", type="primary", key=f"btn_import_{project_name}"):
                try:
                    if uploaded_file.name.endswith('.csv'): idf = pd.read_csv(uploaded_file)
                    else: idf = pd.read_excel(uploaded_file)
                    req = ['Quantity', 'Width (mm)', 'Length (mm)']
                    if all(col in idf.columns for col in req):
                        new_p = []
                        for _, r in idf.iterrows():
                            q = safe_qty(r.get('Quantity', 0))
                            w = safe_float(r.get('Width (mm)', 0))
                            l = safe_float(r.get('Length (mm)', 0))
                            if q > 0 and w > 0 and l > 0:
                                new_p.append({'Quantity': q, 'Width (mm)': w, 'Length (mm)': l})
                        if new_p:
                            if not isinstance(projects[project_name], dict):
                                projects[project_name] = {'panels': new_p, 'manual_packages': [], 'waste_manual_packages': [], 'quotes': [], 'client_info': {}}
                            else:
                                projects[project_name]['panels'] = new_p
                            save_projects(projects)
                            st.success("Dimensions imported successfully! Reloading...")
                            st.rerun()
                        else: st.error("No valid dimensions found. Check data values.")
                    else: st.error("Missing required columns. Please check your headers.")
                except Exception as e: st.error(f"Error reading file: {e}")

    st.write("Project Accessories & Clearance")
    ps_df = pd.DataFrame(power_supplies)
    edited_ps = st.data_editor(ps_df, column_config={'Type': st.column_config.SelectboxColumn(options=list(PS.keys()))}, num_rows='dynamic', hide_index=True, width='stretch', key=f'proj_ps_{project_name}')

    c_col1, c_col2, c_col3, c_col4 = st.columns(4)
    new_remote_qty = c_col1.number_input('Additional Remotes', 0, value=int(remote_qty), key=f'p_rem_{project_name}')
    new_silicone_qty = c_col2.number_input('Silicone Qty', 0, value=int(silicone_qty), key=f'p_sil_{project_name}')
    new_clear_w = c_col3.number_input("Width deduct (mm)", min_value=0.0, value=float(p_clear_w), step=1.0, key=f"p_cw_{project_name}")
    new_clear_h = c_col4.number_input("Length deduct (mm)", min_value=0.0, value=float(p_clear_h), step=1.0, key=f"p_ch_{project_name}")

    st.write("Manual Packages (Order)")
    edited_pkgs = st.data_editor(pd.DataFrame(manual_pkgs), num_rows='dynamic', hide_index=True, width='stretch', key=f'pkg_editor_{project_name}')

    st.write("Manual Waste Packages (Separated)")
    edited_waste_pkgs = st.data_editor(pd.DataFrame(waste_pkgs), num_rows='dynamic', hide_index=True, width='stretch', key=f'waste_pkg_editor_{project_name}')

    c1, c2, c3 = st.columns(3)
    if c1.button('Save project updates', type='primary', width='stretch'):
        clean = []
        for _, r in edited.fillna(0).iterrows():
            q = safe_qty(r.get('Quantity', 0)); w = safe_float(r.get('Width (mm)', 0)); l = safe_float(r.get('Length (mm)', 0))
            if q > 0 and w > 0 and l > 0: clean.append({'Quantity': q, 'Width (mm)': w, 'Length (mm)': l})
        clean_pkgs = []
        for _, r in edited_pkgs.fillna(0).iterrows():
            q = safe_qty(r.get('Quantity', 0)); L = safe_float(r.get('Length (cm)', 0)); W = safe_float(r.get('Width (cm)', 0)); H = safe_float(r.get('Height (cm)', 0)); G = safe_float(r.get('Gross weight (kg)', 0))
            if q > 0 and (L > 0 or W > 0 or H > 0 or G > 0): clean_pkgs.append({'Quantity': q, 'Length (cm)': L, 'Width (cm)': W, 'Height (cm)': H, 'Gross weight (kg)': G})
        clean_waste_pkgs = []
        for _, r in edited_waste_pkgs.fillna(0).iterrows():
            q = safe_qty(r.get('Quantity', 0)); L = safe_float(r.get('Length (cm)', 0)); W = safe_float(r.get('Width (cm)', 0)); H = safe_float(r.get('Height (cm)', 0)); G = safe_float(r.get('Gross weight (kg)', 0))
            if q > 0 and (L > 0 or W > 0 or H > 0 or G > 0): clean_waste_pkgs.append({'Quantity': q, 'Length (cm)': L, 'Width (cm)': W, 'Height (cm)': H, 'Gross weight (kg)': G})
        clean_ps = edited_ps.fillna(0).to_dict('records')

        if not isinstance(projects[project_name], dict):
            projects[project_name] = {'quotes': []}
        projects[project_name]['panels'] = clean
        projects[project_name]['manual_packages'] = clean_pkgs
        projects[project_name]['waste_manual_packages'] = clean_waste_pkgs
        projects[project_name]['power_supplies'] = clean_ps
        projects[project_name]['remote_qty'] = new_remote_qty
        projects[project_name]['silicone_qty'] = new_silicone_qty
        projects[project_name]['clearance_w'] = new_clear_w
        projects[project_name]['clearance_h'] = new_clear_h
        projects[project_name]['client_info'] = {
            'client_type': p_c_type, 'warranty_yes_no': p_c_warr, 'measurement': p_c_meas,
            'client_name': p_c_name, 'client_company': p_c_comp, 'client_address': p_c_addr,
            'client_phone': p_c_phone, 'warranty_text': p_c_wtext, 'transformer_text': p_c_ttext,
            'remote_text': p_c_rtext
        }
        save_projects(projects)
        st.success('Project saved.')
        st.rerun()

    if c2.button('Load into calculator', width='stretch', disabled=not projects.get(project_name)):
        if isinstance(p_data, dict):
            st.session_state.panels = upgrade_panels(copy.deepcopy(p_data.get('panels', [])))
            st.session_state.power_supplies = copy.deepcopy(p_data.get('power_supplies', [{'Type': '30W', 'Quantity': 0}]))
            st.session_state.remote_qty = p_data.get('remote_qty', 0)
            st.session_state.silicone_qty = p_data.get('silicone_qty', 0)
            st.session_state.clearance_w = p_data.get('clearance_w', 2.0)
            st.session_state.clearance_h = p_data.get('clearance_h', 2.0)
            st.session_state.manual_packages = copy.deepcopy(p_data.get('manual_packages') or DEFAULT_PKGS)
            st.session_state.waste_manual_packages = copy.deepcopy(p_data.get('waste_manual_packages') or DEFAULT_PKGS)
            
            c_inf = p_data.get('client_info', {})
            st.session_state.q_client_type = c_inf.get('client_type', 'Company')
            st.session_state.q_warranty_yes_no = c_inf.get('warranty_yes_no', 'Yes')
            st.session_state.q_measurement = c_inf.get('measurement', 'Provided')
            st.session_state.q_client_name = c_inf.get('client_name', '')
            st.session_state.q_client_company = c_inf.get('client_company', '')
            st.session_state.q_client_address = c_inf.get('client_address', '')
            st.session_state.q_client_phone = c_inf.get('client_phone', '')
            st.session_state.q_warranty_text = c_inf.get('warranty_text', '')
            st.session_state.q_transformer_text = c_inf.get('transformer_text', '')
            st.session_state.q_remote_text = c_inf.get('remote_text', '')
            
            st.session_state.c = copy.deepcopy(p_data.get('config', sanitize_config(DEFAULT)))
            st.session_state.auto = p_data.get('auto', False)
        else:
            st.session_state.panels = upgrade_panels(copy.deepcopy(p_data))
            st.session_state.power_supplies = [{'Type': '30W', 'Quantity': 0}]
            st.session_state.remote_qty = 0
            st.session_state.silicone_qty = 0
            st.session_state.clearance_w = 2.0
            st.session_state.clearance_h = 2.0
            st.session_state.manual_packages = copy.deepcopy(DEFAULT_PKGS)
            st.session_state.waste_manual_packages = copy.deepcopy(DEFAULT_PKGS)
            st.session_state.c = sanitize_config(DEFAULT)
        st.session_state.panel_revision += 1
        st.session_state.config_revision += 1
        st.session_state.selected_project = project_name
        st.session_state.app_page = 'Calculator'
        st.rerun()

    confirm = st.checkbox('Confirm project deletion')
    if c3.button('Delete project', width='stretch', disabled=not confirm):
        projects.pop(project_name, None)
        save_projects(projects)
        st.rerun()

    st.divider()
    st.subheader(f"Quote History: {project_name}")

    # Synchronize with live Google Drive files
    live_drive_files = get_live_drive_files(project_name)
    if live_drive_files is not None:
        drive_filenames = {f["name"] for f in live_drive_files}
        drive_urls = {f["name"]: f["url"] for f in live_drive_files}

        # Filter out quotes that were deleted directly in Google Drive
        synced_quotes = []
        for q in p_quotes:
            fname = q.get('filename', '')
            prefix = fname.rsplit('.', 1)[0]
            # Check if either the PDF or DOCX still exists in Drive
            has_in_drive = any(df.startswith(prefix) for df in drive_filenames)
            if has_in_drive:
                # Attach live Drive link if available
                q['drive_url'] = drive_urls.get(fname, drive_urls.get(prefix + '.pdf', drive_urls.get(prefix + '.docx', '')))
                synced_quotes.append(q)

        if len(synced_quotes) != len(p_quotes):
            p_quotes = synced_quotes
            if not isinstance(projects[project_name], dict):
                projects[project_name] = {'panels': projects[project_name], 'manual_packages': [], 'waste_manual_packages': [], 'quotes': []}
            projects[project_name]['quotes'] = p_quotes
            save_projects(projects)

    if p_quotes:
        # Display clean history table
        display_data = []
        for q in p_quotes:
            display_data.append({
                "Date": q.get('timestamp', ''),
                "Filename": q.get('filename', ''),
                "Transparency": q.get('transparency', 'Unknown'),
                "Status": q.get('status', 'Generated')
            })
        st.dataframe(pd.DataFrame(display_data), hide_index=True, width='stretch')

        st.write("Manage Quotes:")
        q_cols = st.columns([4, 2, 2])
        q_options = [f"{q['timestamp']} | {q['filename']} | {q.get('transparency', 'Unknown')}" for q in p_quotes]
        sel_q = q_cols[0].selectbox("Select Quote", q_options, label_visibility="collapsed")
        sel_idx = q_options.index(sel_q)
        selected_quote = p_quotes[sel_idx]

        # 1. Mark as Latest Shared
        if q_cols[1].button("Mark Latest Shared", width='stretch'):
            target_transparency = selected_quote.get('transparency')
            if not target_transparency:
                for t in ['88%', '92%', '96%']:
                    if t in selected_quote['filename']:
                        target_transparency = t
                        break

            for q in p_quotes:
                q_trans = q.get('transparency')
                if not q_trans:
                    for t in ['88%', '92%', '96%']:
                        if t in q['filename']:
                            q_trans = t
                            break
                if q['status'] == 'Latest Shared' and q_trans == target_transparency:
                    q['status'] = 'Archived'

            p_quotes[sel_idx]['status'] = 'Latest Shared'
            if not isinstance(projects[project_name], dict):
                projects[project_name] = {'panels': projects[project_name], 'manual_packages': [], 'waste_manual_packages': [], 'quotes': []}
            projects[project_name]['quotes'] = p_quotes
            save_projects(projects)
            st.success(f"Quote marked as Latest Shared.")
            st.rerun()

        # 2. Delete Quote (From JSON AND from Google Drive)
        if q_cols[2].button("Delete Quote", type="primary", width='stretch'):
            deleted_q = p_quotes.pop(sel_idx)
            prefix = deleted_q['filename'].rsplit('.', 1)[0]

            # Trash the files in Google Drive
            delete_quote_from_drive(project_name, prefix)

            # Also clean local container disk if present
            safe_proj = "".join(ch for ch in project_name if ch.isalnum() or ch in (' ', '_', '-')).strip() or "Manual_Entry"
            base_dir = os.path.dirname(os.path.abspath(__file__))
            proj_dir = os.path.join(base_dir, "Generated Quotes", safe_proj)
            for ext in ['.pdf', '.docx']:
                f_path = os.path.join(proj_dir, prefix + ext)
                if os.path.exists(f_path):
                    try: os.remove(f_path)
                    except: pass

            if not isinstance(projects[project_name], dict):
                projects[project_name] = {'panels': projects[project_name], 'manual_packages': [], 'waste_manual_packages': [], 'quotes': []}
            projects[project_name]['quotes'] = p_quotes
            save_projects(projects)
            st.success(f"Quote '{deleted_q['filename']}' deleted from project and Google Drive.")
            st.rerun()

        # Direct Google Drive link for selected quote
        if selected_quote.get('drive_url'):
            st.markdown(f"☁️ [Open Selected Quote in Google Drive]({selected_quote['drive_url']})")

    else:
        st.info("No quotes have been generated for this project yet.")

    st.stop()


# CALCULATOR PAGE
if page == 'Calculator':
    st.title('LUMEON PDLC Film Price Calculator')

    st.subheader('Project selection')
    project_names = ['Manual entry'] + sorted(projects)
    current = st.session_state.get('selected_project', 'Manual entry')
    current = current if current in project_names else 'Manual entry'
    selected = st.selectbox('Project', project_names, index=project_names.index(current), key='calculator_project_selector')

    pc1, pc2, pc3 = st.columns(3)
    if pc1.button('Load selected project', type='primary', width='stretch', disabled=selected == 'Manual entry'):
        p_data = projects.get(selected, {})
        if isinstance(p_data, dict):
            if p_data.get('panels'):
                st.session_state.panels = upgrade_panels(copy.deepcopy(p_data['panels']))
            st.session_state.power_supplies = copy.deepcopy(p_data.get('power_supplies', [{'Type': '30W', 'Quantity': 0}]))
            st.session_state.remote_qty = p_data.get('remote_qty', 0)
            st.session_state.silicone_qty = p_data.get('silicone_qty', 0)
            st.session_state.clearance_w = p_data.get('clearance_w', 2.0)
            st.session_state.clearance_h = p_data.get('clearance_h', 2.0)
            st.session_state.manual_packages = copy.deepcopy(p_data.get('manual_packages') or DEFAULT_PKGS)
            st.session_state.waste_manual_packages = copy.deepcopy(p_data.get('waste_manual_packages') or DEFAULT_PKGS)
            c_inf = p_data.get('client_info', {})
            st.session_state.q_client_type = c_inf.get('client_type', 'Company')
            st.session_state.q_warranty_yes_no = c_inf.get('warranty_yes_no', 'Yes')
            st.session_state.q_measurement = c_inf.get('measurement', 'Provided')
            st.session_state.q_client_name = c_inf.get('client_name', '')
            st.session_state.q_client_company = c_inf.get('client_company', '')
            st.session_state.q_client_address = c_inf.get('client_address', '')
            st.session_state.q_client_phone = c_inf.get('client_phone', '')
            st.session_state.q_warranty_text = c_inf.get('warranty_text', '')
            st.session_state.q_transformer_text = c_inf.get('transformer_text', '')
            st.session_state.q_remote_text = c_inf.get('remote_text', '')
            
            st.session_state.c = copy.deepcopy(p_data.get('config', sanitize_config(DEFAULT)))
            st.session_state.auto = p_data.get('auto', False)
        else:
            st.session_state.panels = upgrade_panels(copy.deepcopy(p_data))
            st.session_state.power_supplies = [{'Type': '30W', 'Quantity': 0}]
            st.session_state.remote_qty = 0
            st.session_state.silicone_qty = 0
            st.session_state.clearance_w = 2.0
            st.session_state.clearance_h = 2.0
            st.session_state.manual_packages = copy.deepcopy(DEFAULT_PKGS)
            st.session_state.waste_manual_packages = copy.deepcopy(DEFAULT_PKGS)
            
            st.session_state.q_client_name = ''
            st.session_state.q_client_company = ''
            st.session_state.q_client_address = ''
            st.session_state.q_client_phone = ''
            st.session_state.q_warranty_text = ''
            st.session_state.q_transformer_text = ''
            st.session_state.q_remote_text = ''
            st.session_state.q_client_type = 'Company'
            st.session_state.q_warranty_yes_no = 'Yes'
            st.session_state.q_measurement = 'Provided'
            
        st.session_state.panel_revision += 1
        st.session_state.config_revision += 1
        st.session_state.selected_project = selected
        st.rerun()

    if pc2.button('Save to selected project', width='stretch', disabled=selected == 'Manual entry'):
        if not isinstance(projects[selected], dict):
            projects[selected] = {'quotes': []}
        projects[selected]['panels'] = copy.deepcopy(st.session_state.get('current_panels_output', st.session_state.panels))
        projects[selected]['manual_packages'] = copy.deepcopy(st.session_state.manual_packages)
        projects[selected]['waste_manual_packages'] = copy.deepcopy(st.session_state.waste_manual_packages)
        projects[selected]['power_supplies'] = copy.deepcopy(st.session_state.power_supplies)
        projects[selected]['remote_qty'] = st.session_state.remote_qty
        projects[selected]['silicone_qty'] = st.session_state.silicone_qty
        projects[selected]['clearance_w'] = st.session_state.clearance_w
        projects[selected]['clearance_h'] = st.session_state.clearance_h
        
        projects[selected]['config'] = sanitize_config(st.session_state.c)
        projects[selected]['auto'] = st.session_state.auto
        
        projects[selected]['client_info'] = {
            'client_type': st.session_state.q_client_type,
            'warranty_yes_no': st.session_state.q_warranty_yes_no,
            'measurement': st.session_state.q_measurement,
            'client_name': st.session_state.q_client_name,
            'client_company': st.session_state.q_client_company,
            'client_address': st.session_state.q_client_address,
            'client_phone': st.session_state.q_client_phone,
            'warranty_text': st.session_state.q_warranty_text,
            'transformer_text': st.session_state.q_transformer_text,
            'remote_text': st.session_state.q_remote_text
        }
        save_projects(projects)
        st.success(f"Project '{selected}' successfully updated!")

    if pc3.button('Clear project data', width='stretch'):
        st.session_state.panels = [{'Quantity': 0, 'Width (mm)': 0.0, 'Length (mm)': 0.0}]
        st.session_state.manual_packages = copy.deepcopy(DEFAULT_PKGS)
        st.session_state.waste_manual_packages = copy.deepcopy(DEFAULT_PKGS)
        st.session_state.power_supplies = [{'Type': '30W', 'Quantity': 0}]
        st.session_state.remote_qty = 0
        st.session_state.silicone_qty = 0
        st.session_state.q_client_name = ''
        st.session_state.q_client_company = ''
        st.session_state.q_client_address = ''
        st.session_state.q_client_phone = ''
        st.session_state.q_warranty_text = ''
        st.session_state.q_transformer_text = ''
        st.session_state.q_remote_text = ''
        st.session_state.q_client_type = 'Company'
        st.session_state.q_warranty_yes_no = 'Yes'
        st.session_state.q_measurement = 'Provided'
        st.session_state.c = sanitize_config(DEFAULT)
        st.session_state.panel_revision += 1
        st.session_state.config_revision += 1
        st.session_state.selected_project = 'Manual entry'
        st.rerun()

    if c.get('method') == 'Manual packages':
        st.subheader('Project Manual Package Details')
        st.caption('Stored directly inside the Project (not in Global Config). Gross weight must include packed contents.')
        mdf = st.data_editor(
            pd.DataFrame(st.session_state.manual_packages),
            num_rows='dynamic', hide_index=True, width='stretch',
            key=f'manual_package_editor_{st.session_state.panel_revision}'
        )
        package_rows = []
        for _, r in mdf.fillna(0).iterrows():
            package_rows.append({
                'Quantity': safe_qty(r.get('Quantity', 0)),
                'Length (cm)': safe_float(r.get('Length (cm)', 0)),
                'Width (cm)': safe_float(r.get('Width (cm)', 0)),
                'Height (cm)': safe_float(r.get('Height (cm)', 0)),
                'Gross weight (kg)': safe_float(r.get('Gross weight (kg)', 0))
            })
        st.session_state.manual_packages = package_rows

        if float(c.get('waste', 0.0)) > 0:
            st.markdown("**Separated Waste Package Handling**")
            waste_opts = ['Auto-calculate', 'Enter manually']
            curr_w_mode = c.get('waste_pkg_mode', 'Auto-calculate')
            if curr_w_mode not in waste_opts: curr_w_mode = 'Auto-calculate'
            selected_w_mode = st.radio(
                "How would you like to handle the separated waste package dimensions?",
                waste_opts,
                index=waste_opts.index(curr_w_mode),
                format_func=lambda x: "Auto-calculate and append separated waste package" if x == 'Auto-calculate' else "Enter separated waste package dimensions manually",
                horizontal=True,
                key=f"waste_pkg_mode_radio_{st.session_state.panel_revision}"
            )
            c['waste_pkg_mode'] = selected_w_mode
            st.session_state.c = c

            if selected_w_mode == 'Enter manually':
                w_mdf = st.data_editor(
                    pd.DataFrame(st.session_state.waste_manual_packages),
                    num_rows='dynamic', hide_index=True, width='stretch',
                    key=f'waste_manual_package_editor_{st.session_state.panel_revision}'
                )
                w_package_rows = []
                for _, r in w_mdf.fillna(0).iterrows():
                    w_package_rows.append({
                        'Quantity': safe_qty(r.get('Quantity', 0)),
                        'Length (cm)': safe_float(r.get('Length (cm)', 0)),
                        'Width (cm)': safe_float(r.get('Width (cm)', 0)),
                        'Height (cm)': safe_float(r.get('Height (cm)', 0)),
                        'Gross weight (kg)': safe_float(r.get('Gross weight (kg)', 0))
                    })
                st.session_state.waste_manual_packages = w_package_rows

        st.info('Chargeable weight is the larger of total gross weight and total volumetric weight. Volumetric divisor: 5,000.')

    df_preview = pd.DataFrame(st.session_state.panels)
    st.subheader('Panel dimensions')
    df = st.data_editor(df_preview, num_rows='dynamic', hide_index=True, width='stretch', key=f'panel_editor_{st.session_state.panel_revision}')
    st.session_state.current_panels_output = df.fillna(0).to_dict('records')
    live_area = calc_area(st.session_state.current_panels_output)
    st.caption(f'Physical Area: **{live_area:.4f} sqm**')

    if any(0 < safe_float(p.get('Width (mm)', 0)) < 50 or 0 < safe_float(p.get('Length (mm)', 0)) < 50 for p in st.session_state.current_panels_output):
        st.warning('⚠️ One or more panel dimensions are smaller than 50 mm. Make sure dimensions are entered in **millimeters (mm)** (e.g. `700` and `2000`), not meters!')

    st.subheader("Project Accessories")
    st.write("Power Supplies")
    ps_df = pd.DataFrame(st.session_state.power_supplies)
    edited_ps = st.data_editor(ps_df, column_config={'Type': st.column_config.SelectboxColumn(options=list(PS.keys()))}, num_rows='dynamic', hide_index=True, width='stretch', key=f'calc_ps_editor_{st.session_state.panel_revision}')
    st.session_state.power_supplies = edited_ps.fillna(0).to_dict('records')

    ac1, ac2 = st.columns(2)
    st.session_state.remote_qty = ac1.number_input('Additional Remotes per PS', 0, value=int(st.session_state.remote_qty), key=f'rem_{st.session_state.panel_revision}')
    st.session_state.silicone_qty = ac2.number_input('Silicone qty', 0, value=int(st.session_state.silicone_qty), key=f'sil_{st.session_state.panel_revision}')

    st.subheader("Manufacturing & Installation Cut-List")
    st.caption("Apply specific clearance reductions for Width and Length to all selected project dimensions.")
    cl_col1, cl_col2 = st.columns(2)
    cw_key = f"cw_{st.session_state.panel_revision}"
    ch_key = f"ch_{st.session_state.panel_revision}"
    clearance_w_mm = cl_col1.number_input("Width deduction (mm)", min_value=0.0, value=float(st.session_state.clearance_w), step=1.0, key=cw_key)
    clearance_h_mm = cl_col2.number_input("Length deduction (mm)", min_value=0.0, value=float(st.session_state.clearance_h), step=1.0, key=ch_key)
    st.session_state.clearance_w = clearance_w_mm
    st.session_state.clearance_h = clearance_h_mm

    adj_df = df.copy()
    for col in ['Quantity', 'Width (mm)', 'Length (mm)']:
        if col in adj_df.columns:
            adj_df[col] = pd.to_numeric(adj_df[col], errors='coerce').fillna(0)

    adj_df['Adjusted Width (mm)'] = (adj_df['Width (mm)'] - clearance_w_mm).clip(lower=0)
    adj_df['Adjusted Length (mm)'] = (adj_df['Length (mm)'] - clearance_h_mm).clip(lower=0)
    adj_df['Total Area (sqm)'] = adj_df['Quantity'] * (adj_df['Adjusted Width (mm)'] / 1000.0) * (adj_df['Adjusted Length (mm)'] / 1000.0)

    display_columns = ['Quantity', 'Width (mm)', 'Length (mm)', 'Adjusted Width (mm)', 'Adjusted Length (mm)', 'Total Area (sqm)']
    cutlist_csv = adj_df[display_columns].round(4)
    
    for c_name in display_columns:
        if c_name != 'Quantity':
            cutlist_csv[c_name] = cutlist_csv[c_name].map(lambda x: f"{x:g}" if not pd.isna(x) else "")

    acc_data = []
    remotes_per_ps = 1 + int(st.session_state.remote_qty)
    for ps in st.session_state.power_supplies:
        q = safe_qty(ps.get('Quantity', 0))
        if q > 0:
            acc_data.append({
                'Quantity': str(q),
                'Width (mm)': '-',
                'Length (mm)': '-',
                'Adjusted Width (mm)': f"{ps.get('Type', '')} Power Supply",
                'Adjusted Length (mm)': f"w/ {remotes_per_ps} Remote(s) each",
                'Total Area (sqm)': '-'
            })

    if int(st.session_state.silicone_qty) > 0:
        acc_data.append({
            'Quantity': str(int(st.session_state.silicone_qty)),
            'Width (mm)': '-',
            'Length (mm)': '-',
            'Adjusted Width (mm)': 'Silicone Tube(s)',
            'Adjusted Length (mm)': '-',
            'Total Area (sqm)': '-'
        })

    if acc_data:
        sep = pd.DataFrame([{'Quantity': '---', 'Width (mm)': '---', 'Length (mm)': '---', 'Adjusted Width (mm)': '---', 'Adjusted Length (mm)': '---', 'Total Area (sqm)': '---'}])
        cutlist_csv = pd.concat([cutlist_csv, sep, pd.DataFrame(acc_data)], ignore_index=True)

    cutlist_total_area = adj_df['Total Area (sqm)'].sum()
    summary_row = pd.DataFrame([{'Quantity': 'TOTAL', 'Width (mm)': '', 'Length (mm)': '', 'Adjusted Width (mm)': '', 'Adjusted Length (mm)': '', 'Total Area (sqm)': f"{cutlist_total_area:.3f}"}])
    cutlist_csv = pd.concat([cutlist_csv, summary_row], ignore_index=True)

    st.dataframe(cutlist_csv, hide_index=True, width='stretch')
    
    csv_string = cutlist_csv.to_csv(index=False)
    st.download_button(label="Download Cut-List (CSV)", data=csv_string.encode('utf-8'), file_name=f"{st.session_state.get('selected_project', 'Project')}_cutlist.csv", mime="text/csv", key='csv_dl_button')

    a, b = st.columns(2)
    if a.button('Calculate price for target margin', type='primary'):
        st.session_state.auto = True
    if b.button('Use manual selling price', type='primary'):
        st.session_state.auto = False
        st.session_state.manual_notice = True
        st.rerun()

    detail, k = costs(
        df, c,
        st.session_state.power_supplies,
        st.session_state.silicone_qty,
        st.session_state.remote_qty,
        st.session_state.manual_packages,
        st.session_state.waste_manual_packages
    )
    if k['area'] <= 0:
        st.info('Enter at least one valid panel.')
        st.stop()

    manual_sell_price_amd = to_amd(float(c.get('manual_sell', 52925.0)), c.get('manual_sell_cur', 'AMD'), c['rate'])
    billing_area = k['chargeable_area'] if c.get('enforce_500mm', True) else k['area']
    rev = solve(k, c) if st.session_state.auto else billing_area * manual_sell_price_amd

    r = profit(rev, k, c)
    sellamd = rev / billing_area
    cols = st.columns(4)
    cols[0].metric('Selling price / sqm', f'{sellamd:,.0f} AMD', f'${sellamd/c["rate"]:,.2f}')
    cols[1].metric('Revenue', f'{rev:,.0f} AMD', f'${rev/c["rate"]:,.2f}')
    cols[2].metric('Gross profit', f'{r["gross_profit"]:,.0f} AMD', f'${r["gross_profit"]/c["rate"]:,.2f}')
    cols[3].metric('Net profit', f'{r["net_profit"]:,.0f} AMD', f'${r["net_profit"]/c["rate"]:,.2f} | {r["margin"]:.2f}%')

    st.subheader('Package and chargeable weight calculation')
    q = st.columns(5)
    q[0].metric('Packages', k['packages'])
    q[1].metric('Packaging tare', f'{k["tare"]:.2f} kg')
    q[2].metric('Gross weight', f'{k["gross"]:.2f} kg')
    q[3].metric('Volumetric weight', f'{k["volumetric"]:.2f} kg')
    q[4].metric('Chargeable weight', f'{k["chargeable"]:.2f} kg')
    st.caption(k['description'])
    if k['warning']:
        st.warning(k['warning'])

    st.subheader('Automatic labor calculation')
    labor_rate_val = float(c.get('labor_rate_sqm', 10000.)) if c['labor_mode'] == 'Per sqm' else float(c.get('labor_rate_hour', 10000.))
    labor_rate_cur = c.get('labor_cur', 'AMD') if c['labor_mode'] == 'Per sqm' else c.get('labor_hour_cur', 'AMD')
    labor_rate_unit = 'sqm' if c['labor_mode'] == 'Per sqm' else 'hour'

    l1, l2, l3 = st.columns(3)
    l1.metric('Labor rate', f"{labor_rate_val:,.0f} {labor_rate_cur}/{labor_rate_unit}")
    l2.metric('Labor cost', f'{k["labor"]:,.0f} AMD')
    l3.metric('Estimated labor time', f'{k["labor_hours"]:,.2f} hours')
    st.caption(f'Calculation mode: {c["labor_mode"]}. Large panels: 2 hours each when height > 1.5 m AND width > 0.6 m. All other panels: 1.25 hours each.')
    st.dataframe(detail.round(3), hide_index=True, width='stretch')

    st.subheader('Cost summary')
    
    if c.get('enable_turnover_tax', True):
        turnover_tax_str = f'Turnover Tax (Tentative: {r["tentative_tax"]:,.0f}, Min Floor: {r["min_tax"]:,.0f})'
        turnover_row = (turnover_tax_str, r['turnover_tax'])
    else:
        turnover_row = ('Turnover Tax (Disabled)', 0.0)
    
    rows = [
        ('Film', k['film_amd']),
        ('Accessories', k['accessories_cost']),
        ('Alibaba Fee', k['alibaba_fee']),
        ('Shipping', k['shipping']),
        ('Clearance', to_amd(float(c.get('clearance', 0)), c.get('clearance_cur', 'AMD'), c['rate'])),
        ('Import tax', k['duty']),
        ('Labor', k['labor']),
        ('Other costs', to_amd(float(c.get('other', 0)), c.get('other_cur', 'AMD'), c['rate'])),
        ('Base cost before input tax', k['base_total']),
        (f'Input VAT / tax ({c.get("input_tax_pct", 10.0):.1f}% of revenue)', r['input_tax']),
        ('Output VAT', r['output_vat']),
        ('Total cost incl. input tax', r['total_cost']),
        ('Revenue', rev),
        ('Gross profit', r['gross_profit']),
        turnover_row,
        ('Dividend Tax', r['dividend_tax']),
        ('Net profit', r['net_profit'])
    ]
    out = pd.DataFrame(rows, columns=['Item', 'AMD'])
    out['USD equivalent'] = out.AMD / c['rate']
    st.dataframe(out.round(2), hide_index=True, width='stretch')

    st.subheader("Generate Quote Document")
    with st.form("quote_generation_form"):
        q_col1, q_col2, q_col3 = st.columns(3)
        client_type = q_col1.selectbox("Client Type", ["Company", "Personal"], key="q_client_type")
        warranty_yes_no = q_col2.selectbox("Include Warranty?", ["Yes", "No"], key="q_warranty_yes_no")
        measurement = q_col3.selectbox("Measurements", ["Provided", "Measured"], key="q_measurement")

        c_col1, c_col2 = st.columns(2)
        client_name = c_col1.text_input("Client Name", key="q_client_name")
        client_company = c_col2.text_input("Client Company", key="q_client_company")
        client_address = c_col1.text_input("Client Address", key="q_client_address")
        client_phone = c_col2.text_input("Client Phone", key="q_client_phone")

        t_col1, t_col2, t_col3 = st.columns(3)
        warranty_text = t_col1.text_input("Warranty Details (Bookmark)", key="q_warranty_text")
        transformer_text = t_col2.text_input("Transformer Details (Bookmark)", key="q_transformer_text")
        remote_text = t_col3.text_input("Remote Details (Bookmark)", key="q_remote_text")

        q_bot1, q_bot2 = st.columns([2, 1])
        manual_date = q_bot1.date_input("Issue Date (Optional)", value=None, key="q_manual_date")
        compile_pdf = q_bot2.checkbox("Compile Official PDF (Uses CloudConvert credits)", value=True, help="Uncheck when just testing dimensions or pricing to save your daily credits.")

        submitted = st.form_submit_button("Generate Quote", type="primary", width='stretch')

    if submitted:
        if not client_name:
            st.error("Please enter a Client Name.")
        else:
            with st.spinner("Generating document..."):
                proj_name = st.session_state.get('selected_project', 'Manual entry')

                if proj_name != 'Manual entry':
                    if not isinstance(projects.get(proj_name), dict):
                        projects[proj_name] = {'quotes': []}
                    projects[proj_name]['client_info'] = {
                        'client_type': client_type, 'warranty_yes_no': warranty_yes_no, 'measurement': measurement,
                        'client_name': client_name, 'client_company': client_company, 'client_address': client_address,
                        'client_phone': client_phone, 'warranty_text': warranty_text, 'transformer_text': transformer_text,
                        'remote_text': remote_text
                    }
                    save_projects(projects)

                file_path, msg = generate_quote(
                    client_type=client_type,
                    warranty_yes_no=warranty_yes_no,
                    measurement=measurement,
                    client_name=client_name,
                    client_company=client_company,
                    client_address=client_address,
                    client_phone=client_phone,
                    warranty_text=warranty_text,
                    transformer_text=transformer_text,
                    remote_text=remote_text,
                    manual_date=manual_date,
                    unit_price=sellamd,
                    total_amd=rev,
                    panels_df=pd.DataFrame(st.session_state.get('current_panels_output', st.session_state.panels)),
                    transparency=c["film"],
                    project_name=proj_name,
                    enforce_500mm=c.get('enforce_500mm', True),
                    compile_pdf=compile_pdf
                )

                if file_path and os.path.exists(file_path):
                    st.session_state.generated_quote_path = file_path
                    st.session_state.quote_success_msg = "Quote Generated Successfully!"

                    # Identify both DOCX and PDF paths
                    if file_path.endswith('.pdf'):
                        pdf_path = file_path
                        docx_path = file_path[:-4] + '.docx'
                    else:
                        docx_path = file_path
                        pdf_path = file_path[:-5] + '.pdf'

                    # Automatically upload both files to Google Drive (in project subfolder)
                    drive_links = []
                    if os.path.exists(docx_path):
                        link_docx = upload_file_to_drive(docx_path, project_name=proj_name)
                        if link_docx:
                            drive_links.append(("Word", link_docx))

                    if os.path.exists(pdf_path):
                        link_pdf = upload_file_to_drive(pdf_path, project_name=proj_name)
                        if link_pdf:
                            drive_links.append(("PDF", link_pdf))

                    st.session_state.drive_links = drive_links

                    # Update project quote history
                    if proj_name != 'Manual entry':
                        synced_projects = load_projects()
                        if not isinstance(synced_projects.get(proj_name), dict):
                            synced_projects[proj_name] = {'panels': synced_projects.get(proj_name, []), 'manual_packages': [], 'waste_manual_packages': [], 'quotes': [], 'client_info': {}}
                        if 'quotes' not in synced_projects[proj_name]:
                            synced_projects[proj_name]['quotes'] = []

                        synced_projects[proj_name]['quotes'].append({
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'filename': os.path.basename(file_path),
                            'status': 'Generated',
                            'transparency': c["film"]
                        })
                        save_projects(synced_projects)
                        projects[proj_name] = synced_projects[proj_name]
                else:
                    st.error(msg)
                    
    if st.session_state.get('generated_quote_path'):
        st.success(st.session_state.get('quote_success_msg', 'Ready for download!'))
        file_path = st.session_state.generated_quote_path

        # Identify both PDF and DOCX paths automatically
        if file_path.endswith('.pdf'):
            pdf_path = file_path
            docx_path = file_path[:-4] + '.docx'
        else:
            docx_path = file_path
            pdf_path = file_path[:-5] + '.pdf'

        col1, col2 = st.columns(2)

        # 1. Download PDF button (if generated)
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f_pdf:
                col1.download_button(
                    label="📕 Download PDF",
                    data=f_pdf,
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf",
                    type="primary",
                    width="stretch"
                )

        # 2. Download Word (.docx) button
        if os.path.exists(docx_path):
            with open(docx_path, "rb") as f_docx:
                col2.download_button(
                    label="📄 Download Word (.docx)",
                    data=f_docx,
                    file_name=os.path.basename(docx_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    width="stretch"
                )

        # Google Drive Direct Links
        if st.session_state.get('drive_links'):
            st.markdown("---")
            for ftype, url in st.session_state.drive_links:
                st.markdown(f"☁️ [Open {ftype} in Google Drive]({url})")
