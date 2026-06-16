import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import re

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def parse_date(date_str):
    if not date_str:
        return None
    for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None

def parse_mix_ratio(mix_ratio_str, gas_type):
    if not mix_ratio_str:
        if gas_type == 'ACM':
            return {'Argon': 0.9073, 'CO2': 0.1928} # Defaults for 90/10 mix (7 Cum -> 6.3512 Cum Arg, 1.35 KG CO2)
        elif gas_type == 'AHM':
            return {'Argon': 0.9886} # Default for AHM 92/8
        return {}
    
    mix_ratio_str = mix_ratio_str.upper()
    arg_match = re.search(r'(\d+)%\s*(?:ARG|ARGON)', mix_ratio_str)
    co2_match = re.search(r'(\d+)%\s*(?:CO2|CARBON)', mix_ratio_str)
    n2_match = re.search(r'(\d+)%\s*(?:N2|NITROGEN)', mix_ratio_str)
    oxy_match = re.search(r'(\d+)%\s*(?:O2|OXY|OXYGEN)', mix_ratio_str)
    
    parts = {}
    if arg_match: parts['Argon'] = float(arg_match.group(1)) / 100.0
    if co2_match: parts['CO2'] = float(co2_match.group(1)) / 100.0
    if n2_match: parts['N2'] = float(n2_match.group(1)) / 100.0
    if oxy_match: parts['Oxygen'] = float(oxy_match.group(1)) / 100.0
    
    if not parts:
        # Check slash format like 90/10 or 80/20
        slash_match = re.search(r'(\d+)\s*/\s*(\d+)', mix_ratio_str)
        if slash_match:
            val1 = float(slash_match.group(1))
            val2 = float(slash_match.group(2))
            if gas_type == 'ACM':
                if val1 == 90 and val2 == 10:
                    return {'Argon': 0.9073, 'CO2': 0.1928}
                elif val1 == 80 and val2 == 20:
                    return {'Argon': 0.80, 'CO2': 0.20}
                else:
                    total = val1 + val2
                    return {'Argon': val1 / total, 'CO2': val2 / total}
            elif gas_type == 'AHM':
                total = val1 + val2
                return {'Argon': val1 / total}
                
    # Fallback if parsing failed
    if not parts:
        if gas_type == 'ACM':
            return {'Argon': 0.9073, 'CO2': 0.1928}
        elif gas_type == 'AHM':
            return {'Argon': 0.9886}
            
    return parts

try:
    creds = Credentials.from_service_account_file(
        "credentials.json",
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    doc = client.open("Cylinder Tracking")
    
    cyl_ws = doc.worksheet("Cylinders")
    cyl_maint_ws = doc.worksheet("Cylinder Maintenance")
    scan_ws = doc.worksheet("Sheet1")
    
    # Load all data
    cyl_rows = cyl_ws.get_all_values()
    maint_rows = cyl_maint_ws.get_all_values()
    scan_rows = scan_ws.get_all_values()
    
    # ── Map maintenance specifications
    maint_data = {}
    if len(maint_rows) > 1:
        for r in maint_rows[1:]:
            uid = r[0].strip().upper()
            if not uid: continue
            maint_data[uid] = {
                'water_capacity': r[1].strip(),
                'fill_pressure': r[2].strip(),
                'gas_capacity': r[3].strip(),
                'unit': r[4].strip(),
                'is_mixture': r[5].strip(),
                'mix_ratio': r[6].strip(),
            }
            
    # ── Map cylinders
    cylinders = []
    if len(cyl_rows) > 1:
        for r in cyl_rows[1:]:
            uid = r[0].strip()
            if not uid: continue
            cylinders.append({
                'uid': uid,
                'gas_type': r[1].strip(),
                'cylinder_type': r[2].strip(),
                'owner': r[3].strip(),
                'status': r[4].strip(),
                'location': r[5].strip(),
                'last_activity': r[6].strip()
            })
            
    # ── Map scans
    scans = []
    if len(scan_rows) > 1:
        for r in scan_rows[1:]:
            scans.append({
                'date': r[0].strip(),
                'time': r[1].strip(),
                'driver': r[2].strip(),
                'action': r[3].strip(),
                'uid': r[4].strip(),
                'customer': r[5].strip() if len(r) > 5 else ''
            })

    print(f"Loaded {len(cylinders)} cylinders, {len(maint_data)} maintenance records, {len(scans)} scans.")

    # ── Test Calculation logic for Table 1 (Filled Cylinder Inventory)
    # We define our standard products config
    products_config = [
        {'id': 'arg_pura', 'name': 'ARG Pura', 'gas_type': 'ARG', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'acm_90_10', 'name': 'ACM (90.10)_', 'gas_type': 'ACM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.3512, 'unit': 'Cum'},
        {'id': 'co2_90_10', 'name': 'Co2 (90.10)_', 'gas_type': 'ACM', 'cylinder_type': 'Standard', 'gas_per_cyl': 1.35, 'unit': 'KG', 'is_virtual': True},
        {'id': 'co2_pure', 'name': 'Co2', 'gas_type': 'CO2', 'cylinder_type': 'Standard', 'gas_per_cyl': 30.0, 'unit': 'KG'},
        {'id': 'n2_cyl', 'name': 'N2 Cyl', 'gas_type': 'N2', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'oxygen_pure', 'name': 'OXYGEN', 'gas_type': 'OXY', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'ahm_92_08', 'name': 'AHM(92.08)', 'gas_type': 'AHM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.92, 'unit': 'Cum'},
        {'id': 'ahm_98_02', 'name': 'AHM (98.02)', 'gas_type': 'AHM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.98, 'unit': 'Cum'},
        {'id': 'arg_dura', 'name': 'ARG Dura', 'gas_type': 'ARG', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.0, 'unit': 'Cum'},
        {'id': 'n2_dura', 'name': 'N2Dura', 'gas_type': 'N2', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.88, 'unit': 'Cum'},
        {'id': 'oxygen_dura', 'name': 'Oxygen Dura', 'gas_type': 'OXY', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.0, 'unit': 'Cum'}
    ]

    # Initialize results
    t1_rows = {p['id']: {**p, 'filled_count': 0, 'total_gas': 0.0} for p in products_config}

    # Count and calculate
    for c in cylinders:
        if c['status'] != 'Filled':
            continue
        
        uid_upper = c['uid'].upper()
        maint = maint_data.get(uid_upper, {})
        
        gas_type = c['gas_type'].upper()
        cyl_type = c['cylinder_type'].capitalize()
        mix_ratio = maint.get('mix_ratio', '')
        
        # Determine capacity
        try:
            capacity = float(maint.get('gas_capacity') or 0.0) if maint.get('gas_capacity') else None
        except ValueError:
            capacity = None
            
        # Match product ID
        pid = None
        if gas_type == 'ARG':
            pid = 'arg_pura' if cyl_type == 'Standard' else 'arg_dura'
        elif gas_type == 'CO2':
            pid = 'co2_pure'
        elif gas_type == 'N2':
            pid = 'n2_cyl' if cyl_type == 'Standard' else 'n2_dura'
        elif gas_type == 'OXY':
            pid = 'oxygen_pure' if cyl_type == 'Standard' else 'oxygen_dura'
        elif gas_type == 'ACM':
            pid = 'acm_90_10' # Default
        elif gas_type == 'AHM':
            if '98' in mix_ratio:
                pid = 'ahm_98_02'
            else:
                pid = 'ahm_92_08'
                
        if pid:
            if pid == 'acm_90_10':
                # Split ACM
                t1_rows['acm_90_10']['filled_count'] += 1
                t1_rows['co2_90_10']['filled_count'] += 1
                
                # Apply ratios
                cap_val = capacity if capacity is not None else 7.0
                t1_rows['acm_90_10']['total_gas'] += cap_val * 0.9073
                t1_rows['co2_90_10']['total_gas'] += cap_val * 0.1928
            else:
                t1_rows[pid]['filled_count'] += 1
                cap_val = capacity if capacity is not None else t1_rows[pid]['gas_per_cyl']
                t1_rows[pid]['total_gas'] += cap_val

    # Print Table 1 Results
    print("\nTable 1 — Cyl Status Dispatch Stock (Filled Cylinder Inventory):")
    print(f"{'Product':<15} | {'Filled':<6} | {'Gas/Cyl':<8} | {'Total':<10} | {'Unit':<5}")
    print("-" * 55)
    total_physical_filled = 0
    total_cum = 0.0
    total_kg = 0.0
    
    for r_id, r in t1_rows.items():
        if not r.get('is_virtual'):
            total_physical_filled += r['filled_count']
        if r['unit'] == 'Cum':
            total_cum += r['total_gas']
        elif r['unit'] == 'KG':
            total_kg += r['total_gas']
            
        print(f"{r['name']:<15} | {r['filled_count']:<6} | {r['gas_per_cyl']:<8.4f} | {r['total_gas']:<10.4f} | {r['unit']:<5}")
        
    print("-" * 55)
    print(f"{'TOTAL FILLED':<15} | {total_physical_filled:<6} | {'':<8} | {f'{total_cum:.2f} Cum + {total_kg:.2f} KG':<10} | Mixed")

except Exception as e:
    print("Error:", e)
