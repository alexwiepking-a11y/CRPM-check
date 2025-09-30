import pandas as pd
from datetime import datetime
import os
from collections import defaultdict
from tqdm import tqdm
import webbrowser

# ============================================================================  
# CONFIGURATION SECTION  
# ============================================================================  
CITY_TAX_HOTELS = [
    'AMS', 'AMA','AMZ', 'RTM', 'NYT', 'NYB', 'PGL', 
    'PCG', 'POP', 'PLD', 'GEN', 'ZUR', 'RIT', 'KLB'
]

EXCLUDED_HOTELS = [
    'ITA', 'VRS', 'VRSM',
    'NEW_HOTEL1', 'NEW_HOTEL2'
]

CREATE_DASHBOARD = True
AUTO_OPEN_DASHBOARD = True
BATCH_SIZE = 1000

OUTPUT_FOLDER = "CRPM_Results"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

print(f"🏨 City tax will be checked for {len(CITY_TAX_HOTELS)} hotels: {', '.join(CITY_TAX_HOTELS)}")
# ============================================================================  

def should_check_city_tax(hotel_code):
    return hotel_code in CITY_TAX_HOTELS

# --- Exceptions loader (from test_crpm, cleaned) ---  
def load_exceptions(exceptions_file="source_CRPM_exceptions.xlsx"):
    if not os.path.exists(exceptions_file):
        # Create template
        template_df = pd.DataFrame({
            'Rule_Type': [
                'Hotel_Specific', 
                'Country_Pattern', 
                'Hotel_Pattern',
                'Country_Rate_Pattern',
                'Hotel_Rate_Specific',
                'Hotel_Rate_Pattern'
            ],
            'Field': ['VAT', 'VAT', 'Subaccount', 'VAT', 'Subaccount', 'Subaccount'],
            'Hotel_Code': ['HOTEL001', '', 'NYB,NYT', '', 'HOTEL001', 'NYB,NYT'],
            'Rate_Code': ['', '', '', 'MRYC,MRYE,MRYF,MRYA,MRYG', 'SPECIAL_RATE', 'MRYC,MRYE,MRYF,MRYA,MRYG'],
            'Country': ['', 'UK', '', 'UK', '', ''],
            'Current_Value': ['Without', 'Without', '108000A', 'Without', '108000A', '108000A'],
            'Standard_Value': ['Reduced', 'Reduced', '108000', 'Normal', '108000', '108000'],
            'Reason': [
                'Local regulation', 
                'UK hotels can use Without VAT', 
                'NYB/NYT use special subaccount for ALL rates',
                'UK special rates approved for Without VAT',
                'This hotel-rate combo uses special subaccount',
                'NYB/NYT use special subaccount for SPECIFIC rates only'
            ],
            'Approved_By': ['Manager']*6,
            'Date_Added': ['2024-01-15']*6,
            'Status': ['Active']*6,
            'Priority': ['High','Medium','Low','Medium','High','Medium'],
            'Review_Date': ['2025-01-15','2025-06-15','2024-12-15','2025-03-15','2025-01-15','2025-03-15'],
            'Notes': ['']*6
        })
        template_df.to_excel(exceptions_file, index=False)
        print(f"📝 Created exceptions template: {exceptions_file}")
        return pd.DataFrame()

    df = pd.read_excel(exceptions_file, dtype=str).fillna('')
    if 'Status' not in df.columns:
        df['Status'] = 'Active'
    else:
        df['Status'] = df['Status'].replace('', 'Active')
        df.loc[df['Status'].isna(), 'Status'] = 'Active'

    active = df[df['Status'].str.lower() == 'active']
    inactive = df[df['Status'].str.lower() != 'active']

    print(f"✅ Loaded {len(df)} exceptions ({len(active)} active, {len(inactive)} inactive)")
    return active

# --- Exception matching (from source_crpm) ---  
def is_deviation_accepted(exceptions_df, hotel_code, rate_code, country, field, current_value, standard_value):
    if len(exceptions_df) == 0:
        return False, None
    
    for _, exc in exceptions_df.iterrows():
        if exc['Field'].lower() != field.lower():
            continue
        if (str(exc.get('Current_Value','')).lower().strip() != str(current_value).lower().strip() or
            str(exc.get('Standard_Value','')).lower().strip() != str(standard_value).lower().strip()):
            continue

        match = False
        rt = exc['Rule_Type']

        if rt == 'Hotel_Specific' and hotel_code == exc.get('Hotel_Code',''):
            match = True
        elif rt == 'Country_Pattern' and country == exc.get('Country',''):
            match = True
        elif rt == 'Hotel_Pattern':
            if hotel_code in [h.strip() for h in exc.get('Hotel_Code','').split(',')]:
                match = True
        elif rt == 'Country_Rate_Pattern':
            if country == exc.get('Country',''):
                if rate_code in [r.strip() for r in exc.get('Rate_Code','').split(',')]:
                    match = True
        elif rt == 'Hotel_Rate_Specific':
            if hotel_code == exc.get('Hotel_Code','') and rate_code == exc.get('Rate_Code',''):
                match = True
        elif rt == 'Hotel_Rate_Pattern':
            if hotel_code in [h.strip() for h in exc.get('Hotel_Code','').split(',')]:
                if rate_code in [r.strip() for r in exc.get('Rate_Code','').split(',')]:
                    match = True
        elif rt == 'Rate_Pattern':
            if rate_code in [r.strip() for r in exc.get('Rate_Code','').split(',')]:
                match = True

        if match:
            return True, {
                'rule_type': rt,
                'reason': exc.get('Reason',''),
                'approved_by': exc.get('Approved_By',''),
                'priority': exc.get('Priority','Medium'),
                'review_date': exc.get('Review_Date',''),
                'notes': exc.get('Notes','')
            }
    return False, None

# --- Boolean normalizer ---  
def normalize_boolean(val):
    val_str = str(val).strip().lower()
    if val_str in ["true","1","1.0","yes","y"]:
        return "Yes"
    if val_str in ["false","0","0.0","no","n","nan","none"]:
        return "No"
    return "Unknown"

# ============================================================================  
# MAIN  
# ============================================================================  
print("🚀 Starting CRPM analysis...")

exceptions_df = load_exceptions("source_CRPM_exceptions.xlsx")

file = "source_CRPM_check.xlsx"
try:
    print("📊 Loading data...")
    data = pd.read_excel(file, sheet_name="data", dtype=str)
    standard = pd.read_excel(file, sheet_name="standard", dtype=str)
    data.columns = data.columns.str.strip()
    standard.columns = standard.columns.str.strip()
except Exception as e:
    raise RuntimeError(f"❌ Error loading {file}: {e}")

data["CityTax_Current"] = data["Is subject to city tax current"].apply(normalize_boolean)
standard["CityTax_Standard"] = standard["Standard City tax"].apply(normalize_boolean)

merged = data.merge(standard, how="left", on="Hotel code")
analysis = merged[merged["Standard subaccount"].notna()].copy()
print(f"📈 Analyzing {len(analysis)} entries with standards")

true_devs, accepted_devs = [], []
city_tax_skipped = 0

# --- Core deviation detection (from source_crpm, kept clean) ---  
for batch_start in tqdm(range(0, len(analysis), BATCH_SIZE), desc="Processing compliance"):
    batch = analysis.iloc[batch_start:batch_start+BATCH_SIZE]
    for _, row in batch.iterrows():
        hotel, rate, country = row['Hotel code'], row.get('Code',''), row.get('Country','')
        issues, accepted, exc_meta = [], [], []

        # Subaccount
        sub_c, sub_s = str(row["Sub account current"]).strip(), str(row["Standard subaccount"]).strip()
        if sub_c != sub_s:
            txt = f"Subaccount: '{sub_c}' → '{sub_s}'"
            ok, meta = is_deviation_accepted(exceptions_df, hotel, rate, country, "Subaccount", sub_c, sub_s)
            (accepted if ok else issues).append(txt)
            if ok: exc_meta.append(meta)

        # VAT
        vat_c, vat_s = str(row["Vat type current"]).strip(), str(row["Standard VAT"]).strip()
        if vat_c.lower() != vat_s.lower():
            txt = f"VAT: '{vat_c}' → '{vat_s}'"
            ok, meta = is_deviation_accepted(exceptions_df, hotel, rate, country, "VAT", vat_c, vat_s)
            (accepted if ok else issues).append(txt)
            if ok: exc_meta.append(meta)

        # City tax
        if should_check_city_tax(hotel):
            city_c, city_s = row["CityTax_Current"], row["CityTax_Standard"]
            if city_c != city_s:
                txt = f"City Tax: '{city_c}' → '{city_s}'"
                ok, meta = is_deviation_accepted(exceptions_df, hotel, rate, country, "CityTax", city_c, city_s)
                (accepted if ok else issues).append(txt)
                if ok: exc_meta.append(meta)
        else:
            city_tax_skipped += 1

        base = {
            'Hotel_Code': hotel,
            'Rate_Code': rate,
            'Country': country,
            'Sub_Account_Current': sub_c,
            'Sub_Account_Standard': sub_s,
            'VAT_Current': vat_c,
            'VAT_Standard': vat_s,
            'CityTax_Current': row["CityTax_Current"],
            'CityTax_Standard': row["CityTax_Standard"],
            'CityTax_Checked': should_check_city_tax(hotel),
        }

        if issues:
            r = base.copy()
            r['Deviation_Details'] = ' | '.join(issues)
            r['Status'] = 'NEEDS_FIXING'
            r['Priority'] = 'High' if len(issues) > 1 else 'Medium'
            true_devs.append(r)

        if accepted:
            r = base.copy()
            r['Deviation_Details'] = ' | '.join(accepted)
            r['Status'] = 'ACCEPTED'
            if exc_meta:
                r.update(exc_meta[0])
            accepted_devs.append(r)

print(f"🏨 City tax skipped for {city_tax_skipped} entries")

# === Results summary ===
total = len(analysis)
raw_compliant = total - len(true_devs) - len(accepted_devs)
true_compliant = total - len(true_devs)
print("="*80)
print("📊 COMPLIANCE SUMMARY")
print("="*80)
print(f"Total rate plans: {total}")
print(f"Perfect matches: {raw_compliant} ({raw_compliant/total*100:.1f}%)")
print(f"Accepted deviations: {len(accepted_devs)}")
print(f"Issues requiring attention: {len(true_devs)} ({len(true_devs)/total*100:.1f}%)")

# === Save outputs ===
true_df = pd.DataFrame(true_devs)
acc_df = pd.DataFrame(accepted_devs)

true_file = os.path.join(OUTPUT_FOLDER, "true_deviations.xlsx")
acc_file = os.path.join(OUTPUT_FOLDER, "accepted_deviations.xlsx")

true_df.to_excel(true_file, index=False)
acc_df.to_excel(acc_file, index=False)

print(f"💾 Saved true deviations to {true_file}")
print(f"💾 Saved accepted deviations to {acc_file}")

# === Dashboard generation ===
if CREATE_DASHBOARD:
    dash_file = os.path.join(OUTPUT_FOLDER, "compliance_dashboard.html")
    compliance_rate = (true_compliant / total * 100) if total else 0
    accepted_rate = (len(accepted_devs) / total * 100) if total else 0
    issue_rate = (len(true_devs) / total * 100) if total else 0

    html = f"""
    <html>
    <head>
    <title>CRPM Compliance Dashboard</title>
    <style>
    body {{ font-family: Arial, sans-serif; margin: 30px; }}
    h1 {{ color: #333; }}
    .metric {{ padding: 10px; margin: 10px; border-radius: 8px; display:inline-block; }}
    .ok {{ background:#d4edda; }}
    .warn {{ background:#fff3cd; }}
    .bad {{ background:#f8d7da; }}
    </style>
    </head>
    <body>
    <h1>CRPM Compliance Dashboard</h1>
    <div class='metric ok'>✅ Perfect matches: {raw_compliant} ({raw_compliant/total*100:.1f}%)</div>
    <div class='metric warn'>🟡 Accepted deviations: {len(accepted_devs)} ({accepted_rate:.1f}%)</div>
    <div class='metric bad'>❌ Issues: {len(true_devs)} ({issue_rate:.1f}%)</div>
    <hr>
    <p>City tax skipped for {city_tax_skipped} entries.</p>
    <p>Files saved in: {OUTPUT_FOLDER}</p>
    </body>
    </html>
    """
    with open(dash_file, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"📊 Dashboard saved to {dash_file}")
    if AUTO_OPEN_DASHBOARD:
        webbrowser.open('file://' + os.path.realpath(dash_file))
