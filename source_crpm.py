import pandas as pd
from datetime import datetime
import os
from collections import defaultdict

# ============================================================================
# CONFIGURATION SECTION - EASY TO MODIFY
# ============================================================================
# Hotels that actually use city tax - modify this list as needed
CITY_TAX_HOTELS = [
    'AMS', 'AMA','AMZ', 'RTM', 'NYT', 'NYB', 'PGL', 
    'PCG', 'POP', 'PLD', 'GEN', 'ZUR', 'RIT', 'KLB'
]

# Add new hotels here when needed:
# CITY_TAX_HOTELS.append('NEW_HOTEL_CODE')
# or
# CITY_TAX_HOTELS.extend(['HOTEL1', 'HOTEL2', 'HOTEL3'])

EXCLUDED_HOTELS = [
    'ITA', 'VRS', 'VRSM',
    'NEW_HOTEL1', 'NEW_HOTEL2'  # Add new ones here
]

print(f"🏨 City tax will be checked for {len(CITY_TAX_HOTELS)} hotels: {', '.join(CITY_TAX_HOTELS)}")
# ============================================================================

def should_check_city_tax(hotel_code):
    """Return True only for hotels that actually use city tax"""
    return hotel_code in CITY_TAX_HOTELS

def load_exceptions(exceptions_file="source_CRPM_exceptions.xlsx"):
    """Load accepted deviations from Excel file"""
    if not os.path.exists(exceptions_file):
        # Create enhanced exceptions template with rate support
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
            'Approved_By': ['Manager', 'Manager', 'Manager', 'Manager', 'Manager', 'Manager'],
            'Date_Added': ['2024-01-15', '2024-01-15', '2024-01-15', '2024-01-15', '2024-01-15', '2024-01-15'],
            'Status': ['Active', 'Active', 'Active', 'Active', 'Active', 'Active'],
            'Priority': ['High', 'Medium', 'Low', 'Medium', 'High', 'Medium'],
            'Review_Date': ['2025-01-15', '2025-06-15', '2024-12-15', '2025-03-15', '2025-01-15', '2025-03-15'],
            'Notes': ['', '', '', '', '', '']
        })
        template_df.to_excel(exceptions_file, index=False)
        print(f"📝 Created enhanced exceptions template: {exceptions_file}")
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(exceptions_file, dtype=str)
        # Fill NaN values with empty strings for easier comparison
        df = df.fillna('')
        
        # Fix the Status column handling
        if 'Status' not in df.columns:
            df['Status'] = 'Active'  # Default all to Active if column doesn't exist
        else:
            # Replace empty/NaN status values with 'Active'
            df['Status'] = df['Status'].replace('', 'Active')
            df.loc[df['Status'].isna(), 'Status'] = 'Active'
        
        # Filter for active exceptions - fix the filtering logic
        active_exceptions = df[df['Status'].str.lower() == 'active']
        
        print(f"✅ Loaded {len(df)} total exceptions")
        print(f"✅ {len(active_exceptions)} exceptions are ACTIVE")
        print(f"⏸️ {len(df) - len(active_exceptions)} exceptions are INACTIVE")
        
        # Show inactive exceptions for debugging
        inactive_exceptions = df[df['Status'].str.lower() != 'active']
        if len(inactive_exceptions) > 0:
            print("\n⚠️ INACTIVE EXCEPTIONS (will be ignored):")
            for idx, exc in inactive_exceptions.iterrows():
                rule_type = exc.get('Rule_Type', 'Unknown')
                field = exc.get('Field', 'Unknown')
                hotel = exc.get('Hotel_Code', '')
                reason = exc.get('Reason', '')
                status = exc.get('Status', '')
                print(f"   - {rule_type} | {field} | {hotel} | Status: '{status}' | {reason}")
        
        # Check for exceptions needing review
        today = datetime.now()
        review_needed = 0
        for _, exception in active_exceptions.iterrows():
            review_date = exception.get('Review_Date', '')
            if review_date and review_date != '':
                try:
                    review_dt = pd.to_datetime(review_date)
                    if review_dt <= today:
                        review_needed += 1
                except:
                    pass
        
        if review_needed > 0:
            print(f"⚠️ {review_needed} active exceptions need review - check Review_Date column")
        
        return active_exceptions
    except Exception as e:
        print(f"⚠️ Could not load exceptions: {e}")
        return pd.DataFrame()

def is_deviation_accepted(exceptions_df, hotel_code, rate_code, country, field, current_value, standard_value):
    """Enhanced function to check if a deviation is in the accepted exceptions"""
    if len(exceptions_df) == 0:
        return False, None
    
    for idx, exception in exceptions_df.iterrows():
        if exception['Field'].lower() != field.lower():
            continue
        
        # Check if values match
        if (str(exception.get('Current_Value', '')).lower().strip() != str(current_value).lower().strip() or
            str(exception.get('Standard_Value', '')).lower().strip() != str(standard_value).lower().strip()):
            continue
        
        match_found = False
        rule_type = exception['Rule_Type']
        
        # Hotel-specific exception
        if rule_type == 'Hotel_Specific':
            if hotel_code == exception.get('Hotel_Code', ''):
                match_found = True
        
        # Country-based exception
        elif rule_type == 'Country_Pattern':
            if country == exception.get('Country', ''):
                match_found = True
        
        # Hotel pattern exception
        elif rule_type == 'Hotel_Pattern':
            hotel_patterns = str(exception.get('Hotel_Code', '')).split(',')
            hotel_patterns = [p.strip() for p in hotel_patterns]
            if hotel_code in hotel_patterns:
                match_found = True
        
        # Country + Rate pattern exception
        elif rule_type == 'Country_Rate_Pattern':
            if country == exception.get('Country', ''):
                rate_patterns = str(exception.get('Rate_Code', '')).split(',')
                rate_patterns = [p.strip() for p in rate_patterns]
                if rate_code in rate_patterns:
                    match_found = True
        
        # Hotel + Rate specific exception
        elif rule_type == 'Hotel_Rate_Specific':
            if (hotel_code == exception.get('Hotel_Code', '') and 
                rate_code == exception.get('Rate_Code', '')):
                match_found = True
        
        # Hotel + Rate pattern exception
        elif rule_type == 'Hotel_Rate_Pattern':
            hotel_patterns = str(exception.get('Hotel_Code', '')).split(',')
            hotel_patterns = [p.strip() for p in hotel_patterns]
            if hotel_code in hotel_patterns:
                rate_patterns = str(exception.get('Rate_Code', '')).split(',')
                rate_patterns = [p.strip() for p in rate_patterns]
                if rate_code in rate_patterns:
                    match_found = True
        
        # Rate pattern only
        elif rule_type == 'Rate_Pattern':
            rate_patterns = str(exception.get('Rate_Code', '')).split(',')
            rate_patterns = [p.strip() for p in rate_patterns]
            if rate_code in rate_patterns:
                match_found = True
        
        if match_found:
            return True, {
                'rule_type': rule_type,
                'reason': exception.get('Reason', ''),
                'approved_by': exception.get('Approved_By', ''),
                'priority': exception.get('Priority', 'Medium'),
                'review_date': exception.get('Review_Date', ''),
                'notes': exception.get('Notes', '')
            }
    
    return False, None

def generate_exception_suggestions(true_deviations, min_occurrences=3):
    """Suggest new exception rules based on patterns in deviations"""
    suggestions = []
    
    # Group by deviation patterns
    patterns = defaultdict(list)
    
    for dev in true_deviations:
        # Create pattern keys
        hotel_key = dev['Hotel_Code']
        country_key = dev['Country']
        rate_key = dev['Rate_Code']
        
        # Parse deviation details to get field and values
        if 'VAT:' in dev['Deviation_Details']:
            field = 'VAT'
            detail = dev['Deviation_Details'].split('VAT: ')[1].split(' |')[0]
            if ' → ' in detail:
                current, standard = detail.replace("'", "").split(' → ')
                patterns[f"VAT_{current}_{standard}"].append({
                    'hotel': hotel_key, 'country': country_key, 'rate': rate_key
                })
        
        if 'Subaccount:' in dev['Deviation_Details']:
            field = 'Subaccount'
            detail = dev['Deviation_Details'].split('Subaccount: ')[1].split(' |')[0]
            if ' → ' in detail:
                current, standard = detail.replace("'", "").split(' → ')
                patterns[f"Subaccount_{current}_{standard}"].append({
                    'hotel': hotel_key, 'country': country_key, 'rate': rate_key
                })
    
    # Analyze patterns and suggest rules
    for pattern_key, occurrences in patterns.items():
        if len(occurrences) >= min_occurrences:
            field, current_val, standard_val = pattern_key.split('_', 2)
            
            # Check if it's a country pattern
            countries = set(occ['country'] for occ in occurrences)
            hotels = set(occ['hotel'] for occ in occurrences)
            rates = set(occ['rate'] for occ in occurrences)
            
            if len(countries) == 1 and len(hotels) > 1:
                country = list(countries)[0]
                if len(rates) > 1:
                    # Country + Rate pattern
                    suggestions.append({
                        'Rule_Type': 'Country_Rate_Pattern',
                        'Field': field,
                        'Hotel_Code': '',
                        'Rate_Code': ','.join(sorted(rates)),
                        'Country': country,
                        'Current_Value': current_val,
                        'Standard_Value': standard_val,
                        'Reason': f'Multiple {country} hotels with {field} deviation ({len(occurrences)} instances)',
                        'Approved_By': '[TO_BE_APPROVED]',
                        'Date_Added': datetime.now().strftime('%Y-%m-%d'),
                        'Status': 'Inactive',
                        'Priority': 'High' if len(occurrences) > 10 else 'Medium',
                        'Occurrences': len(occurrences)
                    })
                else:
                    # Country pattern
                    suggestions.append({
                        'Rule_Type': 'Country_Pattern',
                        'Field': field,
                        'Hotel_Code': '',
                        'Rate_Code': '',
                        'Country': country,
                        'Current_Value': current_val,
                        'Standard_Value': standard_val,
                        'Reason': f'All {country} hotels with {field} deviation ({len(occurrences)} instances)',
                        'Approved_By': '[TO_BE_APPROVED]',
                        'Date_Added': datetime.now().strftime('%Y-%m-%d'),
                        'Status': 'Inactive',
                        'Priority': 'High' if len(occurrences) > 10 else 'Medium',
                        'Occurrences': len(occurrences)
                    })
            
            elif len(hotels) <= 5 and len(rates) > 1:
                # Hotel + Rate pattern
                suggestions.append({
                    'Rule_Type': 'Hotel_Rate_Pattern',
                    'Field': field,
                    'Hotel_Code': ','.join(sorted(hotels)),
                    'Rate_Code': ','.join(sorted(rates)),
                    'Country': '',
                    'Current_Value': current_val,
                    'Standard_Value': standard_val,
                    'Reason': f'Specific hotels with specific rates {field} deviation ({len(occurrences)} instances)',
                    'Approved_By': '[TO_BE_APPROVED]',
                    'Date_Added': datetime.now().strftime('%Y-%m-%d'),
                    'Status': 'Inactive',
                    'Priority': 'Medium',
                    'Occurrences': len(occurrences)
                })
    
    return suggestions

print("🚀 Starting ENHANCED CRPM analysis with city tax configuration...")

# Load exceptions
exceptions_df = load_exceptions("source_CRPM_exceptions.xlsx")

# Load data
file = "source_CRPM_check.xlsx"
data = pd.read_excel(file, sheet_name="data", dtype=str)
standard = pd.read_excel(file, sheet_name="standard", dtype=str)

# Clean column names
data.columns = data.columns.str.strip()
standard.columns = standard.columns.str.strip()

print(f"📊 Loaded {len(data)} rate plans")
print(f"📋 Loaded {len(standard)} hotel standards")

# Normalize city tax
def normalize_boolean(val):
    val_str = str(val).strip().lower()
    if val_str in ["true", "1", "1.0", "yes", "y"]:
        return "Yes"
    if val_str in ["false", "0", "0.0", "no", "n", "nan", "none"]:
        return "No"
    return "Unknown"

data["CityTax_Current"] = data["Is subject to city tax current"].apply(normalize_boolean)
standard["CityTax_Standard"] = standard["Standard City tax"].apply(normalize_boolean)

# Merge 
merged = data.merge(standard, how="left", on="Hotel code")

# Filter out entries without standards
analysis = merged[merged["Standard subaccount"].notna()].copy()
print(f"📈 Analyzing {len(analysis)} entries with standards")

# Find deviations and categorize them
true_deviations = []
accepted_deviations = []
city_tax_skipped = 0

print("🔍 Checking for deviations (with intelligent city tax handling)...")

for idx, row in analysis.iterrows():
    issues = []
    accepted_issues = []
    exception_details = []
    
    hotel_code = row['Hotel code']
    rate_code = row.get('Code', '')
    country = row.get('Country', '')
    
    # Check subaccount
    sub_current = str(row["Sub account current"]).strip()
    sub_standard = str(row["Standard subaccount"]).strip()
    if sub_current != sub_standard:
        issue_text = f"Subaccount: '{sub_current}' → '{sub_standard}'"
        is_accepted, exception_info = is_deviation_accepted(exceptions_df, hotel_code, rate_code, country, "Subaccount", sub_current, sub_standard)
        if is_accepted:
            accepted_issues.append(issue_text + " (ACCEPTED)")
            exception_details.append(exception_info)
        else:
            issues.append(issue_text)
    
    # Check VAT
    vat_current = str(row["Vat type current"]).strip()
    vat_standard = str(row["Standard VAT"]).strip()
    if vat_current.lower() != vat_standard.lower():
        issue_text = f"VAT: '{vat_current}' → '{vat_standard}'"
        is_accepted, exception_info = is_deviation_accepted(exceptions_df, hotel_code, rate_code, country, "VAT", vat_current, vat_standard)
        if is_accepted:
            accepted_issues.append(issue_text + " (ACCEPTED)")
            exception_details.append(exception_info)
        else:
            issues.append(issue_text)
    
    # Check City Tax - ONLY for hotels that actually use it
    if should_check_city_tax(hotel_code):
        city_current = row["CityTax_Current"]
        city_standard = row["CityTax_Standard"] 
        if city_current != city_standard:
            issue_text = f"City Tax: '{city_current}' → '{city_standard}'"
            is_accepted, exception_info = is_deviation_accepted(exceptions_df, hotel_code, rate_code, country, "CityTax", city_current, city_standard)
            if is_accepted:
                accepted_issues.append(issue_text + " (ACCEPTED)")
                exception_details.append(exception_info)
            else:
                issues.append(issue_text)
    else:
        city_tax_skipped += 1
    
    # Create comprehensive deviation record
    base_record = {
        'Hotel_Code': hotel_code,
        'Rate_Code': rate_code,
        'Rate_Name': row.get('Name', ''),
        'Country': country,
        'Sub_Account_Current': sub_current,
        'Sub_Account_Standard': sub_standard,
        'Sub_Account_Match': sub_current == sub_standard,
        'VAT_Current': vat_current,
        'VAT_Standard': vat_standard,
        'VAT_Match': vat_current.lower() == vat_standard.lower(),
        'CityTax_Current': row["CityTax_Current"],
        'CityTax_Standard': row["CityTax_Standard"],
        'CityTax_Match': row["CityTax_Current"] == row["CityTax_Standard"],
        'CityTax_Checked': should_check_city_tax(hotel_code),
        'Service_Type_Current': row.get('Service type current', ''),
        'Valid_From_Current': row.get('Valid from current', '')
    }
    
    # Add to appropriate list
    if issues:
        record = base_record.copy()
        record['Deviation_Details'] = ' | '.join(issues)
        record['Status'] = 'NEEDS_FIXING'
        record['Priority'] = 'High' if len(issues) > 1 else 'Medium'
        true_deviations.append(record)
    
    if accepted_issues:
        record = base_record.copy()
        record['Deviation_Details'] = ' | '.join(accepted_issues)
        record['Status'] = 'ACCEPTED'
        # Add exception metadata
        if exception_details:
            record['Exception_Rule_Type'] = exception_details[0]['rule_type']
            record['Exception_Reason'] = exception_details[0]['reason']
            record['Approved_By'] = exception_details[0]['approved_by']
            record['Priority'] = exception_details[0]['priority']
            record['Review_Date'] = exception_details[0]['review_date']
        accepted_deviations.append(record)

print(f"🏨 City tax skipped for {city_tax_skipped:,} entries from non-city-tax hotels")

# Generate exception suggestions
print("🤖 Analyzing patterns for potential new exception rules...")
suggestions = generate_exception_suggestions(true_deviations, min_occurrences=3)

# Generate statistics
total_entries = len(analysis)
raw_compliant = len(analysis) - len(true_deviations) - len(accepted_deviations)
true_compliant = len(analysis) - len(true_deviations)

print("\n" + "="*80)
print("📈 ENHANCED COMPLIANCE SUMMARY WITH CITY TAX INTELLIGENCE")
print("="*80)
print(f"Total rate plans analyzed: {total_entries:,}")
print(f"Perfect matches: {raw_compliant:,} ({raw_compliant/total_entries*100:.1f}%)")
print(f"Accepted deviations: {len(accepted_deviations):,}")
print(f"True compliance (with exceptions): {true_compliant:,} ({true_compliant/total_entries*100:.1f}%)")
print(f"Issues requiring attention: {len(true_deviations):,} ({len(true_deviations)/total_entries*100:.1f}%)")
print(f"City tax checks skipped (non-applicable hotels): {city_tax_skipped:,}")

if suggestions:
    print(f"🤖 Suggested new exception rules: {len(suggestions)}")

# Create outputs
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
files_created = []

# 1. Issues to fix with priority sorting
if true_deviations:
    df = pd.DataFrame(true_deviations)
    # Sort by priority and number of issues
    priority_order = {'High': 3, 'Medium': 2, 'Low': 1}
    df['Priority_Score'] = df['Priority'].map(priority_order)
    df['Issue_Count'] = df['Deviation_Details'].str.count('\|') + 1
    df = df.sort_values(['Priority_Score', 'Issue_Count'], ascending=[False, False])
    df = df.drop(['Priority_Score', 'Issue_Count'], axis=1)
    
    filename = f"CRPM_Issues_To_Fix_PRIORITIZED_{timestamp}.xlsx"
    df.to_excel(filename, index=False)
    files_created.append(filename)
    print(f"\n🚨 Issues to fix (prioritized): {filename}")

# 2. Accepted deviations with full metadata
if accepted_deviations:
    df = pd.DataFrame(accepted_deviations)
    filename = f"CRPM_Accepted_Deviations_DETAILED_{timestamp}.xlsx"
    df.to_excel(filename, index=False)
    files_created.append(filename)
    print(f"✅ Accepted deviations (detailed): {filename}")

# 3. Exception rule suggestions
if suggestions:
    df = pd.DataFrame(suggestions)
    df = df.sort_values(['Priority', 'Occurrences'], ascending=[False, False])
    filename = f"CRPM_Suggested_Exception_Rules_{timestamp}.xlsx"
    df.to_excel(filename, index=False)
    files_created.append(filename)
    print(f"🤖 Suggested exception rules: {filename}")

# 4. Executive summary
summary_data = {
    'Metric': [
        'Total Rate Plans',
        'Perfect Compliance',
        'Accepted Deviations',
        'Issues Needing Fix',
        'Compliance Rate (%)',
        'City Tax Checks Skipped',
        'Top Issue Type',
        'Countries with Most Issues',
        'Suggested New Rules'
    ],
    'Value': [
        f"{total_entries:,}",
        f"{raw_compliant:,}",
        f"{len(accepted_deviations):,}",
        f"{len(true_deviations):,}",
        f"{true_compliant/total_entries*100:.1f}%",
        f"{city_tax_skipped:,}",
        "To be calculated",
        "To be calculated",
        f"{len(suggestions)}"
    ]
}

# Calculate top issues
if true_deviations:
    issues_summary = defaultdict(int)
    countries_summary = defaultdict(int)
    
    for dev in true_deviations:
        # Count issue types
        if 'VAT:' in dev['Deviation_Details']:
            issues_summary['VAT Deviations'] += 1
        if 'Subaccount:' in dev['Deviation_Details']:
            issues_summary['Subaccount Deviations'] += 1
        if 'City Tax:' in dev['Deviation_Details']:
            issues_summary['City Tax Deviations'] += 1
        
        countries_summary[dev['Country']] += 1
    
    if issues_summary:
        top_issue = max(issues_summary.items(), key=lambda x: x[1])
        summary_data['Value'][6] = f"{top_issue[0]} ({top_issue[1]} cases)"
    
    if countries_summary:
        top_countries = sorted(countries_summary.items(), key=lambda x: x[1], reverse=True)[:3]
        summary_data['Value'][7] = ", ".join([f"{country} ({count})" for country, count in top_countries])

summary_df = pd.DataFrame(summary_data)
filename = f"CRPM_Executive_Summary_{timestamp}.xlsx"
summary_df.to_excel(filename, index=False)
files_created.append(filename)
print(f"📊 Executive summary: {filename}")

print(f"\n📁 Total files created: {len(files_created)}")
for file in files_created:
    print(f"   📄 {file}")

if suggestions:
    print(f"\n🎯 TOP SUGGESTED EXCEPTION RULES:")
    for suggestion in suggestions[:3]:
        print(f"   • {suggestion['Rule_Type']} for {suggestion['Field']}: {suggestion['Occurrences']} instances")
        print(f"     Reason: {suggestion['Reason']}")

print(f"\n💡 CONFIGURATION TIP:")
print(f"   To add new city tax hotels, modify the CITY_TAX_HOTELS list at the top of this script.")
print(f"   To exclude hotels from analysis, modify the EXCLUDED_HOTELS list at the top of this script.")
print(f"   Current city tax hotels: {len(CITY_TAX_HOTELS)} configured")
print(f"   Current excluded hotels: {len(EXCLUDED_HOTELS)} configured")

print("\n🏁 Enhanced analysis complete with intelligent city tax handling!")