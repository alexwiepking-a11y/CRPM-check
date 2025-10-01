import pandas as pd
from datetime import datetime, timedelta
import os
from collections import defaultdict
from tqdm import tqdm
import webbrowser
from exceptions import load_exceptions, is_deviation_accepted, generate_exception_suggestions
from dashboard import create_actionable_dashboard
import logging
# ...existing imports...
import argparse

parser = argparse.ArgumentParser(description="CRPM Compliance Checker")
parser.add_argument('--input', type=str, default="source/source_CRPM_check.xlsx", help="Path to input Excel file")
parser.add_argument('--exceptions', type=str, default="source/source_CRPM_exceptions.xlsx", help="Path to exceptions Excel file")
parser.add_argument('--output', type=str, default="output", help="Output directory")
args = parser.parse_args()

logging.basicConfig(
    level=logging.INFO,  # Change to logging.DEBUG for more details
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ============================================================================
# PATH CONFIGURATION
# ============================================================================
# Use command-line arguments for file paths
file = args.input
exceptions_file = args.exceptions
OUTPUT_BASE_DIR = args.output

# Clean timestamp: YYYY-MM-DD_HHMM
timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")

# Output directory → e.g. output/2025-09-30_1615
output_dir = os.path.join(OUTPUT_BASE_DIR, timestamp)
os.makedirs(output_dir, exist_ok=True)

# ============================================================================
# CONFIGURATION SECTION - EASY TO MODIFY
# ============================================================================
# Hotels that actually use city tax - modify this list as needed
CITY_TAX_HOTELS = [
    'AMS', 'AMA','AMZ', 'RTM', 'NYT', 'NYB', 'PGL', 
    'PCG', 'POP', 'PLD', 'GEN', 'ZUR', 'RIT', 'KLB'
]

EXCLUDED_HOTELS = [
    'ITA', 'VRS', 'VRSM',
    'NEW_HOTEL1', 'NEW_HOTEL2'  # Add new ones here
]

# Simple configuration options
CREATE_DASHBOARD = True  # Set to False to skip HTML dashboard
AUTO_OPEN_DASHBOARD = True  # Set to False to not auto-open browser
BATCH_SIZE = 1000  # Increase for faster processing of large files

logging.info(f"🏨 City tax will be checked for {len(CITY_TAX_HOTELS)} hotels: {', '.join(CITY_TAX_HOTELS)}")
# ============================================================================

def should_check_city_tax(hotel_code):
    """Return True if the hotel uses city tax, otherwise False."""
    return hotel_code in CITY_TAX_HOTELS

# Load exceptions
exceptions_df = load_exceptions(exceptions_file)

# Load data
logging.info("📊 Loading data files...")
try:
    data = pd.read_excel(file, sheet_name="data", dtype=str)
    standard = pd.read_excel(file, sheet_name="standard", dtype=str)
    
    # Clean column names
    data.columns = data.columns.str.strip()
    standard.columns = standard.columns.str.strip()
    
    logging.info(f"✅ Loaded {len(data)} rate plans")
    logging.info(f"✅ Loaded {len(standard)} hotel standards")
    
except Exception as e:
    logging.error(f"❌ Error loading data: {e}")
    logging.error("Make sure 'source_CRPM_check.xlsx' exists and has 'data' and 'standard' sheets")
    exit(1)

# Check for required columns
required_data_cols = ["Hotel code", "Is subject to city tax current", "Sub account current", "Vat type current"]
required_standard_cols = ["Hotel code", "Standard subaccount", "Standard VAT", "Standard City tax"]

missing_data_cols = [col for col in required_data_cols if col not in data.columns]
missing_standard_cols = [col for col in required_standard_cols if col not in standard.columns]

if missing_data_cols:
    logging.error(f"Missing columns in 'data' sheet: {missing_data_cols}")
    exit(1)
if missing_standard_cols:
    logging.error(f"Missing columns in 'standard' sheet: {missing_standard_cols}")
    exit(1)

# Normalize city tax
def normalize_boolean(val):
    """Convert various representations of boolean values to 'Yes', 'No', or 'Unknown'."""
    val_str = str(val).strip().lower()
    if val_str in ["true", "1", "1.0", "yes", "y"]:
        return "Yes"
    if val_str in ["false", "0", "0.0", "no", "n", "nan", "none"]:
        return "No"
    return "Unknown"

try:
    data["CityTax_Current"] = data["Is subject to city tax current"].apply(normalize_boolean)
except Exception as e:
    logging.error(f"Error normalizing city tax values: {e}")
    exit(1)

standard["CityTax_Standard"] = standard["Standard City tax"].apply(normalize_boolean)

# Merge 
merged = data.merge(standard, how="left", on="Hotel code")

# Filter out entries without standards
analysis = merged[merged["Standard subaccount"].notna()].copy()
logging.info(f"📈 Analyzing {len(analysis)} entries with standards")

# Add deviation columns using vectorized operations
analysis['Subaccount_Deviates'] = analysis['Sub account current'].str.strip() != analysis['Standard subaccount'].str.strip()
analysis['VAT_Deviates'] = analysis['Vat type current'].str.strip().str.lower() != analysis['Standard VAT'].str.strip().str.lower()
analysis['CityTax_Deviates'] = analysis.apply(
    lambda row: row['CityTax_Current'] != row['CityTax_Standard'] if should_check_city_tax(row['Hotel code']) else False,
    axis=1
)

# Now, you can filter deviations much faster:
deviation_mask = analysis['Subaccount_Deviates'] | analysis['VAT_Deviates'] | analysis['CityTax_Deviates']
deviation_rows = analysis[deviation_mask]

# You can then loop only over deviation_rows for further processing (exceptions, etc.)
true_deviations = []
accepted_deviations = []
city_tax_skipped = 0

logging.info("🔍 Checking for deviations...")

# Process in batches with progress bar
batch_size = BATCH_SIZE
for batch_start in tqdm(range(0, len(deviation_rows), batch_size), desc="Processing compliance"):
    batch_end = min(batch_start + batch_size, len(deviation_rows))
    batch = deviation_rows.iloc[batch_start:batch_end]
    
    for idx, row in batch.iterrows():
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

logging.info(f"🏨 City tax skipped for {city_tax_skipped:,} entries from non-city-tax hotels")

# Generate exception suggestions
logging.info("🤖 Analyzing patterns for potential new exception rules...")
suggestions = generate_exception_suggestions(true_deviations, min_occurrences=3)

# Generate statistics
total_entries = len(analysis)
raw_compliant = len(analysis) - len(true_deviations) - len(accepted_deviations)
true_compliant = len(analysis) - len(true_deviations)
processing_time = 0  # Simple version doesn't track time

# Store results for dashboard
results = {
    'total_entries': total_entries,
    'compliance_rate': true_compliant/total_entries*100,
    'true_deviations': true_deviations,
    'accepted_deviations': accepted_deviations,
    'suggestions': suggestions,
    'processing_time': processing_time
}

logging.info("\n" + "="*80)
logging.info("📈 SIMPLIFIED ENHANCED COMPLIANCE SUMMARY")
logging.info("="*80)
logging.info(f"Total rate plans analyzed: {total_entries:,}")
logging.info(f"Perfect matches: {raw_compliant:,} ({raw_compliant/total_entries*100:.1f}%)")
logging.info(f"Accepted deviations: {len(accepted_deviations):,}")
logging.info(f"True compliance (with exceptions): {true_compliant:,} ({true_compliant/total_entries*100:.1f}%)")
logging.info(f"Issues requiring attention: {len(true_deviations):,} ({len(true_deviations)/total_entries*100:.1f}%)")
logging.info(f"City tax checks skipped (non-applicable hotels): {city_tax_skipped:,}")

if suggestions:
    logging.info(f"🤖 Suggested new exception rules: {len(suggestions)}")

files_created = []

# 1. Issues to fix with priority sorting
if true_deviations:
    df = pd.DataFrame(true_deviations)
    # Sort by priority and number of issues
    priority_order = {'High': 3, 'Medium': 2, 'Low': 1}
    df['Priority_Score'] = df['Priority'].map(priority_order)
    df['Issue_Count'] = df['Deviation_Details'].str.count(r'\|') + 1  # Fixed regex
    df = df.sort_values(['Priority_Score', 'Issue_Count'], ascending=[False, False])
    df = df.drop(['Priority_Score', 'Issue_Count'], axis=1)
    
    filename = os.path.join(output_dir, f"CRPM_Issues_To_Fix_PRIORITIZED_{timestamp}.xlsx")
    df.to_excel(filename, index=False)
    files_created.append(filename)
    logging.info(f"\n🚨 Issues to fix (prioritized): {os.path.basename(filename)}")

# 2. Accepted deviations with full metadata
if accepted_deviations:
    df = pd.DataFrame(accepted_deviations)
    filename = os.path.join(output_dir, f"CRPM_Accepted_Deviations_DETAILED_{timestamp}.xlsx")
    df.to_excel(filename, index=False)
    files_created.append(filename)
    logging.info(f"✅ Accepted deviations (detailed): {os.path.basename(filename)}")

# 3. Exception rule suggestions
if suggestions:
    df = pd.DataFrame(suggestions)
    df = df.sort_values(['Priority', 'Occurrences'], ascending=[False, False])
    filename = os.path.join(output_dir, f"CRPM_Suggested_Exception_Rules_{timestamp}.xlsx")
    df.to_excel(filename, index=False)
    files_created.append(filename)
    logging.info(f"🤖 Suggested exception rules: {os.path.basename(filename)}")

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
filename = os.path.join(output_dir, f"CRPM_Executive_Summary_{timestamp}.xlsx")
summary_df.to_excel(filename, index=False)
files_created.append(filename)
logging.info(f"📊 Executive summary: {os.path.basename(filename)}")

# 5. Create HTML Dashboard (if enabled)
if CREATE_DASHBOARD:
    logging.info("🎨 Creating HTML dashboard...")
    try:
        dashboard_file = create_actionable_dashboard(results, output_dir, timestamp)
        files_created.append(dashboard_file)
        logging.info(f"🌟 Dashboard created: {os.path.basename(dashboard_file)}")
        
        # Auto-open dashboard
        if AUTO_OPEN_DASHBOARD:
            try:
                webbrowser.open(f'file:///{os.path.abspath(dashboard_file)}')
                logging.info("🚀 Dashboard opened in web browser!")
            except Exception as e:
                logging.warning("⚠️ Could not auto-open dashboard. Please open manually.")
                
    except Exception as e:
        logging.warning(f"⚠️ Could not create dashboard: {e}")

logging.info(f"\n📁 Total files created: {len(files_created)}")
logging.info(f"📂 Output folder: {output_dir}")
for file in files_created:
    logging.info(f"   📄 {os.path.basename(file)}")

if suggestions:
    logging.info(f"\n🎯 TOP SUGGESTED EXCEPTION RULES:")
    for suggestion in suggestions[:3]:
        logging.info(f"   • {suggestion['Rule_Type']} for {suggestion['Field']}: {suggestion['Occurrences']} instances")
        logging.info(f"     Reason: {suggestion['Reason']}")

logging.info(f"\n💡 CONFIGURATION TIPS:")
logging.info(f"   • To add new city tax hotels, modify the CITY_TAX_HOTELS list at the top")
logging.info(f"   • To disable dashboard, set CREATE_DASHBOARD = False at the top")
logging.info(f"   • To skip auto-opening browser, set AUTO_OPEN_DASHBOARD = False at the top")
logging.info(f"   • Current city tax hotels: {len(CITY_TAX_HOTELS)} configured")

logging.info("\n🏁 Simplified enhanced analysis complete!")