import pandas as pd
from datetime import datetime, timedelta
import os
from collections import defaultdict
from tqdm import tqdm
import webbrowser

# ============================================================================
# PATH CONFIGURATION
# ============================================================================
SOURCE_DIR = "source"
OUTPUT_BASE_DIR = "output"

# Input files
file = os.path.join(SOURCE_DIR, "source_CRPM_check.xlsx")
exceptions_file = os.path.join(SOURCE_DIR, "source_CRPM_exceptions.xlsx")

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

print(f"🏨 City tax will be checked for {len(CITY_TAX_HOTELS)} hotels: {', '.join(CITY_TAX_HOTELS)}")
# ============================================================================

def should_check_city_tax(hotel_code):
    """Return True only for hotels that actually use city tax"""
    return hotel_code in CITY_TAX_HOTELS

def load_exceptions(exceptions_file="source_CRPM_exceptions.xlsx"):
    """Load accepted deviations from Excel file - FIXED VERSION"""
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
        
        # Fix the Status column handling - THIS IS THE KEY FIX
        if 'Status' not in df.columns:
            df['Status'] = 'Active'  # Default all to Active if column doesn't exist
        else:
            # Replace empty/NaN status values with 'Active'
            df['Status'] = df['Status'].replace('', 'Active')
            df.loc[df['Status'].isna(), 'Status'] = 'Active'
        
        # Filter for active exceptions - CORRECTED LOGIC
        active_exceptions = df[df['Status'].str.lower() == 'active']
        inactive_exceptions = df[df['Status'].str.lower() != 'active']
        
        print(f"✅ Loaded {len(df)} total exceptions")
        print(f"✅ {len(active_exceptions)} exceptions are ACTIVE")
        
        if len(inactive_exceptions) > 0:
            print(f"⏸️ {len(inactive_exceptions)} exceptions are INACTIVE")
            print("\n⚠️ INACTIVE EXCEPTIONS (will be ignored):")
            for idx, exc in inactive_exceptions.head(3).iterrows():  # Show first 3
                rule_type = exc.get('Rule_Type', 'Unknown')
                field = exc.get('Field', 'Unknown')
                hotel = exc.get('Hotel_Code', '')
                status = exc.get('Status', '')
                reason = exc.get('Reason', '')
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
        
        # Parse deviation details to get field and values - FIXED REGEX
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

def create_actionable_dashboard(results, output_dir, timestamp):
    """Create a truly actionable dashboard focused on decisions and quick wins"""
    dashboard_file = os.path.join(output_dir, f"CRPM_Dashboard_{timestamp}.html")
    
    # Calculate metrics
    total_entries = results['total_entries']
    compliance_rate = results['compliance_rate']
    issues_count = len(results['true_deviations'])
    accepted_count = len(results['accepted_deviations'])
    
    # ACTIONABLE ANALYSIS - Find the patterns that matter
    
    # 1. QUICK WINS - Hotels with many issues (80/20 rule)
    hotel_issue_count = defaultdict(int)
    hotel_details = defaultdict(lambda: {'issues': [], 'countries': set(), 'priorities': defaultdict(int)})
    
    for dev in results['true_deviations']:
        hotel = dev['Hotel_Code']
        hotel_issue_count[hotel] += 1
        hotel_details[hotel]['issues'].append(dev['Deviation_Details'])
        hotel_details[hotel]['countries'].add(dev['Country'])
        hotel_details[hotel]['priorities'][dev['Priority']] += 1
    
    # Sort by issue count - these are your quick wins
    top_problem_hotels = sorted(hotel_issue_count.items(), key=lambda x: x[1], reverse=True)[:10]
    
    # Calculate impact: fixing top 5 hotels = what % improvement?
    top_5_issues = sum(count for _, count in top_problem_hotels[:5])
    impact_percentage = (top_5_issues / issues_count * 100) if issues_count > 0 else 0
    
    # 2. PATTERN DETECTION - What rules would fix the most issues?
    pattern_impact = defaultdict(lambda: {'count': 0, 'hotels': set(), 'countries': set(), 'examples': []})
    
    for dev in results['true_deviations']:
        details = dev['Deviation_Details']
        if 'VAT:' in details:
            # Extract VAT pattern
            vat_part = details.split('VAT: ')[1].split(' |')[0]
            if ' → ' in vat_part:
                current, standard = vat_part.replace("'", "").split(' → ')
                pattern_key = f"VAT_{current}_to_{standard}"
                pattern_impact[pattern_key]['count'] += 1
                pattern_impact[pattern_key]['hotels'].add(dev['Hotel_Code'])
                pattern_impact[pattern_key]['countries'].add(dev['Country'])
                if len(pattern_impact[pattern_key]['examples']) < 3:
                    pattern_impact[pattern_key]['examples'].append(f"{dev['Hotel_Code']} ({dev['Country']})")
        
        if 'Subaccount:' in details:
            sub_part = details.split('Subaccount: ')[1].split(' |')[0]
            if ' → ' in sub_part:
                current, standard = sub_part.replace("'", "").split(' → ')
                pattern_key = f"Subaccount_{current}_to_{standard}"
                pattern_impact[pattern_key]['count'] += 1
                pattern_impact[pattern_key]['hotels'].add(dev['Hotel_Code'])
                pattern_impact[pattern_key]['countries'].add(dev['Country'])
                if len(pattern_impact[pattern_key]['examples']) < 3:
                    pattern_impact[pattern_key]['examples'].append(f"{dev['Hotel_Code']} ({dev['Country']})")
    
    # Top patterns that would have biggest impact
    top_patterns = sorted(pattern_impact.items(), key=lambda x: x[1]['count'], reverse=True)[:5]
    
    # 3. EFFICIENCY ANALYSIS - How well are exception rules working?
    rule_efficiency = {
        'total_processed': total_entries,
        'auto_approved': accepted_count,
        'needs_review': issues_count,
        'efficiency_rate': (accepted_count / (accepted_count + issues_count) * 100) if (accepted_count + issues_count) > 0 else 100,
        'manual_work_remaining': issues_count
    }
    
    # 4. COUNTRY/REGION FOCUS - Which regions need attention?
    country_analysis = defaultdict(lambda: {'issues': 0, 'hotels': set(), 'types': defaultdict(int)})
    for dev in results['true_deviations']:
        country = dev['Country']
        country_analysis[country]['issues'] += 1
        country_analysis[country]['hotels'].add(dev['Hotel_Code'])
        if 'VAT:' in dev['Deviation_Details']:
            country_analysis[country]['types']['VAT'] += 1
        if 'Subaccount:' in dev['Deviation_Details']:
            country_analysis[country]['types']['Subaccount'] += 1
        if 'City Tax:' in dev['Deviation_Details']:
            country_analysis[country]['types']['City Tax'] += 1
    
    top_countries = sorted(country_analysis.items(), key=lambda x: x[1]['issues'], reverse=True)[:5]
    
    # Create the HTML with ACTIONABLE insights
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CRPM Action Dashboard - {timestamp}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        .dashboard {{
            background: white;
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            max-width: 1400px;
            margin: 0 auto;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid #667eea;
        }}
        .header h1 {{
            color: #333;
            margin: 0;
            font-size: 2.5rem;
        }}
        .subtitle {{
            color: #666;
            font-size: 1.1rem;
            margin-top: 10px;
        }}
        .metrics-bar {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }}
        .metric-box {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }}
        .metric-box:hover {{
            transform: translateY(-5px);
        }}
        .metric-label {{
            font-size: 0.9rem;
            opacity: 0.9;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 0 0 10px 0;
        }}
        .metric-value {{
            font-size: 2rem;
            font-weight: bold;
            margin: 0;
        }}
        .key-insights {{
            background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
            padding: 25px;
            border-radius: 15px;
            margin: 20px 0;
            color: #333;
        }}
        .key-insights h2 {{
            margin: 0 0 15px 0;
            color: #d63384;
        }}
        .insight-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }}
        .insight-card {{
            background: rgba(255,255,255,0.9);
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}
        .insight-number {{
            font-size: 2.5rem;
            font-weight: bold;
            color: #d63384;
            margin: 10px 0;
        }}
        .quick-wins {{
            background: #d4edda;
            border: 2px solid #28a745;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .quick-wins h2 {{
            color: #155724;
            margin: 0 0 15px 0;
        }}
        .hotel-card {{
            background: white;
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            border-left: 4px solid #28a745;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .hotel-info h4 {{
            margin: 0;
            color: #333;
        }}
        .hotel-info p {{
            margin: 5px 0 0 0;
            color: #666;
            font-size: 0.9rem;
        }}
        .issue-count {{
            background: #dc3545;
            color: white;
            padding: 8px 15px;
            border-radius: 20px;
            font-weight: bold;
        }}
        .pattern-insights {{
            background: #e3f2fd;
            border: 2px solid #2196f3;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .pattern-insights h2 {{
            color: #1976d2;
            margin: 0 0 15px 0;
        }}
        .pattern-card {{
            background: white;
            padding: 20px;
            margin: 10px 0;
            border-radius: 10px;
            border-left: 4px solid #2196f3;
        }}
        .pattern-card h4 {{
            margin: 0 0 10px 0;
            color: #1976d2;
        }}
        .impact-badge {{
            background: #2196f3;
            color: white;
            padding: 5px 10px;
            border-radius: 15px;
            font-size: 0.8rem;
            display: inline-block;
            margin-bottom: 10px;
        }}
        .action-plan {{
            background: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .action-plan h2 {{
            color: #856404;
            margin: 0 0 15px 0;
        }}
        .action-step {{
            background: white;
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            border-left: 4px solid #ffc107;
        }}
        .action-step h4 {{
            margin: 0 0 5px 0;
            color: #856404;
        }}
        .priority-high {{ color: #dc3545; font-weight: bold; }}
        .priority-medium {{ color: #ffc107; font-weight: bold; }}
        .priority-low {{ color: #28a745; font-weight: bold; }}
        .efficiency-meter {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
        }}
        .efficiency-bar {{
            background: #e9ecef;
            height: 20px;
            border-radius: 10px;
            overflow: hidden;
            margin: 10px 0;
        }}
        .efficiency-fill {{
            background: linear-gradient(90deg, #28a745, #20c997);
            height: 100%;
            transition: width 0.3s ease;
        }}
    </style>
</head>
<body>
<div class="dashboard">
    <div class="header">
        <h1>🎯 CRPM Action Dashboard</h1>
        <div class="subtitle">Focus on What Matters Most - {datetime.now().strftime('%B %d, %Y')}</div>
    </div>

    <!-- METRICS SUMMARY BAR -->
    <div class="metrics-bar">
        <div class="metric-box">
            <h3 class="metric-label">Total Rate Plans</h3>
            <p class="metric-value">{total_entries:,}</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Compliance Rate</h3>
            <p class="metric-value">{compliance_rate:.1f}%</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Issues to Fix</h3>
            <p class="metric-value">{issues_count:,}</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Accepted Deviations</h3>
            <p class="metric-value">{accepted_count:,}</p>
        </div>
    </div>

    <!-- KEY INSIGHTS SECTION -->
    <div class="key-insights">
        <h2>🚀 Key Insights - What You Need to Know</h2>
        <div class="insight-grid">
            <div class="insight-card">
                <div class="insight-number">{impact_percentage:.0f}%</div>
                <p><strong>Quick Win Impact</strong><br>Fixing top 5 hotels = {impact_percentage:.0f}% improvement</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{len(top_patterns)}</div>
                <p><strong>Automation Opportunities</strong><br>{sum(p[1]['count'] for p in top_patterns)} issues could be auto-fixed</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{rule_efficiency['efficiency_rate']:.0f}%</div>
                <p><strong>Current Automation</strong><br>{accepted_count:,} deviations auto-approved</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{len(top_problem_hotels)}</div>
                <p><strong>Focus Hotels</strong><br>These hotels need immediate attention
                </div>
            </div>
        </div>
        
        <!-- QUICK WINS SECTION -->
        <div class="quick-wins">
            <h2>🎯 Quick Wins - Fix These First (80/20 Rule)</h2>
            <p><strong>Strategy:</strong> These {len(top_problem_hotels[:5])} hotels account for {top_5_issues} issues ({impact_percentage:.0f}% of all problems). Fix these for maximum impact!</p>
            
            {f'''
            {"".join([f"""
            <div class="hotel-card">
                <div class="hotel-info">
                    <h4>Hotel {hotel}</h4>
                    <p>{len(hotel_details[hotel]['countries'])} countries • Priority: {max(hotel_details[hotel]['priorities'].items(), key=lambda x: x[1])[0] if hotel_details[hotel]['priorities'] else 'Medium'}</p>
                    <p><strong>Sample issues:</strong> {hotel_details[hotel]['issues'][0][:60]}{'...' if len(hotel_details[hotel]['issues'][0]) > 60 else ''}</p>
                </div>
                <div class="issue-count">{count} issues</div>
            </div>
            """ for hotel, count in top_problem_hotels[:5]])}
            ''' if top_problem_hotels else '''
            <div class="insight-card">
                <h3>🎉 Excellent Distribution!</h3>
                <p>No single hotel dominates your issues. Problems are well-distributed, indicating good overall management.</p>
            </div>
            '''}
        </div>
        
        <!-- AUTOMATION OPPORTUNITIES -->
        <div class="pattern-insights">
            <h2>🤖 Automation Opportunities - Create These Rules</h2>
            <p><strong>Smart Suggestion:</strong> These patterns appear frequently. Create exception rules to handle them automatically.</p>
            
            {f'''
            {"".join([f"""
            <div class="pattern-card">
                <div class="impact-badge">Would fix {data['count']} issues automatically</div>
                <h4>{pattern.replace('_', ' ').title()}</h4>
                <p><strong>Scope:</strong> {len(data['hotels'])} hotels across {len(data['countries'])} countries</p>
                <p><strong>Examples:</strong> {', '.join(data['examples'])}</p>
                <p><strong>Recommended Rule:</strong> Create a {'Country_Pattern' if len(data['countries']) == 1 else 'Hotel_Pattern'} exception for this deviation</p>
            </div>
            """ for pattern, data in top_patterns])}
            ''' if top_patterns else '''
            <div class="insight-card">
                <h3>🔧 Well Optimized!</h3>
                <p>No major automation opportunities found. Your exception rules are handling patterns effectively.</p>
            </div>
            '''}
        </div>
        
        <!-- REGIONAL FOCUS -->
        <div class="pattern-insights">
            <h2>🌍 Regional Focus - Where to Concentrate Efforts</h2>
            {f'''
            {"".join([f"""
            <div class="pattern-card">
                <h4>{country} - {data['issues']} issues</h4>
                <p><strong>Hotels affected:</strong> {len(data['hotels'])} hotels need attention</p>
                <p><strong>Main problems:</strong> {', '.join([f"{issue_type} ({count})" for issue_type, count in sorted(data['types'].items(), key=lambda x: x[1], reverse=True)])}</p>
                <p><strong>Action:</strong> {'Focus on VAT standardization' if data['types'].get('VAT', 0) > data['issues']/2 else 'Review subaccount configurations' if data['types'].get('Subaccount', 0) > data['issues']/2 else 'Mixed issues - needs detailed review'}</p>
            </div>
            """ for country, data in top_countries])}
            ''' if top_countries else '''
            <div class="insight-card">
                <h3>🌟 Global Excellence!</h3>
                <p>Issues are well-distributed across regions. No single country requires special focus.</p>
            </div>
            '''}
        </div>
        
        <!-- ACTION PLAN -->
        <div class="action-plan">
            <h2>📋 Your 30-Day Action Plan</h2>
            
            <div class="action-step">
                <h4>Week 1: Quick Wins (Immediate Impact)</h4>
                <p>• Review top {min(3, len(top_problem_hotels))} hotels: {', '.join([hotel for hotel, _ in top_problem_hotels[:3]])}</p>
                <p>• Expected impact: Fix ~{sum(count for _, count in top_problem_hotels[:3])} issues ({sum(count for _, count in top_problem_hotels[:3])/issues_count*100:.0f}% improvement)</p>
            </div>
            
            <div class="action-step">
                <h4>Week 2-3: Automation Rules</h4>
                <p>• Implement {min(2, len(top_patterns))} new exception rules for top patterns</p>
                <p>• Expected impact: Auto-handle {sum(data['count'] for _, data in top_patterns[:2])} future similar issues</p>
            </div>
            
            <div class="action-step">
                <h4>Week 4: Regional Focus</h4>
                <p>• Address {top_countries[0][0] if top_countries else 'regional'} issues systematically</p>
                <p>• Set up monitoring for early detection of similar patterns</p>
            </div>
            
            <div class="action-step">
                <h4>Ongoing: Monitor & Optimize</h4>
                <p>• Target: Achieve {min(98, compliance_rate + 5):.0f}% compliance rate</p>
                <p>• Review exception rule effectiveness monthly</p>
            </div>
        </div>
        
        <!-- EFFICIENCY TRACKING -->
        <div class="efficiency-meter">
            <h3>🚀 Automation Efficiency</h3>
            <p>Your exception rules are handling <strong>{rule_efficiency['efficiency_rate']:.0f}%</strong> of deviations automatically</p>
            <div class="efficiency-bar">
                <div class="efficiency-fill" style="width: {rule_efficiency['efficiency_rate']}%"></div>
            </div>
            <p><small>{accepted_count:,} auto-approved • {issues_count:,} need manual review • Target: 95% automation</small></p>
        </div>
        
        <div style="text-align: center; margin-top: 30px; color: #666; font-size: 0.9rem;">
            <p>📊 Actionable CRPM Dashboard - Focus on Impact, Not Just Data</p>
            <p>Next recommended review: {(datetime.now() + timedelta(days=7)).strftime('%B %d, %Y')}</p>
        </div>
    </div>
</body>
</html>
"""
    dashboard_file = os.path.join(output_dir, f"CRPM_Dashboard_{timestamp}.html")
    with open(dashboard_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    return dashboard_file

print("🚀 Starting SIMPLIFIED ENHANCED CRPM analysis...")

# Load exceptions
exceptions_df = load_exceptions(exceptions_file)

# Load data
print("📊 Loading data files...")
data = pd.read_excel(file, sheet_name="data", dtype=str)
standard = pd.read_excel(file, sheet_name="standard", dtype=str)

try:
    print("📊 Loading data files...")
    data = pd.read_excel(file, sheet_name="data", dtype=str)
    standard = pd.read_excel(file, sheet_name="standard", dtype=str)
    
    # Clean column names
    data.columns = data.columns.str.strip()
    standard.columns = standard.columns.str.strip()
    
    print(f"✅ Loaded {len(data)} rate plans")
    print(f"✅ Loaded {len(standard)} hotel standards")
    
except Exception as e:
    print(f"❌ Error loading data: {e}")
    print("Make sure 'source_CRPM_check.xlsx' exists and has 'data' and 'standard' sheets")
    exit(1)

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

print("🔍 Checking for deviations...")

# Process in batches with progress bar
batch_size = BATCH_SIZE
for batch_start in tqdm(range(0, len(analysis), batch_size), desc="Processing compliance"):
    batch_end = min(batch_start + batch_size, len(analysis))
    batch = analysis.iloc[batch_start:batch_end]
    
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

print(f"🏨 City tax skipped for {city_tax_skipped:,} entries from non-city-tax hotels")

# Generate exception suggestions
print("🤖 Analyzing patterns for potential new exception rules...")
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

print("\n" + "="*80)
print("📈 SIMPLIFIED ENHANCED COMPLIANCE SUMMARY")
print("="*80)
print(f"Total rate plans analyzed: {total_entries:,}")
print(f"Perfect matches: {raw_compliant:,} ({raw_compliant/total_entries*100:.1f}%)")
print(f"Accepted deviations: {len(accepted_deviations):,}")
print(f"True compliance (with exceptions): {true_compliant:,} ({true_compliant/total_entries*100:.1f}%)")
print(f"Issues requiring attention: {len(true_deviations):,} ({len(true_deviations)/total_entries*100:.1f}%)")
print(f"City tax checks skipped (non-applicable hotels): {city_tax_skipped:,}")

if suggestions:
    print(f"🤖 Suggested new exception rules: {len(suggestions)}")

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
    print(f"\n🚨 Issues to fix (prioritized): {os.path.basename(filename)}")

# 2. Accepted deviations with full metadata
if accepted_deviations:
    df = pd.DataFrame(accepted_deviations)
    filename = os.path.join(output_dir, f"CRPM_Accepted_Deviations_DETAILED_{timestamp}.xlsx")
    df.to_excel(filename, index=False)
    files_created.append(filename)
    print(f"✅ Accepted deviations (detailed): {os.path.basename(filename)}")

# 3. Exception rule suggestions
if suggestions:
    df = pd.DataFrame(suggestions)
    df = df.sort_values(['Priority', 'Occurrences'], ascending=[False, False])
    filename = os.path.join(output_dir, f"CRPM_Suggested_Exception_Rules_{timestamp}.xlsx")
    df.to_excel(filename, index=False)
    files_created.append(filename)
    print(f"🤖 Suggested exception rules: {os.path.basename(filename)}")

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
print(f"📊 Executive summary: {os.path.basename(filename)}")

# 5. Create HTML Dashboard (if enabled)
if CREATE_DASHBOARD:
    print("🎨 Creating HTML dashboard...")
    try:
        dashboard_file = create_actionable_dashboard(results, output_dir, timestamp)
        files_created.append(dashboard_file)
        print(f"🌟 Dashboard created: {os.path.basename(dashboard_file)}")
        
        # Auto-open dashboard
        if AUTO_OPEN_DASHBOARD:
            try:
                webbrowser.open(f'file:///{os.path.abspath(dashboard_file)}')
                print("🚀 Dashboard opened in web browser!")
            except:
                print("⚠️ Could not auto-open dashboard. Please open manually.")
                
    except Exception as e:
        print(f"⚠️ Could not create dashboard: {e}")

print(f"\n📁 Total files created: {len(files_created)}")
print(f"📂 Output folder: {output_dir}")
for file in files_created:
    print(f"   📄 {os.path.basename(file)}")

if suggestions:
    print(f"\n🎯 TOP SUGGESTED EXCEPTION RULES:")
    for suggestion in suggestions[:3]:
        print(f"   • {suggestion['Rule_Type']} for {suggestion['Field']}: {suggestion['Occurrences']} instances")
        print(f"     Reason: {suggestion['Reason']}")

print(f"\n💡 CONFIGURATION TIPS:")
print(f"   • To add new city tax hotels, modify the CITY_TAX_HOTELS list at the top")
print(f"   • To disable dashboard, set CREATE_DASHBOARD = False at the top")
print(f"   • To skip auto-opening browser, set AUTO_OPEN_DASHBOARD = False at the top")
print(f"   • Current city tax hotels: {len(CITY_TAX_HOTELS)} configured")

print("\n🏁 Simplified enhanced analysis complete!")