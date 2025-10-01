import pandas as pd
import os
from datetime import datetime
from collections import defaultdict

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
